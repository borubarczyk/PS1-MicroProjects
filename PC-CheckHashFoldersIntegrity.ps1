Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Debug -------------------------------------------------------------------
# DEBUG (nie zaśmiecaj pipeline!)
$script:DEBUG = $true
$script:LogPath = Join-Path -Path $env:TEMP -ChildPath "PC-CheckHashFoldersIntegrity.log"

function Write-DebugLog {
    param([string]$Message)
    if (-not $script:DEBUG) { return }
    try {
        $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
        $line = "[DBG $ts] $Message"
        # tylko do pliku (i NIC do pipeline!)
        [System.IO.File]::AppendAllText($script:LogPath, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    } catch { }
}

function Get-KeysFromMap {
    param([object]$Map)
    if ($null -eq $Map) { return @() }
    try {
        if ($Map -is [System.Collections.IDictionary])       { return @($Map.Keys) }
        elseif ($Map.PSObject.Properties['Keys'])            { return @($Map.Keys) }
        else                                                 { return @() }
    } catch { return @() }
}


# --- Helpers -----------------------------------------------------------------
function Get-MapValue {
    param(
        [Parameter(Mandatory = $true)][object]$Dict,
        [Parameter(Mandatory = $true)][string]$Key
    )
    if ($null -eq $Dict) { return $null }
    try {
        # Hashtable or object exposing ContainsKey
        if ($Dict.PSObject.Methods['ContainsKey']) {
            if ($Dict.ContainsKey($Key)) { return $Dict[$Key] } else { return $null }
        }
        # Generic Dictionary: TryGetValue
        if ($Dict.PSObject.Methods['TryGetValue']) {
            $out = $null
            if ($Dict.TryGetValue($Key, [ref]$out)) { return $out } else { return $null }
        }
        # Fallback for IDictionary
        if ($Dict -is [System.Collections.IDictionary]) {
            return $Dict[$Key]
        }
    }
    catch { }
    return $null
}

# --- Narzędzia ---------------------------------------------------------------

function Get-FileHashSafe {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "ERROR:EmptyPath" }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, 'Open', 'Read', 'Read')
        $hashBytes = $sha.ComputeHash($stream)
        ($hashBytes | ForEach-Object { $_.ToString("x2") }) -join ''
    }
    catch {
        "ERROR:" + $_.Exception.Message
    }
    finally {
        if ($stream) { $stream.Dispose() }
        $sha.Dispose()
    }
}

function Get-FilesIndex {
    param(
        [Parameter(Mandatory)][string]$Root,
        [System.Windows.Forms.ProgressBar]$ProgressBar = $null,
        [System.Windows.Forms.Label]$StatusLabel = $null
    )
    $rootClean = (Get-Item -LiteralPath $Root).FullName.TrimEnd('\', '/')
    Write-DebugLog "Get-FilesIndex: Root='$Root' rootClean='$rootClean'"
    $files = Get-ChildItem -LiteralPath $rootClean -Recurse -File -ErrorAction Stop
    $index = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    $total = [math]::Max(1, $files.Count)
    Write-DebugLog "Get-FilesIndex: files=$($files.Count)"
    $i = 0
    foreach ($f in $files) {
        $rel = $f.FullName.Substring($rootClean.Length).TrimStart('\', '/')
        $index[$rel] = [pscustomobject]@{
            Rel              = $rel
            FullName         = $f.FullName
            Length           = $f.Length
            LastWriteTimeUtc = $f.LastWriteTimeUtc
        }
        $i++
        if ($ProgressBar) { $ProgressBar.Value = [int](($i / $total) * 100) }
        if ($StatusLabel) { $StatusLabel.Text = "Indeksuję: $i / $total"; [System.Windows.Forms.Application]::DoEvents() }
    }
    Write-DebugLog "Get-FilesIndex: index.Count=$($index.Count)"
    return $index
}

function Compute-FastFolderSignature {
    param([object]$Index)  # było [hashtable] — powodowało null dla Dictionary

    if ($null -eq $Index) { Write-DebugLog "Compute-FastFolderSignature: Index null"; return '' }

    $keys = Get-KeysFromMap $Index
    Write-DebugLog "Compute-FastFolderSignature: Keys.Count=$(@($keys).Count)"

    $sha = [System.Security.Cryptography.SHA256]::Create()
    $sb  = New-Object System.Text.StringBuilder

    foreach ($k in ($keys | Sort-Object)) {
        $item = Get-MapValue -Dict $Index -Key ([string]$k)
        if ($null -eq $item) { continue }
        [void]$sb.AppendLine(('{0}|{1}|{2:o}' -f $k, $item.Length, $item.LastWriteTimeUtc))
    }

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($sb.ToString())
    $hash  = ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join ''
    $sha.Dispose()
    Write-DebugLog "Compute-FastFolderSignature: Done sig=$hash"
    return $hash
}



function Compare-FoldersGuiCore {
    param(
        [string]$SourcePath,
        [string]$DestPath,
        [switch]$UseHash,
        [switch]$FastPrecheck,
        [switch]$TrustFastPrecheck,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel
    )

    $StatusLabel.Text = "Indeksuję źródło…"; [System.Windows.Forms.Application]::DoEvents()
    $src = Get-FilesIndex -Root $SourcePath -ProgressBar $ProgressBar -StatusLabel $StatusLabel

    $StatusLabel.Text = "Indeksuję cel…"; [System.Windows.Forms.Application]::DoEvents()
    $dst = Get-FilesIndex -Root $DestPath -ProgressBar $ProgressBar -StatusLabel $StatusLabel

    Write-DebugLog ("Compare: srcType={0} dstType={1}" -f $src.GetType().FullName, $dst.GetType().FullName)
    Write-DebugLog ("Compare: src.Count={0} dst.Count={1}" -f $src.Count, $dst.Count)

    if ($FastPrecheck) {
        $StatusLabel.Text = "Szybki wstępny test (rozmiar+czas)…"; [System.Windows.Forms.Application]::DoEvents()
        $sigSrc = Compute-FastFolderSignature -Index $src
        $sigDst = Compute-FastFolderSignature -Index $dst
        Write-DebugLog "Compare: Precheck sigSrc=$sigSrc sigDst=$sigDst"

        if ($sigSrc -eq $sigDst) {
            if ($TrustFastPrecheck) {
                return ,([pscustomobject]@{
                    Path='(cały katalog)'; Status='LikelySame (fast precheck)'
                    SourceSize=$null; DestSize=$null; SourceTime=$null; DestTime=$null
                    SourceHash=$null; DestHash=$null; Source=$SourcePath; Destination=$DestPath
                })
            } else {
                $StatusLabel.Text = "Precheck zgodny → weryfikuję hashami plików…"
                [System.Windows.Forms.Application]::DoEvents()
            }
        } else {
            $StatusLabel.Text = "Precheck różny → pełne porównanie…"
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    # --- KLUCZE: unia ścieżek (bez HashSet i bez ToArray)
    $srcKeys = Get-KeysFromMap $src
    $dstKeys = Get-KeysFromMap $dst
    $allKeys = @($srcKeys + $dstKeys) | Sort-Object -Unique

    Write-DebugLog ("Compare: srcKeys={0} dstKeys={1} union={2}" -f $srcKeys.Count, $dstKeys.Count, $allKeys.Count)

    $total = [math]::Max(1, $allKeys.Count)
    $idx   = 0
    $rows  = New-Object System.Collections.Generic.List[object]

    foreach ($rel in $allKeys) {
        $idx++
        if ($ProgressBar) { $ProgressBar.Value = [int](($idx/$total)*100) }
        if ($StatusLabel) { $StatusLabel.Text = "Porównuję: $idx / $total — $rel"; [System.Windows.Forms.Application]::DoEvents() }

        $s = Get-MapValue -Dict $src -Key $rel
        $d = Get-MapValue -Dict $dst -Key $rel
        if ($null -eq $s -and $null -eq $d) { Write-DebugLog "Compare: both missing '$rel' (shouldn't happen)"; continue }

        if ($null -eq $s) {
            $rows.Add([pscustomobject]@{
                Path=$rel; Status='OnlyInDestination'
                SourceSize=$null; DestSize=$d.Length
                SourceTime=$null; DestTime=$d.LastWriteTimeUtc
                SourceHash=$null; DestHash= if ($UseHash) { Get-FileHashSafe $d.FullName } else { $null }
                Source=$null; Destination=$d.FullName
            })
            continue
        }
        if ($null -eq $d) {
            $rows.Add([pscustomobject]@{
                Path=$rel; Status='OnlyInSource'
                SourceSize=$s.Length; DestSize=$null
                SourceTime=$s.LastWriteTimeUtc; DestTime=$null
                SourceHash= if ($UseHash) { Get-FileHashSafe $s.FullName } else { $null }; DestHash=$null
                Source=$s.FullName; Destination=$null
            })
            continue
        }

        $sameMeta = ($s.Length -eq $d.Length) -and ($s.LastWriteTimeUtc -eq $d.LastWriteTimeUtc)
        $same     = $sameMeta
        $sHash = $null; $dHash = $null
        if ($UseHash) {
            $sHash = Get-FileHashSafe -Path $s.FullName
            $dHash = Get-FileHashSafe -Path $d.FullName
            $same  = ($sHash -eq $dHash)
            if (-not $same) { Write-DebugLog "Hash diff '$rel' s=$sHash d=$dHash" }
        }

        $rows.Add([pscustomobject]@{
            Path=$rel
            Status= if ($same) { 'Same' } elseif ($sameMeta) { 'Different (content)' } else { 'Different (meta/content)' }
            SourceSize=$s.Length; DestSize=$d.Length
            SourceTime=$s.LastWriteTimeUtc; DestTime=$d.LastWriteTimeUtc
            SourceHash=$sHash; DestHash=$dHash
            Source=$s.FullName; Destination=$d.FullName
        })
    }

    Write-DebugLog "Compare: FINISH rows=$($rows.Count)"
    return $rows
}


function New-ResultsDataTable {
    param([System.Collections.IEnumerable]$Items)
    Write-DebugLog "New-ResultsDataTable: inputCount=$((@($Items)).Count)"
    $dt = New-Object System.Data.DataTable 'Results'
    foreach ($col in 'Path', 'Status', 'SourceSize', 'DestSize', 'SourceTime', 'DestTime', 'SourceHash', 'DestHash', 'Source', 'Destination') {
        [void]$dt.Columns.Add($col)
    }
    foreach ($it in $Items) {
        $row = $dt.NewRow()
        foreach ($col in $dt.Columns) {
            try { $row[$col.ColumnName] = ($it.($col.ColumnName)) }
            catch { Write-DebugLog "New-ResultsDataTable: set cell failed col=$($col.ColumnName) err=$($_.Exception.Message)" }
        }
        [void]$dt.Rows.Add($row)
    }
    Write-DebugLog "New-ResultsDataTable: rows=$($dt.Rows.Count) cols=$($dt.Columns.Count)"
    $dt
}

# --- GUI ---------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Porównanie folderów (hash domyślnie)"
$form.Size = New-Object System.Drawing.Size(1100, 700)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::White
$form.MinimumSize = New-Object System.Drawing.Size(900, 600)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font

$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Text = "Folder źródłowy:"
$lblSrc.Location = '15,15'; $lblSrc.AutoSize = $true
$form.Controls.Add($lblSrc)

$txtSrc = New-Object System.Windows.Forms.TextBox
$txtSrc.Location = '15,35'; $txtSrc.Size = '920,25'
$txtSrc.Anchor = 'Top,Left,Right'
$form.Controls.Add($txtSrc)

$btnSrc = New-Object System.Windows.Forms.Button
$btnSrc.Text = "Wybierz…"; $btnSrc.Location = '945,33'
$btnSrc.Anchor = 'Top,Right'
$btnSrc.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$btnSrc.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Wybierz folder źródłowy"
        if ($dlg.ShowDialog() -eq 'OK') { $txtSrc.Text = $dlg.SelectedPath }
    })
$form.Controls.Add($btnSrc)

$lblDst = New-Object System.Windows.Forms.Label
$lblDst.Text = "Folder docelowy (np. udział sieciowy):"
$lblDst.Location = '15,70'; $lblDst.AutoSize = $true
$form.Controls.Add($lblDst)

$txtDst = New-Object System.Windows.Forms.TextBox
$txtDst.Location = '15,90'; $txtDst.Size = '920,25'
$txtDst.Anchor = 'Top,Left,Right'
$form.Controls.Add($txtDst)

$btnDst = New-Object System.Windows.Forms.Button
$btnDst.Text = "Wybierz…"; $btnDst.Location = '945,88'
$btnDst.Anchor = 'Top,Right'
$btnDst.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$btnDst.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Wybierz folder docelowy"
        if ($dlg.ShowDialog() -eq 'OK') { $txtDst.Text = $dlg.SelectedPath }
    })
$form.Controls.Add($btnDst)

$chkUseHash = New-Object System.Windows.Forms.CheckBox
$chkUseHash.Text = "Użyj hashy (SHA-256) — PEWNE porównanie"
$chkUseHash.Checked = $true
$chkUseHash.Location = '15,125'; $chkUseHash.AutoSize = $true
$form.Controls.Add($chkUseHash)

$chkPre = New-Object System.Windows.Forms.CheckBox
$chkPre.Text = "Szybki wstępny test (rozmiar+czas) — niepewny, ale szybki"
$chkPre.Checked = $false
$chkPre.Location = '15,149'; $chkPre.AutoSize = $true
$form.Controls.Add($chkPre)

$chkTrustPre = New-Object System.Windows.Forms.CheckBox
$chkTrustPre.Text = "Zatrzymaj po wstępnym teście, jeśli zgodny (NIE 100% pewne)"
$chkTrustPre.Checked = $false
$chkTrustPre.Location = '35,173'; $chkTrustPre.AutoSize = $true
$form.Controls.Add($chkTrustPre)

$chkOnlyDiff = New-Object System.Windows.Forms.CheckBox
$chkOnlyDiff.Text = "Pokaż tylko różnice"
$chkOnlyDiff.Checked = $true
$chkOnlyDiff.Location = '15,197'; $chkOnlyDiff.AutoSize = $true
$form.Controls.Add($chkOnlyDiff)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Start"; $btnRun.Location = '15,225'
$btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$form.Controls.Add($btnRun)
$form.AcceptButton = $btnRun

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Eksportuj CSV…"; $btnExport.Location = '95,225'; $btnExport.Enabled = $false
$btnExport.FlatStyle = [System.Windows.Forms.FlatStyle]::System
$form.Controls.Add($btnExport)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = '205,228'; $progress.Size = '740,20'
$progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Gotowe."; $lblStatus.AutoSize = $true; $lblStatus.Location = '955,230'
$lblStatus.Anchor = 'Top,Right'
$form.Controls.Add($lblStatus)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = '15,260'; $grid.Size = '1060,390'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.Anchor = 'Top,Bottom,Left,Right'
$grid.RowHeadersVisible = $false
$grid.AutoGenerateColumns = $true
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
$grid.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None
$grid.GridColor = [System.Drawing.Color]::FromArgb(233, 233, 235)
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$grid.RowTemplate.Height = 22
$grid.EnableHeadersVisualStyles = $false
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grid.DefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(250, 250, 250)
$null = $grid.GetType().GetProperty('DoubleBuffered', [System.Reflection.BindingFlags] 'NonPublic,Instance').SetValue($grid, $true, $null)
# Grid is added to layout container later

# Zmienna na wyniki
$script:lastResults = @()

# Podsumowanie wyników w etykiecie (różne/razem)
$totalCount = ($script:lastResults | Measure-Object).Count
$diffCount  = ($script:lastResults | Where-Object { $_.Status -notlike 'Same*' } | Measure-Object).Count
$lblStatus.Text = "Zakończono. Różne: $diffCount / Razem: $totalCount"
function Update-Grid {
    $items = $script:lastResults
    if ($null -eq $items) { $items = @() }
    if ($chkOnlyDiff.Checked) { $items = $items | Where-Object { $_.Status -notlike 'Same*' } }

    $cols = @('Path', 'Status', 'SourceSize', 'DestSize', 'SourceTime', 'DestTime', 'SourceHash', 'DestHash', 'Source', 'Destination')
    Write-DebugLog "Update-Grid(manual): items=$((@($items)).Count)"

    $grid.SuspendLayout()
    try {
        $grid.Visible = $true
        $grid.DataSource = $null
        $grid.AutoGenerateColumns = $false
        $grid.ColumnHeadersVisible = $true
        $grid.Columns.Clear()
        foreach ($name in $cols) {
            $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $col.Name = $name; $col.HeaderText = $name
            if ($name -eq 'Path') { $col.FillWeight = 300 }
            elseif ($name -eq 'Status') { $col.FillWeight = 120 }
            else { $col.FillWeight = 140 }
            [void]$grid.Columns.Add($col)
        }
        $grid.Rows.Clear()
        foreach ($it in $items) {
            $values = @()
            foreach ($name in $cols) { $values += $it.$name }
            [void]$grid.Rows.Add($values)
        }
        $grid.ClearSelection()
        $grid.BringToFront()
    }
    finally {
        $grid.ResumeLayout()
    }
    [System.Windows.Forms.Application]::DoEvents()
    $grid.Refresh()
    Write-DebugLog "Update-Grid(manual): gridCols=$($grid.Columns.Count) gridRows=$($grid.Rows.Count)"
    $btnExport.Enabled = (@($items).Count -gt 0)
}

# Kolorowanie wierszy wg Status po zbindowaniu, bezpiecznie (1x per refresh)
# Disable DataBindingComplete coloring when using manual fill
# (uncomment and adapt if needed)
# $grid.add_DataBindingComplete({})

# Double-click → otwórz plik w Explorerze (najpierw źródło, jak brak to cel)
$grid.add_CellDoubleClick({
        if ($null -eq $_) { return }
        if ($_.RowIndex -lt 0) { return }
        if ($grid.Rows.Count -le $_.RowIndex) { return }
        $row = $grid.Rows[$_.RowIndex]
        if ($null -eq $row) { return }
        $src = $row.Cells['Source'].Value
        $dst = $row.Cells['Destination'].Value
        $pathToOpen = if ($src -and (Test-Path -LiteralPath $src)) { $src } elseif ($dst -and (Test-Path -LiteralPath $dst)) { $dst } else { $null }
        if ($pathToOpen) {
            Start-Process explorer "/select,`"$pathToOpen`""
        }
    })

# Zmiana filtra
$chkOnlyDiff.add_CheckedChanged({ Update-Grid })

$btnRun.Add_Click({
        try {
            Write-DebugLog "Run: CLICK Start"
            $grid.DataSource = $null
            $btnExport.Enabled = $false
            $progress.Value = 0
            $lblStatus.Text = "Sprawdzam ścieżki…"

            $src = $txtSrc.Text.Trim()
            $dst = $txtDst.Text.Trim()
            Write-DebugLog "Run: src='$src' dst='$dst'"
            if (-not (Test-Path -LiteralPath $src)) { [System.Windows.Forms.MessageBox]::Show("Nie istnieje źródło:`n$src"); return }
            if (-not (Test-Path -LiteralPath $dst)) { [System.Windows.Forms.MessageBox]::Show("Nie istnieje cel:`n$dst"); return }

            $lblStatus.Text = "Porównuję…"; [System.Windows.Forms.Application]::DoEvents()

            $results = Compare-FoldersGuiCore -SourcePath $src -DestPath $dst `
                -UseHash:($chkUseHash.Checked) `
                -FastPrecheck:($chkPre.Checked) `
                -TrustFastPrecheck:($chkTrustPre.Checked) `
                -ProgressBar $progress -StatusLabel $lblStatus

            $script:lastResults = $results
            Write-DebugLog "Run: results=$(@($results).Count)"
            Update-Grid

            # Podsumowanie wyników w etykiecie (różne/razem)
            try {
                $totalCount = ($script:lastResults | Measure-Object).Count
                $diffCountLbl = ($script:lastResults | Where-Object { $_.Status -notlike 'Same*' } | Measure-Object).Count
                $lblStatus.Text = "Zakonczono. Roznice: $diffCountLbl / Razem: $totalCount"
            }
            catch {}

            $diffCount = ($results | Where-Object { $_.Status -notlike 'Same*' }).Count
            Write-DebugLog "Run: diffCount=$diffCount"
            $lblStatus.Text = "Zakończono. Różne: $diffCount"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Błąd: " + $_.Exception.Message)
            $lblStatus.Text = "Błąd."
        }
    })

$btnExport.Add_Click({
        if (-not $script:lastResults -or $script:lastResults.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Brak danych do eksportu."); return
        }
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = "CSV|*.csv"
        $dlg.FileName = "roznice.csv"
        if ($dlg.ShowDialog() -eq 'OK') {
            try {
                # Eksportuj to, co widać (z filtrem)
                $view = if ($chkOnlyDiff.Checked) {
                    $script:lastResults | Where-Object { $_.Status -notlike 'Same*' }
                }
                else { $script:lastResults }
                # Explicit projection to avoid CSV falling back to string Length
                $view | Select-Object Path, Status, SourceSize, DestSize, SourceTime, DestTime, SourceHash, DestHash, Source, Destination |
                Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
                [System.Windows.Forms.MessageBox]::Show("Zapisano:`n$($dlg.FileName)")
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Nie udało się zapisać: " + $_.Exception.Message)
            }
        }
    })

# --- Layout: switch to responsive panels ------------------------------------

try {
    $form.SuspendLayout()

    # Root container with clear row strategy
    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = [System.Windows.Forms.DockStyle]::Fill
    $root.AutoSize = $false
    $root.RowCount = 4
    $null = $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $root.RowStyles[3].SizeType = [System.Windows.Forms.SizeType]::Percent
    $root.RowStyles[3].Height = 100
    $null = $root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    # Paths (src/dst) grid: 3 columns (label, textbox, button)
    $paths = New-Object System.Windows.Forms.TableLayoutPanel
    $paths.Dock = [System.Windows.Forms.DockStyle]::Top
    $paths.AutoSize = $true
    $paths.ColumnCount = 3
    $null = $paths.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $paths.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $null = $paths.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $txtSrc.Dock = [System.Windows.Forms.DockStyle]::Fill
    $txtDst.Dock = [System.Windows.Forms.DockStyle]::Fill
    $btnSrc.AutoSize = $true
    $btnDst.AutoSize = $true
    $lblSrc.AutoSize = $true
    $lblDst.AutoSize = $true

    $lblSrc.Margin = New-Object System.Windows.Forms.Padding(8, 10, 8, 4)
    $txtSrc.Margin = New-Object System.Windows.Forms.Padding(8, 8, 8, 4)
    $btnSrc.Margin = New-Object System.Windows.Forms.Padding(8, 8, 8, 4)
    $lblDst.Margin = New-Object System.Windows.Forms.Padding(8, 6, 8, 10)
    $txtDst.Margin = New-Object System.Windows.Forms.Padding(8, 4, 8, 8)
    $btnDst.Margin = New-Object System.Windows.Forms.Padding(8, 4, 8, 8)

    $paths.Controls.Add($lblSrc, 0, 0)
    $paths.Controls.Add($txtSrc, 1, 0)
    $paths.Controls.Add($btnSrc, 2, 0)
    $paths.Controls.Add($lblDst, 0, 1)
    $paths.Controls.Add($txtDst, 1, 1)
    $paths.Controls.Add($btnDst, 2, 1)

    # Options grid: 2x2 to avoid overflow and keep inside window
    $options = New-Object System.Windows.Forms.TableLayoutPanel
    $options.Dock = [System.Windows.Forms.DockStyle]::Top
    $options.AutoSize = $true
    $options.ColumnCount = 2
    $options.RowCount = 2
    $null = $options.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $null = $options.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    $null = $options.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $options.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    foreach ($chk in @($chkUseHash, $chkPre, $chkTrustPre, $chkOnlyDiff)) {
        $chk.Margin = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
        $chk.Anchor = 'Left'
        $chk.AutoSize = $true
    }
    $options.Controls.Add($chkUseHash, 0, 0)
    $options.Controls.Add($chkPre, 1, 0)
    $options.Controls.Add($chkTrustPre, 0, 1)
    $options.Controls.Add($chkOnlyDiff, 1, 1)

    # Actions row: Start | Export | [Progress fills] | Status
    $actions = New-Object System.Windows.Forms.TableLayoutPanel
    $actions.Dock = [System.Windows.Forms.DockStyle]::Top
    $actions.AutoSize = $true
    $actions.ColumnCount = 4
    $null = $actions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $actions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $null = $actions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $null = $actions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $btnRun.AutoSize = $true
    $btnExport.AutoSize = $true
    $progress.Dock = [System.Windows.Forms.DockStyle]::Fill
    $lblStatus.AutoSize = $true
    $lblStatus.Margin = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)

    $actions.Controls.Add($btnRun, 0, 0)
    $actions.Controls.Add($btnExport, 1, 0)
    $actions.Controls.Add($progress, 2, 0)
    $actions.Controls.Add($lblStatus, 3, 0)

    # Results grid fills remaining space (inside a bordered panel for visibility)
    $resultsPanel = New-Object System.Windows.Forms.Panel
    $resultsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $resultsPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $resultsPanel.Margin = New-Object System.Windows.Forms.Padding(8, 6, 8, 8)
    $grid.MinimumSize = New-Object System.Drawing.Size(200, 150)
    $grid.ColumnHeadersVisible = $true
    $grid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $resultsPanel.Controls.Add($grid)

    # Build final layout using a single TableLayoutPanel (4 rows)
    # Rows: 0=paths (Auto), 1=options (Auto), 2=actions (Auto), 3=results (Fill)
    $root.Controls.Add($paths, 0, 0)
    $root.Controls.Add($options, 0, 1)
    $root.Controls.Add($actions, 0, 2)
    $root.Controls.Add($resultsPanel, 0, 3)
    $resultsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill

    # Apply to form
    $form.Controls.Clear()
    $form.Controls.Add($root)
}
finally {
    $form.ResumeLayout()
}

# Ensure grid is visible with headers before scan
try {
    $grid.Visible = $true
    $grid.ColumnHeadersVisible = $true
    $empty = New-ResultsDataTable -Items @()
    $grid.DataSource = $empty
    $grid.BringToFront()
}
catch {}

[void]$form.ShowDialog()
