Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

# --- Narzędzia ---------------------------------------------------------------

function Get-FileHashSafe {
    param([string]$Path)
    # SHA256 bez wczytywania całego pliku do pamięci
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
        $hashBytes = $sha.ComputeHash($stream)
        ($hashBytes | ForEach-Object { $_.ToString("x2") }) -join ''
    } catch {
        "ERROR:" + $_.Exception.Message
    } finally {
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
    $rootClean = (Get-Item -LiteralPath $Root).FullName.TrimEnd('\','/')
    $files = Get-ChildItem -LiteralPath $rootClean -Recurse -File -ErrorAction Stop
    $index = @{}
    $total = [math]::Max(1, $files.Count)
    $i = 0
    foreach ($f in $files) {
        $rel = $f.FullName.Substring($rootClean.Length).TrimStart('\','/')
        $index[$rel] = [pscustomobject]@{
            Rel              = $rel
            FullName         = $f.FullName
            Length           = $f.Length
            LastWriteTimeUtc = $f.LastWriteTimeUtc
        }
        $i++
        if ($ProgressBar) {
            $ProgressBar.Value = [int](($i/$total)*100)
        }
        if ($StatusLabel) {
            $StatusLabel.Text = "Indeksuję: $i / $total"
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    return $index
}

function Compute-FastFolderSignature {
    param([hashtable]$Index) # index: Rel -> obj(Length, LastWriteTimeUtc)
    # Lekki podpis: nazwapliku|rozmiar|czas -> SHA256 (niepewne, ale szybkie)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $sb  = New-Object System.Text.StringBuilder
    foreach ($k in ($Index.Keys | Sort-Object)) {
        $item = $Index[$k]
        [void]$sb.AppendLine(('{0}|{1}|{2:o}' -f $k, $item.Length, $item.LastWriteTimeUtc))
    }
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($sb.ToString())
    $hashBytes = $sha.ComputeHash($bytes)
    $sha.Dispose()
    ($hashBytes | ForEach-Object { $_.ToString("x2") }) -join ''
}

function Compare-FoldersGuiCore {
    param(
        [string]$SourcePath,
        [string]$DestPath,
        [switch]$UseHash,
        [switch]$FastPrecheck,
        [switch]$TrustFastPrecheck, # jeżeli true i fast precheck zgodny → kończymy (NIE 100% pewne)
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel
    )

    # Krok 1: indeksy
    $StatusLabel.Text = "Indeksuję źródło…"
    [System.Windows.Forms.Application]::DoEvents()
    $src = Get-FilesIndex -Root $SourcePath -ProgressBar $ProgressBar -StatusLabel $StatusLabel

    $StatusLabel.Text = "Indeksuję cel…"
    [System.Windows.Forms.Application]::DoEvents()
    $dst = Get-FilesIndex -Root $DestPath -ProgressBar $ProgressBar -StatusLabel $StatusLabel

    if ($FastPrecheck) {
        $StatusLabel.Text = "Szybki wstępny test (rozmiar+czas)…"
        [System.Windows.Forms.Application]::DoEvents()
        $sigSrc = Compute-FastFolderSignature -Index $src
        $sigDst = Compute-FastFolderSignature -Index $dst
        if ($sigSrc -eq $sigDst) {
            if ($TrustFastPrecheck) {
                return ,([pscustomobject]@{
                    Path='(cały katalog)'
                    Status='LikelySame (fast precheck)'
                    SourceSize=$null; DestSize=$null
                    SourceTime=$null; DestTime=$null
                    Source=$SourcePath; Destination=$DestPath
                })
            } else {
                # kontynuujemy do pełnego hashowania dla pewności
                $StatusLabel.Text = "Precheck zgodny → weryfikuję hashami plików…"
                [System.Windows.Forms.Application]::DoEvents()
            }
        } else {
            $StatusLabel.Text = "Precheck różny → pełne porównanie…"
            [System.Windows.Forms.Application]::DoEvents()
        }
    }

    # Krok 2: porównanie
    $allKeys = @($src.Keys + $dst.Keys) | Sort-Object -Unique
    $total   = [math]::Max(1, $allKeys.Count)
    $idx     = 0
    $rows    = New-Object System.Collections.Generic.List[object]

    foreach ($rel in $allKeys) {
        $idx++
        if ($ProgressBar) { $ProgressBar.Value = [int](($idx/$total)*100) }
        if ($StatusLabel) { $StatusLabel.Text = "Porównuję: $idx / $total — $rel"; [System.Windows.Forms.Application]::DoEvents() }

        $s = $src[$rel]; $d = $dst[$rel]

        if (-not $s) {
            $rows.Add([pscustomobject]@{
                Path=$rel; Status='OnlyInDestination'
                SourceSize=$null; DestSize=$d.Length
                SourceTime=$null; DestTime=$d.LastWriteTimeUtc
                Source=$null; Destination=$d.FullName
            })
            continue
        }
        if (-not $d) {
            $rows.Add([pscustomobject]@{
                Path=$rel; Status='OnlyInSource'
                SourceSize=$s.Length; DestSize=$null
                SourceTime=$s.LastWriteTimeUtc; DestTime=$null
                Source=$s.FullName; Destination=$null
            })
            continue
        }

        $sameMeta = ($s.Length -eq $d.Length) -and ($s.LastWriteTimeUtc -eq $d.LastWriteTimeUtc)
        $same = $sameMeta

        if ($UseHash) {
            # Hashy używamy zawsze (pewność). Jeżeli metadane identyczne, i tak sprawdzamy treść.
            $sHash = Get-FileHashSafe -Path $s.FullName
            $dHash = Get-FileHashSafe -Path $d.FullName
            $same  = ($sHash -eq $dHash)
        }

        $rows.Add([pscustomobject]@{
            Path=$rel
            Status= if($same){ 'Same' } else { if($sameMeta) { 'Different (content)' } else { 'Different (meta/content)' } }
            SourceSize=$s.Length; DestSize=$d.Length
            SourceTime=$s.LastWriteTimeUtc; DestTime=$d.LastWriteTimeUtc
            Source=$s.FullName; Destination=$d.FullName
        })
    }

    return $rows
}

# --- GUI ---------------------------------------------------------------------

$form                 = New-Object System.Windows.Forms.Form
$form.Text            = "Porównanie folderów (hash domyślnie)"
$form.Size            = New-Object System.Drawing.Size(1000,650)
$form.StartPosition   = "CenterScreen"

$lblSrc = New-Object System.Windows.Forms.Label
$lblSrc.Text = "Folder źródłowy:"
$lblSrc.Location = New-Object System.Drawing.Point(15,15)
$lblSrc.AutoSize = $true
$form.Controls.Add($lblSrc)

$txtSrc = New-Object System.Windows.Forms.TextBox
$txtSrc.Location = New-Object System.Drawing.Point(15,35)
$txtSrc.Size = New-Object System.Drawing.Size(820,25)
$form.Controls.Add($txtSrc)

$btnSrc = New-Object System.Windows.Forms.Button
$btnSrc.Text = "Wybierz…"
$btnSrc.Location = New-Object System.Drawing.Point(845,33)
$btnSrc.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Wybierz folder źródłowy"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtSrc.Text = $dlg.SelectedPath }
})
$form.Controls.Add($btnSrc)

$lblDst = New-Object System.Windows.Forms.Label
$lblDst.Text = "Folder docelowy (np. udział sieciowy):"
$lblDst.Location = New-Object System.Drawing.Point(15,70)
$lblDst.AutoSize = $true
$form.Controls.Add($lblDst)

$txtDst = New-Object System.Windows.Forms.TextBox
$txtDst.Location = New-Object System.Drawing.Point(15,90)
$txtDst.Size = New-Object System.Drawing.Size(820,25)
$form.Controls.Add($txtDst)

$btnDst = New-Object System.Windows.Forms.Button
$btnDst.Text = "Wybierz…"
$btnDst.Location = New-Object System.Drawing.Point(845,88)
$btnDst.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Wybierz folder docelowy"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtDst.Text = $dlg.SelectedPath }
})
$form.Controls.Add($btnDst)

# Opcje
$chkUseHash = New-Object System.Windows.Forms.CheckBox
$chkUseHash.Text = "Użyj hashy (SHA-256) — PEWNE porównanie"
$chkUseHash.Checked = $true
$chkUseHash.Location = New-Object System.Drawing.Point(15,128)
$chkUseHash.AutoSize = $true
$form.Controls.Add($chkUseHash)

$chkPre = New-Object System.Windows.Forms.CheckBox
$chkPre.Text = "Szybki wstępny test (rozmiar+czas) — niepewny, ale szybki"
$chkPre.Checked = $false
$chkPre.Location = New-Object System.Drawing.Point(15,152)
$chkPre.AutoSize = $true
$form.Controls.Add($chkPre)

$chkTrustPre = New-Object System.Windows.Forms.CheckBox
$chkTrustPre.Text = "Zatrzymaj po wstępnym teście, jeśli zgodny (NIE 100% pewne)"
$chkTrustPre.Checked = $false
$chkTrustPre.Location = New-Object System.Drawing.Point(35,176)
$chkTrustPre.AutoSize = $true
$form.Controls.Add($chkTrustPre)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Start"
$btnRun.Location = New-Object System.Drawing.Point(15,210)
$form.Controls.Add($btnRun)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Eksportuj CSV…"
$btnExport.Location = New-Object System.Drawing.Point(100,210)
$btnExport.Enabled = $false
$form.Controls.Add($btnExport)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(200,213)
$progress.Size = New-Object System.Drawing.Size(660,20)
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Gotowe."
$lblStatus.AutoSize = $true
$lblStatus.Location = New-Object System.Drawing.Point(870,213)
$form.Controls.Add($lblStatus)

# Grid wyników
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(15,250)
$grid.Size = New-Object System.Drawing.Size(950,340)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AutoSizeColumnsMode = 'Fill'
$form.Controls.Add($grid)

# Zmienna na wyniki
$script:lastResults = @()

$btnRun.Add_Click({
    try {
        $grid.DataSource = $null
        $btnExport.Enabled = $false
        $progress.Value = 0
        $lblStatus.Text = "Sprawdzam ścieżki…"

        $src = $txtSrc.Text.Trim()
        $dst = $txtDst.Text.Trim()
        if (-not (Test-Path -LiteralPath $src)) { [System.Windows.Forms.MessageBox]::Show("Nie istnieje źródło: `n$src"); return }
        if (-not (Test-Path -LiteralPath $dst)) { [System.Windows.Forms.MessageBox]::Show("Nie istnieje cel: `n$dst"); return }

        $lblStatus.Text = "Porównuję…"
        [System.Windows.Forms.Application]::DoEvents()

        $results = Compare-FoldersGuiCore -SourcePath $src -DestPath $dst `
            -UseHash:($chkUseHash.Checked) `
            -FastPrecheck:($chkPre.Checked) `
            -TrustFastPrecheck:($chkTrustPre.Checked) `
            -ProgressBar $progress -StatusLabel $lblStatus

        $script:lastResults = $results
        $grid.DataSource = $results
        $btnExport.Enabled = $true

        $lblStatus.Text = "Zakończono. Różne: " + (($results | Where-Object { $_.Status -notlike 'Same*' }).Count)
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
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:lastResults | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Zapisano: `n" + $dlg.FileName)
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Nie udało się zapisać: " + $_.Exception.Message)
        }
    }
})

[void]$form.ShowDialog()
