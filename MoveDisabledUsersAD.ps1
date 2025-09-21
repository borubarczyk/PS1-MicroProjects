#Requires -Modules ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===================== PICKER OU =====================

<#
.SYNOPSIS
Wyświetla okno wyboru jednostki organizacyjnej (OU) w Active Directory.

.DESCRIPTION
Buduje drzewo OU z katalogu AD i zwraca DN wybranego OU. Może opcjonalnie
ograniczyć drzewo do wskazanego `-SearchBase` i wymusić wybór potomka
(`-RequireUnderBase`). Obsługuje szybkie filtrowanie, Enter/podwójny klik,
preselekcję po nazwie oraz zwraca `$null` przy anulowaniu.

.PARAMETER Title
Tytuł okna dialogowego.

.PARAMETER SearchBase
DN OU, od którego budowane jest drzewo (włącznie). Gdy pominięty, budowane jest pełne drzewo domeny.

.PARAMETER RequireUnderBase
Wymusza wybór OU będącego potomkiem `-SearchBase` (nie sam `-SearchBase`).

.PARAMETER PreselectByName
Nazwa OU do wstępnego zaznaczenia (dopasowanie bez rozróżniania wielkości liter).

.OUTPUTS
System.String. Zwraca DN wybranego OU lub `$null` jeśli anulowano.

.EXAMPLE
$dn = Show-OUChooser -Title 'Wybierz OU' -SearchBase 'OU=Users,DC=contoso,DC=local' -RequireUnderBase

.NOTES
Wymaga: RSAT/ActiveDirectory, .NET WinForms. Zalecana sesja PowerShell w STA.
#>
function Show-OUChooser {
    param(
        [string]$Title = "Wybierz OU",
        [string]$SearchBase,          # jeżeli podasz: pokaże drzewo od tego OU (wraz z nim)
        [switch]$RequireUnderBase,    # wymuś wybór POD SearchBase (nie samo SearchBase)
        [string]$PreselectByName      # np. "NIEAKTYWNI"
    )

    # Reset znaczników diagnostycznych
    $script:OUChooser_LastError = $null
    $script:OUChooser_LastDialogResult = $null
    # --- UI ---
    $form               = New-Object System.Windows.Forms.Form
    $form.Text          = $Title
    $form.StartPosition = 'CenterParent'
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.Font          = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.MinimumSize   = New-Object System.Drawing.Size(640,700)
    $form.Size          = New-Object System.Drawing.Size(720,760)

    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'; $root.ColumnCount = 1; $root.RowCount = 4
    [void]$root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent,100))
    [void]$root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Filtr nazwy OU:"; $lblFilter.AutoSize = $true; $lblFilter.Padding = '8,8,8,0'

    $tbFilter = New-Object System.Windows.Forms.TextBox
    $tbFilter.Dock = 'Top'; $tbFilter.Margin = '8,0,8,8'; $tbFilter.Width = 400

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Dock = 'Fill'; $tree.HideSelection = $false
    $tree.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelButtons.Dock='Top'; $panelButtons.AutoSize=$true
    $panelButtons.FlowDirection=[System.Windows.Forms.FlowDirection]::RightToLeft
    $panelButtons.Padding='8,8,8,8'; $panelButtons.WrapContents=$false

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text="OK"; $btnOK.AutoSize=$false; $btnOK.Width=120; $btnOK.Height=32; $btnOK.Margin='6,0,0,0'
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text="Anuluj"; $btnCancel.AutoSize=$false; $btnCancel.Width=120; $btnCancel.Height=32

    $panelButtons.Controls.Add($btnOK) | Out-Null
    $panelButtons.Controls.Add($btnCancel) | Out-Null

    $root.Controls.Add($lblFilter,0,0)
    $root.Controls.Add($tbFilter,0,1)
    $root.Controls.Add($tree,0,2)
    $root.Controls.Add($panelButtons,0,3)
    $form.Controls.Add($root)
    $form.AcceptButton = $btnOK; $form.CancelButton = $btnCancel

    # --- Dane AD -> Drzewo ---
    function Get-ParentDn([string]$dn) { if ($dn -and $dn.Contains(',')) { return $dn.Substring($dn.IndexOf(',')+1) } else { return $null } }

    try {
        $ouList = @()
        $baseOu = $null

        if ($SearchBase) {
            $baseOu = Get-ADOrganizationalUnit -Identity $SearchBase -Properties Name,DistinguishedName -ErrorAction Stop
            $ouList += $baseOu
            $ouList += Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase -SearchScope Subtree -ErrorAction Stop
        } else {
            $ouList = Get-ADOrganizationalUnit -Filter * -SearchScope Subtree -ErrorAction Stop
        }

        if (-not $ouList -or $ouList.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Nie znaleziono żadnych OU do wyświetlenia.","Uwaga",'OK','Information') | Out-Null
            return $null
        }

        $nodes = @{}
        foreach ($ou in $ouList) {
            if (-not $ou.DistinguishedName) { continue }
            $nodes[$ou.DistinguishedName] = New-Object System.Windows.Forms.TreeNode -Property @{
                Text = $ou.Name
                Tag  = $ou.DistinguishedName
            }
        }

        foreach ($ou in $ouList) {
            $dn = [string]$ou.DistinguishedName
            $parentDn = Get-ParentDn $dn
            if ($parentDn -and $nodes.ContainsKey($parentDn)) {
                [void]$nodes[$parentDn].Nodes.Add($nodes[$dn])
            }
        }

        $tree.BeginUpdate()
        $tree.Nodes.Clear()

        if ($SearchBase -and $baseOu -and $nodes.ContainsKey($baseOu.DistinguishedName)) {
            [void]$tree.Nodes.Add($nodes[$baseOu.DistinguishedName])
            if ($tree.Nodes.Count -gt 0) { $tree.Nodes[0].Expand() }
        } else {
            # Bez SearchBase: wyświetl tylko te, których rodzic nie jest OU (czyli top-level pod DC=...)
            foreach ($kv in $nodes.GetEnumerator()) {
                $dn = $kv.Key
                $parentDn = Get-ParentDn $dn
                if (-not $nodes.ContainsKey($parentDn)) { [void]$tree.Nodes.Add($kv.Value) }
            }
            foreach ($n in $tree.Nodes) { $n.Expand() }
        }
        $tree.EndUpdate()
    } catch {
        $script:OUChooser_LastError = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Nie udało się pobrać OU: $($_.Exception.Message)","Błąd",'OK','Error') | Out-Null
        return $null
    }

    # UX skróty
    $tree.Add_NodeMouseDoubleClick({ param($s,$e) if ($e -and $e.Node) { $tree.SelectedNode = $e.Node; $btnOK.PerformClick() } })
    $tree.Add_DoubleClick({ if ($tree.SelectedNode) { $btnOK.PerformClick() } })
    $tree.Add_KeyDown({ param($s,$e) if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter -and $tree.SelectedNode) { $e.Handled=$true; $btnOK.PerformClick() } })

    # Preselect
    if ($PreselectByName) {
        $stack = New-Object System.Collections.Stack
        foreach ($rootNode in $tree.Nodes) { $stack.Push($rootNode) }
        $found = $null
        while ($stack.Count -gt 0 -and -not $found) {
            $node = [System.Windows.Forms.TreeNode]$stack.Pop()
            if ($node.Text -ieq $PreselectByName) { $found = $node; break }
            foreach ($c in $node.Nodes) { $stack.Push($c) }
        }
        if ($found) { $tree.SelectedNode = $found; $found.EnsureVisible() }
    }

    # Filtrowanie
    $tbFilter.Add_TextChanged({
        $q = $tbFilter.Text
        if ([string]::IsNullOrWhiteSpace($q)) {
            foreach ($n in $tree.Nodes) { $n.ExpandAll() }
        } else {
            function ShowMatches($node){
                $isMatch = $node.Text -like "*$q*"
                $childHas = $false
                foreach ($c in $node.Nodes) { if (ShowMatches $c) { $childHas = $true } }
                if ($isMatch -or $childHas) { $node.Expand(); return $true } else { $node.Collapse(); return $false }
            }
            foreach ($n in $tree.Nodes) { ShowMatches $n | Out-Null }
        }
    })

    # OK -> zwrot DN z Tag
    $selectedDN = $null
    $btnOK.Add_Click({
        if (-not $tree.SelectedNode -or -not $tree.SelectedNode.Tag) {
            [System.Windows.Forms.MessageBox]::Show("Wybierz konkretne OU z drzewa.","Uwaga",'OK','Information') | Out-Null
            return
        }
        $dn = [string]$tree.SelectedNode.Tag

        if ($RequireUnderBase -and $SearchBase) {
            # Musi być POTOMEK (czyli DN kończy się na ,SearchBase) i nie może być równe SearchBase
            $escBase = [regex]::Escape($SearchBase)
            $isDesc = ($dn -imatch ",$escBase$") -and ($dn -ine $SearchBase)
            if (-not $isDesc) {
                [System.Windows.Forms.MessageBox]::Show("Wybierz OU, które jest POTOMKIEM bazowego OU.","Uwaga",'OK','Information') | Out-Null
                return
            }
        }

        $selectedDN = $dn
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $btnCancel.Add_Click({ $script:OUChooser_LastError = $null;  $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })

    $dialogResult = $form.ShowDialog()
    $script:OUChooser_LastDialogResult = $dialogResult
    if (-not $selectedDN -and $dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($tree.SelectedNode -and $tree.SelectedNode.Tag) {
            $selectedDN = [string]$tree.SelectedNode.Tag
        }
    }
    return $selectedDN
}


# ===================== GŁÓWNE OKNO =====================
function Show-MoveDisabledUsersGUI {
    <#
    .SYNOPSIS
    Interakcyjny GUI do przenoszenia zablokowanych użytkowników AD do wybranego OU.

    .DESCRIPTION
    Okno pozwala:
    - wybrać bazowe OU (np. „NIEAKTYWNI”),
    - przeskanować katalog i znaleźć zablokowane konta spoza bazowego OU,
    - zaznaczyć wybrane konta (z przyciskami „Zaznacz/Odznacz wszystkie”),
    - wybrać dowolne docelowe OU,
    - przenieść konta (z opcją „WhatIf” – suchy bieg, bez zmian).
    W logu pojawiają się informacje o działaniu, błędach i podsumowaniu.

    .NOTES
    Wymaga: modułu ActiveDirectory (RSAT), dostępu do kontrolera domeny,
    uprawnień do przenoszenia obiektów oraz środowiska z Windows Forms.
    Zalecana sesja PowerShell w STA (`powershell.exe -STA`).

    .EXAMPLE
    Show-MoveDisabledUsersGUI
    Uruchamia GUI aplikacji.
    
    .CONFIGURATION
    Ostatnio uzywane OU sa zapisywane do ukrytego pliku '.move-disabled-users-ad.config.json'
    w katalogu skryptu. Plik jest tworzony automatycznie i aktualizowany
    po wyborze bazowego/docelowego OU oraz przy zamknieciu okna.
    Jesli zapis w katalogu skryptu nie powiedzie sie (np. blokada Nextcloud/AV),
    konfiguracja zostanie zapisana do %LOCALAPPDATA%\MoveDisabledUsersAD\config.json,
    a lokalizacja zostanie zarejestrowana w logu. Zastosowano mechanizm retry i zapis
    przez plik tymczasowy w celu ograniczenia bledow dostepu.
    #>

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.Forms.MessageBox]::Show("Brak modułu ActiveDirectory. Zainstaluj RSAT / AD PowerShell Tools.","Błąd",'OK','Error') | Out-Null
        return
    }
    Import-Module ActiveDirectory -ErrorAction Stop

    $form              = New-Object System.Windows.Forms.Form
    $form.Text         = "Przenoszenie zablokowanych użytkowników AD"
    $form.StartPosition= 'CenterScreen'
    $form.AutoScaleMode= [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.Font         = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.MinimumSize  = New-Object System.Drawing.Size(1000,760)
    $form.Size         = New-Object System.Drawing.Size(1100,820)

    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'
    $root.ColumnCount = 1
    $root.RowCount = 6
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,60)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,40)) )

    # ---- BAZA (wiersz 0) ----
    $pBase = New-Object System.Windows.Forms.TableLayoutPanel
    $pBase.Dock='Fill'; $pBase.ColumnCount=3; $pBase.AutoSize=$true
    $pBase.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    $pBase.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)) )
    $pBase.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )

    $lblBase = New-Object System.Windows.Forms.Label
    $lblBase.Text = "Bazowe OU (drzewo NIEAKTYWNI):"
    $lblBase.AutoSize = $true

    $tbBase = New-Object System.Windows.Forms.TextBox
    $tbBase.ReadOnly = $true
    $tbBase.Dock = 'Fill'
    $tbBase.Margin = '0,0,8,0'

    $btnPickBase = New-Object System.Windows.Forms.Button
    $btnPickBase.Text = "Wybierz bazowe…"
    $btnPickBase.AutoSize = $false
    $btnPickBase.Width  = 140
    $btnPickBase.Height = 32

    $pBase.Controls.Add($lblBase,0,0)
    $pBase.Controls.Add($tbBase,1,0)
    $pBase.Controls.Add($btnPickBase,2,0)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Text = "Status: brak bazowego OU"
    $lblStatus.AutoSize = $true
    $lblStatus.Padding = '0,4,0,8'

    $pBaseWrap = New-Object System.Windows.Forms.FlowLayoutPanel
    $pBaseWrap.Dock='Fill'; $pBaseWrap.AutoSize=$false; $pBaseWrap.FlowDirection='TopDown'
    $pBaseWrap.WrapContents = $false
    $pBaseWrap.Controls.Add($pBase) | Out-Null
    $pBaseWrap.Controls.Add($lblStatus) | Out-Null

    # ---- SKAN (wiersz 1) ----
    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = "Skanuj zablokowanych (spoza bazowego OU)"
    $btnScan.AutoSize = $false
    $btnScan.Width  = 360
    $btnScan.Height = 32
    $btnScan.Margin = '8,4,0,8'
    $btnScan.Enabled  = $false

    # ---- GRID (wiersz 2) ----
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = 'Fill'
    $grid.ReadOnly = $false
    $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $true
    $grid.AutoGenerateColumns = $false
    $grid.RowHeadersVisible = $false
    $grid.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $grid.AutoSizeColumnsMode = 'AllCells'

    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.HeaderText = "Wybierz"
    $colCheck.Width = 70
    $grid.Columns.Add($colCheck) | Out-Null

    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = "Name"
    $colName.DataPropertyName = "Name"
    $colName.AutoSizeMode = 'Fill'
    $grid.Columns.Add($colName) | Out-Null

    $colSam = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSam.HeaderText = "SamAccountName"
    $colSam.DataPropertyName = "SamAccountName"
    $colSam.Width = 180
    $grid.Columns.Add($colSam) | Out-Null

    $colDn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colDn.HeaderText = "DistinguishedName"
    $colDn.DataPropertyName = "DistinguishedName"
    $colDn.AutoSizeMode = 'Fill'
    $grid.Columns.Add($colDn) | Out-Null

    # ---- TARGET (wiersz 3) ----
    $pTarget = New-Object System.Windows.Forms.TableLayoutPanel
    $pTarget.Dock='Top'; $pTarget.ColumnCount=3; $pTarget.AutoSize=$true
    $pTarget.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    $pTarget.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)) )
    $pTarget.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )

    $lblTarget = New-Object System.Windows.Forms.Label
    $lblTarget.Text = "Docelowe OU:"
    $lblTarget.AutoSize = $true

    $tbTarget = New-Object System.Windows.Forms.TextBox
    $tbTarget.ReadOnly = $true
    $tbTarget.Dock = 'Fill'

    $btnPickTarget = New-Object System.Windows.Forms.Button
    $btnPickTarget.Text= "Wybierz docelowe…"
    $btnPickTarget.AutoSize = $false
    $btnPickTarget.Width  = 140
    $btnPickTarget.Height = 32
    $btnPickTarget.Enabled = $false

    $pTarget.Controls.Add($lblTarget,0,0)
    $pTarget.Controls.Add($tbTarget,1,0)
    $pTarget.Controls.Add($btnPickTarget,2,0)

    # ---- AKCJE (wiersz 4) ----
    $pActions = New-Object System.Windows.Forms.FlowLayoutPanel
    $pActions.Dock='Top'
    $pActions.AutoSize = $true
    $pActions.WrapContents = $false
    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = "Zaznacz wszystkie"
    $btnSelectAll.AutoSize = $false
    $btnSelectAll.Width = 160
    $btnSelectAll.Height = 32

    $btnClearAll = New-Object System.Windows.Forms.Button
    $btnClearAll.Text = "Odznacz wszystkie"
    $btnClearAll.AutoSize = $false
    $btnClearAll.Width = 160
    $btnClearAll.Height = 32
    $btnClearAll.Margin = '6,0,0,0'

    $chkWhatIf = New-Object System.Windows.Forms.CheckBox
    $chkWhatIf.Text = "WhatIf (suchy bieg – bez zmian)"
    $chkWhatIf.AutoSize = $true
    $chkWhatIf.Checked  = $true   # DOMYŚLNIE ZAZNACZONE

    $btnMove = New-Object System.Windows.Forms.Button
    $btnMove.Text = "Przenieś zaznaczonych"
    $btnMove.AutoSize = $false
    $btnMove.Width  = 220
    $btnMove.Height = 32
    $btnMove.Margin = '12,0,0,0'

    $pActions.Controls.Add($btnSelectAll) | Out-Null
    $pActions.Controls.Add($btnClearAll) | Out-Null
    $pActions.Controls.Add($chkWhatIf)   | Out-Null
    $pActions.Controls.Add($btnMove)     | Out-Null

    # Handlery zaznacz/odznacz wszystkie
    $btnSelectAll.Add_Click({
        foreach ($r in $grid.Rows) { try { $r.Cells[0].Value = $true } catch {} }
    })
    $btnClearAll.Add_Click({
        foreach ($r in $grid.Rows) { try { $r.Cells[0].Value = $false } catch {} }
    })

    # ---- LOG (wiersz 5) ----
    $tbLog = New-Object System.Windows.Forms.TextBox
    $tbLog.Multiline = $true
    $tbLog.ScrollBars = 'Both'
    $tbLog.ReadOnly = $true
    $tbLog.Dock = 'Fill'
    $tbLog.Font = New-Object System.Drawing.Font("Consolas",10)

    # Złożenie
    $root.Controls.Add($pBaseWrap,0,0)
    $root.Controls.Add($btnScan,  0,1)
    $root.Controls.Add($grid,     0,2)
    $root.Controls.Add($pTarget,  0,3)
    $root.Controls.Add($pActions, 0,4)
    $root.Controls.Add($tbLog,    0,5)
    $form.Controls.Add($root)

    # ===== Helpery =====
    function Get-ConfigPath {
        try {
            $scriptPath = if ($PSCommandPath) { $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
            $scriptDir  = Split-Path -Parent $scriptPath
            return (Join-Path $scriptDir '.move-disabled-users-ad.config.json')
        } catch { return (Join-Path (Get-Location).Path '.move-disabled-users-ad.config.json') }
    }

    function Get-AppDataConfigPath {
        try {
            $dir = Join-Path $env:LOCALAPPDATA 'MoveDisabledUsersAD'
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            return (Join-Path $dir 'config.json')
        } catch { return (Join-Path (Get-Location).Path 'config.json') }
    }

    function Get-TempConfigPath {
        try {
            $dir = Join-Path $env:TEMP 'MoveDisabledUsersAD'
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            return (Join-Path $dir 'config.json')
        } catch { return (Join-Path (Get-Location).Path 'config.tmp.json') }
    }

    function Load-Config {
        $primary  = Get-ConfigPath
        $fallback = Get-AppDataConfigPath
        $tempPath = Get-TempConfigPath
        foreach ($path in @($primary,$fallback,$tempPath)) {
            if (-not (Test-Path -LiteralPath $path)) { continue }
            try {
                $raw = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
                if ([string]::IsNullOrWhiteSpace($raw)) { continue }
                return ($raw | ConvertFrom-Json -ErrorAction Stop)
            } catch {
                $tbLog.AppendText("Uwaga: nieudane wczytanie konfiguracji z $($path): $($_.Exception.Message)`r`n")
                continue
            }
        }
        return $null
    }

    function Save-Config([string]$BaseOU,[string]$TargetOU){
        $obj = [PSCustomObject]@{
            BaseOU   = $BaseOU
            TargetOU = $TargetOU
            LastUsed = (Get-Date).ToString('o')
            Version  = '1.0'
        }
        $json = $obj | ConvertTo-Json -Depth 3
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)

        function Write-Robust([string]$path){
            $dir = Split-Path -Parent $path
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $maxAttempts = 5; $attempt = 0
            while ($attempt -lt $maxAttempts) {
                $attempt++
                $tmp = Join-Path $dir ('.' + [IO.Path]::GetFileName($path) + '.' + [guid]::NewGuid().ToString('N') + '.tmp')
                try {
                    [System.IO.File]::WriteAllText($tmp,$json,$utf8NoBom)
                    if (Test-Path -LiteralPath $path) {
                        try { (Get-Item -LiteralPath $path -Force).Attributes = 'Normal' } catch {}
                        try {
                            [System.IO.File]::Replace($tmp, $path, $null, $true)
                        } catch {
                            $bak = Join-Path $dir ([IO.Path]::GetFileName($path) + '.' + [guid]::NewGuid().ToString('N') + '.bak')
                            try {
                                try { (Get-Item -LiteralPath $path -Force).Attributes = 'Normal' } catch {}
                                Move-Item -LiteralPath $path -Destination $bak -Force
                                Move-Item -LiteralPath $tmp -Destination $path -Force
                                Remove-Item -LiteralPath $bak -Force -ErrorAction SilentlyContinue
                            } catch {
                                try { [System.IO.File]::Copy($tmp,$path,$true) } finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
                            }
                        }
                    } else {
                        try { Move-Item -LiteralPath $tmp -Destination $path -Force } finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
                    }
                    try { (Get-Item -LiteralPath $path).Attributes = ((Get-Item -LiteralPath $path).Attributes -bor [System.IO.FileAttributes]::Hidden) } catch {}
                    return $true
                } catch {
                    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
                    if ($attempt -lt $maxAttempts) { Start-Sleep -Milliseconds (200 * $attempt) } else { return $false }
                }
            }
            return $false
        }

        if (Write-Robust (Get-ConfigPath)) { return }
        if (Write-Robust (Get-AppDataConfigPath)) { $tbLog.AppendText("Zapisano konfiguracje w lokalizacji uzytkownika: " + (Get-AppDataConfigPath) + "`r`n"); return }
        if (Write-Robust (Get-TempConfigPath)) { $tbLog.AppendText("Zapisano konfiguracje w lokalizacji tymczasowej: " + (Get-TempConfigPath) + "`r`n"); return }
        $tbLog.AppendText("Uwaga: nie udalo sie zapisac konfiguracji w zadnej lokalizacji.`r`n")
    }function Get-DisabledUsersOutsideOU([string]$BaseOU){
        Get-ADUser -Filter 'Enabled -eq $false' -Properties DistinguishedName,SamAccountName,Name |
            Where-Object { $_.DistinguishedName -notlike "*$BaseOU*" } |
            Select-Object Name,SamAccountName,DistinguishedName |
            Sort-Object Name
    }

    # Konfiguracja lub autowybór startowy
    $cfg = Load-Config
    if ($cfg -and $cfg.BaseOU) {
        $tbBase.Text = [string]$cfg.BaseOU
        $lblStatus.Text = "Bazowe OU: $($tbBase.Text)"
        $btnScan.Enabled = $true
        $btnPickTarget.Enabled = $true
        if ($cfg.TargetOU) { $tbTarget.Text = [string]$cfg.TargetOU }
        if ([string]::IsNullOrWhiteSpace($tbTarget.Text)) {
            $tbLog.AppendText("Wczytano konfiguracje: bazowe=" + $tbBase.Text + "`r`n")
        } else {
            $tbLog.AppendText("Wczytano konfiguracje: bazowe=" + $tbBase.Text + ", docelowe=" + $tbTarget.Text + "`r`n")
        }
    } else {
        try {
            $maybe = Get-ADOrganizationalUnit -LDAPFilter '(ou=NIEAKTYWNI)' -SearchScope Subtree -ErrorAction Stop |
                     Select-Object -First 1 -ExpandProperty DistinguishedName
            if ($maybe) {
                $tbBase.Text = $maybe
                $lblStatus.Text = "Bazowe OU: $maybe"
                $btnScan.Enabled = $true
                $btnPickTarget.Enabled = $true
                $tbLog.AppendText("Autowybrano bazowe OU: $maybe`r`n")
            Save-Config -BaseOU $tbBase.Text -TargetOU $tbTarget.Text
            }
        } catch {}
    }

    # Wybór bazowego OU
    $btnPickBase.Add_Click({
        $sel = Show-OUChooser -Title "Wybierz BAZOWE OU (np. NIEAKTYWNI)" -PreselectByName "NIEAKTYWNI"
        if ($sel) {
            $tbBase.Text = $sel
            $lblStatus.Text = "Bazowe OU: $sel"
            $tbTarget.Text = ""
            $btnScan.Enabled = $true
            $btnPickTarget.Enabled = $true
            $tbLog.AppendText("Wybrano bazowe OU: $sel`r`n")
            Save-Config -BaseOU $tbBase.Text -TargetOU $tbTarget.Text
        } else {
            if ($script:OUChooser_LastError) { $tbLog.AppendText("Blad wyboru bazowego OU: $script:OUChooser_LastError (DialogResult: $script:OUChooser_LastDialogResult)`r`n") } else { $tbLog.AppendText("Anulowano wybór bazowego OU. (DialogResult: $script:OUChooser_LastDialogResult)`r`n") }
        }
    })

    # Skan
    $btnScan.Add_Click({
        $grid.Rows.Clear()
        $tbLog.AppendText("Skanowanie…`r`n")
        $base = $tbBase.Text
        if ([string]::IsNullOrWhiteSpace($base)) {
            [System.Windows.Forms.MessageBox]::Show("Wybierz najpierw BAZOWE OU.","Uwaga",'OK','Information') | Out-Null
            $tbLog.AppendText("Błąd: brak bazowego OU.`r`n")
            return
        }
        try {
            $users = Get-DisabledUsersOutsideOU -BaseOU $base
            foreach ($u in $users) {
                $idx = $grid.Rows.Add()
                $grid.Rows[$idx].Cells[0].Value = $false
                $grid.Rows[$idx].Cells[1].Value = $u.Name
                $grid.Rows[$idx].Cells[2].Value = $u.SamAccountName
                $grid.Rows[$idx].Cells[3].Value = $u.DistinguishedName
            }
            $tbLog.AppendText("Znaleziono: $($grid.Rows.Count) rekordów.`r`n")
            if ($grid.Rows.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Brak zablokowanych użytkowników SPOZA bazowego OU.","Info",'OK','Information') | Out-Null
            }
        } catch {
            $tbLog.AppendText("Błąd skanowania: $($_.Exception.Message)`r`n")
            [System.Windows.Forms.MessageBox]::Show("Błąd skanowania: $($_.Exception.Message)","Błąd",'OK','Error') | Out-Null
        }
    })

    # Wybór docelowego OU
    $btnPickTarget.Add_Click({
        $base = $tbBase.Text
        if ([string]::IsNullOrWhiteSpace($base)) {
            [System.Windows.Forms.MessageBox]::Show("Najpierw wybierz bazowe OU.","Uwaga",'OK','Information') | Out-Null
            $tbLog.AppendText("Błąd: próba wyboru docelowego bez bazowego OU.`r`n")
            return
        }
        $sel = Show-OUChooser -Title "Wybierz DOCELOWE OU (dowolne)"
        if ($sel) {
            $tbTarget.Text = $sel
            $tbLog.AppendText("Wybrano docelowe OU: $sel`r`n")
            Save-Config -BaseOU $tbBase.Text -TargetOU $tbTarget.Text
        } else {
            if ($script:OUChooser_LastError) { $tbLog.AppendText("Blad wyboru docelowego OU: $script:OUChooser_LastError (DialogResult: $script:OUChooser_LastDialogResult)`r`n") } else { $tbLog.AppendText("Anulowano wybór docelowego OU. (DialogResult: $script:OUChooser_LastDialogResult)`r`n") }
        }
    })

    # Przenoszenie
    $btnMove.Add_Click({
        $base   = $tbBase.Text
        $target = $tbTarget.Text
        if ([string]::IsNullOrWhiteSpace($base))   { [System.Windows.Forms.MessageBox]::Show("Wybierz bazowe OU.","Uwaga",'OK','Information') | Out-Null; $tbLog.AppendText("Błąd: brak bazowego OU.`r`n"); return }
        if ([string]::IsNullOrWhiteSpace($target)) { [System.Windows.Forms.MessageBox]::Show("Wybierz docelowe OU.","Uwaga",'OK','Information') | Out-Null; $tbLog.AppendText("Błąd: brak docelowego OU.`r`n"); return }

        # Zbieranie zaznaczonych
        $selected = @()
        foreach ($r in $grid.Rows) {
            if ($r.Cells[0].Value -eq $true) {
                $selected += [PSCustomObject]@{
                    Name = $r.Cells[1].Value
                    Sam  = $r.Cells[2].Value
                    DN   = $r.Cells[3].Value
                }
            }
        }
        if (-not $selected) {
            [System.Windows.Forms.MessageBox]::Show("Nie zaznaczono żadnych użytkowników.","Uwaga",'OK','Information') | Out-Null
            $tbLog.AppendText("Brak zaznaczonych do przeniesienia.`r`n")
            return
        }

        $whatIf = $chkWhatIf.Checked
        $ok = [System.Windows.Forms.MessageBox]::Show(
            ("Potwierdź przeniesienie {0} użytkownik(ów){1}." -f $selected.Count, $(if($whatIf){" (WhatIf)"}else{""})),
            "Potwierdzenie",'OKCancel','Question'
        )
        if ($ok -ne 'OK') { $tbLog.AppendText("Przerwano przez użytkownika.`r`n"); return }

        $moved = 0; $errors = 0
        foreach ($u in $selected) {
            try {
                if ($u.DN -like "*$target*") {
                    $tbLog.AppendText("Pominięto (już w docelowym): $($u.Name)`r`n"); continue
                }
                if ($whatIf) {
                    $tbLog.AppendText("WhatIf: $($u.Name) -> $target`r`n")
                } else {
                    Move-ADObject -Identity $u.DN -TargetPath $target -ErrorAction Stop
                    $tbLog.AppendText("Przeniesiono: $($u.Name) -> $target`r`n")
                }
                $moved++
            } catch {
                $errors++
                $tbLog.AppendText("Błąd dla $($u.Name): $($_.Exception.Message)`r`n")
            }
        }

        [System.Windows.Forms.MessageBox]::Show(("Zakończono. Sukcesów: {0}, błędów: {1}{2}" -f $moved,$errors, $(if($whatIf){" (WhatIf)"}else{""})),"Raport",'OK','Information') | Out-Null
    })

    $form.Add_Shown({ $form.Activate() }) | Out-Null
    $form.Add_FormClosing({ Save-Config -BaseOU $tbBase.Text -TargetOU $tbTarget.Text }) | Out-Null
    [void]$form.ShowDialog()
}

# ===================== START =====================
Show-MoveDisabledUsersGUI

<#
.SYNOPSIS
Przenoszenie zablokowanych użytkowników AD do wybranego OU – interfejs graficzny.

.DESCRIPTION
Skrypt uruchamia okno WinForms do zarządzania przenoszeniem zablokowanych
użytkowników AD. Umożliwia wybór bazowego OU, skan kont spoza tego OU,
zaznaczanie elementów (także masowo), wybór dowolnego OU docelowego
i przeniesienie z opcją „WhatIf”.

.REQUIREMENTS
- Windows PowerShell z RSAT/ActiveDirectory,
- dostęp do kontrolera domeny (AD Web Services),
- .NET Windows Forms, sesja STA zalecana,
- odpowiednie uprawnienia do przenoszenia obiektów AD.

.EXAMPLE
.\\Przenoszenie zablokowanych użytkowników AD.ps1
Uruchamia narzędzie GUI.

.NOTES
Autor: (uzupełnij)
Wersja: 1.0
Historia: Dodano wybór dowolnego docelowego OU, przyciski zaznacz/odznacz wszystkie,
diagnozę błędów wyboru OU, poprawiono rozmiary i marginesy UI.
.CONFIGURATION
Ostatnio uzywane OU sa zapisywane do ukrytego pliku '.move-disabled-users-ad.config.json'
w katalogu skryptu. Plik jest tworzony automatycznie i aktualizowany
po wyborze bazowego/docelowego OU oraz przy zamknieciu okna.
Skrypt stosuje mechanizm retry oraz zapis przez plik tymczasowy,
aby ograniczyć problemy z blokadami (np. aplikacje sync/AV).
#>
