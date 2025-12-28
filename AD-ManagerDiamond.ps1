#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ======= KONFIGURACJA APLIKACJI =======
$App = [ordered]@{
    Title            = 'Domain Ops - Remote Admin GUI'
    Width            = 1500
    Height           = 820
    DefaultSharePath = 'C$\Windows\Temp\DomainOps' # gdzie kopiujemy instalatory
    Modules          = @()  # kontenery rejestracji modułów (zakładek)
}

# ======= NARZĘDZIA WSPÓLNE =======
function New-FixedColumn {
    param(
        [string]$text,
        [int]$width
    )
    (New-Object System.Windows.Forms.ColumnHeader -Property @{Text = $text; Width = $width }) 
}

function Show-Error([string]$msg, [Exception]$ex = $null) {
    [System.Windows.Forms.MessageBox]::Show(($msg + (if ($ex) "`r`n`r`n$($ex.Message)" else '')), 'Błąd', 'OK', 'Error') | Out-Null
}

function Format-Bytes([long]$bytes) {
    if ($bytes -lt 1KB) { return "$bytes B" }
    elseif ($bytes -lt 1MB) { return "{0:N2} KB" -f ($bytes / 1KB) }
    elseif ($bytes -lt 1GB) { return "{0:N2} MB" -f ($bytes / 1MB) }
    else { return "{0:N2} GB" -f ($bytes / 1GB) }
}

# Globalny stan (poświadczenia, sesje CIM, itp.)
$State = [ordered]@{
    Cred            = $null
    UseCurrentCreds = $true
    CimSessions     = @{}   # ComputerName -> CimSession
    AdComputers     = @()   # cache wyników z AD
}

# Prosty logger do okienka na dole
$script:CurrentModule = $null

function Write-Log([string]$msg, [string]$level = 'INFO') {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $modulePart = if ($script:CurrentModule) { "[Module=$($script:CurrentModule)] " } else { '' }
    $line = "[{0}] [{1}] {2}{3}" -f $ts, $level, $modulePart, $msg
    $Global:txtLog.AppendText($line + [Environment]::NewLine)
}

function Invoke-ModuleAction {
    param(
        [string]$ModuleName,
        [scriptblock]$Action
    )
    $prev = $script:CurrentModule
    try {
        $script:CurrentModule = $ModuleName
        & $Action
    }
    finally {
        $script:CurrentModule = $prev
    }
}

# Pobieranie komputerów z AD
function Get-AdComputersUI {
    try {
        if (-not (Get-Module -ListAvailable ActiveDirectory)) {
            throw "Brak modułu ActiveDirectory (RSAT). Zainstaluj RSAT i spróbuj ponownie."
        }
        Import-Module ActiveDirectory -ErrorAction Stop | Out-Null

        $filter = if ($txtNameFilter.Text) { "(Name -like '*$($txtNameFilter.Text.Replace('*','').Replace('?',''))*')" } else { '*' }
        $searchBase = if ($txtSearchBase.Text.Trim()) { $txtSearchBase.Text.Trim() } else { $null }

        Write-Log "Pobieram komputery z AD (Filter=$filter; SearchBase=$searchBase)..."
        $params = @{ Filter = $filter; Properties = @('OperatingSystem', 'LastLogonDate') }
        if ($searchBase) { $params.SearchBase = $searchBase }

        $State.AdComputers = @(Get-ADComputer @params | Sort-Object Name)
        $clbComputers.Items.Clear()
        foreach ($c in $State.AdComputers) {
            [void]$clbComputers.Items.Add($c.Name)
        }
        Write-Log "Załadowano: $($State.AdComputers.Count) komputerów."
    }
    catch {
        Show-Error "Nie udało się pobrać komputerów z AD." $_
        Write-Log "AD błąd: $($_.Exception.Message)" 'ERROR'
    }
}

# Zwraca wybrane komputery z listy
function Get-SelectedComputers {
    $list = @()
    foreach ($idx in $clbComputers.CheckedIndices) {
        $list += $clbComputers.Items[$idx]
    }
    return @($list | Sort-Object -Unique)
}

# Tworzy/odświeża sesje CIM do wskazanych komputerów
function Connect-Cim([string[]]$Computers, [switch]$UseDCOM) {
    $created = 0
    foreach ($c in $Computers) {
        if ($State.CimSessions.ContainsKey($c)) { continue }
        try {
            $opt = if ($UseDCOM) { New-CimSessionOption -Protocol DCOM } else { New-CimSessionOption -Protocol Wsman }
            if ($State.UseCurrentCreds -or -not $State.Cred) {
                $s = New-CimSession -ComputerName $c -SessionOption $opt -ErrorAction Stop
            }
            else {
                $s = New-CimSession -ComputerName $c -Credential $State.Cred -SessionOption $opt -ErrorAction Stop
            }
            $State.CimSessions[$c] = $s
            $created++
        }
        catch {
            Write-Log "CIM do $c nie powiodła się: $($_.Exception.Message)" 'WARN'
        }
    }
    if ($created -gt 0) { Write-Log "Utworzono $created nowych sesji CIM." }
}

function Close-Cim([string[]]$Computers) {
    foreach ($c in $Computers) {
        if ($State.CimSessions.ContainsKey($c)) {
            try { $State.CimSessions[$c] | Remove-CimSession -ErrorAction Stop } catch {}
            $State.CimSessions.Remove($c) | Out-Null
        }
    }
}

# Invoke-Command helper
# --- PATCH: poprawiona wersja Invoke-Remote (obsługa hashtable wg nazw parametrów) ---
function Invoke-Remote {
    param(
        [string]$ComputerName,
        [scriptblock]$ScriptBlock,
        $Arg
    )
    $p = @{
        ComputerName = $ComputerName
        ScriptBlock  = $ScriptBlock
        ErrorAction  = 'Stop'
    }
    if (-not $State.UseCurrentCreds -and $State.Cred) { $p.Credential = $State.Cred }

    $argsList = @()
    if ($null -ne $Arg) {
        if ($Arg -is [hashtable]) {
            $paramBlock = $ScriptBlock.Ast.ParamBlock
            if ($paramBlock) {
                foreach ($param in $paramBlock.Parameters) {
                    $name = $param.Name.VariablePath.UserPath
                    $argsList += $Arg[$name]
                }
            }
            else {
                $argsList = @($Arg)
            }
        }
        elseif ($Arg -is [object[]]) {
            $argsList = $Arg
        }
        else {
            $argsList = @($Arg)
        }
    }
    return Invoke-Command @p -ArgumentList $argsList
}
# --- KONIEC PATCHA ---


# ======= BUDOWA GUI =======
$form = New-Object System.Windows.Forms.Form
$form.Text = $App.Title
$form.Width = $App.Width
$form.Height = $App.Height
$form.StartPosition = 'CenterScreen'

# Glowny layout: naglowek, przestrzen robocza, log
$layoutRoot = New-Object System.Windows.Forms.TableLayoutPanel
$layoutRoot.Dock = 'Fill'
$layoutRoot.RowCount = 3
$layoutRoot.ColumnCount = 1
$null = $layoutRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 52)))
$null = $layoutRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $layoutRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 140)))
$null = $layoutRoot.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($layoutRoot)

# GORA: panel poswiadczen
$panelTop = New-Object System.Windows.Forms.FlowLayoutPanel
$panelTop.Height = 60
$panelTop.Dock = 'Fill'
$panelTop.BackColor = [System.Drawing.Color]::FromArgb(245, 245, 245)
$panelTop.WrapContents = $false
$panelTop.FlowDirection = 'LeftToRight'
$panelTop.AutoScroll = $true
$panelTop.Padding = '10,14,10,10'
$layoutRoot.Controls.Add($panelTop, 0, 0)

$chkUseCurrent = New-Object System.Windows.Forms.CheckBox
$chkUseCurrent.Text = 'Uzyj biezacych poswiadczen'
$chkUseCurrent.Checked = $true
$chkUseCurrent.AutoSize = $true
$chkUseCurrent.Margin = New-Object System.Windows.Forms.Padding(0, 0, 20, 0)
$panelTop.Controls.Add($chkUseCurrent)

$btnCred = New-Object System.Windows.Forms.Button
$btnCred.Text = 'Zmien poswiadczenia...'
$btnCred.Left = 240
$btnCred.Top = 12
$btnCred.Width = 160
$btnCred.Margin = New-Object System.Windows.Forms.Padding(0, 0, 20, 0)
$panelTop.Controls.Add($btnCred)

$lblCred = New-Object System.Windows.Forms.Label
$lblCred.Text = '(biezacy uzytkownik)'
$lblCred.AutoSize = $true
$lblCred.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
$panelTop.Controls.Add($lblCred)

# Panel glowny (split: lewy AD, prawy TabControl)
$splitMain = New-Object System.Windows.Forms.SplitContainer
$splitMain.Dock = 'Fill'
$splitMain.SplitterDistance = 420
$splitMain.Panel1MinSize = 360
$splitMain.Orientation = 'Vertical'
$layoutRoot.Controls.Add($splitMain, 0, 1)

$form.Add_Shown({
        $splitMain.Panel2MinSize = 700
        $maxDistance = $splitMain.Width - $splitMain.Panel2MinSize
        if ($maxDistance -lt $splitMain.Panel1MinSize) {
            $splitMain.Panel2MinSize = [Math]::Max(200, $splitMain.Width - ($splitMain.Panel1MinSize + 10))
            $maxDistance = $splitMain.Width - $splitMain.Panel2MinSize
        }
        $splitMain.SplitterDistance = [Math]::Max(
            $splitMain.Panel1MinSize,
            [Math]::Min($maxDistance, 480)
        )
    })

# Lewy panel: AD
$grpAd = New-Object System.Windows.Forms.GroupBox
$grpAd.Text = 'Active Directory — Komputery'
$grpAd.Dock = 'Fill'
$splitMain.Panel1.Controls.Add($grpAd)

$lblSearchBase = New-Object System.Windows.Forms.Label
$lblSearchBase.Text = 'SearchBase (OU, opcjonalnie):'
$lblSearchBase.Left = 12; $lblSearchBase.Top = 24; $lblSearchBase.AutoSize = $true
$grpAd.Controls.Add($lblSearchBase)

$txtSearchBase = New-Object System.Windows.Forms.TextBox
$txtSearchBase.Left = 12; $txtSearchBase.Top = 44; $txtSearchBase.Width = 330
$txtSearchBase.Anchor = 'Top,Left,Right'
$grpAd.Controls.Add($txtSearchBase)

$lblNameFilter = New-Object System.Windows.Forms.Label
$lblNameFilter.Text = 'Filtr nazwy (wildcard *):'
$lblNameFilter.Left = 12; $lblNameFilter.Top = 74; $lblNameFilter.AutoSize = $true
$grpAd.Controls.Add($lblNameFilter)

$txtNameFilter = New-Object System.Windows.Forms.TextBox
$txtNameFilter.Left = 12; $txtNameFilter.Top = 94; $txtNameFilter.Width = 210
$txtNameFilter.Anchor = 'Top,Left,Right'
$grpAd.Controls.Add($txtNameFilter)

$btnLoadAD = New-Object System.Windows.Forms.Button
$btnLoadAD.Text = 'Załaduj'
$btnLoadAD.Left = 230; $btnLoadAD.Top = 92; $btnLoadAD.Width = 110
$btnLoadAD.Anchor = 'Top,Right'
$grpAd.Controls.Add($btnLoadAD)

$clbComputers = New-Object System.Windows.Forms.CheckedListBox
$clbComputers.Left = 12; $clbComputers.Top = 130; $clbComputers.Width = 330; $clbComputers.Height = 500
$clbComputers.CheckOnClick = $true
$clbComputers.Anchor = 'Top,Bottom,Left,Right'
$grpAd.Controls.Add($clbComputers)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = 'Zaznacz wszystko'
$btnSelectAll.Left = 12; $btnSelectAll.Top = 640; $btnSelectAll.Width = 150
$btnSelectAll.Anchor = 'Bottom,Left'
$grpAd.Controls.Add($btnSelectAll)

$btnClearSel = New-Object System.Windows.Forms.Button
$btnClearSel.Text = 'Wyczyść zaznaczenie'
$btnClearSel.Left = 192; $btnClearSel.Top = 640; $btnClearSel.Width = 150
$btnClearSel.Anchor = 'Bottom,Left'
$grpAd.Controls.Add($btnClearSel)

$grpAd.Add_Resize({
        $margin = 24
        $availWidth = [Math]::Max(220, $grpAd.ClientSize.Width - $margin)
        $txtSearchBase.Width = $availWidth
        $nameWidth = [Math]::Max(150, $availWidth - $btnLoadAD.Width - 16)
        $txtNameFilter.Width = $nameWidth
        $btnLoadAD.Left = $txtNameFilter.Left + $nameWidth + 8
        $clbComputers.Width = $availWidth
        $clbComputers.Height = [Math]::Max(120, $grpAd.ClientSize.Height - 200)
        $btnSelectAll.Top = $grpAd.ClientSize.Height - 42
        $btnClearSel.Top = $btnSelectAll.Top
    })

# Prawy panel: TabControl (moduły)
$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$splitMain.Panel2.Controls.Add($tabs)

# Dol: log
$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.ReadOnly = $true
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.Dock = 'Fill'
$txtLog.Height = 140
$layoutRoot.Controls.Add($txtLog, 0, 2)
$Global:txtLog = $txtLog

# ======= REAKCJE UI (poświadczenia, AD) =======
$chkUseCurrent.add_CheckedChanged({
        $State.UseCurrentCreds = $chkUseCurrent.Checked
        if ($State.UseCurrentCreds) {
            $lblCred.Text = "(bieżący użytkownik: $env:USERDOMAIN\$env:USERNAME)"
        }
        else {
            $lblCred.Text = if ($State.Cred) { "Używane: $($State.Cred.UserName)" } else { "(brak — ustaw poświadczenia)" }
        }
    })

$btnCred.Add_Click({
        try {
            $cred = Get-Credential -Message 'Poświadczenia do zdalnych operacji'
            if ($cred) {
                $State.Cred = $cred
                $chkUseCurrent.Checked = $false
                $lblCred.Text = "Używane: $($cred.UserName)"
                Write-Log "Ustawiono poświadczenia $($cred.UserName)."
            }
        }
        catch {
            Show-Error "Nie udało się pobrać poświadczeń." $_
        }
    })

$btnLoadAD.Add_Click({ Get-AdComputersUI })
$btnSelectAll.Add_Click({
        for ($i = 0; $i -lt $clbComputers.Items.Count; $i++) { $clbComputers.SetItemChecked($i, $true) }
    })
$btnClearSel.Add_Click({
        for ($i = 0; $i -lt $clbComputers.Items.Count; $i++) { $clbComputers.SetItemChecked($i, $false) }
    })

# ======= INFRA: rejestracja modułów (każdy moduł = zakładka) =======
function Register-ModuleTab {
    param(
        [string]$Name,
        [scriptblock]$Builder # dostaje $tab,$getTargets
    )
    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $Name
    $tabs.TabPages.Add($tab)
    # Pomocnicza funkcja do pobierania hostów
    $getTargets = {
        $targetsRaw = Get-SelectedComputers
        $targets = @($targetsRaw)
        if ($targets.Count -eq 0 -or ($targets.Count -eq 1 -and [string]::IsNullOrWhiteSpace([string]$targets[0]))) {
            Show-Error "Wybierz najpierw komputery z listy po lewej."
            throw "Brak host�w"
        }
        $targets
    }
    $script:getTargets = $getTargets
    & $Builder $tab $getTargets
    $App.Modules += $Name
}

# ======= MODUŁY (zakładki) =======

# 1) Wiersz poleceń (PS/cmd)
Register-ModuleTab -Name 'Polecenia' -Builder {
    param($tab, $getTargets)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Polecenia do uruchomienia (PowerShell lub CMD):'
    $lbl.Left = 12; $lbl.Top = 12; $lbl.AutoSize = $true
    $tab.Controls.Add($lbl)

    $rbPS = New-Object System.Windows.Forms.RadioButton
    $rbPS.Text = 'PowerShell'
    $rbPS.Checked = $true
    $rbPS.Left = 12; $rbPS.Top = 36
    $tab.Controls.Add($rbPS)

    $rbCmd = New-Object System.Windows.Forms.RadioButton
    $rbCmd.Text = 'cmd.exe'
    $rbCmd.Left = 120; $rbCmd.Top = 36
    $tab.Controls.Add($rbCmd)

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.Left = 12; $txt.Top = 64; $txt.Width = 870; $txt.Height = 120
    $txt.Font = New-Object System.Drawing.Font('Consolas', 10)
    $tab.Controls.Add($txt)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Uruchom na zaznaczonych'
    $btnRun.Left = 900; $btnRun.Top = 64; $btnRun.Width = 220; $btnRun.Height = 34
    $tab.Controls.Add($btnRun)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 200; $grid.Width = 1110; $grid.Height = 470
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btnRun.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $rbPS = $ctx.Controls.rbPS
                $txt = $ctx.Controls.txt
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    $cmd = $txt.Text.Trim()
                    if (-not $cmd) { Show-Error "Wpisz polecenia."; return }
                    $rows = @()
                    foreach ($t in $targets) {
                        Write-Log "[$t] uruchamiam polecenia..."
                        if ($rbPS.Checked) {
                            $sb = [scriptblock]::Create($cmd)
                            $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        }
                        else {
                            $sb = { param($c) cmd.exe /c $c 2>&1 }
                            $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb -Arg @{ c = $cmd }
                        }
                        $text = ($out | Out-String).Trim()
                        $rows += [pscustomobject]@{Komputer = $t; Wynik = $text }
                    }
                    $grid.DataSource = $rows
                    Write-Log "Zakończono."
                }
                catch {
                    Write-Log $_.Exception.Message 'ERROR'
                }
            }
        })
}

# 2) BitLocker (manage-bde -status, backup do AD)
Register-ModuleTab -Name 'BitLocker' -Builder {
    param($tab, $getTargets)

    $btnStatus = New-Object System.Windows.Forms.Button
    $btnStatus.Text = 'Sprawdz status'
    $btnStatus.Left = 12; $btnStatus.Top = 12; $btnStatus.Width = 160
    $tab.Controls.Add($btnStatus)

    $btnBackup = New-Object System.Windows.Forms.Button
    $btnBackup.Text = 'Backup kluczy do AD (C:)'
    $btnBackup.Left = 190; $btnBackup.Top = 12; $btnBackup.Width = 220
    $tab.Controls.Add($btnBackup)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btnStatus.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    $rows = @()
                    foreach ($t in $targets) {
                        Write-Log "[$t] sprawdzam BitLocker..."
                        $sb = {
                            $vols = @()
                            if (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue) {
                                foreach ($vol in Get-BitLockerVolume) {
                                    $vols += [pscustomobject]@{
                                        Volume           = $vol.MountPoint
                                        ConversionStatus = $vol.VolumeStatus
                                        Protection       = $vol.ProtectionStatus
                                        Lock             = $vol.LockStatus
                                        Version          = $vol.EncryptionMethod
                                    }
                                }
                            }
                            if (-not $vols) {
                                $exe = Join-Path $env:SystemRoot 'System32\manage-bde.exe'
                                if (-not (Test-Path $exe)) { throw 'manage-bde.exe is not available na hoscie.' }
                                $text = (& $exe -status 2>&1) | Out-String
                                $current = @{}
                                foreach ($line in $text -split "`r?`n") {
                                    if ($line -match 'Volume [A-Z]:') {
                                        if ($current.Count -gt 0) { $vols += [pscustomobject]$current; $current = @{} }
                                        $current.Volume = ($line -replace 'Volume ', '').Trim()
                                    }
                                    if ($line -match '^\s*Conversion Status:\s*(.+)$') { $current.ConversionStatus = $matches[1].Trim() }
                                    if ($line -match '^\s*BitLocker Version:\s*(.+)$') { $current.Version = $matches[1].Trim() }
                                    if ($line -match '^\s*Protection Status:\s*(.+)$') { $current.Protection = $matches[1].Trim() }
                                    if ($line -match '^\s*Lock Status:\s*(.+)$') { $current.Lock = $matches[1].Trim() }
                                }
                                if ($current.Count -gt 0) { $vols += [pscustomobject]$current }
                            }
                            $vols
                        }
                        $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        if (-not $out) {
                            $rows += [pscustomobject]@{Komputer = $t; Wolumin = '(brak informacji)'; Ochrona = 'brak danych'; Konwersja = '-'; Lock = '-'; Wersja = '-' }
                        }
                        else {
                            foreach ($v in $out) {
                                $rows += [pscustomobject]@{
                                    Komputer  = $t
                                    Wolumin   = $v.Volume
                                    Ochrona   = $v.Protection
                                    Konwersja = $v.ConversionStatus
                                    Lock      = $v.Lock
                                    Wersja    = $v.Version
                                }
                            }
                        }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })

    $btnBackup.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                try {
                    $targets = & $getTargets
                    foreach ($t in $targets) {
                        Write-Log "[$t] backup kluczy C: do AD..."
                        $sb = {
                            try {
                                if (Get-Command Backup-BitLockerKeyProtector -ErrorAction SilentlyContinue) {
                                    $vol = Get-BitLockerVolume -MountPoint 'C:'
                                    $kps = $vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }
                                    if (-not $kps) { return 'Brak RecoveryPassword do backupu.' }
                                    foreach ($kp in $kps) {
                                        Backup-BitLockerKeyProtector -MountPoint 'C:' -KeyProtectorId $kp.KeyProtectorId | Out-Null
                                    }
                                    'Backup-BitLockerKeyProtector wykonany.'
                                }
                                else {
                                    $exe = Join-Path $env:SystemRoot 'System32\manage-bde.exe'
                                    if (-not (Test-Path $exe)) { throw 'manage-bde.exe is not available na hoscie.' }
                                    $output = (& $exe -protectors -get C: -Type RecoveryPassword) -join "`n"
                                    $ids = [regex]::Matches($output, 'ID:\s*({[^}]+})') | ForEach-Object { $_.Groups[1].Value }
                                    if ($ids.Count -eq 0) { return 'Nie znaleziono RecoveryPassword.' }
                                    foreach ($id in $ids) {
                                        & $exe -protectors -adbackup C: -id $id | Out-Null
                                    }
                                    'manage-bde -protectors -adbackup wykonany.'
                                }
                            }
                            catch { $_.Exception.Message }
                        }
                        $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        Write-Log "[$t] $out"
                    }
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# 3) Dyski (WMI: Win32_LogicalDisk)
Register-ModuleTab -Name 'Dyski' -Builder {
    param($tab, $getTargets)

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = 'Odśwież'
    $btnRefresh.Left = 12; $btnRefresh.Top = 12; $btnRefresh.Width = 120
    $tab.Controls.Add($btnRefresh)

    $btnClean = New-Object System.Windows.Forms.Button
    $btnClean.Text = 'Wyczyść TEMP + Kosz'
    $btnClean.Left = 150; $btnClean.Top = 12; $btnClean.Width = 180
    $tab.Controls.Add($btnClean)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true
    $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btnRefresh.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    Connect-Cim -Computers $targets
                    $rows = @()
                    foreach ($t in $targets) {
                        if (-not $State.CimSessions.ContainsKey($t)) { continue }
                        $s = $State.CimSessions[$t]
                        $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -CimSession $s
                        foreach ($d in $disks) {
                            $rows += [pscustomobject]@{
                                Komputer     = $t
                                Dysk         = $d.DeviceID
                                SystemPlikow = $d.FileSystem
                                Wolne        = (Format-Bytes $d.FreeSpace)
                                Rozmiar      = (Format-Bytes $d.Size)
                                Zajetosc     = [math]::Round((($d.Size - $d.FreeSpace) / $d.Size) * 100, 1)
                            }
                        }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })

    $btnClean.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                try {
                    $targets = & $getTargets
                    foreach ($t in $targets) {
                        Write-Log "[$t] czyszczenie TEMP i Kosza..."
                        $sb = {
                            try {
                                Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
                                if (Get-Command Clear-RecycleBin -ErrorAction SilentlyContinue) { Clear-RecycleBin -Force -ErrorAction SilentlyContinue }
                                'OK'
                            }
                            catch { $_.Exception.Message }
                        }
                        $r = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        Write-Log "[$t] wynik: $r"
                    }
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# 4) Usługi (przegląd i sterowanie)
Register-ModuleTab -Name 'Usługi' -Builder {
    param($tab, $getTargets)

    $lblHost = New-Object System.Windows.Forms.Label
    $lblHost.Text = 'Komputer (jedna nazwa):'
    $lblHost.Left = 12; $lblHost.Top = 16; $lblHost.AutoSize = $true
    $tab.Controls.Add($lblHost)

    $cmbHost = New-Object System.Windows.Forms.ComboBox
    $cmbHost.Left = 140; $cmbHost.Top = 12; $cmbHost.Width = 240; $cmbHost.DropDownStyle = 'DropDownList'
    $tab.Controls.Add($cmbHost)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Załaduj usługi'
    $btnLoad.Left = 400; $btnLoad.Top = 12; $btnLoad.Width = 140
    $tab.Controls.Add($btnLoad)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Left = 560; $txtFilter.Top = 14; $txtFilter.Width = 200
    $tab.Controls.Add($txtFilter)
    $lblF = New-Object System.Windows.Forms.Label
    $lblF.Text = 'Filtr'; $lblF.Left = 520; $lblF.Top = 16; $lblF.AutoSize = $true
    $tab.Controls.Add($lblF)

    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.Text = 'Start'
    $btnStart.Left = 780; $btnStart.Top = 12; $btnStart.Width = 100
    $tab.Controls.Add($btnStart)

    $btnStop = New-Object System.Windows.Forms.Button
    $btnStop.Text = 'Stop'
    $btnStop.Left = 890; $btnStop.Top = 12; $btnStop.Width = 100
    $tab.Controls.Add($btnStop)

    $btnRestart = New-Object System.Windows.Forms.Button
    $btnRestart.Text = 'Restart'
    $btnRestart.Left = 1000; $btnRestart.Top = 12; $btnRestart.Width = 100
    $tab.Controls.Add($btnRestart)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true
    $grid.SelectionMode = 'FullRowSelect'
    $grid.MultiSelect = $false
    $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    # Inicjalizacja listy hostów
    $tab.Add_Enter({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $combo = $ctx.Controls.cmbHost
                $combo.Items.Clear()
                foreach ($i in Get-SelectedComputers) { [void]$combo.Items.Add($i) }
                if ($combo.Items.Count -gt 0) { $combo.SelectedIndex = 0 }
            }
        })

    $btnLoad.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $combo = $ctx.Controls.cmbHost
                $gridCtrl = $ctx.Controls.grid
                $filterBox = $ctx.Controls.txtFilter
                try {
                    $selectedHost = $combo.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                    Connect-Cim -Computers @($selectedHost)
                    $s = $State.CimSessions[$selectedHost]
                    $svcs = Get-CimInstance -ClassName Win32_Service -CimSession $s | Sort-Object Name
                    if ($filterBox.Text) { $svcs = $svcs | Where-Object { $_.Name -like "*$($filterBox.Text)*" -or $_.DisplayName -like "*$($filterBox.Text)*" } }
                    $gridCtrl.DataSource = $svcs | Select-Object Name, DisplayName, State, StartMode, ProcessId
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })

    foreach ($btn in @($btnStart, $btnStop, $btnRestart)) {
        $btn.Add_Click({
                Invoke-InModuleContext -SourceControl $this -Action {
                    param($ctx)
                    $combo = $ctx.Controls.cmbHost
                    $gridCtrl = $ctx.Controls.grid
                    $btnLoad = $ctx.Controls.btnLoad
                    try {
                        $selectedHost = $combo.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                        if (-not $gridCtrl.SelectedRows) { Show-Error "Zaznacz jedna usluge."; return }
                        $name = $gridCtrl.SelectedRows[0].Cells['Name'].Value
                        Write-Log "[$selectedHost] $($this.Text) uslugi $name ..."
                        $sb = {
                            param($n, $op)
                            $svc = Get-Service -Name $n -ErrorAction Stop
                            switch ($op) {
                                'Start' { Start-Service -InputObject $svc -ErrorAction Stop; 'Started' }
                                'Stop' { Stop-Service  -InputObject $svc -Force -ErrorAction Stop; 'Stopped' }
                                'Restart' { Restart-Service -InputObject $svc -Force -ErrorAction Stop; 'Restarted' }
                            }
                        }
                        $out = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ n = $name; op = $this.Text }
                        Write-Log ("[{0}] {1}: {2}" -f $selectedHost, $name, $out)
                        $btnLoad.PerformClick()
                    }
                    catch { Write-Log $_.Exception.Message 'ERROR' }
                }
            })
    }
}

# 5) Konta lokalne (WMI: Win32_UserAccount LocalAccount=True)
Register-ModuleTab -Name 'Konta lokalne' -Builder {
    param($tab, $getTargets)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'Pokaż konta'; $btn.Left = 12; $btn.Top = 12; $btn.Width = 140
    $tab.Controls.Add($btn)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btn.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    Connect-Cim -Computers $targets
                    $rows = @()
                    foreach ($t in $targets) {
                        if (-not $State.CimSessions.ContainsKey($t)) { continue }
                        $s = $State.CimSessions[$t]
                        $users = Get-CimInstance Win32_UserAccount -CimSession $s -Filter "LocalAccount=True"
                        foreach ($u in $users) {
                            $rows += [pscustomobject]@{Komputer = $t; Nazwa = $u.Name; PełnaNazwa = $u.FullName; Włączone = (-not $u.Disabled); SID = $u.SID }
                        }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# 6) Udostępnienia (WMI: Win32_Share)
Register-ModuleTab -Name 'Udziały (shary)' -Builder {
    param($tab, $getTargets)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'Pokaż udziały'; $btn.Left = 12; $btn.Top = 12; $btn.Width = 140
    $tab.Controls.Add($btn)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btn.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    Connect-Cim -Computers $targets
                    $rows = @()
                    foreach ($t in $targets) {
                        if (-not $State.CimSessions.ContainsKey($t)) { continue }
                        $s = $State.CimSessions[$t]
                        $shares = Get-CimInstance Win32_Share -CimSession $s
                        foreach ($sh in $shares) {
                            $rows += [pscustomobject]@{Komputer = $t; Nazwa = $sh.Name; Ścieżka = $sh.Path; Typ = $sh.Type; Opis = $sh.Description }
                        }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# 7) Instalacja oprogramowania (MSI/EXE)
Register-ModuleTab -Name 'Instalacja softu' -Builder {
    param($tab, $getTargets)
    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text = 'Ścieżka do instalatora (MSI/EXE; lokalna lub UNC):'
    $lbl1.Left = 12; $lbl1.Top = 16; $lbl1.AutoSize = $true
    $tab.Controls.Add($lbl1)

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Left = 12; $txtPath.Top = 36; $txtPath.Width = 780
    $tab.Controls.Add($txtPath)

    $lblArgs = New-Object System.Windows.Forms.Label
    $lblArgs.Text = 'Argumenty ciche (np. MSI: /qn; EXE: /S, /quiet):'
    $lblArgs.Left = 12; $lblArgs.Top = 68; $lblArgs.AutoSize = $true
    $tab.Controls.Add($lblArgs)

    $txtArgs = New-Object System.Windows.Forms.TextBox
    $txtArgs.Left = 12; $txtArgs.Top = 88; $txtArgs.Width = 780
    $tab.Controls.Add($txtArgs)

    $btnInstall = New-Object System.Windows.Forms.Button
    $btnInstall.Text = 'Zainstaluj na zaznaczonych'
    $btnInstall.Left = 820; $btnInstall.Top = 36; $btnInstall.Width = 300; $btnInstall.Height = 38
    $tab.Controls.Add($btnInstall)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 140; $grid.Width = 1110; $grid.Height = 530
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btnInstall.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $txtPath = $ctx.Controls.txtPath
                $txtArgs = $ctx.Controls.txtArgs
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    $path = $txtPath.Text.Trim(); if (-not $path) { Show-Error 'Podaj ścieżkę instalatora.'; return }
                    $installArgs = $txtArgs.Text.Trim()
                    $rows = @()
                    foreach ($t in $targets) {
                        Write-Log "[$t] przygotowuję instalację..."
                        # utwórz zdalny katalog roboczy i skopiuj plik
                        $work = "C:\Windows\Temp\DomainOps"
                        $sbPrep = { param($w) if (-not (Test-Path $w)) { New-Item -Path $w -ItemType Directory -Force | Out-Null } $w }
                        Invoke-Remote -ComputerName $t -ScriptBlock $sbPrep -Arg @{ w = $work } | Out-Null

                        # Kopia pliku
                        try {
                            $sess = if ($State.UseCurrentCreds -or -not $State.Cred) { New-PSSession -ComputerName $t } else { New-PSSession -ComputerName $t -Credential $State.Cred }
                            Copy-Item -Path $path -Destination $work -ToSession $sess -Force
                            Remove-PSSession $sess
                            $file = [System.IO.Path]::GetFileName($path)
                            $remoteFile = Join-Path $work $file

                            # Uruchomienie
                            $sbRun = {
                                param($f, $a)
                                $ext = [System.IO.Path]::GetExtension($f)
                                if ($ext -ieq '.msi') {
                                    $arguments = "/i `"$f`" /qn $a"
                                    $exe = 'msiexec.exe'
                                }
                                else {
                                    $arguments = "`"$f`" $a"
                                    $exe = $f
                                }
                                $p = Start-Process -FilePath $exe -ArgumentList $arguments -PassThru -Wait
                                [pscustomobject]@{ExitCode = $p.ExitCode; File = $f }
                            }
                            $out = Invoke-Remote -ComputerName $t -ScriptBlock $sbRun -Arg @{ f = $remoteFile; a = $installArgs }
                            $rows += [pscustomobject]@{Komputer = $t; Plik = $out.File; KodWyjscia = $out.ExitCode }
                            Write-Log "[$t] zakończono instalację (kod=$($out.ExitCode))."
                        }
                        catch {
                            $rows += [pscustomobject]@{Komputer = $t; Plik = $path; KodWyjscia = 'Kopia/Start błąd'; Uwagi = $_.Exception.Message }
                            Write-Log "[$t] błąd instalacji: $($_.Exception.Message)" 'ERROR'
                        }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# ======= DODATKOWE 5 FUNKCJI „NAJCZĘŚCIEJ W DOMENIE” =======

# 8) GPUpdate (wymuszenie odświeżenia zasad)
Register-ModuleTab -Name 'GPUpdate' -Builder {
    param($tab, $getTargets)
    $btnU = New-Object System.Windows.Forms.Button
    $btnU.Text = 'Wymuś gpupdate /force'
    $btnU.Left = 12; $btnU.Top = 12; $btnU.Width = 220
    $tab.Controls.Add($btnU)

    $btnU.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                try {
                    $targets = & $getTargets
                    foreach ($t in $targets) {
                        Write-Log "[$t] gpupdate /force..."
                        # jeśli środowisko nie ma Invoke-GPUpdate, użyj gpupdate w zdalnej sesji
                        $sb = { gpupdate /force /target:computer 2>&1 | Out-String }
                        $r = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        Write-Log "[$t] GPUpdate: $((($r|Out-String).Trim()) -replace '[\r\n]+',' | ')"
                    }
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# 9) Programy zainstalowane (z rejestru Uninstall x64/x86)
Register-ModuleTab -Name 'Programy (zainstalowane)' -Builder {
    param($tab, $getTargets)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'Pobierz listę'; $btn.Left = 12; $btn.Top = 12; $btn.Width = 160
    $tab.Controls.Add($btn)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btn.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    $sb = {
                        $paths = @(
                            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
                            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
                        )
                        $items = foreach ($p in $paths) {
                            if (Test-Path $p) {
                                Get-ItemProperty $p | Where-Object { $_.DisplayName } | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
                            }
                        }
                        $items
                    }
                    $rows = @()
                    foreach ($t in $targets) {
                        Write-Log "[$t] czytam rejestr Uninstall..."
                        $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        foreach ($i in $out) {
                            $rows += [pscustomobject]@{
                                Komputer = $t; Nazwa = $i.DisplayName; Wersja = $i.DisplayVersion; Wydawca = $i.Publisher; Data = $i.InstallDate
                            }
                        }
                    }
                    $grid.DataSource = $rows | Sort-Object Komputer, Nazwa
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# 10) Windows Update (PSWindowsUpdate jeżeli dostępny)
Register-ModuleTab -Name 'Windows Update' -Builder {
    param($tab, $getTargets)
    $btnCheck = New-Object System.Windows.Forms.Button
    $btnCheck.Text = 'Skanuj dostępne aktualizacje'
    $btnCheck.Left = 12; $btnCheck.Top = 12; $btnCheck.Width = 240
    $tab.Controls.Add($btnCheck)

    $btnInstall = New-Object System.Windows.Forms.Button
    $btnInstall.Text = 'Zainstaluj (jeśli PSWindowsUpdate)'
    $btnInstall.Left = 270; $btnInstall.Top = 12; $btnInstall.Width = 260
    $tab.Controls.Add($btnInstall)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btnCheck.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    $rows = @()
                    foreach ($t in $targets) {
                        Write-Log "[$t] sprawdzam dostępność modułu PSWindowsUpdate..."
                        $sb = {
                            $has = Get-Module -ListAvailable -Name PSWindowsUpdate
                            if ($has) {
                                Import-Module PSWindowsUpdate -Force
                                try { Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -WhatIf } catch { $_ }
                            }
                            else {
                                # tryb awaryjny (tylko skan) — nieinstalacyjny
                                try { Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile -Command "UsoClient StartScan"' -PassThru | Out-Null } catch {}
                                'PSWindowsUpdate brak — wymuszono skan USOClient (jeśli wspierane).'
                            }
                        }
                        $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        $text = ($out | Out-String).Trim()
                        $rows += [pscustomobject]@{Komputer = $t; Wynik = $text }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })

    $btnInstall.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                try {
                    $targets = & $getTargets
                    foreach ($t in $targets) {
                        Write-Log "[$t] próba instalacji z PSWindowsUpdate..."
                        $sb = {
                            if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
                                Import-Module PSWindowsUpdate -Force
                                Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -AutoReboot
                                'Zlecono instalację aktualizacji.'
                            }
                            else {
                                'Brak PSWindowsUpdate na hoście — zainstaluj moduł.'
                            }
                        }
                        $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        Write-Log "[$t] $out"
                    }
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# 11) Zdarzenia (ostatnie 24h: błędy/ostrzeżenia)
Register-ModuleTab -Name 'Zdarzenia (24h)' -Builder {
    param($tab, $getTargets)
    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'Pobierz'; $btn.Left = 12; $btn.Top = 12; $btn.Width = 140
    $tab.Controls.Add($btn)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btn.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    $rows = @()
                    foreach ($t in $targets) {
                        Write-Log "[$t] czytam logi (24h, Error/Warning)..."
                        $sb = {
                            $start = (Get-Date).AddDays(-1)
                            Get-WinEvent -FilterHashtable @{ Level = 1, 2, 3; StartTime = $start } -ErrorAction SilentlyContinue |
                            Select-Object TimeCreated, Id, LevelDisplayName, ProviderName, LogName, Message -First 200
                        }
                        $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        foreach ($e in $out) {
                            $rows += [pscustomobject]@{
                                Komputer = $t; Czas = $e.TimeCreated; Poziom = $e.LevelDisplayName; ID = $e.Id; Źródło = $e.ProviderName; Log = $e.LogName; Wiadomość = $e.Message
                            }
                        }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# 12) Restart/Shutdown + Uptime
Register-ModuleTab -Name 'Zasilanie/Uptime' -Builder {
    param($tab, $getTargets)
    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = 'Pokaż uptime'; $btnUp.Left = 12; $btnUp.Top = 12; $btnUp.Width = 150
    $tab.Controls.Add($btnUp)

    $btnRe = New-Object System.Windows.Forms.Button
    $btnRe.Text = 'Restart'; $btnRe.Left = 180; $btnRe.Top = 12; $btnRe.Width = 120
    $tab.Controls.Add($btnRe)

    $btnSh = New-Object System.Windows.Forms.Button
    $btnSh.Text = 'Wyłącz'; $btnSh.Left = 310; $btnSh.Top = 12; $btnSh.Width = 120
    $tab.Controls.Add($btnSh)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btnUp.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    Connect-Cim -Computers $targets
                    $rows = @()
                    foreach ($t in $targets) {
                        if (-not $State.CimSessions.ContainsKey($t)) { continue }
                        $os = Get-CimInstance Win32_OperatingSystem -CimSession $State.CimSessions[$t]
                        $lboot = $os.LastBootUpTime
                        $uptime = (Get-Date) - $lboot
                        $rows += [pscustomobject]@{Komputer = $t; OstatniStart = $lboot; Uptime = ("{0:%d}d {0:hh}h {0:mm}m" -f $uptime) }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })

    $btnRe.Add_Click({
            try {
                $targets = & $getTargets
                foreach ($t in $targets) {
                    Write-Log "[$t] Restart-Computer..."
                    $p = @{ ComputerName = $t; Force = $true; ErrorAction = 'SilentlyContinue' }
                    if (-not $State.UseCurrentCreds -and $State.Cred) { $p.Credential = $State.Cred }
                    Restart-Computer @p
                }
            }
            catch { Write-Log $_.Exception.Message 'ERROR' }
        })

    $btnSh.Add_Click({
            try {
                $targets = & $getTargets
                foreach ($t in $targets) {
                    Write-Log "[$t] Stop-Computer..."
                    $p = @{ ComputerName = $t; Force = $true; ErrorAction = 'SilentlyContinue' }
                    if (-not $State.UseCurrentCreds -and $State.Cred) { $p.Credential = $State.Cred }
                    Stop-Computer @p
                }
            }
            catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# 13) Defender (status + szybki skan)
Register-ModuleTab -Name 'Defender' -Builder {
    param($tab, $getTargets)
    $btnS = New-Object System.Windows.Forms.Button
    $btnS.Text = 'Status'; $btnS.Left = 12; $btnS.Top = 12; $btnS.Width = 120
    $tab.Controls.Add($btnS)

    $btnQ = New-Object System.Windows.Forms.Button
    $btnQ.Text = 'Szybki skan'; $btnQ.Left = 150; $btnQ.Top = 12; $btnQ.Width = 140
    $tab.Controls.Add($btnQ)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btnS.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    $rows = @()
                    foreach ($t in $targets) {
                        $sb = { if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) { Get-MpComputerStatus } else { 'Brak modułu Defender (Get-MpComputerStatus)' } }
                        $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        if ($out -is [string]) {
                            $rows += [pscustomobject]@{Komputer = $t; Status = $out }
                        }
                        else {
                            $rows += [pscustomobject]@{Komputer = $t; Realtime = $out.RealTimeProtectionEnabled; AV = $out.AntivirusEnabled; OstatniaAktualizacja = $out.AntivirusSignatureLastUpdated }
                        }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })

    $btnQ.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                try {
                    $targets = & $getTargets
                    foreach ($t in $targets) {
                        $sb = { if (Get-Command Start-MpScan -ErrorAction SilentlyContinue) { Start-MpScan -ScanType QuickScan; 'Skan zlecony' } else { 'Brak Start-MpScan' } }
                        $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb
                        Write-Log "[$t] $out"
                    }
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
}

# 14) Konto komputera (AD/Trust): test/napraw kanału, reset w AD, włącz/wyłącz
Register-ModuleTab -Name 'Konto komputera (AD)' -Builder {
    param($tab, $getTargets)

    $lblDC = New-Object System.Windows.Forms.Label
    $lblDC.Text = 'Kontroler domeny (opcjonalnie):'
    $lblDC.Left = 12; $lblDC.Top = 16; $lblDC.AutoSize = $true
    $tab.Controls.Add($lblDC)

    $txtDC = New-Object System.Windows.Forms.TextBox
    $txtDC.Left = 200; $txtDC.Top = 12; $txtDC.Width = 220
    $tab.Controls.Add($txtDC)

    $btnTrust = New-Object System.Windows.Forms.Button
    $btnTrust.Text = 'Test+Napraw kanał (na hoście)'
    $btnTrust.Left = 440; $btnTrust.Top = 12; $btnTrust.Width = 220
    $tab.Controls.Add($btnTrust)

    $btnResetAD = New-Object System.Windows.Forms.Button
    $btnResetAD.Text = 'Reset konta w AD'
    $btnResetAD.Left = 670; $btnResetAD.Top = 12; $btnResetAD.Width = 160
    $tab.Controls.Add($btnResetAD)

    $btnEnable = New-Object System.Windows.Forms.Button
    $btnEnable.Text = 'Włącz konto (AD)'
    $btnEnable.Left = 840; $btnEnable.Top = 12; $btnEnable.Width = 140
    $tab.Controls.Add($btnEnable)

    $btnDisable = New-Object System.Windows.Forms.Button
    $btnDisable.Text = 'Wyłącz konto (AD)'
    $btnDisable.Left = 990; $btnDisable.Top = 12; $btnDisable.Width = 140
    $tab.Controls.Add($btnDisable)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btnTrust.Add_Click({
            Invoke-InModuleContext -SourceControl $this -Action {
                param($ctx)
                $txtDC = $ctx.Controls.txtDC
                $grid = $ctx.Controls.grid
                try {
                    $targets = & $getTargets
                    $dc = $txtDC.Text.Trim()
                    $rows = @()
                    foreach ($t in $targets) {
                        Write-Log "[$t] test/naprawa kanału zaufania..."
                        $sb = {
                            param($dc)
                            try {
                                $srv = if ($dc) { @{Server = $dc } } else { @{} }
                                $ok = Test-ComputerSecureChannel @srv -ErrorAction Stop
                                if (-not $ok) {
                                    Reset-ComputerMachinePassword @srv -ErrorAction Stop
                                    [pscustomobject]@{Akcja = 'Repair'; Wynik = 'Reset-ComputerMachinePassword'; Szczegoly = 'Kanał naprawiony' }
                                }
                                else {
                                    [pscustomobject]@{Akcja = 'Test'; Wynik = 'OK'; Szczegoly = 'Kanał poprawny' }
                                }
                            }
                            catch {
                                try {
                                    $srv = if ($dc) { @{Server = $dc } } else { @{} }
                                    Reset-ComputerMachinePassword @srv -ErrorAction Stop
                                    [pscustomobject]@{Akcja = 'Repair'; Wynik = 'Reset wykonany'; Szczegoly = $_.Exception.Message }
                                }
                                catch {
                                    [pscustomobject]@{Akcja = 'Repair'; Wynik = 'Błąd'; Szczegoly = $_.Exception.Message }
                                }
                            }
                        }
                        $out = Invoke-Remote -ComputerName $t -ScriptBlock $sb -Arg @{ dc = $dc }
                        foreach ($o in $out) { $rows += [pscustomobject]@{Komputer = $t; Akcja = $o.Akcja; Wynik = $o.Wynik; Szczegoly = $o.Szczegoly } }
                    }
                    $grid.DataSource = $rows
                }
                catch { Write-Log $_.Exception.Message 'ERROR' }
            }
        })
    foreach ($btnAct in @(
            @{Btn = $btnResetAD; What = 'Reset' },
            @{Btn = $btnEnable; What = 'Enable' },
            @{Btn = $btnDisable; What = 'Disable' }
        )) {
        $currentAct = $btnAct
        $currentAct.Btn.Add_Click({
                Invoke-InModuleContext -SourceControl $this -Action {
                    param($ctx)
                    $txtDC = $ctx.Controls.txtDC
                    $grid = $ctx.Controls.grid
                    try {
                        if (-not (Get-Module -ListAvailable ActiveDirectory)) { throw "Brak modulu ActiveDirectory (RSAT)." }
                        Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
                        $targets = & $getTargets
                        $dc = $txtDC.Text.Trim()
                        $rows = @()
                        foreach ($t in $targets) {
                            switch ($currentAct.What) {
                                'Reset' { $p = @{Identity = $t; ErrorAction = 'Stop' }; if ($dc) { $p.Server = $dc }; Reset-ADComputer @p; $act = 'Reset-ADComputer' }
                                'Enable' { $p = @{Identity = $t; ErrorAction = 'Stop' }; if ($dc) { $p.Server = $dc }; Enable-ADAccount @p; $act = 'Enable-ADAccount' }
                                'Disable' { $p = @{Identity = $t; ErrorAction = 'Stop' }; if ($dc) { $p.Server = $dc }; Disable-ADAccount @p; $act = 'Disable-ADAccount' }
                            }
                            $rows += [pscustomobject]@{Komputer = $t; Akcja = $act; Wynik = 'OK' }
                        }
                        $grid.DataSource = $rows
                    }
                    catch {
                        Show-Error "Operacja AD nie powiodla sie." $_
                        Write-Log $_.Exception.Message 'ERROR'
                    }
                }
            })
    }
}

# 15) Zmiana nazwy komputera (pojedynczo/wsadowo)
Register-ModuleTab -Name 'Zmiana nazwy' -Builder {
    param($tab, $getTargets)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Załaduj zaznaczone do siatki'
    $btnLoad.Left = 12; $btnLoad.Top = 12; $btnLoad.Width = 220
    $tab.Controls.Add($btnLoad)

    $lblPrefix = New-Object System.Windows.Forms.Label
    $lblPrefix.Text = 'Prefiks:'; $lblPrefix.Left = 250; $lblPrefix.Top = 16; $lblPrefix.AutoSize = $true
    $tab.Controls.Add($lblPrefix)

    $txtPrefix = New-Object System.Windows.Forms.TextBox
    $txtPrefix.Left = 300; $txtPrefix.Top = 12; $txtPrefix.Width = 140
    $tab.Controls.Add($txtPrefix)

    $lblStart = New-Object System.Windows.Forms.Label
    $lblStart.Text = 'Start #:'; $lblStart.Left = 450; $lblStart.Top = 16; $lblStart.AutoSize = $true
    $tab.Controls.Add($lblStart)

    $numStart = New-Object System.Windows.Forms.NumericUpDown
    $numStart.Left = 510; $numStart.Top = 12; $numStart.Width = 70; $numStart.Minimum = 0; $numStart.Maximum = 100000; $numStart.Value = 1
    $tab.Controls.Add($numStart)

    $lblPad = New-Object System.Windows.Forms.Label
    $lblPad.Text = 'Zerowanie (np. 3→001):'; $lblPad.Left = 590; $lblPad.Top = 16; $lblPad.AutoSize = $true
    $tab.Controls.Add($lblPad)

    $numPad = New-Object System.Windows.Forms.NumericUpDown
    $numPad.Left = 750; $numPad.Top = 12; $numPad.Width = 60; $numPad.Minimum = 1; $numPad.Maximum = 8; $numPad.Value = 3
    $tab.Controls.Add($numPad)

    $lblSuffix = New-Object System.Windows.Forms.Label
    $lblSuffix.Text = 'Sufiks:'; $lblSuffix.Left = 820; $lblSuffix.Top = 16; $lblSuffix.AutoSize = $true
    $tab.Controls.Add($lblSuffix)

    $txtSuffix = New-Object System.Windows.Forms.TextBox
    $txtSuffix.Left = 870; $txtSuffix.Top = 12; $txtSuffix.Width = 120
    $tab.Controls.Add($txtSuffix)

    $btnAuto = New-Object System.Windows.Forms.Button
    $btnAuto.Text = 'Autonumeracja'
    $btnAuto.Left = 1000; $btnAuto.Top = 12; $btnAuto.Width = 120
    $tab.Controls.Add($btnAuto)

    $chkRestart = New-Object System.Windows.Forms.CheckBox
    $chkRestart.Text = 'Restart po zmianie'
    $chkRestart.Left = 12; $chkRestart.Top = 44; $chkRestart.Checked = $true
    $tab.Controls.Add($chkRestart)

    $btnRename = New-Object System.Windows.Forms.Button
    $btnRename.Text = 'Zmień nazwy'
    $btnRename.Left = 160; $btnRename.Top = 40; $btnRename.Width = 160
    $tab.Controls.Add($btnRename)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 80; $grid.Width = 1110; $grid.Height = 590
    $grid.AllowUserToAddRows = $false
    $grid.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{Name = 'StaraNazwa'; HeaderText = 'Stara nazwa'; Width = 300 }))
    $grid.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{Name = 'NowaNazwa'; HeaderText = 'Nowa nazwa'; Width = 300 }))
    $tab.Controls.Add($grid)

    $btnLoad.Add_Click({
            $grid.Rows.Clear()
            foreach ($c in & $getTargets) { [void]$grid.Rows.Add(@($c, '')) }
        })

    $btnAuto.Add_Click({
            $start = [int]$numStart.Value; $pad = [int]$numPad.Value
            for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
                $n = $start + $i
                $grid.Rows[$i].Cells['NowaNazwa'].Value = $txtPrefix.Text + ($n.ToString(("D$pad"))) + $txtSuffix.Text
            }
        })

    $btnRename.Add_Click({
            try {
                if ($grid.Rows.Count -eq 0) { Show-Error "Załaduj najpierw hosty do siatki."; return }
                $credToPass = if (-not $State.UseCurrentCreds -and $State.Cred) { $State.Cred } else { $null }
                for ($i = 0; $i -lt $grid.Rows.Count; $i++) {
                    $old = $grid.Rows[$i].Cells['StaraNazwa'].Value
                    $new = $grid.Rows[$i].Cells['NowaNazwa'].Value
                    if ([string]::IsNullOrWhiteSpace($new) -or $old -eq $new) { continue }
                    Write-Log "[$old] zmiana nazwy na '$new'..."
                    $sb = {
                        param($newName, [pscredential]$cred, [bool]$doRestart)
                        $p = @{ NewName = $newName; Force = $true; ErrorAction = 'Stop' }
                        if ($cred) { $p.DomainCredential = $cred }
                        if ($doRestart) { $p.Restart = $true }
                        Rename-Computer @p
                        'OK'
                    }
                    try {
                        $res = Invoke-Remote -ComputerName $old -ScriptBlock $sb -Arg @{ newName = $new; cred = $credToPass; doRestart = $chkRestart.Checked }
                        Write-Log "[$old] wynik: $res"
                    } catch {
                        Write-Log "[$old] błąd zmiany nazwy: $($_.Exception.Message)" 'ERROR'
                    }
                }
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# 16) Udziały: tworzenie/usuwanie + podgląd
Register-ModuleTab -Name 'Udziały (zarządzanie)' -Builder {
    param($tab, $getTargets)

    $lblHost = New-Object System.Windows.Forms.Label
    $lblHost.Text = 'Komputer (jedna nazwa):'
    $lblHost.Left = 12; $lblHost.Top = 16; $lblHost.AutoSize = $true
    $tab.Controls.Add($lblHost)

    $cmbHost = New-Object System.Windows.Forms.ComboBox
    $cmbHost.Left = 140; $cmbHost.Top = 12; $cmbHost.Width = 220; $cmbHost.DropDownStyle = 'DropDownList'
    $tab.Controls.Add($cmbHost)

    $tab.Add_Enter({
            $cmbHost.Items.Clear()
            foreach ($h in Get-SelectedComputers) { [void]$cmbHost.Items.Add($h) }
            if ($cmbHost.Items.Count -gt 0) { $cmbHost.SelectedIndex = 0 }
        })

    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Text = 'Odśwież listę'
    $btnRefresh.Left = 370; $btnRefresh.Top = 12; $btnRefresh.Width = 120
    $tab.Controls.Add($btnRefresh)

    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text = 'Ścieżka lokalna:'
    $lblPath.Left = 12; $lblPath.Top = 52; $lblPath.AutoSize = $true
    $tab.Controls.Add($lblPath)

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Left = 120; $txtPath.Top = 48; $txtPath.Width = 360
    $tab.Controls.Add($txtPath)

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Text = 'Nazwa udziału:'
    $lblName.Left = 500; $lblName.Top = 52; $lblName.AutoSize = $true
    $tab.Controls.Add($lblName)

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Left = 600; $txtName.Top = 48; $txtName.Width = 160
    $tab.Controls.Add($txtName)

    $lblDesc = New-Object System.Windows.Forms.Label
    $lblDesc.Text = 'Opis:'
    $lblDesc.Left = 770; $lblDesc.Top = 52; $lblDesc.AutoSize = $true
    $tab.Controls.Add($lblDesc)

    $txtDesc = New-Object System.Windows.Forms.TextBox
    $txtDesc.Left = 810; $txtDesc.Top = 48; $txtDesc.Width = 310
    $tab.Controls.Add($txtDesc)

    $lblFA = New-Object System.Windows.Forms.Label
    $lblFA.Text = 'FullAccess (grupy, ,/;):'
    $lblFA.Left = 12; $lblFA.Top = 84; $lblFA.AutoSize = $true
    $tab.Controls.Add($lblFA)

    $txtFA = New-Object System.Windows.Forms.TextBox
    $txtFA.Left = 170; $txtFA.Top = 80; $txtFA.Width = 300
    $tab.Controls.Add($txtFA)

    $lblCA = New-Object System.Windows.Forms.Label
    $lblCA.Text = 'ChangeAccess:'
    $lblCA.Left = 480; $lblCA.Top = 84; $lblCA.AutoSize = $true
    $tab.Controls.Add($lblCA)

    $txtCA = New-Object System.Windows.Forms.TextBox
    $txtCA.Left = 570; $txtCA.Top = 80; $txtCA.Width = 240
    $tab.Controls.Add($txtCA)

    $lblRA = New-Object System.Windows.Forms.Label
    $lblRA.Text = 'ReadAccess:'
    $lblRA.Left = 820; $lblRA.Top = 84; $lblRA.AutoSize = $true
    $tab.Controls.Add($lblRA)

    $txtRA = New-Object System.Windows.Forms.TextBox
    $txtRA.Left = 900; $txtRA.Top = 80; $txtRA.Width = 220
    $tab.Controls.Add($txtRA)

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = 'Utwórz udział'
    $btnCreate.Left = 12; $btnCreate.Top = 116; $btnCreate.Width = 170
    $tab.Controls.Add($btnCreate)

    $btnDelete = New-Object System.Windows.Forms.Button
    $btnDelete.Text = 'Usuń udział (po nazwie)'
    $btnDelete.Left = 192; $btnDelete.Top = 116; $btnDelete.Width = 190
    $tab.Controls.Add($btnDelete)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 160; $grid.Width = 1110; $grid.Height = 510
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $refreshAction = {
        try {
            $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
            Write-Log "[$selectedHost] odświeżam listę udziałów..."
            $sb = {
                if (Get-Command Get-SmbShare -ErrorAction SilentlyContinue) {
                    Get-SmbShare | Select-Object Name, Path, Description, ScopeName, CurrentUsers
                } else {
                    Get-CimInstance Win32_Share | Select-Object Name, Path, Description, Type
                }
            }
            $out = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb
            $grid.DataSource = @($out)
        } catch { Write-Log $_.Exception.Message 'ERROR' }
    }

    $btnRefresh.Add_Click($refreshAction)

    $btnCreate.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                $path = $txtPath.Text.Trim(); $name = $txtName.Text.Trim()
                if (-not $path -or -not $name) { Show-Error "Podaj ścieżkę i nazwę udziału."; return }
                $desc = $txtDesc.Text.Trim()
                $fa = $txtFA.Text.Trim(); $ca = $txtCA.Text.Trim(); $ra = $txtRA.Text.Trim()
                Write-Log "[$selectedHost] tworzę udział $name → $path ..."
                $sb = {
                    param($name, $path, $desc, $fa, $ca, $ra)
                    if (-not (Test-Path $path)) { New-Item -Path $path -ItemType Directory -Force | Out-Null }
                    $split = { param($s) if ([string]::IsNullOrWhiteSpace($s)) { @() } else { $s -split '[,;]\s*' } }
                    $full = & $split $fa
                    $chg = & $split $ca
                    $read = & $split $ra

                    if (Get-Command New-SmbShare -ErrorAction SilentlyContinue) {
                        $p = @{Name = $name; Path = $path; ErrorAction = 'Stop' }
                        if ($desc) { $p.Description = $desc }
                        if ($full.Count -gt 0) { $p.FullAccess = $full }
                        if ($chg.Count -gt 0) { $p.ChangeAccess = $chg }
                        if ($read.Count -gt 0) { $p.ReadAccess = $read }
                        New-SmbShare @p | Out-Null
                        'OK (SMB)'
                    } else {
                        $type = 0
                        $r = ([wmiclass]"\\.\root\cimv2:Win32_Share").Create($path, $name, $type, $null, $desc)
                        if ($r.ReturnValue -ne 0) { throw "Win32_Share.Create zwrócił $($r.ReturnValue)" }
                        $perms = @()
                        foreach ($u in $full) { if ($u) { $perms += "${u}:(OI)(CI)F" } }
                        foreach ($u in $chg) { if ($u) { $perms += "${u}:(OI)(CI)M" } }
                        foreach ($u in $read) { if ($u) { $perms += "${u}:(OI)(CI)R" } }
                        foreach ($pmt in $perms) { icacls $path /grant $pmt | Out-Null }
                        'OK (Win32_Share + NTFS ACL)'
                    }
                }
                $out = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{
                    name = $name; path = $path; desc = $desc; fa = $fa; ca = $ca; ra = $ra
                }
                Write-Log "[$selectedHost] wynik: $out"
                & $refreshAction
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })

    $btnDelete.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                $name = $txtName.Text.Trim(); if (-not $name) { Show-Error "Podaj nazwę udziału do usunięcia."; return }
                Write-Log "[$selectedHost] usuwam udział $name ..."
                $sb = {
                    param($name)
                    if (Get-Command Remove-SmbShare -ErrorAction SilentlyContinue) {
                        Remove-SmbShare -Name $name -Force -ErrorAction Stop
                    } else {
                        $s = Get-CimInstance Win32_Share -Filter "Name='$name'"
                        if ($s) { Invoke-CimMethod -InputObject $s -MethodName Delete | Out-Null }
                    }
                    'Usunięto'
                }
                $out = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ name = $name }
                Write-Log "[$selectedHost] wynik: $out"
                & $refreshAction
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# 17) Sterowniki (PnP Signed) + filtr + eksport CSV
Register-ModuleTab -Name 'Sterowniki (PnP)' -Builder {
    param($tab, $getTargets)

    $lblF = New-Object System.Windows.Forms.Label
    $lblF.Text = 'Filtr (nazwa/producent/dostawca):'
    $lblF.Left = 12; $lblF.Top = 16; $lblF.AutoSize = $true
    $tab.Controls.Add($lblF)

    $txtF = New-Object System.Windows.Forms.TextBox
    $txtF.Left = 220; $txtF.Top = 12; $txtF.Width = 300
    $tab.Controls.Add($txtF)

    $btnGet = New-Object System.Windows.Forms.Button
    $btnGet.Text = 'Pobierz'
    $btnGet.Left = 540; $btnGet.Top = 12; $btnGet.Width = 120
    $tab.Controls.Add($btnGet)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Eksport CSV'
    $btnExport.Left = 670; $btnExport.Top = 12; $btnExport.Width = 120
    $tab.Controls.Add($btnExport)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btnGet.Add_Click({
            try {
                $targets = & $getTargets
                Connect-Cim -Computers $targets
                $rows = @()
                $filter = $txtF.Text.Trim()
                foreach ($t in $targets) {
                    if (-not $State.CimSessions.ContainsKey($t)) { continue }
                    Write-Log "[$t] pobieram sterowniki (Win32_PnPSignedDriver)..."
                    $drv = Get-CimInstance Win32_PnPSignedDriver -CimSession $State.CimSessions[$t] |
                    Select-Object DeviceName, DriverVersion, DriverDate, Manufacturer, DriverProviderName, InfName, IsSigned
                    if ($filter) {
                        $drv = $drv | Where-Object {
                            $_.DeviceName -like "*$filter*" -or
                            $_.Manufacturer -like "*$filter*" -or
                            $_.DriverProviderName -like "*$filter*"
                        }
                    }
                    foreach ($d in $drv) {
                        $rows += [pscustomobject]@{
                            Komputer = $t; Urządzenie = $d.DeviceName; Wersja = $d.DriverVersion; Data = $d.DriverDate
                            Producent = $d.Manufacturer; Dostawca = $d.DriverProviderName; INF = $d.InfName; Podpisany = $d.IsSigned
                        }
                    }
                }
                $grid.DataSource = $rows
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })

    $btnExport.Add_Click({
            try {
                if (-not $grid.DataSource) { Show-Error "Brak danych do eksportu."; return }
                $dlg = New-Object System.Windows.Forms.SaveFileDialog
                $dlg.Filter = 'CSV (*.csv)|*.csv'
                $dlg.FileName = 'Sterowniki.csv'
                if ($dlg.ShowDialog() -eq 'OK') {
                    @($grid.DataSource) | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $dlg.FileName
                    Write-Log "Zapisano CSV: $($dlg.FileName)"
                }
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# 18) Lokalni administratorzy — podgląd/edycja
Register-ModuleTab -Name 'Lokalni administratorzy' -Builder {
    param($tab, $getTargets)

    $lblHost = New-Object System.Windows.Forms.Label
    $lblHost.Text = 'Komputer (jedna nazwa):'
    $lblHost.Left = 12; $lblHost.Top = 16; $lblHost.AutoSize = $true
    $tab.Controls.Add($lblHost)

    $cmbHost = New-Object System.Windows.Forms.ComboBox
    $cmbHost.Left = 140; $cmbHost.Top = 12; $cmbHost.Width = 240; $cmbHost.DropDownStyle = 'DropDownList'
    $tab.Controls.Add($cmbHost)

    $tab.Add_Enter({
            $cmbHost.Items.Clear()
            foreach ($h in Get-SelectedComputers) { [void]$cmbHost.Items.Add($h) }
            if ($cmbHost.Items.Count -gt 0) { $cmbHost.SelectedIndex = 0 }
        })

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Pokaż członków'
    $btnLoad.Left = 400; $btnLoad.Top = 12; $btnLoad.Width = 140
    $tab.Controls.Add($btnLoad)

    $txtAcct = New-Object System.Windows.Forms.TextBox
    $txtAcct.Left = 12; $txtAcct.Top = 48; $txtAcct.Width = 360
    $tab.Controls.Add($txtAcct)
    $lblAcct = New-Object System.Windows.Forms.Label
    $lblAcct.Text = 'Konto/grupa do dodania (DOMENA\użytkownik lub grupa):'
    $lblAcct.Left = 12; $lblAcct.Top = 74; $lblAcct.AutoSize = $true
    $tab.Controls.Add($lblAcct)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = 'Dodaj do "Administratorzy"'
    $btnAdd.Left = 384; $btnAdd.Top = 46; $btnAdd.Width = 200
    $tab.Controls.Add($btnAdd)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = 'Usuń zaznaczonego'
    $btnDel.Left = 600; $btnDel.Top = 46; $btnDel.Width = 160
    $tab.Controls.Add($btnDel)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 108; $grid.Width = 1110; $grid.Height = 562
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $false
    $tab.Controls.Add($grid)

    $loadAction = {
        try {
            $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
            Write-Log "[$selectedHost] odczyt grupy Lokalni Administratorzy..."
            $sb = {
                function Get-NetLocalAdminsFallback {
                    $raw = (net localgroup administrators) | Out-String
                    $lines = $raw -split "`r?`n"
                    $body = $false; $items = @()
                    foreach ($l in $lines) {
                        if ($l -match '^-{3,}') { $body = -not $body; continue }
                        if ($body -and $l.Trim()) {
                            $n = $l.Trim()
                            if ($n -notmatch '^(Polecenie zostało|The command completed)') { $items += [pscustomobject]@{Name = $n; ObjectClass = '(net)'; PrincipalSource = 'Unknown'; SID = $null } }
                        }
                    }
                    $items
                }
                if (Get-Command Get-LocalGroupMember -ErrorAction SilentlyContinue) {
                    try {
                        Get-LocalGroupMember -Group 'Administrators' | Select-Object Name, ObjectClass, PrincipalSource, SID
                    } catch { Get-NetLocalAdminsFallback }
                } else { Get-NetLocalAdminsFallback }
            }
            $out = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb
            $grid.DataSource = @($out)
        } catch { Write-Log $_.Exception.Message 'ERROR' }
    }
    $btnLoad.Add_Click($loadAction)

    $btnAdd.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                $acct = $txtAcct.Text.Trim(); if (-not $acct) { Show-Error "Podaj konto/grupę (DOMENA\\użytkownik)."; return }
                Write-Log "[$selectedHost] dodaję $acct do lokalnej grupy Administratorzy..."
                $sb = {
                    param($member)
                    if (Get-Command Add-LocalGroupMember -ErrorAction SilentlyContinue) {
                        Add-LocalGroupMember -Group 'Administrators' -Member $member -ErrorAction Stop
                        'OK (Add-LocalGroupMember)'
                    } else {
                        cmd.exe /c "net localgroup administrators `"$member`" /add" | Out-Null
                        'OK (net localgroup)'
                    }
                }
                $r = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ member = $acct }
                Write-Log "[$selectedHost] $r"
                & $loadAction
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })

    $btnDel.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                if (-not $grid.SelectedRows) { Show-Error "Zaznacz pozycję do usunięcia."; return }
                $name = $grid.SelectedRows[0].Cells['Name'].Value
                Write-Log "[$selectedHost] usuwam $name z lokalnych Administratorów..."
                $sb = {
                    param($member)
                    if (Get-Command Remove-LocalGroupMember -ErrorAction SilentlyContinue) {
                        Remove-LocalGroupMember -Group 'Administrators' -Member $member -ErrorAction Stop
                        'OK (Remove-LocalGroupMember)'
                    } else {
                        cmd.exe /c "net localgroup administrators `"$member`" /delete" | Out-Null
                        'OK (net localgroup)'
                    }
                }
                $r = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ member = $name }
                Write-Log "[$selectedHost] $r"
                & $loadAction
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# 19) Zaplanowane zadania — podgląd/uruchamianie/wyłączanie/usuwanie + proste tworzenie
Register-ModuleTab -Name 'Zadania (Harmonogram)' -Builder {
    param($tab, $getTargets)

    $lblHost = New-Object System.Windows.Forms.Label
    $lblHost.Text = 'Komputer:'
    $lblHost.Left = 12; $lblHost.Top = 16; $lblHost.AutoSize = $true
    $tab.Controls.Add($lblHost)

    $cmbHost = New-Object System.Windows.Forms.ComboBox
    $cmbHost.Left = 80; $cmbHost.Top = 12; $cmbHost.Width = 220; $cmbHost.DropDownStyle = 'DropDownList'
    $tab.Controls.Add($cmbHost)

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = 'Filtr (nazwa/ścieżka):'
    $lblFilter.Left = 320; $lblFilter.Top = 16; $lblFilter.AutoSize = $true
    $tab.Controls.Add($lblFilter)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Left = 460; $txtFilter.Top = 12; $txtFilter.Width = 300
    $tab.Controls.Add($txtFilter)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Pobierz'
    $btnLoad.Left = 780; $btnLoad.Top = 12; $btnLoad.Width = 100
    $tab.Controls.Add($btnLoad)

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = 'Uruchom'
    $btnRun.Left = 890; $btnRun.Top = 12; $btnRun.Width = 80
    $tab.Controls.Add($btnRun)

    $btnEnable = New-Object System.Windows.Forms.Button
    $btnEnable.Text = 'Włącz'
    $btnEnable.Left = 980; $btnEnable.Top = 12; $btnEnable.Width = 70
    $tab.Controls.Add($btnEnable)

    $btnDisable = New-Object System.Windows.Forms.Button
    $btnDisable.Text = 'Wyłącz'
    $btnDisable.Left = 1060; $btnDisable.Top = 12; $btnDisable.Width = 70
    $tab.Controls.Add($btnDisable)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = 'Usuń'
    $btnDel.Left = 1140; $btnDel.Top = 12; $btnDel.Width = 60
    $tab.Controls.Add($btnDel)

    $grpNew = New-Object System.Windows.Forms.GroupBox
    $grpNew.Text = 'Utwórz proste zadanie (SYSTEM)'
    $grpNew.Left = 12; $grpNew.Top = 48; $grpNew.Width = 1188; $grpNew.Height = 100
    $tab.Controls.Add($grpNew)

    $lblTN = New-Object System.Windows.Forms.Label
    $lblTN.Text = 'Nazwa zadania:'; $lblTN.Left = 12; $lblTN.Top = 24; $lblTN.AutoSize = $true
    $grpNew.Controls.Add($lblTN)
    $txtTN = New-Object System.Windows.Forms.TextBox
    $txtTN.Left = 110; $txtTN.Top = 20; $txtTN.Width = 240
    $grpNew.Controls.Add($txtTN)

    $lblAct = New-Object System.Windows.Forms.Label
    $lblAct.Text = 'Akcja (program):'; $lblAct.Left = 370; $lblAct.Top = 24; $lblAct.AutoSize = $true
    $grpNew.Controls.Add($lblAct)
    $txtAct = New-Object System.Windows.Forms.TextBox
    $txtAct.Left = 480; $txtAct.Top = 20; $txtAct.Width = 280
    $grpNew.Controls.Add($txtAct)

    $lblArg = New-Object System.Windows.Forms.Label
    $lblArg.Text = 'Argumenty:'; $lblArg.Left = 770; $lblArg.Top = 24; $lblArg.AutoSize = $true
    $grpNew.Controls.Add($lblArg)
    $txtArg = New-Object System.Windows.Forms.TextBox
    $txtArg.Left = 840; $txtArg.Top = 20; $txtArg.Width = 330
    $grpNew.Controls.Add($txtArg)

    $lblTrig = New-Object System.Windows.Forms.Label
    $lblTrig.Text = 'Trigger:'; $lblTrig.Left = 12; $lblTrig.Top = 60; $lblTrig.AutoSize = $true
    $grpNew.Controls.Add($lblTrig)
    $cmbTrig = New-Object System.Windows.Forms.ComboBox
    $cmbTrig.Left = 70; $cmbTrig.Top = 56; $cmbTrig.Width = 150; $cmbTrig.DropDownStyle = 'DropDownList'
    $cmbTrig.Items.AddRange(@('Na logowanie', 'Codziennie o HH:MM', 'Ręczny (OnDemand)'))
    $cmbTrig.SelectedIndex = 0
    $grpNew.Controls.Add($cmbTrig)

    $txtTime = New-Object System.Windows.Forms.TextBox
    $txtTime.Left = 230; $txtTime.Top = 56; $txtTime.Width = 60
    $txtTime.Text = '07:00'
    $grpNew.Controls.Add($txtTime)

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = 'Utwórz (SYSTEM, Highest)'
    $btnCreate.Left = 310; $btnCreate.Top = 54; $btnCreate.Width = 180
    $grpNew.Controls.Add($btnCreate)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 160; $grid.Width = 1188; $grid.Height = 510
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $false
    $tab.Controls.Add($grid)

    $tab.Add_Enter({
            $cmbHost.Items.Clear()
            foreach ($h in Get-SelectedComputers) { [void]$cmbHost.Items.Add($h) }
            if ($cmbHost.Items.Count -gt 0) { $cmbHost.SelectedIndex = 0 }
        })

    $loadAction = {
        try {
            $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
            Write-Log "[$selectedHost] odczyt zadań Harmonogramu..."
            $sb = {
                param($flt)
                $tasks = Get-ScheduledTask
                if ($flt) {
                    $tasks = $tasks | Where-Object { $_.TaskName -like "*$flt*" -or $_.TaskPath -like "*$flt*" }
                }
                $rows = @()
                foreach ($t in $tasks) {
                    try {
                        $i = Get-ScheduledTaskInfo -TaskName $t.TaskName -TaskPath $t.TaskPath
                        $rows += [pscustomobject]@{
                            TaskName = $t.TaskName; TaskPath = $t.TaskPath; State = $i.State; Enabled = $t.Enabled
                            LastRun = $i.LastRunTime; NextRun = $i.NextRunTime; Author = $t.Author; Description = $t.Description
                        }
                    } catch {
                        $rows += [pscustomobject]@{
                            TaskName = $t.TaskName; TaskPath = $t.TaskPath; State = '(brak informacji)'; Enabled = $t.Enabled
                            LastRun = $null; NextRun = $null; Author = $t.Author; Description = $t.Description
                        }
                    }
                }
                $rows
            }
            $out = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ flt = $txtFilter.Text.Trim() }
            $grid.DataSource = @($out | Sort-Object TaskPath, TaskName)
        } catch { Write-Log $_.Exception.Message 'ERROR' }
    }
    $btnLoad.Add_Click($loadAction)

    foreach ($pair in @(
            @{Btn = $btnRun; Op = 'Run' },
            @{Btn = $btnEnable; Op = 'Enable' },
            @{Btn = $btnDisable; Op = 'Disable' },
            @{Btn = $btnDel; Op = 'Delete' }
        )) {
        $pair.Btn.Add_Click({
                try {
                    $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                    if (-not $grid.SelectedRows) { Show-Error "Zaznacz zadanie."; return }
                    $tn = $grid.SelectedRows[0].Cells['TaskName'].Value
                    $tp = $grid.SelectedRows[0].Cells['TaskPath'].Value
                    Write-Log "[$selectedHost] $($pair.Op) zadania $tp$tn ..."
                    $sb = {
                        param($tp, $tn, $op)
                        switch ($op) {
                            'Run' { Start-ScheduledTask -TaskPath $tp -TaskName $tn; 'Started' }
                            'Enable' { Enable-ScheduledTask -TaskPath $tp -TaskName $tn; 'Enabled' }
                            'Disable' { Disable-ScheduledTask -TaskPath $tp -TaskName $tn; 'Disabled' }
                            'Delete' { Unregister-ScheduledTask -TaskPath $tp -TaskName $tn -Confirm:$false; 'Deleted' }
                        }
                    }
                    $r = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ tp = $tp; tn = $tn; op = $pair.Op }
                    Write-Log "[$selectedHost] $r"
                    $btnLoad.PerformClick()
                } catch { Write-Log $_.Exception.Message 'ERROR' }
            })
    }

    $btnCreate.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                $tn = $txtTN.Text.Trim(); $exe = $txtAct.Text.Trim()
                if (-not $tn -or -not $exe) { Show-Error "Podaj nazwę i ścieżkę programu."; return }
                $taskArgs = $txtArg.Text.Trim(); $trig = $cmbTrig.Text; $time = $txtTime.Text.Trim()
                Write-Log "[$selectedHost] tworze zadanie $tn -> $exe $taskArgs ($trig)..."
                $sb = {
                    param($tn, $exe, $taskArgs, $trig, $time)
                    $act = New-ScheduledTaskAction -Execute $exe -Argument $taskArgs
                    $prin = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest
                    switch ($trig) {
                        'Na logowanie' { $tr = New-ScheduledTaskTrigger -AtLogOn }
                        'Codziennie o HH:MM' {
                            $h, $m = $time.Split(':'); $dt = (Get-Date).Date.AddHours([int]$h).AddMinutes([int]$m)
                            $tr = New-ScheduledTaskTrigger -Daily -At $dt.TimeOfDay
                        }
                        default { $tr = $null }
                    }
                    if ($tr) { $t = New-ScheduledTask -Action $act -Trigger $tr -Principal $prin }
                    else { $t = New-ScheduledTask -Action $act -Principal $prin -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries) }
                    Register-ScheduledTask -TaskName $tn -InputObject $t -Force | Out-Null
                    'OK'
                }
                $r = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ tn = $tn; exe = $exe; taskArgs = $taskArgs; trig = $trig; time = $time }
                Write-Log "[$selectedHost] $r"
                $btnLoad.PerformClick()
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# 20) Zapora Windows — podgląd/enable/disable + szybkie tworzenie i usuwanie
Register-ModuleTab -Name 'Zapora Windows' -Builder {
    param($tab, $getTargets)

    $lblHost = New-Object System.Windows.Forms.Label
    $lblHost.Text = 'Komputer:'; $lblHost.Left = 12; $lblHost.Top = 16; $lblHost.AutoSize = $true
    $tab.Controls.Add($lblHost)

    $cmbHost = New-Object System.Windows.Forms.ComboBox
    $cmbHost.Left = 80; $cmbHost.Top = 12; $cmbHost.Width = 220; $cmbHost.DropDownStyle = 'DropDownList'
    $tab.Controls.Add($cmbHost)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Pokaż reguły (Inbound)'
    $btnLoad.Left = 320; $btnLoad.Top = 12; $btnLoad.Width = 180
    $tab.Controls.Add($btnLoad)

    $btnToggle = New-Object System.Windows.Forms.Button
    $btnToggle.Text = 'Włącz/wyłącz zaznaczoną'
    $btnToggle.Left = 510; $btnToggle.Top = 12; $btnToggle.Width = 190
    $tab.Controls.Add($btnToggle)

    $btnDel = New-Object System.Windows.Forms.Button
    $btnDel.Text = 'Usuń po nazwie'
    $btnDel.Left = 710; $btnDel.Top = 12; $btnDel.Width = 140
    $tab.Controls.Add($btnDel)

    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = 'Nowa reguła: Nazwa / Port / Protokół'; $lblNew.Left = 12; $lblNew.Top = 48; $lblNew.AutoSize = $true
    $tab.Controls.Add($lblNew)
    $txtRule = New-Object System.Windows.Forms.TextBox
    $txtRule.Left = 220; $txtRule.Top = 44; $txtRule.Width = 260; $txtRule.Text = 'MojaReguła'
    $tab.Controls.Add($txtRule)
    $numPort = New-Object System.Windows.Forms.NumericUpDown
    $numPort.Left = 490; $numPort.Top = 44; $numPort.Width = 80; $numPort.Minimum = 1; $numPort.Maximum = 65535; $numPort.Value = 5985
    $tab.Controls.Add($numPort)
    $cmbProto = New-Object System.Windows.Forms.ComboBox
    $cmbProto.Left = 580; $cmbProto.Top = 44; $cmbProto.Width = 90; $cmbProto.DropDownStyle = 'DropDownList'
    $cmbProto.Items.AddRange(@('TCP', 'UDP')); $cmbProto.SelectedIndex = 0
    $tab.Controls.Add($cmbProto)
    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Text = 'Utwórz regułę (Allow, Inbound, Any profile)'
    $btnCreate.Left = 680; $btnCreate.Top = 42; $btnCreate.Width = 340
    $tab.Controls.Add($btnCreate)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 80; $grid.Width = 1110; $grid.Height = 590
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $false
    $tab.Controls.Add($grid)

    $tab.Add_Enter({
            $cmbHost.Items.Clear()
            foreach ($h in Get-SelectedComputers) { [void]$cmbHost.Items.Add($h) }
            if ($cmbHost.Items.Count -gt 0) { $cmbHost.SelectedIndex = 0 }
        })

    $loadAction = {
        try {
            $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
            Write-Log "[$selectedHost] odczyt reguł zapory (Inbound)..."
            $sb = {
                if (Get-Command Get-NetFirewallRule -ErrorAction SilentlyContinue) {
                    $rules = Get-NetFirewallRule -Direction Inbound | Get-NetFirewallRule
                    $rows = foreach ($r in $rules) {
                        $pf = (Get-NetFirewallPortFilter -AssociatedNetFirewallRule $r -ErrorAction SilentlyContinue)
                        [pscustomobject]@{
                            Name = $r.Name; DisplayName = $r.DisplayName; Enabled = $r.Enabled; Action = $r.Action; Profile = $r.Profile
                            Protocol = ($pf.Protocol); LocalPort = ($pf.LocalPort -join ','); Program = $r.Program; Group = $r.Group
                        }
                    }
                    $rows
                } else {
                    'Brak modułu NetSecurity — użyj netsh'
                }
            }
            $out = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb
            $grid.DataSource = @($out)
        } catch { Write-Log $_.Exception.Message 'ERROR' }
    }
    $btnLoad.Add_Click($loadAction)

    $btnToggle.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                if (-not $grid.SelectedRows) { Show-Error "Zaznacz regułę."; return }
                $name = $grid.SelectedRows[0].Cells['Name'].Value
                Write-Log "[$selectedHost] przełączam regułę zapory $name..."
                $sb = {
                    param($name)
                    if (Get-Command Get-NetFirewallRule -ErrorAction SilentlyContinue) {
                        $r = Get-NetFirewallRule -Name $name -ErrorAction Stop
                        if ($r.Enabled -eq 'True') { Disable-NetFirewallRule -Name $name | Out-Null; 'Disabled' }
                        else { Enable-NetFirewallRule -Name $name | Out-Null; 'Enabled' }
                    } else { 'netsh only — brak toggle' }
                }
                $r = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ name = $name }
                Write-Log "[$selectedHost] $r"
                & $loadAction
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })

    $btnCreate.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                $name = $txtRule.Text.Trim(); $port = [int]$numPort.Value; $proto = $cmbProto.Text
                if (-not $name) { Show-Error "Podaj nazwę reguły."; return }
                Write-Log "[$selectedHost] tworzę inbound allow $proto/$port ..."
                $sb = {
                    param($name, $proto, $port)
                    if (Get-Command New-NetFirewallRule -ErrorAction SilentlyContinue) {
                        New-NetFirewallRule -DisplayName $name -Name $name -Direction Inbound -Action Allow -Enabled True -Protocol $proto -LocalPort $port -Profile Any | Out-Null
                        'OK (New-NetFirewallRule)'
                    } else {
                        netsh advfirewall firewall add rule name="$name" dir=in action=allow protocol=$proto localport=$port
                        'OK (netsh)'
                    }
                }
                $r = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ name = $name; proto = $proto; port = $port }
                Write-Log "[$selectedHost] $r"
                & $loadAction
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })

    $btnDel.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                $name = $txtRule.Text.Trim(); if (-not $name) { Show-Error "Podaj nazwę reguły do usunięcia."; return }
                Write-Log "[$selectedHost] usuwam regułę $name ..."
                $sb = {
                    param($name)
                    if (Get-Command Remove-NetFirewallRule -ErrorAction SilentlyContinue) {
                        Remove-NetFirewallRule -Name $name -ErrorAction Stop | Out-Null
                        'Deleted (Remove-NetFirewallRule)'
                    } else {
                        netsh advfirewall firewall delete rule name="$name" dir=in | Out-Null
                        'Deleted (netsh)'
                    }
                }
                $r = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ name = $name }
                Write-Log "[$selectedHost] $r"
                & $loadAction
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# 21) LAPS — podgląd hasła/wygaśnięcia + wymuszenie rotacji
Register-ModuleTab -Name 'LAPS (AD)' -Builder {
    param($tab, $getTargets)

    $btnGet = New-Object System.Windows.Forms.Button
    $btnGet.Text = 'Pokaż LAPS dla zaznaczonych'
    $btnGet.Left = 12; $btnGet.Top = 12; $btnGet.Width = 260
    $tab.Controls.Add($btnGet)

    $btnRotate = New-Object System.Windows.Forms.Button
    $btnRotate.Text = 'Wymuś rotację hasła'
    $btnRotate.Left = 280; $btnRotate.Top = 12; $btnRotate.Width = 200
    $tab.Controls.Add($btnRotate)

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Kopiuj hasło z zaznaczonego'
    $btnCopy.Left = 490; $btnCopy.Top = 12; $btnCopy.Width = 220
    $tab.Controls.Add($btnCopy)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $false
    $tab.Controls.Add($grid)

    function Convert-FileTimeLocal([string]$ft) {
        try {
            if ([string]::IsNullOrWhiteSpace($ft)) { return $null }
            [DateTime]::FromFileTimeUtc([int64]$ft).ToLocalTime()
        } catch { $null }
    }

    $btnGet.Add_Click({
            try {
                if (-not (Get-Module -ListAvailable ActiveDirectory)) { throw "Brak modułu ActiveDirectory (RSAT)." }
                Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
                $targets = & $getTargets
                $rows = @()
                foreach ($c in $targets) {
                    Write-Log "[AD:$c] pobieram atrybuty LAPS..."
                    try {
                        $obj = Get-ADComputer -Identity $c -Properties * -ErrorAction Stop
                        $legacyPwd = $obj.'ms-Mcs-AdmPwd'
                        $legacyExp = Convert-FileTimeLocal $obj.'ms-Mcs-AdmPwdExpirationTime'
                        $winPwd = $null
                        $winExp = $null
                        # Preferuj oficjalny cmdlet jeśli dostępny (Windows LAPS)
                        if (Get-Command Get-LapsADPassword -ErrorAction SilentlyContinue) {
                            try {
                                $lp = Get-LapsADPassword -Identity $c -ErrorAction Stop
                                if ($lp -and $lp.Password -and $lp.ExpirationTime) {
                                    $winPwd = $lp.Password
                                    $winExp = $lp.ExpirationTime.ToLocalTime()
                                }
                            } catch {}
                        }
                        if (-not $winPwd) {
                            $winPwd = $obj.'msLAPS-Password'
                            $winExp = $obj.'msLAPS-PasswordExpirationTime'
                            if ($winExp -and ($winExp -is [string])) { try { $winExp = [datetime]::Parse($winExp).ToLocalTime() } catch {} }
                        }
                        if ($legacyPwd) {
                            $rows += [pscustomobject]@{Komputer = $c; Rozwiązanie = 'LAPS (legacy)'; Hasło = $legacyPwd; Wygasa = $legacyExp; Info = '' }
                        } elseif ($winPwd) {
                            $rows += [pscustomobject]@{Komputer = $c; Rozwiązanie = 'Windows LAPS'; Hasło = $winPwd; Wygasa = $winExp; Info = '' }
                        } else {
                            $rows += [pscustomobject]@{Komputer = $c; Rozwiązanie = 'Brak/No access'; Hasło = '(niedostępne)'; Wygasa = $null; Info = 'Brak uprawnień lub nie skonfigurowano LAPS' }
                        }
                    } catch {
                        $rows += [pscustomobject]@{Komputer = $c; Rozwiązanie = 'Błąd'; Hasło = ''; Wygasa = $null; Info = $_.Exception.Message }
                    }
                }
                $grid.DataSource = $rows
            } catch { Show-Error "Nie mogę odczytać LAPS z AD." $_; Write-Log $_.Exception.Message 'ERROR' }
        })

    $btnRotate.Add_Click({
            try {
                if (-not (Get-Module -ListAvailable ActiveDirectory)) { throw "Brak modułu ActiveDirectory (RSAT)." }
                Import-Module ActiveDirectory -ErrorAction Stop | Out-Null
                $targets = & $getTargets
                foreach ($c in $targets) {
                    Write-Log "[AD:$c] wymuszam rotację hasła LAPS..."
                    try {
                        $ok = $false
                        if (Get-Command Reset-LapsPassword -ErrorAction SilentlyContinue) {
                            Reset-LapsPassword -Identity $c -ErrorAction Stop | Out-Null
                            $ok = $true
                        }
                        if (-not $ok) {
                            # legacy LAPS: ustaw datę ważności na 0, co wymusi odświeżenie
                            Set-ADComputer -Identity $c -Replace @{'ms-Mcs-AdmPwdExpirationTime' = '0' } -ErrorAction Stop
                        }
                        Write-Log "[AD:$c] zlecono rotację."
                    } catch {
                        Write-Log "[AD:$c] błąd rotacji: $($_.Exception.Message)" 'ERROR'
                    }
                }
            } catch { Show-Error "Operacja wymaga RSAT/AD." $_; Write-Log $_.Exception.Message 'ERROR' }
        })

    $btnCopy.Add_Click({
            try {
                if (-not $grid.SelectedRows) { Show-Error "Zaznacz pozycję z hasłem."; return }
                $lapsPassword = $grid.SelectedRows[0].Cells['Haslo'].Value
                if ([string]::IsNullOrWhiteSpace($lapsPassword) -or $lapsPassword -eq '(niedostepne)') { Show-Error "Brak hasla do skopiowania."; return }
                [System.Windows.Forms.Clipboard]::SetText($lapsPassword)
                Write-Log "Skopiowano hasło LAPS do schowka (lokalnie)."
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# 22) Certyfikaty — LM\My / LM\Root + eksport .cer
Register-ModuleTab -Name 'Certyfikaty (LM)' -Builder {
    param($tab, $getTargets)

    $lblHost = New-Object System.Windows.Forms.Label
    $lblHost.Text = 'Komputer:'; $lblHost.Left = 12; $lblHost.Top = 16; $lblHost.AutoSize = $true
    $tab.Controls.Add($lblHost)

    $cmbHost = New-Object System.Windows.Forms.ComboBox
    $cmbHost.Left = 80; $cmbHost.Top = 12; $cmbHost.Width = 220; $cmbHost.DropDownStyle = 'DropDownList'
    $tab.Controls.Add($cmbHost)

    $cmbStore = New-Object System.Windows.Forms.ComboBox
    $cmbStore.Left = 320; $cmbStore.Top = 12; $cmbStore.Width = 160; $cmbStore.DropDownStyle = 'DropDownList'
    $cmbStore.Items.AddRange(@('My', 'Root', 'TrustedPublisher', 'CA'))
    $cmbStore.SelectedIndex = 0
    $tab.Controls.Add($cmbStore)

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Left = 500; $txtFilter.Top = 12; $txtFilter.Width = 300
    $tab.Controls.Add($txtFilter)
    $lblF = New-Object System.Windows.Forms.Label
    $lblF.Text = 'Filtr (Subject/Thumbprint):'; $lblF.Left = 500; $lblF.Top = 36; $lblF.AutoSize = $true
    $tab.Controls.Add($lblF)

    $btnLoad = New-Object System.Windows.Forms.Button
    $btnLoad.Text = 'Pobierz'
    $btnLoad.Left = 820; $btnLoad.Top = 12; $btnLoad.Width = 100
    $tab.Controls.Add($btnLoad)

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = 'Eksport .cer zaznaczonego'
    $btnExport.Left = 930; $btnExport.Top = 12; $btnExport.Width = 190
    $tab.Controls.Add($btnExport)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 60; $grid.Width = 1110; $grid.Height = 610
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $false
    $tab.Controls.Add($grid)

    $tab.Add_Enter({
            $cmbHost.Items.Clear()
            foreach ($h in Get-SelectedComputers) { [void]$cmbHost.Items.Add($h) }
            if ($cmbHost.Items.Count -gt 0) { $cmbHost.SelectedIndex = 0 }
        })

    $btnLoad.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                $store = $cmbStore.Text; $flt = $txtFilter.Text.Trim()
                Write-Log "[$selectedHost] certyfikaty LocalMachine\\$store ..."
                $sb = {
                    param($store, $flt)
                    $path = "Cert:\LocalMachine\$store"
                    $list = Get-ChildItem -Path $path -ErrorAction Stop | ForEach-Object {
                        $eku = $_.EnhancedKeyUsageList | ForEach-Object { $_.FriendlyName } | Where-Object { $_ } | Sort-Object -Unique
                        [pscustomobject]@{
                            Subject = $_.Subject; Thumbprint = $_.Thumbprint; NotBefore = $_.NotBefore; NotAfter = $_.NotAfter
                            FriendlyName = $_.FriendlyName; HasPrivateKey = $_.HasPrivateKey; EKU = ($eku -join '; ')
                        }
                    }
                    if ($flt) { $list = $list | Where-Object { $_.Subject -like "*$flt*" -or $_.Thumbprint -like "*$flt*" } }
                    $list
                }
                $out = Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ store = $store; flt = $flt }
                $grid.DataSource = @($out | Sort-Object Subject)
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })

    $btnExport.Add_Click({
            try {
                $selectedHost = $cmbHost.Text; if (-not $selectedHost) { Show-Error "Wybierz komputer."; return }
                if (-not $grid.SelectedRows) { Show-Error "Zaznacz certyfikat."; return }
                $store = $cmbStore.Text
                $thumb = $grid.SelectedRows[0].Cells['Thumbprint'].Value
                $dlg = New-Object System.Windows.Forms.SaveFileDialog
                $dlg.Filter = 'CER (*.cer)|*.cer'; $dlg.FileName = "$selectedHost-$thumb.cer"
                if ($dlg.ShowDialog() -ne 'OK') { return }
                Write-Log "[$selectedHost] eksportuję $thumb z LocalMachine\\$store do $($dlg.FileName) ..."
                # zapis na hoście i pobranie pliku
                $remoteTmp = "C:\Windows\Temp\DomainOps\$thumb.cer"
                $sb = {
                    param($store, $thumb, $outFile)
                    if (-not (Test-Path (Split-Path -Path $outFile -Parent))) { New-Item -ItemType Directory -Force -Path (Split-Path -Path $outFile -Parent) | Out-Null }
                    $cert = Get-ChildItem -Path ("Cert:\LocalMachine\$store\$thumb")
                    Export-Certificate -Cert $cert -FilePath $outFile -Force | Out-Null
                    'OK'
                }
                Invoke-Remote -ComputerName $selectedHost -ScriptBlock $sb -Arg @{ store = $store; thumb = $thumb; outFile = $remoteTmp } | Out-Null
                $sess = if ($State.UseCurrentCreds -or -not $State.Cred) { New-PSSession -ComputerName $selectedHost } else { New-PSSession -ComputerName $selectedHost -Credential $State.Cred }
                Copy-Item -FromSession $sess -Path $remoteTmp -Destination $dlg.FileName -Force
                Remove-PSSession $sess
                Write-Log "[$selectedHost] eksport zakończony."
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# 23) Diagnostyka łączności — Ping / WSMan(5985) / WinRM HTTPS(5986) / SMB(445) / RDP(3389) / WMI(135)
Register-ModuleTab -Name 'Diagnostyka łączności' -Builder {
    param($tab, $getTargets)

    $btn = New-Object System.Windows.Forms.Button
    $btn.Text = 'Testuj łączność dla zaznaczonych'
    $btn.Left = 12; $btn.Top = 12; $btn.Width = 260
    $tab.Controls.Add($btn)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Left = 12; $grid.Top = 52; $grid.Width = 1110; $grid.Height = 618
    $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false
    $tab.Controls.Add($grid)

    $btn.Add_Click({
            try {
                $targets = & $getTargets
                $rows = @()
                foreach ($t in $targets) {
                    Write-Log "[$t] test łączności..."
                    $ping = Test-Connection -ComputerName $t -Count 1 -Quiet -ErrorAction SilentlyContinue
                    $wsman = $false; try { Test-WSMan -ComputerName $t -ErrorAction Stop | Out-Null; $wsman = $true } catch {}
                    $t5986 = (Test-NetConnection -ComputerName $t -Port 5986 -WarningAction SilentlyContinue).TcpTestSucceeded
                    $smb = (Test-NetConnection -ComputerName $t -Port 445  -WarningAction SilentlyContinue).TcpTestSucceeded
                    $rdp = (Test-NetConnection -ComputerName $t -Port 3389 -WarningAction SilentlyContinue).TcpTestSucceeded
                    $wmi = (Test-NetConnection -ComputerName $t -Port 135  -WarningAction SilentlyContinue).TcpTestSucceeded
                    $rows += [pscustomobject]@{
                        Komputer = $t; Ping = $ping; WSMan5985 = $wsman; WinRM5986 = $t5986; SMB445 = $smb; RDP3389 = $rdp; WMI135 = $wmi
                    }
                }
                $grid.DataSource = $rows
            } catch { Write-Log $_.Exception.Message 'ERROR' }
        })
}

# ======= START UI =======
$form.Add_Shown({ $chkUseCurrent.Checked = $true })
[void]$form.ShowDialog()
