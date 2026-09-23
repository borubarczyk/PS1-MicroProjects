# Skrypt PowerShell z GUI do analizy uprawnień folderów
# Wyszukuje foldery, w których uprawnienia (ACE) nadano bezpośrednio użytkownikowi zamiast grupie.

# Import wymaganych modułów
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:job = $null

# Tworzenie formularza GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = "Analiza uprawnień folderów"
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"

# Ścieżka do analizy
$tbPath = New-Object System.Windows.Forms.TextBox
$tbPath.Location = New-Object System.Drawing.Point(20, 22)
$tbPath.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($tbPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Wybierz..."
$btnBrowse.Location = New-Object System.Drawing.Point(330, 20)
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Wybierz folder do analizy"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tbPath.Text = $dlg.SelectedPath }
})
$form.Controls.Add($btnBrowse)

# Przycisk do uruchomienia analizy
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Uruchom analizę"
$btnStart.Size = New-Object System.Drawing.Size(110, 23)
$btnStart.Location = New-Object System.Drawing.Point(410, 20)
$form.Controls.Add($btnStart)

# Lista wyników
$Results = New-Object System.Windows.Forms.TextBox
$Results.Multiline = $true
$Results.ReadOnly = $true
$Results.ScrollBars = "Both"
$Results.WordWrap = $false
$Results.Location = New-Object System.Drawing.Point(20, 60)
$Results.Size = New-Object System.Drawing.Size(540, 250)
$form.Controls.Add($Results)

# Pasek postępu
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 320)
$progressBar.Size = New-Object System.Drawing.Size(540, 20)
$form.Controls.Add($progressBar)

# Zadanie w tle działa w osobnym procesie - wyniki odbieramy timerem w wątku GUI
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 500
$timer.Add_Tick({
    if (-not $script:job) { return }
    foreach ($line in @(Receive-Job -Job $script:job -ErrorAction SilentlyContinue)) {
        $Results.AppendText("$line`r`n")
    }
    if ($script:job.State -in 'Completed', 'Failed', 'Stopped') {
        $timer.Stop()
        $Results.AppendText("--- Zakończono ($($script:job.State)) ---`r`n")
        Remove-Job -Job $script:job -Force
        $script:job = $null
        $progressBar.Style = 'Blocks'
        $progressBar.Value = 0
        $btnStart.Enabled = $true
    }
})

$btnStart.Add_Click({
    $rootPath = $tbPath.Text.Trim()
    if (-not $rootPath -or -not (Test-Path -LiteralPath $rootPath -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show("Wybierz istniejący folder do analizy.", "Błąd") | Out-Null
        return
    }

    $btnStart.Enabled = $false
    $Results.Clear()
    $progressBar.Style = 'Marquee'

    # Uruchomienie analizy w tle
    $script:job = Start-Job -ArgumentList $rootPath -ScriptBlock {
        param($rootPath)

        $typeCache = @{}
        # Zwraca 'User', 'Group' lub 'Other' dla wpisu ACL
        function Get-PrincipalType {
            param($Identity)
            $key = $Identity.Value
            if ($typeCache.ContainsKey($key)) { return $typeCache[$key] }
            $type = 'Other'
            try {
                $sid = $Identity.Translate([System.Security.Principal.SecurityIdentifier])
                $account = $Identity.Value
                if ($account -match '^(BUILTIN|NT AUTHORITY|NT SERVICE|APPLICATION PACKAGE AUTHORITY|CREATOR OWNER|Everyone)' -or $sid.Value -notmatch '^S-1-5-21-') {
                    $type = 'Other'
                } else {
                    try {
                        $entry = [ADSI]"LDAP://<SID=$($sid.Value)>"
                        $classes = @($entry.Properties['objectClass'])
                        if ($classes -contains 'group') { $type = 'Group' }
                        elseif ($classes -contains 'computer') { $type = 'Other' }
                        elseif ($classes -contains 'user') { $type = 'User' }
                    } catch { }
                    if ($type -eq 'Other') {
                        # Konta lokalne komputera
                        $name = ($account -split '\\')[-1]
                        try {
                            $local = [ADSI]"WinNT://$env:COMPUTERNAME/$name"
                            if ($local.SchemaClassName -eq 'User') { $type = 'User' }
                            elseif ($local.SchemaClassName -eq 'Group') { $type = 'Group' }
                        } catch { }
                    }
                }
            } catch { }
            $typeCache[$key] = $type
            return $type
        }

        $folders = @(Get-Item -LiteralPath $rootPath) + @(Get-ChildItem -LiteralPath $rootPath -Directory -Recurse -Force -ErrorAction SilentlyContinue)
        foreach ($folder in $folders) {
            $folderPath = $folder.FullName
            try {
                $acl = Get-Acl -LiteralPath $folderPath -ErrorAction Stop
                foreach ($entry in $acl.Access) {
                    if ((Get-PrincipalType $entry.IdentityReference) -eq 'User') {
                        "Znaleziono uprawnienia użytkownika w folderze: $folderPath"
                        "  Użytkownik: $($entry.IdentityReference)"
                        "  Prawa: $($entry.FileSystemRights)$(if ($entry.IsInherited) { ' (dziedziczone)' })"
                        ""
                    }
                }
            } catch {
                "Nie można uzyskać ACL dla folderu ${folderPath}: $($_.Exception.Message)"
            }
        }
    }
    $timer.Start()
})

# Pokaż formularz
$form.Add_Shown({ $form.Activate() })
$form.Add_FormClosing({
    $timer.Stop()
    if ($script:job) { Stop-Job -Job $script:job; Remove-Job -Job $script:job -Force }
})
[void]$form.ShowDialog()
