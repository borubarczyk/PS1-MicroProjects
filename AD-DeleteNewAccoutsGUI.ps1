Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Konfiguracja logu ---
$logFile = "$env:USERPROFILE\Desktop\AD_Delete_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Function Write-Log {
    param([string]$msg)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $msg"
    $line | Out-File -FilePath $logFile -Append -Encoding utf8
}

# Import modułu AD (jeśli nie zaimportowany)
Try {
    if (-not (Get-Module -Name ActiveDirectory)) {
        Import-Module ActiveDirectory -ErrorAction Stop
    }
} Catch {
    [System.Windows.Forms.MessageBox]::Show("Nie można załadować modułu ActiveDirectory. Upewnij się, że RSAT/AD module jest zainstalowany i masz odpowiednie uprawnienia.`nBłąd: $($_.Exception.Message)", "Błąd", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    Exit
}

# --- UI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Nowo utworzone konta w AD"
$form.Size = New-Object System.Drawing.Size(900,600)
$form.StartPosition = "CenterScreen"

# Label i combobox dla wyboru minut
$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "Pokaż konta utworzone w ciągu ostatnich:"
$lbl.AutoSize = $true
$lbl.Location = New-Object System.Drawing.Point(10,14)
$form.Controls.Add($lbl)

$combo = New-Object System.Windows.Forms.ComboBox
$combo.DropDownStyle = 'DropDownList'
$combo.Items.AddRange(@("5","10","15","30","45","60"))
$combo.SelectedIndex = 2 # domyślnie 15
$combo.Location = New-Object System.Drawing.Point(260,10)
$combo.Width = 80
$form.Controls.Add($combo)

# Refresh button
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Odśwież"
$btnRefresh.Location = New-Object System.Drawing.Point(360,8)
$btnRefresh.Width = 90
$form.Controls.Add($btnRefresh)

# Select All button
$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Zaznacz wszystko"
$btnSelectAll.Location = New-Object System.Drawing.Point(460,8)
$btnSelectAll.Width = 120
$form.Controls.Add($btnSelectAll)

# Deselect All button
$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Text = "Odznacz wszystko"
$btnDeselectAll.Location = New-Object System.Drawing.Point(590,8)
$btnDeselectAll.Width = 120
$form.Controls.Add($btnDeselectAll)

# Info label
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Text = "Wybierz konta i kliknij 'Usuń konto z AD' aby usunąć zaznaczone konta."
$lblInfo.AutoSize = $true
$lblInfo.Location = New-Object System.Drawing.Point(10,40)
$form.Controls.Add($lblInfo)

# ListView (z checkboxami)
$listView = New-Object System.Windows.Forms.ListView
$listView.Location = New-Object System.Drawing.Point(10,70)
$listView.Size = New-Object System.Drawing.Size(860,420)
$listView.View = 'Details'
$listView.FullRowSelect = $true
$listView.GridLines = $true
$listView.CheckBoxes = $true
$listView.MultiSelect = $true

# Kolumny: DisplayName, SamAccountName, whenCreated, DistinguishedName
$listView.Columns.Add("DisplayName", 220) > $null
$listView.Columns.Add("SamAccountName", 150) > $null
$listView.Columns.Add("whenCreated", 160) > $null
$listView.Columns.Add("DistinguishedName", 320) > $null

$form.Controls.Add($listView)

# Delete button
$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Text = "Usuń konto z AD"
$btnDelete.Location = New-Object System.Drawing.Point(10,510)
$btnDelete.Size = New-Object System.Drawing.Size(140,30)
$btnDelete.BackColor = [System.Drawing.Color]::FromArgb(220,50,50)
$btnDelete.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnDelete)

# Disable/Enable button for safer option (opcjonalne) - zamiast kasować można wyłączyć
$btnDisable = New-Object System.Windows.Forms.Button
$btnDisable.Text = "Wyłącz konto (disable)"
$btnDisable.Location = New-Object System.Drawing.Point(160,510)
$btnDisable.Size = New-Object System.Drawing.Size(160,30)
$form.Controls.Add($btnDisable)

# Status label
$status = New-Object System.Windows.Forms.Label
$status.Text = "Status: gotowy"
$status.AutoSize = $true
$status.Location = New-Object System.Drawing.Point(340,516)
$form.Controls.Add($status)

# --- Funkcje pomocnicze ---
Function Get-NewADUsers {
    param([int]$minutes)
    $status.Text = "Status: pobieranie..."
    $form.Refresh()
    $startTime = (Get-Date).AddMinutes(-$minutes)
    Write-Log "Pobieranie kont utworzonych od $startTime (ostatnie $minutes minut)."

    # Uzyskanie użytkowników - jeśli w Twoim środowisku AD jest dużo użytkowników możesz preferować bardziej zoptymalizowane filtry/LDAP query
    Try {
        $users = Get-ADUser -Filter * -Properties whenCreated,displayName,sAMAccountName,distinguishedName |
                 Where-Object { $_.whenCreated -ge $startTime } |
                 Sort-Object -Property whenCreated -Descending
        Write-Log "Znaleziono $($users.Count) użytkowników."
        return $users
    } Catch {
        Write-Log "Błąd przy pobieraniu użytkowników: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Błąd przy pobieraniu użytkowników: $($_.Exception.Message)", "Błąd", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return @()
    } Finally {
        $status.Text = "Status: pobieranie zakończone"
    }
}

Function Populate-List {
    param([int]$minutes)
    $listView.Items.Clear()
    $users = Get-NewADUsers -minutes $minutes
    foreach ($u in $users) {
        $lvi = New-Object System.Windows.Forms.ListViewItem($u.DisplayName)
        $lvi.SubItems.Add($u.sAMAccountName)
        $lvi.SubItems.Add($u.whenCreated.ToString("yyyy-MM-dd HH:mm:ss"))
        $lvi.SubItems.Add($u.DistinguishedName)
        # przechowaj DN w Tag dla późniejszego użycia
        $lvi.Tag = $u.DistinguishedName
        $listView.Items.Add($lvi) > $null
    }
    $status.Text = "Status: załadowano $($listView.Items.Count) rekordów"
}

# --- Eventy ---
$btnRefresh.Add_Click({
    $minutes = [int]$combo.SelectedItem
    if (-not $minutes) { $minutes = 15 }
    Populate-List -minutes $minutes
})

# Zaznacz/Odznacz wszystko
$btnSelectAll.Add_Click({
    foreach ($it in $listView.Items) { $it.Checked = $true }
})
$btnDeselectAll.Add_Click({
    foreach ($it in $listView.Items) { $it.Checked = $false }
})

# Double-click na wierszu pokaże szczegóły
$listView.Add_DoubleClick({
    if ($listView.SelectedItems.Count -gt 0) {
        $item = $listView.SelectedItems[0]
        $dn = $item.Tag
        $txt = "DisplayName: $($item.Text)`nSamAccountName: $($item.SubItems[1].Text)`nwhenCreated: $($item.SubItems[2].Text)`nDistinguishedName: $dn"
        [System.Windows.Forms.MessageBox]::Show($txt, "Szczegóły konta", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
})

# Delete button click
$btnDelete.Add_Click({
    $checked = @()
    foreach ($it in $listView.Items) {
        if ($it.Checked) { $checked += $it }
    }

    if ($checked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Brak zaznaczonych kont do usunięcia.", "Uwaga", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    # Potwierdzenie - pokaz listę samAccountName do usunięcia
    $names = ($checked | ForEach-Object { $_.SubItems[1].Text }) -join "`n"
    $confirmText = "Zostaną usunięte następujące konta z AD:`n`n$names`n`nCzy na pewno chcesz kontynuować? Ta operacja jest nieodwracalna."
    $res = [System.Windows.Forms.MessageBox]::Show($confirmText, "Potwierdź usunięcie", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    # Usuń każde konto
    foreach ($it in $checked) {
        $dn = $it.Tag
        $sam = $it.SubItems[1].Text
        Try {
            Write-Log "Usuwanie konta: $sam, DN: $dn"
            # Usuń obiekt AD - użyj Remove-ADObject (usuwa każdy obiekt AD). Wymaga uprawnień.
            Remove-ADObject -Identity $dn -Confirm:$false -ErrorAction Stop
            Write-Log "Usunięto: $sam"
        } Catch {
            Write-Log "Błąd usuwania $sam : $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Błąd usuwania $sam : $($_.Exception.Message)", "Błąd", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    # Odśwież listę po usunięciu
    $minutes = [int]$combo.SelectedItem
    Populate-List -minutes $minutes
})

# Disable account instead of deleting (bezpieczniejsza opcja)
$btnDisable.Add_Click({
    $checked = @()
    foreach ($it in $listView.Items) {
        if ($it.Checked) { $checked += $it }
    }
    if ($checked.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Brak zaznaczonych kont do wyłączenia.", "Uwaga", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $names = ($checked | ForEach-Object { $_.SubItems[1].Text }) -join "`n"
    $confirmText = "Zostaną wyłączone (AccountDisabled) następujące konta:`n`n$names`n`nKont użytkownicy zostaną zablokowani ale nie usunięci. Kontynuować?"
    $res = [System.Windows.Forms.MessageBox]::Show($confirmText, "Potwierdź wyłączenie", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($res -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    foreach ($it in $checked) {
        $dn = $it.Tag
        $sam = $it.SubItems[1].Text
        Try {
            Write-Log "Wyłączanie konta: $sam, DN: $dn"
            Set-ADUser -Identity $sam -Enabled $false -ErrorAction Stop
            Write-Log "Wyłączono: $sam"
        } Catch {
            Write-Log "Błąd wyłączania $sam : $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Błąd wyłączania $sam : $($_.Exception.Message)", "Błąd", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    $minutes = [int]$combo.SelectedItem
    Populate-List -minutes $minutes
})

# Przy starcie załaduj listę
Populate-List -minutes ([int]$combo.SelectedItem)

# Pokaż formularz
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

# Na końcu pokaż lokalizację logu
Write-Host "Log operacji: $logFile"
