# Import wymaganych modułów
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
Import-Module ActiveDirectory

# Tworzenie głównego okna
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD User Manager"
$form.Size = New-Object System.Drawing.Size(600,400)
$form.StartPosition = "CenterScreen"

# Lista rozwijana z użytkownikami
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Location = New-Object System.Drawing.Point(10,10)
$comboBox.Size = New-Object System.Drawing.Size(200,20)
$comboBox.DropDownStyle = "DropDown"
$comboBox.AutoCompleteMode = "SuggestAppend"
$comboBox.AutoCompleteSource = "CustomSource"

# Pole tekstowe z informacjami
$textBox = New-Object System.Windows.Forms.RichTextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(560,200)
$textBox.ReadOnly = $true
$textBox.ScrollBars = "Vertical"

# Przycisk Zablokuj konto
# To-Do - Dodaj potwierdzenie przed zablokowaniem konta oraz zmień tekst przycisku na "Odblokuj konto" jeśli konto jest zablokowane
$lockButton = New-Object System.Windows.Forms.Button
$lockButton.Location = New-Object System.Drawing.Point(10,250)
$lockButton.Size = New-Object System.Drawing.Size(100,30)
$lockButton.Text = "Zablokuj konto"

# Przycisk Reset hasła
$resetButton = New-Object System.Windows.Forms.Button
$resetButton.Location = New-Object System.Drawing.Point(120,250)
$resetButton.Size = New-Object System.Drawing.Size(100,30)
$resetButton.Text = "Reset hasła"

# Pobieranie użytkowników
# To-Do - Dodaj obsługę błędów na wypadek braku uprawnień lub problemów z połączeniem oraz pobieraj dane użytkownika tylko na podstawie wybranego konta
try {
    $users = Get-ADUser -Filter * -ResultSetSize $null | Select-Object -ExpandProperty SamAccountName
    foreach ($sam in $users) {
        if ($sam) {
            $comboBox.Items.Add($sam) | Out-Null
            $comboBox.AutoCompleteCustomSource.Add($sam) | Out-Null
        }
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Błąd podczas pobierania użytkowników: $($_.Exception.Message)", "Błąd")
}

# Funkcja pokazująca informacje o użytkowniku
function Show-UserInfo {
    $selectedUser = $comboBox.Text
    if ($selectedUser) {
        try {
            $user = Get-ADUser -Identity $selectedUser -Properties Enabled,LockedOut,LastBadPasswordAttempt,badPwdCount,LastLogonDate,PasswordLastSet,PasswordExpired,DisplayName,Created
            # Obliczanie daty wygaśnięcia hasła
            $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
            $passwordExpires = if ($user.PasswordLastSet) { $user.PasswordLastSet.AddDays($maxPasswordAge) } else { "Nigdy nie ustawiono" }
            
            $textBox.Clear()

            $textBox.AppendText("Nazwa użytkownika: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.SamAccountName)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Nazwa wyświetlana: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.DisplayName)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Konto utworzone: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.Created)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Konto aktywne: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.Enabled)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Ostatnie logowanie: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.LastLogonDate)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Hasło ostatnio ustawione: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.PasswordLastSet)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Hasło ważne do: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$passwordExpires`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Hasło wygasłe: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.PasswordExpired)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Zablokowany (lockout): ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.LockedOut)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Ost. bledna proba: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.LastBadPasswordAttempt)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            $textBox.AppendText("Liczba blednych prob: ")
            $textBox.SelectionColor = [System.Drawing.Color]::Green
            $textBox.AppendText("$($user.badPwdCount)`r`n")
            $textBox.SelectionColor = $textBox.ForeColor

            # Ustawienia przyciskow wg statusu
            if ($user.Enabled) { $lockButton.Text = "Zablokuj konto" } else { $lockButton.Text = "Odblokuj (aktywuj) konto" }
            if ($unlockButton) { $unlockButton.Enabled = [bool]$user.LockedOut }
        }
        catch {
            $textBox.Clear()
            $errorMessage = "Błąd: Nie znaleziono użytkownika lub brak uprawnień`r`n" + $_.Exception.Message
            $textBox.SelectionColor = [System.Drawing.Color]::Red
            $textBox.AppendText($errorMessage)
            $textBox.SelectionColor = $textBox.ForeColor
        }
    }
}

# Funkcja blokowania konta
$lockButton.Add_Click({
    $selectedUser = $comboBox.Text
    if ($selectedUser) {
        try {
            $u = Get-ADUser -Identity $selectedUser -Properties Enabled; if ($u.Enabled) { $confirm = [System.Windows.Forms.MessageBox]::Show("Czy na pewno zablokowac (wylaczyc) konto '" + $selectedUser + "'?","Potwierdzenie",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning); if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) { Disable-ADAccount -Identity $selectedUser; [System.Windows.Forms.MessageBox]::Show("Konto zostalo zablokowane (wylaczone)", "Sukces") | Out-Null } } else { $confirm = [System.Windows.Forms.MessageBox]::Show("Czy na pewno odblokowac (wlaczyc) konto '" + $selectedUser + "'?","Potwierdzenie",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question); if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) { Enable-ADAccount -Identity $selectedUser; [System.Windows.Forms.MessageBox]::Show("Konto zostalo odblokowane (wlaczone)", "Sukces") | Out-Null } }
            [System.Windows.Forms.MessageBox]::Show("Konto zostało zablokowane", "Sukces")
            Show-UserInfo
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Błąd podczas blokowania konta: $($_.Exception.Message)", "Błąd")
        }
    }
})

# Funkcja resetowania hasła
$resetButton.Add_Click({
    $selectedUser = $comboBox.Text
    if ($selectedUser) {
        # Tworzenie okna resetu hasła
        $resetForm = New-Object System.Windows.Forms.Form
        $resetForm.Text = "Reset hasła"
        $resetForm.Size = New-Object System.Drawing.Size(300,150)
        $resetForm.StartPosition = "CenterScreen"

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
        $label.Text = "Wprowadź nowe hasło:"

        $passBox = New-Object System.Windows.Forms.TextBox
        $passBox.Location = New-Object System.Drawing.Point(10,40)
        $passBox.Size = New-Object System.Drawing.Size(260,20)
        $passBox.UseSystemPasswordChar = $true

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(10,70)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $resetForm.Controls.AddRange(@($label,$passBox,$okButton))
        
        if ($resetForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                Set-ADAccountPassword -Identity $selectedUser -NewPassword (ConvertTo-SecureString $passBox.Text -AsPlainText -Force)
                [System.Windows.Forms.MessageBox]::Show("Hasło zostało zresetowane", "Sukces")
                Show-UserInfo
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Błąd podczas resetowania hasła: $($_.Exception.Message)", "Błąd")
            }
        }
    }
})

# Aktualizacja informacji przy zmianie użytkownika
$comboBox.Add_SelectedIndexChanged({Show-UserInfo})
$comboBox.Add_KeyUp({
    if ($_.KeyCode -eq "Enter") {
        Show-UserInfo
    }
})

# Przycisk odblokowania konta po zbyt wielu probach logowania (lockout)
$unlockButton = New-Object System.Windows.Forms.Button
$unlockButton.Location = New-Object System.Drawing.Point(230,250)
$unlockButton.Size = New-Object System.Drawing.Size(160,30)
$unlockButton.Text = "Odblokuj (lockout)"
$unlockButton.Enabled = $false
$unlockButton.Add_Click({
    $selectedUser = $comboBox.Text
    if (-not $selectedUser) { return }
    try {
        $u = Get-ADUser -Identity $selectedUser -Properties LockedOut
        if (-not $u.LockedOut) {
            [System.Windows.Forms.MessageBox]::Show("Konto nie jest zablokowane probami logowania.", "Info") | Out-Null
            return
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show("Odblokowac konto po zbyt wielu probach logowania?", "Potwierdzenie", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        Unlock-ADAccount -Identity $selectedUser
        [System.Windows.Forms.MessageBox]::Show("Konto zostalo odblokowane (lockout)", "Sukces") | Out-Null
        Show-UserInfo
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Blad podczas odblokowywania konta: $($_.Exception.Message)", "Blad") | Out-Null
    }
})

# Dodanie kontrolek do formularza
$form.Controls.AddRange(@($comboBox,$textBox,$lockButton,$resetButton,$unlockButton))

# Pokazanie formularza
$form.ShowDialog()
