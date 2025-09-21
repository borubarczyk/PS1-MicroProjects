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
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(560,200)
$textBox.Multiline = $true
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
    $users = Get-ADUser -Filter * -Properties Enabled,LastLogonDate,PasswordLastSet,PasswordExpired
    foreach ($user in $users) {
        if ($user.SamAccountName) {
            $comboBox.Items.Add($user.SamAccountName) | Out-Null
            $comboBox.AutoCompleteCustomSource.Add($user.SamAccountName)
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
            $user = Get-ADUser -Identity $selectedUser -Properties Enabled,LastLogonDate,PasswordLastSet,PasswordExpired,DisplayName,Created
            # Obliczanie daty wygaśnięcia hasła
            $maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.Days
            $passwordExpires = if ($user.PasswordLastSet) { $user.PasswordLastSet.AddDays($maxPasswordAge) } else { "Nigdy nie ustawiono" }
            
            $info = "Nazwa użytkownika: $($user.SamAccountName)`r`n"
            $info += "Nazwa wyświetlana: $($user.DisplayName)`r`n"
            $info += "Konto utworzone: $($user.Created)`r`n"
            $info += "Konto aktywne: $($user.Enabled)`r`n"
            $info += "Ostatnie logowanie: $($user.LastLogonDate)`r`n"
            $info += "Hasło ostatnio ustawione: $($user.PasswordLastSet)`r`n"
            $info += "Hasło ważne do: $passwordExpires`r`n"
            $info += "Hasło wygasłe: $($user.PasswordExpired)"
            $textBox.Text = $info
        }
        catch {
            $textBox.Text = "Błąd: Nie znaleziono użytkownika lub brak uprawnień`r`n$($_.Exception.Message)"
        }
    }
}

# Funkcja blokowania konta
$lockButton.Add_Click({
    $selectedUser = $comboBox.Text
    if ($selectedUser) {
        try {
            Disable-ADAccount -Identity $selectedUser
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

# Dodanie kontrolek do formularza
$form.Controls.AddRange(@($comboBox,$textBox,$lockButton,$resetButton))

# Pokazanie formularza
$form.ShowDialog()
