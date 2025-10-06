# Skrypt PowerShell z GUI do analizy uprawnień folderów

# Import wymaganych modułów
Add-Type -AssemblyName System.Windows.Forms

# Tworzenie formularza GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = "Analiza uprawnień folderów"
$form.Size = New-Object System.Drawing.Size(600, 400)

# Przycisk do uruchomienia analizy
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Uruchom analizę"
$btnStart.Location = New-Object System.Drawing.Point(20, 20)
$btnStart.Add_Click({
    # Uruchomienie analizy w tle
    Start-Job -ScriptBlock {
        # Ścieżka do analizy (można dodać wybór folderu)
        $rootPath = "C:\Ścieżka\do\analizy"  # Zmień na odpowiednią ścieżkę
        $forbiddenUserFolders = @(
            "C:\Ścieżka\do\zakazanych\folder1",
            "C:\Ścieżka\do\zakazanych\folder2"
        )

        # Funkcja sprawdzająca, czy to grupa
        function Is-Group {
            param ($Identity)
            if ($Identity -like "*Group*" -or $Identity -like "*G*") {
                return $true
            }
            return $false
        }

        # Główna pętla analizy
        Get-ChildItem -Path $rootPath -Directory -Recurse | ForEach-Object {
            $folderPath = $_.FullName
            try {
                $acl = Get-Acl -Path $folderPath
                $entries = $acl.AccessControlEntries

                foreach ($entry in $entries) {
                    $identity = $entry.IdentityReference
                    $rights = $entry.FileSystemRights

                    if (-not (Is-Group $identity)) {
                        # Przekazanie wyniku do GUI
                        Invoke-Command -ScriptBlock {
                            $form.Results.AppendText("Znaleziono uprawnienia użytkownika w folderze: $using:folderPath`n")
                            $form.Results.AppendText("  Użytkownik: $identity`n")
                            $form.Results.AppendText("  Prawa: $rights`n`n")
                        }
                    }
                }
            } catch {
                Invoke-Command -ScriptBlock {
                    $form.Results.AppendText("Nie można uzyskać ACL dla folderu $using:folderPath: $_`n")
                }
            }
        }
    }
})

# Listę wyników
$Results = New-Object System.Windows.Forms.TextBox
$Results.Multiline = $true
$Results.ScrollBars = "Both"
$Results.Location = New-Object System.Drawing.Point(20, 60)
$Results.Size = New-Object System.Drawing.Size(500, 250)
$form.Controls.Add($Results)

# Pasek postępu
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 320)
$progressBar.Size = New-Object System.Drawing.Size(500, 20)
$form.Controls.Add($progressBar)

# Dodanie kontrolki do formularza
$form.Controls.Add($btnStart)

# Pokaż formularz
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
