<#
.SYNOPSIS
Interaktywne narzędzie GUI do liczenia hashy (MD5/SHA1/SHA256) dla wielu plików i folderów.

.DESCRIPTION
Tworzy aplikację Windows Forms pozwalającą wskazać pliki lub katalogi (również metodą drag & drop),
oblicza ich wartości skrótów wybranym algorytmem, prezentuje wyniki i zapamiętuje historię sesji.
Zawiera pasek postępu i komunikaty dla błędnych ścieżek, dzięki czemu na bieżąco widać status pracy.

.EXAMPLE
PS> .\File-HashChecker.ps1
Uruchamia okno „Sprawdzanie hashy plików” w którym można wskazać elementy do przeliczenia.

.NOTES
Obsługuje zaznaczenie wielu pozycji w oknie dialogowym i automatyczne rozwijanie folderów do plików.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-ErrorMessage {
    param($err)
    if ($err -and $err.Exception) {
        return $err.Exception.Message
    }
    elseif ($err) {
        return $err.ToString()
    }
    return "Nieznany błąd."
}

# Okno główne
$form = New-Object System.Windows.Forms.Form
$form.Text = "Sprawdzanie hashy plików"
$form.Size = New-Object System.Drawing.Size(750,500)
$form.StartPosition = "CenterScreen"

# Lista wyboru algorytmu
$combo = New-Object System.Windows.Forms.ComboBox
$combo.Location = New-Object System.Drawing.Point(20,20)
$combo.Size = New-Object System.Drawing.Size(120,20)
$combo.Items.AddRange(@("MD5","SHA1","SHA256"))
$combo.SelectedIndex = 2
$form.Controls.Add($combo)

# Przycisk wyboru plików
$button = New-Object System.Windows.Forms.Button
$button.Text = "Wybierz pliki"
$button.Location = New-Object System.Drawing.Point(160,18)
$button.Size = New-Object System.Drawing.Size(120,25)
$form.Controls.Add($button)

# Zakładki na wyniki i historię
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(20,60)
$tabControl.Size = New-Object System.Drawing.Size(690,330)
$form.Controls.Add($tabControl)

$resultsTab = New-Object System.Windows.Forms.TabPage
$resultsTab.Text = "Wyniki"
$tabControl.Controls.Add($resultsTab)

$historyTab = New-Object System.Windows.Forms.TabPage
$historyTab.Text = "Historia"
$tabControl.Controls.Add($historyTab)

# Pole tekstowe na wynik
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.ScrollBars = "Both"
$textBox.WordWrap = $false
$textBox.Font = New-Object System.Drawing.Font("Consolas",10)
$textBox.Dock = "Fill"
$resultsTab.Controls.Add($textBox)

# Historia poprzednich wyników
$historyTextBox = New-Object System.Windows.Forms.TextBox
$historyTextBox.Multiline = $true
$historyTextBox.ScrollBars = "Both"
$historyTextBox.WordWrap = $false
$historyTextBox.Font = New-Object System.Drawing.Font("Consolas",10)
$historyTextBox.Dock = "Fill"
$historyTextBox.ReadOnly = $true
$historyTextBox.BackColor = [System.Drawing.SystemColors]::Window
$historyTab.Controls.Add($historyTextBox)

# Pasek postępu
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20,410)
$progressBar.Size = New-Object System.Drawing.Size(690,20)
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Dialog wyboru plików
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Multiselect = $true
$dialog.Title = "Wybierz pliki do sprawdzenia"

# Funkcja obsługująca liczenie hashy i aktualizowanie UI
$processFiles = {
    param([string[]]$files)

    if (-not $files -or $files.Count -eq 0) {
        return
    }

    $algo = $combo.SelectedItem
    if (-not $algo) {
        return
    }

    $resolvedFiles = New-Object System.Collections.Generic.List[string]
    $infoMessages = New-Object System.Collections.Generic.List[string]

    foreach ($path in $files) {
        try {
            if (-not (Test-Path -LiteralPath $path)) {
                $infoMessages.Add("Nie znaleziono ścieżki: $path")
                continue
            }

            $item = Get-Item -LiteralPath $path -ErrorAction Stop
            if ($item.PSIsContainer) {
                try {
                    $childFiles = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse -ErrorAction Stop)
                    if (-not $childFiles) {
                        $infoMessages.Add("Folder pusty: $($item.FullName)")
                    }
                    else {
                        foreach ($child in $childFiles) {
                            $resolvedFiles.Add($child.FullName)
                        }
                        $infoMessages.Add("Folder: $($item.FullName) (dodano $($childFiles.Count) plików)")
                    }
                }
                catch {
                    $infoMessages.Add("Błąd przy folderze: $($item.FullName)`r`n$(Get-ErrorMessage $_)")
                }
            }
            else {
                $resolvedFiles.Add($item.FullName)
            }
        }
        catch {
            $infoMessages.Add("Błąd przy ścieżce: $path`r`n$(Get-ErrorMessage $_)")
        }
    }

    $textBox.Clear()
    if ($infoMessages.Count -gt 0) {
        foreach ($msg in $infoMessages) {
            $textBox.AppendText("$msg`r`n`r`n")
        }
    }

    if ($resolvedFiles.Count -eq 0) {
        $textBox.AppendText("Brak plików do sprawdzenia.`r`n")
        return
    }

    $progressBar.Value = 0
    $progressBar.Maximum = $resolvedFiles.Count
    $progressBar.Step = 1
    $progressBar.Visible = $true
    $progressBar.Refresh()

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

    try {
        foreach ($file in $resolvedFiles) {
            try {
                $hash = Get-FileHash -LiteralPath $file -Algorithm $algo -ErrorAction Stop
                $textBox.AppendText(
                    "Plik: $($file)`r`n" +
                    "${algo}: $($hash.Hash)`r`n`r`n"
                )
            }
            catch {
                $errorMessage = Get-ErrorMessage $_
                $textBox.AppendText(
                    "Błąd przy pliku: $file`r`n$errorMessage`r`n`r`n"
                )
            }

            if ($progressBar.Value -lt $progressBar.Maximum) {
                $progressBar.PerformStep()
            }
            [System.Windows.Forms.Application]::DoEvents()
        }

        $currentRunText = $textBox.Text
        $historyTextBox.AppendText(
            "[$timestamp] Algorytm: $algo`r`n$currentRunText`r`n------------------------------`r`n"
        )
    }
    finally {
        $progressBar.Visible = $false
        $progressBar.Value = 0
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

# Obsługa przeciągnij i upuść
$form.AllowDrop = $true
$textBox.AllowDrop = $true
$historyTextBox.AllowDrop = $true

$dragEnterHandler = {
    param($_sender, $_eventArgs)
    if ($_eventArgs.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $_eventArgs.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
    else {
        $_eventArgs.Effect = [System.Windows.Forms.DragDropEffects]::None
    }
}

$dragDropHandler = {
    param($_sender, $_eventArgs)
    if ($_eventArgs.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $files = [string[]]$_eventArgs.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        & $processFiles $files
    }
}

$form.Add_DragEnter($dragEnterHandler)
$form.Add_DragDrop($dragDropHandler)
$textBox.Add_DragEnter($dragEnterHandler)
$textBox.Add_DragDrop($dragDropHandler)
$historyTextBox.Add_DragEnter($dragEnterHandler)
$historyTextBox.Add_DragDrop($dragDropHandler)

# Akcja po kliknięciu
$button.Add_Click({
    if ($dialog.ShowDialog() -eq "OK") {
        & $processFiles $dialog.FileNames
    }
})

# Start GUI
[void]$form.ShowDialog()
