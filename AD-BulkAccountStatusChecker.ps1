#Requires -Modules ActiveDirectory

# --- Moduł AD + szybka weryfikacja domeny (opcjonalnie)
Import-Module ActiveDirectory -ErrorAction Stop
try { [void](Get-ADDomain) } catch { [System.Windows.Forms.MessageBox]::Show("Brak połączenia z AD: $($_.Exception.Message)"); return }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# --- Główne okno
$form                = New-Object System.Windows.Forms.Form
$form.Text           = "Sprawdzanie statusu kont AD"
$form.Size           = New-Object System.Drawing.Size(650,720)
$form.StartPosition  = "CenterScreen"
$form.MaximizeBox    = $false

$pad = 12

# --- Etykiety
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point($pad,$pad)
$label.Size     = New-Object System.Drawing.Size(600,20)
$label.Text     = "Wprowadź loginy (jeden na linię):"
$form.Controls.Add($label)

$countLabel = New-Object System.Windows.Forms.Label
$countLabel.Location = New-Object System.Drawing.Point($pad, $pad+20)
$countLabel.Size     = New-Object System.Drawing.Size(600,20)
$countLabel.Text     = "Liczba kont: 0 | Unikalne: 0"
$form.Controls.Add($countLabel)

# --- Wejście
$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Location = New-Object System.Drawing.Point($pad, $pad+45)
$inputBox.Size     = New-Object System.Drawing.Size(610,140)
$inputBox.Multiline = $true
$inputBox.ScrollBars = "Vertical"
$inputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$inputBox.Add_TextChanged({
    $lines = $inputBox.Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $uniq  = $lines | Select-Object -Unique
    $countLabel.Text = "Liczba kont: {0} | Unikalne: {1}" -f $lines.Count, $uniq.Count
})
$form.Controls.Add($inputBox)

# --- Panel sterowania
$filterCombo = New-Object System.Windows.Forms.ComboBox
$filterCombo.Location = New-Object System.Drawing.Point($pad, 210)
$filterCombo.Size     = New-Object System.Drawing.Size(180,24)
$filterCombo.DropDownStyle = "DropDownList"
[void]$filterCombo.Items.AddRange(@("Wszystkie","Tylko włączone","Tylko wyłączone"))
$filterCombo.SelectedIndex = 0
$form.Controls.Add($filterCombo)

$sortCheckBox = New-Object System.Windows.Forms.CheckBox
$sortCheckBox.Location = New-Object System.Drawing.Point(210, 212)
$sortCheckBox.Size     = New-Object System.Drawing.Size(160,20)
$sortCheckBox.Text     = "Sortuj alfabetycznie"
$sortCheckBox.Checked  = $true
$form.Controls.Add($sortCheckBox)

# --- Przyciski akcji
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(400, 208)
$button.Size     = New-Object System.Drawing.Size(120,28)
$button.Text     = "Sprawdź konta"
$form.Controls.Add($button)

$exportBtn = New-Object System.Windows.Forms.Button
$exportBtn.Location = New-Object System.Drawing.Point(535, 208)
$exportBtn.Size     = New-Object System.Drawing.Size(85,28)
$exportBtn.Text     = "Eksport CSV"
$exportBtn.Enabled  = $false
$form.Controls.Add($exportBtn)

# --- Wyniki
$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point($pad, 245)
$outputBox.Size     = New-Object System.Drawing.Size(610,380)
$outputBox.ReadOnly = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($outputBox)

# --- Status zbiorczy
$statsLabel = New-Object System.Windows.Forms.Label
$statsLabel.Location = New-Object System.Drawing.Point($pad, 630)
$statsLabel.Size     = New-Object System.Drawing.Size(610,20)
$statsLabel.Text     = "Znalezione: 0 | Włączone: 0 | Wyłączone: 0 | Błędy: 0"
$form.Controls.Add($statsLabel)

# --- Pasek postępu
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point($pad, 655)
$progress.Size     = New-Object System.Drawing.Size(610,18)
$progress.Style    = 'Continuous'
$form.Controls.Add($progress)

# --- Bufor wyników do eksportu
$script:resultsForExport = @()

function Write-Line([string]$text, [System.Drawing.Color]$color) {
    $outputBox.SelectionColor = $color
    $outputBox.AppendText($text + "`r`n")
    $outputBox.SelectionColor = [System.Drawing.Color]::Black
}

# --- BackgroundWorker dla responsywności
$bw = New-Object System.ComponentModel.BackgroundWorker
$bw.WorkerReportsProgress = $true

$bw.add_DoWork({
    param($sender, $e)

    $input = $e.Argument
    $logins = $input.Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique

    $total = [math]::Max(1, $logins.Count)
    $localResults = New-Object System.Collections.Generic.List[object]

    for ($i=0; $i -lt $logins.Count; $i++) {
        $login = $logins[$i]
        $rec = [PSCustomObject]@{ Login=$login; Enabled=$null; Name=$null; Error=$null }

        try {
            $user = Get-ADUser -Identity $login -Properties Enabled,Name -ErrorAction Stop
            if ($null -ne $user) {
                $rec.Enabled = [bool]$user.Enabled
                $rec.Name    = $user.Name
            } else {
                $rec.Error = "NotFound"
            }
        } catch {
            $rec.Error = $_.Exception.Message
        }

        $localResults.Add($rec) | Out-Null
        $pct = [int](($i+1)/$total*100)
        $sender.ReportProgress($pct)  # to tylko pasek
    }

    $e.Result = $localResults
})

$bw.add_ProgressChanged({
    param($sender, $e)
    $progress.Value = [Math]::Min(100, [Math]::Max(0, $e.ProgressPercentage))
})

$bw.add_RunWorkerCompleted({
    param($sender, $e)

    $outputBox.Clear()
    $progress.Value = 0
    $exportBtn.Enabled = $false
    $button.Enabled = $true

    if ($e.Error) {
        [System.Windows.Forms.MessageBox]::Show("Błąd: $($e.Error.Message)")
        return
    }

    $results = $e.Result

    # sort + zestawienia
    if ($sortCheckBox.Checked) {
        $results = $results | Sort-Object Login
    }

    $enabledCnt   = ($results | Where-Object { $_.Enabled -eq $true }).Count
    $disabledCnt  = ($results | Where-Object { $_.Enabled -eq $false }).Count
    $errorCnt     = ($results | Where-Object { $_.Enabled -eq $null }).Count
    $foundCnt     = $enabledCnt + $disabledCnt

    # filtr + render
    $filter = $filterCombo.SelectedItem
    $script:resultsForExport = @()

    foreach ($r in $results) {
        $show = switch ($true) {
            { $r.Enabled -eq $null } { $filter -eq "Wszystkie"; break }
            { $filter -eq "Wszystkie" } { $true; break }
            { $filter -eq "Tylko włączone" -and $r.Enabled } { $true; break }
            { $filter -eq "Tylko wyłączone" -and -not $r.Enabled } { $true; break }
            default { $false }
        }

        if ($show) {
            if ($r.Enabled -eq $null) {
                Write-Line ("Konto: {0} | {1} | BŁĄD: {2}" -f $r.Login, ($r.Name ?? "-"), ($r.Error ?? "Nie znaleziono")), [System.Drawing.Color]::Black
            } elseif ($r.Enabled) {
                Write-Line ("Konto: {0} | {1} | Status: WŁĄCZONE" -f $r.Login, ($r.Name ?? "-")), [System.Drawing.Color]::Green
            } else {
                Write-Line ("Konto: {0} | {1} | Status: WYŁĄCZONE" -f $r.Login, ($r.Name ?? "-")), [System.Drawing.Color]::Red
            }
        }

        # zawsze zbieraj do eksportu (pełne dane)
        $script:resultsForExport += [PSCustomObject]@{
            Login   = $r.Login
            Name    = $r.Name
            Enabled = $(if ($r.Enabled -eq $null) { $null } else { [bool]$r.Enabled })
            Error   = $r.Error
        }
    }

    $statsLabel.Text = "Znalezione: {0} | Włączone: {1} | Wyłączone: {2} | Błędy: {3}" -f $foundCnt, $enabledCnt, $disabledCnt, $errorCnt
    $exportBtn.Enabled = $script:resultsForExport.Count -gt 0
})

# --- Zdarzenia przycisków
$button.Add_Click({
    # Lock UI i odpal worker
    if (-not $bw.IsBusy) {
        $button.Enabled = $false
        $outputBox.Clear()
        $statsLabel.Text = "Przetwarzanie..."
        $bw.RunWorkerAsync(@{ Text = $inputBox.Text })
    }
})

$exportBtn.Add_Click({
    if ($script:resultsForExport.Count -eq 0) { return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV (*.csv)|*.csv"
    $dlg.FileName = "wyniki_ad.csv"
    if ($dlg.ShowDialog() -eq "OK") {
        $script:resultsForExport | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("Zapisano: $($dlg.FileName)")
    }
})

# --- Start
[void]$form.ShowDialog()
