#Requires -Modules ActiveDirectory

#region Active Directory Setup
Import-Module ActiveDirectory -ErrorAction Stop
try { [void](Get-ADDomain) } catch { [System.Windows.Forms.MessageBox]::Show("Brak polaczenia z AD: $($_.Exception.Message)"); return }
#endregion Active Directory Setup

#region WinForms Bootstrapping
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
#endregion WinForms Bootstrapping

#region GUI Construction

#region Main Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Sprawdzanie statusu kont AD"
$form.Size = New-Object System.Drawing.Size(760, 740)
$form.MinimumSize = New-Object System.Drawing.Size(720, 620)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false

$pad = 12

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = [System.Windows.Forms.DockStyle]::Fill
$layout.ColumnCount = 1
$layout.RowCount = 4
$layout.Padding = New-Object System.Windows.Forms.Padding($pad)
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 42.5)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 7.5)))
$form.Controls.Add($layout)
#endregion Main Form

#region Input Section
$inputGroup = New-Object System.Windows.Forms.GroupBox
$inputGroup.Text = "Lista kont do sprawdzenia"
$inputGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
$inputGroup.MinimumSize = New-Object System.Drawing.Size(0, 220)

$inputLayout = New-Object System.Windows.Forms.TableLayoutPanel
$inputLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$inputLayout.ColumnCount = 1
$inputLayout.RowCount = 3
$inputLayout.Padding = New-Object System.Windows.Forms.Padding(8)
$inputLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$inputLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$inputLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

$label = New-Object System.Windows.Forms.Label
$label.AutoSize = $true
 $label.Text = "Wprowadź loginy (po jednym na linii):"

$countLabel = New-Object System.Windows.Forms.Label
$countLabel.AutoSize = $true
$countLabel.Text = "Liczba kont: 0 | Unikalne: 0"

$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Multiline = $true
$inputBox.ScrollBars = "Vertical"
$inputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$inputBox.MinimumSize = New-Object System.Drawing.Size(0, 150)
$inputBox.Height = 150
$inputBox.Dock = [System.Windows.Forms.DockStyle]::Fill

$inputLayout.Controls.Add($label, 0, 0)
$inputLayout.Controls.Add($countLabel, 0, 1)
$inputLayout.Controls.Add($inputBox, 0, 2)
$inputGroup.Controls.Add($inputLayout)
$layout.Controls.Add($inputGroup, 0, 0)
#endregion Input Section

#region Options Section
$optionsGroup = New-Object System.Windows.Forms.GroupBox
$optionsGroup.Text = "Filtry i sterowanie"
$optionsGroup.Dock = [System.Windows.Forms.DockStyle]::Fill

$optionsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$optionsLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$optionsLayout.ColumnCount = 6
$optionsLayout.RowCount = 1
$optionsLayout.Padding = New-Object System.Windows.Forms.Padding(8)
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$optionsLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

$filterLabel = New-Object System.Windows.Forms.Label
$filterLabel.Text = "Filtr:"
$filterLabel.AutoSize = $true
$filterLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left

$filterCombo = New-Object System.Windows.Forms.ComboBox
$filterCombo.Width = 180
$filterCombo.DropDownStyle = "DropDownList"
$filterCombo.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$filterCombo.Margin = New-Object System.Windows.Forms.Padding(6, 0, 12, 0)
[void]$filterCombo.Items.AddRange(@("Wszystkie", "Tylko włączone", "Tylko wyłączone"))
$filterCombo.SelectedIndex = 0

$sortCheckBox = New-Object System.Windows.Forms.CheckBox
$sortCheckBox.Text = "Sortuj alfabetycznie"
$sortCheckBox.Checked = $true
$sortCheckBox.AutoSize = $true
$sortCheckBox.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$sortCheckBox.Margin = New-Object System.Windows.Forms.Padding(0, 2, 12, 2)

$button = New-Object System.Windows.Forms.Button
 $button.Text = "Sprawdź konta"
$button.AutoSize = $true
$button.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$button.Anchor = [System.Windows.Forms.AnchorStyles]::Right
$button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)

$exportBtn = New-Object System.Windows.Forms.Button
$exportBtn.Text = "Eksport CSV"
$exportBtn.Enabled = $false
$exportBtn.AutoSize = $true
$exportBtn.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$exportBtn.Anchor = [System.Windows.Forms.AnchorStyles]::Right

$spacer = New-Object System.Windows.Forms.Panel
$spacer.Dock = [System.Windows.Forms.DockStyle]::Fill

$optionsLayout.Controls.Add($filterLabel, 0, 0)
$optionsLayout.Controls.Add($filterCombo, 1, 0)
$optionsLayout.Controls.Add($sortCheckBox, 2, 0)
$optionsLayout.Controls.Add($spacer, 3, 0)
$optionsLayout.Controls.Add($button, 4, 0)
$optionsLayout.Controls.Add($exportBtn, 5, 0)
$optionsGroup.Controls.Add($optionsLayout)
$layout.Controls.Add($optionsGroup, 0, 1)
#endregion Options Section

#region Output Section
$outputGroup = New-Object System.Windows.Forms.GroupBox
 $outputGroup.Text = "Wyniki zapytań"
$outputGroup.Dock = [System.Windows.Forms.DockStyle]::Fill
$outputGroup.Padding = New-Object System.Windows.Forms.Padding(8)

$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.ReadOnly = $true
$outputBox.ScrollBars = "Both"
$outputBox.WordWrap = $false
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$outputBox.MinimumSize = New-Object System.Drawing.Size(0, 150)
$outputBox.Height = 150
$outputBox.Dock = [System.Windows.Forms.DockStyle]::Fill

$outputGroup.Controls.Add($outputBox)
$layout.Controls.Add($outputGroup, 0, 2)
#endregion Output Section

#region Status Section
$statsLabel = New-Object System.Windows.Forms.Label
 $statsLabel.Text = "Znalezione: 0 | Włączone: 0 | Wyłączone: 0 | Błędy: 0"
$statsLabel.Dock = [System.Windows.Forms.DockStyle]::Fill
$statsLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$statsLabel.AutoSize = $false

$statusLayout = New-Object System.Windows.Forms.TableLayoutPanel
$statusLayout.Dock = [System.Windows.Forms.DockStyle]::Fill
$statusLayout.ColumnCount = 1
$statusLayout.RowCount = 1
$statusLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$statusLayout.Margin = New-Object System.Windows.Forms.Padding(8, 10, 8, 0)

$statusLayout.Controls.Add($statsLabel, 0, 0)
$layout.Controls.Add($statusLayout, 0, 3)
#endregion Status Section

#endregion GUI Construction

#region Data Helpers
$script:resultsForExport = @()

function Write-Line ([string]$text, [System.Drawing.Color]$color) {
    if ($null -eq $color) { $color = [System.Drawing.Color]::Black }
    $outputBox.SelectionColor = $color
    $outputBox.AppendText($text + "`r`n")
    $outputBox.SelectionColor = [System.Drawing.Color]::Black
}
#endregion Data Helpers

#region Rendering
function Render-Results {
    $outputBox.Clear()
    if (-not $script:resultsForExport -or $script:resultsForExport.Count -eq 0) { return }

    $results = $script:resultsForExport
    if ($sortCheckBox.Checked) {
        $results = $results | Sort-Object Login
    }

    $filter = $filterCombo.SelectedItem
    $filterAll = ($filter -eq 'Wszystkie')
    $filterEnabled = ($filter -eq 'Tylko wlaczone' -or  $filter -eq 'Tylko włączone')
    $filterDisabled = ($filter -eq 'Tylko wylaczone' -or  $filter -eq 'Tylko wyłączone')
    $L = [char]0x0141  # Ł
    $Aog = [char]0x0104 # Ą

    foreach ($r in $results) {
        $show = switch ($true) {
            { $r.Enabled -eq $null } { $filterAll; break }
            { $filterAll } { $true; break }
            { $filterEnabled -and $r.Enabled } { $true; break }
            { $filterDisabled -and -not $r.Enabled } { $true; break }
            default { $false }
        }

        if ($show) {
            if ($r.Enabled -eq $null) {
                $blad = "B" + $L + $Aog + "D"
                Write-Line ("Konto: {0} | {1} | {2}: {3}" -f $r.Login, ($r.Name ?? "-"), $blad, ($r.Error ?? "Nie znaleziono")) ([System.Drawing.Color]::Black)
            }
            elseif ($r.Enabled) {
                $wlaczone = "W" + $L + $Aog + "CZONE"
                Write-Line ("Konto: {0} | {1} | Status: {2}" -f $r.Login, ($r.Name ?? "-"), $wlaczone) ([System.Drawing.Color]::Green)
            }
            else {
                $wylaczone = "WY" + $L + $Aog + "CZONE"
                Write-Line ("Konto: {0} | {1} | Status: {2}" -f $r.Login, ($r.Name ?? "-"), $wylaczone) ([System.Drawing.Color]::Red)
            }
        }
    }
}
#endregion Rendering

#region Processing Logic
function Invoke-AccountCheck {
    $button.Enabled = $false
    $exportBtn.Enabled = $false
    $outputBox.Clear()
    $statsLabel.Text = "Przetwarzanie..."

    $logins = @($inputBox.Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $logins = @($logins | Select-Object -Unique)

    if (-not $logins -or $logins.Count -eq 0) {
        $statsLabel.Text = "Brak loginów do sprawdzenia"
        $button.Enabled = $true
        return
    }

    $total = [math]::Max(1, $logins.Count)
    $results = New-Object System.Collections.Generic.List[object]

    for ($i = 0; $i -lt $logins.Count; $i++) {
        $login = $logins[$i]
        $rec = [PSCustomObject]@{
            Login   = $login
            Enabled = $null
            Name    = $null
            Error   = $null
        }

        try {
            $user = Get-ADUser -Identity $login -Properties Enabled, Name -ErrorAction Stop
            if ($null -ne $user) {
                $rec.Enabled = [bool]$user.Enabled
                $rec.Name = $user.Name
            }
            else {
                $rec.Error = "NotFound"
            }
        }
        catch {
            $rec.Error = $_.Exception.Message
        }

        $results.Add($rec) | Out-Null
        [System.Windows.Forms.Application]::DoEvents()
    }

    $results = $results.ToArray()
    if ($sortCheckBox.Checked) {
        $results = $results | Sort-Object Login
    }

    $enabledCnt  = ($results | Where-Object { $_.Enabled -eq $true }).Count
    $disabledCnt = ($results | Where-Object { $_.Enabled -eq $false }).Count
    $errorCnt    = ($results | Where-Object { $_.Enabled -eq $null }).Count
    $foundCnt    = $enabledCnt + $disabledCnt

    $script:resultsForExport = @()
    foreach ($r in $results) {
        $script:resultsForExport += [PSCustomObject]@{
            Login   = $r.Login
            Name    = $r.Name
            Enabled = $(if ($r.Enabled -eq $null) { $null } else { [bool]$r.Enabled })
            Error   = $r.Error
        }
    }

    $statsLabel.Text = "Znalezione: {0} | Włączone: {1} | Wyłączone: {2} | Błędy: {3}" -f $foundCnt, $enabledCnt, $disabledCnt, $errorCnt
    $exportBtn.Enabled = $script:resultsForExport.Count -gt 0
    Render-Results
    $button.Enabled = $true
}
#endregion Processing Logic

#region Event Handlers
$button.Add_Click({ Invoke-AccountCheck })

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

$inputBox.Add_TextChanged({
    $lines = $inputBox.Text -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $uniq = $lines | Select-Object -Unique
    $countLabel.Text = "Liczba kont: {0} | Unikalne: {1}" -f $lines.Count, $uniq.Count
})

# Odświeżanie widoku po zmianie filtra/sortowania
$filterCombo.Add_SelectedIndexChanged({ Render-Results })
$sortCheckBox.Add_CheckedChanged({ Render-Results })
#endregion Event Handlers

#region Application Start
[void]$form.ShowDialog()
#endregion Application Start



