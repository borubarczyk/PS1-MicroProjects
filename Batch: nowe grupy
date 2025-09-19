# Created by BK
#Requires -Modules ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Ensure-ADModule {
    if (-not (Get-Module -ListAvailable ActiveDirectory)) {
        throw "Brak modułu ActiveDirectory (RSAT)."
    }
    Import-Module ActiveDirectory -ErrorAction Stop
}

function Select-OU {
    $ous = Get-ADOrganizationalUnit -Filter * -Properties CanonicalName |
      Sort-Object CanonicalName |
      Select-Object @{n='OU';e={$_.Name}},
                    @{n='DN';e={$_.DistinguishedName}},
                    @{n='Canonical';e={$_.CanonicalName}}
    $sel = $ous | Out-GridView -Title "Wybierz OU docelowe" -PassThru
    if (-not $sel) { throw "Anulowano wybór OU." }
    $sel.DN
}

function Normalize-Sam([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return "" }
    $sam = ($name -replace '[\\/:"\*\?\<\>\|;=,\+\@\[\]\(\)\{\}]','')
    $sam = $sam -replace '\s+','_'
    if ($sam.Length -gt 20) { $sam = $sam.Substring(0,20) }
    return $sam
}

function New-GroupsUI([string]$OuDn) {
    $form                = New-Object Windows.Forms.Form
    $form.Text           = "Batch: nowe grupy w OU = $OuDn"
    $form.StartPosition  = 'CenterScreen'
    $form.Size           = [Drawing.Size]::new(900, 520)
    $form.TopMost        = $true
    $font                = New-Object Drawing.Font('Segoe UI',10)

    $lbl = New-Object Windows.Forms.Label
    $lbl.Text = "Podaj grupy poniżej (możesz wkleić z Excela: Nazwa [TAB] Opis [TAB] Scope [TAB] Kategoria)"
    $lbl.AutoSize = $true; $lbl.Font=$font; $lbl.Location='10,10'
    $form.Controls.Add($lbl)

    $grid = New-Object Windows.Forms.DataGridView
    $grid.Location = '10,40'
    $grid.Size     = [Drawing.Size]::new(860,380)
    $grid.Font     = $font
    $grid.AllowUserToAddRows = $true
    $grid.AllowUserToResizeRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = 'CellSelect'
    $grid.AutoSizeColumnsMode = 'Fill'

    # Kolumny
    $colName = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $colName.HeaderText = 'Nazwa (Name)'; $colName.Name='Name'; $colName.FillWeight=160

    $colSam = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $colSam.HeaderText = 'sAMAccountName'; $colSam.Name='sAM'; $colSam.FillWeight=120

    $colDesc = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $colDesc.HeaderText = 'Opis (description)'; $colDesc.Name='Desc'; $colDesc.FillWeight=220

    $colScope = New-Object Windows.Forms.DataGridViewComboBoxColumn
    $colScope.HeaderText='Scope'; $colScope.Name='Scope'
    $colScope.Items.AddRange(@('Global','DomainLocal','Universal')); $colScope.FillWeight=90

    $colCat = New-Object Windows.Forms.DataGridViewComboBoxColumn
    $colCat.HeaderText='Kategoria'; $colCat.Name='Category'
    $colCat.Items.AddRange(@('Security','Distribution')); $colCat.FillWeight=90

    $colRes = New-Object Windows.Forms.DataGridViewTextBoxColumn
    $colRes.HeaderText='Status'; $colRes.Name='Result'; $colRes.ReadOnly=$true; $colRes.FillWeight=120

    $grid.Columns.AddRange(
    [System.Windows.Forms.DataGridViewColumn[]]@(
        $colName,$colSam,$colDesc,$colScope,$colCat,$colRes
    )
)
    $form.Controls.Add($grid)

    # Auto-generowanie sAM z Name
    $grid.add_CellEndEdit({
        param($s,$e)
        if ($grid.Columns[$e.ColumnIndex].Name -eq 'Name') {
            $val = $grid.Rows[$e.RowIndex].Cells['Name'].Value
            if ($null -ne $val -and "$val".Trim().Length -gt 0) {
                $grid.Rows[$e.RowIndex].Cells['sAM'].Value = Normalize-Sam "$val"
            }
        }
    })

    # Przyciski
    $btnPaste = New-Object Windows.Forms.Button
    $btnPaste.Text='Wklej ze schowka'; $btnPaste.Font=$font
    $btnPaste.Location='10,430'; $btnPaste.Size=[Drawing.Size]::new(150,35)

    $btnAdd = New-Object Windows.Forms.Button
    $btnAdd.Text='Dodaj 10 wierszy'; $btnAdd.Font=$font
    $btnAdd.Location='170,430'; $btnAdd.Size=[Drawing.Size]::new(150,35)

    $btnCreate = New-Object Windows.Forms.Button
    $btnCreate.Text='Utwórz wszystkie'; $btnCreate.Font=$font
    $btnCreate.Location='700,430'; $btnCreate.Size=[Drawing.Size]::new(170,35)

    $form.Controls.AddRange(@($btnPaste,$btnAdd,$btnCreate))

    # Wklej: zakładamy TSV (Excel), kolumny: Name [tab] Desc [tab] Scope [tab] Category
    $btnPaste.Add_Click({
        try {
            $txt = [Windows.Forms.Clipboard]::GetText()
            if ([string]::IsNullOrWhiteSpace($txt)) { return }
            $lines = $txt -split "(`r`n|`n|`r)"
            foreach ($ln in $lines) {
                if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                $parts = $ln -split "`t"
                $name  = $parts[0].Trim()
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $desc  = if ($parts.Count -ge 2) { $parts[1] } else { "" }
                $scope = if ($parts.Count -ge 3 -and @('Global','DomainLocal','Universal') -contains $parts[2]) { $parts[2] } else { 'Global' }
                $cat   = if ($parts.Count -ge 4 -and @('Security','Distribution') -contains $parts[3]) { $parts[3] } else { 'Security' }

                $idx = $grid.Rows.Add()
                $grid.Rows[$idx].Cells['Name'].Value = $name
                $grid.Rows[$idx].Cells['sAM'].Value  = Normalize-Sam $name
                $grid.Rows[$idx].Cells['Desc'].Value = $desc
                $grid.Rows[$idx].Cells['Scope'].Value= $scope
                $grid.Rows[$idx].Cells['Category'].Value = $cat
            }
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message,"Błąd wklejania") | Out-Null
        }
    })

    $btnAdd.Add_Click({
        1..10 | ForEach-Object { [void]$grid.Rows.Add() }
    })

    $btnCreate.Add_Click({
        $btnCreate.Enabled = $false
        try {
            for ($i=0; $i -lt $grid.Rows.Count; $i++) {
                $row = $grid.Rows[$i]
                if ($row.IsNewRow) { continue }

                $name = "$($row.Cells['Name'].Value)".Trim()
                $sam  = "$($row.Cells['sAM'].Value)".Trim()
                $desc = "$($row.Cells['Desc'].Value)".Trim()
                $scp  = "$($row.Cells['Scope'].Value)"
                $cat  = "$($row.Cells['Category'].Value)"

                if ([string]::IsNullOrWhiteSpace($name)) { $row.Cells['Result'].Value = "Pominięto: brak Nazwy"; continue }
                if ([string]::IsNullOrWhiteSpace($sam))  { $sam = Normalize-Sam $name }

                try {
                    if (Get-ADGroup -LDAPFilter "(sAMAccountName=$sam)" -ErrorAction SilentlyContinue) {
                        $row.Cells['Result'].Value = "Istnieje (sAM=$sam)"
                        continue
                    }

                    New-ADGroup -Name $name `
                        -SamAccountName $sam `
                        -GroupScope $scp `
                        -GroupCategory $cat `
                        -Path $OuDn `
                        -Description $desc `
                        -ErrorAction Stop

                    $row.Cells['Result'].Value = "OK"
                }
                catch {
                    $row.Cells['Result'].Value = "Błąd: $($_.Exception.Message)"
                }
            }
            [Windows.Forms.MessageBox]::Show("Zakończono tworzenie grup.","Gotowe",[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        }
        finally {
            $btnCreate.Enabled = $true
        }
    })

    [void]$form.ShowDialog()
}

# === MAIN ===
try {
    Ensure-ADModule
    $ou = Select-OU
    New-GroupsUI -OuDn $ou
}
catch {
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Błąd",[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    throw
}
