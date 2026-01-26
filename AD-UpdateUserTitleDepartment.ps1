#requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Uniwersalny kreator GUI do hurtowej zmiany atrybutów kont Active Directory.
.DESCRIPTION
    Narzędzie prowadzi operatora przez pięć etapów:
        1. Import danych z CSV i ich podgląd.
        2. Mapowanie kolumn CSV na atrybuty AD.
        3. Weryfikację istnienia kont (kluczem jest zawsze samAccountName).
        4. Podsumowanie zmian (co było / co będzie) i wykonanie aktualizacji.
        5. Możliwość cofnięcia zapisanych zmian na podstawie pliku rollback.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Warning "Uruchom skrypt w trybie STA (powershell.exe -STA), aby uniknąć problemów z GUI."
}

$ErrorActionPreference = 'Stop'

function Initialize-ActiveDirectoryModule {
    try {
        if (-not (Get-Module -Name ActiveDirectory)) {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Nie można załadować modułu ActiveDirectory.`nBłąd: $($_.Exception.Message)",
            "Błąd krytyczny",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        exit 1
    }
}

function Initialize-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -Path $Path)) {
        [void](New-Item -Path $Path -ItemType Directory -Force)
    }
}

function Show-DialogMessage {
    param(
        [Parameter(Mandatory)][string]$Text,
        [string]$Title = "Informacja",
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    [System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    ) | Out-Null
}

function Get-CsvValue {
    param(
        [Parameter(Mandatory)][psobject]$Row,
        [Parameter(Mandatory)][string]$Column
    )
    $prop = $Row.PSObject.Properties[$Column]
    if ($prop) { return $prop.Value }
    return $null
}

function ConvertTo-PlainString {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Array]) {
        $parts = @()
        foreach ($item in $Value) {
            if ($null -eq $item) { continue }
            $parts += [string]$item
        }
        return ($parts -join ';')
    }
    return [string]$Value
}

function Format-DisplayValue {
    param([object]$Value)
    $text = ConvertTo-PlainString -Value $Value
    if ([string]::IsNullOrWhiteSpace($text)) { return "(pusto)" }
    return $text
}

function Get-CsvDelimiter {
    param([Parameter(Mandatory)][string]$Path)
    $firstLine = (Get-Content -Path $Path -TotalCount 1 -ErrorAction Stop)
    if (-not $firstLine) { return ',' }
    $candidates = @{
        ';' = ($firstLine -split ';').Count
        ',' = ($firstLine -split ',').Count
        "`t" = ($firstLine -split "`t").Count
        '|' = ($firstLine -split '\|').Count
    }
    return ($candidates.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1).Key
}

function ConvertTo-DataTable {
    param(
        [Parameter(Mandatory)][System.Collections.IEnumerable]$InputObject,
        [Parameter(Mandatory)][string[]]$Columns,
        [int]$MaxRows = 500
    )
    $table = New-Object System.Data.DataTable
    foreach ($col in $Columns) {
        [void]$table.Columns.Add($col)
    }
    $rowCounter = 0
    foreach ($row in $InputObject) {
        $dataRow = $table.NewRow()
        foreach ($col in $Columns) {
            $value = Get-CsvValue -Row $row -Column $col
            $dataRow[$col] = [string]$value
        }
        [void]$table.Rows.Add($dataRow)
        $rowCounter++
        if ($rowCounter -ge $MaxRows) { break }
    }
    return $table
}

$maxPreviewRows = 500

function Clear-CsvPreview {
    if ($null -ne $gridPreview) {
        $gridPreview.Rows.Clear()
        $gridPreview.Columns.Clear()
    }
    if ($null -ne $lblPreviewPlaceholder) {
        $lblPreviewPlaceholder.Visible = $true
        $lblPreviewPlaceholder.BringToFront()
    }
    if ($null -ne $lblPreviewStatus) {
        $lblPreviewStatus.Text = "Podgląd: brak danych."
    }
}

function Show-CsvPreview {
    if (-not $gridPreview) { return }
    $gridPreview.SuspendLayout()
    try {
        $gridPreview.Rows.Clear()
        $gridPreview.Columns.Clear()
        if (-not $script:AppState.Data -or $script:AppState.Data.Count -eq 0) {
            Clear-CsvPreview
            return
        }
        foreach ($colName in $script:AppState.Columns) {
            $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $col.Name = $colName
            $col.HeaderText = $colName
            $col.ReadOnly = $true
            $col.AutoSizeMode = 'DisplayedCells'
            $col.MinimumWidth = 60
            [void]$gridPreview.Columns.Add($col)
        }
        $limit = [Math]::Min($maxPreviewRows, $script:AppState.Data.Count)
        for ($idx = 0; $idx -lt $limit; $idx++) {
            $row = $script:AppState.Data[$idx]
            $values = foreach ($colName in $script:AppState.Columns) {
                [string](Get-CsvValue -Row $row -Column $colName)
            }
            [void]$gridPreview.Rows.Add($values)
        }
        $gridPreview.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
        if ($limit -gt 0) {
            if ($null -ne $lblPreviewPlaceholder) {
                $lblPreviewPlaceholder.Visible = $false
                $lblPreviewPlaceholder.SendToBack()
            }
            if ($null -ne $lblPreviewStatus) {
                $lblPreviewStatus.Text = "Podgląd: $limit wierszy (max $maxPreviewRows)."
            }
        } else {
            Clear-CsvPreview
        }
        Write-AppLog "Podglad CSV: $limit rekordow wyswietlonych."
    } finally {
        $gridPreview.ResumeLayout()
    }
}

function Get-DefaultKeyColumn {
    if (-not $script:AppState.Columns -or $script:AppState.Columns.Count -eq 0) { return $null }
    foreach ($name in $script:AppState.Columns) {
        if ($name -match '^(?i)sam.*account') { return $name }
    }
    return $script:AppState.Columns | Select-Object -First 1
}

function Update-MappingGrid {
    param([Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid)
    $Grid.Rows.Clear()
    $comboCol = [System.Windows.Forms.DataGridViewComboBoxColumn]$Grid.Columns['CsvColumn']
    $comboCol.Items.Clear()
    if ($script:AppState.Columns) {
        $comboCol.Items.AddRange([object[]]$script:AppState.Columns)
    }
    if (-not $script:AppState.Columns -or $script:AppState.Columns.Count -eq 0) { return }
    $keyColumn = Get-DefaultKeyColumn
    $rowIndex = $Grid.Rows.Add($true, $keyColumn, 'samAccountName', 'Klucz konta (wymagany)')
    $row = $Grid.Rows[$rowIndex]
    $row.Tag = 'key'
    $row.Cells['AdAttribute'].ReadOnly = $true
    $row.Cells['Use'].ReadOnly = $true
    $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(235,235,235)
}

function Add-MappingRow {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid,
        [string]$CsvColumn,
        [string]$Attribute
    )
    $index = $Grid.Rows.Add($true, $CsvColumn, $Attribute, '')
    $Grid.Rows[$index].Tag = 'attr'
    return $Grid.Rows[$index]
}

function Set-MappingDefaults {
    param([Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid)
    foreach ($column in $script:AppState.Columns) {
        if (-not $column -or $column -ieq 'samAccountName') { continue }
        if (-not ($script:AllowedAttributes -contains $column)) { continue }
        $exists = $false
        foreach ($row in $Grid.Rows) {
            if ($row.IsNewRow) { continue }
            $attr = [string]$row.Cells['AdAttribute'].Value
            if ($attr -and $attr -ieq $column) {
                $exists = $true
                break
            }
        }
        if (-not $exists) {
            Add-MappingRow -Grid $Grid -CsvColumn $column -Attribute $column | Out-Null
        }
    }
}

function Get-MappingRows {
    param([Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid)
    $result = @()
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        $csvColumn = [string]$row.Cells['CsvColumn'].Value
        $attr = [string]$row.Cells['AdAttribute'].Value
        $isKey = ($row.Tag -eq 'key') -or ($attr -and $attr -ieq 'samAccountName')
        $useValue = if ($isKey) { $true } else { [bool]$row.Cells['Use'].Value }
        $result += [pscustomobject]@{
            Use = $useValue
            CsvColumn = $csvColumn
            Attribute = if ($isKey) { 'samAccountName' } else { $attr }
            IsKey = $isKey
        }
    }
    return $result
}

function Test-MappingDefinition {
    param([Parameter(Mandatory)][array]$Mappings)
    if (-not $Mappings -or $Mappings.Count -eq 0) {
        Show-DialogMessage "Brak zdefiniowanych mapowań. Dodaj co najmniej jedną kolumnę." "Mapowanie" ([System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    $keyEntry = $Mappings | Where-Object { $_.IsKey } | Select-Object -First 1
    if (-not $keyEntry -or [string]::IsNullOrWhiteSpace($keyEntry.CsvColumn)) {
        Show-DialogMessage "Wskaż kolumnę zawierającą samAccountName." "Mapowanie" ([System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    $updateEntries = $Mappings | Where-Object {
        (-not $_.IsKey) -and $_.Use -and -not [string]::IsNullOrWhiteSpace($_.Attribute) -and -not [string]::IsNullOrWhiteSpace($_.CsvColumn)
    }
    if (-not $updateEntries -or $updateEntries.Count -eq 0) {
        Show-DialogMessage "Brak aktywnych mapowań atrybutów. Wybierz kolumny i nazwy atrybutów AD." "Mapowanie" ([System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    return $true
}

function Invoke-Verification {
    param(
        [Parameter(Mandatory)][array]$Mappings,
        [Parameter(Mandatory)][System.Windows.Forms.ListView]$ListView,
        [Parameter(Mandatory)][System.Windows.Forms.Label]$SummaryLabel
    )
    $ListView.BeginUpdate()
    $ListView.Items.Clear()
    try {
        if (-not (Test-MappingDefinition -Mappings $Mappings)) { return $false }
        $keyMapping = $Mappings | Where-Object { $_.IsKey } | Select-Object -First 1
        $updateMappings = $Mappings | Where-Object { (-not $_.IsKey) -and $_.Use }
        $attributes = ($updateMappings | ForEach-Object { $_.Attribute } | Sort-Object -Unique)
        $attributes = $attributes + @('samAccountName','DistinguishedName')
        $attributes = $attributes | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique

        $uniqueRows = @{}
        $missingKeyRows = 0
        foreach ($row in $script:AppState.Data) {
            $sam = Get-CsvValue -Row $row -Column $keyMapping.CsvColumn
            $sam = if ($null -eq $sam) { '' } else { [string]$sam }
            if ([string]::IsNullOrWhiteSpace($sam)) {
                $missingKeyRows++
                $item = New-Object System.Windows.Forms.ListViewItem("(brak)")
                $item.SubItems.Add("CSV") | Out-Null
                $item.SubItems.Add("Brak wartości samAccountName w wierszu pliku.") | Out-Null
                [void]$ListView.Items.Add($item)
                continue
            }
            $key = $sam.ToLowerInvariant()
            $uniqueRows[$key] = [pscustomobject]@{ Sam = $sam; Row = $row }
        }

        $script:AppState.Verified = @()
        $found = 0
        $problems = $missingKeyRows
        foreach ($entry in $uniqueRows.GetEnumerator() | Sort-Object { $_.Value.Sam }) {
            $sam = $entry.Value.Sam
            try {
                $user = Get-ADUser -Identity $sam -Properties $attributes -ErrorAction Stop
                $current = @{}
                foreach ($attr in $attributes) {
                    $current[$attr] = $user.$attr
                }
                $script:AppState.Verified += [pscustomobject]@{
                    SamAccountName = $sam
                    Exists = $true
                    DistinguishedName = $user.DistinguishedName
                    CurrentValues = $current
                    Row = $entry.Value.Row
                }
                $item = New-Object System.Windows.Forms.ListViewItem($sam)
                $item.SubItems.Add("OK") | Out-Null
                $item.SubItems.Add($user.DistinguishedName) | Out-Null
                [void]$ListView.Items.Add($item)
                $found++
            } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
                $script:AppState.Verified += [pscustomobject]@{
                    SamAccountName = $sam
                    Exists = $false
                    DistinguishedName = $null
                    CurrentValues = @{}
                    Row = $entry.Value.Row
                }
                $item = New-Object System.Windows.Forms.ListViewItem($sam)
                $item.SubItems.Add("Nie znaleziono") | Out-Null
                $item.SubItems.Add("Konto nie istnieje w AD.") | Out-Null
                [void]$ListView.Items.Add($item)
                $problems++
            } catch {
                $script:AppState.Verified += [pscustomobject]@{
                    SamAccountName = $sam
                    Exists = $false
                    DistinguishedName = $null
                    CurrentValues = @{}
                    Row = $entry.Value.Row
                }
                $item = New-Object System.Windows.Forms.ListViewItem($sam)
                $item.SubItems.Add("Błąd") | Out-Null
                $item.SubItems.Add($_.Exception.Message) | Out-Null
                [void]$ListView.Items.Add($item)
                $problems++
            }
        }

        $SummaryLabel.Text = "Znaleziono $found kont(a). Problemy: $problems."
        Write-AppLog "Weryfikacja zakończona. OK: $found, Problemy: $problems."
        return $true
    } finally {
        $ListView.EndUpdate()
    }
}

function Build-ChangePreview {
    param([Parameter(Mandatory)][array]$Mappings)
    $script:AppState.ChangePreview = @()
    $updateMappings = $Mappings | Where-Object { (-not $_.IsKey) -and $_.Use }
    foreach ($entry in $script:AppState.Verified) {
        if (-not $entry.Exists) { continue }
        foreach ($map in $updateMappings) {
            $newRaw = Get-CsvValue -Row $entry.Row -Column $map.CsvColumn
            $newValue = if ($null -eq $newRaw) { '' } else { [string]$newRaw }
            $newTrim = $newValue.Trim()
            $current = $entry.CurrentValues[$map.Attribute]
            $currentString = ConvertTo-PlainString -Value $current
            $currentTrim = $currentString.Trim()

            $shouldUpdate = $false
            $statusText = ""
            $newDisplay = $newTrim
            if ([string]::IsNullOrWhiteSpace($newTrim)) {
                $newDisplay = ''
                if ($script:AppState.AllowEmptyClear) {
                    if ($currentTrim.Length -gt 0) {
                        $shouldUpdate = $true
                        $statusText = "Do czyszczenia"
                    } else {
                        $statusText = "Bez zmian (już pusto)"
                    }
                } else {
                    if ($currentTrim.Length -gt 0) {
                        $statusText = "Pominięto (puste w CSV, czyszczenie wyłączone)"
                    } else {
                        $statusText = "Bez zmian (pusto)"
                    }
                }
            } else {
                if ($currentTrim -cne $newTrim) {
                    $shouldUpdate = $true
                    $statusText = "Do zmiany"
                } else {
                    $statusText = "Bez zmian (identyczna wartość)"
                }
            }
            if ($shouldUpdate -and -not $statusText) {
                $statusText = "Do zmiany"
            } elseif (-not $shouldUpdate -and -not $statusText) {
                $statusText = "Bez zmian"
            }

            $script:AppState.ChangePreview += [pscustomobject]@{
                SamAccountName   = $entry.SamAccountName
                Attribute        = $map.Attribute
                CurrentValue     = $current
                NewValue         = $newTrim
                DistinguishedName = $entry.DistinguishedName
                CurrentDisplay   = $currentString
                NewDisplay       = $newDisplay
                WillUpdate       = $shouldUpdate
                StatusText       = $statusText
            }
        }
    }
}

function Update-ChangesGrid {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid,
        [Parameter(Mandatory)][System.Windows.Forms.Label]$SummaryLabel
    )
    $Grid.Rows.Clear()
    foreach ($change in $script:AppState.ChangePreview) {
        $currentText = Format-DisplayValue -Value $change.CurrentValue
        $newText = Format-DisplayValue -Value $change.NewValue
        $index = $Grid.Rows.Add(
            $change.SamAccountName,
            $change.Attribute,
            $currentText,
            $newText,
            $change.StatusText
        )
        $row = $Grid.Rows[$index]
        $row.Tag = $change
        if ($change.WillUpdate) {
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
        } else {
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DimGray
        }
    }
    $userCount = ($script:AppState.ChangePreview | Select-Object -ExpandProperty SamAccountName -Unique | Measure-Object).Count
    $pending = @($script:AppState.ChangePreview | Where-Object { $_.WillUpdate }).Count
    $SummaryLabel.Text = "Podgląd obejmuje $($script:AppState.ChangePreview.Count) pozycji (do zmiany: $pending) na $userCount kontach."
    if ($Grid.Columns.Count -gt 0) {
        $Grid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
    }
}
function Update-ChangeRowStatus {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid,
        [Parameter(Mandatory)][pscustomobject]$Change,
        [Parameter(Mandatory)][string]$StatusText
    )
    foreach ($row in $Grid.Rows) {
        if ($row.IsNewRow) { continue }
        if ($row.Tag -eq $Change) {
            $row.Cells['Status'].Value = $StatusText
            break
        }
    }
}

function New-RollbackSnapshot {
    param([Parameter(Mandatory)][System.Collections.IEnumerable]$Changes)
    $effective = @()
    if ($Changes) {
        $effective = @($Changes | Where-Object { $_.WillUpdate })
    }
    if (-not $effective -or $effective.Count -eq 0) { return $null }

    $snapshot = [ordered]@{
        CreatedAt = Get-Date
        CreatedBy = $env:USERNAME
        CsvPath = $script:AppState.CsvPath
        AllowEmptyClear = $script:AppState.AllowEmptyClear
        Items = @()
    }

    foreach ($group in ($effective | Group-Object -Property SamAccountName)) {
        $item = [ordered]@{
            SamAccountName = $group.Name
            DistinguishedName = $group.Group[0].DistinguishedName
            Changes = @()
        }
        foreach ($change in $group.Group) {
            $item.Changes += [ordered]@{
                Attribute = $change.Attribute
                OldValue = $change.CurrentValue
                NewValue = $change.NewValue
                OldDisplay = Format-DisplayValue -Value $change.CurrentValue
                NewDisplay = Format-DisplayValue -Value $change.NewValue
            }
        }
        $snapshot.Items += $item
    }

    $fileName = "AD-BulkUpdater_{0}.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss')
    $path = Join-Path $script:AppState.RollbackFolder $fileName
    $snapshot | ConvertTo-Json -Depth 6 | Out-File -FilePath $path -Encoding utf8
    Write-AppLog "Zapisano plik cofania: $path"
    return $path
}

function Update-RollbackList {
    param([Parameter(Mandatory)][System.Windows.Forms.ListView]$ListView)
    $ListView.BeginUpdate()
    $ListView.Items.Clear()
    try {
        if (-not (Test-Path -Path $script:AppState.RollbackFolder)) { return }
        $files = Get-ChildItem -Path $script:AppState.RollbackFolder -Filter 'AD-BulkUpdater_*.json' -File -ErrorAction SilentlyContinue |
                 Sort-Object -Property LastWriteTime -Descending
        foreach ($file in $files) {
            $meta = $null
            try {
                $meta = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json -Depth 5
            } catch {
                $meta = $null
            }
            $count = if ($meta -and $meta.Items) { $meta.Items.Count } else { 0 }
            $csvName = if ($meta -and $meta.CsvPath) { Split-Path $meta.CsvPath -Leaf } else { '' }
            $item = New-Object System.Windows.Forms.ListViewItem($file.LastWriteTime.ToString("yyyy-MM-dd HH:mm"))
            $item.SubItems.Add($file.Name) | Out-Null
            $item.SubItems.Add("Konta: $count $csvName") | Out-Null
            $item.Tag = $file.FullName
            [void]$ListView.Items.Add($item)
        }
    } finally {
        $ListView.EndUpdate()
    }
}

function Update-RollbackPreview {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid,
        [Parameter(Mandatory)][System.Windows.Forms.Label]$StatusLabel
    )
    $Grid.Rows.Clear()
    if (-not (Test-Path -Path $Path)) {
        $script:AppState.SelectedRollback = $null
        $StatusLabel.Text = "Nie znaleziono pliku cofania."
        return
    }
    try {
        $snapshot = Get-Content -Path $Path -Raw | ConvertFrom-Json -Depth 6
    } catch {
        Show-DialogMessage "Nie można odczytać pliku cofania.`n$($_.Exception.Message)" "Cofnięcie" ([System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    $script:AppState.SelectedRollback = $snapshot
    if (-not $snapshot.Items) {
        $StatusLabel.Text = "Plik nie zawiera żadnych zmian."
        return
    }
    foreach ($item in $snapshot.Items) {
        foreach ($change in $item.Changes) {
            $oldText = $change.OldDisplay
            if ([string]::IsNullOrWhiteSpace($oldText)) {
                $oldText = Format-DisplayValue -Value $change.OldValue
            }
            $newText = $change.NewDisplay
            if ([string]::IsNullOrWhiteSpace($newText)) {
                $newText = Format-DisplayValue -Value $change.NewValue
            }
            $rowIndex = $Grid.Rows.Add(
                $item.SamAccountName,
                $change.Attribute,
                $oldText,
                $newText,
                ''
            )
            $Grid.Rows[$rowIndex].Tag = [pscustomobject]@{ User = $item; Change = $change }
        }
    }
    $StatusLabel.Text = "Przywrócenie obejmie $($snapshot.Items.Count) kont(a)."
    if ($Grid.Columns.Count -gt 0) {
        $Grid.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
    }
}

function Invoke-Rollback {
    param(
        [Parameter(Mandatory)][pscustomobject]$Snapshot,
        [Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid,
        [Parameter(Mandatory)][System.Windows.Forms.Label]$StatusLabel
    )
    if (-not $Snapshot -or -not $Snapshot.Items) {
        Show-DialogMessage "Brak danych do cofnięcia." "Cofnięcie" ([System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $restored = 0
    foreach ($item in $Snapshot.Items) {
        $replace = @{}
        $clear = New-Object System.Collections.Generic.List[string]
        foreach ($change in $item.Changes) {
            $oldValue = $change.OldValue
            $oldText = ConvertTo-PlainString -Value $oldValue
            if ([string]::IsNullOrWhiteSpace($oldText)) {
                [void]$clear.Add($change.Attribute)
            } else {
                $replace[$change.Attribute] = $oldValue
            }
        }
        $status = "OK"
        try {
            if ($replace.Count -gt 0) {
                Set-ADUser -Identity $item.DistinguishedName -Replace $replace -ErrorAction Stop
            }
            if ($clear.Count -gt 0) {
                Set-ADUser -Identity $item.DistinguishedName -Clear ($clear.ToArray()) -ErrorAction Stop
            }
            $restored++
            Write-AppLog "Cofnięto zmiany dla $($item.SamAccountName)."
        } catch {
            $status = "Błąd: $($_.Exception.Message)"
            Write-AppLog "Błąd przy cofaniu $($item.SamAccountName): $($_.Exception.Message)"
        }
        foreach ($row in $Grid.Rows) {
            if ($row.IsNewRow) { continue }
            $tag = $row.Tag
            if ($tag -and $tag.User -eq $item) {
                $row.Cells['Status'].Value = $status
            }
        }
    }
    $StatusLabel.Text = "Cofnięto $restored kont."
}

function Invoke-ApplyChanges {
    param(
        [Parameter(Mandatory)][bool]$WhatIf,
        [Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid,
        [Parameter(Mandatory)][System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.ListView]$RollbackList = $null
    )
    $changes = @()
    if ($script:AppState.ChangePreview) {
        $changes = @($script:AppState.ChangePreview | Where-Object { $_.WillUpdate })
    }
    if (-not $changes -or $changes.Count -eq 0) {
        Show-DialogMessage "Brak zaplanowanych zmian. Upewnij się, że dane w zakładce 4 zawierają pozycje oznaczone jako 'Do zmiany'." "Aktualizacja" ([System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }

    $snapshotPath = $null
    if (-not $WhatIf) {
        $snapshotPath = New-RollbackSnapshot -Changes $changes
    }

    $processed = 0
    foreach ($group in ($changes | Group-Object -Property SamAccountName)) {
        $dn = $group.Group[0].DistinguishedName
        $replace = @{}
        $clearList = New-Object System.Collections.Generic.List[string]

        foreach ($change in $group.Group) {
            if ([string]::IsNullOrWhiteSpace($change.NewValue)) {
                [void]$clearList.Add($change.Attribute)
            } else {
                $replace[$change.Attribute] = $change.NewValue
            }
        }

        $status = "WHATIF"
        try {
            if (-not $WhatIf) {
                if ($replace.Count -gt 0) {
                    Set-ADUser -Identity $dn -Replace $replace -ErrorAction Stop
                }
                if ($clearList.Count -gt 0) {
                    Set-ADUser -Identity $dn -Clear ($clearList.ToArray()) -ErrorAction Stop
                }
                $status = "OK"
                $processed++
                $clearedNames = if ($clearList.Count -gt 0) { ($clearList.ToArray() -join ',') } else { '' }
                Write-AppLog "Zmieniono konto $($group.Name): $($replace.Keys -join ',') (czyszczenie: $clearedNames)"
            } else {
                Write-AppLog "Symulacja zmian dla konta $($group.Name)."
            }
        } catch {
            $status = "Błąd: $($_.Exception.Message)"
            Write-AppLog "Błąd podczas aktualizacji $($group.Name): $($_.Exception.Message)"
        }

        foreach ($change in $group.Group) {
            Update-ChangeRowStatus -Grid $Grid -Change $change -StatusText $status
        }
    }

    if ($WhatIf) {
        $StatusLabel.Text = "Symulacja zakończona. Sprawdź kolumnę Status."
    } else {
        $StatusLabel.Text = "Zmieniono $processed kont. Plik cofania: $snapshotPath"
        if ($RollbackList -and $snapshotPath) {
            Update-RollbackList -ListView $RollbackList
        }
    }
}
Initialize-ActiveDirectoryModule

$documentsPath = [Environment]::GetFolderPath('MyDocuments')
$baseFolder = Join-Path $documentsPath 'AD-BulkAttributeUpdater'
$logFolder = Join-Path $baseFolder 'Logs'
$rollbackFolder = Join-Path $baseFolder 'Rollback'
Initialize-Folder $baseFolder
Initialize-Folder $logFolder
Initialize-Folder $rollbackFolder

$script:AppState = [ordered]@{
    CsvPath = $null
    Data = @()
    Columns = @()
    Mapping = @()
    Verified = @()
    ChangePreview = @()
    AllowEmptyClear = $false
    LastDelimiter = ','
    RollbackFolder = $rollbackFolder
    LogFile = Join-Path $logFolder ("AD-BulkUpdater_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    SelectedRollback = $null
    VerificationStale = $true
}
"=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') START ===" | Set-Content -Path $script:AppState.LogFile -Encoding utf8

function Write-AppLog {
    param([Parameter(Mandatory)][string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $script:AppState.LogFile -Append -Encoding utf8
}

$script:AllowedAttributes = @(
    'title','department','description','company','mail','telephoneNumber','mobile','otherTelephone',
    'homePhone','ipPhone','pager','facsimileTelephoneNumber','physicalDeliveryOfficeName','streetAddress',
    'postOfficeBox','l','st','postalCode','co','countryCode','employeeID','employeeNumber','manager',
    'givenName','sn','displayName','initials','office','info','wWWHomePage',
    'extensionAttribute1','extensionAttribute2','extensionAttribute3','extensionAttribute4','extensionAttribute5',
    'extensionAttribute6','extensionAttribute7','extensionAttribute8','extensionAttribute9','extensionAttribute10',
    'extensionAttribute11','extensionAttribute12','extensionAttribute13','extensionAttribute14','extensionAttribute15'
)

$script:AttributeAutoComplete = New-Object System.Windows.Forms.AutoCompleteStringCollection
$script:AttributeAutoComplete.AddRange($script:AllowedAttributes)

$delimiterOptions = @(
    [pscustomobject]@{ Label = "Auto"; Value = "auto" },
    [pscustomobject]@{ Label = "Średnik (;)"; Value = ";" },
    [pscustomobject]@{ Label = "Przecinek (,)"; Value = "," },
    [pscustomobject]@{ Label = "Tabulator"; Value = "`t" },
    [pscustomobject]@{ Label = "Kreska (|)"; Value = "|" }
)

function Import-CsvData {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$DelimiterMode
    )
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "Nie wskazano ścieżki do pliku."
    }
    if (-not (Test-Path -Path $Path)) {
        throw "Plik nie istnieje: $Path"
    }
    $delimiter = if ($DelimiterMode -eq 'auto') { Get-CsvDelimiter -Path $Path } else { $DelimiterMode }
    $rows = Import-Csv -Path $Path -Delimiter $delimiter
    if (-not $rows -or $rows.Count -eq 0) {
        throw "Plik nie zawiera danych."
    }
    $columns = $rows[0].PSObject.Properties | Select-Object -ExpandProperty Name
    if (-not $columns -or $columns.Count -eq 0) {
        throw "Plik nie zawiera nagłówków."
    }
    $script:AppState.CsvPath = $Path
    $script:AppState.Data = $rows
    $script:AppState.Columns = $columns
    $script:AppState.LastDelimiter = $delimiter
    Write-AppLog "Zaimportowano CSV: $Path (separator '$delimiter', wiersze $($rows.Count))."
}
# --- GUI ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Bulk Attribute Updater"
$form.Size = New-Object System.Drawing.Size(1280, 860)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(1100, 720)
$form.KeyPreview = $true
$form.add_KeyDown({
    if ($_.KeyCode -eq 'Escape') { $form.Close() }
})

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$form.Controls.Add($tabControl)

# Tab 1: Import danych
$tabImport = New-Object System.Windows.Forms.TabPage
$tabImport.Text = "1. Import danych"
[void]$tabControl.TabPages.Add($tabImport)

$lblImport = New-Object System.Windows.Forms.Label
$lblImport.Text = "Wybierz plik CSV zawierający dane do aktualizacji."
$lblImport.AutoSize = $true
$lblImport.Location = New-Object System.Drawing.Point(15,15)
$tabImport.Controls.Add($lblImport)

$txtCsvPath = New-Object System.Windows.Forms.TextBox
$txtCsvPath.Location = New-Object System.Drawing.Point(15,45)
$txtCsvPath.Width = 650
$tabImport.Controls.Add($txtCsvPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Przeglądaj..."
$btnBrowse.Location = New-Object System.Drawing.Point(680,43)
$btnBrowse.Width = 100
$tabImport.Controls.Add($btnBrowse)

$lblDelimiter = New-Object System.Windows.Forms.Label
$lblDelimiter.Text = "Separator:"
$lblDelimiter.AutoSize = $true
$lblDelimiter.Location = New-Object System.Drawing.Point(800,47)
$tabImport.Controls.Add($lblDelimiter)

$comboDelimiter = New-Object System.Windows.Forms.ComboBox
$comboDelimiter.DropDownStyle = 'DropDownList'
$comboDelimiter.Location = New-Object System.Drawing.Point(870,43)
$comboDelimiter.Width = 150
$comboDelimiter.DisplayMember = 'Label'
$comboDelimiter.ValueMember = 'Value'
$comboDelimiter.DataSource = $delimiterOptions
$comboDelimiter.FormattingEnabled = $true
$comboDelimiter.add_Format({
    param($source,$formatEvent)
    $item = $formatEvent.ListItem
    if ($null -eq $item) { return }
    $labelProp = $item.PSObject.Properties['Label']
    if ($labelProp) {
        $formatEvent.Value = $labelProp.Value
    }
})
if ($comboDelimiter.Items.Count -gt 0) {
    $comboDelimiter.SelectedIndex = 0
}
$tabImport.Controls.Add($comboDelimiter)

$btnLoad = New-Object System.Windows.Forms.Button
$btnLoad.Text = "Importuj"
$btnLoad.Location = New-Object System.Drawing.Point(1040,43)
$btnLoad.Width = 100
$tabImport.Controls.Add($btnLoad)

$previewPanel = New-Object System.Windows.Forms.Panel
$previewPanel.Location = New-Object System.Drawing.Point(15,90)
$previewPanel.Size = New-Object System.Drawing.Size(1220,650)
$previewPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$previewPanel.BorderStyle = 'FixedSingle'
$previewPanel.BackColor = [System.Drawing.Color]::White
$tabImport.Controls.Add($previewPanel)

$gridPreview = New-Object System.Windows.Forms.DataGridView
$gridPreview.Dock = 'Fill'
$gridPreview.ReadOnly = $true
$gridPreview.AllowUserToAddRows = $false
$gridPreview.AllowUserToDeleteRows = $false
$gridPreview.AutoSizeColumnsMode = 'DisplayedCells'
$gridPreview.AutoGenerateColumns = $false
$gridPreview.RowHeadersVisible = $false
$gridPreview.BackgroundColor = [System.Drawing.Color]::White
$gridPreview.GridColor = [System.Drawing.Color]::LightGray
$gridPreview.BorderStyle = 'None'
$gridPreview.ScrollBars = 'Both'
$previewPanel.Controls.Add($gridPreview)

$lblPreviewPlaceholder = New-Object System.Windows.Forms.Label
$lblPreviewPlaceholder.Text = "Podgląd CSV pojawi się po imporcie (wyświetlane maks. 500 wierszy)."
$lblPreviewPlaceholder.Dock = 'Fill'
$lblPreviewPlaceholder.TextAlign = 'MiddleCenter'
$lblPreviewPlaceholder.ForeColor = [System.Drawing.Color]::Gray
$lblPreviewPlaceholder.BackColor = [System.Drawing.Color]::FromArgb(10,0,0,0)
$lblPreviewPlaceholder.Font = New-Object System.Drawing.Font($form.Font.FontFamily, 12, [System.Drawing.FontStyle]::Italic)
$previewPanel.Controls.Add($lblPreviewPlaceholder)
$lblPreviewPlaceholder.BringToFront()

$lblPreviewStatus = New-Object System.Windows.Forms.Label
$lblPreviewStatus.AutoSize = $true
$lblPreviewStatus.Location = New-Object System.Drawing.Point(15,745)
$lblPreviewStatus.Text = "Podgląd: brak danych."
$tabImport.Controls.Add($lblPreviewStatus)

$lblImportStatus = New-Object System.Windows.Forms.Label
$lblImportStatus.AutoSize = $true
$lblImportStatus.Location = New-Object System.Drawing.Point(15,765)
$lblImportStatus.Text = "Nie załadowano danych."
$tabImport.Controls.Add($lblImportStatus)

Clear-CsvPreview

# Tab 2: Mapowanie
$tabMapping = New-Object System.Windows.Forms.TabPage
$tabMapping.Text = "2. Mapowanie kolumn"
[void]$tabControl.TabPages.Add($tabMapping)

$lblMappingInfo = New-Object System.Windows.Forms.Label
$lblMappingInfo.Text = "Wskaż, która kolumna CSV odpowiada konkretnej właściwości AD. samAccountName jest wymagany jako klucz."
$lblMappingInfo.AutoSize = $true
$lblMappingInfo.Location = New-Object System.Drawing.Point(15,15)
$tabMapping.Controls.Add($lblMappingInfo)

$gridMapping = New-Object System.Windows.Forms.DataGridView
$gridMapping.Location = New-Object System.Drawing.Point(15,45)
$gridMapping.Size = New-Object System.Drawing.Size(1020,620)
$gridMapping.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$gridMapping.AllowUserToAddRows = $false
$gridMapping.AllowUserToDeleteRows = $false
$gridMapping.SelectionMode = 'FullRowSelect'
$gridMapping.MultiSelect = $true
$gridMapping.RowHeadersVisible = $false
$gridMapping.AutoSizeColumnsMode = 'Fill'
$tabMapping.Controls.Add($gridMapping)

$colUse = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colUse.HeaderText = "Użyj"
$colUse.Name = "Use"
$colUse.Width = 60
[void]$gridMapping.Columns.Add($colUse)

$colCsv = New-Object System.Windows.Forms.DataGridViewComboBoxColumn
$colCsv.HeaderText = "Kolumna CSV"
$colCsv.Name = "CsvColumn"
$colCsv.FlatStyle = 'Flat'
[void]$gridMapping.Columns.Add($colCsv)

$colAttr = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colAttr.HeaderText = "Atrybut AD"
$colAttr.Name = "AdAttribute"
$colAttr.Width = 180
[void]$gridMapping.Columns.Add($colAttr)

$colDesc = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDesc.HeaderText = "Opis"
$colDesc.Name = "Description"
$colDesc.ReadOnly = $true
$colDesc.Width = 250
[void]$gridMapping.Columns.Add($colDesc)

$btnAddMapping = New-Object System.Windows.Forms.Button
$btnAddMapping.Text = "Dodaj atrybut"
$btnAddMapping.Location = New-Object System.Drawing.Point(1050,45)
$btnAddMapping.Width = 180
$tabMapping.Controls.Add($btnAddMapping)

$btnRemoveMapping = New-Object System.Windows.Forms.Button
$btnRemoveMapping.Text = "Usuń zaznaczone"
$btnRemoveMapping.Location = New-Object System.Drawing.Point(1050,85)
$btnRemoveMapping.Width = 180
$tabMapping.Controls.Add($btnRemoveMapping)

$btnGuessMapping = New-Object System.Windows.Forms.Button
$btnGuessMapping.Text = "Dopasuj automatycznie"
$btnGuessMapping.Location = New-Object System.Drawing.Point(1050,125)
$btnGuessMapping.Width = 180
$tabMapping.Controls.Add($btnGuessMapping)

$lblMappingStatus = New-Object System.Windows.Forms.Label
$lblMappingStatus.AutoSize = $true
$lblMappingStatus.Location = New-Object System.Drawing.Point(15,680)
$lblMappingStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$lblMappingStatus.Text = "Załaduj plik CSV, aby rozpocząć mapowanie."
$tabMapping.Controls.Add($lblMappingStatus)

$gridMapping.add_EditingControlShowing({
    param($source,$eventData)
    if ($gridMapping.CurrentCell -and $gridMapping.CurrentCell.OwningColumn.Name -eq 'AdAttribute') {
        $tb = $eventData.Control -as [System.Windows.Forms.TextBox]
        if ($tb) {
            $tb.AutoCompleteMode = 'SuggestAppend'
            $tb.AutoCompleteSource = 'CustomSource'
            $tb.AutoCompleteCustomSource = $script:AttributeAutoComplete
        }
    }
})
$gridMapping.add_CellValueChanged({
    $script:AppState.VerificationStale = $true
})
$gridMapping.add_CurrentCellDirtyStateChanged({
    if ($gridMapping.IsCurrentCellDirty) {
        $gridMapping.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
    }
})
# Tab 3: Weryfikacja
$tabVerify = New-Object System.Windows.Forms.TabPage
$tabVerify.Text = "3. Weryfikacja kont"
[void]$tabControl.TabPages.Add($tabVerify)

$btnVerify = New-Object System.Windows.Forms.Button
$btnVerify.Text = "Sprawdź konta"
$btnVerify.Location = New-Object System.Drawing.Point(15,15)
$btnVerify.Width = 150
$tabVerify.Controls.Add($btnVerify)

$lvVerify = New-Object System.Windows.Forms.ListView
$lvVerify.Location = New-Object System.Drawing.Point(15,50)
$lvVerify.Size = New-Object System.Drawing.Size(1220,640)
$lvVerify.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$lvVerify.View = 'Details'
$lvVerify.FullRowSelect = $true
$lvVerify.GridLines = $true
$tabVerify.Controls.Add($lvVerify)
[void]$lvVerify.Columns.Add("samAccountName",200)
[void]$lvVerify.Columns.Add("Status",140)
[void]$lvVerify.Columns.Add("Informacja",850)

$lblVerifySummary = New-Object System.Windows.Forms.Label
$lblVerifySummary.AutoSize = $true
$lblVerifySummary.Location = New-Object System.Drawing.Point(15,700)
$lblVerifySummary.Text = "Nie uruchomiono weryfikacji."
$tabVerify.Controls.Add($lblVerifySummary)

# Tab 4: Aktualizacja
$tabApply = New-Object System.Windows.Forms.TabPage
$tabApply.Text = "4. Aktualizacja i podgląd"
[void]$tabControl.TabPages.Add($tabApply)

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Text = "Tryb testowy (WHATIF)"
$chkWhatIf.Location = New-Object System.Drawing.Point(15,15)
$chkWhatIf.AutoSize = $true
$chkWhatIf.Checked = $true
$tabApply.Controls.Add($chkWhatIf)

$chkAllowEmpty = New-Object System.Windows.Forms.CheckBox
$chkAllowEmpty.Text = "Pozwól czyścić atrybuty gdy w CSV jest pusto"
$chkAllowEmpty.Location = New-Object System.Drawing.Point(200,15)
$chkAllowEmpty.AutoSize = $true
$tabApply.Controls.Add($chkAllowEmpty)

$btnApply = New-Object System.Windows.Forms.Button
$btnApply.Text = "Wykonaj aktualizację"
$btnApply.Location = New-Object System.Drawing.Point(500,12)
$btnApply.Width = 180
$tabApply.Controls.Add($btnApply)

$gridChanges = New-Object System.Windows.Forms.DataGridView
$gridChanges.Location = New-Object System.Drawing.Point(15,50)
$gridChanges.Size = New-Object System.Drawing.Size(1220,640)
$gridChanges.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$gridChanges.AllowUserToAddRows = $false
$gridChanges.AllowUserToDeleteRows = $false
$gridChanges.AllowUserToResizeRows = $false
$gridChanges.ReadOnly = $true
$gridChanges.RowHeadersVisible = $false
$gridChanges.SelectionMode = 'FullRowSelect'
$gridChanges.MultiSelect = $false
$gridChanges.AutoSizeColumnsMode = 'DisplayedCells'
$gridChanges.AutoSizeRowsMode = 'DisplayedCells'
$gridChanges.ScrollBars = 'Both'
$gridChanges.ColumnHeadersHeightSizeMode = 'AutoSize'
$tabApply.Controls.Add($gridChanges)

function Add-FillColumn {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.DataGridView]$Grid,
        [Parameter(Mandatory)][string]$name,
        [Parameter(Mandatory)][string]$header,
        [int]$fillWeight = 100,
        [string]$AutoSizeMode = 'Fill'
    )
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $name
    $col.HeaderText = $header
    $col.AutoSizeMode = $AutoSizeMode
    $col.FillWeight = $fillWeight
    $col.MinimumWidth = 80
    $col.ReadOnly = $true
    [void]$Grid.Columns.Add($col)
}

Add-FillColumn -Grid $gridChanges -name "SamAccountName" -header "samAccountName" -fillWeight 120 -AutoSizeMode 'DisplayedCells'
Add-FillColumn -Grid $gridChanges -name "Attribute" -header "Atrybut" -fillWeight 110 -AutoSizeMode 'DisplayedCells'
Add-FillColumn -Grid $gridChanges -name "CurrentValue" -header "Aktualna wartość" -fillWeight 150 -AutoSizeMode 'DisplayedCells'
Add-FillColumn -Grid $gridChanges -name "NewValue" -header "Nowa wartość" -fillWeight 150 -AutoSizeMode 'DisplayedCells'
Add-FillColumn -Grid $gridChanges -name "Status" -header "Status" -fillWeight 90 -AutoSizeMode 'DisplayedCells'

$lblApplyStatus = New-Object System.Windows.Forms.Label
$lblApplyStatus.AutoSize = $true
$lblApplyStatus.Location = New-Object System.Drawing.Point(15,700)
$lblApplyStatus.Text = "Brak zmian do wyświetlenia."
$tabApply.Controls.Add($lblApplyStatus)

# Tab 5: Cofnięcie
$tabRollback = New-Object System.Windows.Forms.TabPage
$tabRollback.Text = "5. Cofnięcie zmian"
[void]$tabControl.TabPages.Add($tabRollback)

$btnRefreshRollback = New-Object System.Windows.Forms.Button
$btnRefreshRollback.Text = "Odśwież listę"
$btnRefreshRollback.Location = New-Object System.Drawing.Point(15,15)
$btnRefreshRollback.Width = 150
$tabRollback.Controls.Add($btnRefreshRollback)

$btnUndo = New-Object System.Windows.Forms.Button
$btnUndo.Text = "Cofnij wybrane zmiany"
$btnUndo.Location = New-Object System.Drawing.Point(180,15)
$btnUndo.Width = 190
$tabRollback.Controls.Add($btnUndo)

$lvRollback = New-Object System.Windows.Forms.ListView
$lvRollback.Location = New-Object System.Drawing.Point(15,50)
$lvRollback.Size = New-Object System.Drawing.Size(400,640)
$lvRollback.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$lvRollback.View = 'Details'
$lvRollback.FullRowSelect = $true
$lvRollback.GridLines = $true
$tabRollback.Controls.Add($lvRollback)
[void]$lvRollback.Columns.Add("Data",150)
[void]$lvRollback.Columns.Add("Plik",120)
[void]$lvRollback.Columns.Add("Opis",120)

$gridRollbackPreview = New-Object System.Windows.Forms.DataGridView
$gridRollbackPreview.Location = New-Object System.Drawing.Point(430,50)
$gridRollbackPreview.Size = New-Object System.Drawing.Size(805,640)
$gridRollbackPreview.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$gridRollbackPreview.AllowUserToAddRows = $false
$gridRollbackPreview.AllowUserToDeleteRows = $false
$gridRollbackPreview.AllowUserToResizeRows = $false
$gridRollbackPreview.ReadOnly = $true
$gridRollbackPreview.RowHeadersVisible = $false
$gridRollbackPreview.SelectionMode = 'FullRowSelect'
$gridRollbackPreview.MultiSelect = $false
$gridRollbackPreview.AutoSizeColumnsMode = 'DisplayedCells'
$gridRollbackPreview.AutoSizeRowsMode = 'DisplayedCells'
$gridRollbackPreview.ScrollBars = 'Both'
$gridRollbackPreview.ColumnHeadersHeightSizeMode = 'AutoSize'
$tabRollback.Controls.Add($gridRollbackPreview)

Add-FillColumn -Grid $gridRollbackPreview -name "Sam" -header "samAccountName" -fillWeight 120 -AutoSizeMode 'DisplayedCells'
Add-FillColumn -Grid $gridRollbackPreview -name "Attr" -header "Atrybut" -fillWeight 110 -AutoSizeMode 'DisplayedCells'
Add-FillColumn -Grid $gridRollbackPreview -name "Before" -header "Stara wartość" -fillWeight 150 -AutoSizeMode 'DisplayedCells'
Add-FillColumn -Grid $gridRollbackPreview -name "After" -header "Nowa wartość" -fillWeight 150 -AutoSizeMode 'DisplayedCells'
Add-FillColumn -Grid $gridRollbackPreview -name "Status" -header "Status" -fillWeight 90 -AutoSizeMode 'DisplayedCells'

$lblRollbackStatus = New-Object System.Windows.Forms.Label
$lblRollbackStatus.AutoSize = $true
$lblRollbackStatus.Location = New-Object System.Drawing.Point(430,700)
$lblRollbackStatus.Text = "Wybierz plik cofania, aby zobaczyć szczegóły."
$tabRollback.Controls.Add($lblRollbackStatus)
# --- Zdarzenia ---
$btnBrowse.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "CSV (*.csv)|*.csv|Wszystkie pliki (*.*)|*.*"
    if (-not [string]::IsNullOrWhiteSpace($txtCsvPath.Text)) {
        try {
            $dialog.InitialDirectory = Split-Path -Path $txtCsvPath.Text -Parent
        } catch { }
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtCsvPath.Text = $dialog.FileName
    }
})

$btnLoad.Add_Click({
    try {
        $selected = $comboDelimiter.SelectedItem
        $delimiter = if ($selected.Value) { $selected.Value } else { 'auto' }
        Import-CsvData -Path $txtCsvPath.Text -DelimiterMode $delimiter
        Show-CsvPreview
        Update-MappingGrid -Grid $gridMapping
        Set-MappingDefaults -Grid $gridMapping
        $lblImportStatus.Text = "Załadowano $($script:AppState.Data.Count) wierszy. Separator: '$($script:AppState.LastDelimiter)'."
        $lblMappingStatus.Text = "Skonfiguruj mapowanie, a następnie przejdź do weryfikacji."
        $script:AppState.VerificationStale = $true
        & $runVerification
    } catch {
        Show-DialogMessage "Nie udało się zaimportować CSV.`n$($_.Exception.Message)" "Import danych" ([System.Windows.Forms.MessageBoxIcon]::Error)
        $lblImportStatus.Text = "Błąd importu. Sprawdź logi."
        Clear-CsvPreview
        Write-AppLog "Błąd importu: $($_.Exception.Message)"
    }
})

$btnAddMapping.Add_Click({
    if (-not $script:AppState.Columns -or $script:AppState.Columns.Count -eq 0) {
        Show-DialogMessage "Najpierw zaimportuj plik CSV." "Mapowanie" ([System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Add-MappingRow -Grid $gridMapping | Out-Null
    $script:AppState.VerificationStale = $true
})

$btnRemoveMapping.Add_Click({
    foreach ($row in $gridMapping.SelectedRows) {
        if ($row.Tag -eq 'key') { continue }
        $gridMapping.Rows.Remove($row)
    }
    $script:AppState.VerificationStale = $true
})

$btnGuessMapping.Add_Click({
    if (-not $script:AppState.Columns -or $script:AppState.Columns.Count -eq 0) {
        Show-DialogMessage "Najpierw zaimportuj plik CSV." "Mapowanie" ([System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Set-MappingDefaults -Grid $gridMapping
    $lblMappingStatus.Text = "Zastosowano automatyczne dopasowanie."
    $script:AppState.VerificationStale = $true
})

$runVerification = {
    if (-not $script:AppState.Data -or $script:AppState.Data.Count -eq 0) {
        return
    }
    if (-not $script:AppState.VerificationStale -and $script:AppState.Verified.Count -gt 0) {
        return
    }
    $mappings = Get-MappingRows -Grid $gridMapping
    $form.Cursor = 'WaitCursor'
    try {
        if (Invoke-Verification -Mappings $mappings -ListView $lvVerify -SummaryLabel $lblVerifySummary) {
            $script:AppState.Mapping = $mappings
            Build-ChangePreview -Mappings $mappings
            Update-ChangesGrid -Grid $gridChanges -SummaryLabel $lblApplyStatus
            $script:AppState.VerificationStale = $false
        }
    } finally {
        $form.Cursor = 'Default'
    }
}

$btnVerify.Add_Click({
    if (-not $script:AppState.Data -or $script:AppState.Data.Count -eq 0) {
        Show-DialogMessage "Załaduj dane CSV zanim rozpoczniesz weryfikację." "Weryfikacja" ([System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $script:AppState.VerificationStale = $true
    & $runVerification
})

$tabControl.add_SelectedIndexChanged({
    if ($tabControl.SelectedTab -eq $tabVerify -or $tabControl.SelectedTab -eq $tabApply) {
        $script:AppState.VerificationStale = $true
        & $runVerification
    }
})

$chkAllowEmpty.Add_CheckedChanged({
    $script:AppState.AllowEmptyClear = $chkAllowEmpty.Checked
    if ($script:AppState.Mapping.Count -gt 0 -and $script:AppState.Verified.Count -gt 0) {
        Build-ChangePreview -Mappings $script:AppState.Mapping
        Update-ChangesGrid -Grid $gridChanges -SummaryLabel $lblApplyStatus
    }
})

$btnApply.Add_Click({
    $whatIf = $chkWhatIf.Checked
    if (-not $whatIf) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Zostaną zapisane zmiany w AD. Czy na pewno chcesz kontynuować?",
            "Potwierdź operację",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }
    $form.Cursor = 'WaitCursor'
    try {
        Invoke-ApplyChanges -WhatIf:$whatIf -Grid $gridChanges -StatusLabel $lblApplyStatus -RollbackList $lvRollback
    } finally {
        $form.Cursor = 'Default'
    }
})

$btnRefreshRollback.Add_Click({
    Update-RollbackList -ListView $lvRollback
    $lblRollbackStatus.Text = "Lista plików cofania została odświeżona."
})
$lvRollback.Add_SelectedIndexChanged({
    if ($lvRollback.SelectedItems.Count -eq 0) { return }
    $path = $lvRollback.SelectedItems[0].Tag
    if ($path) {
        Update-RollbackPreview -Path $path -Grid $gridRollbackPreview -StatusLabel $lblRollbackStatus
    }
})

$btnUndo.Add_Click({
    if (-not $script:AppState.SelectedRollback) {
        Show-DialogMessage "Wybierz plik cofania z listy." "Cofnięcie" ([System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Czy na pewno chcesz przywrócić wartości z wybranego pliku?",
        "Potwierdź cofnięcie",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    $form.Cursor = 'WaitCursor'
    try {
        Invoke-Rollback -Snapshot $script:AppState.SelectedRollback -Grid $gridRollbackPreview -StatusLabel $lblRollbackStatus
        Update-RollbackList -ListView $lvRollback
    } finally {
        $form.Cursor = 'Default'
    }
})

$form.Add_Shown({
    Update-RollbackList -ListView $lvRollback
})

[void]$form.ShowDialog()
