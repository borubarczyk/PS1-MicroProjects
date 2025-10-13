<# 
    AD_Swap_GUI.ps1
    - Wklej loginy (SamAccountName) w lewym polu
    - Kliknij "Sprawdź w AD" -> w tabeli zobaczysz Imię/Nazwisko
    - Zaznacz "Zamień?" dla tych, gdzie trzeba zamienić
    - Kliknij "Zamień imię <-> nazwisko (zaznaczone)"
    - "Tryb testowy" => nie zapisuje zmian w AD
    Wymaga: Import-Module ActiveDirectory
#>

# --- WSTĘPNE SPRAWDZENIA ---
try { Import-Module ActiveDirectory -ErrorAction Stop }
catch {
    [System.Windows.Forms.MessageBox]::Show("Nie udało się załadować modułu ActiveDirectory.`nZainstaluj RSAT/AD Module.","Błąd", 'OK', 'Error') | Out-Null
    return
}

if ($host.Runspace.ApartmentState -ne 'STA') {
    Write-Warning "PowerShell nie działa w STA. Uruchom ten skrypt poleceniem: powershell -STA -File .\AD_Swap_GUI.ps1"
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- UTYLITKI ---
function New-RowObject {
    param(
        [bool]$Swap=$false,[string]$Login,[string]$Istnieje,[string]$Imie,[string]$Nazwisko,[string]$DisplayName,[string]$DN
    )
    $obj = New-Object PSObject -Property ([ordered]@{
        'Zamień?'    = $Swap
        'Login'      = $Login
        'Istnieje'   = $Istnieje
        'Imię'       = $Imie
        'Nazwisko'   = $Nazwisko
        'Wyświetlana'= $DisplayName
        'DN'         = $DN
    })
    return $obj
}

function Get-UsersFromAD {
    param([string[]]$Loginy)

    $items = New-Object System.Collections.ArrayList
    foreach ($l in $Loginy) {
        $login = $l.Trim()
        if ([string]::IsNullOrWhiteSpace($login)) { continue }

        $user = $null
        try {
            $user = Get-ADUser -Identity $login -Properties GivenName,Surname,DisplayName,DistinguishedName -ErrorAction Stop
            # Domyślnie NIE zaznaczamy, ale spróbujmy auto-sugestii:
            # Jeśli GivenName i Surname wyglądają na odwrócone względem DisplayName "Nazwisko Imię"
            $autoSwap = $false
            if ($user.GivenName -and $user.Surname -and $user.DisplayName) {
                $disp = $user.DisplayName -replace '\s+',' ' 
                $gn = $user.GivenName.Trim()
                $sn = $user.Surname.Trim()
                # Prosta heurystyka: DisplayName == "Nazwisko Imię" albo "Imię Nazwisko"
                $isSN_GN = ($disp -eq "$sn $gn")
                $isGN_SN = ($disp -eq "$gn $sn")
                # Zasugeruj zamianę, jeżeli wygląda na "Nazwisko Imię" a GivenName==Nazwisko i Surname==Imię
                if (($gn -match '^[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż-]+$') -and ($sn -match '^[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż-]+$')) {
                    if ($isSN_GN -and ($gn -cmatch $sn) -and ($sn -cmatch $gn)) { $autoSwap = $true }
                }
            }

            $null = $items.Add( (New-RowObject -Swap:$autoSwap -Login:$login -Istnieje:'TAK' -Imie:$user.GivenName -Nazwisko:$user.Surname -DisplayName:$user.DisplayName -DN:$user.DistinguishedName) )
        }
        catch {
            $null = $items.Add( (New-RowObject -Swap:$false -Login:$login -Istnieje:'NIE' -Imie:'' -Nazwisko:'' -DisplayName:'' -DN:'') )
        }
    }
    return $items
}

function Export-GridToCsv {
    param($data,[string]$Path)
    $sb = New-Object System.Text.StringBuilder
    $null = $sb.AppendLine("Zamień?;Login;Istnieje;Imię;Nazwisko;Wyświetlana;DN")
    foreach ($row in $data) {
        $line = '{0};{1};{2};{3};{4};{5};{6}' -f ($row.'Zamień?' -as [bool]), $row.Login, $row.Istnieje, $row.Imię, $row.Nazwisko, $row.'Wyświetlana', $row.DN
        $null = $sb.AppendLine($line)
    }
    [IO.File]::WriteAllText($Path, $sb.ToString(), [Text.UTF8Encoding]::new($false))
}

# --- FORMULARZ ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD: sprawdzanie loginów i zamiana Imię <-> Nazwisko"
$form.Size = New-Object System.Drawing.Size(1000,650)
$form.StartPosition = 'CenterScreen'

# Panele
$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Dock = 'Fill'
$form.Controls.Add($rightPanel)

# Uklad w prawej czesci: 3 wiersze (gora: panel z przyciskami; srodek: siatka; dol: loginy)
$rightLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rightLayout.Dock = 'Fill'
$rightLayout.ColumnCount = 1
$rightLayout.RowCount = 3
$null = $rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40)))
$null = $rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 180)))
$rightLayout.Padding = [System.Windows.Forms.Padding]::new(0)
$rightLayout.Margin  = [System.Windows.Forms.Padding]::new(0)
$rightPanel.Controls.Add($rightLayout)

# Panel dolny zawierajacy pole do wklejenia loginow i przyciski akcji
$bottomPanel = New-Object System.Windows.Forms.Panel
$bottomPanel.Dock = 'Fill'
$bottomPanel.Height = 180

# Pole do wklejenia loginów
$lblInput = New-Object System.Windows.Forms.Label
$lblInput.Text = "Wklej loginy (SamAccountName), po jednym w linii:"
$lblInput.AutoSize = $true
$lblInput.Location = '10,10'
$bottomPanel.Controls.Add($lblInput)

$txtLoginy = New-Object System.Windows.Forms.TextBox
$txtLoginy.Multiline = $true
$txtLoginy.ScrollBars = 'Vertical'
$txtLoginy.Location = '10,35'
$txtLoginy.Size = New-Object System.Drawing.Size(280,100)
$txtLoginy.Anchor = 'Top,Left,Bottom,Right'
$bottomPanel.Controls.Add($txtLoginy)

$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Text = "Sprawdź w AD"
$btnCheck.Location = '10,145'
$btnCheck.Width = 130
$btnCheck.Anchor = 'Bottom,Left'
$bottomPanel.Controls.Add($btnCheck)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = "Wyczyść"
$btnClear.Location = '160,145'
$btnClear.Width = 130
$btnClear.Anchor = 'Bottom,Left'
$bottomPanel.Controls.Add($btnClear)

# Prawa strona: kontrolki nad gridem
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = 'Fill'
$topPanel.Height = 40

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text = "Tryb testowy (bez zmian w AD)"
$chkDryRun.AutoSize = $true
$chkDryRun.Checked = $true
$chkDryRun.Location = '10,10'
$topPanel.Controls.Add($chkDryRun)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Zaznacz wszystko"
$btnSelectAll.Location = '280,7'
$btnSelectAll.Width = 130
$btnSelectAll.Anchor = 'Top,Right'
$topPanel.Controls.Add($btnSelectAll)

$btnUnselectAll = New-Object System.Windows.Forms.Button
$btnUnselectAll.Text = "Odznacz wszystko"
$btnUnselectAll.Location = '420,7'
$btnUnselectAll.Width = 130
$btnUnselectAll.Anchor = 'Top,Right'
$topPanel.Controls.Add($btnUnselectAll)

$btnSwap = New-Object System.Windows.Forms.Button
$btnSwap.Text = "Zamień imię <-> nazwisko (zaznaczone)"
$btnSwap.Location = '560,7'
$btnSwap.Width = 220
$btnSwap.Anchor = 'Top,Right'
$topPanel.Controls.Add($btnSwap)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Eksportuj CSV"
$btnExport.Location = '790,7'
$btnExport.Width = 120
$btnExport.Anchor = 'Top,Right'
$topPanel.Controls.Add($btnExport)

# Grid
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AutoGenerateColumns = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $true
$grid.ReadOnly = $false
$grid.BackgroundColor = [System.Drawing.SystemColors]::Window
$grid.AutoSizeColumnsMode = 'DisplayedCells'
$grid.BorderStyle = 'FixedSingle'

# Umiesc kontener i siatke w ukladzie tabeli zamiast polegac na kolejnosci Dock
$rightLayout.Controls.Add($topPanel, 0, 0)
$rightLayout.Controls.Add($grid, 0, 1)
$rightLayout.Controls.Add($bottomPanel, 0, 2)

# Responsywny uklad przyciskow w topPanel
$reposition = {
    param($sender,$e)
    $pad = 10
    $w = $topPanel.ClientSize.Width
    if ($w -le 0) { return }
    $x = $w - $pad

    $btnExport.Location = New-Object System.Drawing.Point(($x - $btnExport.Width), 7)
    $x -= ($btnExport.Width + $pad)
    $btnSwap.Location   = New-Object System.Drawing.Point(($x - $btnSwap.Width), 7)
    $x -= ($btnSwap.Width + $pad)
    $btnUnselectAll.Location = New-Object System.Drawing.Point(($x - $btnUnselectAll.Width), 7)
    $x -= ($btnUnselectAll.Width + $pad)
    $leftStart = [Math]::Max($chkDryRun.Right + $pad, $x - $btnSelectAll.Width)
    $btnSelectAll.Location = New-Object System.Drawing.Point($leftStart, 7)
}

$topPanel.Add_Resize($reposition)
$null = $reposition.Invoke($topPanel, [System.EventArgs]::Empty)

# Kolumny
$colSwap = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSwap.HeaderText = "Zamień?"
$colSwap.DataPropertyName = 'Zamień?'
$colSwap.Width = 70
$grid.Columns.Add($colSwap)

foreach ($colName in @('Login','Istnieje','Imię','Nazwisko','Wyświetlana','DN')) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $colName
    $col.DataPropertyName = $colName
    if ($colName -eq 'DN') { $col.Width = 260 } elseif ($colName -eq 'Wyświetlana') { $col.Width = 160 } else { $col.Width = 110 }
    $col.ReadOnly = $true
    $grid.Columns.Add($col)
}

  # Ustawienia widocznosci/scrollingu, aby kolumny od lewej byly zawsze widoczne
  $grid.RowHeadersVisible = $false
  $grid.ScrollBars = 'Both'
  try { $grid.FirstDisplayedScrollingColumnIndex = 0 } catch {}

  # Ustal nazwy kolumn (Name) na podstawie naglowkow, aby dzialalo indeksowanie po nazwie
foreach ($c in $grid.Columns) {
    if ([string]::IsNullOrEmpty($c.Name)) { $c.Name = $c.HeaderText }
}

# Uporzadkuj kolejnosc wyswietlania kolumn (lewa -> prawa)
try {
    $grid.Columns['Zamien?'].DisplayIndex = 0
    $grid.Columns['Login'].DisplayIndex   = 1
    $grid.Columns['Istnieje'].DisplayIndex= 2
    $grid.Columns['Imie'].DisplayIndex    = 3
    $grid.Columns['Nazwisko'].DisplayIndex= 4
    $grid.Columns['Wyswietlana'].DisplayIndex = 5
    $grid.Columns['DN'].DisplayIndex      = 6
    # Przypnij pierwsze kolumny, aby zawsze byly widoczne po lewej
    $grid.Columns['Zamien?'].Frozen = $true
    $grid.Columns['Login'].Frozen   = $true
    $grid.Columns['Istnieje'].Frozen= $true
} catch {}

# Status
$status = New-Object System.Windows.Forms.StatusStrip
$form.Controls.Add($status)
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text = "Gotowy."
$status.Items.Add($lblStatus) | Out-Null

# Źródło danych
$table = New-Object System.Collections.ArrayList
$binding = New-Object System.Windows.Forms.BindingSource
$binding.DataSource = $table
$grid.DataSource = $binding

# Po odswiezeniu danych ustaw widok na pierwsza kolumne
$grid.add_DataBindingComplete({
    try { $grid.FirstDisplayedScrollingColumnIndex = 0 } catch {}
})

# --- ZDARZENIA ---
$btnClear.Add_Click({
    $txtLoginy.Clear()
    $table.Clear()
    $binding.ResetBindings($true)
    $lblStatus.Text = "Wyczyszczono."
})

$btnCheck.Add_Click({
    $lblStatus.Text = "Sprawdzanie w AD..."
    $form.UseWaitCursor = $true
    try {
        $loginy = ($txtLoginy.Text -split "`r?`n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        $table.Clear()
        $results = Get-UsersFromAD -Loginy $loginy
        foreach ($r in $results) { [void]$table.Add($r) }
        $binding.ResetBindings($true)
        $lblStatus.Text = "Zakończono sprawdzanie. Znaleziono: $($results.Count) pozycji."
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Błąd podczas sprawdzania: $($_.Exception.Message)","Błąd",'OK','Error') | Out-Null
        $lblStatus.Text = "Błąd sprawdzania."
    }
    finally { $form.UseWaitCursor = $false }
})

$btnSelectAll.Add_Click({
    foreach ($row in $grid.Rows) { $row.Cells[0].Value = $true }
    $lblStatus.Text = "Zaznaczono wszystko."
})

$btnUnselectAll.Add_Click({
    foreach ($row in $grid.Rows) { $row.Cells[0].Value = $false }
    $lblStatus.Text = "Odznaczono wszystko."
})

$btnSwap.Add_Click({
    $doit = -not $chkDryRun.Checked
    $count = 0
    $errors = 0
    $form.UseWaitCursor = $true
    try {
        foreach ($row in $grid.Rows) {
            $wantSwap = [bool]$row.Cells[0].Value
            $exists = $row.Cells[$grid.Columns['Istnieje'].Index].Value -eq 'TAK'
            if ($wantSwap -and $exists) {
                $login  = $row.Cells[$grid.Columns['Login'].Index].Value
                $given  = $row.Cells[$grid.Columns['Imię'].Index].Value
                $sn     = $row.Cells[$grid.Columns['Nazwisko'].Index].Value
                try {
                    if ($doit) {
                        Set-ADUser -Identity $login -GivenName $sn -Surname $given -ErrorAction Stop
                    }
                    # odśwież wiersz (lokalnie zamieniamy widok niezależnie od dry-run)
                    $row.Cells[$grid.Columns['Imię'].Index].Value = $sn
                    $row.Cells[$grid.Columns['Nazwisko'].Index].Value = $given
                    $row.Cells[0].Value = $false
                    $count++
                }
                catch {
                    $errors++
                }
            }
        }
        if ($doit) {
            $lblStatus.Text = "Zamieniono $count rekord(ów). Błędów: $errors."
        } else {
            $lblStatus.Text = "TRYB TESTOWY: zasymulowano zamianę dla $count rekord(ów). Błędów: $errors."
        }
    }
    finally { $form.UseWaitCursor = $false }
})

$btnExport.Add_Click({
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Title = "Zapisz wynik do CSV"
    $sfd.Filter = "CSV (*.csv)|*.csv"
    $sfd.FileName = "wynik_AD.csv"
    if ($sfd.ShowDialog() -eq 'OK') {
        try {
            Export-GridToCsv -data $table -Path $sfd.FileName
            $lblStatus.Text = "Zapisano: $($sfd.FileName)"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Nie udało się zapisać: $($_.Exception.Message)","Błąd",'OK','Error') | Out-Null
        }
    }
})

# --- START ---
$form.Add_Shown({ $txtLoginy.Focus(); try { $grid.FirstDisplayedScrollingColumnIndex = 0 } catch {} })
[void]$form.ShowDialog()
