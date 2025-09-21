#region Init
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

# Hidden JSON config in start directory (priority over inline defaults)
$tabNames = @('Uczen','Student','Pracownik','Wykladowca','Inne')
$ConfigFileName = '.kreator-kont-ad.config.json'
$StartDir = try { (Get-Location).Path } catch { $PSScriptRoot }
if (-not $StartDir) { $StartDir = Split-Path -Parent $PSCommandPath }
$ConfigPath = Join-Path -Path $StartDir -ChildPath $ConfigFileName

$Domain_Defaults_Default = @{
  Uczen      = 'uczen.example.com'
  Student    = 'student.example.com'
  Pracownik  = 'firma.local'
  Wykladowca = 'wykladowcy.firma.local'
  Inne       = 'inne.firma.local'
}
$OU_Defaults_Default = @{
  Uczen      = 'OU=2025/2026,OU=Student,OU=UsersO365,DC=ad,DC=example,DC=com'
  Student    = 'OU=2024/2025,OU=Student,OU=UsersO365,DC=ad,DC=example,DC=com'
  Pracownik  = 'OU=Pracownicy,DC=firma,DC=local'
  Wykladowca = 'OU=Wykladowcy,DC=firma,DC=local'
  Inne       = 'OU=Inne,DC=firma,DC=local'
}
$Cities_Default = @('Warszawa','Olsztyn','Katowice','Poznań','Szczecin','Człuchów','Lublin')

$Domain_Defaults = $Domain_Defaults_Default.Clone()
$OU_Defaults      = $OU_Defaults_Default.Clone()
$Cities           = @($Cities_Default)
$ConfigStatusMessage = $null

function Test-ConfigObject([object]$cfg){
  if (-not $cfg) { return $false }
  foreach ($k in 'Domain_Defaults','OU_Defaults') { if (-not ($cfg.PSObject.Properties.Name -contains $k)) { return $false } }
  foreach ($t in $tabNames) { if (-not ($cfg.Domain_Defaults.PSObject.Properties.Name -contains $t)) { return $false } }
  foreach ($t in $tabNames) { if (-not ($cfg.OU_Defaults.PSObject.Properties.Name -contains $t)) { return $false } }
  return $true
}

if (Test-Path -LiteralPath $ConfigPath) {
  try {
    $cfg = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json -ErrorAction Stop
    if (Test-ConfigObject $cfg) {
      $Domain_Defaults = @{}
      foreach ($k in $tabNames) { $Domain_Defaults[$k] = [string]$cfg.Domain_Defaults.$k }
      $OU_Defaults = @{}
      foreach ($k in $tabNames) { $OU_Defaults[$k] = [string]$cfg.OU_Defaults.$k }
      if ($cfg.PSObject.Properties.Name -contains 'Cities' -and $cfg.Cities) { $Cities = @($cfg.Cities) }
      $ConfigStatusMessage = "Wczytano konfigurację: $ConfigPath"
    } else { $ConfigStatusMessage = "Konfiguracja niepoprawna – używam domyślnej. $ConfigPath" }
  } catch { $ConfigStatusMessage = "Błąd odczytu konfiguracji – używam domyślnej. $_" }
}
else {
  try {
    $cfgObj = [pscustomobject]@{ Domain_Defaults=$Domain_Defaults_Default; OU_Defaults=$OU_Defaults_Default; Cities=$Cities_Default }
    $json = $cfgObj | ConvertTo-Json -Depth 6
    $null = New-Item -Path $ConfigPath -ItemType File -Force
    Set-Content -Path $ConfigPath -Value $json -Encoding UTF8
    (Get-Item -LiteralPath $ConfigPath).Attributes = (Get-Item -LiteralPath $ConfigPath).Attributes -bor [System.IO.FileAttributes]::Hidden
    $ConfigStatusMessage = "Utworzono domyślny plik konfiguracyjny: $ConfigPath"
  } catch { $ConfigStatusMessage = "Nie udało się utworzyć pliku konfiguracyjnego: $_" }
}

# utils
function Remove-PolishDiacritics {
  param([string]$text)
  if (-not $text) { return $null }
  $map = @{ 'ą'='a'; 'ć'='c'; 'ę'='e'; 'ł'='l'; 'ń'='n'; 'ó'='o'; 'ś'='s'; 'ż'='z'; 'ź'='z' }
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $text.ToLower().ToCharArray()) { if ($map.ContainsKey($ch)) { [void]$sb.Append($map[$ch]) } else { [void]$sb.Append($ch) } }
  return ($sb.ToString() -replace '[^a-z0-9\.]','')
}

function Get-RandChar([Parameter(Mandatory)][string]$Pool){
  $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
  $bytes = New-Object 'Byte[]' 4; $rng.GetBytes($bytes)
  $idx = [math]::Abs([BitConverter]::ToInt32($bytes,0)) % $Pool.Length
  $Pool[$idx]
}
function Test-Sequential([char]$Prev,[char]$Curr){ if (-not $Prev){return $false}; return ([math]::Abs([int][char]$Prev - [int][char]$Curr) -eq 1) }
function New-RandomPassword([ValidateRange(8,128)][int]$Length=12){
  $chars=@{Digits='123456789';Lower='abcdefghjkmnpqrstuvwxyz';Upper='ABCDEFGHJKMNPQRSTUVWXYZ';Symbols='#$%&?@'}
  $sb=New-Object System.Text.StringBuilder
  [void]$sb.Append((Get-RandChar -Pool ($chars.Lower+$chars.Upper)))
  foreach($k in 'Digits','Lower','Upper','Symbols'){ if($sb.Length -ge $Length){break}; $ch=Get-RandChar -Pool $chars[$k]; if(Test-Sequential $sb[$sb.Length-1] $ch){$ch=Get-RandChar -Pool $chars[$k]}; [void]$sb.Append($ch) }
  $all=($chars.Digits+$chars.Lower+$chars.Upper+$chars.Symbols)
  while($sb.Length -lt $Length){ $ch=Get-RandChar -Pool $all; $p=if($sb.Length -gt 0){$sb[$sb.Length-1]}else{[char]0}; $p2=if($sb.Length -gt 1){$sb[$sb.Length-2]}else{[char]0}; if(Test-Sequential $p $ch){continue}; if($p -eq $ch -and $p2 -eq $ch){continue}; [void]$sb.Append($ch) }
  $pass=$sb.ToString(); $ok=($pass -cmatch '[0-9]') -and ($pass -cmatch '[a-z]') -and ($pass -cmatch '[A-Z]') -and ($pass -cmatch '[#\$%&\?@]')
  if($ok){$pass}else{ New-RandomPassword -Length $Length }
}

function Test-ADModule {
  try {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { throw 'Brak modułu ActiveDirectory (RSAT).' }
    if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory -ErrorAction Stop | Out-Null }
    return $true
  } catch {
    [System.Windows.Forms.MessageBox]::Show("Brak modułu ActiveDirectory. Zainstaluj RSAT: Active Directory i spróbuj ponownie.`r`n$_", 'Błąd modułu AD',[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    return $false
  }
}

function Get-AvailableDomains {
  if (-not (Test-ADModule)) { return @() }
  try {
    $forest = Get-ADForest -ErrorAction Stop
    @($forest.RootDomain,$forest.Domains,$forest.UPNSuffixes) | ForEach-Object { $_ } | Where-Object { $_ } | Select-Object -Unique | Sort-Object
  } catch { @() }
}

function Show-OUChooser { param([string]$Title='Wybierz OU')
  if (-not (Test-ADModule)) { return $null }
  $form = New-Object Windows.Forms.Form
  $form.Text=$Title; $form.StartPosition='CenterParent'; $form.Size=[Drawing.Size]::new(700,750)
  $form.Font = New-Object Drawing.Font('Segoe UI',10)
  $tree = New-Object Windows.Forms.TreeView; $tree.Dock='Fill'; $tree.HideSelection=$false
  $ok=New-Object Windows.Forms.Button; $ok.Text='OK'; $ok.Width=120; $ok.Height=32
  $cancel=New-Object Windows.Forms.Button; $cancel.Text='Anuluj'; $cancel.Width=120; $cancel.Height=32
  $panel = New-Object Windows.Forms.FlowLayoutPanel; $panel.Dock='Top'; $panel.FlowDirection='RightToLeft'; $panel.Controls.Add($ok)|Out-Null; $panel.Controls.Add($cancel)|Out-Null
  $root = New-Object Windows.Forms.TableLayoutPanel; $root.Dock='Fill'; $root.RowCount=2
  [void]$root.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent,100)))
  [void]$root.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize)))
  $root.Controls.Add($tree,0,0); $root.Controls.Add($panel,0,1); $form.Controls.Add($root)
  try {
    $nodes=@{}
    $ous = Get-ADOrganizationalUnit -Filter * -SearchScope Subtree -ErrorAction Stop | Sort-Object DistinguishedName
    foreach($ou in $ous){ $dn=[string]$ou.DistinguishedName; $nodes[$dn]=New-Object Windows.Forms.TreeNode -Property @{Text=$ou.Name; Tag=$dn} }
    foreach($ou in $ous){ $dn=[string]$ou.DistinguishedName; $parent= if($dn.Contains(',')){$dn.Substring($dn.IndexOf(',')+1)}else{$null}; if($parent -and $nodes.ContainsKey($parent)){[void]$nodes[$parent].Nodes.Add($nodes[$dn])} }
    foreach($kv in $nodes.GetEnumerator()){ $dn=$kv.Key; $parent= if($dn.Contains(',')){$dn.Substring($dn.IndexOf(',')+1)}else{$null}; if(-not $nodes.ContainsKey($parent)){[void]$tree.Nodes.Add($kv.Value)} }
    foreach($n in $tree.Nodes){ $n.Expand() }
  } catch {
    [Windows.Forms.MessageBox]::Show("Nie udało się pobrać OU: $($_.Exception.Message)",'Błąd') | Out-Null; return $null
  }
  $selected=$null
  $ok.Add_Click({ if ($tree.SelectedNode -and $tree.SelectedNode.Tag){ $selected=[string]$tree.SelectedNode.Tag; $form.Close() } })
  $cancel.Add_Click({ $selected=$null; $form.Close() })
  [void]$form.ShowDialog(); return $selected
}
#endregion

#region GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Kreator Kont AD'
$form.Size = New-Object System.Drawing.Size(1000, 820)
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false; $form.MinimizeBox = $false

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Size = New-Object System.Drawing.Size(960, 540)
$tabs.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($tabs)

$tb_logg_box = New-Object System.Windows.Forms.RichTextBox
$tb_logg_box.Size = New-Object System.Drawing.Size(960, 100)
$tb_logg_box.Location = New-Object System.Drawing.Point(10, 700)
$tb_logg_box.ReadOnly = $true
$form.Controls.Add($tb_logg_box)

function Write-ToTextBox {
  param([string]$Text, [ValidateSet('Info','Warning','Error')][string]$Type='Info')
  $Date = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  switch ($Type){ 'Info'{$color=[Drawing.Color]::Green;$p='[INFO]'}'Warning'{$color=[Drawing.Color]::Orange;$p='[WARNING]'}'Error'{$color=[Drawing.Color]::Red;$p='[ERROR]'} }
  $tb_logg_box.SelectionStart=$tb_logg_box.TextLength; $tb_logg_box.SelectionLength=0; $tb_logg_box.SelectionColor=$color
  $tb_logg_box.AppendText("$Date $p $Text`r`n"); $tb_logg_box.ScrollToCaret(); $tb_logg_box.SelectionColor=$tb_logg_box.ForeColor
}

# Usage label
$lblUsage = New-Object System.Windows.Forms.Label
$lblUsage.Text = 'Wklej: Imię[TAB]Nazwisko (Student: +NrAlbumu). Wybierz domenę i OU dla aktywnej zakładki. Następnie Dodaj konta.'
$lblUsage.AutoSize = $true
$lblUsage.Location = New-Object System.Drawing.Point(10, 555)
$form.Controls.Add($lblUsage)

# Miasto
$lblMiasto = New-Object System.Windows.Forms.Label
$lblMiasto.Text = 'Domyślne miasto dla wszystkich:'
$lblMiasto.AutoSize = $true
$lblMiasto.Location = New-Object System.Drawing.Point(10, 580)
$form.Controls.Add($lblMiasto)

$cbMiasto = New-Object System.Windows.Forms.ComboBox
$cbMiasto.Location = New-Object System.Drawing.Point(220, 578)
$cbMiasto.Width = 220
$cbMiasto.DropDownStyle = 'DropDownList'
$form.Controls.Add($cbMiasto)

# Domena
$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Text = 'Domena/UPN dla zakładki:'
$lblDomain.AutoSize = $true
$lblDomain.Location = New-Object System.Drawing.Point(10, 610)
$form.Controls.Add($lblDomain)

$tbDomain_Edit = New-Object System.Windows.Forms.TextBox
$tbDomain_Edit.Location = New-Object System.Drawing.Point(220, 608)
$tbDomain_Edit.Width = 320
$tbDomain_Edit.ReadOnly = $true
$tbDomain_Edit.Text = '(używana domena dla aktywnej zakładki)'
$form.Controls.Add($tbDomain_Edit)

$cbDomain = New-Object System.Windows.Forms.ComboBox
$cbDomain.Location = New-Object System.Drawing.Point(550, 608)
$cbDomain.Width = 200; $cbDomain.DropDownStyle='DropDownList'
$cbDomain.Add_SelectedIndexChanged({ if ($tabs.SelectedTab -and $cbDomain.SelectedItem){ $sel=[string]$cbDomain.SelectedItem; $Domain_Defaults[$tabs.SelectedTab.Text]=$sel; $tbDomain_Edit.Text=$sel } })
$form.Controls.Add($cbDomain)

# OU
$lblOU = New-Object System.Windows.Forms.Label
$lblOU.Text = 'OU dla zakładki:'
$lblOU.AutoSize = $true
$lblOU.Location = New-Object System.Drawing.Point(10, 640)
$form.Controls.Add($lblOU)

$tbOU_Edit = New-Object System.Windows.Forms.TextBox
$tbOU_Edit.Location = New-Object System.Drawing.Point(220, 638)
$tbOU_Edit.Width = 500
$tbOU_Edit.ReadOnly = $true
$tbOU_Edit.Text = 'OU zdefiniowane dla aktywnej zakładki'
$form.Controls.Add($tbOU_Edit)

$btnEditOU = New-Object System.Windows.Forms.Button
$btnEditOU.Text = 'Odblokuj edycje OU'
$btnEditOU.Size = New-Object System.Drawing.Size(150, 25)
$btnEditOU.Location = New-Object System.Drawing.Point(730, 636)
$btnEditOU.Add_Click({ $tbOU_Edit.ReadOnly = -not $tbOU_Edit.ReadOnly; $btnEditOU.Text = if($tbOU_Edit.ReadOnly){'Odblokuj edycje OU'}else{'Zablokuj edycje OU'} })
$form.Controls.Add($btnEditOU)

$btnChooseOU = New-Object System.Windows.Forms.Button
$btnChooseOU.Text = 'Wybierz OU'
$btnChooseOU.Size = New-Object System.Drawing.Size(120, 25)
$btnChooseOU.Location = New-Object System.Drawing.Point(860, 636)
$btnChooseOU.Add_Click({ $dn = Show-OUChooser -Title 'Wybierz OU docelowe'; if ($dn){ $tbOU_Edit.Text=$dn; if ($tabs.SelectedTab){ $OU_Defaults[$tabs.SelectedTab.Text]=$dn } } })
$form.Controls.Add($btnChooseOU)

$tabs.add_SelectedIndexChanged({ if ($tabs.SelectedTab -and $tabs.SelectedTab.Text){ $tabName=$tabs.SelectedTab.Text; if ($tbDomain_Edit.ReadOnly -and $Domain_Defaults.ContainsKey($tabName)){$tbDomain_Edit.Text=$Domain_Defaults[$tabName]} if ($tbOU_Edit.ReadOnly -and $OU_Defaults.ContainsKey($tabName)){$tbOU_Edit.Text=$OU_Defaults[$tabName]} if ($cbDomain.Items.Contains($Domain_Defaults[$tabName])){ $cbDomain.SelectedItem=$Domain_Defaults[$tabName] } } })

# Tabs + grids
$grids=@{}
foreach($name in $tabNames){
  $tab = New-Object System.Windows.Forms.TabPage; $tab.Text=$name
  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.EditMode=[Windows.Forms.DataGridViewEditMode]::EditOnEnter
  $grid.Size = New-Object System.Drawing.Size(920, 430)
  $grid.Location = New-Object System.Drawing.Point(10, 10)
  $grid.AllowUserToAddRows=$true; $grid.AllowUserToDeleteRows=$true
  $grid.ColumnHeadersHeightSizeMode = [Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
  $grid.SelectionMode=[Windows.Forms.DataGridViewSelectionMode]::FullRowSelect; $grid.MultiSelect=$false
  switch($name){
    'Uczen'   { $columns=@('Imie','Nazwisko','NazwaWyswietlana','Login','Miasto','Email','Haslo') }
    'Student' { $columns=@('Imie','Nazwisko','NrAlbumu','NazwaWyswietlana','Login','Miasto','Email','Haslo') }
    default   { $columns=@('Imie','Nazwisko','NazwaWyswietlana','Login','Miasto','Email','Haslo') }
  }
  foreach($col in $columns){[void]$grid.Columns.Add($col,$col)}
  $yBtns=450

  $pasteBtn = New-Object System.Windows.Forms.Button
  $pasteBtn.Text='Wklej dane ze schowka'
  $pasteBtn.Size=New-Object System.Drawing.Size(200,25)
  $pasteBtn.Location=New-Object System.Drawing.Point(10,$yBtns)
  $pasteBtn.Add_Click({ try{ $selectedTab=$tabs.SelectedTab.Text; $targetGrid=$grids[$selectedTab]; $clip=Get-Clipboard -Raw; if([string]::IsNullOrWhiteSpace($clip)){ Write-ToTextBox 'Schowek jest pusty.' 'Warning'; return }
    $lines=$clip -split "`r?`n"; $targetGrid.SuspendLayout(); foreach($line in $lines){ if([string]::IsNullOrWhiteSpace($line)){continue}; $vals=$line -split "`t"; $exp = switch($selectedTab){ 'Uczen'{2} 'Student'{3} default{2} }; if($vals.Count -lt $exp){ Write-ToTextBox "Za malo kolumn dla $selectedTab" 'Warning'; continue }; if($vals.Count -gt $exp){ Write-ToTextBox "Zbyt wiele kolumn dla $selectedTab" 'Warning'; continue }; $ri=$targetGrid.Rows.Add(); $row=$targetGrid.Rows[$ri]; if($vals.Count -ge 1){$row.Cells['Imie'].Value=$vals[0].Trim()}; if($vals.Count -ge 2){$row.Cells['Nazwisko'].Value=$vals[1].Trim()}; if($selectedTab -eq 'Student' -and $vals.Count -ge 3){$row.Cells['NrAlbumu'].Value=$vals[2].Trim()} }
    $domain = if (-not $tbDomain_Edit.ReadOnly -and -not [string]::IsNullOrWhiteSpace($tbDomain_Edit.Text)){$tbDomain_Edit.Text.Trim()}else{$Domain_Defaults[$selectedTab]}
    for($r=0;$r -lt $targetGrid.Rows.Count;$r++){ $rRow=$targetGrid.Rows[$r]; if($rRow.IsNewRow){continue}; $imi=$rRow.Cells['Imie'].Value; $naz=$rRow.Cells['Nazwisko'].Value; if([string]::IsNullOrWhiteSpace($imi) -or [string]::IsNullOrWhiteSpace($naz)){continue}; $login=Remove-PolishDiacritics(("{0}.{1}" -f $imi.Substring(0,1),$naz)); if($selectedTab -eq 'Student'){ $album=$rRow.Cells['NrAlbumu'].Value; if(-not [string]::IsNullOrWhiteSpace($album)){ $rRow.Cells['Login'].Value=$album.Trim(); $rRow.Cells['Email'].Value=("{0}@{1}" -f $album.Trim(),$domain); $rRow.Cells['NazwaWyswietlana'].Value="$imi $naz ($album)" } else { $rRow.Cells['Login'].Value=$login; $rRow.Cells['Email'].Value="$login@$domain"; $rRow.Cells['NazwaWyswietlana'].Value="$imi $naz (Student)" } } else { $rRow.Cells['Login'].Value=$login; $rRow.Cells['Email'].Value="$login@$domain"; $rRow.Cells['NazwaWyswietlana'].Value="$imi $naz ($selectedTab)" }
      if(-not $rRow.Cells['Miasto'].Value){ $rRow.Cells['Miasto'].Value=$cbMiasto.SelectedItem }
      if(-not $rRow.Cells['Haslo'].Value){ $rRow.Cells['Haslo'].Value=New-RandomPassword }
    }
    Write-ToTextBox "Wklejono dane dla zakladki $selectedTab." 'Info' } catch { Write-ToTextBox "Błąd podczas wklejania danych: $_" 'Error' } finally { $targetGrid.ResumeLayout() } })

  $checkBtn = New-Object System.Windows.Forms.Button
  $checkBtn.Text='Sprawdz konta'
  $checkBtn.Size=New-Object System.Drawing.Size(150,25)
  $checkBtn.Location=New-Object System.Drawing.Point(220,$yBtns)
  $checkBtn.Add_Click({ if(-not (Test-ADModule)){return}; $selectedTab=$tabs.SelectedTab.Text; $targetGrid=$grids[$selectedTab]; Write-ToTextBox "Sprawdzam loginy dla zakladki $selectedTab..." 'Info'; for($r=0;$r -lt $targetGrid.Rows.Count;$r++){ $row=$targetGrid.Rows[$r]; if($row.IsNewRow){continue}; $loginVal=$row.Cells['Login'].Value; if([string]::IsNullOrWhiteSpace($loginVal)){continue}; $adUser=Get-ADUser -Filter "SamAccountName -eq '$loginVal'" -ErrorAction SilentlyContinue; if($adUser){ Write-ToTextBox "Login $loginVal jest ZAJETY w AD." 'Warning'; if($selectedTab -eq 'Uczen'){ $orig=$loginVal; $cnt=1; while(Get-ADUser -Filter "SamAccountName -eq '$loginVal'" -ErrorAction SilentlyContinue){ $loginVal="$orig$cnt"; $cnt++ }; $row.Cells['Login'].Value=$loginVal; $domain= if(-not $tbDomain_Edit.ReadOnly -and -not [string]::IsNullOrWhiteSpace($tbDomain_Edit.Text)){$tbDomain_Edit.Text.Trim()}else{$Domain_Defaults[$selectedTab]}; $row.Cells['Email'].Value="$loginVal@$domain"; Write-ToTextBox "Zmieniono login w wierszu #$r na: $loginVal" 'Info' } } else { Write-ToTextBox "Login $loginVal jest wolny." 'Info' } } })

  $refreshBtn = New-Object System.Windows.Forms.Button
  $refreshBtn.Text='Odswiez dane'
  $refreshBtn.Size=New-Object System.Drawing.Size(150,25)
  $refreshBtn.Location=New-Object System.Drawing.Point(400,$yBtns)
  $refreshBtn.Add_Click({ $selectedTab=$tabs.SelectedTab.Text; $targetGrid=$grids[$selectedTab]; $domain= if(-not $tbDomain_Edit.ReadOnly -and -not [string]::IsNullOrWhiteSpace($tbDomain_Edit.Text)){$tbDomain_Edit.Text.Trim()}else{$Domain_Defaults[$selectedTab]}; Write-ToTextBox "Odswiezam dane w zakladce: $selectedTab" 'Info'; for($r=0;$r -lt $targetGrid.Rows.Count;$r++){ $rRow=$targetGrid.Rows[$r]; if($rRow.IsNewRow){continue}; $imi=$rRow.Cells['Imie'].Value; $naz=$rRow.Cells['Nazwisko'].Value; if([string]::IsNullOrWhiteSpace($imi) -or [string]::IsNullOrWhiteSpace($naz)){continue}; $login=Remove-PolishDiacritics(("{0}.{1}" -f $imi.Substring(0,1),$naz)); if($selectedTab -eq 'Student'){ $album=$rRow.Cells['NrAlbumu'].Value; if(-not [string]::IsNullOrWhiteSpace($album)){ $rRow.Cells['Login'].Value=$album.Trim(); $rRow.Cells['Email'].Value=("{0}@{1}" -f $album.Trim(),$domain); $rRow.Cells['NazwaWyswietlana'].Value="$imi $naz ($album)" } else { $rRow.Cells['Login'].Value=$login; $rRow.Cells['Email'].Value="$login@$domain"; $rRow.Cells['NazwaWyswietlana'].Value="$imi $naz (Student)" } } else { $rRow.Cells['Login'].Value=$login; $rRow.Cells['Email'].Value="$login@$domain"; $rRow.Cells['NazwaWyswietlana'].Value="$imi $naz ($selectedTab)" }; if(-not $rRow.Cells['Miasto'].Value){$rRow.Cells['Miasto'].Value=$cbMiasto.SelectedItem}; if(-not $rRow.Cells['Haslo'].Value){$rRow.Cells['Haslo'].Value=New-RandomPassword} } })

  $tab.Controls.Add($grid); $tab.Controls.Add($pasteBtn); $tab.Controls.Add($checkBtn); $tab.Controls.Add($refreshBtn)
  $tabs.Controls.Add($tab); $grids[$name]=$grid
}

# Bottom buttons
$btnDodaj = New-Object System.Windows.Forms.Button
$btnDodaj.Text='Dodaj konta'
$btnDodaj.Location=New-Object System.Drawing.Point(10, 660)
$btnDodaj.Size=New-Object System.Drawing.Size(150,30)
$btnDodaj.Add_Click({ if(-not (Test-ADModule)){return}; $selectedTab=$tabs.SelectedTab.Text; $grid=$grids[$selectedTab]; $ouPath= if(-not $tbOU_Edit.ReadOnly -and -not [string]::IsNullOrWhiteSpace($tbOU_Edit.Text)){$tbOU_Edit.Text.Trim()}else{$OU_Defaults[$selectedTab]}; $domain= if(-not $tbDomain_Edit.ReadOnly -and -not [string]::IsNullOrWhiteSpace($tbDomain_Edit.Text)){$tbDomain_Edit.Text.Trim()}else{$Domain_Defaults[$selectedTab]}; if(-not $ouPath){Write-ToTextBox "Brak OU dla zakladki $selectedTab." 'Error'; return}; if(-not $domain){Write-ToTextBox "Brak domeny dla zakladki $selectedTab." 'Error'; return}; for($i=0;$i -lt $grid.Rows.Count;$i++){ $row=$grid.Rows[$i]; if($row.IsNewRow){continue}; try{ $imie=if($row.Cells['Imie'].Value){$row.Cells['Imie'].Value.ToString().Trim()}else{''}; $nazwisko=if($row.Cells['Nazwisko'].Value){$row.Cells['Nazwisko'].Value.ToString().Trim()}else{''}; if([string]::IsNullOrWhiteSpace($imie) -or [string]::IsNullOrWhiteSpace($nazwisko)){ Write-ToTextBox "Pominieto wiersz #$($i): brak imienia/nazwiska." 'Warning'; continue }; $miasto= if($row.Cells['Miasto'].Value){$row.Cells['Miasto'].Value.ToString().Trim()}else{$cbMiasto.SelectedItem}
      if($selectedTab -eq 'Student'){
        $nrAlbumu= if($row.Cells['NrAlbumu'].Value){$row.Cells['NrAlbumu'].Value.ToString().Trim()}else{''}; if([string]::IsNullOrWhiteSpace($nrAlbumu)){ Write-ToTextBox "Brak numeru albumu dla studenta: $imie $nazwisko" 'Warning'; continue }; if(Get-ADUser -Filter "SamAccountName -eq '$nrAlbumu'" -ErrorAction SilentlyContinue){ Write-ToTextBox "Uzytkownik $nrAlbumu juz istnieje - pomijam." 'Warning'; continue }
        $login=$nrAlbumu; $email="$login@$domain"; $upn=$email; $displayName="$imie $nazwisko ($nrAlbumu)"
        # unique CN in OU
        $cnBase=$login; $cn=$cnBase; $cnCount=0; while(Get-ADObject -LDAPFilter "(cn=$cn)" -SearchBase $ouPath -SearchScope OneLevel -ErrorAction SilentlyContinue){ $cnCount++; $cn="$cnBase$cnCount" }
      } else {
        $loginBase=Remove-PolishDiacritics(("{0}.{1}" -f $imie.Substring(0,1),$nazwisko)); $originalLogin=$loginBase; $counter=0; while(Get-ADUser -Filter "SamAccountName -eq '$loginBase'" -ErrorAction SilentlyContinue){ $counter++; $loginBase="$originalLogin$counter" }; $login=$loginBase; $email="$login@$domain"; $upn=$email; $displayName="$imie $nazwisko ($selectedTab)"
        if(Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -ErrorAction SilentlyContinue){ $base=$login; $j=1; do{ $login="$base$j"; $upn="$login@$domain"; $email=$upn; $j++ } while(Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -ErrorAction SilentlyContinue) }
        $cnBase=$login; $cn=$cnBase; $cnCount=0; while(Get-ADObject -LDAPFilter "(cn=$cn)" -SearchBase $ouPath -SearchScope OneLevel -ErrorAction SilentlyContinue){ $cnCount++; $cn="$cnBase$cnCount" }
      }

      $existingPass= if($row.Cells['Haslo'].Value){$row.Cells['Haslo'].Value.ToString()}else{''}
      $password= if([string]::IsNullOrWhiteSpace($existingPass)){ New-RandomPassword -Length 12 } else { $existingPass }
      $row.Cells['Haslo'].Value=$password; $secure=ConvertTo-SecureString $password -AsPlainText -Force
      New-ADUser -Name $cn -SamAccountName $login -UserPrincipalName $upn -DisplayName $displayName -GivenName $imie -Surname $nazwisko -EmailAddress $email -City $miasto -Path $ouPath -AccountPassword $secure -Enabled $true
      Write-ToTextBox "Dodano uzytkownika $login" 'Info'; $row.Cells['Login'].Value=$login; $row.Cells['Email'].Value=$email; $row.Cells['NazwaWyswietlana'].Value=$displayName
    } catch { Write-ToTextBox "Blad przy dodawaniu (wiersz #$i): $_" 'Error' } }
  Write-ToTextBox 'Synchronizacja z M365...' 'Info'; try { Start-ADSyncSyncCycle -PolicyType Delta } catch { Write-ToTextBox "Blad Start-ADSyncSyncCycle: $_" 'Warning' } })
$form.Controls.Add($btnDodaj)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Eksportuj dane do CSV'
$btnExport.Location = New-Object System.Drawing.Point(170, 660)
$btnExport.Size = New-Object System.Drawing.Size(180, 30)
$btnExport.Add_Click({ try{ $selectedTab=$tabs.SelectedTab.Text; $grid=$grids[$selectedTab]; if(-not $grid){ Write-ToTextBox "Brak gridu dla zakladki $selectedTab" 'Warning'; return }; $dlg=New-Object System.Windows.Forms.SaveFileDialog; $dlg.Filter='CSV file (*.csv)|*.csv|All files (*.*)|*.*'; $dlg.Title='Zapisz dane do CSV'; if($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK){return}; $path=$dlg.FileName; $colNames=$grid.Columns | ForEach-Object { $_.Name }; $lines=New-Object System.Collections.Generic.List[string]; $lines.Add(($colNames -join ',')); for($r=0;$r -lt $grid.Rows.Count;$r++){ $row=$grid.Rows[$r]; if($row.IsNewRow){continue}; $vals= foreach($cn in $colNames){ $val=$row.Cells[$cn].Value; if($null -eq $val){$val=''}; '"'+$val.ToString().Replace('"','""')+'"' }; $lines.Add(($vals -join ',')) }; [System.IO.File]::WriteAllLines($path,$lines,[System.Text.Encoding]::UTF8); Write-ToTextBox "Wyeksportowano dane do pliku: $path" 'Info' } catch { Write-ToTextBox "Blad eksportu CSV: $_" 'Error' } })
$form.Controls.Add($btnExport)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text='Wyczysc zakladke'
$btnClear.Size=New-Object System.Drawing.Size(150,30)
$btnClear.Location=New-Object System.Drawing.Point(360, 660)
$btnClear.Add_Click({ $selectedTab=$tabs.SelectedTab.Text; $grids[$selectedTab].Rows.Clear(); Write-ToTextBox "Wyczyszczono dane z zakladki $selectedTab." 'Info' })
$form.Controls.Add($btnClear)

$btnPomoc = New-Object System.Windows.Forms.Button
$btnPomoc.Text='Pomoc'
$btnPomoc.Size=New-Object System.Drawing.Size(100,30)
$btnPomoc.Location=New-Object System.Drawing.Point(760, 660)
$btnPomoc.Add_Click({
    $msg = @"
Skladnia wklejania (TAB):
- Uczen/Pracownik/Wykladowca/Inne: Imie[TAB]Nazwisko
- Student: Imie[TAB]Nazwisko[TAB]NrAlbumu

Konfiguracja (priorytet nad domyslna):
$ConfigPath
Tworzona automatycznie przy pierwszym uruchomieniu.
"@
    [Windows.Forms.MessageBox]::Show($msg,'Pomoc - Kreator Kont AD') | Out-Null
})
$form.Controls.Add($btnPomoc)

$btnZamknij = New-Object System.Windows.Forms.Button
$btnZamknij.Text='Zamknij'
$btnZamknij.Size=New-Object System.Drawing.Size(100,30)
$btnZamknij.Location=New-Object System.Drawing.Point(870, 660)
$btnZamknij.Add_Click({ $form.Close() })
$form.Controls.Add($btnZamknij)

$form.Topmost=$true
$form.Add_Shown({ $form.Activate(); try { $cbMiasto.Items.Clear(); [void]$cbMiasto.Items.AddRange($Cities); if($cbMiasto.Items.Count -gt 0){$cbMiasto.SelectedIndex=0} } catch {}; try { $cbDomain.Items.Clear(); [void]$cbDomain.Items.AddRange((Get-AvailableDomains)); if($tabs.SelectedTab){ $dnm=$Domain_Defaults[$tabs.SelectedTab.Text]; if($cbDomain.Items.Contains($dnm)){ $cbDomain.SelectedItem=$dnm } } } catch {}; if($tabs.SelectedTab){ $tabName=$tabs.SelectedTab.Text; if($Domain_Defaults.ContainsKey($tabName)){$tbDomain_Edit.Text=$Domain_Defaults[$tabName]}; if($OU_Defaults.ContainsKey($tabName)){$tbOU_Edit.Text=$OU_Defaults[$tabName]} } if($ConfigStatusMessage){ Write-ToTextBox $ConfigStatusMessage 'Info' } })
[void]$form.ShowDialog()
#endregion

