#region Init
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Stop'

# Hidden JSON config in start directory (priority over inline defaults)
$tabNames = @('Uczen', 'Student', 'Pracownik', 'Wykladowca', 'Inne')
# Legacy per-repo config removed. Only .AD-BulkUserCreator.json is used.

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
$Cities_Default = @('Warszawa', 'Olsztyn', 'Katowice', 'Poznań', 'Szczecin', 'Człuchów', 'Lublin')

$Domain_Defaults = $Domain_Defaults_Default.Clone()
$OU_Defaults = $OU_Defaults_Default.Clone()
$Cities = @($Cities_Default)

# Per-tab settings (defaults). These are persisted to .AD-BulkUserCreator.json
$LoginFormat_ByTab = @{}
$DisplayNameFormat_ByTab = @{}
# Per-tab WhatIf flags (not persisted; default OFF)
$WhatIf_ByTab = @{}
foreach ($t in $tabNames) {
  $LoginFormat_ByTab[$t] = 'i.nazwisko'
  $DisplayNameFormat_ByTab[$t] = '{Imie} {Nazwisko} ({Rola})'
}

# Guard to prevent saving while initializing UI
$script:IsInitializing = $true

function Write-ToTextBox {
  param(
    [string]$Text,
    [ValidateSet('Info', 'Warning', 'Error', 'Success')][string]$Type = 'Info'
  )

  $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  switch ($Type) {
    'Info' { $color = [Drawing.Color]::FromArgb(0, 120, 215); $prefix = '[INFO]' }
    'Warning' { $color = [Drawing.Color]::FromArgb(199, 125, 0); $prefix = '[WARNING]' }
    'Error' { $color = [Drawing.Color]::FromArgb(192, 0, 0); $prefix = '[ERROR]' }
    'Success' { $color = [Drawing.Color]::FromArgb(16, 124, 16); $prefix = '[SUCCESS]' }
    default { $color = $tb_logg_box.ForeColor; $prefix = '[INFO]' }
  }

  $entry = '{0} {1} {2}' -f $timestamp, $prefix, $Text
  $tb_logg_box.SelectionStart = $tb_logg_box.TextLength
  $tb_logg_box.SelectionLength = 0
  $tb_logg_box.SelectionColor = $color
  $tb_logg_box.AppendText($entry + [Environment]::NewLine)
  $tb_logg_box.ScrollToCaret()
  $tb_logg_box.SelectionColor = $tb_logg_box.ForeColor
}

function Test-ConfigObject([object]$cfg) {
  if (-not $cfg) { return $false }
  foreach ($k in 'Domain_Defaults', 'OU_Defaults') { if (-not ($cfg.PSObject.Properties.Name -contains $k)) { return $false } }
  foreach ($t in $tabNames) { if (-not ($cfg.Domain_Defaults.PSObject.Properties.Name -contains $t)) { return $false } }
  foreach ($t in $tabNames) { if (-not ($cfg.OU_Defaults.PSObject.Properties.Name -contains $t)) { return $false } }
  return $true
}

# Removed legacy config messages list

function Save-Defaults {
  try {
    Set-ABCSettings
    return $true
  }
  catch {
    Write-ToTextBox ('Nie udalo sie zapisac konfiguracji: ' + $_.Exception.Message) 'Warning'
    return $false
  }
}

# utils
function Remove-PolishDiacritics {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }

  $normalized = $Text.Normalize([System.Text.NormalizationForm]::FormD)
  $builder = New-Object System.Text.StringBuilder
  foreach ($char in $normalized.ToCharArray()) {
    if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($char) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$builder.Append($char)
    }
  }
  $clean = $builder.ToString().Normalize([System.Text.NormalizationForm]::FormC)

  $clean = $clean.Replace([char]0x0142, 'l')
  $clean = $clean.Replace([char]0x0141, 'L')

  return $clean
}


function Get-Login {
  param(
    [string]$Imie,
    [string]$Nazwisko,
    [string]$Format = 'i.nazwisko'
  )

  $i = (Remove-PolishDiacritics($Imie)).Trim().ToLower()
  $n = (Remove-PolishDiacritics($Nazwisko)).Trim().ToLower()

  $i = ($i -replace "[\s\-']", '')
  $n = ($n -replace "[\s\-']", '')
  if ([string]::IsNullOrWhiteSpace($i) -or [string]::IsNullOrWhiteSpace($n)) { return '' }

  $firstLetter = if ($i.Length -gt 0) { $i.Substring(0, 1) } else { '' }

  switch ($Format) {
    'i.nazwisko' { return '{0}.{1}' -f $firstLetter, $n }
    'inazwisko' { return '{0}{1}' -f $firstLetter, $n }
    'imie.nazwisko' { return '{0}.{1}' -f $i, $n }
    'nazwisko.imie' { return '{0}.{1}' -f $n, $i }
    default { return '{0}.{1}' -f $firstLetter, $n }
  }
}
# === SETTINGS ===
$Script:SettingsPath = try { Join-Path -Path $PSScriptRoot -ChildPath '.AD-BulkUserCreator.json' } catch { Join-Path -Path $env:USERPROFILE -ChildPath '.AD-BulkUserCreator.json' }
$Script:Settings = @{
  LoginFormat       = 'i.nazwisko'
  EmailDomain       = 'example.edu'
  DisplayNameFormat = '{Imie} {Nazwisko} ({Rola})'
}



function Get-ABCSettings {
  if (Test-Path -LiteralPath $Script:SettingsPath) {
    try {
      $json = Get-Content -LiteralPath $Script:SettingsPath -Raw -ErrorAction Stop
      $data = $json | ConvertFrom-Json -ErrorAction Stop
    } catch {
      try { [Windows.Forms.MessageBox]::Show('Plik ustawien jest uszkodzony i zostanie podmieniony na szablon.','Ustawienia', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning) | Out-Null } catch {}
      if (Write-DefaultConfig) {
        try {
          $json = Get-Content -LiteralPath $Script:SettingsPath -Raw -ErrorAction Stop
          $data = $json | ConvertFrom-Json -ErrorAction Stop
        } catch {}
      }
    }
    if ($null -ne $data) {
      if ($data.LoginFormat) { $Script:Settings.LoginFormat = [string]$data.LoginFormat }
      if ($data.EmailDomain) { $Script:Settings.EmailDomain = [string]$data.EmailDomain }
      if ($data.DisplayNameFormat) { $Script:Settings.DisplayNameFormat = [string]$data.DisplayNameFormat }
      if ($data.PSObject.Properties.Name -contains 'LoginFormatByTab' -and $data.LoginFormatByTab) {
        foreach ($k in $tabNames) { if ($data.LoginFormatByTab.PSObject.Properties.Name -contains $k) { $LoginFormat_ByTab[$k] = [string]$data.LoginFormatByTab.$k } }
      } elseif ($data.LoginFormat) {
        foreach ($k in $tabNames) { $LoginFormat_ByTab[$k] = [string]$data.LoginFormat }
      }
      if ($data.PSObject.Properties.Name -contains 'DisplayNameFormatByTab' -and $data.DisplayNameFormatByTab) {
        foreach ($k in $tabNames) { if ($data.DisplayNameFormatByTab.PSObject.Properties.Name -contains $k) { $DisplayNameFormat_ByTab[$k] = [string]$data.DisplayNameFormatByTab.$k } }
      } elseif ($data.DisplayNameFormat) {
        foreach ($k in $tabNames) { $DisplayNameFormat_ByTab[$k] = [string]$data.DisplayNameFormat }
      }
      if ($data.PSObject.Properties.Name -contains 'Domain_Defaults' -and $data.Domain_Defaults) {
        $script:Domain_Defaults = @{}
        foreach ($k in $tabNames) {
          if ($data.Domain_Defaults.PSObject.Properties.Name -contains $k) {
            $script:Domain_Defaults[$k] = [string]$data.Domain_Defaults.$k
          }
        }
      }
      if ($data.PSObject.Properties.Name -contains 'OU_Defaults' -and $data.OU_Defaults) {
        $script:OU_Defaults = @{}
        foreach ($k in $tabNames) {
          if ($data.OU_Defaults.PSObject.Properties.Name -contains $k) {
            $script:OU_Defaults[$k] = [string]$data.OU_Defaults.$k
          }
        }
      }
      if ($data.PSObject.Properties.Name -contains 'Cities' -and $data.Cities) {
        $script:Cities = @($data.Cities)
      }
      Write-ToTextBox "Wczytano ustawienia z $($Script:SettingsPath)" 'Info'
    } else {
      Write-ToTextBox "Brak danych ustawien po wczytaniu, uzywam domyslnych." 'Warning'
    }
  } else {
    try { [Windows.Forms.MessageBox]::Show('Brak pliku ustawien – zostanie utworzony szablon.','Ustawienia', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information) | Out-Null } catch {}
    if (Write-DefaultConfig) { Write-ToTextBox "Utworzono plik ustawien: $($Script:SettingsPath)" 'Info' }
  }
}

function Set-ABCSettings {
  try {
    $out = [ordered]@{
      EmailDomain            = $Script:Settings.EmailDomain
      LoginFormat            = $Script:Settings.LoginFormat
      DisplayNameFormat      = $Script:Settings.DisplayNameFormat
      LoginFormatByTab       = $LoginFormat_ByTab
      DisplayNameFormatByTab = $DisplayNameFormat_ByTab
      Domain_Defaults        = $Domain_Defaults
      OU_Defaults            = $OU_Defaults
      Cities                 = $Cities
    }
    $out | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Script:SettingsPath -Encoding UTF8
    Write-ToTextBox "Zapisano ustawienia do $($Script:SettingsPath)" 'Info'
  }
  catch {
    Write-ToTextBox "Błąd zapisu ustawień: $_" 'Error'
  }
}
function Write-DefaultConfig {
  try {
    $out = [ordered]@{
      EmailDomain            = $Script:Settings.EmailDomain
      LoginFormat            = $Script:Settings.LoginFormat
      DisplayNameFormat      = $Script:Settings.DisplayNameFormat
      LoginFormatByTab       = $LoginFormat_ByTab
      DisplayNameFormatByTab = $DisplayNameFormat_ByTab
      Domain_Defaults        = $Domain_Defaults
      OU_Defaults            = $OU_Defaults
      Cities                 = $Cities
    }
    $out | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Script:SettingsPath -Encoding UTF8
    return $true
  } catch {
    Write-ToTextBox ('Nie udalo sie zapisac domyslnego pliku ustawien: ' + $_.Exception.Message) 'Error'
    try { [Windows.Forms.MessageBox]::Show('Nie udalo sie zapisac domyslnego pliku ustawien: ' + $_.Exception.Message,'Blad', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error) | Out-Null } catch {}
    return $false
  }
}

function Show-ABCSettingsDialog {
  $dlg = New-Object System.Windows.Forms.Form
  $dlg.Text = 'Ustawienia'
  $dlg.StartPosition = 'CenterParent'
  $dlg.AutoSize = $true
  $dlg.AutoSizeMode = 'GrowAndShrink'

  $table = New-Object System.Windows.Forms.TableLayoutPanel
  $table.AutoSize = $true
  $table.ColumnCount = 2
  $table.RowCount = 4
  $table.Padding = [System.Windows.Forms.Padding]::new(10)
  $dlg.Controls.Add($table)

  # Login format
  $lblFmt = New-Object System.Windows.Forms.Label
  $lblFmt.Text = 'Format loginu'
  $lblFmt.AutoSize = $true
  $table.Controls.Add($lblFmt, 0, 0)

  $cbFmt = New-Object System.Windows.Forms.ComboBox
  $cbFmt.DropDownStyle = 'DropDownList'
  $null = $cbFmt.Items.AddRange(@('i.nazwisko', 'inazwisko', 'imie.nazwisko', 'nazwisko.imie'))
  $cbFmt.SelectedItem = $Script:Settings.LoginFormat
  $table.Controls.Add($cbFmt, 1, 0)

  # Email domain
  $lblDom = New-Object System.Windows.Forms.Label
  $lblDom.Text = 'Domena e-mail'
  $lblDom.AutoSize = $true
  $table.Controls.Add($lblDom, 0, 1)

  $tbDom = New-Object System.Windows.Forms.TextBox
  $tbDom.Text = $Script:Settings.EmailDomain
  $tbDom.Width = 240
  $table.Controls.Add($tbDom, 1, 1)

  # Display name
  $lblDn = New-Object System.Windows.Forms.Label
  $lblDn.Text = 'Format nazwy wyświetlanej'
  $lblDn.AutoSize = $true
  $table.Controls.Add($lblDn, 0, 2)

  $tbDn = New-Object System.Windows.Forms.TextBox
  $tbDn.Text = $Script:Settings.DisplayNameFormat
  $tbDn.Width = 240
  $table.Controls.Add($tbDn, 1, 2)

  # Buttons
  $pBtns = New-Object System.Windows.Forms.FlowLayoutPanel
  $pBtns.FlowDirection = 'RightToLeft'
  $pBtns.AutoSize = $true
  $table.Controls.Add($pBtns, 0, 3)
  $table.SetColumnSpan($pBtns, 2)

  $btnOK = New-Object System.Windows.Forms.Button
  $btnOK.Text = 'Zapisz'
  $btnOK.Add_Click({
      $Script:Settings.LoginFormat = [string]$cbFmt.SelectedItem
      $Script:Settings.EmailDomain = [string]$tbDom.Text.Trim()
      $Script:Settings.DisplayNameFormat = [string]$tbDn.Text
      Set-ABCSettings
      $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
      $dlg.Close()
    })
  & $styleButton $btnOK
  $pBtns.Controls.Add($btnOK)

  $btnCancel = New-Object System.Windows.Forms.Button
  $btnCancel.Text = 'Anuluj'
  $btnCancel.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })
  & $styleButton $btnCancel
  $pBtns.Controls.Add($btnCancel)

  [void]$dlg.ShowDialog()
}

function Get-UniqueLogin {
  param(
    [string]$BaseLogin,
    [string]$Rola # 'Student','Wykladowca','Uczen','Pracownik','Inne' itp.
  )
  if ([string]::IsNullOrWhiteSpace($BaseLogin)) { return $BaseLogin }
  # Tylko Student i Wykladowca mają numerację wg wymagań
  $needsNumbering = @('Student', 'Wykładowca', 'Wykladowca') -contains $Rola
  if (-not $needsNumbering) { return $BaseLogin }
  try {
    $existing = Get-ADUser -LDAPFilter "(sAMAccountName=$BaseLogin*)" -Properties sAMAccountName | Select-Object -ExpandProperty sAMAccountName
  }
  catch {
    Write-ToTextBox "Get-ADUser nieosiągalne podczas sprawdzania loginów: $_" 'Warning'
    return $BaseLogin
  }
  if (-not $existing) { return $BaseLogin }
  if ($existing -notcontains $BaseLogin) { return $BaseLogin } # bazowy wolny, choć istnieją inne z podobnym prefiksem
  # Szukaj wolnego sufiksu numerycznego
  for ($i = 1; $i -lt 1000; $i++) {
    $candidate = "$BaseLogin$i"
    if ($existing -notcontains $candidate) { return $candidate }
  }
  return $BaseLogin
}


function Get-RandChar([Parameter(Mandatory)][string]$Pool) {
  $rng = New-Object System.Security.Cryptography.RNGCryptoServiceProvider
  $bytes = New-Object 'Byte[]' 4; $rng.GetBytes($bytes)
  $idx = [math]::Abs([BitConverter]::ToInt32($bytes, 0)) % $Pool.Length
  $Pool[$idx]
}
function Test-Sequential([char]$Prev, [char]$Curr) { if (-not $Prev) { return $false }; return ([math]::Abs([int][char]$Prev - [int][char]$Curr) -eq 1) }
function New-RandomPassword([ValidateRange(8, 128)][int]$Length = 12) {
  $chars = @{Digits = '123456789'; Lower = 'abcdefghjkmnpqrstuvwxyz'; Upper = 'ABCDEFGHJKMNPQRSTUVWXYZ'; Symbols = '#$%&?@' }
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append((Get-RandChar -Pool ($chars.Lower + $chars.Upper)))
  foreach ($k in 'Digits', 'Lower', 'Upper', 'Symbols') { if ($sb.Length -ge $Length) { break }; $ch = Get-RandChar -Pool $chars[$k]; if (Test-Sequential $sb[$sb.Length - 1] $ch) { $ch = Get-RandChar -Pool $chars[$k] }; [void]$sb.Append($ch) }
  $all = ($chars.Digits + $chars.Lower + $chars.Upper + $chars.Symbols)
  while ($sb.Length -lt $Length) { $ch = Get-RandChar -Pool $all; $p = if ($sb.Length -gt 0) { $sb[$sb.Length - 1] }else { [char]0 }; $p2 = if ($sb.Length -gt 1) { $sb[$sb.Length - 2] }else { [char]0 }; if (Test-Sequential $p $ch) { continue }; if ($p -eq $ch -and $p2 -eq $ch) { continue }; [void]$sb.Append($ch) }
  $pass = $sb.ToString(); $ok = ($pass -cmatch '[0-9]') -and ($pass -cmatch '[a-z]') -and ($pass -cmatch '[A-Z]') -and ($pass -cmatch '[#\$%&\?@]')
  if ($ok) { $pass }else { New-RandomPassword -Length $Length }
}

function Test-ADModule {
  try {
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { throw 'Brak modułu ActiveDirectory (RSAT).' }
    if (-not (Get-Module -Name ActiveDirectory)) { Import-Module ActiveDirectory -ErrorAction Stop | Out-Null }
    return $true
  }
  catch {
    [System.Windows.Forms.MessageBox]::Show("Brak modułu ActiveDirectory. Zainstaluj RSAT: Active Directory i spróbuj ponownie.`r
$_", 'Błąd modułu AD', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    return $false
  }
}

function Get-AvailableDomains {
  if (-not (Test-ADModule)) { return @() }
  try {
    $forest = Get-ADForest -ErrorAction Stop
    @($forest.RootDomain, $forest.Domains, $forest.UPNSuffixes) | ForEach-Object { $_ } | Where-Object { $_ } | Select-Object -Unique | Sort-Object
  }
  catch { @() }
}

function Show-OUChooser {
  param(
    [string]$Title = 'Wybierz OU',
    [System.Windows.Forms.Form]$Owner = $null,
    [string]$InitialDn = $null
  )
  if (-not (Test-ADModule)) { return $null }

  $dialog = New-Object Windows.Forms.Form
  $dialog.Text = $Title
  $dialog.Font = New-Object Drawing.Font('Segoe UI', 10)
  $dialog.Size = [Drawing.Size]::new(700, 750)
  $dialog.StartPosition = if ($Owner) { 'CenterParent' } else { 'CenterScreen' }
  $dialog.ShowInTaskbar = -not [bool]$Owner
  $dialog.MaximizeBox = $false
  $dialog.MinimizeBox = $false
  $dialog.TopMost = $true
  $dialog.Tag = $null

  $tree = New-Object Windows.Forms.TreeView
  $tree.Dock = 'Fill'
  $tree.HideSelection = $false

  $ok = New-Object Windows.Forms.Button
  $ok.Text = 'Wybierz'
  $ok.AutoSize = $true
  $ok.Padding = [Windows.Forms.Padding]::new(12, 6, 12, 6)

  $cancel = New-Object Windows.Forms.Button
  $cancel.Text = 'Anuluj'
  $cancel.AutoSize = $true
  $cancel.Padding = [Windows.Forms.Padding]::new(12, 6, 12, 6)

  $buttonBar = New-Object Windows.Forms.FlowLayoutPanel
  $buttonBar.Dock = 'Fill'
  $buttonBar.FlowDirection = [Windows.Forms.FlowDirection]::RightToLeft
  $buttonBar.AutoSize = $true
  $buttonBar.AutoSizeMode = 'GrowAndShrink'
  $buttonBar.Padding = [Windows.Forms.Padding]::new(0, 12, 0, 0)
  $buttonBar.WrapContents = $false
  $buttonBar.Controls.Add($ok) | Out-Null
  $buttonBar.Controls.Add($cancel) | Out-Null
  $cancel.Margin = [Windows.Forms.Padding]::new(0, 0, 8, 0)

  $root = New-Object Windows.Forms.TableLayoutPanel
  $root.Dock = 'Fill'
  $root.RowCount = 2
  $root.ColumnCount = 1
  [void]$root.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent, 100)))
  [void]$root.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize)))
  $root.Controls.Add($tree, 0, 0)
  $root.Controls.Add($buttonBar, 0, 1)
  $dialog.Controls.Add($root)
  $dialog.AcceptButton = $ok
  $dialog.CancelButton = $cancel

  $ok.Add_Click({
      if ($tree.SelectedNode -and $tree.SelectedNode.Tag) {
        $dialog.Tag = [string]$tree.SelectedNode.Tag
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
      }
    })
  $tree.Add_NodeMouseDoubleClick({
      if ($_.Node -and $_.Node.Tag) {
        $dialog.Tag = [string]$_.Node.Tag
        $dialog.DialogResult = [System.Windows.Forms.DialogResult]::OK
      }
    })
  $cancel.Add_Click({
      $dialog.Tag = $null
      $dialog.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    })

  try {
    $nodes = @{}
    $ous = Get-ADOrganizationalUnit -Filter * -SearchScope Subtree -ErrorAction Stop | Sort-Object DistinguishedName
    foreach ($ou in $ous) {
      $dn = [string]$ou.DistinguishedName
      $nodes[$dn] = New-Object Windows.Forms.TreeNode -Property @{ Text = $ou.Name; Tag = $dn }
    }
    foreach ($ou in $ous) {
      $dn = [string]$ou.DistinguishedName
      $parent = if ($dn.Contains(',')) { $dn.Substring($dn.IndexOf(',') + 1) } else { $null }
      if ($parent -and $nodes.ContainsKey($parent)) {
        [void]$nodes[$parent].Nodes.Add($nodes[$dn])
      }
    }
    foreach ($item in $nodes.GetEnumerator()) {
      $dn = $item.Key
      $parent = if ($dn.Contains(',')) { $dn.Substring($dn.IndexOf(',') + 1) } else { $null }
      if (-not $nodes.ContainsKey($parent)) {
        [void]$tree.Nodes.Add($item.Value)
      }
    }
    foreach ($node in $tree.Nodes) { $node.Expand() }
    if ($InitialDn -and $nodes.ContainsKey($InitialDn)) {
      $tree.SelectedNode = $nodes[$InitialDn]
      $tree.SelectedNode.EnsureVisible()
    }
  }
  catch {
    [Windows.Forms.MessageBox]::Show("Nie udalo sie pobrac OU: $($_.Exception.Message)", 'Blad') | Out-Null
    return $null
  }

  $previousTopMost = $null
  if ($null -ne $Owner) {
    $previousTopMost = $Owner.TopMost
    $Owner.TopMost = $false
  }

  try {
    $dialogResult = if ($null -ne $Owner) { $dialog.ShowDialog($Owner) } else { $dialog.ShowDialog() }
  }
  finally {
    if ($null -ne $Owner) {
      $Owner.TopMost = $previousTopMost
    }
  }

  if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK -and $dialog.Tag) {
    return [string]$dialog.Tag
  }

  return $null
}



#endregion

#region GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Kreator Kont AD'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.StartPosition = 'CenterScreen'
$form.ClientSize = New-Object System.Drawing.Size(1000, 780)
$form.MinimumSize = New-Object System.Drawing.Size(900, 700)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false

$layoutRoot = New-Object System.Windows.Forms.TableLayoutPanel
$layoutRoot.Dock = 'Fill'
$layoutRoot.ColumnCount = 1
$layoutRoot.RowCount = 5
$layoutRoot.Padding = [System.Windows.Forms.Padding]::new(18, 18, 18, 18)
$layoutRoot.BackColor = [System.Drawing.Color]::Transparent
[void]$layoutRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$layoutRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$layoutRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$layoutRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$layoutRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 150)))
$form.Controls.Add($layoutRoot)

$lblUsage = New-Object System.Windows.Forms.Label
$lblUsage.Text = 'Wklej Imie i Nazwisko (Student dodatkowo NrAlbumu). Wybierz domene i OU dla aktywnej zakladki, nastepnie dodaj konta.'
$lblUsage.AutoSize = $true
$lblUsage.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 12)
$lblUsage.ForeColor = [System.Drawing.Color]::FromArgb(55, 61, 69)
$layoutRoot.Controls.Add($lblUsage, 0, 0)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = 'Fill'
$tabs.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 12)
$layoutRoot.Controls.Add($tabs, 0, 1)

$configPanel = New-Object System.Windows.Forms.TableLayoutPanel
$configPanel.ColumnCount = 3
$configPanel.RowCount = 4
$configPanel.Dock = 'Fill'
$configPanel.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 12)
$configPanel.AutoSize = $true
$configPanel.AutoSizeMode = 'GrowAndShrink'
$configPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
$configPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$configPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$configPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$configPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$configPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$configPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$layoutRoot.Controls.Add($configPanel, 0, 2)

$styleButton = {
  param(
    [System.Windows.Forms.Button]$Button,
    [switch]$Primary
  )
  $Button.AutoSize = $true
  $Button.UseVisualStyleBackColor = $true
  $Button.Padding = [System.Windows.Forms.Padding]::new(12, 6, 12, 6)
  $Button.Margin = [System.Windows.Forms.Padding]::new(8, 0, 0, 0)
  if ($Primary) {
    $Button.Font = New-Object System.Drawing.Font($Button.Font, [System.Drawing.FontStyle]::Bold)
  }
}

$lblMiasto = New-Object System.Windows.Forms.Label
$lblMiasto.Text = 'Domyslne miasto:'
$lblMiasto.AutoSize = $true
$lblMiasto.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 6)
$configPanel.Controls.Add($lblMiasto, 0, 0)

$cbMiasto = New-Object System.Windows.Forms.ComboBox
$cbMiasto.DropDownStyle = 'DropDownList'
$cbMiasto.Width = 260
$cbMiasto.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 6)
$configPanel.Controls.Add($cbMiasto, 1, 0)
$configPanel.SetColumnSpan($cbMiasto, 2)

$lblDomain = New-Object System.Windows.Forms.Label
$lblDomain.Text = 'Domena/UPN:'
$lblDomain.AutoSize = $true
$lblDomain.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 6)
$configPanel.Controls.Add($lblDomain, 0, 1)

$cbDomain = New-Object System.Windows.Forms.ComboBox
$cbDomain.DropDownStyle = 'DropDownList'
$cbDomain.Width = 300
$cbDomain.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 6)
$configPanel.Controls.Add($cbDomain, 1, 1)
$configPanel.SetColumnSpan($cbDomain, 2)

$tbDomain_Edit = New-Object System.Windows.Forms.TextBox
$tbDomain_Edit.ReadOnly = $true
$tbDomain_Edit.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$tbDomain_Edit.BackColor = $form.BackColor
$tbDomain_Edit.Dock = 'Fill'
$tbDomain_Edit.Text = '(uzywana domena dla aktywnej zakladki)'
$tbDomain_Edit.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 6)
$configPanel.Controls.Add($tbDomain_Edit, 0, 2)
$configPanel.SetColumnSpan($tbDomain_Edit, 3)

$lblOU = New-Object System.Windows.Forms.Label
$lblOU.Text = 'OU docelowe:'
$lblOU.AutoSize = $true
$lblOU.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 0)
$configPanel.Controls.Add($lblOU, 0, 3)

$tbOU_Edit = New-Object System.Windows.Forms.TextBox
$tbOU_Edit.ReadOnly = $true
$tbOU_Edit.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$tbOU_Edit.BackColor = $form.BackColor
$tbOU_Edit.Dock = 'Fill'
$tbOU_Edit.Text = 'OU zdefiniowane dla aktywnej zakladki'
$tbOU_Edit.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 0)
$configPanel.Controls.Add($tbOU_Edit, 1, 3)

$btnChooseOU = New-Object System.Windows.Forms.Button
$btnChooseOU.Text = 'Wybierz OU'
& $styleButton $btnChooseOU
$btnChooseOU.Margin = [System.Windows.Forms.Padding]::new(8, 0, 0, 0)
$configPanel.Controls.Add($btnChooseOU, 2, 3)
$btnChooseOU.Add_Click({
    $tabName = if ($tabs.SelectedTab -and $tabs.SelectedTab.Text) { [string]$tabs.SelectedTab.Text } else { $null }
    $currentDn = if ($tabName -and $OU_Defaults.ContainsKey($tabName)) { $OU_Defaults[$tabName] } else { $null }
    $dn = Show-OUChooser -Title 'Wybierz OU docelowe' -Owner $form -InitialDn $currentDn
    if ($dn) {
      $tbOU_Edit.Text = $dn
      if ($tabName) {
        $OU_Defaults[$tabName] = $dn
        if (Save-Defaults) {
          Write-ToTextBox "Ustawiono OU dla zakladki ${tabName}: $dn" 'Info'
        }
        else {
          Write-ToTextBox "Nie udalo sie zapisac OU dla zakladki ${tabName}." 'Warning'
        }
      }
    }
  })

# --- Format loginu ---
[void]$configPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$lblFormat = New-Object System.Windows.Forms.Label
$lblFormat.Text = 'Format loginu:'
$lblFormat.AutoSize = $true
$lblFormat.Margin = [System.Windows.Forms.Padding]::new(0, 12, 12, 6)
$configPanel.Controls.Add($lblFormat, 0, 4)

$cbLoginFormat = New-Object System.Windows.Forms.ComboBox
$cbLoginFormat.DropDownStyle = 'DropDownList'
$null = $cbLoginFormat.Items.AddRange(@('i.nazwisko', 'inazwisko', 'imie.nazwisko', 'nazwisko.imie'))
$cbLoginFormat.SelectedIndex = 0
$cbLoginFormat.Margin = [System.Windows.Forms.Padding]::new(0, 8, 0, 0)
$cbLoginFormat.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$configPanel.Controls.Add($cbLoginFormat, 1, 4)
$cbLoginFormat.Visible = $true
# Ustaw wstępnie wg aktywnej zakładki i reaguj na zmianę
# ToolTip for login format (used to show Student lock info)
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.InitialDelay = 200
$toolTip.ReshowDelay = 200
$toolTip.AutoPopDelay = 5000

$toolTip.SetToolTip($cbLoginFormat, '')
try {
  if ($tabs.SelectedTab) {
    $lf = $LoginFormat_ByTab[$tabs.SelectedTab.Text]
    $idx = $cbLoginFormat.Items.IndexOf($lf)
    if ($idx -ge 0) { $cbLoginFormat.SelectedIndex = $idx }
  }
}
catch {}
$cbLoginFormat.Add_SelectedIndexChanged({
  if ($script:IsInitializing) { return }
  if ($cbLoginFormat.SelectedItem -and $tabs.SelectedTab) {
    $LoginFormat_ByTab[$tabs.SelectedTab.Text] = [string]$cbLoginFormat.SelectedItem; $Script:Settings.LoginFormat = [string]$cbLoginFormat.SelectedItem
    try { Set-ABCSettings } catch {}
  }
})

# --- DisplayName format ---
[void]$configPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$lblDN = New-Object System.Windows.Forms.Label
$lblDN.Text = 'Format nazwy wyświetlanej:'
$lblDN.AutoSize = $true
$lblDN.Margin = [System.Windows.Forms.Padding]::new(0, 12, 12, 6)
$configPanel.Controls.Add($lblDN, 0, 5)

$tbDN = New-Object System.Windows.Forms.TextBox
$tbDN.Text = [string]$Script:Settings.DisplayNameFormat
$tbDN.Dock = 'Fill'
$tbDN.Margin = [System.Windows.Forms.Padding]::new(0, 8, 0, 0)
$configPanel.Controls.Add($tbDN, 1, 5)
$configPanel.SetColumnSpan($tbDN, 2)
$tbDN.Add_TextChanged({
  if ($script:IsInitializing) { return }
  if ($tabs.SelectedTab) {
    $DisplayNameFormat_ByTab[$tabs.SelectedTab.Text] = [string]$tbDN.Text
    try { Set-ABCSettings } catch {}
  }
})
# Info label dla zakladki Student (format loginu zablokowany)
$lblStudentFmtInfo = New-Object System.Windows.Forms.Label
$lblStudentFmtInfo.AutoSize = $true
$lblStudentFmtInfo.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
$lblStudentFmtInfo.Text = [string]::Empty
$configPanel.Controls.Add($lblStudentFmtInfo, 2, 4)
$configPanel.SetColumnSpan($cbLoginFormat, 1)
$actionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionPanel.AutoSize = $true
$actionPanel.AutoSizeMode = 'GrowAndShrink'
$actionPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$actionPanel.WrapContents = $false
$actionPanel.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 12)
$actionPanel.Padding = [System.Windows.Forms.Padding]::new(0)
$actionPanel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$layoutRoot.Controls.Add($actionPanel, 0, 3)

$tb_logg_box = New-Object System.Windows.Forms.RichTextBox
$tb_logg_box.Dock = 'Fill'
$tb_logg_box.ReadOnly = $true
$tb_logg_box.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$tb_logg_box.BackColor = [System.Drawing.SystemColors]::Window
$tb_logg_box.Margin = [System.Windows.Forms.Padding]::new(0)
$tb_logg_box.Font = New-Object System.Drawing.Font('Consolas', 9)
$tb_logg_box.ForeColor = [System.Drawing.Color]::FromArgb(32, 32, 32)
$layoutRoot.Controls.Add($tb_logg_box, 0, 4)

$cbDomain.Add_SelectedIndexChanged({
    if ($script:IsInitializing) { return }
    if ($tabs.SelectedTab -and $cbDomain.SelectedItem) {
      $tabName = [string]$tabs.SelectedTab.Text
      $sel = [string]$cbDomain.SelectedItem
      $Domain_Defaults[$tabName] = $sel
      $tbDomain_Edit.Text = $sel
      $Script:Settings.EmailDomain = $sel
      if (-not (Save-Defaults)) {
        Write-ToTextBox "Nie udalo sie zapisac domeny dla zakladki ${tabName}." 'Warning'
      }
    }
  })

$tabs.Add_SelectedIndexChanged({
    $script:IsInitializing = $true
    if ($tabs.SelectedTab -and $tabs.SelectedTab.Text) {
      $tabName = $tabs.SelectedTab.Text
      if ($Domain_Defaults.ContainsKey($tabName)) {
        $tbDomain_Edit.Text = $Domain_Defaults[$tabName]
        $Script:Settings.EmailDomain = [string]$Domain_Defaults[$tabName]
        if ($cbDomain.Items.Contains($Domain_Defaults[$tabName])) {
          $cbDomain.SelectedItem = $Domain_Defaults[$tabName]
        }
      }
      # Per-tab formats
      try {
        $lf = $LoginFormat_ByTab[$tabName]
        $idx = $cbLoginFormat.Items.IndexOf($lf)
        if ($idx -ge 0) { $cbLoginFormat.SelectedIndex = $idx }
      }
      catch {}
      try {
        if ($DisplayNameFormat_ByTab.ContainsKey($tabName)) { $tbDN.Text = [string]$DisplayNameFormat_ByTab[$tabName] }
      }
      catch {}
      if ($OU_Defaults.ContainsKey($tabName)) {
        $tbOU_Edit.Text = $OU_Defaults[$tabName]
      }
      # Toggle login format UI for Student tab
      if ($tabName -eq 'Student') {
        $cbLoginFormat.Enabled = $false
        if ($lblStudentFmtInfo) { $lblStudentFmtInfo.Text = 'Student: login = NrAlbumu' }
        if ($toolTip) { $toolTip.SetToolTip($cbLoginFormat, 'Student: login = NrAlbumu') }
      }
      else {
        $cbLoginFormat.Enabled = $true
        if ($lblStudentFmtInfo) { $lblStudentFmtInfo.Text = '' }
      }
    }
    $script:IsInitializing = $false
  })
$grids = @{}
foreach ($name in $tabNames) {
  $tab = New-Object System.Windows.Forms.TabPage
  $tab.Text = $name

  $tabLayout = New-Object System.Windows.Forms.TableLayoutPanel
  $tabLayout.Dock = 'Fill'
  $tabLayout.ColumnCount = 1
  $tabLayout.RowCount = 2
  [void]$tabLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$tabLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  $tabLayout.Padding = [System.Windows.Forms.Padding]::new(0)
  $tabLayout.Margin = [System.Windows.Forms.Padding]::new(0)

  $grid = New-Object System.Windows.Forms.DataGridView
  $grid.Dock = 'Fill'
  $grid.Margin = [System.Windows.Forms.Padding]::new(0)
  $grid.BackgroundColor = [System.Drawing.SystemColors]::Window
  $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
  $grid.RowHeadersVisible = $false
  $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
  $grid.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
  $grid.AllowUserToAddRows = $true
  $grid.AllowUserToDeleteRows = $true
  $grid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
  $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
  $grid.MultiSelect = $false
  $grid.EnableHeadersVisualStyles = $false
  $grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.SystemColors]::ControlLight
  $grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(45, 52, 60)
  $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.SystemColors]::ControlLightLight

  switch ($name) {
    'Uczen' { $columns = @('Imie', 'Nazwisko', 'NazwaWyswietlana', 'Login', 'Miasto', 'Email', 'Haslo') }
    'Student' { $columns = @('Imie', 'Nazwisko', 'NrAlbumu', 'NazwaWyswietlana', 'Login', 'Miasto', 'Email', 'Haslo') }
    default { $columns = @('Imie', 'Nazwisko', 'NazwaWyswietlana', 'Login', 'Miasto', 'Email', 'Haslo') }
  }
  foreach ($col in $columns) { [void]$grid.Columns.Add($col, $col) }

  $buttonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttonsPanel.Dock = 'Fill'
  $buttonsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $buttonsPanel.AutoSize = $true
  $buttonsPanel.AutoSizeMode = 'GrowAndShrink'
  $buttonsPanel.WrapContents = $false
  $buttonsPanel.Margin = [System.Windows.Forms.Padding]::new(0, 8, 0, 0)
  $buttonsPanel.Padding = [System.Windows.Forms.Padding]::new(0)

  $pasteBtn = New-Object System.Windows.Forms.Button
  $pasteBtn.Text = 'Wklej dane ze schowka'
  & $styleButton $pasteBtn
  $pasteBtn.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
  $pasteBtn.Add_Click({
      try {
        $selectedTab = $tabs.SelectedTab.Text
        $targetGrid = $grids[$selectedTab]
        $clip = Get-Clipboard -Raw
        if ([string]::IsNullOrWhiteSpace($clip)) { Write-ToTextBox 'Schowek jest pusty.' 'Warning'; return }

        $linesText = $clip -split "`r?
"
        $targetGrid.SuspendLayout()
        foreach ($line in $linesText) {
          if ([string]::IsNullOrWhiteSpace($line)) { continue }
          $vals = $line -split "`t"
          $exp = switch ($selectedTab) { 'Uczen' { 2 } 'Student' { 3 } default { 2 } }
          if ($vals.Count -lt $exp) { Write-ToTextBox "Za malo kolumn dla $selectedTab" 'Warning'; continue }
          if ($vals.Count -gt $exp) { Write-ToTextBox "Zbyt wiele kolumn dla $selectedTab" 'Warning'; continue }
          $ri = $targetGrid.Rows.Add()
          $row = $targetGrid.Rows[$ri]
          if ($vals.Count -ge 1) { $row.Cells['Imie'].Value = ([string]$vals[0]).Trim() }
          if ($vals.Count -ge 2) { $row.Cells['Nazwisko'].Value = ([string]$vals[1]).Trim() }
          if ($selectedTab -eq 'Student' -and $vals.Count -ge 3) { $row.Cells['NrAlbumu'].Value = ([string]$vals[2]).Trim() }
        }

        $domain = if (-not $tbDomain_Edit.ReadOnly -and -not [string]::IsNullOrWhiteSpace($tbDomain_Edit.Text)) { $tbDomain_Edit.Text.Trim() } else { $Domain_Defaults[$selectedTab] }
        for ($r = 0; $r -lt $targetGrid.Rows.Count; $r++) {
          $rRow = $targetGrid.Rows[$r]
          if ($rRow.IsNewRow) { continue }
          $imi = [string]$rRow.Cells['Imie'].Value
          $naz = [string]$rRow.Cells['Nazwisko'].Value
          if ([string]::IsNullOrWhiteSpace($imi) -or [string]::IsNullOrWhiteSpace($naz)) { continue }
          $fmt = $LoginFormat_ByTab[$selectedTab]
          $login = Get-Login $imi $naz $fmt
          if ($selectedTab -eq 'Student') {
            $album = [string]$rRow.Cells['NrAlbumu'].Value
            if (-not [string]::IsNullOrWhiteSpace($album)) {
              $albumClean = $album.Trim()
              $rRow.Cells['Login'].Value = $albumClean
              $rRow.Cells['Email'].Value = ("{0}@{1}" -f $albumClean, $domain)
              $rRow.Cells['NazwaWyswietlana'].Value = "$imi $naz ($albumClean)"
            }
            else {
              $rRow.Cells['Login'].Value = $login
              $rRow.Cells['Email'].Value = "$login@$domain"
              $rRow.Cells['NazwaWyswietlana'].Value = "$imi $naz (Student)"
            }
          }
          else {
            $rRow.Cells['Login'].Value = $login
            $rRow.Cells['Email'].Value = "$login@$domain"
            $rRow.Cells['NazwaWyswietlana'].Value = "$imi $naz ($selectedTab)"
          }
          if (-not $rRow.Cells['Miasto'].Value) { $rRow.Cells['Miasto'].Value = $cbMiasto.SelectedItem }
          if (-not $rRow.Cells['Haslo'].Value) { $rRow.Cells['Haslo'].Value = New-RandomPassword }
        }

        Write-ToTextBox "Wklejono dane dla zakladki $selectedTab." 'Info'
        # Po wklejeniu automatycznie uruchom sprawdzanie (jak klikniecie "Sprawdz")
      }
      catch {
        Write-ToTextBox "Blad podczas wklejania danych: $_" 'Error'
      }
      finally {
        $targetGrid.ResumeLayout()
        try { $checkBtn.PerformClick() } catch {}
      }
    })

  # Przyciski: Sprawdź i Odśwież dane
  $checkBtn = New-Object System.Windows.Forms.Button
  $checkBtn.Text = 'Sprawdź'
  & $styleButton $checkBtn
  $checkBtn.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
  $checkBtn.Add_Click({
      $selectedTab = $tabs.SelectedTab.Text
      $grid = $grids[$selectedTab]
      if (-not $grid) { Write-ToTextBox "Brak gridu dla zakladki $selectedTab" 'Warning'; return }
      $logins = @{}
      for ($r = 0; $r -lt $grid.Rows.Count; $r++) {
        $row = $grid.Rows[$r]; if ($row.IsNewRow) { continue }
        $imi = [string]$row.Cells['Imie'].Value
        $naz = [string]$row.Cells['Nazwisko'].Value
        if ($selectedTab -eq 'Student') {
          $album = [string]$row.Cells['NrAlbumu'].Value
          if ([string]::IsNullOrWhiteSpace($album)) { Write-ToTextBox "Wiersz $($r): brak NrAlbumu" 'Warning' }
        }
        else {
          if ([string]::IsNullOrWhiteSpace($imi) -or [string]::IsNullOrWhiteSpace($naz)) { Write-ToTextBox "Wiersz $($r): brak imienia/nazwiska" 'Warning' }
        }
        $loginVal = [string]$row.Cells['Login'].Value
        if (-not [string]::IsNullOrWhiteSpace($loginVal)) {
          if ($logins.ContainsKey($loginVal)) { Write-ToTextBox "Duplikat loginu: $loginVal (wiersze $($logins[$loginVal]), $r)" 'Error' } else { $logins[$loginVal] = $r }
        }
      }
      Write-ToTextBox "Sprawdzanie zakończone." 'Info'
  
      # Auto-sugestia unikalnych loginów dla Student/Wykładowca
      for ($r2 = 0; $r2 -lt $grid.Rows.Count; $r2++) {
        $row2 = $grid.Rows[$r2]; if ($row2.IsNewRow) { continue }
        $role = $selectedTab
        $login2 = [string]$row2.Cells['Login'].Value
        if ([string]::IsNullOrWhiteSpace($login2)) { continue }
        $unique = $login2
        try {
          $existing = @(
            Get-ADUser -LDAPFilter "(sAMAccountName=$login2*)" -Properties sAMAccountName -ErrorAction Stop |
            Select-Object -ExpandProperty sAMAccountName
          )
          if ($existing -and ($existing -contains $login2)) {
            for ($i = 1; $i -lt 10000; $i++) {
              $candidate = "$login2$i"
              if ($existing -notcontains $candidate) { $unique = $candidate; break }
            }
          }
        }
        catch {
          Write-ToTextBox "Get-ADUser nieosiagalne podczas sprawdzania loginow: $_" 'Warning'
        }
        if ($unique -ne $login2) {
          $row2.Cells['Login'].Value = $unique
          # Uaktualnij e-mail zgodnie z nowym loginem i bieżącą domeną
          $domain = if (-not $tbDomain_Edit.ReadOnly -and -not [string]::IsNullOrWhiteSpace($tbDomain_Edit.Text)) { $tbDomain_Edit.Text.Trim() } else { $Domain_Defaults[$selectedTab] }
          if ($row2.Cells -and $row2.Cells["Email"]) { $row2.Cells['Email'].Value = ("{0}@{1}" -f $unique, $domain) }
          Write-ToTextBox "Zmieniono login na unikalny: $login2 -> $unique (wiersz $r2)" 'Info'
        }
      }

    })
  $refreshBtn = New-Object System.Windows.Forms.Button
  $refreshBtn.Text = 'Odśwież dane'
  & $styleButton $refreshBtn
  $refreshBtn.Add_Click({
      try {
        $selectedTab = $tabs.SelectedTab.Text
        $grid = $grids[$selectedTab]
        if (-not $grid) { Write-ToTextBox "Brak gridu dla zakladki $selectedTab" 'Warning'; return }

        $domain = if (-not $tbDomain_Edit.ReadOnly -and -not [string]::IsNullOrWhiteSpace($tbDomain_Edit.Text)) { $tbDomain_Edit.Text.Trim() } else { $Domain_Defaults[$selectedTab] }
        $fmt = $LoginFormat_ByTab[$selectedTab]

        for ($r = 0; $r -lt $grid.Rows.Count; $r++) {
          $row = $grid.Rows[$r]; if ($row.IsNewRow) { continue }
          $imi = [string]$row.Cells['Imie'].Value
          $naz = [string]$row.Cells['Nazwisko'].Value

          if ($selectedTab -eq 'Student') {
            $album = ([string]$row.Cells['NrAlbumu'].Value).Trim()
            if (-not [string]::IsNullOrWhiteSpace($album)) {
              $row.Cells['Login'].Value = $album
              $row.Cells['Email'].Value = ("{0}@{1}" -f $album, $Script:Settings.EmailDomain)
              $row.Cells['NazwaWyswietlana'].Value = "$imi $naz ($album)"
            }
          }
          else {
            if ([string]::IsNullOrWhiteSpace($imi) -or [string]::IsNullOrWhiteSpace($naz)) { continue }
            $login = Get-Login $imi $naz $fmt
            $row.Cells['Login'].Value = $login
            $row.Cells['Email'].Value = "$login@$domain"
            if (-not $row.Cells['NazwaWyswietlana'].Value) { $row.Cells['NazwaWyswietlana'].Value = "$imi $naz ($selectedTab)" }
          }

          if (-not $row.Cells['Miasto'].Value) { $row.Cells['Miasto'].Value = $cbMiasto.SelectedItem }
          if (-not $row.Cells['Haslo'].Value) { $row.Cells['Haslo'].Value = New-RandomPassword }
        }
        Write-ToTextBox "Odświeżono dane dla zakładki $selectedTab." 'Info'
      }
      catch {
        Write-ToTextBox "Błąd odświeżania: $_" 'Error'
      }
    })
  $buttonsPanel.Controls.Add($pasteBtn) | Out-Null

  # Usunięto osobny przycisk Ustawienia – konfiguracja przeniesiona do głównego okna
  # WhatIf per tab (default: OFF)
  $cbWhatIf = New-Object System.Windows.Forms.CheckBox
  $cbWhatIf.Text = 'WhatIf'
  $cbWhatIf.Checked = $false
  $buttonsPanel.Controls.Add($cbWhatIf) | Out-Null
  $WhatIf_ByTab[$name] = $cbWhatIf

  $btnCreate = New-Object System.Windows.Forms.Button
  $btnCreate.Text = 'Utwórz konta'
  & $styleButton $btnCreate
  $btnCreate.Add_Click({
      try {
        $selectedTab = $tabs.SelectedTab.Text
        $grid = $grids[$selectedTab]
        if (-not $grid) { Write-ToTextBox "Brak gridu dla zakładki $selectedTab" 'Warning'; return }

        foreach ($row in $grid.Rows) {
          if ($row.IsNewRow) { continue }
          $imi = [string]$row.Cells['Imie'].Value
          $naz = [string]$row.Cells['Nazwisko'].Value
          $login = [string]$row.Cells['Login'].Value
          $dnFmt = if ($DisplayNameFormat_ByTab.ContainsKey($selectedTab)) { [string]$DisplayNameFormat_ByTab[$selectedTab] } else { [string]$Script:Settings.DisplayNameFormat }
          $dn = $dnFmt.Replace('{Imie}', $imi).Replace('{Nazwisko}', $naz).Replace('{Rola}', $selectedTab)
          # CN/Name in AD must be unique in the container; keep DisplayName readable,
          # but make Name unique by appending trailing digits from login if present,
          # otherwise append the login itself.
          $cn = $dn
          if ($login -match '\d+$') {
            $cn = "$dn $($Matches[0])"
          }
          else {
            $cn = "$dn ($login)"
          }
          $email = if ($selectedTab -eq 'Student' -and $row.Cells['NrAlbumu'].Value) {
            "{0}@{1}" -f ([string]$row.Cells['NrAlbumu'].Value), $Script:Settings.EmailDomain
          }
          else {
            "{0}@{1}" -f $login, $Script:Settings.EmailDomain
          }
          if ([string]::IsNullOrWhiteSpace($login)) { Write-ToTextBox "Pomijam wiersz bez loginu" 'Warning'; continue }

          # dopilnuj unikalności jeszcze raz
          $loginU = Get-UniqueLogin -BaseLogin $login -Rola $selectedTab
          if ($loginU -ne $login) { $row.Cells['Login'].Value = $loginU; $login = $loginU }

          $PlainPassword = [string]$row.Cells['Haslo'].Value
          if ([string]::IsNullOrWhiteSpace($pwd)) { $PlainPassword = New-RandomPassword; $row.Cells['Haslo'].Value = $pwd }

          try {
            $params = @{
              Name              = $cn
              GivenName         = $imi
              Surname           = $naz
              SamAccountName    = $login
              UserPrincipalName = "$login@$($Script:Settings.EmailDomain)"
              DisplayName       = $dn
              EmailAddress      = $email
              Enabled           = $true
              Path              = if (-not [string]::IsNullOrWhiteSpace($tbOU_Edit.Text)) { $tbOU_Edit.Text } else { $OU_Defaults[$selectedTab] }
              AccountPassword   = (ConvertTo-SecureString $PlainPassword -AsPlainText -Force)
            }
            if ($WhatIf_ByTab.ContainsKey($selectedTab) -and $WhatIf_ByTab[$selectedTab].Checked) {
              Write-ToTextBox ("[WHATIF] New-ADUser " + ($params | Out-String)) 'Info'
            }
            else {
              New-ADUser @params
              Write-ToTextBox "Utworzono konto: $login ($dn)" 'Success'
            }
          }
          catch {
            Write-ToTextBox "Błąd tworzenia konta $($login): $_" 'Error'
          }
        }
      }
      catch {
        Write-ToTextBox "Błąd procesu tworzenia kont: $_" 'Error'
      }
    })
  $buttonsPanel.Controls.Add($btnCreate) | Out-Null

  $buttonsPanel.Controls.Add($checkBtn) | Out-Null
  $buttonsPanel.Controls.Add($refreshBtn) | Out-Null

  $tabLayout.Controls.Add($grid, 0, 0)
  $tabLayout.Controls.Add($buttonsPanel, 0, 1)
  $tab.Controls.Add($tabLayout)

  $grids[$name] = $grid
  [void]$tabs.TabPages.Add($tab)
}


$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Eksportuj dane do CSV'
& $styleButton $btnExport
$btnExport.Add_Click({
    try {
      $selectedTab = $tabs.SelectedTab.Text
      $grid = $grids[$selectedTab]
      if (-not $grid) { Write-ToTextBox "Brak gridu dla zakladki $selectedTab" 'Warning'; return }

      $dlg = New-Object System.Windows.Forms.SaveFileDialog
      $dlg.Filter = 'CSV file (*.csv)|*.csv|All files (*.*)|*.*'
      $dlg.Title = 'Zapisz dane do CSV'
      if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
      $path = $dlg.FileName

      $colNames = $grid.Columns | ForEach-Object { $_.Name }
      $linesToSave = New-Object System.Collections.Generic.List[string]
      $linesToSave.Add(($colNames -join ',')) | Out-Null

      for ($r = 0; $r -lt $grid.Rows.Count; $r++) {
        $row = $grid.Rows[$r]
        if ($row.IsNewRow) { continue }
        $vals = foreach ($cn in $colNames) {
          $val = $row.Cells[$cn].Value
          if ($null -eq $val) { $val = '' }
          '"' + $val.ToString().Replace('"', '""') + '"'
        }
        $linesToSave.Add(($vals -join ',')) | Out-Null
      }

      [System.IO.File]::WriteAllLines($path, $linesToSave, [System.Text.Encoding]::UTF8)
      Write-ToTextBox "Wyeksportowano dane do pliku: $path" 'Info'
    }
    catch {
      Write-ToTextBox "Blad eksportu CSV: $_" 'Error'
    }
  })

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'Wyczysc zakladke'
& $styleButton $btnClear
$btnClear.Add_Click({ $selectedTab = $tabs.SelectedTab.Text; $grids[$selectedTab].Rows.Clear(); Write-ToTextBox "Wyczyszczono dane z zakladki $selectedTab." 'Info' })

$btnTopMost = New-Object System.Windows.Forms.Button
$btnTopMost.Text = 'Zawsze na wierzchu'
& $styleButton $btnTopMost

$btnPomoc = New-Object System.Windows.Forms.Button
$btnPomoc.Text = 'Pomoc'
& $styleButton $btnPomoc

$btnZamknij = New-Object System.Windows.Forms.Button
$btnZamknij.Text = 'Zamknij'
& $styleButton $btnZamknij
$btnZamknij.Margin = [System.Windows.Forms.Padding]::new(8, 0, 0, 0)
$btnZamknij.Add_Click({ $form.Close() })

$btnPomoc.Add_Click({
    $msg = @"
Skladnia wklejania (TAB):
- Uczen/Pracownik/Wykladowca/Inne: Imie[TAB]Nazwisko
- Student: Imie[TAB]Nazwisko[TAB]NrAlbumu

Ustawienia (plik JSON):
$Script:SettingsPath
Plik tworzy się przy zapisie ustawień.
"@
    [Windows.Forms.MessageBox]::Show($msg, 'Pomoc - Kreator Kont AD') | Out-Null
  })

$actionPanel.Controls.Add($btnDodaj) | Out-Null
$actionPanel.Controls.Add($btnExport) | Out-Null
$actionPanel.Controls.Add($btnClear) | Out-Null
$actionPanel.Controls.Add($btnTopMost) | Out-Null
$actionPanel.Controls.Add($btnPomoc) | Out-Null
$actionPanel.Controls.Add($btnZamknij) | Out-Null

$form.TopMost = $false

$updateTopMostState = {
  $btnTopMost.Text = if ($form.TopMost) { 'Zawsze na wierzchu: wl.' } else { 'Zawsze na wierzchu: wyl.' }
}

$btnTopMost.Add_Click({
    $form.TopMost = -not $form.TopMost
    & $updateTopMostState
  })

& $updateTopMostState

$form.Add_Shown({
    $script:IsInitializing = $true
    $form.Activate()
    try {
      $cbMiasto.Items.Clear()
      [void]$cbMiasto.Items.AddRange($Cities)
      if ($cbMiasto.Items.Count -gt 0) { $cbMiasto.SelectedIndex = 0 }
    }
    catch {}
    try {
      $cbDomain.Items.Clear()
      [void]$cbDomain.Items.AddRange((Get-AvailableDomains))
      if ($tabs.SelectedTab) {
        $dnm = $Domain_Defaults[$tabs.SelectedTab.Text]
        if ($cbDomain.Items.Contains($dnm)) { $cbDomain.SelectedItem = $dnm }
      }
      if ($cbDomain.SelectedItem) { $Script:Settings.EmailDomain = [string]$cbDomain.SelectedItem }
    }
    catch {}
    if ($tabs.SelectedTab) {
      $tabName = $tabs.SelectedTab.Text
      if ($Domain_Defaults.ContainsKey($tabName)) { $tbDomain_Edit.Text = $Domain_Defaults[$tabName] }
      if ($OU_Defaults.ContainsKey($tabName)) { $tbOU_Edit.Text = $OU_Defaults[$tabName] }
      # Toggle login format UI for active tab (Student locked)
      if ($tabName -eq 'Student') {
        $cbLoginFormat.Enabled = $false
        if ($lblStudentFmtInfo) { $lblStudentFmtInfo.Text = 'Student: login = NrAlbumu' }
      }
      else {
        $cbLoginFormat.Enabled = $true
        if ($lblStudentFmtInfo) { $lblStudentFmtInfo.Text = '' }
      }
      # Init per-tab formats on first show
      try {
        $lf = $LoginFormat_ByTab[$tabName]
        $idx = $cbLoginFormat.Items.IndexOf($lf)
        if ($idx -ge 0) { $cbLoginFormat.SelectedIndex = $idx }
      }
      catch {}
      try { if ($DisplayNameFormat_ByTab.ContainsKey($tabName)) { $tbDN.Text = [string]$DisplayNameFormat_ByTab[$tabName] } } catch {}
    }
    & $updateTopMostState
    # Legacy config message suppressed; using .AD-BulkUserCreator.json only
    $script:IsInitializing = $false
  })

Get-ABCSettings
[void]$form.ShowDialog()

