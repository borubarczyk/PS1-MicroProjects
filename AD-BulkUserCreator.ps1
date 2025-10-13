<#
.SYNOPSIS
Bulk Active Directory user creation wizard with a modern WinForms UI.

.DESCRIPTION
This PowerShell script launches a Windows Forms application that helps you paste
and validate people data, generate logins/emails, pick a UPN suffix and a target
OU, and create accounts in Active Directory in bulk. The tool supports multiple
roles (tabs): Uczen (Pupil), Student, Pracownik (Employee), Wykladowca (Lecturer),
and Inne (Other). It persists per‑tab settings (login/display name formats, defaults)
to a JSON file in the script folder, validates common input mistakes, and surfaces
friendly error hints for typical AD creation problems.

.FEATURES
- Paste data from clipboard (TAB‑separated); auto‑check runs immediately.
- Per‑tab login/display name formats; Student login is fixed to NrAlbumu.
- UPN/email built from the domain selected in the GUI (Domena/UPN).
- Target OU selection with a tree browser; CN is made unique when needed.
- AD collision checks and friendly error messages in Polish (e.g., login/CN exists,
  password/complexity, permissions, invalid CN, UPN/domain connectivity).
- Name normalization (e.g., "IMIE" -> "Imie"); validation flags rows with digits or
  special characters in name/surname; rows color‑coded:
  • Red    = account exists in AD
  • Green  = login is available
  • Orange = invalid data (name/surname issues)
- Random password generator with complexity constraints.
- WhatIf mode per tab.
- System shell stock icons on key buttons.
- Export current grid to CSV.

.REQUIREMENTS
- Windows PowerShell 5.1 or PowerShell 7+ on Windows.
- RSAT ActiveDirectory module installed and importable.
- Permissions to create users in the selected OU.
- Windows desktop session (WinForms UI).

.USAGE
- Run the script: .\AD-BulkUserCreator.ps1
- Select a tab (role). Paste data from clipboard:
  • Uczen/Pracownik/Wykladowca/Inne:  Imie<TAB>Nazwisko
  • Student:                           Imie<TAB>Nazwisko<TAB>NrAlbumu
- Pick Domena/UPN and choose OU. Adjust formats if applicable and click
  "Sprawdz" to validate or "Utwórz konta" to create users.

.DATA PERSISTENCE
- Settings are saved to: $Script:SettingsPath (.AD-BulkUserCreator.json)
  and include per‑tab formats, OU/domain defaults, and cities list.

.NOTES
- Student display name is fixed to "Imie Nazwisko (NrAlbumu)" and login = NrAlbumu.
- UPN/email are generated from the GUI‑selected domain per tab.
- The tool colors invalid input rows and will not attempt to auto‑suggest logins
  for them until corrected.

.OUTPUTS
None. Displays a GUI and writes operational logs to the on‑screen log window.

.LINK
Active Directory PowerShell Module: https://learn.microsoft.com/powershell/module/activedirectory/
WinForms in PowerShell: https://learn.microsoft.com/powershell/scripting/samples/creating-windows-forms
SHGetStockIconInfo: https://learn.microsoft.com/windows/win32/api/shellapi/nf-shellapi-shgetstockiconinfo
#>

#region Init
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Interop & Icons
# P/Invoke: Shell stock icons (SHGetStockIconInfo)
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace ShellInterop {
    public static class StockIconNative
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHSTOCKICONINFO
        {
            public UInt32 cbSize;
            public IntPtr hIcon;
            public Int32 iSysImageIndex;
            public Int32 iIcon;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szPath;
        }

        [DllImport("Shell32.dll", SetLastError = false)]
        private static extern int SHGetStockIconInfo(int siid, uint uFlags, ref SHSTOCKICONINFO psii);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool DestroyIcon(IntPtr hIcon);

        private const uint SHGSI_ICON = 0x000000100;
        private const uint SHGSI_LARGEICON = 0x000000000;
        private const uint SHGSI_SMALLICON = 0x000000001;

        public static IntPtr GetHIcon(int siid, bool small)
        {
            SHSTOCKICONINFO sii = new SHSTOCKICONINFO();
            sii.cbSize = (UInt32)System.Runtime.InteropServices.Marshal.SizeOf(typeof(SHSTOCKICONINFO));
            uint flags = SHGSI_ICON | (small ? SHGSI_SMALLICON : SHGSI_LARGEICON);
            int hr = SHGetStockIconInfo(siid, flags, ref sii);
            if (hr != 0) return IntPtr.Zero;
            return sii.hIcon;
        }
    }
}
"@

# Helper: set shell stock icon on a Button (with safe fallback)
function Set-ButtonStockIcon {
  param(
    [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
    [Parameter(Mandatory)][int]$Siid,
    [switch]$Small
  )
  try {
    $h = [ShellInterop.StockIconNative]::GetHIcon($Siid, [bool]$Small)
    if ($h -ne [IntPtr]::Zero) {
      $ico = [System.Drawing.Icon]::FromHandle($h)
      try {
        $Button.Image = $ico.ToBitmap()
        $Button.ImageAlign = [System.Drawing.ContentAlignment]::MiddleLeft
        $Button.TextImageRelation = [System.Windows.Forms.TextImageRelation]::ImageBeforeText
      }
      finally {
        [ShellInterop.StockIconNative]::DestroyIcon($h) | Out-Null
        try { $ico.Dispose() } catch {}
      }
    }
    <# [Removed misplaced validation block]
      # (removed)
      # (removed)
        # (removed)
      }
      Write-ToTextBox ("Uwaga: błędne dane w wierszach: {0}" -f (($badLps | Sort-Object) -join ', ')) 'Warning'
      try { Show-InvalidDataDialog -Grid $Grid -Rows $bad } catch {}
    #>
  }
  catch {}
}
#endregion Interop & Icons

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

#region Logging
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

#region Configuration
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

# Typical operational limit for sAMAccountName in many environments (legacy 20 chars).
# Used for validation only; does not auto-truncate.
$Script:MaxLoginLength = 20



function Get-ABCSettings {
  if (Test-Path -LiteralPath $Script:SettingsPath) {
    try {
      $json = Get-Content -LiteralPath $Script:SettingsPath -Raw -ErrorAction Stop
      $data = $json | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
      try { [Windows.Forms.MessageBox]::Show('Plik ustawien jest uszkodzony i zostanie podmieniony na szablon.', 'Ustawienia', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning) | Out-Null } catch {}
      if (Write-DefaultConfig) {
        try {
          $json = Get-Content -LiteralPath $Script:SettingsPath -Raw -ErrorAction Stop
          $data = $json | ConvertFrom-Json -ErrorAction Stop
        }
        catch {}
      }
    }
    if ($null -ne $data) {
      if ($data.LoginFormat) { $Script:Settings.LoginFormat = [string]$data.LoginFormat }
      if ($data.EmailDomain) { $Script:Settings.EmailDomain = [string]$data.EmailDomain }
      if ($data.DisplayNameFormat) { $Script:Settings.DisplayNameFormat = [string]$data.DisplayNameFormat }
      if ($data.PSObject.Properties.Name -contains 'LoginFormatByTab' -and $data.LoginFormatByTab) {
        foreach ($k in $tabNames) { if ($data.LoginFormatByTab.PSObject.Properties.Name -contains $k) { $LoginFormat_ByTab[$k] = [string]$data.LoginFormatByTab.$k } }
      }
      elseif ($data.LoginFormat) {
        foreach ($k in $tabNames) { $LoginFormat_ByTab[$k] = [string]$data.LoginFormat }
      }
      if ($data.PSObject.Properties.Name -contains 'DisplayNameFormatByTab' -and $data.DisplayNameFormatByTab) {
        foreach ($k in $tabNames) { if ($data.DisplayNameFormatByTab.PSObject.Properties.Name -contains $k) { $DisplayNameFormat_ByTab[$k] = [string]$data.DisplayNameFormatByTab.$k } }
      }
      elseif ($data.DisplayNameFormat) {
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
    }
    else {
      Write-ToTextBox "Brak danych ustawien po wczytaniu, uzywam domyslnych." 'Warning'
    }
  }
  else {
    try { [Windows.Forms.MessageBox]::Show('Brak pliku ustawien – zostanie utworzony szablon.', 'Ustawienia', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information) | Out-Null } catch {}
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
  }
  catch {
    Write-ToTextBox ('Nie udalo sie zapisac domyslnego pliku ustawien: ' + $_.Exception.Message) 'Error'
    try { [Windows.Forms.MessageBox]::Show('Nie udalo sie zapisac domyslnego pliku ustawien: ' + $_.Exception.Message, 'Blad', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error) | Out-Null } catch {}
    return $false
  }
}

#region GUI Helpers
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

#region AD Helpers (Login Uniqueness)
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
    $existsExact = Get-ADUser -LDAPFilter "(sAMAccountName=$BaseLogin)" -ErrorAction Stop
  }
  catch {
    Write-ToTextBox "Get-ADUser nieosiągalne podczas sprawdzania loginów: $_" 'Warning'
    return $BaseLogin
  }
  if (-not $existsExact) { return $BaseLogin }
  for ($i = 1; $i -lt 10000; $i++) {
    $candidate = "$BaseLogin$i"
    try {
      $existsCand = Get-ADUser -LDAPFilter "(sAMAccountName=$candidate)" -ErrorAction Stop
    }
    catch {
      Write-ToTextBox "Get-ADUser nieosiągalne podczas sprawdzania loginów (kandydat): $_" 'Warning'
      return $BaseLogin
    }
    if (-not $existsCand) { return $candidate }
  }
  return $BaseLogin
}


# Proper-case converter for Polish names (e.g., "IMIE NAZWISKO" -> "Imię Nazwisko",
# handles multi-part and hyphenated names like "Anna-Maria").
#region Names & Validation
## Column name constants (single source of truth)
$ColLp               = 'Lp'
$ColImie             = 'Imię'
$ColNazwisko         = 'Nazwisko'
$ColNrAlbumu         = 'NrAlbumu'
$ColNazwaWyswietlana = 'NazwaWyswietlana'
$ColLogin            = 'Login'
$ColMiasto           = 'Miasto'
$ColEmail            = 'Email'
$ColHaslo            = 'Haslo'

 function Convert-ToPolishProperName {
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
  $pl = [System.Globalization.CultureInfo]::GetCultureInfo('pl-PL')
  $t = $Text.Trim()
  # normalize internal whitespace to single spaces
  $t = ($t -replace '\s+', ' ')
  $words = $t -split ' '
  $outWords = New-Object System.Collections.Generic.List[string]
  foreach ($w in $words) {
    if ([string]::IsNullOrWhiteSpace($w)) { continue }
    # Handle hyphenated parts separately so both sides are title-cased
    $parts = $w -split '-'
    $outParts = New-Object System.Collections.Generic.List[string]
    foreach ($p in $parts) {
      if ([string]::IsNullOrWhiteSpace($p)) { continue }
      $lower = $p.ToLower($pl)
      $tc = $pl.TextInfo.ToTitleCase($lower)
      [void]$outParts.Add($tc)
    }
    [void]$outWords.Add([string]::Join('-', $outParts))
  }
  return [string]::Join(' ', $outWords)
}
#endregion AD Helpers (Login Uniqueness)
#endregion GUI Helpers
#endregion Configuration
#endregion Logging

# Walidacja pól imię/nazwisko – zwraca $true jeśli poprawne, inaczej $false i powód
function Test-NameValid {
  param(
    [string]$Text,
    [string]$Field = 'Pole'
  )
  $trim = if ($null -ne $Text) { $Text.Trim() } else { '' }
  if ([string]::IsNullOrWhiteSpace($trim)) { return , $false, ("Brak {0}" -f $Field) }
  # Niedozwolone: cyfry, kropki, znaki specjalne (poza myślnikiem i spacją)
  if ($trim -match "\d") { return , $false, ("{0} zawiera cyfry" -f $Field) }
  if ($trim -match "[^\p{L} \-]") { return , $false, ("{0} zawiera niedozwolone znaki" -f $Field) }
  # Zbyt wiele spacji (dwie i więcej obok siebie)
  if ($trim -match "\s{2,}") { return , $false, ("{0} zawiera wielokrotne spacje" -f $Field) }
  return , $true, ''
}

# Oznacza w siatce wiersze z błędnym Imię/Nazwisko oraz wypisuje listę w logu
function Invoke-RowNameValidation {
  param([System.Windows.Forms.DataGridView]$Grid)
  try {
    if (-not $Grid) { return }
    $colorInvalid = [System.Drawing.Color]::Orange
    $bad = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $Grid.Rows.Count; $i++) {
      $r = $Grid.Rows[$i]; if ($r.IsNewRow) { continue }
      $imi = [string]$r.Cells['Imię'].Value
      $naz = [string]$r.Cells[$ColNazwisko].Value
      $resI = Test-NameValid -Text $imi -Field 'Imię'
      $okI  = [bool]$resI[0]
      $resN = Test-NameValid -Text $naz -Field 'Nazwisko'
      $okN  = [bool]$resN[0]
      $login = [string]$r.Cells[$ColLogin].Value
      $tooLong = $false
      try {
        if (-not [string]::IsNullOrWhiteSpace($login)) {
          $max = if ($Script:MaxLoginLength) { [int]$Script:MaxLoginLength } else { 20 }
          if ($login.Length -gt $max) { $tooLong = $true }
        }
      } catch {}
      if (-not $okI -or -not $okN -or $tooLong) {
        $r.DefaultCellStyle.BackColor = $colorInvalid
        [void]$bad.Add($i)
      }
    }
    if ($bad.Count -gt 0) { Write-ToTextBox ("Uwaga: błędne dane w wierszach: {0}" -f (Get-BadRowList -Grid $Grid -Rows $bad)) 'Warning' }
    if ($bad.Count -gt 0) { try { Show-InvalidDataDialog -Grid $Grid -Rows $bad } catch {} }
  }
  catch {}
}

# Uaktualnia kolumnę "Lp" na podstawie kolejności wierszy (1..N)
function Update-GridIndexColumn {
  param([System.Windows.Forms.DataGridView]$Grid)
  try {
    if (-not $Grid) { return }
    $n = 1
    for ($i = 0; $i -lt $Grid.Rows.Count; $i++) {
      $r = $Grid.Rows[$i]; if ($r.IsNewRow) { continue }
      try { $r.Cells['Lp'].Value = $n } catch {}
      $n++
    }
  }
  catch {}
}

# Zwraca listę numerów Lp (lub indeksów) dla podanych wierszy w siatce
function Get-BadRowList {
  param(
    [Parameter(Mandatory)] [System.Windows.Forms.DataGridView] $Grid,
    [Parameter(Mandatory)] [System.Collections.IEnumerable] $Rows
  )
  $list = New-Object System.Collections.Generic.List[object]
  foreach ($idx in ($Rows | Sort-Object)) {
    try {
      $lpv = $Grid.Rows[$idx].Cells[$ColLp].Value
      if ($null -ne $lpv -and $lpv.ToString().Trim() -ne '') { [void]$list.Add($lpv) } else { [void]$list.Add($idx) }
    } catch { [void]$list.Add($idx) }
  }
  return (($list | Sort-Object) -join ', ')
}
function Show-InvalidDataDialog {
  param(
    [Parameter(Mandatory)] [System.Windows.Forms.DataGridView] $Grid,
    [Parameter(Mandatory)] [System.Collections.IEnumerable] $Rows
  )
  try {
    $details = New-Object System.Collections.Generic.List[string]
    $max = if ($Script:MaxLoginLength) { [int]$Script:MaxLoginLength } else { 20 }
    foreach($i in $Rows){
      if ($i -isnot [int]) { continue }
      $r = $Grid.Rows[$i]; if (-not $r -or $r.IsNewRow) { continue }
      $reasons = @()
      $imi = [string]$r.Cells[$ColImie].Value
      $naz = [string]$r.Cells['Nazwisko'].Value
      $resI = Test-NameValid -Text $imi -Field 'Imię'
      if (-not [bool]$resI[0]) { $reasons += [string]$resI[1] }
      $resN = Test-NameValid -Text $naz -Field 'Nazwisko'
      if (-not [bool]$resN[0]) { $reasons += [string]$resN[1] }
      $login = [string]$r.Cells[$ColLogin].Value
      if (-not [string]::IsNullOrWhiteSpace($login) -and $login.Length -gt $max) {
        $reasons += ("Login przekracza {0} znaków" -f $max)
      }
      $dispRow = $i + 1
      try {
        $lpv2 = $r.Cells[$ColLp].Value
        if ($null -ne $lpv2 -and $lpv2.ToString().Trim() -ne '') { $dispRow = $lpv2 }
      } catch {}
      if ($reasons.Count -gt 0) { [void]$details.Add(("Wiersz {0}: {1}" -f $dispRow, ($reasons -join '; '))) }
    }
    if ($details.Count -gt 0) {
      [System.Windows.Forms.MessageBox]::Show(($details -join [Environment]::NewLine), 'Błędne dane', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
  } catch {}
}
#endregion Names & Validation

# Sprawdza typowe konflikty przed utworzeniem konta w AD
#region AD Pre-checks & Errors
function Test-ADPreCreateConflicts {
  param(
    [Parameter(Mandatory)][string]$Sam,
    [Parameter(Mandatory)][string]$CN,
    [Parameter(Mandatory)][string]$Path
  )
  $res = [ordered]@{ LoginExists = $false; NameExists = $false; Messages = New-Object System.Collections.Generic.List[string] }
  try {
    $u = Get-ADUser -LDAPFilter "(sAMAccountName=$Sam)" -ErrorAction Stop
    if ($u) { $res.LoginExists = $true; $null = $res.Messages.Add("Login zajęty w AD: $Sam") }
  }
  catch {}
  if (-not [string]::IsNullOrWhiteSpace($Path)) {
    try {
      $o = Get-ADObject -LDAPFilter "(name=$CN)" -SearchBase $Path -SearchScope OneLevel -ErrorAction Stop
      if ($o) { $res.NameExists = $true; $null = $res.Messages.Add("Nazwa (CN/Name) już istnieje w OU: '$CN'") }
    }
    catch {}
  }
  return [pscustomobject]$res
}

# Mapowanie bledow AD na przyjazne komunikaty PL
function Get-FriendlyADError {
  param([Parameter(Mandatory)]$ErrorRecord)
  try {
    $msg = [string]$ErrorRecord.Exception.Message
  }
  catch { $msg = [string]$ErrorRecord }
  $m = $msg.ToLowerInvariant()
  if ($m -match 'already.*exist|already in use|in use') { return 'Obiekt już istnieje (login/UPN/CN). Zmień login lub nazwę.' }
  if ($m -match 'access is denied|insufficient access rights') { return 'Brak uprawnień do tworzenia w wybranym OU.' }
  if ($m -match 'server is unwilling to process the request') { return 'Serwer odmówił żądania. Sprawdź politykę haseł, wymagane atrybuty i poprawność UPN.' }
  if ($m -match 'constraint violation|password.*requirement|complexit') { return 'Hasło nie spełnia wymagań domeny. Zmień hasło lub zasady.' }
  if ($m -match 'invalid dn syntax|object name is invalid|bad name') { return 'Nieprawidłowa nazwa (CN) lub znaki niedozwolone.' }
  if ($m -match 'no such object|cannot find an object with identity') { return 'Nieprawidłowa ścieżka OU lub brak obiektu docelowego.' }
  if ($m -match 'naming violation') { return 'Naruszenie zasad nazewnictwa AD (CN/Name).' }
  if ($m -match 'the specified domain either does not exist|unable to contact the global catalog') { return 'Wybrana domena/UPN niedostępna lub błąd łączności z AD.' }
  return ("{0}" -f $msg)
}
#endregion AD Pre-checks & Errors


#region Password Generation
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
#endregion Password Generation

#region AD Module & OU Picker
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
#endregion AD Module & OU Picker



#endregion

# Helper: validate/check current tab data (extracted from Sprawdź handler)
#region Data Check & Auto-suggest
function Invoke-ActiveTabCheck {
  try {
    $selectedTab = $tabs.SelectedTab.Text
    $grid = $grids[$selectedTab]
    if (-not $grid) { Write-ToTextBox "Brak gridu dla zakladki $selectedTab" 'Warning'; return }
    # Kolory statusu
    $colorExists = [System.Drawing.Color]::LightCoral    # czerwony: konto istnieje
    $colorFree = [System.Drawing.Color]::PaleGreen     # zielony: konto wolne
    # Wyczyść wcześniejsze kolory
    foreach ($r in 0..($grid.Rows.Count - 1)) {
      $rowClr = $grid.Rows[$r]
      if ($rowClr -and -not $rowClr.IsNewRow) { $rowClr.DefaultCellStyle.BackColor = [System.Drawing.Color]::Empty }
    }

    # Normalizuj Imię/Nazwisko do formy wlaściwej (np. "IMIE" -> "Imię")
    for ($rFix = 0; $rFix -lt $grid.Rows.Count; $rFix++) {
      $rowFix = $grid.Rows[$rFix]; if ($rowFix.IsNewRow) { continue }
      if ($rowFix.Cells['Imię']) {
        $rawImie = [string]$rowFix.Cells['Imię'].Value
        if (-not [string]::IsNullOrWhiteSpace($rawImie)) {
          $normImie = Convert-ToPolishProperName $rawImie
          $rowFix.Cells['Imię'].Value = $normImie
        }
      }
      if ($rowFix.Cells['Nazwisko']) {
        $rawNaz = [string]$rowFix.Cells['Nazwisko'].Value
        if (-not [string]::IsNullOrWhiteSpace($rawNaz)) {
          $normNaz = Convert-ToPolishProperName $rawNaz
          $rowFix.Cells['Nazwisko'].Value = $normNaz
        }
      }
    }

    # Student: nie sprawdzamy unikalnosci ani nie modyfikujemy loginow.
    # Tylko informujemy przy sprawdzaniu, ze konto juz istnieje (jesli jest w AD).
    if ($selectedTab -eq 'Student') {
      for ($r = 0; $r -lt $grid.Rows.Count; $r++) {
        $row = $grid.Rows[$r]; if ($row.IsNewRow) { continue }
        $imiS = [string]$row.Cells['Imię'].Value
        $nazS = [string]$row.Cells['Nazwisko'].Value
        $album = [string]$row.Cells['NrAlbumu'].Value
        if ([string]::IsNullOrWhiteSpace($album)) { Write-ToTextBox "Wiersz ${r}: brak NrAlbumu" 'Warning'; continue }
        $sam = $album.Trim()
        # Jesli dane w wierszu sa niepoprawne, nie koloruj na zielono/czerwono (zostanie nadpisane przez walidacje na pomaranczowo)
        $isInvalid = $false
        try {
          $resI = Test-NameValid -Text $imiS -Field 'Imię'
          $resN = Test-NameValid -Text $nazS -Field 'Nazwisko'
          $isInvalid = (-not [bool]$resI[0]) -or (-not [bool]$resN[0])
          if (-not [string]::IsNullOrWhiteSpace($sam)) {
            $max = if ($Script:MaxLoginLength) { [int]$Script:MaxLoginLength } else { 20 }
            if ($sam.Length -gt $max) { $isInvalid = $true }
          }
        } catch {}
        try {
          $user = Get-ADUser -LDAPFilter "(sAMAccountName=$sam)" -Properties sAMAccountName -ErrorAction Stop
          if (-not $isInvalid) {
            if ($user) {
              Write-ToTextBox "Konto studenta juz istnieje w AD: $sam (wiersz $r)" 'Info'
              $row.DefaultCellStyle.BackColor = $colorExists
            }
            else {
              $row.DefaultCellStyle.BackColor = $colorFree
            }
          }
          # Aktualizuj nazwę wyświetlaną wg znormalizowanych Imię/Nazwisko
          if ($row.Cells['NazwaWyswietlana']) {
            if (-not [string]::IsNullOrWhiteSpace($sam)) {
              $row.Cells['NazwaWyswietlana'].Value = ("{0} {1} ({2})" -f $imiS, $nazS, $sam)
            }
            else {
              $row.Cells['NazwaWyswietlana'].Value = ("{0} {1}" -f $imiS, $nazS)
            }
          }
        }
        catch {
          Write-ToTextBox "Get-ADUser nieosiagalne podczas sprawdzania kont studenta: $_" 'Warning'
        }
      }
      try { Invoke-RowNameValidation -Grid $grid } catch {}
      Write-ToTextBox "Sprawdzanie zakończone." 'Info'
      return
    }

    # Dla pozostalych zakladek: walidacja podstawowa + duplikaty loginow
    $logins = @{}
    for ($r = 0; $r -lt $grid.Rows.Count; $r++) {
      $row = $grid.Rows[$r]; if ($row.IsNewRow) { continue }
      $imi = [string]$row.Cells['Imię'].Value
      $naz = [string]$row.Cells['Nazwisko'].Value
      if ([string]::IsNullOrWhiteSpace($imi) -or [string]::IsNullOrWhiteSpace($naz)) { Write-ToTextBox "Wiersz $($r): brak imienia/nazwiska" 'Warning' }
      $loginVal = [string]$row.Cells['Login'].Value
      if (-not [string]::IsNullOrWhiteSpace($loginVal)) {
        if ($logins.ContainsKey($loginVal)) { Write-ToTextBox "Duplikat loginu: $loginVal (wiersze $($logins[$loginVal]), $r)" 'Error' } else { $logins[$loginVal] = $r }
      }
    }
    try { Invoke-RowNameValidation -Grid $grid } catch {}
    Write-ToTextBox "Sprawdzanie zakończone." 'Info'

    # Auto-sugestia unikalnych loginow (nie dotyczy Student)
    for ($r2 = 0; $r2 -lt $grid.Rows.Count; $r2++) {
      $row2 = $grid.Rows[$r2]; if ($row2.IsNewRow) { continue }
      $login2 = [string]$row2.Cells['Login'].Value
      if ([string]::IsNullOrWhiteSpace($login2)) { continue }
      $base = $login2
      $unique = $base
      try {
        $existsExact = Get-ADUser -LDAPFilter "(sAMAccountName=$base)" -ErrorAction Stop
        if ($existsExact) {
          for ($i = 1; $i -lt 10000; $i++) {
            $candidate = "$base$i"
            # unikaj kolizji z innymi wierszami w siatce
            $inGrid = $false
            for ($q = 0; $q -lt $grid.Rows.Count; $q++) {
              if ($q -eq $r2) { continue }
              $rOther = $grid.Rows[$q]; if ($rOther.IsNewRow) { continue }
              $otherLogin = [string]$rOther.Cells['Login'].Value
              if ($otherLogin -and ($otherLogin -ieq $candidate)) { $inGrid = $true; break }
            }
            if ($inGrid) { continue }
            $existsCand = Get-ADUser -LDAPFilter "(sAMAccountName=$candidate)" -ErrorAction Stop
            if (-not $existsCand) { $unique = $candidate; break }
          }
        }
      }
      catch {
        Write-ToTextBox "Get-ADUser nieosiagalne podczas sprawdzania loginow: $_" 'Warning'
      }
      if ($unique -ne $base) {
        $row2.Cells['Login'].Value = $unique
        # Uaktualnij e-mail zgodnie z nowym loginem i bieżącą domeną z GUI
        $domain = if ($cbDomain.SelectedItem) { [string]$cbDomain.SelectedItem } else { $Domain_Defaults[$selectedTab] }
        if ($row2.Cells -and $row2.Cells['Email']) { $row2.Cells['Email'].Value = ("{0}@{1}" -f $unique, $domain) }
        Write-ToTextBox "Zmieniono login na unikalny: $base -> $unique (wiersz $r2)" 'Info'
      }
    }

    # Koloruj wiersze wg istnienia konta w AD (dokładne dopasowanie sAMAccountName)
    for ($r3 = 0; $r3 -lt $grid.Rows.Count; $r3++) {
      $row3 = $grid.Rows[$r3]; if ($row3.IsNewRow) { continue }
      $login3 = [string]$row3.Cells['Login'].Value
      if ([string]::IsNullOrWhiteSpace($login3)) { continue }
      # Pomiń kolorowanie na zielono/czerwono, jeżeli wiersz jest niepoprawny (Imię/Nazwisko lub długość loginu)
      $skipColor = $false
      try {
        $im3 = [string]$row3.Cells[$ColImie].Value
        $na3 = [string]$row3.Cells[$ColNazwisko].Value
        $ri = Test-NameValid -Text $im3 -Field 'Imię'
        $rn = Test-NameValid -Text $na3 -Field 'Nazwisko'
        if (-not [bool]$ri[0] -or -not [bool]$rn[0]) { $skipColor = $true }
        $max3 = if ($Script:MaxLoginLength) { [int]$Script:MaxLoginLength } else { 20 }
        if ($login3.Length -gt $max3) { $skipColor = $true }
      } catch {}
      if ($skipColor) { continue }
      try {
        $u = Get-ADUser -LDAPFilter "(sAMAccountName=$login3)" -Properties sAMAccountName -ErrorAction Stop
        if ($u) { $row3.DefaultCellStyle.BackColor = $colorExists } else { $row3.DefaultCellStyle.BackColor = $colorFree }
      }
      catch {
        # Brak koloru, jesli AD niedostepne
      }
    }
  }
  catch {
    Write-ToTextBox "Blad sprawdzania: $_" 'Error'
  }
}
#endregion Data Check & Auto-suggest

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
[void]$layoutRoot.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
$form.Controls.Add($layoutRoot)

$lblUsage = New-Object System.Windows.Forms.Label
$lblUsage.Text = 'Wklej Imię i Nazwisko (Student dodatkowo NrAlbumu). Wybierz domenę i OU dla aktywnej zakładki, następnie dodaj konta.'
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
  $Button.Padding = [System.Windows.Forms.Padding]::new(8, 4, 8, 4)
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
$cbMiasto.Width = 180
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

## usunięto informacyjny wiersz z wybraną domeną – zbędny, bo wybór jest w ComboBox

$tbOU_Edit = New-Object System.Windows.Forms.TextBox
$tbOU_Edit.ReadOnly = $true
$tbOU_Edit.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$tbOU_Edit.BackColor = $form.BackColor

# Przycisk wyboru OU przeniesiony pod listę Domena/UPN,
# a ścieżka OU poniżej – bez dodatkowej etykiety
$tbOU_Edit.Dock = 'Fill'
$tbOU_Edit.Text = 'OU zdefiniowane dla aktywnej zakladki'
$tbOU_Edit.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 0)

$btnChooseOU = New-Object System.Windows.Forms.Button
$btnChooseOU.Text = 'Wybierz OU'
& $styleButton $btnChooseOU
try { Set-ButtonStockIcon -Button $btnChooseOU -Siid 4 -Small } catch {}
# Zmniejsz czcionke o 1pt i zmniejsz padding, aby zmiescic ikone i tekst
try {
  $newSize = [Math]::Max(6, [int]$btnChooseOU.Font.Size - 1)
  $btnChooseOU.Font = New-Object System.Drawing.Font($btnChooseOU.Font.FontFamily, $newSize, $btnChooseOU.Font.Style)
}
catch {}
$btnChooseOU.Padding = [System.Windows.Forms.Padding]::new(6, 2, 6, 2)
# Dopasuj wysokość do comboboxów (Domena/Miasto)
$btnChooseOU.AutoSize = $false
$btnChooseOU.Height = $cbDomain.Height
$btnChooseOU.Width = [Math]::Max($btnChooseOU.PreferredSize.Width + 10, 140)
$btnChooseOU.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 4)

# Etykieta i panel w jednej linii (wiersz 2): [Ścieżka OU:] [Wybierz OU] [ścieżka]
$lblOU = New-Object System.Windows.Forms.Label
$lblOU.Text = 'Ścieżka OU:'
$lblOU.AutoSize = $true
$lblOU.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 0)
$configPanel.Controls.Add($lblOU, 0, 2)

$ouInline = New-Object System.Windows.Forms.TableLayoutPanel
$ouInline.ColumnCount = 2
$ouInline.RowCount = 1
$ouInline.AutoSize = $true
$ouInline.AutoSizeMode = 'GrowAndShrink'
$ouInline.Dock = 'Fill'
$ouInline.ColumnStyles.Clear()
[void]$ouInline.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$ouInline.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$tbOU_Edit.Margin = [System.Windows.Forms.Padding]::new(8, 0, 0, 0)
$tbOU_Edit.Dock = 'Fill'
$ouInline.Controls.Add($btnChooseOU, 0, 0)
$ouInline.Controls.Add($tbOU_Edit, 1, 0)
$configPanel.Controls.Add($ouInline, 1, 2)
$configPanel.SetColumnSpan($ouInline, 2)

# Umieszczenie: przycisk pod Domena/UPN (wiersz 2), ścieżka OU w kolejnym wierszu (3)
$configPanel.Controls.Add($btnChooseOU, 1, 2)
$configPanel.SetColumnSpan($btnChooseOU, 2)
$configPanel.Controls.Add($tbOU_Edit, 0, 3)
$configPanel.SetColumnSpan($tbOU_Edit, 3)
# Przenies do ukladu w jednej linii (usun ze starych komorek i dodaj do panelu)
$configPanel.Controls.Remove($btnChooseOU)
$configPanel.Controls.Remove($tbOU_Edit)
[void]$ouInline.Controls.Add($btnChooseOU)
[void]$ouInline.Controls.Add($tbOU_Edit)
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
$lblFormat.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 6)
$configPanel.Controls.Add($lblFormat, 0, 4)

$cbLoginFormat = New-Object System.Windows.Forms.ComboBox
$cbLoginFormat.DropDownStyle = 'DropDownList'
$null = $cbLoginFormat.Items.AddRange(@('i.nazwisko', 'inazwisko', 'imie.nazwisko', 'nazwisko.imie'))
$cbLoginFormat.SelectedIndex = 0
$cbLoginFormat.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 6)
$cbLoginFormat.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$cbLoginFormat.Width = 300
$configPanel.Controls.Add($cbLoginFormat, 1, 4)
$configPanel.SetColumnSpan($cbLoginFormat, 2)
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
$lblDN.Margin = [System.Windows.Forms.Padding]::new(0, 0, 12, 6)
$configPanel.Controls.Add($lblDN, 0, 5)

$tbDN = New-Object System.Windows.Forms.TextBox
$tbDN.Text = [string]$Script:Settings.DisplayNameFormat
$tbDN.Dock = 'Fill'
$tbDN.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 6)
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
# Info label nie jest dodawany do panelu, aby ComboBox mógł zajmować pełną szerokość (kolumny 1-2)
# $configPanel.Controls.Add($lblStudentFmtInfo, 2, 4)
$actionPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionPanel.AutoSize = $true
$actionPanel.AutoSizeMode = 'GrowAndShrink'
$actionPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
$actionPanel.WrapContents = $false
$actionPanel.Margin = [System.Windows.Forms.Padding]::new(0, 0, 0, 6)
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
      # Toggle login/display name UI for Student tab
      if ($tabName -eq 'Student') {
        $cbLoginFormat.Enabled = $false
        $tbDN.ReadOnly = $true
        $tbDN.Text = 'Imię Nazwisko (NrAlbumu)'
        if ($lblStudentFmtInfo) { $lblStudentFmtInfo.Text = 'Student: login = NrAlbumu; Nazwa wyświetlana = Imię Nazwisko (NrAlbumu)' }
        if ($toolTip) { $toolTip.SetToolTip($cbLoginFormat, 'Student: login = NrAlbumu') }
      }
      else {
        $cbLoginFormat.Enabled = $true
        $tbDN.ReadOnly = $false
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
  $grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells
  $grid.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
  $grid.AllowUserToResizeColumns = $true
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
    'Uczen' { $columns = @('Lp', 'Imię', 'Nazwisko', 'NazwaWyswietlana', 'Login', 'Miasto', 'Email', 'Haslo') }
    'Student' { $columns = @('Lp', 'Imię', 'Nazwisko', 'NrAlbumu', 'NazwaWyswietlana', 'Login', 'Miasto', 'Email', 'Haslo') }
    default { $columns = @('Lp', 'Imię', 'Nazwisko', 'NazwaWyswietlana', 'Login', 'Miasto', 'Email', 'Haslo') }
  }
  foreach ($col in $columns) { [void]$grid.Columns.Add($col, $col) }
  try { $grid.Columns['Lp'].ReadOnly = $true } catch {}
  if ($name -eq 'Student') {
    try { $grid.Columns['NazwaWyswietlana'].ReadOnly = $true } catch {}
  }

  # Wiersz przycisków: lewa grupa (Wklej/Usuń/WhatIf) + prawa grupa (Utwórz/Sprawdź/Odśwież)
  $buttonsRow = New-Object System.Windows.Forms.TableLayoutPanel
  $buttonsRow.Dock = 'Fill'
  $buttonsRow.AutoSize = $true
  $buttonsRow.AutoSizeMode = 'GrowAndShrink'
  $buttonsRow.Margin = [System.Windows.Forms.Padding]::new(0, 4, 0, 0)
  $buttonsRow.ColumnCount = 2
  $buttonsRow.RowCount = 1
  $buttonsRow.ColumnStyles.Clear()
  [void]$buttonsRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
  [void]$buttonsRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
  $buttonsLeft = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttonsLeft.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $buttonsLeft.AutoSize = $true
  $buttonsLeft.AutoSizeMode = 'GrowAndShrink'
  $buttonsLeft.WrapContents = $false
  $buttonsLeft.Dock = 'Fill'
  $buttonsLeft.Padding = [System.Windows.Forms.Padding]::new(0)
  $buttonsRight = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttonsRight.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
  $buttonsRight.AutoSize = $true
  $buttonsRight.AutoSizeMode = 'GrowAndShrink'
  $buttonsRight.WrapContents = $false
  $buttonsRight.Padding = [System.Windows.Forms.Padding]::new(0)
  $buttonsRow.Controls.Add($buttonsLeft, 0, 0)
  $buttonsRow.Controls.Add($buttonsRight, 1, 0)

  $pasteBtn = New-Object System.Windows.Forms.Button
  $pasteBtn.Text = 'Wklej dane ze schowka'
  & $styleButton $pasteBtn
  $pasteBtn.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
  try { Set-ButtonStockIcon -Button $pasteBtn -Siid 55 -Small } catch {}
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
          if ($vals.Count -ge 1) { $row.Cells['Imię'].Value = ([string]$vals[0]).Trim() }
          if ($vals.Count -ge 2) { $row.Cells['Nazwisko'].Value = ([string]$vals[1]).Trim() }
          if ($selectedTab -eq 'Student' -and $vals.Count -ge 3) { $row.Cells['NrAlbumu'].Value = ([string]$vals[2]).Trim() }
        }

        $domain = if ($cbDomain.SelectedItem) { [string]$cbDomain.SelectedItem } else { $Domain_Defaults[$selectedTab] }
        for ($r = 0; $r -lt $targetGrid.Rows.Count; $r++) {
          $rRow = $targetGrid.Rows[$r]
          if ($rRow.IsNewRow) { continue }
          $imi = [string]$rRow.Cells['Imię'].Value
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
              $rRow.Cells['NazwaWyswietlana'].Value = "$imi $naz"
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
        try { Update-GridIndexColumn -Grid $targetGrid } catch {}
        try { Invoke-ActiveTabCheck } catch {}
      }
    })

  # Przyciski: Sprawdź i Odśwież dane
  $checkBtn = New-Object System.Windows.Forms.Button
  $checkBtn.Text = 'Sprawdź'
  & $styleButton $checkBtn
  $checkBtn.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
  $checkBtn.Add_Click({
      Invoke-ActiveTabCheck
      return
    })
  $refreshBtn = New-Object System.Windows.Forms.Button
  $refreshBtn.Text = 'Odśwież dane'
  & $styleButton $refreshBtn
  $refreshBtn.Add_Click({
      try {
        $selectedTab = $tabs.SelectedTab.Text
        $grid = $grids[$selectedTab]
        if (-not $grid) { Write-ToTextBox "Brak gridu dla zakladki $selectedTab" 'Warning'; return }

        $domain = if ($cbDomain.SelectedItem) { [string]$cbDomain.SelectedItem } else { $Domain_Defaults[$selectedTab] }
        $fmt = $LoginFormat_ByTab[$selectedTab]

        for ($r = 0; $r -lt $grid.Rows.Count; $r++) {
          $row = $grid.Rows[$r]; if ($row.IsNewRow) { continue }
          $imi = [string]$row.Cells['Imię'].Value
          $naz = [string]$row.Cells['Nazwisko'].Value

          if ($selectedTab -eq 'Student') {
            $album = ([string]$row.Cells['NrAlbumu'].Value).Trim()
            if (-not [string]::IsNullOrWhiteSpace($album)) {
              $row.Cells['Login'].Value = $album
              $row.Cells['Email'].Value = ("{0}@{1}" -f $album, $domain)
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
  $buttonsLeft.Controls.Add($pasteBtn) | Out-Null

  # Usuń zaznaczony wiersz (pojedynczy)
  $btnDeleteRow = New-Object System.Windows.Forms.Button
  $btnDeleteRow.Text = 'Usuń wiersz'
  & $styleButton $btnDeleteRow
  $btnDeleteRow.Margin = [System.Windows.Forms.Padding]::new(0, 0, 8, 0)
  try { Set-ButtonStockIcon -Button $btnDeleteRow -Siid 84 -Small } catch {}
  $btnDeleteRow.Add_Click({
      try {
        $selectedTab = $tabs.SelectedTab.Text
        $grid = $grids[$selectedTab]
        if (-not $grid) { Write-ToTextBox "Brak gridu dla zakładki $selectedTab" 'Warning'; return }
        $row = $grid.CurrentRow
        if (-not $row -or $row.IsNewRow) { Write-ToTextBox 'Nie wybrano prawidłowego wiersza do usunięcia.' 'Warning'; return }
        $idx = $row.Index
        $grid.Rows.RemoveAt($idx)
        Write-ToTextBox "Usunięto wiersz $idx w zakładce $selectedTab." 'Info'
      }
      catch {
        Write-ToTextBox "Błąd usuwania wiersza: $_" 'Error'
      }
    })
  $buttonsLeft.Controls.Add($btnDeleteRow) | Out-Null

  # Usunięto osobny przycisk Ustawienia – konfiguracja przeniesiona do głównego okna
  # WhatIf per tab (default: OFF)
  $cbWhatIf = New-Object System.Windows.Forms.CheckBox
  $cbWhatIf.Text = 'WhatIf'
  $cbWhatIf.Checked = $false
  $buttonsLeft.Controls.Add($cbWhatIf) | Out-Null
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
          $imi = [string]$row.Cells['Imię'].Value
          $naz = [string]$row.Cells['Nazwisko'].Value
          $login = [string]$row.Cells['Login'].Value
          # DisplayName: Student ma sztywno "Imie Nazwisko (NrAlbumu)", bez roli i bez edycji formatu
          if ($selectedTab -eq 'Student') {
            $albumForDn = ([string]$row.Cells['NrAlbumu'].Value).Trim()
            if (-not [string]::IsNullOrWhiteSpace($albumForDn)) { $dn = "$imi $naz ($albumForDn)" } else { $dn = "$imi $naz" }
          }
          else {
            $dnFmt = if ($DisplayNameFormat_ByTab.ContainsKey($selectedTab)) { [string]$DisplayNameFormat_ByTab[$selectedTab] } else { [string]$Script:Settings.DisplayNameFormat }
            $dn = $dnFmt.Replace('{Imie}', $imi).Replace('{Nazwisko}', $naz).Replace('{Rola}', $selectedTab)
          }
          # CN/Name: dla Student nie dopisujemy nr albumu drugi raz
          # (CN = DisplayName, czyli "Imie Nazwisko (NrAlbumu)").
          # Dla pozostalych ról utrzymujemy dotychczasowe reguly unikalnosci.
          if ($selectedTab -eq 'Student') {
            $cn = $dn
          }
          else {
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
          }
          $domainForAccount = if ($cbDomain.SelectedItem) { [string]$cbDomain.SelectedItem } else { $Domain_Defaults[$selectedTab] }
          $email = if ($selectedTab -eq 'Student' -and $row.Cells['NrAlbumu'].Value) {
            "{0}@{1}" -f ([string]$row.Cells['NrAlbumu'].Value), $domainForAccount
          }
          else {
            "{0}@{1}" -f $login, $domainForAccount
          }
          if ([string]::IsNullOrWhiteSpace($login)) { Write-ToTextBox "Pomijam wiersz bez loginu" 'Warning'; continue }

          # dopilnuj unikalności jeszcze raz
          $loginU = Get-UniqueLogin -BaseLogin $login -Rola $selectedTab
          if ($loginU -ne $login) { $row.Cells['Login'].Value = $loginU; $login = $loginU }

          $PlainPassword = [string]$row.Cells['Haslo'].Value
          if ([string]::IsNullOrWhiteSpace($PlainPassword)) { $PlainPassword = New-RandomPassword; $row.Cells['Haslo'].Value = $PlainPassword }

          try {
            $targetPath = if (-not [string]::IsNullOrWhiteSpace($tbOU_Edit.Text)) { $tbOU_Edit.Text } else { $OU_Defaults[$selectedTab] }
            # Pre-check typowych konfliktów przed New-ADUser
            $conf = Test-ADPreCreateConflicts -Sam $login -CN $cn -Path $targetPath
            if ($conf.LoginExists) { Write-ToTextBox "Login już istnieje w AD: $login — zmień login (wiersz: $($row.Index))" 'Error'; continue }
            if ($conf.NameExists) { Write-ToTextBox "Nazwa (CN/Name) w OU już istnieje: '$cn' — zmień format nazwy lub wybierz inne OU (wiersz: $($row.Index))" 'Error'; continue }
            $params = @{
              Name              = $cn
              GivenName         = $imi
              Surname           = $naz
              SamAccountName    = $login
              UserPrincipalName = "$login@$domainForAccount"
              DisplayName       = $dn
              EmailAddress      = $email
              Enabled           = $true
              Path              = $targetPath
              AccountPassword   = (ConvertTo-SecureString $PlainPassword -AsPlainText -Force)
            }
            if ($WhatIf_ByTab.ContainsKey($selectedTab) -and $WhatIf_ByTab[$selectedTab].Checked) {
              Write-ToTextBox ("[WHATIF] New-ADUser " + ($params | Out-String)) 'Info'
            }
            else {
              try {
                New-ADUser @params
                Write-ToTextBox "Utworzono konto: $login ($dn)" 'Success'
              }
              catch {
                $hint = Get-FriendlyADError -ErrorRecord $_
                Write-ToTextBox ("Błąd tworzenia konta {0}: {1}" -f $login, $hint) 'Error'
                continue
              }
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
  # Prawa grupa (wyrównana do prawej kolumny)
  # Prawa grupa (wyrównana do prawej kolumny)
  try { Set-ButtonStockIcon -Button $btnCreate -Siid 96 -Small } catch {}
  $buttonsRight.Controls.Add($btnCreate) | Out-Null

  try { Set-ButtonStockIcon -Button $checkBtn -Siid 22 -Small } catch {}
  $buttonsRight.Controls.Add($checkBtn) | Out-Null
  try { Set-ButtonStockIcon -Button $refreshBtn -Siid 79 -Small } catch {}
  $buttonsRight.Controls.Add($refreshBtn) | Out-Null

  $tabLayout.Controls.Add($grid, 0, 0)
  $tabLayout.Controls.Add($buttonsRow, 0, 1)
  $tab.Controls.Add($tabLayout)

  $grids[$name] = $grid
  [void]$tabs.TabPages.Add($tab)
}


$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Eksportuj dane do CSV'
& $styleButton $btnExport
try { Set-ButtonStockIcon -Button $btnExport -Siid 1 -Small } catch {}
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
try { Set-ButtonStockIcon -Button $btnClear -Siid 31 -Small } catch {}
$btnClear.Add_Click({ $selectedTab = $tabs.SelectedTab.Text; $grids[$selectedTab].Rows.Clear(); Write-ToTextBox "Wyczyszczono dane z zakladki $selectedTab." 'Info' })

$btnTopMost = New-Object System.Windows.Forms.Button
$btnTopMost.Text = 'Zawsze na wierzchu'
& $styleButton $btnTopMost
try { Set-ButtonStockIcon -Button $btnTopMost -Siid 106 -Small } catch {}

$btnPomoc = New-Object System.Windows.Forms.Button
$btnPomoc.Text = 'Pomoc'
& $styleButton $btnPomoc
try { Set-ButtonStockIcon -Button $btnPomoc -Siid 23 -Small } catch {}

$btnZamknij = New-Object System.Windows.Forms.Button
$btnZamknij.Text = 'Zamknij'
& $styleButton $btnZamknij
try { Set-ButtonStockIcon -Button $btnZamknij -Siid 80 -Small } catch {}
$btnZamknij.Margin = [System.Windows.Forms.Padding]::new(8, 0, 0, 0)
$btnZamknij.Add_Click({ $form.Close() })

$btnPomoc.Add_Click({
    $msg = @"
Opis
- Skrypt służy do masowego tworzenia kont AD z danych wklejanych do siatki (zakładki: Uczeń, Student, Pracownik, Wykładowca, Inne).
- Po wklejeniu danych automatycznie uruchamia się Sprawdź: normalizuje imiona i nazwiska do formy „Imię Nazwisko”, sprawdza dostępność loginów w AD i koloruje wiersze (czerwony = konto istnieje, zielony = wolne). Dla Studentów login = NrAlbumu i tylko raportowane jest istnienie konta (bez modyfikacji loginu).
- Domena dla UPN/e‑mail jest brana z listy „Domena/UPN” w GUI (per zakładka). OU docelowe wybierasz przyciskiem „Wybierz OU”.
- Tworzenie kont ustawia: CN/Name, DisplayName, UPN, e‑mail, hasło i Enabled. Dla Student DisplayName/CN = „Imię Nazwisko (NrAlbumu)”. Jeśli login jest zajęty, skrypt proponuje unikalny login zaczynając od 1 (np. j.kowalski1).
- Dostępne akcje: Wklej dane, Usuń wiersz, Sprawdź, Odśwież dane, Utwórz konta, Eksportuj CSV, Wyczyszcz zakładkę, Zawsze na wierzchu, Pomoc, Zamknij.

Skladnia wklejania (TAB)
- Uczen/Pracownik/Wykladowca/Inne: Imie[TAB]Nazwisko
- Student: Imie[TAB]Nazwisko[TAB]NrAlbumu

Ustawienia i wymagania
- Ustawienia zapisywane są w pliku: $Script:SettingsPath (tworzy się przy zapisie/zmianie ustawień).
- Wymagany moduł RSAT: ActiveDirectory oraz uprawnienia do tworzenia kont w wybranym OU.
"@
    [Windows.Forms.MessageBox]::Show($msg, 'Pomoc – Kreator Kont AD') | Out-Null
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
      # info o wybranej domenie pod comboboxem usunięte – brak potrzeby aktualizacji pola
      if ($OU_Defaults.ContainsKey($tabName)) { $tbOU_Edit.Text = $OU_Defaults[$tabName] }
      # Toggle login/display name UI for active tab (Student locked)
      if ($tabName -eq 'Student') {
        $cbLoginFormat.Enabled = $false
        $tbDN.ReadOnly = $true
        $tbDN.Text = 'Imię Nazwisko (NrAlbumu)'
        if ($lblStudentFmtInfo) { $lblStudentFmtInfo.Text = 'Student: login = NrAlbumu; Nazwa wyświetlana = Imię Nazwisko (NrAlbumu)' }
      }
      else {
        $cbLoginFormat.Enabled = $true
        $tbDN.ReadOnly = $false
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


