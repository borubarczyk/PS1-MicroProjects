<#
.SYNOPSIS
    Domain Ops (AD-ManagerDiamond) - centrum administracji domeną: zdalne zarządzanie komputerami,
    użytkownicy i komputery w Active Directory. Interfejs WPF.

.DESCRIPTION
    Program jest podzielony na przestrzenie robocze (przełącznik na górze okna):
      - Zarządzanie zdalne  - operacje na zaznaczonych komputerach przez PowerShell Remoting (WinRM):
                              diagnostyka, sesje i profile użytkowników, usługi, procesy, dyski, zdarzenia,
                              oprogramowanie i aktualizacje, bezpieczeństwo, udziały, polecenia, instalacje,
      - Użytkownicy AD      - konta użytkowników: szczegóły, hasła i blokady, stan konta, grupy, atrybuty,
                              raporty (zablokowane, nieaktywne, wygasające hasła...) i źródło blokad,
      - Komputery AD        - konta komputerów: informacje, kanał zaufania, LAPS, klucze BitLocker,
                              zmiana nazwy, grupy i raporty (nieaktywne, systemy, bez LAPS...).
    Każda przestrzeń ma listę obiektów docelowych (komputery albo użytkownicy) z zaznaczaniem, a moduły
    pogrupowane w kategorie. Wyniki trafiają do tabel z filtrem, sortowaniem, podglądem wiersza,
    kopiowaniem i eksportem CSV. Operacje wykonują się w tle i równolegle - okno nie zawiesza się.

    Rozbudowa: każda przestrzeń to Register-Workspace, każdy moduł to Register-Module. Własne moduły
    można dodawać bez zmiany tego pliku: pliki *.ps1 z folderu AD-ManagerDiamond.Modules (obok skryptu)
    są wczytywane przy starcie. Przykład w README.

.NOTES
    Wymagania:
      - Windows PowerShell 5.1 (WPF, tryb STA - skrypt sam uruchomi się ponownie w odpowiednim trybie),
      - RSAT: moduł ActiveDirectory (przestrzenie AD i wczytywanie komputerów), opcjonalnie moduł LAPS,
      - WinRM na komputerach docelowych i uprawnienia administratora lokalnego.
    Ustawienia: %APPDATA%\AD-ManagerDiamond\settings.json
    Dziennik:   %LOCALAPPDATA%\AD-ManagerDiamond\Logs\DomainOps_RRRRMMDD.log
    Pliki robocze na komputerach: %SystemRoot%\Temp\DomainOps
    Plik musi pozostać zapisany jako UTF-8 z BOM (polskie znaki w Windows PowerShell 5.1).

.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\AD-ManagerDiamond.ps1
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Start: platforma, wersja PowerShell i tryb STA
if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
    throw 'Domain Ops działa wyłącznie w systemie Windows.'
}

# Interfejs jest pisany pod Windows PowerShell 5.1 (WPF w trybie STA). Z PowerShell 7 przełączamy się na powershell.exe,
# chyba że zmienna środowiskowa DOMAINOPS_ALLOW_CORE=1 pozwala zostać w PowerShell 7.
$isCore = ($PSVersionTable.PSEdition -eq 'Core') -and ($env:DOMAINOPS_ALLOW_CORE -ne '1')
$isSta = [System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA
if ($isCore -or -not $isSta) {
    if (-not $PSCommandPath) { throw 'Uruchom skrypt jako plik: powershell.exe -STA -File .\AD-ManagerDiamond.ps1' }
    $exe = if ($isCore) { Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe' } else { (Get-Process -Id $PID).Path }
    Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath)) | Out-Null
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
Add-Type -AssemblyName System.Windows.Forms   # tylko okno wyboru folderu

# Moduł ActiveDirectory nie musi tworzyć dysku AD: (szybszy import, mniej błędów przy braku DC)
$env:ADPS_LoadDefaultDrive = '0'
#endregion

#region Konfiguracja i stan
$script:AppVersion = '4.0'

$script:App = @{
    Name       = 'Domain Ops'
    DataDir    = Join-Path $env:APPDATA 'AD-ManagerDiamond'
    LogDir     = Join-Path $env:LOCALAPPDATA 'AD-ManagerDiamond\Logs'
    ModulesDir = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'AD-ManagerDiamond.Modules' } else { '' }
}
$script:App.SettingsFile = Join-Path $script:App.DataDir 'settings.json'
$script:App.LogFile = Join-Path $script:App.LogDir ('DomainOps_{0:yyyyMMdd}.log' -f (Get-Date))

# Ustawienia zapamiętywane między uruchomieniami
$script:Settings = [ordered]@{
    SearchBase       = ''
    NameFilter       = ''
    OnlyEnabled      = $true
    UserSearchBase   = ''
    UserFilter       = 0
    DomainController = ''
    ThrottleLimit    = 16
    TimeoutSec       = 20
    InactiveDays     = 90
    LastWorkspace    = 'Remote'
    LastModules      = @{}
    WindowWidth      = 1560
    WindowHeight     = 940
    WindowMaximized  = $false
    LogVisible       = $false
    LogHeight        = 190
    DetailVisible    = $true
}

# Poświadczenia bieżącej sesji (nie są zapisywane na dysku)
$script:State = @{
    Credential = $null
    UseCurrent = $true
}

# Interfejs: przestrzenie robocze, definicje i konteksty modułów, kontrolki okna głównego
$script:UI = @{
    Window          = $null
    Workspaces      = [ordered]@{}
    ModuleDefs      = New-Object System.Collections.ArrayList
    Modules         = @{}
    ActiveWorkspace = $null
    ActiveModule    = $null
    LogItems        = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    LogUnread       = 0
    Controls        = @{}
}

# Silnik operacji w tle
$script:Engine = @{
    Pool       = $null
    Operations = New-Object System.Collections.ArrayList
    Timer      = $null
    NextId     = 1
    InTick     = $false
    Deferred   = New-Object System.Collections.ArrayList
}

$script:LogContext = $null
$script:Clipboard = @{ Secret = $null; Timer = $null }

# Kolumny dodawane przez PowerShell Remoting, których nie pokazujemy w tabelach
$script:HiddenProperties = @('PSComputerName', 'RunspaceId', 'PSShowComputerName', 'PSSourceJobInstanceId')
#endregion

#region Narzędzia ogólne
function Write-Log {
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO',
        [AllowEmptyString()][AllowNull()][string]$Module = $script:LogContext
    )
    $now = Get-Date
    try {
        $items = $script:UI.LogItems
        if ($items.Count -ge 3000) { for ($i = 0; $i -lt 500; $i++) { $items.RemoveAt(0) } }
        $entry = [pscustomobject]@{
            Time    = $now.ToString('HH:mm:ss')
            Level   = $Level
            Module  = [string]$Module
            Message = $Message
        }
        $items.Add($entry)
        $list = $script:UI.Controls['logList']
        if ($list -and $list.IsVisible) { $list.ScrollIntoView($entry) }
        if (-not $script:Settings.LogVisible -and ($Level -eq 'WARN' -or $Level -eq 'ERROR')) {
            $script:UI.LogUnread++
            Update-LogBadge
        }
    }
    catch { }
    try {
        if (-not (Test-Path -LiteralPath $script:App.LogDir)) { New-Item -ItemType Directory -Path $script:App.LogDir -Force | Out-Null }
        $prefix = if ($Module) { "[$Module] " } else { '' }
        $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}{3}' -f $now, $Level, $prefix, $Message
        [System.IO.File]::AppendAllText($script:App.LogFile, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    }
    catch { }
}

function Update-LogBadge {
    $badge = $script:UI.Controls['logBadge']
    if (-not $badge) { return }
    if ($script:UI.LogUnread -gt 0) {
        $script:UI.Controls['logBadgeText'].Text = [string][Math]::Min(99, $script:UI.LogUnread)
        $badge.Visibility = 'Visible'
    }
    else { $badge.Visibility = 'Collapsed' }
}

function Get-EffectiveCredential {
    if (-not $script:State.UseCurrent -and $script:State.Credential) { return $script:State.Credential }
    return $null
}

function Get-AdSplat {
    # Parametry -Server/-Credential dla poleceń AD wykonywanych w wątku okna
    $p = @{}
    if ($script:Settings.DomainController) { $p.Server = [string]$script:Settings.DomainController }
    $cred = Get-EffectiveCredential
    if ($cred) { $p.Credential = $cred }
    return $p
}

function Import-AdModule {
    if (Get-Module -Name ActiveDirectory) { return }
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw 'Brak modułu ActiveDirectory (RSAT). Zainstaluj «RSAT: Active Directory Domain Services» i spróbuj ponownie.'
    }
    Import-Module ActiveDirectory -ErrorAction Stop -Verbose:$false | Out-Null
}

function Get-ObjectValue {
    param($InputObject, [string]$Name)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Data.DataRowView]) { $InputObject = $InputObject.Row }
    if ($InputObject -is [System.Data.DataRow]) {
        if (-not $InputObject.Table.Columns.Contains($Name)) { return $null }
        $v = $InputObject[$Name]
        if ($v -is [System.DBNull]) { return $null }
        return $v
    }
    if ($InputObject -is [System.Collections.IDictionary]) { return $InputObject[$Name] }
    $p = $InputObject.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function ConvertTo-CellValue {
    # Zamienia wartość z wyników na postać do tabeli: liczby zostają liczbami (sortowanie),
    # daty -> tekst ISO (sortuje się poprawnie), bool -> Tak/Nie, kolekcje -> tekst
    param($Value)
    if ($null -eq $Value) { return [System.DBNull]::Value }
    if ($Value -is [System.Management.Automation.PSObject]) { $Value = $Value.PSObject.BaseObject }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool]) { if ($Value) { return 'Tak' } else { return 'Nie' } }
    if ($Value -is [datetime]) {
        if ($Value.Year -lt 1700) { return [System.DBNull]::Value }
        return $Value.ToString('yyyy-MM-dd HH:mm:ss')
    }
    if ($Value -is [enum] -or $Value -is [timespan] -or $Value -is [guid] -or $Value -is [char]) { return [string]$Value }
    if ($Value -is [System.ValueType]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable]) {
        return (@($Value | ForEach-Object { [string]$_ }) -join ', ')
    }
    return [string]$Value
}

function ConvertTo-LikeLiteral {
    # Ucieka znaki specjalne wyrażenia LIKE w DataView.RowFilter
    param([string]$Text)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        if ('*%[]'.IndexOf($ch) -ge 0) { [void]$sb.Append('[').Append($ch).Append(']') }
        elseif ($ch -eq "'") { [void]$sb.Append("''") }
        else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function ConvertTo-XmlText([string]$Text) {
    return [System.Security.SecurityElement]::Escape([string]$Text)
}

function Format-LogObject {
    param($InputObject)
    if ($null -eq $InputObject) { return '' }
    if ($InputObject -is [System.Management.Automation.PSObject]) {
        $base = $InputObject.PSObject.BaseObject
        if ($base -is [string] -or $base -is [System.ValueType]) { return [string]$base }
    }
    if ($InputObject -is [string] -or $InputObject -is [System.ValueType]) { return [string]$InputObject }
    $parts = foreach ($p in $InputObject.PSObject.Properties) {
        if ($script:HiddenProperties -contains $p.Name -or $p.Name.StartsWith('__')) { continue }
        $v = ConvertTo-CellValue $p.Value
        if ($v -is [System.DBNull] -or [string]$v -eq '') { continue }
        '{0}: {1}' -f $p.Name, $v
    }
    return (@($parts) -join '; ')
}

function Format-Bytes {
    param([double]$Bytes)
    if ($Bytes -ge 1TB) { return '{0:N2} TB' -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return '{0:N2} GB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N1} MB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N0} KB' -f ($Bytes / 1KB) }
    return '{0:N0} B' -f $Bytes
}

function Split-ListText {
    # "a, b; c" -> @('a','b','c')
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    return @($Text -split '[,;\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Get-ParentDN([string]$DistinguishedName) {
    return ($DistinguishedName -replace '^(?:\\.|[^,])+,', '')
}

function Get-RdnValue([string]$DistinguishedName) {
    $first = [regex]::Match($DistinguishedName, '^(?:\\.|[^,])+').Value
    return (($first -replace '^[^=]+=', '') -replace '\\(.)', '$1')
}

function Get-SafeFileName([string]$Text) {
    return ($Text -replace '[\\/:*?"<>|\s]+', '_').Trim('_')
}

function Test-NetBiosName {
    # Zwraca opis problemu albo pusty tekst, gdy nazwa jest poprawna
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'Brak nowej nazwy' }
    if ($Name.Length -gt 15) { return 'Maksymalnie 15 znaków' }
    if ($Name -notmatch '^[A-Za-z0-9-]+$') { return 'Dozwolone są tylko litery, cyfry i myślnik' }
    if ($Name -match '^\d+$') { return 'Nazwa nie może składać się z samych cyfr' }
    if ($Name.StartsWith('-') -or $Name.EndsWith('-')) { return 'Nazwa nie może zaczynać się ani kończyć myślnikiem' }
    return ''
}

function New-RandomPassword {
    # Hasło spełniające wymagania złożoności (bez znaków łatwych do pomylenia)
    param([int]$Length = 16)
    $sets = @('ABCDEFGHJKLMNPQRSTUVWXYZ', 'abcdefghijkmnopqrstuvwxyz', '23456789', '!@#$%*-_=+?')
    $all = -join $sets
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $bytes = New-Object byte[] 4
        $pick = {
            param([string]$From)
            $rng.GetBytes($bytes)
            $From[[int]([BitConverter]::ToUInt32($bytes, 0) % [uint32]$From.Length)]
        }
        $chars = New-Object System.Collections.Generic.List[char]
        foreach ($set in $sets) { $chars.Add((& $pick $set)) }
        while ($chars.Count -lt [Math]::Max(8, $Length)) { $chars.Add((& $pick $all)) }
        # Tasowanie Fisher-Yates
        for ($i = $chars.Count - 1; $i -gt 0; $i--) {
            $rng.GetBytes($bytes)
            $j = [int]([BitConverter]::ToUInt32($bytes, 0) % [uint32]($i + 1))
            $tmp = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $tmp
        }
        return (-join $chars)
    }
    finally { $rng.Dispose() }
}

function Set-ClipboardText {
    # Schowek bywa chwilowo zajęty przez inny proces - kilka prób
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    for ($i = 0; $i -lt 6; $i++) {
        try {
            if ($Text) { [System.Windows.Clipboard]::SetText($Text) } else { [System.Windows.Clipboard]::Clear() }
            return
        }
        catch {
            if ($i -eq 5) { throw }
            Start-Sleep -Milliseconds 40
        }
    }
}

function Set-ClipboardSecret {
    # Kopiuje poufny tekst do schowka i czyści go po upływie czasu (o ile nadal tam jest)
    param([string]$Text, [int]$Seconds = 60)
    if ([string]::IsNullOrEmpty($Text)) { return }
    Set-ClipboardText $Text
    $script:Clipboard.Secret = $Text
    $timer = $script:Clipboard.Timer
    if ($timer) {
        $timer.Stop()
        $timer.Interval = [TimeSpan]::FromSeconds([Math]::Max(5, $Seconds))
        $timer.Start()
    }
}

function Clear-ClipboardSecret {
    if ($script:Clipboard.Timer) { $script:Clipboard.Timer.Stop() }
    if (-not $script:Clipboard.Secret) { return }
    try {
        if ([System.Windows.Clipboard]::ContainsText() -and [System.Windows.Clipboard]::GetText() -eq $script:Clipboard.Secret) {
            [System.Windows.Clipboard]::Clear()
            Write-Log 'Wyczyszczono poufną wartość ze schowka.' -Module ''
        }
    }
    catch { }
    $script:Clipboard.Secret = $null
}

function ConvertTo-Hashtable {
    # Obiekt z ConvertFrom-Json (PSCustomObject) -> hashtabla
    param($InputObject)
    $h = @{}
    if ($null -eq $InputObject) { return $h }
    if ($InputObject -is [System.Collections.IDictionary]) {
        foreach ($k in $InputObject.Keys) { $h[[string]$k] = $InputObject[$k] }
        return $h
    }
    foreach ($p in $InputObject.PSObject.Properties) { $h[$p.Name] = $p.Value }
    return $h
}

function Import-Settings {
    try {
        if (-not (Test-Path -LiteralPath $script:App.SettingsFile)) { return }
        $json = Get-Content -LiteralPath $script:App.SettingsFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($key in @($script:Settings.Keys)) {
            $p = $json.PSObject.Properties[$key]
            if ($p -and $null -ne $p.Value) { $script:Settings[$key] = $p.Value }
        }
    }
    catch { }
    # Wartości spoza zakresu (np. ręcznie edytowany plik) sprowadzamy do bezpiecznych granic
    $s = $script:Settings
    try { $s.ThrottleLimit = [Math]::Min(64, [Math]::Max(1, [int]$s.ThrottleLimit)) } catch { $s.ThrottleLimit = 16 }
    try { $s.TimeoutSec = [Math]::Min(300, [Math]::Max(5, [int]$s.TimeoutSec)) } catch { $s.TimeoutSec = 20 }
    try { $s.InactiveDays = [Math]::Min(3650, [Math]::Max(1, [int]$s.InactiveDays)) } catch { $s.InactiveDays = 90 }
    try { $s.UserFilter = [Math]::Min(3, [Math]::Max(0, [int]$s.UserFilter)) } catch { $s.UserFilter = 0 }
    try { $s.LogHeight = [Math]::Min(600, [Math]::Max(90, [int]$s.LogHeight)) } catch { $s.LogHeight = 190 }
    try { $s.WindowWidth = [Math]::Max(1100, [int]$s.WindowWidth); $s.WindowHeight = [Math]::Max(680, [int]$s.WindowHeight) } catch { }
    foreach ($b in 'OnlyEnabled', 'WindowMaximized', 'LogVisible', 'DetailVisible') { $s[$b] = [bool]$s[$b] }
    foreach ($t in 'SearchBase', 'NameFilter', 'UserSearchBase', 'DomainController', 'LastWorkspace') { $s[$t] = [string]$s[$t] }
    $s.LastModules = ConvertTo-Hashtable $s.LastModules
}

function Export-Settings {
    try {
        if (-not (Test-Path -LiteralPath $script:App.DataDir)) { New-Item -ItemType Directory -Path $script:App.DataDir -Force | Out-Null }
        $script:Settings | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $script:App.SettingsFile -Encoding UTF8
    }
    catch { }
}
#endregion

#region Motyw (ciemny, spójny z ServerReview / NPS Event Viewer)
$script:ThemeXaml = @'
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <SolidColorBrush x:Key="BgBrush" Color="#0F1318"/>
  <SolidColorBrush x:Key="PanelBrush" Color="#12171D"/>
  <SolidColorBrush x:Key="PanelBorderBrush" Color="#1E242E"/>
  <SolidColorBrush x:Key="CardBrush" Color="#161B22"/>
  <SolidColorBrush x:Key="CardBorderBrush" Color="#242B36"/>
  <SolidColorBrush x:Key="FieldBrush" Color="#1B212A"/>
  <SolidColorBrush x:Key="FieldBorderBrush" Color="#2A323F"/>
  <SolidColorBrush x:Key="HoverBrush" Color="#1A2029"/>
  <SolidColorBrush x:Key="SelectedBrush" Color="#1A2640"/>
  <SolidColorBrush x:Key="SelectedBorderBrush" Color="#2F4478"/>
  <SolidColorBrush x:Key="TextBrush" Color="#E4E8EF"/>
  <SolidColorBrush x:Key="MutedBrush" Color="#8791A5"/>
  <SolidColorBrush x:Key="FaintBrush" Color="#5E6779"/>
  <SolidColorBrush x:Key="AccentBrush" Color="#3E6FE0"/>
  <SolidColorBrush x:Key="FocusBrush" Color="#4C7DF0"/>
  <SolidColorBrush x:Key="OkBrush" Color="#5EE3AE"/>
  <SolidColorBrush x:Key="WarnBrush" Color="#FFC46B"/>
  <SolidColorBrush x:Key="CritBrush" Color="#FF7A86"/>
  <SolidColorBrush x:Key="InfoBrush" Color="#8CC0FF"/>

  <Style x:Key="Card" TargetType="Border">
    <Setter Property="Background" Value="#161B22"/>
    <Setter Property="BorderBrush" Value="#242B36"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="CornerRadius" Value="10"/>
    <Setter Property="Padding" Value="16,14"/>
  </Style>

  <Style x:Key="Glyph" TargetType="TextBlock">
    <Setter Property="FontFamily" Value="Segoe Fluent Icons, Segoe MDL2 Assets"/>
    <Setter Property="FontSize" Value="14"/>
    <Setter Property="VerticalAlignment" Value="Center"/>
  </Style>

  <Style x:Key="ScrollThumb" TargetType="Thumb">
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Thumb">
          <Border CornerRadius="4" Background="#39414F"/>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style TargetType="ScrollBar">
    <Setter Property="Width" Value="10"/>
    <Setter Property="MinWidth" Value="10"/>
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ScrollBar">
          <Grid Background="Transparent">
            <Track x:Name="PART_Track" IsDirectionReversed="True">
              <Track.Thumb><Thumb Style="{StaticResource ScrollThumb}" Margin="2"/></Track.Thumb>
            </Track>
          </Grid>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="Orientation" Value="Horizontal">
        <Setter Property="Width" Value="Auto"/>
        <Setter Property="MinWidth" Value="0"/>
        <Setter Property="Height" Value="10"/>
        <Setter Property="MinHeight" Value="10"/>
        <Setter Property="Template">
          <Setter.Value>
            <ControlTemplate TargetType="ScrollBar">
              <Grid Background="Transparent">
                <Track x:Name="PART_Track" IsDirectionReversed="False">
                  <Track.Thumb><Thumb Style="{StaticResource ScrollThumb}" Margin="2"/></Track.Thumb>
                </Track>
              </Grid>
            </ControlTemplate>
          </Setter.Value>
        </Setter>
      </Trigger>
    </Style.Triggers>
  </Style>

  <Style TargetType="Button">
    <Setter Property="Background" Value="#1B212A"/>
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="BorderBrush" Value="#2A323F"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="Padding" Value="13,6"/>
    <Setter Property="MinHeight" Value="32"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7" Padding="{TemplateBinding Padding}">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#4C7DF0"/></Trigger>
            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#4C7DF0"/></Trigger>
            <Trigger Property="IsPressed" Value="True"><Setter TargetName="b" Property="Opacity" Value="0.8"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.4"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
    <Setter Property="Background" Value="#3E6FE0"/>
    <Setter Property="BorderBrush" Value="#3E6FE0"/>
    <Setter Property="Foreground" Value="White"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#5582EC"/></Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
    <Setter Property="Foreground" Value="#FF8A95"/>
    <Setter Property="BorderBrush" Value="#4A2A32"/>
    <Setter Property="Background" Value="#21181C"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#3A1F26"/></Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="DangerPrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
    <Setter Property="Background" Value="#D9475A"/>
    <Setter Property="BorderBrush" Value="#D9475A"/>
    <Setter Property="Foreground" Value="White"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#E85C6E"/></Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="GhostButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderBrush" Value="Transparent"/>
    <Setter Property="Foreground" Value="#AEB6C4"/>
    <Setter Property="Padding" Value="8,4"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True">
        <Setter Property="Background" Value="#1E252F"/>
        <Setter Property="Foreground" Value="#E4E8EF"/>
      </Trigger>
    </Style.Triggers>
  </Style>

  <Style TargetType="TextBox">
    <Setter Property="Background" Value="#1B212A"/>
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="BorderBrush" Value="#2A323F"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="Padding" Value="8,5"/>
    <Setter Property="MinHeight" Value="32"/>
    <Setter Property="CaretBrush" Value="#E4E8EF"/>
    <Setter Property="SelectionBrush" Value="#3E6FE0"/>
    <Setter Property="VerticalContentAlignment" Value="Center"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="TextBox">
          <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7">
            <Grid Margin="{TemplateBinding Padding}">
              <ScrollViewer x:Name="PART_ContentHost" VerticalAlignment="{TemplateBinding VerticalContentAlignment}"/>
              <TextBlock x:Name="wm" Text="{TemplateBinding Tag}" Foreground="#5E6779" Margin="2,0,0,0"
                         VerticalAlignment="{TemplateBinding VerticalContentAlignment}" IsHitTestVisible="False" Visibility="Collapsed"/>
            </Grid>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="Text" Value=""><Setter TargetName="wm" Property="Visibility" Value="Visible"/></Trigger>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#39445A"/></Trigger>
            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#4C7DF0"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="b" Property="Opacity" Value="0.45"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style x:Key="MultiText" TargetType="TextBox" BasedOn="{StaticResource {x:Type TextBox}}">
    <Setter Property="AcceptsReturn" Value="True"/>
    <Setter Property="AcceptsTab" Value="True"/>
    <Setter Property="TextWrapping" Value="NoWrap"/>
    <Setter Property="VerticalContentAlignment" Value="Top"/>
    <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
    <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
    <Setter Property="FontFamily" Value="Consolas"/>
    <Setter Property="FontSize" Value="12.5"/>
  </Style>
  <Style x:Key="ReadOnlyText" TargetType="TextBox" BasedOn="{StaticResource MultiText}">
    <Setter Property="IsReadOnly" Value="True"/>
    <Setter Property="AcceptsTab" Value="False"/>
    <Setter Property="TextWrapping" Value="Wrap"/>
    <Setter Property="HorizontalScrollBarVisibility" Value="Disabled"/>
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="BorderThickness" Value="0"/>
  </Style>

  <Style TargetType="PasswordBox">
    <Setter Property="Background" Value="#1B212A"/>
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="BorderBrush" Value="#2A323F"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="Padding" Value="8,5"/>
    <Setter Property="MinHeight" Value="32"/>
    <Setter Property="CaretBrush" Value="#E4E8EF"/>
    <Setter Property="SelectionBrush" Value="#3E6FE0"/>
    <Setter Property="VerticalContentAlignment" Value="Center"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="PasswordBox">
          <Border x:Name="b" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                  BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="7">
            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="b" Property="BorderBrush" Value="#4C7DF0"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="ComboBox">
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="MinHeight" Value="32"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ComboBox">
          <Grid>
            <ToggleButton Focusable="False" ClickMode="Press"
                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
              <ToggleButton.Template>
                <ControlTemplate TargetType="ToggleButton">
                  <Border x:Name="bd" Background="#1B212A" BorderBrush="#2A323F" BorderThickness="1" CornerRadius="7">
                    <Path HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,11,0"
                          Data="M 0 0 L 4 4 L 8 0" Stroke="#8791A5" StrokeThickness="1.6"/>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#4C7DF0"/></Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </ToggleButton.Template>
            </ToggleButton>
            <ContentPresenter IsHitTestVisible="False" Margin="10,0,28,0" VerticalAlignment="Center"
                              Content="{TemplateBinding SelectionBoxItem}"
                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
            <Popup IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom" AllowsTransparency="True"
                   Focusable="False" PopupAnimation="Fade">
              <Border Background="#1B212A" BorderBrush="#2F3846" BorderThickness="1" CornerRadius="7" Margin="0,3,0,0" Padding="3"
                      MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}" MaxHeight="380">
                <ScrollViewer>
                  <StackPanel IsItemsHost="True"/>
                </ScrollViewer>
              </Border>
            </Popup>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style TargetType="ComboBoxItem">
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="Padding" Value="10,6"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ComboBoxItem">
          <Border x:Name="bd" Background="Transparent" CornerRadius="5" Padding="{TemplateBinding Padding}">
            <ContentPresenter/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="#22304D"/></Trigger>
            <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="bd" Property="Background" Value="#2C3B5E"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="CheckBox">
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="VerticalAlignment" Value="Center"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="CheckBox">
          <Grid Background="Transparent">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border x:Name="box" Width="17" Height="17" CornerRadius="4" BorderThickness="1" BorderBrush="#3A4352" Background="#1B212A" VerticalAlignment="Center">
              <Path x:Name="mark" Data="M 3.5 8 L 6.5 11 L 12.5 4.5" Stroke="White" StrokeThickness="2" Visibility="Collapsed"
                    StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"/>
            </Border>
            <ContentPresenter x:Name="cp" Grid.Column="1" Margin="8,0,0,0" VerticalAlignment="Center"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="Content" Value="{x:Null}"><Setter TargetName="cp" Property="Margin" Value="0"/></Trigger>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="box" Property="BorderBrush" Value="#4C7DF0"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="box" Property="Background" Value="#3E6FE0"/>
              <Setter TargetName="box" Property="BorderBrush" Value="#3E6FE0"/>
              <Setter TargetName="mark" Property="Visibility" Value="Visible"/>
            </Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="SegmentRadio" TargetType="RadioButton">
    <Setter Property="Foreground" Value="#8791A5"/>
    <Setter Property="Padding" Value="12,5"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border x:Name="b" Background="Transparent" CornerRadius="6" Padding="{TemplateBinding Padding}">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="#E4E8EF"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="b" Property="Background" Value="#28303D"/>
              <Setter Property="Foreground" Value="#E4E8EF"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style x:Key="SegmentHost" TargetType="Border">
    <Setter Property="Background" Value="#161B22"/>
    <Setter Property="BorderBrush" Value="#242B36"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="CornerRadius" Value="8"/>
    <Setter Property="Padding" Value="3"/>
  </Style>

  <Style x:Key="WorkspaceTab" TargetType="RadioButton">
    <Setter Property="Foreground" Value="#8791A5"/>
    <Setter Property="Padding" Value="14,7"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Border x:Name="b" Background="Transparent" CornerRadius="7" Padding="{TemplateBinding Padding}">
            <ContentPresenter VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="#E4E8EF"/></Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="b" Property="Background" Value="#1F2B47"/>
              <Setter Property="Foreground" Value="White"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="NavItem" TargetType="RadioButton">
    <Setter Property="Foreground" Value="#AEB6C4"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Margin" Value="0,1"/>
    <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="RadioButton">
          <Grid>
            <Border x:Name="b" Background="Transparent" CornerRadius="7" Padding="10,7">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
            <Border x:Name="bar" Width="3" CornerRadius="2" Background="#4C7DF0" HorizontalAlignment="Left" Margin="0,8" Visibility="Hidden"/>
          </Grid>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True">
              <Setter TargetName="b" Property="Background" Value="#1A2029"/>
              <Setter Property="Foreground" Value="#E4E8EF"/>
            </Trigger>
            <Trigger Property="IsChecked" Value="True">
              <Setter TargetName="b" Property="Background" Value="#1A2640"/>
              <Setter TargetName="bar" Property="Visibility" Value="Visible"/>
              <Setter Property="Foreground" Value="White"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style x:Key="NavHeader" TargetType="TextBlock">
    <Setter Property="Foreground" Value="#5E6779"/>
    <Setter Property="FontSize" Value="11"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Margin" Value="10,14,0,5"/>
  </Style>

  <Style x:Key="ItemCard" TargetType="ListBoxItem">
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="Padding" Value="8,6"/>
    <Setter Property="Margin" Value="0,1"/>
    <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ListBoxItem">
          <Border x:Name="b" Background="Transparent" BorderBrush="Transparent" BorderThickness="1" CornerRadius="7" Padding="{TemplateBinding Padding}">
            <ContentPresenter/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="b" Property="Background" Value="#1A2029"/></Trigger>
            <Trigger Property="IsSelected" Value="True">
              <Setter TargetName="b" Property="Background" Value="#1A2640"/>
              <Setter TargetName="b" Property="BorderBrush" Value="#2F4478"/>
            </Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style TargetType="ListBox">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="ItemContainerStyle" Value="{StaticResource ItemCard}"/>
    <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Disabled"/>
    <Setter Property="VirtualizingPanel.IsVirtualizing" Value="True"/>
    <Setter Property="VirtualizingPanel.VirtualizationMode" Value="Recycling"/>
  </Style>

  <Style TargetType="ProgressBar">
    <Setter Property="Background" Value="#262D39"/>
    <Setter Property="Foreground" Value="#4C7DF0"/>
    <Setter Property="Height" Value="6"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ProgressBar">
          <Grid>
            <Border x:Name="PART_Track" CornerRadius="3" Background="{TemplateBinding Background}"/>
            <Border x:Name="PART_Indicator" CornerRadius="3" Background="{TemplateBinding Foreground}" HorizontalAlignment="Left"/>
          </Grid>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style x:Key="GridCell" TargetType="DataGridCell">
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="DataGridCell">
          <Border Background="{TemplateBinding Background}" Padding="10,4">
            <ContentPresenter VerticalAlignment="Center"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="IsSelected" Value="True">
        <Setter Property="Background" Value="#233354"/>
        <Setter Property="Foreground" Value="White"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="BoolCell" TargetType="DataGridCell" BasedOn="{StaticResource GridCell}">
    <Style.Triggers>
      <DataTrigger Binding="{Binding RelativeSource={RelativeSource Self}, Path=Content.Text}" Value="Tak"><Setter Property="Foreground" Value="#5EE3AE"/></DataTrigger>
      <DataTrigger Binding="{Binding RelativeSource={RelativeSource Self}, Path=Content.Text}" Value="Nie"><Setter Property="Foreground" Value="#FF7A86"/></DataTrigger>
      <Trigger Property="IsSelected" Value="True">
        <Setter Property="Background" Value="#233354"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="BoolCellInv" TargetType="DataGridCell" BasedOn="{StaticResource GridCell}">
    <Style.Triggers>
      <DataTrigger Binding="{Binding RelativeSource={RelativeSource Self}, Path=Content.Text}" Value="Tak"><Setter Property="Foreground" Value="#FF7A86"/></DataTrigger>
      <DataTrigger Binding="{Binding RelativeSource={RelativeSource Self}, Path=Content.Text}" Value="Nie"><Setter Property="Foreground" Value="#5EE3AE"/></DataTrigger>
      <Trigger Property="IsSelected" Value="True">
        <Setter Property="Background" Value="#233354"/>
      </Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="GridRow" TargetType="DataGridRow">
    <Style.Triggers>
      <DataTrigger Binding="{Binding [__flag]}" Value="crit"><Setter Property="Foreground" Value="#FF7A86"/></DataTrigger>
      <DataTrigger Binding="{Binding [__flag]}" Value="warn"><Setter Property="Foreground" Value="#FFC46B"/></DataTrigger>
      <DataTrigger Binding="{Binding [__flag]}" Value="muted"><Setter Property="Foreground" Value="#7B8496"/></DataTrigger>
      <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#1D2430"/></Trigger>
    </Style.Triggers>
  </Style>
  <Style x:Key="GridHeader" TargetType="DataGridColumnHeader">
    <Setter Property="Background" Value="#1B212A"/>
    <Setter Property="Foreground" Value="#8791A5"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="FontSize" Value="12"/>
    <Setter Property="Padding" Value="10,8"/>
    <Setter Property="BorderBrush" Value="#252C38"/>
    <Setter Property="BorderThickness" Value="0,0,1,1"/>
  </Style>
  <Style x:Key="DarkGrid" TargetType="DataGrid">
    <Style.Resources>
      <SolidColorBrush x:Key="{x:Static SystemColors.ControlBrushKey}" Color="#161B22"/>
    </Style.Resources>
    <Setter Property="Background" Value="#161B22"/>
    <Setter Property="Foreground" Value="#D5DAE3"/>
    <Setter Property="RowBackground" Value="#161B22"/>
    <Setter Property="AlternatingRowBackground" Value="#181E26"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="GridLinesVisibility" Value="Horizontal"/>
    <Setter Property="HorizontalGridLinesBrush" Value="#1F2530"/>
    <Setter Property="HeadersVisibility" Value="Column"/>
    <Setter Property="IsReadOnly" Value="True"/>
    <Setter Property="CanUserAddRows" Value="False"/>
    <Setter Property="CanUserDeleteRows" Value="False"/>
    <Setter Property="CanUserResizeRows" Value="False"/>
    <Setter Property="SelectionMode" Value="Extended"/>
    <Setter Property="SelectionUnit" Value="FullRow"/>
    <Setter Property="ClipboardCopyMode" Value="IncludeHeader"/>
    <Setter Property="MinRowHeight" Value="28"/>
    <Setter Property="FontSize" Value="12.5"/>
    <Setter Property="EnableRowVirtualization" Value="True"/>
    <Setter Property="EnableColumnVirtualization" Value="True"/>
    <Setter Property="VirtualizingPanel.VirtualizationMode" Value="Recycling"/>
    <Setter Property="ColumnHeaderStyle" Value="{StaticResource GridHeader}"/>
    <Setter Property="CellStyle" Value="{StaticResource GridCell}"/>
    <Setter Property="RowStyle" Value="{StaticResource GridRow}"/>
  </Style>

  <Style TargetType="ContextMenu">
    <Setter Property="Background" Value="#1B212A"/>
    <Setter Property="BorderBrush" Value="#2F3846"/>
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ContextMenu">
          <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="4">
            <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Cycle"/>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style TargetType="MenuItem">
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="Cursor" Value="Hand"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="MenuItem">
          <Border x:Name="b" Background="Transparent" CornerRadius="6" Padding="8,6" MinWidth="200">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="26"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <ContentPresenter ContentSource="Icon" VerticalAlignment="Center" TextElement.Foreground="#8791A5"/>
              <ContentPresenter Grid.Column="1" ContentSource="Header" VerticalAlignment="Center"/>
            </Grid>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="b" Property="Background" Value="#24304A"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.4"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style x:Key="MenuSeparator" TargetType="Separator">
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Separator">
          <Border Height="1" Background="#2A323F" Margin="6,4"/>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="ToolTip">
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ToolTip">
          <Border Background="#232A35" BorderBrush="#2F3846" BorderThickness="1" CornerRadius="6" Padding="9,6" MaxWidth="460">
            <ContentPresenter>
              <ContentPresenter.Resources>
                <Style TargetType="TextBlock"><Setter Property="TextWrapping" Value="Wrap"/></Style>
              </ContentPresenter.Resources>
            </ContentPresenter>
          </Border>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>

  <Style TargetType="TreeView">
    <Style.Resources>
      <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" Color="#2A3B63"/>
      <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="White"/>
      <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightBrushKey}" Color="#24304A"/>
      <SolidColorBrush x:Key="{x:Static SystemColors.InactiveSelectionHighlightTextBrushKey}" Color="White"/>
    </Style.Resources>
    <Setter Property="Background" Value="#161B22"/>
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="BorderBrush" Value="#242B36"/>
    <Setter Property="Padding" Value="6"/>
  </Style>
  <Style TargetType="TreeViewItem">
    <Setter Property="Foreground" Value="#E4E8EF"/>
    <Setter Property="Padding" Value="4,3"/>
  </Style>

  <Style TargetType="GridSplitter">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Focusable" Value="False"/>
    <Style.Triggers>
      <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#2F4478"/></Trigger>
    </Style.Triggers>
  </Style>

  <Style x:Key="Chip" TargetType="Border">
    <Setter Property="Background" Value="#1E252F"/>
    <Setter Property="CornerRadius" Value="9"/>
    <Setter Property="Padding" Value="9,2"/>
    <Setter Property="Margin" Value="0,0,6,0"/>
    <Setter Property="VerticalAlignment" Value="Center"/>
  </Style>
</ResourceDictionary>
'@

$script:ThemeDoc = $null
$script:Theme = $null

function Get-Theme {
    # Jedna instancja słownika - style i pędzle dla kontrolek tworzonych w kodzie
    if (-not $script:Theme) { $script:Theme = [System.Windows.Markup.XamlReader]::Parse($script:ThemeXaml) }
    return $script:Theme
}

function New-UiElement {
    # Ładuje XAML z wstrzykniętym motywem: zasoby muszą istnieć w chwili parsowania (StaticResource),
    # dlatego słownik motywu trafia na początek <Root.Resources> korzenia dokumentu.
    param([Parameter(Mandatory)][string]$Xaml)
    if (-not $script:ThemeDoc) {
        $script:ThemeDoc = New-Object System.Xml.XmlDocument
        $script:ThemeDoc.LoadXml($script:ThemeXaml)
    }
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $false
    $doc.LoadXml($Xaml)
    $root = $doc.DocumentElement
    $resName = $root.LocalName + '.Resources'
    $existing = $null
    foreach ($child in @($root.ChildNodes)) { if ($child.LocalName -eq $resName) { $existing = $child } }
    $dictionary = $doc.ImportNode($script:ThemeDoc.DocumentElement, $true)
    if ($existing) {
        # Lokalne zasoby dokumentu dopisujemy za motywem (mogą z niego korzystać)
        foreach ($child in @($existing.ChildNodes)) { [void]$dictionary.AppendChild($child) }
        [void]$root.RemoveChild($existing)
    }
    $resources = $doc.CreateElement($resName, $root.NamespaceURI)
    [void]$resources.AppendChild($dictionary)
    [void]$root.PrependChild($resources)
    $reader = New-Object System.Xml.XmlNodeReader $doc
    return [System.Windows.Markup.XamlReader]::Load($reader)
}

function Get-Brush([string]$Color) {
    $brush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($Color))
    $brush.Freeze()
    return $brush
}
#endregion

#region Fabryka kontrolek modułów
# Moduły budują swój panel parametrów z tych funkcji. Wszystkie kontrolki dostają styl motywu
# (style domyślne z Get-Theme), a obsługa zdarzeń przechodzi przez wspólny dyspozytor, dzięki czemu
# akcja dostaje kontekst modułu ($m) niezależnie od miejsca, w którym została zdefiniowana.

function Get-Glyph([string]$Code) {
    if (-not $Code) { return '' }
    return [string][char][Convert]::ToInt32($Code, 16)
}

function Get-ThemeResource([string]$Key) {
    return (Get-Theme)[$Key]
}

function New-GlyphBlock {
    param([string]$Code, [double]$Size = 14, [string]$Color = '')
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Style = Get-ThemeResource 'Glyph'
    $t.Text = Get-Glyph $Code
    $t.FontSize = $Size
    if ($Color) { $t.Foreground = Get-Brush $Color }
    return $t
}

function New-IconContent {
    # Zawartość przycisku: ikona + tekst
    param([string]$Text, [string]$Icon, [double]$IconSize = 13)
    if (-not $Icon) { return $Text }
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Orientation = 'Horizontal'
    $g = New-GlyphBlock -Code $Icon -Size $IconSize
    if ($Text) { $g.Margin = '0,0,8,0' }
    [void]$sp.Children.Add($g)
    if ($Text) {
        $t = New-Object System.Windows.Controls.TextBlock
        $t.Text = $Text
        $t.VerticalAlignment = 'Center'
        [void]$sp.Children.Add($t)
    }
    return $sp
}

# Rejestr obsługi zdarzeń: kontrolka -> @{ Module; Click; TextChanged; ... }
$script:Handlers = New-Object 'System.Collections.Generic.Dictionary[object,hashtable]'
$script:Dispatchers = @{
    Click            = { param($s, $e) Invoke-ControlHandler -Source $s -EventName 'Click' }
    TextChanged      = { param($s, $e) Invoke-ControlHandler -Source $s -EventName 'TextChanged' }
    Checked          = { param($s, $e) Invoke-ControlHandler -Source $s -EventName 'Checked' }
    Unchecked        = { param($s, $e) Invoke-ControlHandler -Source $s -EventName 'Unchecked' }
    SelectionChanged = { param($s, $e) Invoke-ControlHandler -Source $s -EventName 'SelectionChanged' }
}

function Register-ControlHandler {
    param($Control, [string]$EventName, [hashtable]$Module, [scriptblock]$Action)
    $entry = $null
    if (-not $script:Handlers.TryGetValue($Control, [ref]$entry)) {
        $entry = @{ Module = $Module }
        $script:Handlers[$Control] = $entry
    }
    $entry.Module = $Module
    $isNew = -not $entry.ContainsKey($EventName)
    $entry[$EventName] = $Action
    if ($isNew) { $Control."add_$EventName"($script:Dispatchers[$EventName]) }
}

function Invoke-ControlHandler {
    param($Source, [string]$EventName)
    try {
        $entry = $null
        if (-not $script:Handlers.TryGetValue($Source, [ref]$entry)) { return }
        $action = $entry[$EventName]
        if ($action) { Invoke-UiAction -Module $entry.Module -Action $action -Source $Source }
    }
    catch {
        Write-Log "Błąd obsługi zdarzenia: $($_.Exception.Message)" 'ERROR'
    }
}

function Invoke-UiAction {
    param([hashtable]$Module, [scriptblock]$Action, $Source = $null)
    $previous = $script:LogContext
    $script:LogContext = if ($Module) { $Module.Title } else { $null }
    try {
        $null = & $Action $Module $Source
    }
    catch {
        Write-Log "Błąd: $($_.Exception.Message)" 'ERROR'
        Show-Error 'Operacja nie powiodła się.' $_
    }
    finally {
        $script:LogContext = $previous
    }
}

function Add-ParamRow {
    # Wiersz panelu parametrów z podpisem po lewej (-Title, wyrównany we wszystkich wierszach) i dowolną zawartością
    param([Parameter(Mandatory)][hashtable]$Module, [string]$Title = '', [Parameter(Mandatory)]$Content)
    $container = $Content
    if ($Title) {
        $grid = New-Object System.Windows.Controls.Grid
        $c1 = New-Object System.Windows.Controls.ColumnDefinition
        $c1.Width = New-Object System.Windows.GridLength 128
        $c2 = New-Object System.Windows.Controls.ColumnDefinition
        [void]$grid.ColumnDefinitions.Add($c1)
        [void]$grid.ColumnDefinitions.Add($c2)
        $caption = New-Object System.Windows.Controls.TextBlock
        $caption.Text = $Title
        $caption.Foreground = Get-Brush '#8791A5'
        $caption.FontWeight = 'SemiBold'
        $caption.FontSize = 12
        $caption.Margin = '0,8,12,0'
        $caption.VerticalAlignment = 'Top'
        $caption.TextWrapping = 'Wrap'
        [void]$grid.Children.Add($caption)
        [System.Windows.Controls.Grid]::SetColumn($Content, 1)
        [void]$grid.Children.Add($Content)
        $container = $grid
    }
    $stack = $Module.ParamsStack
    if ($stack.Children.Count -gt 0) { $container.Margin = '0,4,0,0' }
    [void]$stack.Children.Add($container)
    $Module.ParamsCard.Visibility = 'Visible'
    return $Content
}

function Add-ToolbarRow {
    # Wiersz kontrolek (zawijany) w panelu parametrów
    param([Parameter(Mandatory)][hashtable]$Module, [string]$Title = '')
    $wrap = New-Object System.Windows.Controls.WrapPanel
    $wrap.Orientation = 'Horizontal'
    return (Add-ParamRow -Module $Module -Title $Title -Content $wrap)
}

function Add-StretchTextBox {
    # Pole tekstowe na całą szerokość panelu (np. wieloliniowe polecenie)
    param([Parameter(Mandatory)][hashtable]$Module, [string]$Title = '', [switch]$Multiline, [double]$Height = 120, [string]$Placeholder = '', [string]$Text = '')
    $tb = New-Object System.Windows.Controls.TextBox
    if ($Multiline) {
        $tb.Style = Get-ThemeResource 'MultiText'
        $tb.Height = $Height
    }
    $tb.Text = $Text
    if ($Placeholder) { $tb.Tag = $Placeholder }
    $border = New-Object System.Windows.Controls.Border
    $border.Padding = '0,0,8,6'
    $border.Child = $tb
    [void](Add-ParamRow -Module $Module -Title $Title -Content $border)
    return $tb
}

function Add-Separator {
    # Pozioma linia oddzielająca grupy wierszy w panelu parametrów
    param([Parameter(Mandatory)][hashtable]$Module)
    $b = New-Object System.Windows.Controls.Border
    $b.Height = 1
    $b.Background = Get-Brush '#242B36'
    $b.Margin = '0,8,0,6'
    [void]$Module.ParamsStack.Children.Add($b)
}

function Add-TopControl {
    # Dowolny element w panelu parametrów (pełna szerokość)
    param([hashtable]$Module, $Control)
    [void]$Module.ParamsStack.Children.Add($Control)
    $Module.ParamsCard.Visibility = 'Visible'
    return $Control
}

function Add-Label {
    param([Parameter(Mandatory)]$Parent, [Parameter(Mandatory)][AllowEmptyString()][string]$Text, [switch]$Hint, [switch]$Bold, [double]$MaxWidth = 0)
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $Text
    $t.VerticalAlignment = 'Center'
    $t.Margin = '0,0,8,6'
    $t.TextWrapping = 'Wrap'
    if ($MaxWidth -gt 0) { $t.MaxWidth = $MaxWidth }
    if ($Hint) { $t.Foreground = Get-Brush '#7B8496'; $t.FontSize = 12 }
    if ($Bold) { $t.FontWeight = 'SemiBold' }
    [void]$Parent.Children.Add($t)
    return $t
}

function Add-TextBox {
    param([Parameter(Mandatory)]$Parent, [double]$Width = 200, [string]$Text = '', [string]$Placeholder = '', [switch]$Multiline, [double]$Height = 0)
    $tb = New-Object System.Windows.Controls.TextBox
    if ($Multiline) {
        $tb.Style = Get-ThemeResource 'MultiText'
        $tb.Height = if ($Height -gt 0) { $Height } else { 140 }
    }
    $tb.Width = $Width
    $tb.Text = $Text
    if ($Placeholder) { $tb.Tag = $Placeholder }
    $tb.Margin = '0,0,8,6'
    [void]$Parent.Children.Add($tb)
    return $tb
}

function Add-PasswordBox {
    param([Parameter(Mandatory)]$Parent, [double]$Width = 200)
    $pb = New-Object System.Windows.Controls.PasswordBox
    $pb.Width = $Width
    $pb.Margin = '0,0,8,6'
    [void]$Parent.Children.Add($pb)
    return $pb
}

function Add-ComboBox {
    param([Parameter(Mandatory)]$Parent, [string[]]$Items = @(), [double]$Width = 180, [int]$Selected = 0)
    $cb = New-Object System.Windows.Controls.ComboBox
    $cb.Width = $Width
    $cb.Margin = '0,0,8,6'
    foreach ($i in $Items) { [void]$cb.Items.Add($i) }
    if ($cb.Items.Count -gt 0) { $cb.SelectedIndex = [Math]::Min([Math]::Max(0, $Selected), $cb.Items.Count - 1) }
    [void]$Parent.Children.Add($cb)
    return $cb
}

$script:NumericSpecs = New-Object 'System.Collections.Generic.Dictionary[object,hashtable]'
$script:NumericEvents = @{
    PreviewTextInput = {
        param($s, $e)
        if ($e.Text -notmatch '^\d+$') { $e.Handled = $true }
    }
    LostFocus        = {
        param($s, $e)
        try { $s.Text = [string](Get-Num $s) } catch { }
    }
}

function Add-Numeric {
    # Pole liczbowe (TextBox przyjmujący cyfry); wartość odczytuje Get-Num z ograniczeniem do zakresu
    param([Parameter(Mandatory)]$Parent, [int]$Value = 0, [int]$Minimum = 0, [int]$Maximum = 100000, [double]$Width = 80)
    $tb = New-Object System.Windows.Controls.TextBox
    $tb.Width = $Width
    $tb.Margin = '0,0,8,6'
    $tb.Text = [string]$Value
    $tb.HorizontalContentAlignment = 'Right'
    $script:NumericSpecs[$tb] = @{ Min = $Minimum; Max = $Maximum; Default = $Value }
    $tb.add_PreviewTextInput($script:NumericEvents.PreviewTextInput)
    $tb.add_LostFocus($script:NumericEvents.LostFocus)
    [void]$Parent.Children.Add($tb)
    return $tb
}

function Get-Num {
    param([Parameter(Mandatory)]$Control)
    $spec = $null
    if (-not $script:NumericSpecs.TryGetValue($Control, [ref]$spec)) { $spec = @{ Min = 0; Max = [int]::MaxValue; Default = 0 } }
    $v = 0
    if (-not [int]::TryParse(([string]$Control.Text).Trim(), [ref]$v)) { $v = $spec.Default }
    return [Math]::Min($spec.Max, [Math]::Max($spec.Min, $v))
}

function Add-CheckBox {
    param([Parameter(Mandatory)]$Parent, [Parameter(Mandatory)][string]$Text, [bool]$Checked = $false, [string]$ToolTip = '')
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Content = $Text
    $cb.IsChecked = $Checked
    $cb.Margin = '2,0,16,6'
    if ($ToolTip) { $cb.ToolTip = $ToolTip }
    [void]$Parent.Children.Add($cb)
    return $cb
}

function Test-Checked($CheckBox) {
    return ($CheckBox.IsChecked -eq $true)
}

$script:SegmentCounter = 0
function Add-Segmented {
    # Przełącznik segmentowy (zamiast grupy przycisków radiowych); indeks wyboru: Get-SegmentIndex
    param([AllowNull()]$Parent, [Parameter(Mandatory)][string[]]$Items, [int]$Selected = 0, [hashtable]$Module, [scriptblock]$OnChange)
    $script:SegmentCounter++
    $group = 'seg' + $script:SegmentCounter
    $container = New-Object System.Windows.Controls.Border
    $container.Style = Get-ThemeResource 'SegmentHost'
    $container.Margin = '0,0,8,6'
    $container.HorizontalAlignment = 'Left'
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Orientation = 'Horizontal'
    $container.Child = $sp
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $rb = New-Object System.Windows.Controls.RadioButton
        $rb.Style = Get-ThemeResource 'SegmentRadio'
        $rb.GroupName = $group
        $rb.Content = $Items[$i]
        $rb.IsChecked = ($i -eq $Selected)
        if ($OnChange) { Register-ControlHandler -Control $rb -EventName 'Checked' -Module $Module -Action $OnChange }
        [void]$sp.Children.Add($rb)
    }
    if ($Parent) { [void]$Parent.Children.Add($container) }
    return $container
}

function Get-SegmentIndex($Segmented) {
    $i = 0
    foreach ($rb in $Segmented.Child.Children) {
        if ($rb.IsChecked -eq $true) { return $i }
        $i++
    }
    return -1
}

function Set-SegmentIndex($Segmented, [int]$Index) {
    $i = 0
    foreach ($rb in $Segmented.Child.Children) {
        $rb.IsChecked = ($i -eq $Index)
        $i++
    }
}

function New-PlainButton {
    param([string]$Text, [string]$Icon = '', [switch]$Primary, [switch]$Danger, [switch]$Ghost, [string]$ToolTip = '')
    $b = New-Object System.Windows.Controls.Button
    if ($Primary -and $Danger) { $b.Style = Get-ThemeResource 'DangerPrimaryButton' }
    elseif ($Primary) { $b.Style = Get-ThemeResource 'PrimaryButton' }
    elseif ($Danger) { $b.Style = Get-ThemeResource 'DangerButton' }
    elseif ($Ghost) { $b.Style = Get-ThemeResource 'GhostButton' }
    $b.Content = New-IconContent -Text $Text -Icon $Icon
    if ($ToolTip) { $b.ToolTip = $ToolTip }
    return $b
}

function Add-Button {
    # Przycisk akcji modułu - wyłączany automatycznie na czas operacji tego modułu.
    # Pierwszy przycisk -Primary jest domyślną akcją modułu (klawisz F5).
    param(
        [Parameter(Mandatory)]$Parent,
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text,
        [Parameter(Mandatory)][hashtable]$Module,
        [Parameter(Mandatory)][scriptblock]$OnClick,
        [string]$Icon = '',
        [string]$ToolTip = '',
        [switch]$Primary,
        [switch]$Danger,
        [switch]$AlwaysEnabled
    )
    $b = New-PlainButton -Text $Text -Icon $Icon -Primary:$Primary -Danger:$Danger -ToolTip $ToolTip
    $b.Margin = '0,0,8,6'
    Register-ControlHandler -Control $b -EventName 'Click' -Module $Module -Action $OnClick
    if (-not $AlwaysEnabled) { [void]$Module.Buttons.Add($b) }
    if ($Primary -and -not $Module.PrimaryButton) { $Module.PrimaryButton = $b }
    [void]$Parent.Children.Add($b)
    return $b
}

function Add-StatTile {
    # Kafelek z liczbą nad tabelą wyników (np. "Kandydaci do usunięcia: 12")
    param([Parameter(Mandatory)][hashtable]$Module, [Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][string]$Label, [string]$Icon = '', [string]$Value = '–')
    $xaml = @"
<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Padding="14,10" Margin="0,0,10,0">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
    <Border Width="34" Height="34" CornerRadius="8" Background="#1B2333" Margin="0,0,12,0" VerticalAlignment="Center">
      <TextBlock Name="icon" Style="{StaticResource Glyph}" FontSize="15" Foreground="#8CB0FF" HorizontalAlignment="Center"/>
    </Border>
    <StackPanel Grid.Column="1" VerticalAlignment="Center">
      <TextBlock Name="value" FontSize="20" FontWeight="SemiBold" Foreground="White"/>
      <TextBlock Name="label" FontSize="11.5" Foreground="#8791A5" TextTrimming="CharacterEllipsis"/>
    </StackPanel>
  </Grid>
</Border>
"@
    $tile = New-UiElement $xaml
    # Styl ustawiany w kodzie: StaticResource w atrybucie korzenia nie widzi jego własnych zasobów
    $tile.Style = Get-ThemeResource 'Card'
    $tile.Padding = '14,10'
    $tile.FindName('icon').Text = Get-Glyph $(if ($Icon) { $Icon } else { 'E9D2' })
    $tile.FindName('value').Text = $Value
    $tile.FindName('label').Text = $Label
    $Module.Stats[$Key] = $tile
    [void]$Module.StatsGrid.Children.Add($tile)
    $Module.StatsGrid.Columns = $Module.StatsGrid.Children.Count
    $Module.StatsGrid.Visibility = 'Visible'
    return $tile
}

function Set-StatTile {
    param([Parameter(Mandatory)][hashtable]$Module, [Parameter(Mandatory)][string]$Key, [string]$Value = '–', [ValidateSet('', 'ok', 'warn', 'crit', 'info')][string]$Tone = '', [string]$Label = '')
    $tile = $Module.Stats[$Key]
    if (-not $tile) { return }
    $v = $tile.FindName('value')
    $v.Text = $Value
    $v.Foreground = Get-Brush $(switch ($Tone) { 'ok' { '#5EE3AE' } 'warn' { '#FFC46B' } 'crit' { '#FF7A86' } 'info' { '#8CC0FF' } default { '#FFFFFF' } })
    if ($Label) { $tile.FindName('label').Text = $Label }
}

function Reset-StatTiles([hashtable]$Module) {
    foreach ($k in @($Module.Stats.Keys)) { Set-StatTile -Module $Module -Key $k -Value '–' }
}

function Add-RowAction {
    # Pozycja menu kontekstowego tabeli wyników; akcja dostaje ($m, $rows) - zaznaczone wiersze (DataRowView)
    param([Parameter(Mandatory)][hashtable]$Module, [Parameter(Mandatory)][string]$Text, [Parameter(Mandatory)][scriptblock]$Action, [string]$Icon = '', [switch]$Danger, [switch]$Separator)
    [void]$Module.RowActions.Add(@{ Text = $Text; Action = $Action; Icon = $Icon; Danger = [bool]$Danger; Separator = [bool]$Separator })
}
#endregion

#region Okna dialogowe
# Wspólny szablon okna: nagłówek z ikoną, treść (wstawiana w miejsce <!--BODY-->) i przyciski.
# Okna są modalne względem okna głównego. $script:DialogHook pozwala testom obsłużyć okno zamiast ShowDialog.
$script:DialogHook = $null
$script:DialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="480" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" Background="#12171D" Foreground="#E4E8EF" FontFamily="Segoe UI" FontSize="13"
        UseLayoutRounding="True" SnapsToDevicePixels="True" TextOptions.TextFormattingMode="Display">
  <Grid Background="#12171D">
  <Grid Margin="22,20,22,18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid Margin="0,0,0,16">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Border x:Name="dlgIconTile" Width="38" Height="38" CornerRadius="10" Background="#1A2640" Margin="0,0,14,0" VerticalAlignment="Top">
        <TextBlock x:Name="dlgIcon" Style="{StaticResource Glyph}" FontSize="17" Foreground="#8CB0FF" HorizontalAlignment="Center"/>
      </Border>
      <StackPanel Grid.Column="1" VerticalAlignment="Center">
        <TextBlock x:Name="dlgTitle" FontSize="16" FontWeight="SemiBold" Foreground="White" TextWrapping="Wrap"/>
        <TextBlock x:Name="dlgSubtitle" Foreground="#8791A5" TextWrapping="Wrap" Margin="0,3,0,0"/>
      </StackPanel>
    </Grid>
    <Grid x:Name="dlgBody" Grid.Row="1">
<!--BODY-->
    </Grid>
    <Grid Grid.Row="2" Margin="0,18,0,0">
      <StackPanel x:Name="dlgExtra" Orientation="Horizontal" HorizontalAlignment="Left"/>
      <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
        <Button x:Name="btnCancel" Content="Anuluj" MinWidth="96" IsCancel="True"/>
        <Button x:Name="btnOk" Content="OK" MinWidth="96" Margin="8,0,0,0" Style="{StaticResource PrimaryButton}" IsDefault="True"/>
      </StackPanel>
    </Grid>
  </Grid>
  </Grid>
</Window>
'@

$script:DialogTones = @{
    info = @{ Back = '#1A2640'; Fore = '#8CB0FF'; Icon = 'E946' }
    ok   = @{ Back = '#15291F'; Fore = '#5EE3AE'; Icon = 'E930' }
    warn = @{ Back = '#2E2616'; Fore = '#FFC46B'; Icon = 'E7BA' }
    crit = @{ Back = '#2E1A1E'; Fore = '#FF7A86'; Icon = 'EA39' }
}

$script:DialogEvents = @{
    Ok     = {
        param($s, $e)
        $w = [System.Windows.Window]::GetWindow($s)
        if (-not $w) { return }
        $validate = $w.Tag.Validate
        if ($validate) {
            $ok = $false
            try { $ok = [bool](& $validate $w) }
            catch { Show-Error 'Nie można zatwierdzić.' $_; return }
            if (-not $ok) { return }
        }
        Close-Dialog -Window $w -Ok $true
    }
    Cancel = {
        param($s, $e)
        $w = [System.Windows.Window]::GetWindow($s)
        if (-not $w) { return }
        $w.Tag.Result = $false
        if (-not $w.Tag.Modal) { $w.Close() }
    }
}

$script:DwmReady = $null
function Set-DarkTitleBar {
    # Ciemny pasek tytułu (Windows 10 20H1+ / 11). Brak obsługi - okno zostaje z jasnym paskiem.
    param($Window)
    try {
        if ($null -eq $script:DwmReady) {
            $script:DwmReady = $false
            if (-not ('DomainOps.Dwm' -as [type])) {
                Add-Type -Namespace DomainOps -Name Dwm -ErrorAction Stop -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("dwmapi.dll")]
public static extern int DwmSetWindowAttribute(System.IntPtr hwnd, int attribute, ref int value, int size);
'@
            }
            $script:DwmReady = $true
        }
        if (-not $script:DwmReady) { return }
        $hwnd = (New-Object System.Windows.Interop.WindowInteropHelper $Window).Handle
        if ($hwnd -eq [IntPtr]::Zero) { return }
        $on = 1
        if ([DomainOps.Dwm]::DwmSetWindowAttribute($hwnd, 20, [ref]$on, 4) -ne 0) {
            [void][DomainOps.Dwm]::DwmSetWindowAttribute($hwnd, 19, [ref]$on, 4)
        }
        $caption = 0x001D1712   # #12171D jako COLORREF (0x00BBGGRR) - Windows 11
        [void][DomainOps.Dwm]::DwmSetWindowAttribute($hwnd, 35, [ref]$caption, 4)
    }
    catch { }
}

function New-Dialog {
    param(
        [Parameter(Mandatory)][string]$Title,
        [string]$Subtitle = '',
        [string]$Body = '',
        [ValidateSet('info', 'ok', 'warn', 'crit')][string]$Tone = 'info',
        [string]$Icon = '',
        [double]$Width = 480,
        [double]$Height = 0,
        [string]$OkText = 'OK',
        [string]$CancelText = 'Anuluj',
        [switch]$NoCancel,
        [switch]$Danger,
        [switch]$Resizable,
        [scriptblock]$Validate
    )
    $w = New-UiElement ($script:DialogXaml.Replace('<!--BODY-->', $Body))
    $w.Title = $Title
    $w.Width = $Width
    if ($Height -gt 0) {
        $w.SizeToContent = 'Manual'
        $w.Height = $Height
    }
    if ($Resizable) {
        $w.ResizeMode = 'CanResizeWithGrip'
        $w.MinWidth = [Math]::Min($Width, 420)
        $w.MinHeight = 260
    }
    $toneDef = $script:DialogTones[$Tone]
    $w.FindName('dlgIconTile').Background = Get-Brush $toneDef.Back
    $iconBlock = $w.FindName('dlgIcon')
    $iconBlock.Foreground = Get-Brush $toneDef.Fore
    $iconBlock.Text = Get-Glyph $(if ($Icon) { $Icon } else { $toneDef.Icon })
    $w.FindName('dlgTitle').Text = $Title
    $sub = $w.FindName('dlgSubtitle')
    if ($Subtitle) { $sub.Text = $Subtitle } else { $sub.Visibility = 'Collapsed' }
    $ok = $w.FindName('btnOk')
    $ok.Content = $OkText
    if ($Danger) { $ok.Style = Get-ThemeResource 'DangerPrimaryButton' }
    $cancel = $w.FindName('btnCancel')
    $cancel.Content = $CancelText
    if ($NoCancel) {
        $cancel.Visibility = 'Collapsed'
        $ok.IsCancel = $true
    }
    $ok.add_Click($script:DialogEvents.Ok)
    $cancel.add_Click($script:DialogEvents.Cancel)
    $w.Tag = @{ Result = $false; Modal = $false; Validate = $Validate }
    return $w
}

function Invoke-Dialog {
    # Pokazuje okno modalnie; zwraca $true, gdy zatwierdzono
    param([Parameter(Mandatory)][System.Windows.Window]$Window)
    if (-not ($Window.Tag -is [hashtable])) { $Window.Tag = @{} }
    $Window.Tag['Result'] = $false
    $Window.Tag['Modal'] = $false
    $main = $script:UI.Window
    if ($main -and $main.IsVisible -and -not [object]::ReferenceEquals($main, $Window)) { $Window.Owner = $main }
    else { $Window.WindowStartupLocation = 'CenterScreen' }
    $Window.add_SourceInitialized({ param($s, $e) Set-DarkTitleBar $s })
    if ($script:DialogHook) {
        $null = & $script:DialogHook $Window
    }
    else {
        $Window.Tag.Modal = $true
        [void]$Window.ShowDialog()
    }
    return [bool]$Window.Tag.Result
}

function Close-Dialog {
    param([Parameter(Mandatory)]$Window, [bool]$Ok)
    $Window.Tag.Result = $Ok
    if ($Window.Tag.Modal) {
        try { $Window.DialogResult = $Ok } catch { try { $Window.Close() } catch { } }
    }
    else {
        try { $Window.Close() } catch { }
    }
}

function New-DialogText {
    # Akapit tekstu do treści okna (zawijany)
    param([string]$Text, [string]$Color = '#C9D0DC', [double]$Size = 13)
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $Text
    $t.TextWrapping = 'Wrap'
    $t.Foreground = Get-Brush $Color
    $t.FontSize = $Size
    return $t
}

function Show-Message {
    param([string]$Text, [string]$Title = 'Domain Ops', [ValidateSet('info', 'ok', 'warn', 'crit')][string]$Tone = 'info')
    $body = @'
<ScrollViewer MaxHeight="420" VerticalScrollBarVisibility="Auto">
  <TextBlock x:Name="msgText" TextWrapping="Wrap" Foreground="#C9D0DC" LineHeight="19"/>
</ScrollViewer>
'@
    $w = New-Dialog -Title $Title -Body $body -Tone $Tone -NoCancel -Width 500
    $w.FindName('msgText').Text = $Text
    [void](Invoke-Dialog $w)
}

function Show-Warning([string]$Text) {
    Show-Message -Text $Text -Title 'Uwaga' -Tone 'warn'
}

function Show-Error {
    param([string]$Text, $ErrorObject = $null)
    # $ErrorObject może być ErrorRecord ($_ z bloku catch), wyjątkiem albo tekstem
    $detail = if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $ErrorObject.Exception.Message }
    elseif ($ErrorObject -is [System.Exception]) { $ErrorObject.Message }
    elseif ($ErrorObject) { [string]$ErrorObject }
    else { '' }
    $message = if ($detail) { "$Text`r`n`r`n$detail" } else { $Text }
    Show-Message -Text $message -Title 'Błąd' -Tone 'crit'
}

function Confirm-Action {
    # Pytanie z listą obiektów, których dotyczy operacja
    param(
        [Parameter(Mandatory)][string]$Text,
        [string[]]$Items = @(),
        [string]$Title = 'Potwierdzenie',
        [string]$ConfirmText = 'Wykonaj',
        [switch]$Danger
    )
    $body = @'
<StackPanel>
  <TextBlock x:Name="cfText" TextWrapping="Wrap" Foreground="#C9D0DC" LineHeight="19"/>
  <Border x:Name="cfListHost" Style="{StaticResource Card}" Padding="12,8" Margin="0,14,0,0" Visibility="Collapsed">
    <ScrollViewer MaxHeight="230" VerticalScrollBarVisibility="Auto">
      <ItemsControl x:Name="cfList">
        <ItemsControl.ItemTemplate>
          <DataTemplate>
            <StackPanel Orientation="Horizontal" Margin="0,2">
              <Ellipse Width="5" Height="5" Fill="#5E6779" Margin="0,0,9,0" VerticalAlignment="Center"/>
              <TextBlock Text="{Binding}" Foreground="#E4E8EF" TextTrimming="CharacterEllipsis"/>
            </StackPanel>
          </DataTemplate>
        </ItemsControl.ItemTemplate>
      </ItemsControl>
    </ScrollViewer>
  </Border>
  <TextBlock x:Name="cfCount" Foreground="#7B8496" FontSize="12" Margin="2,8,0,0" Visibility="Collapsed"/>
</StackPanel>
'@
    $tone = if ($Danger) { 'crit' } else { 'warn' }
    $w = New-Dialog -Title $Title -Body $body -Tone $tone -OkText $ConfirmText -Danger:$Danger -Width 520
    $w.FindName('cfText').Text = $Text
    $list = @($Items | Where-Object { $_ } | ForEach-Object { [string]$_ })
    if ($list.Count -gt 0) {
        $w.FindName('cfListHost').Visibility = 'Visible'
        $w.FindName('cfList').ItemsSource = @($list | Select-Object -First 200)
        $count = $w.FindName('cfCount')
        $count.Text = if ($list.Count -gt 200) { "Pozycji: $($list.Count) (pokazano 200)" } else { "Pozycji: $($list.Count)" }
        $count.Visibility = 'Visible'
    }
    # Przy operacjach niszczących domyślnym przyciskiem (Enter) jest Anuluj
    if ($Danger) {
        $w.FindName('btnOk').IsDefault = $false
        $w.FindName('btnCancel').IsDefault = $true
        $w.add_ContentRendered({ param($s, $e) [void]$s.FindName('btnCancel').Focus() })
    }
    return (Invoke-Dialog $w)
}

function Show-CredentialDialog {
    param([string]$Message = 'Podaj konto z uprawnieniami administracyjnymi. Poświadczenia są przechowywane tylko w pamięci do zamknięcia programu.', [string]$UserName = '')
    $body = @'
<StackPanel>
  <TextBlock Text="Użytkownik" Foreground="#8791A5" FontSize="12" Margin="0,0,0,5"/>
  <TextBox x:Name="crUser" Tag="DOMENA\login albo login@domena"/>
  <TextBlock Text="Hasło" Foreground="#8791A5" FontSize="12" Margin="0,12,0,5"/>
  <PasswordBox x:Name="crPass"/>
</StackPanel>
'@
    $w = New-Dialog -Title 'Poświadczenia' -Subtitle $Message -Body $body -Icon 'E72E' -OkText 'Użyj' -Width 470 -Validate {
        param($w)
        if ([string]::IsNullOrWhiteSpace($w.FindName('crUser').Text)) { Show-Warning 'Podaj nazwę użytkownika.'; return $false }
        return $true
    }
    $user = $w.FindName('crUser')
    $user.Text = $UserName
    $w.add_ContentRendered({
            param($s, $e)
            if ($s.FindName('crUser').Text) { [void]$s.FindName('crPass').Focus() } else { [void]$s.FindName('crUser').Focus() }
        })
    if (-not (Invoke-Dialog $w)) { return $null }
    $secure = $w.FindName('crPass').SecurePassword.Copy()
    $secure.MakeReadOnly()
    return (New-Object System.Management.Automation.PSCredential($user.Text.Trim(), $secure))
}

function Show-InputDialog {
    # Zwraca wpisany tekst albo $null. -Validate { param($text) } zwraca opis błędu albo pusty tekst.
    param([string]$Title, [string]$Prompt, [string]$Default = '', [string]$Placeholder = '', [switch]$Multiline, [scriptblock]$Validate, [string]$Icon = 'E70F')
    $body = if ($Multiline) {
        '<TextBox x:Name="inText" Style="{StaticResource MultiText}" MinHeight="220"/>'
    }
    else { '<TextBox x:Name="inText"/>' }
    $w = New-Dialog -Title $Title -Subtitle $Prompt -Body $body -Icon $Icon -Width $(if ($Multiline) { 620 } else { 480 }) -Height $(if ($Multiline) { 470 } else { 0 }) -Resizable:$Multiline -Validate {
        param($w)
        $check = $w.Tag.Check
        if ($check) {
            $problem = [string](& $check $w.FindName('inText').Text)
            if ($problem) { Show-Warning $problem; return $false }
        }
        return $true
    }
    $w.Tag.Check = $Validate
    $box = $w.FindName('inText')
    $box.Text = $Default
    if ($Placeholder) { $box.Tag = $Placeholder }
    if ($Multiline) { $w.FindName('btnOk').IsDefault = $false }
    $w.add_ContentRendered({ param($s, $e) $b = $s.FindName('inText'); [void]$b.Focus(); $b.SelectAll() })
    if (-not (Invoke-Dialog $w)) { return $null }
    return $box.Text
}

function Show-PasswordDialog {
    <#
        Ustawienia resetu hasła. Zwraca hashtablę albo $null:
        Mode = 'Generate' (losowe, osobne dla każdego konta) | 'Manual' (Password: SecureString),
        Length, MustChange, Unlock
    #>
    param([string]$Subtitle = 'Nowe hasło dla zaznaczonych kont.', [int]$Count = 1)
    $body = @'
<StackPanel>
  <Border Style="{StaticResource SegmentHost}" HorizontalAlignment="Left">
    <StackPanel Orientation="Horizontal">
      <RadioButton x:Name="pwGen" Style="{StaticResource SegmentRadio}" GroupName="pwMode" Content="Wygeneruj losowe" IsChecked="True"/>
      <RadioButton x:Name="pwManual" Style="{StaticResource SegmentRadio}" GroupName="pwMode" Content="Wpisz hasło"/>
    </StackPanel>
  </Border>
  <StackPanel x:Name="pwGenPanel" Margin="0,14,0,0">
    <StackPanel Orientation="Horizontal">
      <TextBlock Text="Długość" Foreground="#8791A5" VerticalAlignment="Center" Margin="0,0,10,0"/>
      <ComboBox x:Name="pwLength" Width="90"/>
    </StackPanel>
    <TextBlock Foreground="#7B8496" FontSize="12" TextWrapping="Wrap" Margin="0,8,0,0"
               Text="Każde konto dostanie inne hasło. Hasła pojawią się w tabeli wyników jako wartości poufne (podgląd, kopiowanie, eksport)."/>
  </StackPanel>
  <StackPanel x:Name="pwManualPanel" Margin="0,14,0,0" Visibility="Collapsed">
    <TextBlock Text="Hasło" Foreground="#8791A5" FontSize="12" Margin="0,0,0,5"/>
    <PasswordBox x:Name="pw1"/>
    <TextBlock Text="Powtórz hasło" Foreground="#8791A5" FontSize="12" Margin="0,10,0,5"/>
    <PasswordBox x:Name="pw2"/>
  </StackPanel>
  <Border Height="1" Background="#242B36" Margin="0,16,0,12"/>
  <CheckBox x:Name="pwMustChange" Content="Wymagaj zmiany hasła przy następnym logowaniu" IsChecked="True" Margin="0,0,0,8"/>
  <CheckBox x:Name="pwUnlock" Content="Odblokuj konto, jeśli jest zablokowane" IsChecked="True"/>
</StackPanel>
'@
    $w = New-Dialog -Title 'Reset hasła' -Subtitle $Subtitle -Body $body -Icon 'E8D7' -OkText 'Resetuj hasło' -Width 500 -Validate {
        param($w)
        if ($w.FindName('pwManual').IsChecked -eq $true) {
            $p1 = $w.FindName('pw1').Password
            if (-not $p1) { Show-Warning 'Hasło nie może być puste.'; return $false }
            if ($p1 -cne $w.FindName('pw2').Password) { Show-Warning 'Hasła nie są identyczne.'; return $false }
        }
        return $true
    }
    $len = $w.FindName('pwLength')
    foreach ($n in 12, 14, 16, 20, 24) { [void]$len.Items.Add([string]$n) }
    $len.SelectedIndex = 2
    $toggle = {
        param($s, $e)
        $win = [System.Windows.Window]::GetWindow($s)
        $manual = ($win.FindName('pwManual').IsChecked -eq $true)
        $win.FindName('pwManualPanel').Visibility = if ($manual) { 'Visible' } else { 'Collapsed' }
        $win.FindName('pwGenPanel').Visibility = if ($manual) { 'Collapsed' } else { 'Visible' }
    }
    $w.FindName('pwGen').add_Checked($toggle)
    $w.FindName('pwManual').add_Checked($toggle)
    if (-not (Invoke-Dialog $w)) { return $null }
    $result = @{
        Mode       = 'Generate'
        Length     = [int]$len.SelectedItem
        MustChange = ($w.FindName('pwMustChange').IsChecked -eq $true)
        Unlock     = ($w.FindName('pwUnlock').IsChecked -eq $true)
        Password   = $null
    }
    if ($w.FindName('pwManual').IsChecked -eq $true) {
        $result.Mode = 'Manual'
        $secure = $w.FindName('pw1').SecurePassword.Copy()
        $secure.MakeReadOnly()
        $result.Password = $secure
    }
    return $result
}

function Read-NewPassword {
    # Nowe hasło wpisane dwukrotnie; zwraca SecureString albo $null
    param([string]$Subtitle = 'Podaj nowe hasło.', [string]$Title = 'Nowe hasło')
    $body = @'
<StackPanel>
  <TextBlock Text="Hasło" Foreground="#8791A5" FontSize="12" Margin="0,0,0,5"/>
  <PasswordBox x:Name="np1"/>
  <TextBlock Text="Powtórz hasło" Foreground="#8791A5" FontSize="12" Margin="0,10,0,5"/>
  <PasswordBox x:Name="np2"/>
</StackPanel>
'@
    $w = New-Dialog -Title $Title -Subtitle $Subtitle -Body $body -Icon 'E8D7' -OkText 'Ustaw hasło' -Width 470 -Validate {
        param($w)
        $p1 = $w.FindName('np1').Password
        if (-not $p1) { Show-Warning 'Hasło nie może być puste.'; return $false }
        if ($p1 -cne $w.FindName('np2').Password) { Show-Warning 'Hasła nie są identyczne.'; return $false }
        return $true
    }
    $w.add_ContentRendered({ param($s, $e) [void]$s.FindName('np1').Focus() })
    if (-not (Invoke-Dialog $w)) { return $null }
    $secure = $w.FindName('np1').SecurePassword.Copy()
    $secure.MakeReadOnly()
    return $secure
}

function Show-TextDialog {
    param([string]$Title, [string]$Text, [string]$Subtitle = '')
    $body = '<TextBox x:Name="txText" Style="{StaticResource MultiText}" IsReadOnly="True" TextWrapping="NoWrap"/>'
    $w = New-Dialog -Title $Title -Subtitle $Subtitle -Body $body -Icon 'E8A5' -OkText 'Zamknij' -NoCancel -Width 920 -Height 640 -Resizable
    $box = $w.FindName('txText')
    $box.Text = ($Text -replace "`r?`n", "`r`n")
    $copy = New-PlainButton -Text 'Kopiuj' -Icon 'E8C8'
    $copy.add_Click({
            param($s, $e)
            $t = [System.Windows.Window]::GetWindow($s).FindName('txText').Text
            if ($t) { Set-ClipboardText $t; Show-Toast 'Skopiowano do schowka.' 'ok' }
        })
    [void]$w.FindName('dlgExtra').Children.Add($copy)
    [void](Invoke-Dialog $w)
}

function ConvertTo-LdapFilterValue {
    # RFC 4515: znaki specjalne w wartościach filtra LDAP
    param([string]$Text)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        switch ($ch) {
            '\' { [void]$sb.Append('\5c') }
            '*' { [void]$sb.Append('\2a') }
            '(' { [void]$sb.Append('\28') }
            ')' { [void]$sb.Append('\29') }
            ([char]0) { [void]$sb.Append('\00') }
            default { [void]$sb.Append($ch) }
        }
    }
    return $sb.ToString()
}

function Get-OuList {
    # Jednostki organizacyjne i kontenery domeny: @{ Domain = 'contoso.local'; RootDN = 'DC=...'; Items = @(DN...) }
    Import-AdModule
    $ad = Get-AdSplat
    $domain = Get-ADDomain @ad
    $dns = New-Object System.Collections.ArrayList
    foreach ($ou in @(Get-ADOrganizationalUnit -Filter * @ad)) { [void]$dns.Add([string]$ou.DistinguishedName) }
    foreach ($c in @($domain.ComputersContainer, $domain.UsersContainer)) { if ($c) { [void]$dns.Add([string]$c) } }
    return @{ Domain = [string]$domain.DNSRoot; RootDN = [string]$domain.DistinguishedName; Items = @($dns) }
}

function Select-OrganizationalUnit {
    # Wybór OU z drzewa domeny. Zwraca DN, '' (cała domena - tylko z -AllowDomainRoot) albo $null (anulowano).
    param([string]$Title = 'Wybierz jednostkę organizacyjną', [string]$Selected = '', [switch]$AllowDomainRoot)
    $data = Invoke-WithWaitCursor { Get-OuList }
    if (-not $data) { return $null }
    $body = @'
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <TextBox x:Name="ouSearch" Tag="Szukaj jednostki organizacyjnej…" Margin="0,0,0,10"/>
  <Border Grid.Row="1" Style="{StaticResource Card}" Padding="4">
    <TreeView x:Name="ouTree" BorderThickness="0" Background="Transparent"/>
  </Border>
  <TextBlock x:Name="ouPath" Grid.Row="2" Foreground="#7B8496" FontSize="12" TextWrapping="Wrap" Margin="2,10,0,0"/>
</Grid>
'@
    $subtitle = if ($AllowDomainRoot) { 'Zaznacz jednostkę organizacyjną albo korzeń domeny (cała domena).' } else { 'Zaznacz docelową jednostkę organizacyjną.' }
    $w = New-Dialog -Title $Title -Subtitle $subtitle -Body $body -Icon 'E8B7' -OkText 'Wybierz' -Width 560 -Height 680 -Resizable -Validate {
        param($w)
        $item = $w.FindName('ouTree').SelectedItem
        if (-not $item) { Show-Warning 'Zaznacz jednostkę organizacyjną.'; return $false }
        if (-not $w.Tag.AllowRoot -and [string]$item.Tag -eq $w.Tag.RootDN) { Show-Warning 'Wybierz jednostkę organizacyjną (nie korzeń domeny).'; return $false }
        return $true
    }
    $w.Tag.AllowRoot = [bool]$AllowDomainRoot
    $w.Tag.RootDN = $data.RootDN
    $w.Tag.Data = $data
    $w.Tag.Selected = $Selected
    $tree = $w.FindName('ouTree')
    Update-OuTree -Window $w -Filter ''
    $tree.add_SelectedItemChanged({
            param($s, $e)
            $win = [System.Windows.Window]::GetWindow($s)
            $item = $s.SelectedItem
            $win.FindName('ouPath').Text = if ($item) { [string]$item.Tag } else { '' }
        })
    $tree.add_MouseDoubleClick({
            param($s, $e)
            if ($s.SelectedItem -and $e.OriginalSource -is [System.Windows.FrameworkElement]) {
                $ok = [System.Windows.Window]::GetWindow($s).FindName('btnOk')
                $ok.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, $ok)))
            }
        })
    $w.FindName('ouSearch').add_TextChanged({
            param($s, $e)
            Update-OuTree -Window ([System.Windows.Window]::GetWindow($s)) -Filter $s.Text.Trim()
        })
    if (-not (Invoke-Dialog $w)) { return $null }
    $dn = [string]$tree.SelectedItem.Tag
    if ($dn -eq $data.RootDN) { return '' }
    return $dn
}

function New-TreeHeader([string]$Text, [string]$Icon, [string]$Color = '#8CB0FF') {
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Orientation = 'Horizontal'
    $g = New-GlyphBlock -Code $Icon -Size 13 -Color $Color
    $g.Margin = '0,0,8,0'
    [void]$sp.Children.Add($g)
    $t = New-Object System.Windows.Controls.TextBlock
    $t.Text = $Text
    [void]$sp.Children.Add($t)
    return $sp
}

function Update-OuTree {
    # Buduje drzewo; przy filtrze pokazuje pasujące OU wraz z przodkami (rozwinięte)
    param($Window, [string]$Filter)
    $data = $Window.Tag.Data
    $tree = $Window.FindName('ouTree')
    $tree.Items.Clear()
    $root = New-Object System.Windows.Controls.TreeViewItem
    $root.Header = New-TreeHeader -Text $data.Domain -Icon 'E774' -Color '#5EE3AE'
    $root.Tag = $data.RootDN
    $root.IsExpanded = $true
    [void]$tree.Items.Add($root)
    $rootKey = $data.RootDN.ToLowerInvariant()
    $map = @{ $rootKey = $root }
    $items = @($data.Items)
    if ($Filter) {
        # Pasujące + wszyscy ich przodkowie
        $keep = @{}
        foreach ($dn in $items) {
            if ((Get-RdnValue $dn) -like "*$Filter*") {
                $cur = $dn
                while ($cur -and $cur.ToLowerInvariant() -ne $rootKey -and -not $keep.ContainsKey($cur.ToLowerInvariant())) {
                    $keep[$cur.ToLowerInvariant()] = $true
                    $cur = Get-ParentDN $cur
                }
            }
        }
        $items = @($items | Where-Object { $keep.ContainsKey($_.ToLowerInvariant()) })
    }
    $sorted = $items | Where-Object { $_ } | Sort-Object { @($_ -split '(?<!\\),').Count }, { Get-RdnValue $_ }
    $selected = [string]$Window.Tag.Selected
    foreach ($dn in $sorted) {
        $parent = $map[(Get-ParentDN $dn).ToLowerInvariant()]
        if (-not $parent) { $parent = $root }
        $node = New-Object System.Windows.Controls.TreeViewItem
        $isContainer = $dn -match '^CN='
        $node.Header = New-TreeHeader -Text (Get-RdnValue $dn) -Icon $(if ($isContainer) { 'E8B7' } else { 'ED41' }) -Color $(if ($isContainer) { '#8791A5' } else { '#FFC46B' })
        $node.Tag = $dn
        if ($Filter) { $node.IsExpanded = $true }
        [void]$parent.Items.Add($node)
        $map[$dn.ToLowerInvariant()] = $node
        if ($selected -and $dn -eq $selected) {
            $node.IsSelected = $true
            $p = $parent
            while ($p -is [System.Windows.Controls.TreeViewItem]) { $p.IsExpanded = $true; $p = $p.Parent }
        }
    }
    if (-not $tree.SelectedItem -and -not $Filter) { $root.IsSelected = $true }
}

function Find-AdGroup {
    # Wyszukiwanie grup (nazwa, sAMAccountName, opis); zwraca obiekty z Name, Sam, Scope, Category, Description, DN
    param([string]$Text)
    Import-AdModule
    $ad = Get-AdSplat
    $v = ConvertTo-LdapFilterValue $Text
    $filter = if ($v) { "(&(objectCategory=group)(|(name=*$v*)(sAMAccountName=*$v*)(description=*$v*)))" } else { '(objectCategory=group)' }
    $groups = @(Get-ADGroup -LDAPFilter $filter -Properties Description -ResultSetSize 300 @ad)
    return @($groups | Sort-Object Name | ForEach-Object {
            [pscustomobject]@{
                Name        = [string]$_.Name
                Sam         = [string]$_.SamAccountName
                Scope       = [string]$_.GroupScope
                Category    = [string]$_.GroupCategory
                Description = [string]$_.Description
                DN          = [string]$_.DistinguishedName
            }
        })
}

function Select-AdGroups {
    # Wybór grup AD (wielokrotny, z wyszukiwaniem). Zwraca tablicę obiektów grup albo $null.
    param([string]$Title = 'Wybierz grupy', [string]$Subtitle = 'Wyszukaj grupy i zaznacz je na liście (kliknięcie zaznacza / odznacza).')
    $body = @'
<Grid>
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*"/>
    <RowDefinition Height="Auto"/>
  </Grid.RowDefinitions>
  <Grid Margin="0,0,0,10">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <TextBox x:Name="grSearch" Tag="Nazwa grupy, fragment nazwy lub opisu…"/>
    <Button x:Name="grFind" Grid.Column="1" Content="Szukaj" Margin="8,0,0,0" MinWidth="90"/>
  </Grid>
  <Border Grid.Row="1" Style="{StaticResource Card}" Padding="4">
    <Grid>
      <ListBox x:Name="grList" SelectionMode="Multiple">
        <ListBox.ItemTemplate>
          <DataTemplate>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <CheckBox IsHitTestVisible="False" Focusable="False" Margin="0,0,10,0" VerticalAlignment="Center"
                        IsChecked="{Binding IsSelected, Mode=OneWay, RelativeSource={RelativeSource AncestorType=ListBoxItem}}"/>
              <StackPanel Grid.Column="1">
                <TextBlock Text="{Binding [Name]}" Foreground="#E4E8EF"/>
                <TextBlock Text="{Binding [Description]}" Foreground="#7B8496" FontSize="11.5" TextTrimming="CharacterEllipsis"/>
              </StackPanel>
              <Border Grid.Column="2" Style="{StaticResource Chip}" Margin="8,0,0,0">
                <TextBlock Text="{Binding [Scope]}" Foreground="#8791A5" FontSize="11"/>
              </Border>
            </Grid>
          </DataTemplate>
        </ListBox.ItemTemplate>
      </ListBox>
      <TextBlock x:Name="grEmpty" Text="Wpisz nazwę i kliknij Szukaj" Foreground="#5E6779" HorizontalAlignment="Center" VerticalAlignment="Center" IsHitTestVisible="False"/>
    </Grid>
  </Border>
  <TextBlock x:Name="grSelected" Grid.Row="2" Foreground="#8791A5" FontSize="12" Margin="2,10,0,0" TextWrapping="Wrap" Text="Wybrane: brak"/>
</Grid>
'@
    $w = New-Dialog -Title $Title -Subtitle $Subtitle -Body $body -Icon 'E902' -OkText 'Wybierz' -Width 620 -Height 640 -Resizable -Validate {
        param($w)
        if ($w.Tag.Chosen.Count -eq 0) { Show-Warning 'Nie zaznaczono żadnej grupy.'; return $false }
        return $true
    }
    $w.Tag.Chosen = [ordered]@{}
    $w.Tag.Loading = $false
    $table = New-Object System.Data.DataTable
    foreach ($c in 'Name', 'Sam', 'Scope', 'Category', 'Description', 'DN') { [void]$table.Columns.Add($c, [string]) }
    $w.Tag.Table = $table
    $list = $w.FindName('grList')
    $list.ItemsSource = $table.DefaultView
    $search = {
        param($s, $e)
        $win = [System.Windows.Window]::GetWindow($s)
        $text = $win.FindName('grSearch').Text.Trim()
        $found = @()
        try { $found = @(Invoke-WithWaitCursor { Find-AdGroup -Text $text }) }
        catch { Show-Error 'Wyszukiwanie grup nie powiodło się.' $_; return }
        $win.Tag.Loading = $true
        try {
            $t = $win.Tag.Table
            $t.Rows.Clear()
            foreach ($g in $found) { [void]$t.Rows.Add($g.Name, $g.Sam, $g.Scope, $g.Category, $g.Description, $g.DN) }
            $lb = $win.FindName('grList')
            foreach ($drv in $t.DefaultView) {
                if ($win.Tag.Chosen.Contains([string]$drv['DN'])) { [void]$lb.SelectedItems.Add($drv) }
            }
            $win.FindName('grEmpty').Text = if ($t.Rows.Count -eq 0) { 'Nie znaleziono grup' } else { '' }
        }
        finally { $win.Tag.Loading = $false }
    }
    $w.FindName('grFind').add_Click($search)
    $w.FindName('grSearch').add_KeyDown({
            param($s, $e)
            if ($e.Key -eq [System.Windows.Input.Key]::Return) {
                $e.Handled = $true
                $b = [System.Windows.Window]::GetWindow($s).FindName('grFind')
                $b.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, $b)))
            }
        })
    $list.add_SelectionChanged({
            param($s, $e)
            $win = [System.Windows.Window]::GetWindow($s)
            if ($win.Tag.Loading) { return }
            foreach ($drv in $e.AddedItems) {
                $win.Tag.Chosen[[string]$drv['DN']] = [pscustomobject]@{
                    Name = [string]$drv['Name']; Sam = [string]$drv['Sam']; Scope = [string]$drv['Scope']
                    Category = [string]$drv['Category']; Description = [string]$drv['Description']; DN = [string]$drv['DN']
                }
            }
            foreach ($drv in $e.RemovedItems) { $win.Tag.Chosen.Remove([string]$drv['DN']) }
            $names = @($win.Tag.Chosen.Values | ForEach-Object { $_.Name })
            $win.FindName('grSelected').Text = if ($names.Count -gt 0) { 'Wybrane (' + $names.Count + '): ' + ($names -join ', ') } else { 'Wybrane: brak' }
        })
    $w.FindName('btnOk').IsDefault = $false
    if (-not (Invoke-Dialog $w)) { return $null }
    return @($w.Tag.Chosen.Values)
}

function Invoke-WithWaitCursor {
    # Krótka operacja w wątku okna z kursorem oczekiwania
    param([Parameter(Mandatory)][scriptblock]$Action)
    $previous = [System.Windows.Input.Mouse]::OverrideCursor
    [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    try { & $Action }
    finally { [System.Windows.Input.Mouse]::OverrideCursor = $previous }
}

function Show-SettingsDialog {
    $body = @'
<StackPanel>
  <TextBlock Text="Kontroler domeny" Foreground="#8791A5" FontSize="12" Margin="0,0,0,5"/>
  <TextBox x:Name="stDc" Tag="puste = wybór automatyczny"/>
  <Grid Margin="0,14,0,0">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="14"/>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="14"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>
    <StackPanel>
      <TextBlock Text="Równoległe operacje" Foreground="#8791A5" FontSize="12" Margin="0,0,0,5"/>
      <TextBox x:Name="stThrottle" HorizontalContentAlignment="Right"/>
    </StackPanel>
    <StackPanel Grid.Column="2">
      <TextBlock Text="Limit połączenia (s)" Foreground="#8791A5" FontSize="12" Margin="0,0,0,5"/>
      <TextBox x:Name="stTimeout" HorizontalContentAlignment="Right"/>
    </StackPanel>
    <StackPanel Grid.Column="4">
      <TextBlock Text="Nieaktywność (dni)" Foreground="#8791A5" FontSize="12" Margin="0,0,0,5"/>
      <TextBox x:Name="stDays" HorizontalContentAlignment="Right"/>
    </StackPanel>
  </Grid>
  <TextBlock Foreground="#7B8496" FontSize="12" TextWrapping="Wrap" Margin="0,8,0,0"
             Text="Równoległe operacje: 1–64. Limit połączenia WinRM: 5–300 s. Nieaktywność to domyślna wartość raportów kont i profili."/>
  <Border Height="1" Background="#242B36" Margin="0,16,0,12"/>
  <WrapPanel>
    <Button x:Name="stLogs" Margin="0,0,8,6"/>
    <Button x:Name="stData" Margin="0,0,8,6"/>
    <Button x:Name="stPlugins" Margin="0,0,8,6"/>
  </WrapPanel>
  <TextBlock x:Name="stInfo" Foreground="#5E6779" FontSize="11.5" TextWrapping="Wrap" Margin="0,8,0,0"/>
</StackPanel>
'@
    $w = New-Dialog -Title 'Ustawienia' -Subtitle 'Połączenia, wydajność i pliki programu.' -Body $body -Icon 'E713' -OkText 'Zapisz' -Width 560 -Validate {
        param($w)
        foreach ($pair in @(@('stThrottle', 1, 64, 'Równoległe operacje'), @('stTimeout', 5, 300, 'Limit połączenia'), @('stDays', 1, 3650, 'Nieaktywność'))) {
            $v = 0
            if (-not [int]::TryParse($w.FindName($pair[0]).Text.Trim(), [ref]$v) -or $v -lt $pair[1] -or $v -gt $pair[2]) {
                Show-Warning ('{0}: podaj liczbę z zakresu {1}–{2}.' -f $pair[3], $pair[1], $pair[2])
                return $false
            }
        }
        return $true
    }
    $w.FindName('stDc').Text = [string]$script:Settings.DomainController
    $w.FindName('stThrottle').Text = [string]$script:Settings.ThrottleLimit
    $w.FindName('stTimeout').Text = [string]$script:Settings.TimeoutSec
    $w.FindName('stDays').Text = [string]$script:Settings.InactiveDays
    $w.FindName('stLogs').Content = New-IconContent -Text 'Folder dziennika' -Icon 'E838'
    $w.FindName('stData').Content = New-IconContent -Text 'Folder ustawień' -Icon 'E838'
    $w.FindName('stPlugins').Content = New-IconContent -Text 'Folder modułów' -Icon 'EA86'
    $w.FindName('stInfo').Text = "Domain Ops $($script:AppVersion) • PowerShell $($PSVersionTable.PSVersion) • moduły własne: $($script:App.ModulesDir)"
    $w.FindName('stLogs').add_Click({ Open-Folder $script:App.LogDir })
    $w.FindName('stData').add_Click({ Open-Folder $script:App.DataDir })
    $w.FindName('stPlugins').add_Click({ if ($script:App.ModulesDir) { Open-Folder $script:App.ModulesDir } })
    if (-not (Invoke-Dialog $w)) { return $false }
    $script:Settings.DomainController = $w.FindName('stDc').Text.Trim()
    $script:Settings.TimeoutSec = [int]$w.FindName('stTimeout').Text.Trim()
    $script:Settings.InactiveDays = [int]$w.FindName('stDays').Text.Trim()
    Set-EngineThrottle ([int]$w.FindName('stThrottle').Text.Trim())
    Export-Settings
    return $true
}

function Open-Folder([string]$Path) {
    try {
        if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
        Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Path)
    }
    catch { Show-Error "Nie można otworzyć folderu $Path." $_ }
}

function Show-Toast {
    # Krótkie powiadomienie w prawym dolnym rogu okna (znika samo)
    param([Parameter(Mandatory)][string]$Text, [ValidateSet('info', 'ok', 'warn', 'crit')][string]$Tone = 'info', [int]$Seconds = 4)
    $toastHost = $script:UI.Controls['toastHost']
    if (-not $toastHost) { return }
    try {
        $toneDef = $script:DialogTones[$Tone]
        $xaml = @'
<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Background="#1E2530" BorderBrush="#2F3846" BorderThickness="1"
        CornerRadius="9" Padding="14,10" Margin="0,8,0,0" MaxWidth="420" HorizontalAlignment="Right">
  <Border.Effect><DropShadowEffect BlurRadius="18" ShadowDepth="3" Opacity="0.45" Color="Black"/></Border.Effect>
  <StackPanel Orientation="Horizontal">
    <TextBlock Name="icon" Style="{StaticResource Glyph}" FontSize="15" Margin="0,0,10,0"/>
    <TextBlock Name="text" TextWrapping="Wrap" MaxWidth="360" Foreground="#E4E8EF" VerticalAlignment="Center"/>
  </StackPanel>
</Border>
'@
        $toast = New-UiElement $xaml
        $icon = $toast.FindName('icon')
        $icon.Text = Get-Glyph $toneDef.Icon
        $icon.Foreground = Get-Brush $toneDef.Fore
        $toast.FindName('text').Text = $Text
        while ($toastHost.Children.Count -ge 4) { $toastHost.Children.RemoveAt(0) }
        [void]$toastHost.Children.Add($toast)
        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromSeconds([Math]::Max(2, $Seconds))
        $script:ToastTimers[$timer] = $toast
        $timer.add_Tick({
                param($s, $e)
                $s.Stop()
                $t = $null
                if ($script:ToastTimers.TryGetValue($s, [ref]$t)) {
                    [void]$script:ToastTimers.Remove($s)
                    $h = $script:UI.Controls['toastHost']
                    if ($h -and $h.Children.Contains($t)) { $h.Children.Remove($t) }
                }
            })
        $timer.Start()
    }
    catch { }
}
$script:ToastTimers = New-Object 'System.Collections.Generic.Dictionary[object,object]'
#endregion

#region Widok modułu i tabela wyników (DataTable + DataView + DataGrid)
# Kolumny tabeli tworzone są w kodzie w chwili pojawienia się nowej właściwości w wynikach
# (bez ponownego wiązania siatki - sortowanie i przewinięcie zostają). Kolumny ukryte zaczynają się od "__":
#   __search - tekst do filtrowania, __flag - kolor wiersza (crit/warn/muted), __tone - kolor "pigułki"
#   w kolumnach z $m.PillColumns (ok/warn/crit/info), __secret_<kolumna> - prawdziwa wartość poufna.
$script:ModuleViewXaml = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Grid.RowDefinitions>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="Auto"/>
    <RowDefinition Height="*" MinHeight="160"/>
  </Grid.RowDefinitions>

  <Grid Margin="0,0,0,16">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <Border Width="44" Height="44" CornerRadius="11" Background="#1A2640" BorderBrush="#2F4478" BorderThickness="1" VerticalAlignment="Top">
      <TextBlock x:Name="hdrIcon" Style="{StaticResource Glyph}" FontSize="20" Foreground="#8CB0FF" HorizontalAlignment="Center"/>
    </Border>
    <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
      <StackPanel Orientation="Horizontal">
        <TextBlock x:Name="hdrTitle" FontSize="20" FontWeight="SemiBold" Foreground="White"/>
        <Border Style="{StaticResource Chip}" Margin="12,2,0,0">
          <TextBlock x:Name="hdrCategory" FontSize="11" Foreground="#8791A5"/>
        </Border>
      </StackPanel>
      <TextBlock x:Name="hdrDesc" Foreground="#8791A5" TextWrapping="Wrap" Margin="0,4,0,0" MaxWidth="1000" HorizontalAlignment="Left"/>
    </StackPanel>
    <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Top">
      <Border x:Name="busyChip" Style="{StaticResource Chip}" Background="#1A2640" Padding="12,5" Visibility="Collapsed" VerticalAlignment="Center">
        <StackPanel Orientation="Horizontal">
          <Ellipse Width="8" Height="8" Fill="#4C7DF0" Margin="0,0,8,0" VerticalAlignment="Center"/>
          <TextBlock x:Name="busyText" Text="Trwa operacja…" Foreground="#8CC0FF" FontSize="12"/>
        </StackPanel>
      </Border>
      <Button x:Name="btnCollapse" Style="{StaticResource GhostButton}" Padding="8,4" MinHeight="28" ToolTip="Zwiń / rozwiń panel parametrów"/>
    </StackPanel>
  </Grid>

  <Border x:Name="paramsCard" Grid.Row="1" Style="{StaticResource Card}" Padding="16,14,10,8" Margin="0,0,0,12" Visibility="Collapsed">
    <StackPanel x:Name="paramsStack"/>
  </Border>

  <UniformGrid x:Name="statsGrid" Grid.Row="2" Rows="1" Margin="0,0,-10,12" Visibility="Collapsed"/>

  <Grid Grid.Row="3" Margin="2,0,0,8">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="Auto"/>
      <ColumnDefinition Width="Auto"/>
    </Grid.ColumnDefinitions>
    <TextBlock Text="Wyniki" FontSize="14" FontWeight="SemiBold" VerticalAlignment="Center"/>
    <Border Grid.Column="1" Style="{StaticResource Chip}" Margin="10,0,0,0">
      <TextBlock x:Name="countText" Text="0" Foreground="#AEB6C4" FontSize="11.5"/>
    </Border>
    <TextBlock x:Name="resultHint" Grid.Column="2" Foreground="#5E6779" FontSize="12" VerticalAlignment="Center" Margin="6,0,12,0" TextTrimming="CharacterEllipsis"/>
    <TextBox x:Name="filterBox" Grid.Column="3" Width="260" Tag="Filtruj wyniki (Ctrl+F)…" Margin="0,0,8,0"/>
    <CheckBox x:Name="chkReveal" Grid.Column="4" Content="Pokaż poufne" Margin="4,0,12,0" Visibility="Collapsed"/>
    <Button x:Name="btnDetail" Grid.Column="5" Style="{StaticResource GhostButton}" ToolTip="Panel szczegółów wiersza" Margin="0,0,4,0"/>
    <Button x:Name="btnCopy" Grid.Column="6" Style="{StaticResource GhostButton}" ToolTip="Kopiuj zaznaczone wiersze lub całą tabelę (format Excel)" Margin="0,0,4,0"/>
    <Button x:Name="btnExport" Grid.Column="7" ToolTip="Eksport widocznych wierszy do CSV lub raportu HTML"/>
  </Grid>

  <Border Grid.Row="4" Style="{StaticResource Card}" Padding="0">
    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*" MinWidth="300"/>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition x:Name="detailCol" Width="340" MinWidth="0"/>
      </Grid.ColumnDefinitions>
      <DataGrid x:Name="grid" Style="{StaticResource DarkGrid}" AutoGenerateColumns="False" Margin="1,1,1,6" CanUserReorderColumns="True"/>
      <StackPanel x:Name="emptyState" HorizontalAlignment="Center" VerticalAlignment="Center" IsHitTestVisible="False" Margin="20,40,20,20">
        <Border Width="64" Height="64" CornerRadius="32" Background="#1A2029" HorizontalAlignment="Center">
          <TextBlock x:Name="emptyIcon" Style="{StaticResource Glyph}" FontSize="26" Foreground="#4A5568" HorizontalAlignment="Center"/>
        </Border>
        <TextBlock x:Name="emptyText" Text="Brak wyników" FontSize="15" FontWeight="SemiBold" Foreground="#AEB6C4" HorizontalAlignment="Center" Margin="0,14,0,0"/>
        <TextBlock x:Name="emptyHint" Foreground="#5E6779" HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" MaxWidth="440" Margin="0,5,0,0"/>
      </StackPanel>
      <GridSplitter x:Name="detailSplit" Grid.Column="1" Width="5" HorizontalAlignment="Stretch" ResizeBehavior="PreviousAndNext"/>
      <Border x:Name="detailPane" Grid.Column="2" BorderBrush="#242B36" BorderThickness="1,0,0,0">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <Grid Margin="14,12,8,6">
            <TextBlock Text="SZCZEGÓŁY WIERSZA" Foreground="#5E6779" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
            <Button x:Name="btnDetailCopy" Style="{StaticResource GhostButton}" HorizontalAlignment="Right" Padding="6,2" MinHeight="24" ToolTip="Kopiuj szczegóły"/>
          </Grid>
          <TextBox x:Name="detailText" Grid.Row="1" Style="{StaticResource ReadOnlyText}" FontFamily="Segoe UI" FontSize="12.5" Margin="8,0,4,8" Padding="6,0"/>
        </Grid>
      </Border>
    </Grid>
  </Border>
</Grid>
'@

$script:PillTones = @{
    ok   = @('#15291F', '#5EE3AE')
    warn = @('#2E2616', '#FFC46B')
    crit = @('#2E1A1E', '#FF7A86')
    info = @('#1A2640', '#8CC0FF')
}

function New-ModuleView {
    # Tworzy wizualną część modułu i podłącza zdarzenia tabeli
    param([Parameter(Mandatory)][hashtable]$Module)
    $root = New-UiElement $script:ModuleViewXaml
    $Module.Root = $root
    foreach ($n in 'paramsCard', 'paramsStack', 'statsGrid', 'busyChip', 'busyText', 'countText', 'resultHint', 'filterBox', 'chkReveal',
        'btnDetail', 'btnCopy', 'btnExport', 'grid', 'emptyState', 'emptyIcon', 'emptyText', 'emptyHint', 'detailCol', 'detailSplit',
        'detailPane', 'detailText', 'btnDetailCopy', 'btnCollapse') {
        $Module.View_[$n] = $root.FindName($n)
    }
    $Module.ParamsCard = $Module.View_['paramsCard']
    $Module.ParamsStack = $Module.View_['paramsStack']
    $Module.StatsGrid = $Module.View_['statsGrid']
    $Module.Grid = $Module.View_['grid']
    $Module.FilterBox = $Module.View_['filterBox']

    $root.FindName('hdrIcon').Text = Get-Glyph $Module.Icon
    $root.FindName('hdrTitle').Text = $Module.Title
    $root.FindName('hdrCategory').Text = $(if ($Module.WorkspaceTitle) { $Module.WorkspaceTitle + '  •  ' + $Module.Category } else { $Module.Category })
    $desc = $root.FindName('hdrDesc')
    if ($Module.Description) { $desc.Text = $Module.Description } else { $desc.Visibility = 'Collapsed' }

    $v = $Module.View_
    $v.btnDetail.Content = New-IconContent -Text '' -Icon 'E8A1'
    $v.btnCopy.Content = New-IconContent -Text '' -Icon 'E8C8'
    $v.btnExport.Content = New-IconContent -Text 'Eksport' -Icon 'EDE1'
    $v.btnDetailCopy.Content = New-IconContent -Text '' -Icon 'E8C8' -IconSize 11
    $v.btnCollapse.Content = New-IconContent -Text '' -Icon 'E70E' -IconSize 12
    $v.btnCollapse.Visibility = 'Collapsed'
    Register-ControlHandler -Control $v.btnCollapse -EventName 'Click' -Module $Module -Action {
        param($m)
        $m.ParamsCollapsed = -not $m.ParamsCollapsed
        Update-ParamsCollapse -Module $m
    }

    Register-ControlHandler -Control $v.filterBox -EventName 'TextChanged' -Module $Module -Action { param($m) Request-ResultFilter -Module $m }
    Register-ControlHandler -Control $v.btnCopy -EventName 'Click' -Module $Module -Action { param($m) Copy-ResultView -Module $m }
    Register-ControlHandler -Control $v.btnExport -EventName 'Click' -Module $Module -Action { param($m) Export-ResultView -Module $m }
    Register-ControlHandler -Control $v.btnDetail -EventName 'Click' -Module $Module -Action {
        param($m)
        $script:Settings.DetailVisible = -not $script:Settings.DetailVisible
        foreach ($other in $script:UI.Modules.Values) { if ($other.Root) { Update-DetailPane -Module $other } }
    }
    Register-ControlHandler -Control $v.btnDetailCopy -EventName 'Click' -Module $Module -Action {
        param($m)
        $t = $m.View_['detailText'].Text
        if ($t) { Set-ClipboardText $t; Show-Toast 'Skopiowano szczegóły wiersza.' 'ok' }
    }
    Register-ControlHandler -Control $v.chkReveal -EventName 'Checked' -Module $Module -Action { param($m) Set-SecretReveal -Module $m -Reveal $true }
    Register-ControlHandler -Control $v.chkReveal -EventName 'Unchecked' -Module $Module -Action { param($m) Set-SecretReveal -Module $m -Reveal $false }

    $grid = $Module.Grid
    $script:GridModules[$grid] = $Module
    $grid.add_SelectionChanged($script:GridEvents.SelectionChanged)
    $grid.add_MouseDoubleClick($script:GridEvents.MouseDoubleClick)
    $grid.add_ContextMenuOpening($script:GridEvents.ContextMenuOpening)
    $grid.ContextMenu = New-Object System.Windows.Controls.ContextMenu
    Reset-ResultTable -Module $Module
    Update-DetailPane -Module $Module
}

$script:GridModules = New-Object 'System.Collections.Generic.Dictionary[object,hashtable]'
$script:GridEvents = @{
    SelectionChanged   = {
        param($s, $e)
        try {
            $m = $null
            if ($script:GridModules.TryGetValue($s, [ref]$m)) { Update-DetailText -Module $m }
        }
        catch { }
    }
    MouseDoubleClick   = {
        param($s, $e)
        try {
            $m = $null
            if (-not $script:GridModules.TryGetValue($s, [ref]$m)) { return }
            # Tylko podwójne kliknięcie na wierszu (nie na nagłówku ani pasku przewijania)
            $dep = $e.OriginalSource
            while ($dep -and -not ($dep -is [System.Windows.Controls.DataGridRow])) {
                if ($dep -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $dep -is [System.Windows.Controls.Primitives.ScrollBar]) { return }
                if ($dep -is [System.Windows.Media.Visual] -or $dep -is [System.Windows.Media.Media3D.Visual3D]) { $dep = [System.Windows.Media.VisualTreeHelper]::GetParent($dep) }
                else { $dep = $null }
            }
            if (-not $dep) { return }
            $drv = $dep.Item
            if (-not ($drv -is [System.Data.DataRowView])) { return }
            if ($m.RowDoubleClick) { Invoke-UiAction -Module $m -Action $m.RowDoubleClick -Source $drv }
            else { Show-RowDetails -Module $m -Row $drv }
        }
        catch { Write-Log "Nie można wyświetlić szczegółów: $($_.Exception.Message)" 'ERROR' }
    }
    ContextMenuOpening = {
        param($s, $e)
        try {
            $m = $null
            if ($script:GridModules.TryGetValue($s, [ref]$m)) { Update-GridMenu -Module $m }
        }
        catch { }
    }
}

function New-MenuItem {
    param([string]$Text, [string]$Icon = '', [hashtable]$Module, [scriptblock]$Action, [switch]$Danger, [bool]$Enabled = $true)
    $mi = New-Object System.Windows.Controls.MenuItem
    $mi.Header = $Text
    if ($Icon) { $mi.Icon = New-GlyphBlock -Code $Icon -Size 13 -Color $(if ($Danger) { '#FF8A95' } else { '#8791A5' }) }
    if ($Danger) { $mi.Foreground = Get-Brush '#FF8A95' }
    $mi.IsEnabled = $Enabled
    if ($Action) { Register-ControlHandler -Control $mi -EventName 'Click' -Module $Module -Action $Action }
    return $mi
}

function New-MenuSeparator {
    $sep = New-Object System.Windows.Controls.Separator
    $sep.Style = Get-ThemeResource 'MenuSeparator'
    return $sep
}

function Update-GridMenu {
    # Menu kontekstowe budowane przy otwarciu: akcje modułu + kopiowanie/eksport
    param([hashtable]$Module)
    $menu = $Module.Grid.ContextMenu
    foreach ($old in @($menu.Items)) {
        [void]$script:MenuActions.Remove($old)
        [void]$script:Handlers.Remove($old)
    }
    $menu.Items.Clear()
    $rows = @(Get-SelectedResultRows -Module $Module)
    $hasRows = $rows.Count -gt 0
    foreach ($a in $Module.RowActions) {
        if ($a.Separator -and $menu.Items.Count -gt 0) { [void]$menu.Items.Add((New-MenuSeparator)) }
        $item = New-MenuItem -Text $a.Text -Icon $a.Icon -Module $Module -Danger:$a.Danger -Enabled ($hasRows -and -not $Module.Busy) -Action {
            param($m, $s)
            $entry = $null
            [void]$script:MenuActions.TryGetValue($s, [ref]$entry)
            if (-not $entry) { return }
            $selected = @(Get-SelectedResultRows -Module $m | Where-Object { [string](Get-ObjectValue $_ 'Status') -ne 'Błąd' })
            if ($selected.Count -eq 0) { Show-Warning 'Zaznacz wiersze w tabeli wyników.'; return }
            $null = & $entry $m $selected
        }
        $script:MenuActions[$item] = $a.Action
        [void]$menu.Items.Add($item)
    }
    if ($menu.Items.Count -gt 0) { [void]$menu.Items.Add((New-MenuSeparator)) }
    [void]$menu.Items.Add((New-MenuItem -Text 'Szczegóły wiersza' -Icon 'E8A1' -Module $Module -Enabled $hasRows -Action {
                param($m)
                $r = @(Get-SelectedResultRows -Module $m)
                if ($r.Count -gt 0) { Show-RowDetails -Module $m -Row $r[0] }
            }))
    [void]$menu.Items.Add((New-MenuItem -Text 'Kopiuj wartość komórki' -Icon 'E8C8' -Module $Module -Enabled $hasRows -Action {
                param($m)
                $cell = $m.Grid.CurrentCell
                if (-not $cell.Column -or -not ($cell.Item -is [System.Data.DataRowView])) { return }
                $name = [string]$cell.Column.SortMemberPath
                $value = Get-ExportValue -Module $m -Row $cell.Item -Column $name
                Set-ClipboardText ([string]$value)
                Show-Toast "Skopiowano: $name" 'ok' 2
            }))
    [void]$menu.Items.Add((New-MenuItem -Text 'Kopiuj zaznaczone wiersze' -Icon 'E8C8' -Module $Module -Enabled $hasRows -Action { param($m) Copy-ResultView -Module $m -SelectedOnly }))
    [void]$menu.Items.Add((New-MenuItem -Text 'Eksport…' -Icon 'EDE1' -Module $Module -Action { param($m) Export-ResultView -Module $m }))
}
$script:MenuActions = New-Object 'System.Collections.Generic.Dictionary[object,scriptblock]'

function Update-ParamsCollapse {
    param([hashtable]$Module)
    $v = $Module.View_
    $collapsed = [bool]$Module.ParamsCollapsed
    $v.paramsCard.Visibility = if ($collapsed -or $Module.ParamsStack.Children.Count -eq 0) { 'Collapsed' } else { 'Visible' }
    $v.statsGrid.Visibility = if ($collapsed -or $Module.StatsGrid.Children.Count -eq 0) { 'Collapsed' } else { 'Visible' }
    $v.btnCollapse.Content = New-IconContent -Text $(if ($collapsed) { 'Parametry' } else { '' }) -Icon $(if ($collapsed) { 'E70D' } else { 'E70E' }) -IconSize 12
}

function Complete-ModuleView {
    # Po zbudowaniu modułu: elementy zależne od ustawień nadanych w Build
    param([Parameter(Mandatory)][hashtable]$Module)
    $v = $Module.View_
    if ($Module.ParamsStack.Children.Count -gt 0) { $v.btnCollapse.Visibility = 'Visible' }
    if ($Module.SecretColumns.Count -gt 0) { $v.chkReveal.Visibility = 'Visible' }
    $v.emptyIcon.Text = Get-Glyph $(if ($Module.EmptyIcon) { $Module.EmptyIcon } else { $Module.Icon })
    if (-not $Module.EmptyHint -and $Module.PrimaryButton) {
        $label = $Module.PrimaryButton.Content
        if ($label -is [System.Windows.Controls.Panel]) { $label = $label.Children[$label.Children.Count - 1].Text }
        $Module.EmptyHint = switch ($Module.Target) {
            'Computer' { "Zaznacz komputery na liście po lewej i kliknij «$label» (F5)." }
            'User' { "Zaznacz konta na liście po lewej i kliknij «$label» (F5)." }
            default { "Kliknij «$label» (F5), aby pobrać dane." }
        }
    }
    if ($Module.ResultHint) { $v.resultHint.Text = $Module.ResultHint }
    elseif ($Module.RowActions.Count -gt 0) { $v.resultHint.Text = 'Prawy przycisk myszy na wierszach – akcje' }
    else { $v.resultHint.Text = 'Dwuklik na wierszu – szczegóły' }
    Update-ResultCount -Module $Module
}

function Reset-ResultTable {
    param([Parameter(Mandatory)][hashtable]$Module)
    $table = New-Object System.Data.DataTable 'Wyniki'
    foreach ($c in '__search', '__flag', '__tone') { [void]$table.Columns.Add($c, [string]) }
    $view = [System.Data.DataView]::new($table)
    $Module.Table = $table
    $Module.View = $view
    $Module.ColumnIndex = @{}
    $g = $Module.Grid
    if ($g) {
        $g.ItemsSource = $null
        $g.Columns.Clear()
        $g.ItemsSource = $view
    }
    Update-ResultFilter -Module $Module
}

function Add-GridColumn {
    param([hashtable]$Module, [string]$Name)
    $g = $Module.Grid
    if (-not $g -or $Module.ColumnIndex.ContainsKey($Name)) { return }
    $Module.ColumnIndex[$Name] = $true
    if ($Name.StartsWith('__') -or $Module.HiddenColumns -contains $Name) { return }
    $path = '[' + $Name + ']'
    if ($Module.PillColumns -contains $Name) {
        $xaml = @"
<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Border x:Name="pill" CornerRadius="9" Padding="9,2" HorizontalAlignment="Left" Background="#1E252F">
    <TextBlock x:Name="txt" Text="{Binding $(ConvertTo-XmlText $path)}" Foreground="#AEB6C4" FontSize="12" FontWeight="SemiBold"/>
  </Border>
  <DataTemplate.Triggers>
"@
        foreach ($tone in $script:PillTones.Keys) {
            $xaml += "<DataTrigger Binding=`"{Binding [__tone]}`" Value=`"$tone`"><Setter TargetName=`"pill`" Property=`"Background`" Value=`"$($script:PillTones[$tone][0])`"/><Setter TargetName=`"txt`" Property=`"Foreground`" Value=`"$($script:PillTones[$tone][1])`"/></DataTrigger>"
        }
        $xaml += '<DataTrigger Binding="{Binding ' + (ConvertTo-XmlText $path) + '}" Value="{x:Null}"><Setter TargetName="pill" Property="Visibility" Value="Collapsed"/></DataTrigger>'
        $xaml += '</DataTemplate.Triggers></DataTemplate>'
        $col = New-Object System.Windows.Controls.DataGridTemplateColumn
        $col.CellTemplate = [System.Windows.Markup.XamlReader]::Parse($xaml)
        $col.ClipboardContentBinding = New-Object System.Windows.Data.Binding $path
    }
    else {
        $col = New-Object System.Windows.Controls.DataGridTextColumn
        $col.Binding = New-Object System.Windows.Data.Binding $path
        $style = New-Object System.Windows.Style ([System.Windows.Controls.TextBlock])
        $style.Setters.Add((New-Object System.Windows.Setter([System.Windows.Controls.TextBlock]::TextTrimmingProperty, [System.Windows.TextTrimming]::CharacterEllipsis)))
        $col.ElementStyle = $style
        if ($Module.GoodWhenNo -contains $Name) { $col.CellStyle = Get-ThemeResource 'BoolCellInv' }
        elseif ($Module.ColorBools) { $col.CellStyle = Get-ThemeResource 'BoolCell' }
    }
    $col.Header = $Name
    $col.SortMemberPath = $Name
    $col.MaxWidth = 520
    $col.MinWidth = 54
    $g.Columns.Add($col)
}

function Add-ResultRows {
    # Dodaje obiekty jako wiersze tabeli modułu; nowe właściwości tworzą nowe kolumny.
    # -Computer wypełnia pierwszą kolumnę (-TargetColumn, domyślnie "Komputer").
    param([hashtable]$Module, [string]$Computer = '', [object[]]$Objects, [string]$TargetColumn = 'Komputer')
    $table = $Module.Table
    if ($null -eq $table) { return }
    $table.BeginLoadData()
    try {
        foreach ($obj in $Objects) {
            if ($null -eq $obj) { continue }
            $values = New-Object System.Collections.Specialized.OrderedDictionary
            if ($Computer -and $TargetColumn) { $values[$TargetColumn] = $Computer }
            $base = if ($obj -is [System.Management.Automation.PSObject]) { $obj.PSObject.BaseObject } else { $obj }
            if ($base -is [string] -or $base -is [System.ValueType]) {
                $values['Wynik'] = ConvertTo-CellValue $base
            }
            elseif ($base -is [System.Collections.IDictionary]) {
                foreach ($k in $base.Keys) {
                    if ($Computer -and [string]$k -eq $TargetColumn) { continue }
                    $values[[string]$k] = ConvertTo-CellValue $base[$k]
                }
            }
            else {
                foreach ($p in $obj.PSObject.Properties) {
                    if ($script:HiddenProperties -contains $p.Name) { continue }
                    if ($Computer -and $p.Name -eq $TargetColumn) { continue }
                    $values[$p.Name] = ConvertTo-CellValue $p.Value
                }
            }
            foreach ($name in @($values.Keys)) {
                if (-not $table.Columns.Contains($name)) { [void]$table.Columns.Add($name, [object]) }
                if ($Module.SecretColumns -contains $name -and -not $table.Columns.Contains("__secret_$name")) { [void]$table.Columns.Add("__secret_$name", [object]) }
                Add-GridColumn -Module $Module -Name $name
            }
            $row = $table.NewRow()
            $search = New-Object System.Text.StringBuilder
            foreach ($name in $values.Keys) {
                $v = $values[$name]
                if ($Module.SecretColumns -contains $name) {
                    $row["__secret_$name"] = $v
                    if (-not $Module.RevealSecrets -and $v -isnot [System.DBNull] -and [string]$v -ne '') { $v = '••••••••••' }
                    $row[$name] = $v
                    continue
                }
                $row[$name] = $v
                if ($v -isnot [System.DBNull] -and -not $name.StartsWith('__')) { [void]$search.Append([string]$v).Append(' ') }
            }
            $row['__search'] = $search.ToString().ToLowerInvariant()
            if (-not $values.Contains('__flag')) {
                $status = [string]$values['Status']
                if ($status.StartsWith('Błąd')) { $row['__flag'] = 'crit' }
            }
            $table.Rows.Add($row)
        }
    }
    finally {
        $table.EndLoadData()
    }
    Update-ResultCount -Module $Module
}

function Set-ResultValue {
    # Zmienia wartość w istniejącym wierszu (np. Status po wykonaniu akcji) i odświeża tekst wyszukiwania
    param([hashtable]$Module, [Parameter(Mandatory)]$Row, [Parameter(Mandatory)][string]$Column, $Value)
    if ($Row -is [System.Data.DataRowView]) { $Row = $Row.Row }
    $table = $Row.Table
    if (-not $table.Columns.Contains($Column)) {
        [void]$table.Columns.Add($Column, [object])
        Add-GridColumn -Module $Module -Name $Column
    }
    $Row[$Column] = ConvertTo-CellValue $Value
    $search = New-Object System.Text.StringBuilder
    foreach ($c in $table.Columns) {
        $n = $c.ColumnName
        if ($n.StartsWith('__') -or $Module.SecretColumns -contains $n) { continue }
        $v = $Row[$n]
        if ($v -isnot [System.DBNull]) { [void]$search.Append([string]$v).Append(' ') }
    }
    $Row['__search'] = $search.ToString().ToLowerInvariant()
}

function Remove-ResultRows {
    param([hashtable]$Module, [object[]]$Rows)
    foreach ($r in $Rows) {
        $dr = if ($r -is [System.Data.DataRowView]) { $r.Row } else { $r }
        if ($dr -and $dr.RowState -ne [System.Data.DataRowState]::Detached -and $dr.RowState -ne [System.Data.DataRowState]::Deleted) {
            $dr.Table.Rows.Remove($dr)
        }
    }
    Update-ResultCount -Module $Module
}

function Set-SecretReveal {
    param([hashtable]$Module, [bool]$Reveal)
    $Module.RevealSecrets = $Reveal
    $table = $Module.Table
    foreach ($name in $Module.SecretColumns) {
        if (-not $table.Columns.Contains($name)) { continue }
        foreach ($row in $table.Rows) {
            $real = $row["__secret_$name"]
            if ($real -is [System.DBNull] -or [string]$real -eq '') { continue }
            $row[$name] = if ($Reveal) { $real } else { '••••••••••' }
        }
    }
    if ($Reveal) { Write-Log 'Odkryto wartości poufne w tabeli wyników.' 'WARN' -Module $Module.Title }
    Update-DetailText -Module $Module
}

$script:FilterTimer = $null
function Request-ResultFilter {
    # Filtr z opóźnieniem (wpisywanie nie przelicza tabeli przy każdym znaku)
    param([hashtable]$Module)
    if (-not $script:FilterTimer) {
        $script:FilterTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:FilterTimer.Interval = [TimeSpan]::FromMilliseconds(220)
        $script:FilterTimer.add_Tick({
                param($s, $e)
                $s.Stop()
                $m = $script:FilterPending
                $script:FilterPending = $null
                if ($m) { Update-ResultFilter -Module $m }
            })
    }
    if ($script:FilterPending -and -not [object]::ReferenceEquals($script:FilterPending, $Module)) { Update-ResultFilter -Module $script:FilterPending }
    $script:FilterPending = $Module
    $script:FilterTimer.Stop()
    $script:FilterTimer.Start()
}
$script:FilterPending = $null

function Update-ResultFilter {
    param([hashtable]$Module)
    # Uwaga: DataView jest dla PowerShella listą - pusty widok byłby fałszem, stąd porównanie z $null
    if ($null -eq $Module.View) { return }
    $text = if ($Module.FilterBox) { $Module.FilterBox.Text.Trim().ToLowerInvariant() } else { '' }
    $terms = @($text -split '\s+' | Where-Object { $_ })
    $parts = foreach ($t in $terms) {
        if ($t.StartsWith('-') -and $t.Length -gt 1) { "__search NOT LIKE '*{0}*'" -f (ConvertTo-LikeLiteral $t.Substring(1)) }
        else { "__search LIKE '*{0}*'" -f (ConvertTo-LikeLiteral $t) }
    }
    if ($Module.ExtraFilter) { $parts = @($parts) + @('(' + $Module.ExtraFilter + ')') }
    $filter = (@($parts) -join ' AND ')
    try { $Module.View.RowFilter = $filter } catch { $Module.View.RowFilter = '' }
    Update-ResultCount -Module $Module
}

function Update-ResultCount {
    param([hashtable]$Module)
    if ($null -eq $Module.View -or -not $Module.View_) { return }
    $total = $Module.Table.Rows.Count
    $visible = $Module.View.Count
    $g = $Module.Grid
    if ($g) { $g.HeadersVisibility = if ($g.Columns.Count -gt 0) { 'Column' } else { 'None' } }
    $count = $Module.View_['countText']
    if ($count) { $count.Text = if ($visible -eq $total) { [string]$total } else { "$visible z $total" } }
    $empty = $Module.View_['emptyState']
    if ($empty) {
        if ($visible -gt 0) { $empty.Visibility = 'Collapsed' }
        else {
            $empty.Visibility = 'Visible'
            if ($Module.Busy -and $total -eq 0) {
                $Module.View_['emptyText'].Text = 'Trwa pobieranie danych…'
                $Module.View_['emptyHint'].Text = 'Wyniki pojawią się tu na bieżąco dla kolejnych obiektów.'
            }
            elseif ($total -gt 0) {
                $Module.View_['emptyText'].Text = 'Nic nie pasuje do filtra'
                $Module.View_['emptyHint'].Text = 'Zmień lub wyczyść filtr. Słowo poprzedzone minusem wyklucza wiersze.'
            }
            else {
                $Module.View_['emptyText'].Text = $(if ($Module.EmptyText) { $Module.EmptyText } else { 'Brak wyników' })
                $Module.View_['emptyHint'].Text = [string]$Module.EmptyHint
            }
        }
    }
}

function Update-DetailPane {
    param([hashtable]$Module)
    $v = $Module.View_
    if (-not $v -or -not $v['detailPane']) { return }
    if ($script:Settings.DetailVisible) {
        $v.detailPane.Visibility = 'Visible'
        $v.detailSplit.Visibility = 'Visible'
        if ($v.detailCol.Width.Value -lt 120) { $v.detailCol.Width = New-Object System.Windows.GridLength 340 }
    }
    else {
        $v.detailPane.Visibility = 'Collapsed'
        $v.detailSplit.Visibility = 'Collapsed'
        $v.detailCol.Width = New-Object System.Windows.GridLength 0
    }
}

function Format-RowDetails {
    param([hashtable]$Module, $Row)
    $sb = New-Object System.Text.StringBuilder
    foreach ($col in $Module.Table.Columns) {
        $name = $col.ColumnName
        if ($name.StartsWith('__')) { continue }
        $v = Get-ExportValue -Module $Module -Row $Row -Column $name
        if ([string]$v -eq '') { continue }
        $text = [string]$v
        if ($text -match "[\r\n]") { [void]$sb.AppendLine("${name}:").AppendLine($text.Trim()).AppendLine() }
        else { [void]$sb.AppendLine(('{0}:  {1}' -f $name, $text)) }
    }
    return $sb.ToString()
}

function Update-DetailText {
    param([hashtable]$Module)
    $box = $Module.View_['detailText']
    if (-not $box) { return }
    $item = $Module.Grid.SelectedItem
    if ($item -is [System.Data.DataRowView]) { $box.Text = Format-RowDetails -Module $Module -Row $item }
    else { $box.Text = '' }
}

function Show-RowDetails {
    param([hashtable]$Module, $Row)
    if (-not ($Row -is [System.Data.DataRowView])) { return }
    $title = 'Szczegóły'
    foreach ($c in 'Komputer', 'Login', 'Nazwa', 'Name') {
        $v = Get-ObjectValue $Row $c
        if ($v) { $title = "Szczegóły – $v"; break }
    }
    Show-TextDialog -Title $title -Subtitle $Module.Title -Text (Format-RowDetails -Module $Module -Row $Row)
}

function Get-VisibleColumnNames {
    param([hashtable]$Module)
    $cols = @($Module.Grid.Columns | Where-Object { $_.Visibility -eq 'Visible' } | Sort-Object DisplayIndex)
    return @($cols | ForEach-Object { [string]$_.SortMemberPath })
}

function Get-ExportValue {
    # Wartość komórki do eksportu/kopiowania (z maskowaniem wartości poufnych)
    param([hashtable]$Module, $Row, [string]$Column)
    if ($Module.SecretColumns -contains $Column) {
        $real = Get-ObjectValue $Row "__secret_$Column"
        if ($null -eq $real -or [string]$real -eq '') { return '' }
        if ($Module.RevealSecrets) { return $real }
        return '********'
    }
    $v = Get-ObjectValue $Row $Column
    if ($null -eq $v) { return '' }
    return $v
}

function Get-RowValue {
    # Prawdziwa wartość kolumny (także poufnej) - do użycia w akcjach modułów
    param($Row, [string]$Column)
    $secret = Get-ObjectValue $Row "__secret_$Column"
    if ($null -ne $secret) { return $secret }
    return (Get-ObjectValue $Row $Column)
}

function Get-SelectedResultRows {
    # Zaznaczone wiersze tabeli (DataRowView)
    param([hashtable]$Module)
    return @($Module.Grid.SelectedItems | Where-Object { $_ -is [System.Data.DataRowView] })
}

function Get-SelectedRowsByHost {
    # Grupuje zaznaczone wiersze po kolumnie Komputer: host -> lista hashtabel z wartościami kolumn
    param([hashtable]$Module, [string[]]$Columns, [object[]]$Rows = $null, [string]$TargetColumn = 'Komputer')
    $result = [ordered]@{}
    $source = @(if ($null -ne $Rows) { $Rows } else { Get-SelectedResultRows -Module $Module })
    foreach ($drv in $source) {
        $computer = [string](Get-ObjectValue $drv $TargetColumn)
        if (-not $computer -or (Get-ObjectValue $drv 'Status') -eq 'Błąd') { continue }
        $item = @{ __row = $drv }
        $valid = $true
        foreach ($c in $Columns) {
            $v = Get-RowValue $drv $c
            if ($null -eq $v -or [string]$v -eq '') { $valid = $false }
            $item[$c] = $v
        }
        if (-not $valid) { continue }
        if (-not $result.Contains($computer)) { $result[$computer] = New-Object System.Collections.ArrayList }
        [void]$result[$computer].Add($item)
    }
    return $result
}

function Export-ResultView {
    param([hashtable]$Module)
    if ($null -eq $Module.View -or $Module.View.Count -eq 0) { Show-Warning 'Brak danych do eksportu.'; return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv|Raport HTML (*.html)|*.html'
    $dlg.FileName = '{0}_{1:yyyyMMdd_HHmm}' -f (Get-SafeFileName $Module.Title), (Get-Date)
    $dlg.AddExtension = $true
    $dlg.DefaultExt = '.csv'
    $answer = if ($script:UI.Window) { $dlg.ShowDialog($script:UI.Window) } else { $dlg.ShowDialog() }
    if ($answer -ne $true) { return }
    $path = $dlg.FileName
    $columns = Get-VisibleColumnNames -Module $Module
    $rows = @($Module.View | ForEach-Object { $_ })
    if ($path -match '\.html?$') {
        $html = ConvertTo-HtmlReport -Module $Module -Columns $columns -Rows $rows
        [System.IO.File]::WriteAllText($path, $html, (New-Object System.Text.UTF8Encoding($true)))
    }
    else {
        $objects = foreach ($drv in $rows) {
            $o = [ordered]@{}
            foreach ($c in $columns) { $o[$c] = Get-ExportValue -Module $Module -Row $drv -Column $c }
            [pscustomobject]$o
        }
        $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
        $objects | Export-Csv -LiteralPath $path -NoTypeInformation -UseCulture -Encoding $encoding
    }
    Write-Log "Zapisano $($rows.Count) wierszy do pliku $path" 'OK' -Module $Module.Title
    Show-Toast "Zapisano $($rows.Count) wierszy: $([System.IO.Path]::GetFileName($path))" 'ok'
}

function ConvertTo-HtmlReport {
    # Samodzielny raport HTML (ciemny motyw jak w programie) z widocznych wierszy
    param([hashtable]$Module, [string[]]$Columns, [object[]]$Rows)
    $enc = { param($t) [System.Net.WebUtility]::HtmlEncode([string]$t) }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<!DOCTYPE html><html lang="pl"><head><meta charset="utf-8"><title>').Append((& $enc $Module.Title)).Append('</title><style>')
    [void]$sb.Append(':root{--bg:#0F1318;--card:#161B22;--line:#242B36;--text:#E4E8EF;--muted:#8791A5;--ok:#5EE3AE;--warn:#FFC46B;--crit:#FF7A86}')
    [void]$sb.Append('*{box-sizing:border-box}body{margin:0;padding:32px;background:var(--bg);color:var(--text);font:13px/1.45 "Segoe UI",system-ui,sans-serif}')
    [void]$sb.Append('h1{font-size:22px;margin:0 0 4px}.meta{color:var(--muted);margin-bottom:20px}.card{background:var(--card);border:1px solid var(--line);border-radius:10px;overflow:auto}')
    [void]$sb.Append('table{border-collapse:collapse;width:100%}th{position:sticky;top:0;background:#1B212A;color:var(--muted);text-align:left;font-weight:600;padding:9px 12px;border-bottom:1px solid var(--line);white-space:nowrap}')
    [void]$sb.Append('td{padding:7px 12px;border-bottom:1px solid #1F2530;vertical-align:top;white-space:pre-wrap}tr:nth-child(even) td{background:#181E26}')
    [void]$sb.Append('tr.crit td{color:var(--crit)}tr.warn td{color:var(--warn)}tr.muted td{color:#7B8496}.pill{display:inline-block;padding:1px 9px;border-radius:9px;font-weight:600;background:#1E252F}')
    [void]$sb.Append('.ok{color:var(--ok);background:#15291F}.p-warn{color:var(--warn);background:#2E2616}.p-crit{color:var(--crit);background:#2E1A1E}.info{color:#8CC0FF;background:#1A2640}</style></head><body>')
    [void]$sb.Append('<h1>').Append((& $enc $Module.Title)).Append('</h1><div class="meta">')
    $filter = if ($Module.FilterBox) { $Module.FilterBox.Text.Trim() } else { '' }
    $meta = 'Domain Ops {0} • {1:yyyy-MM-dd HH:mm} • {2} • wierszy: {3}' -f $script:AppVersion, (Get-Date), $env:USERNAME, $Rows.Count
    if ($filter) { $meta += " • filtr: $filter" }
    [void]$sb.Append((& $enc $meta)).Append('</div><div class="card"><table><thead><tr>')
    foreach ($c in $Columns) { [void]$sb.Append('<th>').Append((& $enc $c)).Append('</th>') }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($drv in $Rows) {
        $flag = [string](Get-ObjectValue $drv '__flag')
        $tone = [string](Get-ObjectValue $drv '__tone')
        [void]$sb.Append($(if ($flag) { "<tr class=`"$flag`">" } else { '<tr>' }))
        foreach ($c in $Columns) {
            $v = & $enc (Get-ExportValue -Module $Module -Row $drv -Column $c)
            if ($Module.PillColumns -contains $c -and $v) {
                $cls = switch ($tone) { 'ok' { 'ok' } 'warn' { 'p-warn' } 'crit' { 'p-crit' } 'info' { 'info' } default { '' } }
                $v = "<span class=`"pill $cls`">$v</span>"
            }
            [void]$sb.Append('<td>').Append($v).Append('</td>')
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div></body></html>')
    return $sb.ToString()
}

function Copy-ResultView {
    # Kopiuje zaznaczone wiersze (albo wszystkie widoczne) jako tekst rozdzielany tabulatorami (Excel)
    param([hashtable]$Module, [switch]$SelectedOnly)
    if ($null -eq $Module.View -or $Module.View.Count -eq 0) { Show-Warning 'Brak danych do skopiowania.'; return }
    $columns = Get-VisibleColumnNames -Module $Module
    $rows = @(Get-SelectedResultRows -Module $Module)
    if (-not $SelectedOnly -and $rows.Count -le 1) { $rows = @($Module.View | ForEach-Object { $_ }) }
    if ($rows.Count -eq 0) { return }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(($columns -join "`t"))
    foreach ($drv in $rows) {
        $cells = foreach ($c in $columns) { ([string](Get-ExportValue -Module $Module -Row $drv -Column $c)) -replace '[\t\r\n]+', ' ' }
        [void]$sb.AppendLine((@($cells) -join "`t"))
    }
    Set-ClipboardText $sb.ToString()
    Show-Toast "Skopiowano $($rows.Count) wierszy do schowka." 'ok'
}

function Show-GridDialog {
    # Dowolne obiekty w tabeli z filtrem, eksportem i kopiowaniem (okno z widokiem modułu)
    param([string]$Title, [object[]]$Rows, [string[]]$SecretColumns = @(), [string]$Subtitle = '', [string[]]$PillColumns = @())
    $m = New-ModuleContext -Definition @{ Key = 'Dialog_' + [guid]::NewGuid().ToString('N'); Title = $Title; Description = $Subtitle; Category = 'Podgląd'; Icon = 'E8FD'; Workspace = '' }
    $m.SecretColumns = @($SecretColumns)
    $m.PillColumns = @($PillColumns)
    New-ModuleView -Module $m
    Complete-ModuleView -Module $m
    foreach ($r in $Rows) { Add-ResultRows -Module $m -Objects @($r) }
    $w = New-Dialog -Title $Title -Subtitle $Subtitle -Body '' -Icon 'E8FD' -OkText 'Zamknij' -NoCancel -Width 1100 -Height 700 -Resizable
    # Nagłówek modułu jest zbędny w oknie - zostaje pasek wyników i tabela
    $m.Root.RowDefinitions[0].Height = New-Object System.Windows.GridLength 0
    $m.Root.Children[0].Visibility = 'Collapsed'
    [void]$w.FindName('dlgBody').Children.Add($m.Root)
    [void](Invoke-Dialog $w)
    [void]$script:GridModules.Remove($m.Grid)
}
#endregion

#region Silnik operacji w tle
# Każdy obiekt docelowy to osobne zadanie w puli wątków. Tryb 'Remote' wykonuje blok skryptu na komputerze
# przez Invoke-Command (blok dostaje jeden parametr: hashtablę $P). Tryb 'Local' wykonuje blok lokalnie
# z parametrami ($Target, $P, $Ctx) - np. polecenia AD, kopiowanie plików albo połączenie obu podejść.
# Wyniki odbiera timer w wątku okna, więc wszystkie aktualizacje interfejsu są bezpieczne.
$script:WorkerScript = @'
param($Target, $Mode, $ScriptText, $P, $Ctx)
$ProgressPreference = 'SilentlyContinue'
$result = @{ Target = $Target; Ok = $true; Data = @(); Errors = @() }
try {
    $sb = [scriptblock]::Create($ScriptText)
    if ($Mode -eq 'Remote') {
        $icErrors = $null
        $ic = @{
            ComputerName  = $Target
            ScriptBlock   = $sb
            ArgumentList  = @(, $P)
            ErrorAction   = 'SilentlyContinue'
            ErrorVariable = 'icErrors'
        }
        if ($Ctx.Credential) { $ic.Credential = $Ctx.Credential }
        if ($Ctx.SessionOption) { $ic.SessionOption = $Ctx.SessionOption }
        $result.Data = @(Invoke-Command @ic)
        if ($icErrors) { $result.Errors = @($icErrors | ForEach-Object { $_.Exception.Message }) }
    }
    else {
        $raw = @(& $sb $Target $P $Ctx 2>&1)
        $result.Data = @($raw | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] })
        $result.Errors = @($raw | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] } | ForEach-Object { $_.Exception.Message })
    }
}
catch {
    $result.Errors = @($result.Errors) + @($_.Exception.Message)
}
if ($result.Errors.Count -gt 0 -and $result.Data.Count -eq 0) { $result.Ok = $false }
[pscustomobject]$result
'@

function Initialize-Engine {
    if ($script:Engine.Pool) { return }
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    # Zasady wykonywania dotyczą tylko Windows (na innych platformach Open() rzuciłby wyjątek)
    if ([System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT) {
        try { $iss.ExecutionPolicy = [Microsoft.PowerShell.ExecutionPolicy]::Bypass } catch { }
    }
    $pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, [int]$script:Settings.ThrottleLimit, $iss, $Host)
    $pool.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $pool.Open()
    $script:Engine.Pool = $pool
}

function Initialize-EngineTimer {
    if ($script:Engine.Timer) { return }
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(150)
    $timer.add_Tick({ param($s, $e) Update-Operations })
    $script:Engine.Timer = $timer
}

function Set-EngineThrottle([int]$Limit) {
    $script:Settings.ThrottleLimit = [Math]::Min(64, [Math]::Max(1, $Limit))
    if ($script:Engine.Pool) {
        try { [void]$script:Engine.Pool.SetMaxRunspaces($script:Settings.ThrottleLimit) } catch { }
    }
}

function Set-ModuleBusy {
    param([hashtable]$Module, [bool]$Busy, [string]$Text = '')
    $Module.Busy = $Busy
    foreach ($b in $Module.Buttons) { $b.IsEnabled = -not $Busy }
    $chip = $Module.View_['busyChip']
    if ($chip) {
        $chip.Visibility = if ($Busy) { 'Visible' } else { 'Collapsed' }
        if ($Text) { $Module.View_['busyText'].Text = $Text }
    }
    Update-ResultCount -Module $Module
    Update-NavBusy -Module $Module
}

function New-SessionOption {
    # Odpowiednik New-PSSessionOption -OpenTimeout (obiekt tworzony bezpośrednio)
    $option = New-Object System.Management.Automation.Remoting.PSSessionOption
    $option.OpenTimeout = [TimeSpan]::FromSeconds([int]$script:Settings.TimeoutSec)
    return $option
}

function Start-HostOperation {
    <#
        Uruchamia blok skryptu dla każdego obiektu z -Targets (komputery albo dowolne identyfikatory w trybie -Local).
        -Output Grid : wyniki trafiają do tabeli modułu (kolumna -TargetColumn + właściwości obiektów)
        -Output Log  : wyniki wypisywane są w dzienniku
        -Output None : tylko -OnResult/-OnComplete
        -PerTarget   : osobna hashtabla $P dla wybranych obiektów (np. różne usługi na różnych komputerach)
        -OnResult { param($m, $r) }   - po zakończeniu każdego obiektu ($r: Target, Ok, Data, Errors)
        -OnComplete { param($m, $op) } - po zakończeniu wszystkich (nie wywoływane po anulowaniu)
    #>
    param(
        [Parameter(Mandatory)][hashtable]$Module,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Targets,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [switch]$Local,
        [hashtable]$Parameters = @{},
        [hashtable]$PerTarget = @{},
        [ValidateSet('Grid', 'Log', 'None')][string]$Output = 'Grid',
        [string]$TargetColumn = 'Komputer',
        [switch]$Append,
        [scriptblock]$OnResult,
        [scriptblock]$OnComplete
    )
    if ($Module.Busy) {
        Show-Warning "Poprzednia operacja w module «$($Module.Title)» jeszcze trwa. Poczekaj na jej zakończenie lub przerwij ją na pasku stanu."
        return
    }
    $items = @($Targets | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
    if ($items.Count -eq 0) { return }

    Initialize-Engine
    Initialize-EngineTimer
    if ($Output -eq 'Grid' -and -not $Append -and $Module.Grid) { Reset-ResultTable -Module $Module }

    $ctx = @{
        Credential    = Get-EffectiveCredential
        Server        = [string]$script:Settings.DomainController
        SessionOption = New-SessionOption
    }
    $mode = if ($Local) { 'Local' } else { 'Remote' }
    $scriptText = $ScriptBlock.ToString()

    $op = @{
        Id           = $script:Engine.NextId
        Name         = $Name
        Module       = $Module
        Output       = $Output
        TargetColumn = $TargetColumn
        OnResult     = $OnResult
        OnComplete   = $OnComplete
        Items        = New-Object System.Collections.ArrayList
        Total        = $items.Count
        Done         = 0
        Failed       = 0
        Cancelled    = $false
        Started      = Get-Date
    }
    $script:Engine.NextId++

    foreach ($t in $items) {
        $p = $Parameters
        if ($PerTarget.ContainsKey($t)) { $p = $PerTarget[$t] }
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.RunspacePool = $script:Engine.Pool
        [void]$ps.AddScript($script:WorkerScript)
        [void]$ps.AddArgument($t).AddArgument($mode).AddArgument($scriptText).AddArgument($p).AddArgument($ctx)
        $handle = $ps.BeginInvoke()
        [void]$op.Items.Add(@{ Target = $t; PS = $ps; Handle = $handle; Finished = $false })
    }

    [void]$script:Engine.Operations.Add($op)
    Set-ModuleBusy -Module $Module -Busy $true -Text ("{0}: 0/{1}" -f $Name, $items.Count)
    $list = (@($items | Select-Object -First 8) -join ', ')
    if ($items.Count -gt 8) { $list += ", … (+$($items.Count - 8))" }
    Write-Log ("{0} – start ({1}): {2}" -f $Name, $items.Count, $list) -Module $Module.Title
    $script:Engine.Timer.Start()
    Update-StatusBar
}

function Update-Operations {
    # Wywoływane przez timer w wątku okna
    if ($script:Engine.InTick) { return }
    $script:Engine.InTick = $true
    try {
        foreach ($op in @($script:Engine.Operations)) {
            $changed = $false
            foreach ($item in $op.Items) {
                if ($item.Finished -or -not $item.Handle.IsCompleted) { continue }
                $item.Finished = $true
                $changed = $true
                $r = $null
                try {
                    $output = $item.PS.EndInvoke($item.Handle)
                    if ($output -and $output.Count -gt 0) { $r = $output[0] }
                    if (-not $r) {
                        $msg = 'Zadanie nie zwróciło wyniku.'
                        if ($item.PS.Streams.Error.Count -gt 0) { $msg = $item.PS.Streams.Error[0].Exception.Message }
                        $r = [pscustomobject]@{ Target = $item.Target; Ok = $false; Data = @(); Errors = @($msg) }
                    }
                }
                catch {
                    $msg = if ($op.Cancelled) { 'Anulowano.' } else { $_.Exception.Message }
                    $r = [pscustomobject]@{ Target = $item.Target; Ok = $false; Data = @(); Errors = @($msg) }
                }
                finally {
                    try { $item.PS.Dispose() } catch { }
                }
                Complete-OperationItem -Operation $op -Result $r
            }
            if ($changed -and $op.Module.View_['busyText']) {
                $text = '{0}: {1}/{2}' -f $op.Name, $op.Done, $op.Total
                if ($op.Failed -gt 0) { $text += " • błędy: $($op.Failed)" }
                $op.Module.View_['busyText'].Text = $text
            }
            if ($op.Done -ge $op.Total) { Complete-Operation -Operation $op }
        }
    }
    catch {
        Write-Log "Błąd silnika operacji: $($_.Exception.Message)" 'ERROR' -Module ''
    }
    finally {
        $script:Engine.InTick = $false
        if ($script:Engine.Operations.Count -eq 0 -and $script:Engine.Timer) { $script:Engine.Timer.Stop() }
        Update-StatusBar
    }
    # Akcje odroczone (np. okna z wynikami) - wykonywane po zwolnieniu blokady, aby modalne okno
    # nie wstrzymywało odbierania wyników pozostałych operacji
    while ($script:Engine.Deferred.Count -gt 0) {
        $next = $script:Engine.Deferred[0]
        $script:Engine.Deferred.RemoveAt(0)
        Invoke-UiAction -Module $next.Module -Action $next.Action
    }
}

function Invoke-Deferred {
    param([hashtable]$Module, [scriptblock]$Action)
    [void]$script:Engine.Deferred.Add(@{ Module = $Module; Action = $Action })
}

function Complete-OperationItem {
    param([hashtable]$Operation, $Result)
    $Operation.Done++
    $m = $Operation.Module
    $previous = $script:LogContext
    $script:LogContext = $m.Title
    try {
        $errorText = (@($Result.Errors | Where-Object { $_ } | Select-Object -Unique)) -join ' | '
        if (-not $Result.Ok) {
            $Operation.Failed++
            Write-Log ("[{0}] {1}" -f $Result.Target, $errorText) 'ERROR'
            if ($Operation.Output -eq 'Grid') {
                Add-ResultRows -Module $m -Computer $Result.Target -TargetColumn $Operation.TargetColumn -Objects @([pscustomobject]@{ 'Status' = 'Błąd'; 'Szczegóły' = $errorText })
            }
        }
        else {
            if ($errorText) { Write-Log ("[{0}] ostrzeżenia: {1}" -f $Result.Target, $errorText) 'WARN' }
            $data = @($Result.Data | Where-Object { $null -ne $_ })
            switch ($Operation.Output) {
                'Grid' {
                    if ($data.Count -gt 0) { Add-ResultRows -Module $m -Computer $Result.Target -TargetColumn $Operation.TargetColumn -Objects $data }
                    else { Write-Log ("[{0}] brak wyników" -f $Result.Target) }
                }
                'Log' {
                    foreach ($d in $data) {
                        $text = Format-LogObject $d
                        $level = if ($text -match 'Błąd') { 'WARN' } else { 'OK' }
                        Write-Log ("[{0}] {1}" -f $Result.Target, $text) $level
                    }
                }
            }
        }
        if ($Operation.OnResult) {
            try { $null = & $Operation.OnResult $m $Result }
            catch { Write-Log "Błąd obsługi wyniku ($($Result.Target)): $($_.Exception.Message)" 'ERROR' }
        }
    }
    finally {
        $script:LogContext = $previous
    }
}

function Complete-Operation {
    param([hashtable]$Operation)
    [void]$script:Engine.Operations.Remove($Operation)
    $m = $Operation.Module
    Set-ModuleBusy -Module $m -Busy $false
    $seconds = [Math]::Round(((Get-Date) - $Operation.Started).TotalSeconds, 1)
    $okCount = $Operation.Done - $Operation.Failed
    if ($Operation.Cancelled) {
        Write-Log ("{0} – przerwano (zakończone: {1}, czas {2} s)" -f $Operation.Name, $okCount, $seconds) 'WARN' -Module $m.Title
        Show-Toast ("{0}: przerwano" -f $Operation.Name) 'warn'
    }
    else {
        $level = if ($Operation.Failed -gt 0) { 'WARN' } else { 'OK' }
        Write-Log ("{0} – zakończono: {1} OK, {2} z błędem, czas {3} s" -f $Operation.Name, $okCount, $Operation.Failed, $seconds) $level -Module $m.Title
        if ($Operation.Total -gt 1 -or $Operation.Failed -gt 0 -or $seconds -ge 3) {
            $tone = if ($Operation.Failed -eq 0) { 'ok' } elseif ($okCount -gt 0) { 'warn' } else { 'crit' }
            $text = if ($Operation.Failed -eq 0) { '{0}: gotowe ({1}, {2} s)' -f $Operation.Name, $okCount, $seconds }
            else { '{0}: {1} OK, {2} z błędem' -f $Operation.Name, $okCount, $Operation.Failed }
            Show-Toast $text $tone
        }
    }
    if ($Operation.OnComplete -and -not $Operation.Cancelled) {
        $previous = $script:LogContext
        $script:LogContext = $m.Title
        try { $null = & $Operation.OnComplete $m $Operation }
        catch { Write-Log "Błąd po zakończeniu operacji: $($_.Exception.Message)" 'ERROR' }
        finally { $script:LogContext = $previous }
    }
}

function Stop-AllOperations {
    $ops = @($script:Engine.Operations)
    if ($ops.Count -eq 0) { return }
    foreach ($op in $ops) {
        $op.Cancelled = $true
        foreach ($item in $op.Items) {
            if (-not $item.Finished) { try { [void]$item.PS.BeginStop($null, $null) } catch { } }
        }
    }
    Write-Log 'Przerywanie trwających operacji…' 'WARN' -Module ''
}

function Close-Engine {
    try { if ($script:Engine.Timer) { $script:Engine.Timer.Stop() } } catch { }
    $handles = @()
    foreach ($op in @($script:Engine.Operations)) {
        foreach ($item in $op.Items) {
            if (-not $item.Finished) { try { $handles += $item.PS.BeginStop($null, $null) } catch { } }
        }
    }
    foreach ($h in $handles) { try { [void]$h.AsyncWaitHandle.WaitOne(3000) } catch { } }
    if ($script:Engine.Pool) {
        try { $script:Engine.Pool.Close(); $script:Engine.Pool.Dispose() } catch { }
        $script:Engine.Pool = $null
    }
}

function Update-StatusBar {
    $c = $script:UI.Controls
    $label = $c['txtStatus']
    if (-not $label) { return }
    $ops = @($script:Engine.Operations)
    if ($ops.Count -eq 0) {
        $label.Text = 'Gotowe'
        $c.statusDot.Fill = Get-Brush '#5EE3AE'
        $c.prgStatus.Visibility = 'Collapsed'
        $c.btnCancelOps.Visibility = 'Collapsed'
        return
    }
    $total = 0
    $done = 0
    $parts = foreach ($op in $ops) {
        $total += $op.Total
        $done += $op.Done
        $text = '{0}: {1}/{2}' -f $op.Name, $op.Done, $op.Total
        if ($op.Failed -gt 0) { $text += " (błędy: $($op.Failed))" }
        $text
    }
    $label.Text = 'Trwa: ' + (@($parts) -join '   •   ')
    $c.statusDot.Fill = Get-Brush '#4C7DF0'
    $bar = $c.prgStatus
    $bar.Maximum = [Math]::Max(1, $total)
    $bar.Value = [Math]::Min($done, $bar.Maximum)
    $bar.Visibility = 'Visible'
    $c.btnCancelOps.Visibility = 'Visible'
}
#endregion

#region Przestrzenie robocze i moduły (rejestracja i budowa)
function Register-Workspace {
    <#
        Przestrzeń robocza = przycisk na górnym pasku + własna nawigacja modułów.
        -Target: 'Computer' (lista komputerów po lewej), 'User' (lista kont) albo 'None' (bez listy).
    #>
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$Title,
        [string]$Icon = 'E80F',
        [ValidateSet('Computer', 'User', 'None')][string]$Target = 'None',
        [string]$Description = ''
    )
    $script:UI.Workspaces[$Key] = @{
        Key         = $Key
        Title       = $Title
        Icon        = $Icon
        Target      = $Target
        Description = $Description
        NavHost     = $null
        Tab         = $null
        NavItems    = @{}
    }
}

function Register-Module {
    <#
        Moduł = pozycja w nawigacji przestrzeni roboczej, w kategorii -Category.
        -Build { param($m) } buduje panel parametrów i akcje (wywoływany przy pierwszym otwarciu).
        Ponowna rejestracja tego samego klucza zastępuje moduł (np. własną wersją z folderu modułów).
    #>
    param(
        [Parameter(Mandatory)][string]$Workspace,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$Title,
        [string]$Icon = 'E74C',
        [string]$Description = '',
        [string]$Badge = '',
        [Parameter(Mandatory)][scriptblock]$Build
    )
    $definition = @{ Workspace = $Workspace; Category = $Category; Key = $Key; Title = $Title; Icon = $Icon; Description = $Description; Badge = $Badge; Build = $Build }
    for ($i = 0; $i -lt $script:UI.ModuleDefs.Count; $i++) {
        if ($script:UI.ModuleDefs[$i].Key -eq $Key) { $script:UI.ModuleDefs[$i] = $definition; return }
    }
    [void]$script:UI.ModuleDefs.Add($definition)
}

function New-ModuleContext {
    param([hashtable]$Definition)
    $ws = $null
    if ($Definition['Workspace'] -and $script:UI.Workspaces.Contains($Definition['Workspace'])) { $ws = $script:UI.Workspaces[$Definition['Workspace']] }
    return @{
        Key            = $Definition['Key']
        Title          = $Definition['Title']
        Description    = $Definition['Description']
        Category       = $Definition['Category']
        Icon           = $Definition['Icon']
        Workspace      = $Definition['Workspace']
        WorkspaceTitle = $(if ($ws) { $ws.Title } else { '' })
        Target         = $(if ($ws) { $ws.Target } else { 'None' })
        Busy           = $false
        Buttons        = New-Object System.Collections.ArrayList
        PrimaryButton  = $null
        RowActions     = New-Object System.Collections.ArrayList
        Actions        = @{}
        RowDoubleClick = $null
        SecretColumns  = @()
        PillColumns    = @()
        GoodWhenNo     = @()
        HiddenColumns  = @()
        ColorBools     = $false
        RevealSecrets  = $false
        Stats          = @{}
        EmptyText      = ''
        EmptyHint      = ''
        EmptyIcon      = ''
        ResultHint     = ''
        ExtraFilter    = ''
        ParamsCollapsed = $false
        Data           = @{}
        View_          = @{}
        ColumnIndex    = @{}
        Root           = $null
        ParamsCard     = $null
        ParamsStack    = $null
        StatsGrid      = $null
        Grid           = $null
        Table          = $null
        View           = $null
        FilterBox      = $null
    }
}

function Get-ModuleDefinition([string]$Key) {
    foreach ($d in $script:UI.ModuleDefs) { if ($d.Key -eq $Key) { return $d } }
    return $null
}

function Initialize-Module {
    param([hashtable]$Definition)
    $m = New-ModuleContext -Definition $Definition
    $script:UI.Modules[$Definition.Key] = $m
    New-ModuleView -Module $m
    $previous = $script:LogContext
    $script:LogContext = $m.Title
    try { $null = & $Definition.Build $m }
    catch {
        Write-Log "Nie udało się zbudować modułu: $($_.Exception.Message)" 'ERROR'
        Add-Label -Parent (Add-ToolbarRow -Module $m) -Text "Błąd modułu: $($_.Exception.Message)" | Out-Null
    }
    finally { $script:LogContext = $previous }
    Complete-ModuleView -Module $m
    $m.Root.Visibility = 'Collapsed'
    [void]$script:UI.Controls['contentHost'].Children.Add($m.Root)
    return $m
}

function Show-Module {
    param([string]$Key)
    $definition = Get-ModuleDefinition $Key
    if (-not $definition) { return }
    if ($script:UI.ActiveWorkspace -ne $definition.Workspace) { Show-Workspace -Key $definition.Workspace -ModuleKey $Key; return }
    $m = $script:UI.Modules[$Key]
    if (-not $m) { $m = Invoke-WithWaitCursor { Initialize-Module -Definition $definition } }
    $active = $script:UI.ActiveModule
    if ($active -and -not [object]::ReferenceEquals($active, $m)) { $active.Root.Visibility = 'Collapsed' }
    $m.Root.Visibility = 'Visible'
    $script:UI.ActiveModule = $m
    $script:Settings.LastModules[$definition.Workspace] = $Key
    $ws = $script:UI.Workspaces[$definition.Workspace]
    $nav = $ws.NavItems[$Key]
    if ($nav -and $nav.IsChecked -ne $true) { $nav.IsChecked = $true }
}

function Show-Workspace {
    param([string]$Key, [string]$ModuleKey = '')
    if (-not $script:UI.Workspaces.Contains($Key)) { $Key = @($script:UI.Workspaces.Keys)[0] }
    $ws = $script:UI.Workspaces[$Key]
    $script:UI.ActiveWorkspace = $Key
    $script:Settings.LastWorkspace = $Key
    foreach ($other in $script:UI.Workspaces.Values) {
        if ($other.NavHost) { $other.NavHost.Visibility = if ($other.Key -eq $Key) { 'Visible' } else { 'Collapsed' } }
        if ($other.Tab -and $other.Key -eq $Key -and $other.Tab.IsChecked -ne $true) { $other.Tab.IsChecked = $true }
    }
    $c = $script:UI.Controls
    $c.panelComputers.Visibility = if ($ws.Target -eq 'Computer') { 'Visible' } else { 'Collapsed' }
    $c.panelUsers.Visibility = if ($ws.Target -eq 'User') { 'Visible' } else { 'Collapsed' }
    $c.targetColumn.Width = if ($ws.Target -eq 'None') { New-Object System.Windows.GridLength 0 } else { New-Object System.Windows.GridLength 300 }
    $c.txtSubtitle.Text = $ws.Description
    if (-not $ModuleKey) { $ModuleKey = [string]$script:Settings.LastModules[$Key] }
    $def = Get-ModuleDefinition $ModuleKey
    if (-not $def -or $def.Workspace -ne $Key) {
        $def = $null
        foreach ($d in $script:UI.ModuleDefs) { if ($d.Workspace -eq $Key) { $def = $d; break } }
    }
    if ($def) { Show-Module -Key $def.Key }
    elseif ($script:UI.ActiveModule) { $script:UI.ActiveModule.Root.Visibility = 'Collapsed'; $script:UI.ActiveModule = $null }
}

function Update-NavBusy {
    # Kropka przy pozycji nawigacji pokazuje moduły z trwającą operacją
    param([hashtable]$Module)
    if (-not $Module.Workspace -or -not $script:UI.Workspaces.Contains($Module.Workspace)) { return }
    $nav = $script:UI.Workspaces[$Module.Workspace].NavItems[$Module.Key]
    if (-not $nav) { return }
    $dot = $script:NavDots[$nav]
    if ($dot) { $dot.Visibility = if ($Module.Busy) { 'Visible' } else { 'Collapsed' } }
}
$script:NavDots = @{}

function Get-TargetComputers {
    # Komputery zaznaczone na liście po lewej
    param([switch]$Quiet)
    $rows = @($script:UI.HostTable.Select('Sel = true', 'Name ASC'))
    $names = @($rows | ForEach-Object { [string]$_['Name'] } | Where-Object { $_ } | Select-Object -Unique)
    if ($names.Count -eq 0 -and -not $Quiet) { Show-Warning 'Zaznacz komputery na liście po lewej stronie.' }
    return $names
}

function Get-TargetUsers {
    # Konta zaznaczone na liście użytkowników (sAMAccountName)
    param([switch]$Quiet)
    $rows = @($script:UI.UserTable.Select('Sel = true', 'Login ASC'))
    $names = @($rows | ForEach-Object { [string]$_['Login'] } | Where-Object { $_ } | Select-Object -Unique)
    if ($names.Count -eq 0 -and -not $Quiet) { Show-Warning 'Zaznacz konta na liście użytkowników po lewej stronie.' }
    return $names
}
#endregion

#region Listy obiektów docelowych (lewy panel): komputery i użytkownicy
# Oba panele działają tak samo: DataTable + DataView (szybki filtr), ListBox z szablonem wiersza,
# pole wyboru w wierszu zapisuje stan bezpośrednio w DataRow (kolumna Sel), dzięki czemu zaznaczenia
# przetrwają filtrowanie. Kolumna Dot to kolor kropki stanu, Sub - druga linia opisu.
$script:TargetPanelXaml = @'
<Border xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Background="#12171D" BorderBrush="#1E242E" BorderThickness="0,0,1,0">
  <Grid Margin="14,14,12,10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid Margin="2,0,0,10">
      <TextBlock x:Name="pTitle" Foreground="#8791A5" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
      <Border Style="{StaticResource Chip}" HorizontalAlignment="Right" Margin="0">
        <TextBlock x:Name="pCount" Text="0" Foreground="#AEB6C4" FontSize="11"/>
      </Border>
    </Grid>
    <Border Grid.Row="1" Style="{StaticResource Card}" Padding="12,12,12,6" Margin="0,0,0,10">
      <StackPanel x:Name="pSource"/>
    </Border>
    <TextBox x:Name="pSearch" Grid.Row="2" Margin="0,0,0,8"/>
    <Grid Grid.Row="3">
      <ListBox x:Name="pList" SelectionMode="Extended">
        <ListBox.ItemTemplate>
          <DataTemplate>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <CheckBox IsChecked="{Binding [Sel], Mode=OneWay}" Margin="0,0,10,0" VerticalAlignment="Center" Focusable="False"/>
              <Ellipse Grid.Column="1" Width="8" Height="8" Fill="{Binding [Dot]}" Margin="0,0,10,0" VerticalAlignment="Center"/>
              <StackPanel Grid.Column="2">
                <TextBlock Text="{Binding [Title]}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                <TextBlock Text="{Binding [Sub]}" Foreground="#7B8496" FontSize="11.5" TextTrimming="CharacterEllipsis"/>
              </StackPanel>
            </Grid>
          </DataTemplate>
        </ListBox.ItemTemplate>
      </ListBox>
      <StackPanel x:Name="pEmpty" HorizontalAlignment="Center" VerticalAlignment="Center" IsHitTestVisible="False" Margin="10">
        <TextBlock x:Name="pEmptyIcon" Style="{StaticResource Glyph}" FontSize="28" Foreground="#3A4352" HorizontalAlignment="Center"/>
        <TextBlock x:Name="pEmptyText" Foreground="#5E6779" TextAlignment="Center" TextWrapping="Wrap" Margin="0,10,0,0"/>
      </StackPanel>
    </Grid>
    <Grid Grid.Row="4" Margin="0,8,0,0">
      <StackPanel Orientation="Horizontal">
        <Button x:Name="pAll" Style="{StaticResource GhostButton}" ToolTip="Zaznacz widoczne"/>
        <Button x:Name="pNone" Style="{StaticResource GhostButton}" ToolTip="Odznacz wszystkie"/>
        <Button x:Name="pInvert" Style="{StaticResource GhostButton}" ToolTip="Odwróć zaznaczenie widocznych"/>
      </StackPanel>
      <TextBlock x:Name="pSel" HorizontalAlignment="Right" VerticalAlignment="Center" Foreground="#8791A5" FontSize="12"/>
    </Grid>
  </Grid>
</Border>
'@

$script:DotColors = @{ unknown = '#3A4352'; ok = '#5EE3AE'; warn = '#FFC46B'; crit = '#FF7A86'; off = '#5E6779' }

function New-TargetPanel {
    # Wspólny szkielet panelu; zwraca hashtablę z kontrolkami i tabelą
    param([string]$Kind, [string]$Title, [string]$Placeholder, [string]$EmptyIcon, [string]$EmptyText, [string[]]$Columns, [string]$KeyColumn)
    $root = New-UiElement $script:TargetPanelXaml
    $p = @{ Kind = $Kind; Root = $root; KeyColumn = $KeyColumn; EmptyText = $EmptyText }
    foreach ($n in 'pTitle', 'pCount', 'pSource', 'pSearch', 'pList', 'pEmpty', 'pEmptyIcon', 'pEmptyText', 'pAll', 'pNone', 'pInvert', 'pSel') { $p[$n] = $root.FindName($n) }
    $p.pTitle.Text = $Title
    $p.pSearch.Tag = $Placeholder
    $p.pEmptyIcon.Text = Get-Glyph $EmptyIcon
    $p.pEmptyText.Text = $EmptyText
    $p.pAll.Content = New-IconContent -Text '' -Icon 'E8B3' -IconSize 13
    $p.pNone.Content = New-IconContent -Text '' -Icon 'E894' -IconSize 13
    $p.pInvert.Content = New-IconContent -Text '' -Icon 'E8AB' -IconSize 13

    $t = New-Object System.Data.DataTable $Kind
    [void]$t.Columns.Add('Sel', [bool])
    foreach ($c in $Columns) {
        if ($c -eq 'LastLogon' -or $c -eq 'PwdLastSet') { [void]$t.Columns.Add($c, [datetime]) }
        else { [void]$t.Columns.Add($c, [string]) }
    }
    foreach ($c in 'Source', 'Dot', 'Title', 'Sub', 'Status') { if (-not $t.Columns.Contains($c)) { [void]$t.Columns.Add($c, [string]) } }
    $t.Columns['Sel'].DefaultValue = $false
    $t.Columns['Dot'].DefaultValue = $script:DotColors.unknown
    $t.PrimaryKey = [System.Data.DataColumn[]]@($t.Columns[$KeyColumn])
    $t.CaseSensitive = $false
    $view = [System.Data.DataView]::new($t)
    $view.Sort = "$KeyColumn ASC"
    $p.Table = $t
    $p.View = $view
    $p.pList.ItemsSource = $view

    $script:TargetPanels[$p.pList] = $p
    $script:TargetPanels[$p.pSearch] = $p
    $p.pList.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler]$script:TargetEvents.ItemClick)
    $p.pList.add_PreviewKeyDown($script:TargetEvents.KeyDown)
    $p.pSearch.add_TextChanged($script:TargetEvents.SearchChanged)
    foreach ($pair in @(@('pAll', 'CheckVisible'), @('pNone', 'UncheckAll'), @('pInvert', 'InvertVisible'))) {
        $script:TargetPanels[$p[$pair[0]]] = $p
        $script:TargetButtonModes[$p[$pair[0]]] = $pair[1]
        $p[$pair[0]].add_Click($script:TargetEvents.ButtonClick)
    }
    Update-TargetCount $p
    return $p
}

$script:TargetPanels = New-Object 'System.Collections.Generic.Dictionary[object,hashtable]'
$script:TargetButtonModes = New-Object 'System.Collections.Generic.Dictionary[object,string]'
$script:TargetEvents = @{
    ItemClick     = {
        param($s, $e)
        try {
            $cb = $e.OriginalSource
            if (-not ($cb -is [System.Windows.Controls.CheckBox])) { return }
            $drv = $cb.DataContext
            if (-not ($drv -is [System.Data.DataRowView])) { return }
            $p = $script:TargetPanels[$s]
            $value = ($cb.IsChecked -eq $true)
            # Kliknięcie pola w jednym z kilku zaznaczonych wierszy zmienia je wszystkie
            $rows = @($s.SelectedItems | Where-Object { $_ -is [System.Data.DataRowView] })
            if ($rows.Count -gt 1 -and $rows -contains $drv) { foreach ($r in $rows) { $r.Row['Sel'] = $value } }
            else { $drv.Row['Sel'] = $value }
            Update-TargetCount $p
        }
        catch { Write-Log "Błąd zaznaczania: $($_.Exception.Message)" 'ERROR' }
    }
    KeyDown       = {
        param($s, $e)
        try {
            if ($e.Key -ne [System.Windows.Input.Key]::Space) { return }
            $p = $script:TargetPanels[$s]
            $rows = @($s.SelectedItems | Where-Object { $_ -is [System.Data.DataRowView] })
            if ($rows.Count -eq 0) { return }
            $value = -not [bool]$rows[0].Row['Sel']
            foreach ($r in $rows) { $r.Row['Sel'] = $value }
            Update-TargetCount $p
            $e.Handled = $true
        }
        catch { }
    }
    SearchChanged = {
        param($s, $e)
        try { Update-TargetFilter $script:TargetPanels[$s] } catch { }
    }
    ButtonClick   = {
        param($s, $e)
        try { Set-TargetCheck -Panel $script:TargetPanels[$s] -Mode $script:TargetButtonModes[$s] } catch { }
    }
}

function Update-TargetCount {
    param([hashtable]$Panel)
    $t = $Panel.Table
    $selected = @($t.Select('Sel = true')).Count
    $Panel.pCount.Text = [string]$t.Rows.Count
    $Panel.pSel.Text = if ($Panel.View.Count -ne $t.Rows.Count) { "zaznaczone: $selected • widoczne: $($Panel.View.Count)" } else { "zaznaczone: $selected z $($t.Rows.Count)" }
    $Panel.pSel.Foreground = Get-Brush $(if ($selected -gt 0) { '#8CB0FF' } else { '#7B8496' })
    if ($Panel.View.Count -gt 0) { $Panel.pEmpty.Visibility = 'Collapsed' }
    else {
        $Panel.pEmpty.Visibility = 'Visible'
        $Panel.pEmptyText.Text = if ($t.Rows.Count -gt 0) { 'Nic nie pasuje do wyszukiwania' } else { $Panel.EmptyText }
    }
}

function Update-TargetFilter {
    param([hashtable]$Panel)
    $text = $Panel.pSearch.Text.Trim()
    if ($text) {
        $lit = ConvertTo-LikeLiteral $text
        $parts = foreach ($c in $Panel.SearchColumns) { "[$c] LIKE '*$lit*'" }
        $Panel.View.RowFilter = (@($parts) -join ' OR ')
    }
    else { $Panel.View.RowFilter = '' }
    Update-TargetCount $Panel
}

function Set-TargetCheck {
    param([hashtable]$Panel, [ValidateSet('CheckVisible', 'UncheckAll', 'InvertVisible', 'CheckSelected', 'UncheckSelected')][string]$Mode)
    $t = $Panel.Table
    switch ($Mode) {
        'CheckVisible' { foreach ($drv in @($Panel.View | ForEach-Object { $_ })) { $drv.Row['Sel'] = $true } }
        'UncheckAll' { foreach ($r in $t.Rows) { if ([bool]$r['Sel']) { $r['Sel'] = $false } } }
        'InvertVisible' { foreach ($drv in @($Panel.View | ForEach-Object { $_ })) { $drv.Row['Sel'] = -not [bool]$drv.Row['Sel'] } }
        'CheckSelected' { foreach ($drv in @($Panel.pList.SelectedItems)) { $drv.Row['Sel'] = $true } }
        'UncheckSelected' { foreach ($drv in @($Panel.pList.SelectedItems)) { $drv.Row['Sel'] = $false } }
    }
    Update-TargetCount $Panel
}

function ConvertTo-DbValue($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return [System.DBNull]::Value }
    return $Value
}

function Import-TargetRows {
    # Dodaje/aktualizuje pozycje panelu; zachowuje zaznaczenia istniejących pozycji
    param([hashtable]$Panel, [object[]]$Items, [string]$Source, [switch]$Check, [switch]$ReplaceSource)
    $t = $Panel.Table
    $key = $Panel.KeyColumn
    $added = 0
    # Odpięcie widoku od listy na czas importu - tysiące wierszy dodają się wtedy szybko
    $Panel.pList.ItemsSource = $null
    $t.BeginLoadData()
    try {
        if ($ReplaceSource) {
            foreach ($r in @($t.Select(("Source = '{0}' AND Sel = false" -f $Source.Replace("'", "''"))))) { $t.Rows.Remove($r) }
        }
        foreach ($it in $Items) {
            $name = ([string](Get-ObjectValue $it $key)).Trim()
            if (-not $name) { continue }
            $row = $t.Rows.Find($name)
            $isNew = ($null -eq $row)
            if ($isNew) {
                $row = $t.NewRow()
                $row[$key] = $name
                $row['Sel'] = $false
                $row['Source'] = $Source
            }
            elseif ($Source -eq 'AD') { $row['Source'] = 'AD' }
            foreach ($col in $t.Columns) {
                $n = $col.ColumnName
                if ($n -eq $key -or $n -eq 'Sel' -or $n -eq 'Source' -or $n -eq 'Dot' -or $n -eq 'Title' -or $n -eq 'Sub') { continue }
                $v = Get-ObjectValue $it $n
                if ($null -eq $v) { continue }
                if ($col.DataType -eq [datetime]) { if ($v -is [datetime]) { $row[$n] = $v } }
                elseif ($v -is [bool]) { $row[$n] = $(if ($v) { 'Tak' } else { 'Nie' }) }
                else { $row[$n] = ConvertTo-DbValue ([string]$v) }
            }
            if ($Check) { $row['Sel'] = $true }
            & $Panel.Describe $row
            if ($isNew) {
                $t.Rows.Add($row)
                $added++
            }
        }
    }
    finally {
        $t.EndLoadData()
        $t.AcceptChanges()
        $Panel.pList.ItemsSource = $Panel.View
    }
    Update-TargetCount $Panel
    return $added
}

function Remove-TargetRows {
    param([hashtable]$Panel)
    $rows = @($Panel.pList.SelectedItems | ForEach-Object { $_.Row })
    if ($rows.Count -eq 0) { return }
    foreach ($r in $rows) { $Panel.Table.Rows.Remove($r) }
    $Panel.Table.AcceptChanges()
    Update-TargetCount $Panel
}

function Get-TargetSelection {
    # Nazwy z wierszy podświetlonych na liście (menu kontekstowe)
    param([hashtable]$Panel)
    return @($Panel.pList.SelectedItems | ForEach-Object { [string]$_.Row[$Panel.KeyColumn] })
}

function New-SourceRow {
    # Wiersz w karcie źródła danych: kontrolka rozciągana + przyciski za nią
    param($Stretch, [object[]]$After = @(), [string]$Caption = '')
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Margin = '0,0,0,8'
    if ($Caption) {
        $c = New-Object System.Windows.Controls.TextBlock
        $c.Text = $Caption
        $c.Foreground = Get-Brush '#8791A5'
        $c.FontSize = 11.5
        $c.Margin = '1,0,0,4'
        [void]$sp.Children.Add($c)
    }
    $dock = New-Object System.Windows.Controls.DockPanel
    $dock.LastChildFill = $true
    foreach ($a in $After) {
        [System.Windows.Controls.DockPanel]::SetDock($a, 'Right')
        $a.Margin = '6,0,0,0'
        [void]$dock.Children.Add($a)
    }
    [void]$dock.Children.Add($Stretch)
    [void]$sp.Children.Add($dock)
    return $sp
}

function Add-PanelMenu {
    # Menu kontekstowe listy: pozycje @{ Text; Icon; Action = { param($panel, $names) } } albo '-' (separator)
    param([hashtable]$Panel, [object[]]$Items)
    $menu = New-Object System.Windows.Controls.ContextMenu
    foreach ($it in $Items) {
        if ($it -is [string]) { [void]$menu.Items.Add((New-MenuSeparator)); continue }
        $mi = New-MenuItem -Text $it.Text -Icon $it.Icon -Danger:([bool]$it['Danger'])
        $script:PanelMenuActions[$mi] = @{ Panel = $Panel; Action = $it.Action }
        $mi.add_Click($script:PanelMenuClick)
        [void]$menu.Items.Add($mi)
    }
    $Panel.pList.ContextMenu = $menu
}
$script:PanelMenuActions = New-Object 'System.Collections.Generic.Dictionary[object,hashtable]'
$script:PanelMenuClick = {
    param($s, $e)
    try {
        $entry = $script:PanelMenuActions[$s]
        $names = @(Get-TargetSelection -Panel $entry.Panel)
        if ($names.Count -eq 0) { return }
        $null = & $entry.Action $entry.Panel $names
    }
    catch {
        Write-Log "Błąd: $($_.Exception.Message)" 'ERROR' -Module ''
        Show-Error 'Operacja nie powiodła się.' $_
    }
}

function Start-Tool {
    # Uruchamia narzędzie systemowe (konsola MMC, RDP, Eksplorator) - tylko dla poprawnych nazw hostów
    param([string]$FilePath, [string[]]$Arguments, [string]$Name)
    if ($Name -and $Name -notmatch '^[A-Za-z0-9._-]+$') { Show-Warning "Niepoprawna nazwa komputera: $Name"; return }
    try { Start-Process -FilePath $FilePath -ArgumentList $Arguments | Out-Null }
    catch { Show-Error "Nie można uruchomić $FilePath." $_ }
}
#endregion

#region Panel komputerów
function New-ComputerLdapFilter {
    param([string]$NamePattern, [bool]$OnlyEnabled)
    $parts = @('(objectCategory=computer)')
    $pattern = ([string]$NamePattern).Trim()
    if ($pattern) {
        if ($pattern -notmatch '\*') { $pattern = "*$pattern*" }
        # Znaki specjalne LDAP (poza gwiazdką, która jest symbolem wieloznacznym)
        $escaped = $pattern -replace '\\', '\5c' -replace '\(', '\28' -replace '\)', '\29'
        $parts += "(name=$escaped)"
    }
    if ($OnlyEnabled) { $parts += '(!(userAccountControl:1.2.840.113556.1.4.803:=2))' }
    return '(&' + ($parts -join '') + ')'
}

function Initialize-ComputerPanel {
    $p = New-TargetPanel -Kind 'Computers' -Title 'KOMPUTERY' -Placeholder 'Szukaj na liście…' -EmptyIcon 'E7F4' `
        -EmptyText "Lista jest pusta.`nWczytaj komputery z Active Directory albo dodaj nazwy ręcznie." `
        -Columns @('Name', 'OS', 'LastLogon', 'Enabled', 'DNSHostName', 'DN') -KeyColumn 'Name'
    $p.SearchColumns = @('Name', 'OS', 'Source', 'DN', 'Status')
    $p.Describe = {
        param($row)
        $row['Title'] = $row['Name']
        $os = [string]$row['OS']
        $os = $os -replace '^Microsoft ', '' -replace ' (Edition|Evaluation)$', ''
        $parts = @()
        if ($os) { $parts += $os } else { $parts += 'system nieznany' }
        if ([string]$row['Enabled'] -eq 'Nie') { $parts += 'konto wyłączone' }
        if ([string]$row['Status']) { $parts += [string]$row['Status'] }
        $parts += [string]$row['Source']
        $row['Sub'] = $parts -join ' • '
        if ([string]$row['Enabled'] -eq 'Nie' -and [string]$row['Dot'] -eq $script:DotColors.unknown) { $row['Dot'] = $script:DotColors.off }
    }
    $script:UI.ComputerPanel = $p
    $script:UI.HostTable = $p.Table

    $ou = New-Object System.Windows.Controls.TextBox
    $ou.Tag = 'Cała domena'
    $ou.Text = [string]$script:Settings.SearchBase
    $ou.ToolTip = 'Jednostka organizacyjna (DN), z której wczytać komputery'
    $btnOu = New-PlainButton -Icon 'E8B7' -ToolTip 'Wybierz jednostkę organizacyjną'
    $btnOu.Padding = '9,6'
    [void]$p.pSource.Children.Add((New-SourceRow -Stretch $ou -After @($btnOu) -Caption 'Jednostka organizacyjna'))
    $name = New-Object System.Windows.Controls.TextBox
    $name.Tag = 'np. PC-* albo fragment nazwy'
    $name.Text = [string]$script:Settings.NameFilter
    [void]$p.pSource.Children.Add((New-SourceRow -Stretch $name -Caption 'Filtr nazwy'))
    $chk = New-Object System.Windows.Controls.CheckBox
    $chk.Content = 'Tylko włączone konta'
    $chk.IsChecked = [bool]$script:Settings.OnlyEnabled
    $chk.Margin = '1,0,0,10'
    [void]$p.pSource.Children.Add($chk)
    $load = New-PlainButton -Text 'Wczytaj z AD' -Icon 'E896' -Primary
    $add = New-PlainButton -Icon 'E710' -ToolTip 'Dodaj nazwy ręcznie'
    $file = New-PlainButton -Icon 'E8E5' -ToolTip 'Wczytaj nazwy z pliku TXT/CSV'
    foreach ($b in $add, $file) { $b.Padding = '9,6' }
    [void]$p.pSource.Children.Add((New-SourceRow -Stretch $load -After @($file, $add)))
    $p.Ou = $ou
    $p.NameFilter = $name
    $p.OnlyEnabled = $chk

    $tm = New-ModuleContext -Definition @{ Key = '__computers'; Title = 'Lista komputerów'; Category = ''; Icon = 'E7F4'; Workspace = '' }
    $script:UI.Modules['__computers'] = $tm
    [void]$tm.Buttons.Add($load)
    Register-ControlHandler -Control $btnOu -EventName 'Click' -Module $tm -Action {
        $dn = Select-OrganizationalUnit -Title 'Komputery z jednostki organizacyjnej' -Selected $script:UI.ComputerPanel.Ou.Text.Trim() -AllowDomainRoot
        if ($null -ne $dn) { $script:UI.ComputerPanel.Ou.Text = $dn }
    }
    Register-ControlHandler -Control $load -EventName 'Click' -Module $tm -Action { param($m) Start-AdComputerLoad -Module $m }
    Register-ControlHandler -Control $add -EventName 'Click' -Module $tm -Action { Add-ManualComputers }
    Register-ControlHandler -Control $file -EventName 'Click' -Module $tm -Action { Import-ComputerFile }
    $name.add_KeyDown({ param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Return) { Invoke-UiAction -Module $script:UI.Modules['__computers'] -Action { param($m) Start-AdComputerLoad -Module $m } } })

    Add-PanelMenu -Panel $p -Items @(
        @{ Text = 'Zaznacz wybrane'; Icon = 'E73A'; Action = { param($panel) Set-TargetCheck -Panel $panel -Mode CheckSelected } }
        @{ Text = 'Odznacz wybrane'; Icon = 'E739'; Action = { param($panel) Set-TargetCheck -Panel $panel -Mode UncheckSelected } }
        '-'
        @{ Text = 'Pulpit zdalny'; Icon = 'E8AF'; Action = { param($panel, $names) Start-Tool -FilePath 'mstsc.exe' -Arguments @("/v:$($names[0])") -Name $names[0] } }
        @{ Text = 'Zarządzanie komputerem'; Icon = 'E912'; Action = { param($panel, $names) Start-Tool -FilePath 'compmgmt.msc' -Arguments @("/computer:\\$($names[0])") -Name $names[0] } }
        @{ Text = 'Usługi'; Icon = 'E9F5'; Action = { param($panel, $names) Start-Tool -FilePath 'services.msc' -Arguments @("/computer:\\$($names[0])") -Name $names[0] } }
        @{ Text = 'Podgląd zdarzeń'; Icon = 'E81C'; Action = { param($panel, $names) Start-Tool -FilePath 'eventvwr.exe' -Arguments @("\\$($names[0])") -Name $names[0] } }
        @{ Text = 'Udział C$'; Icon = 'E838'; Action = { param($panel, $names) Start-Tool -FilePath 'explorer.exe' -Arguments @("\\$($names[0])\C$") -Name $names[0] } }
        @{ Text = 'Ping ciągły (konsola)'; Icon = 'E968'; Action = { param($panel, $names) Start-Tool -FilePath 'cmd.exe' -Arguments @('/k', "ping -t $($names[0])") -Name $names[0] } }
        '-'
        @{ Text = 'Kopiuj nazwy'; Icon = 'E8C8'; Action = { param($panel, $names) Set-ClipboardText ($names -join [Environment]::NewLine); Show-Toast "Skopiowano nazw: $($names.Count)" 'ok' } }
        @{ Text = 'Usuń z listy'; Icon = 'E74D'; Danger = $true; Action = { param($panel) Remove-TargetRows -Panel $panel } }
    )
    return $p
}

function Start-AdComputerLoad {
    param([hashtable]$Module)
    $p = $script:UI.ComputerPanel
    $script:Settings.SearchBase = $p.Ou.Text.Trim()
    $script:Settings.NameFilter = $p.NameFilter.Text.Trim()
    $script:Settings.OnlyEnabled = ($p.OnlyEnabled.IsChecked -eq $true)
    $params = @{
        LdapFilter = New-ComputerLdapFilter -NamePattern $script:Settings.NameFilter -OnlyEnabled $script:Settings.OnlyEnabled
        SearchBase = $script:Settings.SearchBase
    }
    Start-HostOperation -Module $Module -Name 'Wczytywanie komputerów z AD' -Targets @('Active Directory') -Local -Output None -Parameters $params -ScriptBlock {
        param($Target, $P, $Ctx)
        Import-Module ActiveDirectory -ErrorAction Stop -Verbose:$false
        $q = @{
            LDAPFilter  = $P.LdapFilter
            Properties  = @('OperatingSystem', 'LastLogonDate', 'Enabled', 'DNSHostName')
            ErrorAction = 'Stop'
        }
        if ($P.SearchBase) { $q.SearchBase = $P.SearchBase }
        if ($Ctx.Server) { $q.Server = $Ctx.Server }
        if ($Ctx.Credential) { $q.Credential = $Ctx.Credential }
        Get-ADComputer @q | ForEach-Object {
            [pscustomobject]@{
                Name        = $_.Name
                OS          = $_.OperatingSystem
                LastLogon   = $_.LastLogonDate
                Enabled     = $_.Enabled
                DNSHostName = $_.DNSHostName
                DN          = $_.DistinguishedName
            }
        }
    } -OnResult {
        param($m, $r)
        if (-not $r.Ok) {
            $m.Data.LastError = (@($r.Errors)) -join "`r`n"
            Invoke-Deferred -Module $m -Action { param($m) Show-Error 'Nie udało się wczytać komputerów z Active Directory.' $m.Data.LastError }
            return
        }
        $items = @($r.Data)
        $added = Import-TargetRows -Panel $script:UI.ComputerPanel -Items $items -Source 'AD' -ReplaceSource
        Write-Log ("Wczytano z AD {0} komputerów (nowych na liście: {1})." -f $items.Count, $added) 'OK'
    }
}

function Read-NameList {
    # Nazwy z tekstu (wiersze, przecinki, spacje)
    param([string]$Text)
    return @($Text -split '[\s,;]+' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ } | Select-Object -Unique)
}

function Add-ManualComputers {
    $text = Show-InputDialog -Title 'Dodaj komputery' -Prompt 'Wpisz lub wklej nazwy komputerów – po jednej w wierszu albo rozdzielone przecinkiem lub spacją. Dodane pozycje zostaną zaznaczone.' -Multiline -Icon 'E710'
    if ($null -eq $text) { return }
    $names = @(Read-NameList $text)
    if ($names.Count -eq 0) { return }
    $added = Import-TargetRows -Panel $script:UI.ComputerPanel -Items @($names | ForEach-Object { [pscustomobject]@{ Name = $_ } }) -Source 'ręcznie' -Check
    Write-Log ("Dodano ręcznie {0} komputerów (nowych: {1})." -f $names.Count, $added) 'OK' -Module 'Komputery'
}

function Read-NameFile {
    # Pierwsza kolumna pliku TXT/CSV (bez wiersza nagłówka)
    param([string]$HeaderPattern)
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'Pliki tekstowe i CSV (*.txt;*.csv)|*.txt;*.csv|Wszystkie pliki (*.*)|*.*'
    $answer = if ($script:UI.Window) { $dlg.ShowDialog($script:UI.Window) } else { $dlg.ShowDialog() }
    if ($answer -ne $true) { return $null }
    $names = foreach ($line in (Get-Content -LiteralPath $dlg.FileName -Encoding UTF8)) {
        $first = (($line -split '[,;\t]')[0]).Trim().Trim('"')
        if ($first -and $first -notmatch $HeaderPattern) { $first }
    }
    # Przecinek: pusta lista ma pozostać tablicą (nie $null, który oznacza anulowanie)
    return , @($names | Select-Object -Unique)
}

function Import-ComputerFile {
    $names = Read-NameFile -HeaderPattern '^(name|computer|computername|hostname|nazwa|komputer)$'
    if ($null -eq $names) { return }
    if ($names.Count -eq 0) { Show-Warning 'Plik nie zawiera nazw komputerów.'; return }
    $added = Import-TargetRows -Panel $script:UI.ComputerPanel -Items @($names | ForEach-Object { [pscustomobject]@{ Name = $_ } }) -Source 'plik' -Check
    Write-Log ("Wczytano z pliku {0} komputerów (nowych: {1})." -f $names.Count, $added) 'OK' -Module 'Komputery'
}

function Set-ComputerState {
    # Kolor kropki i opis stanu komputera na liście (np. po teście łączności)
    param([string]$Name, [ValidateSet('unknown', 'ok', 'warn', 'crit', 'off')][string]$State, [string]$Status = '')
    $p = $script:UI.ComputerPanel
    if (-not $p) { return }
    $row = $p.Table.Rows.Find($Name)
    if (-not $row) { return }
    $row['Dot'] = $script:DotColors[$State]
    $row['Status'] = ConvertTo-DbValue $Status
    & $p.Describe $row
}

function Rename-ComputerRow {
    # Po zmianie nazwy komputera aktualizuje pozycję na liście
    param([string]$OldName, [string]$NewName)
    $t = $script:UI.HostTable
    $row = $t.Rows.Find($OldName)
    if (-not $row -or $t.Rows.Find($NewName)) { return }
    $row['Name'] = $NewName
    & $script:UI.ComputerPanel.Describe $row
    $t.AcceptChanges()
}
#endregion

#region Panel użytkowników
function New-UserLdapFilter {
    param([string]$Query, [int]$State)
    $parts = @('(objectCategory=person)', '(objectClass=user)')
    $q = ([string]$Query).Trim()
    if ($q) {
        if ($q -match '\*') {
            $v = $q -replace '\\', '\5c' -replace '\(', '\28' -replace '\)', '\29'
            $parts += "(|(sAMAccountName=$v)(displayName=$v)(userPrincipalName=$v)(mail=$v))"
        }
        else {
            $v = ConvertTo-LdapFilterValue $q
            $parts += "(|(anr=$v)(sAMAccountName=$v*)(userPrincipalName=$v*)(employeeID=$v))"
        }
    }
    switch ($State) {
        1 { $parts += '(!(userAccountControl:1.2.840.113556.1.4.803:=2))' }
        2 { $parts += '(userAccountControl:1.2.840.113556.1.4.803:=2)' }
        3 { $parts += '(lockoutTime>=1)' }
    }
    return '(&' + ($parts -join '') + ')'
}

function Initialize-UserPanel {
    $p = New-TargetPanel -Kind 'Users' -Title 'UŻYTKOWNICY' -Placeholder 'Szukaj na liście…' -EmptyIcon 'E716' `
        -EmptyText "Lista jest pusta.`nWyszukaj konta w Active Directory albo dodaj loginy ręcznie." `
        -Columns @('Login', 'Name', 'Upn', 'Mail', 'Enabled', 'Locked', 'Department', 'Title_', 'LastLogon', 'DN') -KeyColumn 'Login'
    $p.SearchColumns = @('Login', 'Name', 'Upn', 'Mail', 'Department', 'Title_', 'Source', 'DN')
    $p.Describe = {
        param($row)
        $display = [string]$row['Name']
        $row['Title'] = if ($display) { $display } else { $row['Login'] }
        $parts = @([string]$row['Login'])
        if ([string]$row['Title_']) { $parts += [string]$row['Title_'] }
        elseif ([string]$row['Department']) { $parts += [string]$row['Department'] }
        if ([string]$row['Locked'] -eq 'Tak') { $parts += 'zablokowane' }
        elseif ([string]$row['Enabled'] -eq 'Nie') { $parts += 'wyłączone' }
        if ([string]$row['Source'] -ne 'AD') { $parts += [string]$row['Source'] }
        $row['Sub'] = $parts -join ' • '
        $row['Dot'] = if ([string]$row['Locked'] -eq 'Tak') { $script:DotColors.crit }
        elseif ([string]$row['Enabled'] -eq 'Nie') { $script:DotColors.off }
        elseif ([string]$row['Enabled'] -eq 'Tak') { $script:DotColors.ok }
        else { $script:DotColors.unknown }
    }
    $script:UI.UserPanel = $p
    $script:UI.UserTable = $p.Table

    $query = New-Object System.Windows.Controls.TextBox
    $query.Tag = 'Login, nazwisko lub e-mail'
    $query.ToolTip = 'Puste pole = wszystkie konta z wybranej jednostki. Można używać gwiazdki, np. jan*'
    [void]$p.pSource.Children.Add((New-SourceRow -Stretch $query -Caption 'Wyszukaj konta'))
    $ou = New-Object System.Windows.Controls.TextBox
    $ou.Tag = 'Cała domena'
    $ou.Text = [string]$script:Settings.UserSearchBase
    $btnOu = New-PlainButton -Icon 'E8B7' -ToolTip 'Wybierz jednostkę organizacyjną'
    $btnOu.Padding = '9,6'
    [void]$p.pSource.Children.Add((New-SourceRow -Stretch $ou -After @($btnOu) -Caption 'Jednostka organizacyjna'))
    $state = New-Object System.Windows.Controls.ComboBox
    foreach ($i in 'Wszystkie konta', 'Tylko aktywne', 'Tylko wyłączone', 'Tylko zablokowane') { [void]$state.Items.Add($i) }
    $state.SelectedIndex = [int]$script:Settings.UserFilter
    [void]$p.pSource.Children.Add((New-SourceRow -Stretch $state -Caption 'Stan konta'))
    $find = New-PlainButton -Text 'Szukaj w AD' -Icon 'E721' -Primary
    $add = New-PlainButton -Icon 'E710' -ToolTip 'Dodaj loginy ręcznie'
    $file = New-PlainButton -Icon 'E8E5' -ToolTip 'Wczytaj loginy z pliku TXT/CSV'
    foreach ($b in $add, $file) { $b.Padding = '9,6' }
    [void]$p.pSource.Children.Add((New-SourceRow -Stretch $find -After @($file, $add)))
    $p.Query = $query
    $p.Ou = $ou
    $p.State = $state

    $tm = New-ModuleContext -Definition @{ Key = '__users'; Title = 'Lista użytkowników'; Category = ''; Icon = 'E716'; Workspace = '' }
    $script:UI.Modules['__users'] = $tm
    [void]$tm.Buttons.Add($find)
    Register-ControlHandler -Control $btnOu -EventName 'Click' -Module $tm -Action {
        $dn = Select-OrganizationalUnit -Title 'Konta z jednostki organizacyjnej' -Selected $script:UI.UserPanel.Ou.Text.Trim() -AllowDomainRoot
        if ($null -ne $dn) { $script:UI.UserPanel.Ou.Text = $dn }
    }
    Register-ControlHandler -Control $find -EventName 'Click' -Module $tm -Action { param($m) Start-AdUserLoad -Module $m }
    Register-ControlHandler -Control $add -EventName 'Click' -Module $tm -Action { param($m) Add-ManualUsers -Module $m }
    Register-ControlHandler -Control $file -EventName 'Click' -Module $tm -Action { param($m) Import-UserFile -Module $m }
    $query.add_KeyDown({ param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Return) { Invoke-UiAction -Module $script:UI.Modules['__users'] -Action { param($m) Start-AdUserLoad -Module $m } } })

    Add-PanelMenu -Panel $p -Items @(
        @{ Text = 'Zaznacz wybrane'; Icon = 'E73A'; Action = { param($panel) Set-TargetCheck -Panel $panel -Mode CheckSelected } }
        @{ Text = 'Odznacz wybrane'; Icon = 'E739'; Action = { param($panel) Set-TargetCheck -Panel $panel -Mode UncheckSelected } }
        '-'
        @{ Text = 'Kopiuj loginy'; Icon = 'E8C8'; Action = { param($panel, $names) Set-ClipboardText ($names -join [Environment]::NewLine); Show-Toast "Skopiowano loginów: $($names.Count)" 'ok' } }
        @{ Text = 'Kopiuj adresy e-mail'; Icon = 'E715'; Action = {
                param($panel, $names)
                $mails = @(foreach ($n in $names) { $r = $panel.Table.Rows.Find($n); if ($r -and [string]$r['Mail']) { [string]$r['Mail'] } })
                if ($mails.Count -eq 0) { Show-Warning 'Wybrane konta nie mają adresów e-mail.'; return }
                Set-ClipboardText ($mails -join '; ')
                Show-Toast "Skopiowano adresów: $($mails.Count)" 'ok'
            }
        }
        @{ Text = 'Usuń z listy'; Icon = 'E74D'; Danger = $true; Action = { param($panel) Remove-TargetRows -Panel $panel } }
    )
    return $p
}

$script:UserLoadScript = {
    # Wspólny blok: wyszukiwanie (LdapFilter) albo uzupełnianie danych podanych loginów (Logins)
    param($Target, $P, $Ctx)
    Import-Module ActiveDirectory -ErrorAction Stop -Verbose:$false
    $ad = @{ ErrorAction = 'Stop' }
    if ($Ctx.Server) { $ad.Server = $Ctx.Server }
    if ($Ctx.Credential) { $ad.Credential = $Ctx.Credential }
    $props = @('DisplayName', 'UserPrincipalName', 'mail', 'Enabled', 'LockedOut', 'Department', 'Title', 'LastLogonDate')
    $users = @()
    if ($P.Logins) {
        $users = foreach ($login in $P.Logins) {
            try { Get-ADUser -Identity $login -Properties $props @ad }
            catch { Write-Error "Nie znaleziono konta $login`: $($_.Exception.Message)" }
        }
    }
    else {
        $q = @{ LDAPFilter = $P.LdapFilter; Properties = $props; ResultSetSize = 20000 }
        if ($P.SearchBase) { $q.SearchBase = $P.SearchBase }
        $users = @(Get-ADUser @q @ad)
        if ($P.OnlyLocked) { $users = @($users | Where-Object { $_.LockedOut }) }
    }
    foreach ($u in $users) {
        if (-not $u) { continue }
        [pscustomobject]@{
            Login      = $u.SamAccountName
            Name       = $u.DisplayName
            Upn        = $u.UserPrincipalName
            Mail       = $u.mail
            Enabled    = [bool]$u.Enabled
            Locked     = [bool]$u.LockedOut
            Department = $u.Department
            Title_     = $u.Title
            LastLogon  = $u.LastLogonDate
            DN         = $u.DistinguishedName
        }
    }
}

function Start-AdUserLoad {
    param([hashtable]$Module, [string[]]$Logins = @(), [string]$Source = 'AD', [switch]$Check)
    $p = $script:UI.UserPanel
    $params = @{ Logins = @($Logins) }
    if ($Logins.Count -eq 0) {
        $state = [int]$p.State.SelectedIndex
        $script:Settings.UserFilter = $state
        $script:Settings.UserSearchBase = $p.Ou.Text.Trim()
        $params.LdapFilter = New-UserLdapFilter -Query $p.Query.Text -State $state
        $params.SearchBase = $script:Settings.UserSearchBase
        $params.OnlyLocked = ($state -eq 3)
    }
    $Module.Data.Source = $Source
    $Module.Data.Check = [bool]$Check
    $Module.Data.Replace = ($Logins.Count -eq 0)
    $title = if ($Logins.Count -gt 0) { 'Uzupełnianie danych kont z AD' } else { 'Wyszukiwanie kont w AD' }
    Start-HostOperation -Module $Module -Name $title -Targets @('Active Directory') -Local -Output None -Parameters $params -ScriptBlock $script:UserLoadScript -OnResult {
        param($m, $r)
        if (-not $r.Ok) {
            $m.Data.LastError = (@($r.Errors)) -join "`r`n"
            Invoke-Deferred -Module $m -Action { param($m) Show-Error 'Nie udało się pobrać kont z Active Directory.' $m.Data.LastError }
            return
        }
        $items = @($r.Data)
        $panel = $script:UI.UserPanel
        if ($m.Data.Replace) {
            # Nowe wyszukiwanie zastępuje poprzednie wyniki (zaznaczone pozycje zostają)
            $added = Import-TargetRows -Panel $panel -Items $items -Source 'AD' -ReplaceSource
        }
        else {
            $added = Import-TargetRows -Panel $panel -Items $items -Source 'AD' -Check:$m.Data.Check
        }
        if ($items.Count -ge 20000) { Write-Log 'Wyświetlono pierwsze 20000 kont - zawęź wyszukiwanie.' 'WARN' }
        Write-Log ("Znaleziono w AD {0} kont (nowych na liście: {1})." -f $items.Count, $added) 'OK'
    }
}

function Add-UserLogins {
    param([hashtable]$Module, [string[]]$Names, [string]$Source)
    if ($Names.Count -eq 0) { return }
    $logins = @($Names | ForEach-Object { ($_ -replace '^.*\\', '') -replace '@.*$', '' } | Where-Object { $_ } | Select-Object -Unique)
    [void](Import-TargetRows -Panel $script:UI.UserPanel -Items @($logins | ForEach-Object { [pscustomobject]@{ Login = $_ } }) -Source $Source -Check)
    Write-Log ("Dodano {0} kont ({1}); pobieranie danych z AD…" -f $logins.Count, $Source) 'INFO' -Module 'Użytkownicy'
    if (Get-Module -ListAvailable -Name ActiveDirectory) { Start-AdUserLoad -Module $Module -Logins $logins -Source $Source -Check }
}

function Add-ManualUsers {
    param([hashtable]$Module)
    $text = Show-InputDialog -Title 'Dodaj konta' -Prompt 'Wpisz lub wklej loginy (sAMAccountName, DOMENA\login albo UPN) – po jednym w wierszu albo rozdzielone przecinkiem lub spacją.' -Multiline -Icon 'E8FA'
    if ($null -eq $text) { return }
    Add-UserLogins -Module $Module -Names (Read-NameList $text) -Source 'ręcznie'
}

function Import-UserFile {
    param([hashtable]$Module)
    $names = Read-NameFile -HeaderPattern '^(login|samaccountname|user|username|użytkownik|uzytkownik|konto)$'
    if ($null -eq $names) { return }
    if ($names.Count -eq 0) { Show-Warning 'Plik nie zawiera loginów.'; return }
    Add-UserLogins -Module $Module -Names $names -Source 'plik'
}

function Update-UserRow {
    # Aktualizacja pozycji listy po zmianie w AD (np. odblokowanie, wyłączenie konta)
    param([string]$Login, [hashtable]$Values)
    $p = $script:UI.UserPanel
    if (-not $p) { return }
    $row = $p.Table.Rows.Find($Login)
    if (-not $row) { return }
    foreach ($k in $Values.Keys) {
        if (-not $p.Table.Columns.Contains($k)) { continue }
        $v = $Values[$k]
        $row[$k] = if ($v -is [bool]) { $(if ($v) { 'Tak' } else { 'Nie' }) } else { ConvertTo-DbValue $v }
    }
    & $p.Describe $row
}
#endregion

#region Okno główne
$script:MainXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Domain Ops" Width="1560" Height="940" MinWidth="1100" MinHeight="680"
        WindowStartupLocation="CenterScreen" Background="#0F1318" Foreground="#E4E8EF"
        FontFamily="Segoe UI" FontSize="13" UseLayoutRounding="True" SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display">
  <Grid Background="#0F1318">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*" MinHeight="300"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition x:Name="logRow" Height="0"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Background="#12171D" BorderBrush="#1E242E" BorderThickness="0,0,0,1" Padding="16,10">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <Border Width="36" Height="36" CornerRadius="10">
            <Border.Background>
              <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                <GradientStop Color="#4C7DF0" Offset="0"/>
                <GradientStop Color="#8A5CF0" Offset="1"/>
              </LinearGradientBrush>
            </Border.Background>
            <TextBlock Style="{StaticResource Glyph}" Text="&#xE968;" FontSize="17" Foreground="White" HorizontalAlignment="Center"/>
          </Border>
          <StackPanel Margin="11,0,0,0" VerticalAlignment="Center">
            <StackPanel Orientation="Horizontal">
              <TextBlock Text="Domain Ops" FontSize="15" FontWeight="SemiBold" Foreground="White"/>
              <TextBlock x:Name="txtVersion" Foreground="#5E6779" FontSize="11" Margin="7,3,0,0"/>
            </StackPanel>
            <TextBlock x:Name="txtDomain" Foreground="#7B8496" FontSize="11.5"/>
          </StackPanel>
        </StackPanel>
        <Border Grid.Column="1" Style="{StaticResource SegmentHost}" Margin="36,0,0,0" VerticalAlignment="Center">
          <StackPanel x:Name="wsSwitcher" Orientation="Horizontal"/>
        </Border>
        <TextBlock x:Name="txtSubtitle" Grid.Column="2" Foreground="#5E6779" FontSize="12" VerticalAlignment="Center" Margin="18,0,12,0" TextTrimming="CharacterEllipsis"/>
        <StackPanel Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Center">
          <Button x:Name="btnCred" Style="{StaticResource GhostButton}" ToolTip="Konto używane do operacji (kliknij, aby zmienić)" Padding="10,5">
            <StackPanel Orientation="Horizontal">
              <Border Width="26" Height="26" CornerRadius="13" Background="#1F2B47" Margin="0,0,9,0">
                <TextBlock Style="{StaticResource Glyph}" Text="&#xE77B;" FontSize="12" Foreground="#8CB0FF" HorizontalAlignment="Center"/>
              </Border>
              <StackPanel VerticalAlignment="Center">
                <TextBlock x:Name="txtCredUser" Foreground="#E4E8EF" FontSize="12.5"/>
                <TextBlock x:Name="txtCredMode" Foreground="#7B8496" FontSize="11"/>
              </StackPanel>
              <TextBlock Style="{StaticResource Glyph}" Text="&#xE70D;" FontSize="10" Foreground="#7B8496" Margin="10,0,0,0"/>
            </StackPanel>
          </Button>
          <Button x:Name="btnSettings" Style="{StaticResource GhostButton}" ToolTip="Ustawienia" Margin="4,0,0,0" Padding="10,7">
            <TextBlock Style="{StaticResource Glyph}" Text="&#xE713;" FontSize="15"/>
          </Button>
        </StackPanel>
      </Grid>
    </Border>

    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition x:Name="targetColumn" Width="300"/>
        <ColumnDefinition Width="248"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <Grid x:Name="targetHost"/>
      <Border Grid.Column="1" BorderBrush="#1E242E" BorderThickness="0,0,1,0">
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <Grid x:Name="navHost" Margin="10,4,10,14"/>
        </ScrollViewer>
      </Border>
      <Grid Grid.Column="2">
        <Grid x:Name="contentHost" Margin="26,20,26,16"/>
        <StackPanel x:Name="toastHost" HorizontalAlignment="Right" VerticalAlignment="Bottom" Margin="0,0,26,22"/>
      </Grid>
    </Grid>

    <GridSplitter x:Name="logSplitter" Grid.Row="2" Height="5" HorizontalAlignment="Stretch" ResizeDirection="Rows"
                  ResizeBehavior="PreviousAndNext" Visibility="Collapsed"/>
    <Border x:Name="logPanel" Grid.Row="3" Background="#12171D" BorderBrush="#1E242E" BorderThickness="0,1,0,0" Visibility="Collapsed">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Grid Margin="18,6,10,2">
          <TextBlock Text="DZIENNIK OPERACJI" Foreground="#5E6779" FontSize="11" FontWeight="SemiBold" VerticalAlignment="Center"/>
          <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnLogCopy" Style="{StaticResource GhostButton}" ToolTip="Kopiuj dziennik" MinHeight="26" Padding="8,2"/>
            <Button x:Name="btnLogFile" Style="{StaticResource GhostButton}" ToolTip="Otwórz plik dziennika" MinHeight="26" Padding="8,2"/>
            <Button x:Name="btnLogClear" Style="{StaticResource GhostButton}" ToolTip="Wyczyść okno dziennika" MinHeight="26" Padding="8,2"/>
            <Button x:Name="btnLogHide" Style="{StaticResource GhostButton}" ToolTip="Ukryj dziennik (Ctrl+L)" MinHeight="26" Padding="8,2"/>
          </StackPanel>
        </Grid>
        <ListBox x:Name="logList" Grid.Row="1" Margin="8,0,8,6" SelectionMode="Extended" FontSize="12.5">
          <ListBox.ItemContainerStyle>
            <Style TargetType="ListBoxItem" BasedOn="{StaticResource ItemCard}">
              <Setter Property="Padding" Value="8,2"/>
              <Setter Property="Margin" Value="0"/>
            </Style>
          </ListBox.ItemContainerStyle>
          <ListBox.ItemTemplate>
            <DataTemplate>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="62"/>
                  <ColumnDefinition Width="70"/>
                  <ColumnDefinition Width="Auto" MaxWidth="230"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="{Binding Time}" Foreground="#5E6779" FontFamily="Consolas" VerticalAlignment="Center"/>
                <Border x:Name="pill" Grid.Column="1" CornerRadius="8" Padding="7,0" HorizontalAlignment="Left" Background="#1E252F" VerticalAlignment="Center">
                  <TextBlock x:Name="lvl" Text="{Binding Level}" FontSize="10.5" FontWeight="SemiBold" Foreground="#AEB6C4"/>
                </Border>
                <TextBlock Grid.Column="2" Text="{Binding Module}" Foreground="#8CB0FF" Margin="0,0,12,0" TextTrimming="CharacterEllipsis" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="3" Text="{Binding Message}" TextWrapping="Wrap" VerticalAlignment="Center"/>
              </Grid>
              <DataTemplate.Triggers>
                <DataTrigger Binding="{Binding Level}" Value="OK">
                  <Setter TargetName="pill" Property="Background" Value="#15291F"/>
                  <Setter TargetName="lvl" Property="Foreground" Value="#5EE3AE"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Level}" Value="WARN">
                  <Setter TargetName="pill" Property="Background" Value="#2E2616"/>
                  <Setter TargetName="lvl" Property="Foreground" Value="#FFC46B"/>
                </DataTrigger>
                <DataTrigger Binding="{Binding Level}" Value="ERROR">
                  <Setter TargetName="pill" Property="Background" Value="#2E1A1E"/>
                  <Setter TargetName="lvl" Property="Foreground" Value="#FF7A86"/>
                </DataTrigger>
              </DataTemplate.Triggers>
            </DataTemplate>
          </ListBox.ItemTemplate>
        </ListBox>
      </Grid>
    </Border>

    <Border Grid.Row="4" Background="#12171D" BorderBrush="#1E242E" BorderThickness="0,1,0,0" Padding="16,4,10,4">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Ellipse x:Name="statusDot" Width="8" Height="8" Fill="#5EE3AE" VerticalAlignment="Center"/>
        <TextBlock x:Name="txtStatus" Grid.Column="1" Text="Gotowe" Margin="10,0,0,0" Foreground="#AEB6C4" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
        <ProgressBar x:Name="prgStatus" Grid.Column="2" Width="200" Margin="12,0" VerticalAlignment="Center" Visibility="Collapsed"/>
        <Button x:Name="btnCancelOps" Grid.Column="3" Style="{StaticResource DangerButton}" Padding="10,2" MinHeight="26" Visibility="Collapsed" ToolTip="Przerwij wszystkie trwające operacje"/>
        <Button x:Name="btnLogToggle" Grid.Column="4" Style="{StaticResource GhostButton}" MinHeight="26" Padding="8,2" Margin="8,0,0,0" ToolTip="Dziennik operacji (Ctrl+L)">
          <StackPanel Orientation="Horizontal">
            <TextBlock Style="{StaticResource Glyph}" Text="&#xE81C;" FontSize="12" Margin="0,0,7,0"/>
            <TextBlock Text="Dziennik" FontSize="12" VerticalAlignment="Center"/>
            <Border x:Name="logBadge" Background="#D9475A" CornerRadius="8" Padding="6,0" Margin="7,0,0,0" Visibility="Collapsed" VerticalAlignment="Center">
              <TextBlock x:Name="logBadgeText" Foreground="White" FontSize="10.5" FontWeight="SemiBold"/>
            </Border>
          </StackPanel>
        </Button>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

function New-MainWindow {
    $w = New-UiElement $script:MainXaml
    $script:UI.Window = $w
    $c = $script:UI.Controls
    foreach ($n in 'txtVersion', 'txtDomain', 'wsSwitcher', 'txtSubtitle', 'btnCred', 'txtCredUser', 'txtCredMode', 'btnSettings',
        'targetColumn', 'targetHost', 'navHost', 'contentHost', 'toastHost', 'logRow', 'logSplitter', 'logPanel', 'logList',
        'btnLogCopy', 'btnLogFile', 'btnLogClear', 'btnLogHide', 'statusDot', 'txtStatus', 'prgStatus', 'btnCancelOps', 'btnLogToggle',
        'logBadge', 'logBadgeText') {
        $c[$n] = $w.FindName($n)
    }
    $c.txtVersion.Text = "v$($script:AppVersion)"
    $c.txtDomain.Text = if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN.ToLowerInvariant() } else { 'komputer spoza domeny' }
    $c.btnCancelOps.Content = New-IconContent -Text 'Przerwij' -Icon 'E71A' -IconSize 11
    $c.btnLogCopy.Content = New-IconContent -Text '' -Icon 'E8C8' -IconSize 12
    $c.btnLogFile.Content = New-IconContent -Text '' -Icon 'E8E5' -IconSize 12
    $c.btnLogClear.Content = New-IconContent -Text '' -Icon 'E894' -IconSize 12
    $c.btnLogHide.Content = New-IconContent -Text '' -Icon 'E70D' -IconSize 12
    $c.logList.ItemsSource = $script:UI.LogItems

    # Rozmiar z ustawień, ale nie większy niż obszar roboczy ekranu (np. laptop 1366x768)
    $work = [System.Windows.SystemParameters]::WorkArea
    if ($work.Width -gt 0 -and $work.Height -gt 0) {
        $w.MinWidth = [Math]::Min($w.MinWidth, $work.Width)
        $w.MinHeight = [Math]::Min($w.MinHeight, $work.Height)
        $w.Width = [Math]::Min([double]$script:Settings.WindowWidth, $work.Width)
        $w.Height = [Math]::Min([double]$script:Settings.WindowHeight, $work.Height)
    }
    else {
        $w.Width = [double]$script:Settings.WindowWidth
        $w.Height = [double]$script:Settings.WindowHeight
    }
    if ($script:Settings.WindowMaximized) { $w.WindowState = 'Maximized' }

    $c.btnSettings.add_Click({ Invoke-UiAction -Module $null -Action { if (Show-SettingsDialog) { Show-Toast 'Zapisano ustawienia.' 'ok' } } })
    $c.btnCred.add_Click({ param($s, $e) Show-CredentialMenu -Anchor $s })
    $c.btnCancelOps.add_Click({ Stop-AllOperations })
    $c.btnLogToggle.add_Click({ Set-LogVisible (-not $script:Settings.LogVisible) })
    $c.btnLogHide.add_Click({ Set-LogVisible $false })
    $c.btnLogClear.add_Click({ $script:UI.LogItems.Clear() })
    $c.btnLogFile.add_Click({
            if (Test-Path -LiteralPath $script:App.LogFile) { Start-Process -FilePath 'notepad.exe' -ArgumentList ('"{0}"' -f $script:App.LogFile) }
            else { Open-Folder $script:App.LogDir }
        })
    $c.btnLogCopy.add_Click({
            $items = @($script:UI.Controls.logList.SelectedItems)
            if ($items.Count -le 1) { $items = @($script:UI.LogItems) }
            $text = (@($items | ForEach-Object { '{0} [{1}] {2}{3}' -f $_.Time, $_.Level, $(if ($_.Module) { "[$($_.Module)] " } else { '' }), $_.Message }) -join [Environment]::NewLine)
            if ($text) { Set-ClipboardText $text; Show-Toast "Skopiowano wpisów: $($items.Count)" 'ok' }
        })
    $w.add_PreviewKeyDown($script:ShellEvents.PreviewKeyDown)
    $w.add_Closing($script:ShellEvents.Closing)
    $w.add_SourceInitialized({ param($s, $e) Set-DarkTitleBar $s })

    Update-CredentialLabel
    Update-StatusBar
    return $w
}

$script:ShellEvents = @{
    PreviewKeyDown = {
        param($s, $e)
        try {
            # Porównania z typami wyliczeniowymi wprost: ModifierKeys ma własny konwerter tekstu ('None' nie jest rozpoznawane)
            $mods = [System.Windows.Input.Keyboard]::Modifiers
            $none = [System.Windows.Input.ModifierKeys]::None
            $ctrl = [System.Windows.Input.ModifierKeys]::Control
            $key = $e.Key
            $m = $script:UI.ActiveModule
            if ($key -eq [System.Windows.Input.Key]::F5 -and $mods -eq $none) {
                if ($m -and $m.PrimaryButton -and $m.PrimaryButton.IsEnabled) {
                    $b = $m.PrimaryButton
                    $b.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, $b)))
                }
                $e.Handled = $true
            }
            elseif ($key -eq [System.Windows.Input.Key]::F -and $mods -eq $ctrl) {
                if ($m -and $m.FilterBox) { [void]$m.FilterBox.Focus(); $m.FilterBox.SelectAll() }
                $e.Handled = $true
            }
            elseif ($key -eq [System.Windows.Input.Key]::L -and $mods -eq $ctrl) {
                Set-LogVisible (-not $script:Settings.LogVisible)
                $e.Handled = $true
            }
            elseif ($mods -eq $ctrl -and [int]$key -ge [int][System.Windows.Input.Key]::D1 -and [int]$key -le [int][System.Windows.Input.Key]::D9) {
                $index = [int]$key - [int][System.Windows.Input.Key]::D1
                $keys = @($script:UI.Workspaces.Keys)
                if ($index -lt $keys.Count) { Show-Workspace -Key $keys[$index] }
                $e.Handled = $true
            }
        }
        catch { Write-Log "Błąd skrótu klawiszowego: $($_.Exception.Message)" 'ERROR' -Module '' }
    }
    Closing        = {
        param($s, $e)
        try {
            if ($script:Engine.Operations.Count -gt 0) {
                if (-not (Confirm-Action -Text 'Trwają operacje w tle. Zamknięcie programu je przerwie. Zamknąć mimo to?' -Title 'Zamknięcie programu' -ConfirmText 'Zamknij')) {
                    $e.Cancel = $true
                    return
                }
            }
            $script:Settings.WindowMaximized = ($s.WindowState -eq 'Maximized')
            if ($s.WindowState -eq 'Normal') {
                $script:Settings.WindowWidth = [int]$s.ActualWidth
                $script:Settings.WindowHeight = [int]$s.ActualHeight
            }
            if ($script:Settings.LogVisible) { $script:Settings.LogHeight = [int]$script:UI.Controls.logRow.ActualHeight }
            Export-Settings
            Clear-ClipboardSecret
        }
        catch { }
    }
}

function Set-LogVisible {
    param([bool]$Visible)
    $c = $script:UI.Controls
    if (-not $Visible -and $script:Settings.LogVisible -and $c.logRow.ActualHeight -gt 60) {
        $script:Settings.LogHeight = [int]$c.logRow.ActualHeight
    }
    $script:Settings.LogVisible = $Visible
    if ($Visible) {
        $c.logRow.Height = New-Object System.Windows.GridLength ([double][Math]::Max(90, [int]$script:Settings.LogHeight))
        $c.logPanel.Visibility = 'Visible'
        $c.logSplitter.Visibility = 'Visible'
        $script:UI.LogUnread = 0
        Update-LogBadge
        if ($script:UI.LogItems.Count -gt 0) { $c.logList.ScrollIntoView($script:UI.LogItems[$script:UI.LogItems.Count - 1]) }
    }
    else {
        $c.logRow.Height = New-Object System.Windows.GridLength 0
        $c.logPanel.Visibility = 'Collapsed'
        $c.logSplitter.Visibility = 'Collapsed'
    }
}

function Update-CredentialLabel {
    $c = $script:UI.Controls
    if (-not $c['txtCredUser']) { return }
    $cred = Get-EffectiveCredential
    if ($cred) {
        $c.txtCredUser.Text = $cred.UserName
        $c.txtCredMode.Text = 'konto alternatywne'
        $c.txtCredMode.Foreground = Get-Brush '#FFC46B'
    }
    else {
        $c.txtCredUser.Text = if ($env:USERDOMAIN) { "$env:USERDOMAIN\$env:USERNAME" } else { [string]$env:USERNAME }
        $c.txtCredMode.Text = 'bieżące konto'
        $c.txtCredMode.Foreground = Get-Brush '#7B8496'
    }
}

function Show-CredentialMenu {
    param($Anchor)
    $menu = New-Object System.Windows.Controls.ContextMenu
    $current = New-MenuItem -Text ('Bieżące konto ({0}\{1})' -f $env:USERDOMAIN, $env:USERNAME) -Icon $(if (-not (Get-EffectiveCredential)) { 'E73E' } else { 'E77B' })
    $current.add_Click({
            $script:State.UseCurrent = $true
            Update-CredentialLabel
            Write-Log 'Operacje będą wykonywane na bieżącym koncie.' 'INFO' -Module ''
        })
    [void]$menu.Items.Add($current)
    $saved = $script:State.Credential
    if ($saved) {
        $alt = New-MenuItem -Text ('Zapamiętane: {0}' -f $saved.UserName) -Icon $(if (Get-EffectiveCredential) { 'E73E' } else { 'E77B' })
        $alt.add_Click({
                $script:State.UseCurrent = $false
                Update-CredentialLabel
                Write-Log ('Operacje będą wykonywane jako {0}.' -f $script:State.Credential.UserName) 'INFO' -Module ''
            })
        [void]$menu.Items.Add($alt)
    }
    [void]$menu.Items.Add((New-MenuSeparator))
    $other = New-MenuItem -Text 'Inne konto…' -Icon 'E8FA'
    $other.add_Click({
            Invoke-UiAction -Module $null -Action {
                $userName = if ($script:State.Credential) { $script:State.Credential.UserName } else { '' }
                $cred = Show-CredentialDialog -UserName $userName
                if (-not $cred) { return }
                $script:State.Credential = $cred
                $script:State.UseCurrent = $false
                Update-CredentialLabel
                Write-Log ('Operacje będą wykonywane jako {0}.' -f $cred.UserName) 'OK' -Module ''
                Show-Toast "Konto: $($cred.UserName)" 'ok'
            }
        })
    [void]$menu.Items.Add($other)
    if ($saved) {
        $forget = New-MenuItem -Text 'Zapomnij poświadczenia' -Icon 'E74D' -Danger
        $forget.add_Click({
                $script:State.Credential = $null
                $script:State.UseCurrent = $true
                Update-CredentialLabel
                Write-Log 'Usunięto poświadczenia z pamięci.' 'INFO' -Module ''
            })
        [void]$menu.Items.Add($forget)
    }
    $menu.PlacementTarget = $Anchor
    $menu.Placement = 'Bottom'
    $menu.IsOpen = $true
}

function Initialize-Navigation {
    # Przyciski przestrzeni roboczych i nawigacja modułów (kategorie w kolejności rejestracji)
    $c = $script:UI.Controls
    $c.wsSwitcher.Children.Clear()
    $c.navHost.Children.Clear()
    $i = 0
    foreach ($ws in $script:UI.Workspaces.Values) {
        $i++
        $tab = New-Object System.Windows.Controls.RadioButton
        $tab.Style = Get-ThemeResource 'WorkspaceTab'
        $tab.GroupName = 'workspaces'
        $tab.Content = New-IconContent -Text $ws.Title -Icon $ws.Icon
        $tab.Tag = $ws.Key
        $tab.ToolTip = "$($ws.Description) (Ctrl+$i)"
        $tab.add_Click($script:NavEvents.WorkspaceClick)
        [void]$c.wsSwitcher.Children.Add($tab)
        $ws.Tab = $tab

        $panel = New-Object System.Windows.Controls.StackPanel
        $panel.Visibility = 'Collapsed'
        $ws.NavHost = $panel
        $ws.NavItems = @{}
        [void]$c.navHost.Children.Add($panel)
        $defs = @($script:UI.ModuleDefs | Where-Object { $_.Workspace -eq $ws.Key })
        $categories = New-Object System.Collections.ArrayList
        foreach ($d in $defs) { if (-not $categories.Contains($d.Category)) { [void]$categories.Add($d.Category) } }
        foreach ($cat in $categories) {
            $header = New-Object System.Windows.Controls.TextBlock
            $header.Style = Get-ThemeResource 'NavHeader'
            $header.Text = $cat.ToUpperInvariant()
            [void]$panel.Children.Add($header)
            foreach ($d in @($defs | Where-Object { $_.Category -eq $cat })) {
                $item = New-NavItem -Definition $d -Workspace $ws.Key
                $ws.NavItems[$d.Key] = $item
                [void]$panel.Children.Add($item)
            }
        }
    }
    foreach ($d in $script:UI.ModuleDefs) {
        if (-not $script:UI.Workspaces.Contains($d.Workspace)) { Write-Log "Moduł «$($d.Title)» wskazuje nieistniejącą przestrzeń roboczą «$($d.Workspace)»." 'WARN' -Module '' }
    }
}

function New-NavItem {
    param([hashtable]$Definition, [string]$Workspace)
    $rb = New-Object System.Windows.Controls.RadioButton
    $rb.Style = Get-ThemeResource 'NavItem'
    $rb.GroupName = "nav_$Workspace"
    $rb.Tag = $Definition.Key
    if ($Definition.Description) { $rb.ToolTip = $Definition.Description }
    $grid = New-Object System.Windows.Controls.Grid
    foreach ($w in @('Auto', '*', 'Auto', 'Auto')) {
        $cd = New-Object System.Windows.Controls.ColumnDefinition
        $cd.Width = if ($w -eq '*') { New-Object System.Windows.GridLength(1, [System.Windows.GridUnitType]::Star) } else { [System.Windows.GridLength]::Auto }
        [void]$grid.ColumnDefinitions.Add($cd)
    }
    $icon = New-GlyphBlock -Code $Definition.Icon -Size 14
    $icon.Margin = '2,0,11,0'
    $icon.Width = 18
    [void]$grid.Children.Add($icon)
    $text = New-Object System.Windows.Controls.TextBlock
    $text.Text = $Definition.Title
    $text.TextTrimming = 'CharacterEllipsis'
    $text.VerticalAlignment = 'Center'
    [System.Windows.Controls.Grid]::SetColumn($text, 1)
    [void]$grid.Children.Add($text)
    if ($Definition.Badge) {
        # Oznaczenie modułu (np. "nowe") - dyskretna kropka z podpowiedzią, nie zabiera miejsca na nazwę
        $badge = New-Object System.Windows.Shapes.Ellipse
        $badge.Width = 6
        $badge.Height = 6
        $badge.Fill = Get-Brush '#5EE3AE'
        $badge.Margin = '6,0,0,0'
        $badge.VerticalAlignment = 'Center'
        $badge.ToolTip = $Definition.Badge
        [System.Windows.Controls.Grid]::SetColumn($badge, 2)
        [void]$grid.Children.Add($badge)
    }
    $dot = New-Object System.Windows.Shapes.Ellipse
    $dot.Width = 7
    $dot.Height = 7
    $dot.Fill = Get-Brush '#4C7DF0'
    $dot.Margin = '8,0,2,0'
    $dot.VerticalAlignment = 'Center'
    $dot.Visibility = 'Collapsed'
    $dot.ToolTip = 'Trwa operacja'
    [System.Windows.Controls.Grid]::SetColumn($dot, 3)
    [void]$grid.Children.Add($dot)
    $script:NavDots[$rb] = $dot
    $rb.Content = $grid
    $rb.add_Click($script:NavEvents.ModuleClick)
    return $rb
}

$script:NavEvents = @{
    WorkspaceClick = {
        param($s, $e)
        try { Show-Workspace -Key ([string]$s.Tag) }
        catch { Write-Log "Błąd przełączania przestrzeni: $($_.Exception.Message)" 'ERROR' -Module '' }
    }
    ModuleClick    = {
        param($s, $e)
        try { Show-Module -Key ([string]$s.Tag) }
        catch {
            Write-Log "Błąd otwierania modułu: $($_.Exception.Message)" 'ERROR' -Module ''
            Show-Error 'Nie można otworzyć modułu.' $_
        }
    }
}
#endregion

#region Przestrzenie robocze
Register-Workspace -Key 'Remote' -Title 'Zarządzanie zdalne' -Icon 'E7F4' -Target Computer `
    -Description 'Operacje na zaznaczonych komputerach przez PowerShell Remoting (WinRM)'
Register-Workspace -Key 'AdUsers' -Title 'Użytkownicy AD' -Icon 'E716' -Target User `
    -Description 'Konta użytkowników w Active Directory: hasła, blokady, grupy, atrybuty i raporty'
Register-Workspace -Key 'AdComputers' -Title 'Komputery AD' -Icon 'E977' -Target Computer `
    -Description 'Konta komputerów w Active Directory: LAPS, BitLocker, nazwy, grupy i raporty'
#endregion

#region Zarządzanie zdalne: Diagnostyka
Register-Module -Workspace 'Remote' -Category 'Diagnostyka' -Key 'Connectivity' -Title 'Łączność' -Icon 'E968' `
    -Description 'DNS, ping, porty TCP i sesja PowerShell Remoting. Test działa lokalnie – także dla komputerów bez WinRM – i koloruje kropki na liście komputerów.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Stan')
    $row = Add-ToolbarRow -Module $m -Title 'Parametry'
    Add-Label -Parent $row -Text 'Porty TCP' | Out-Null
    $m.Ports = Add-TextBox -Parent $row -Width 220 -Text '5985, 5986, 445, 3389, 135'
    Add-Label -Parent $row -Text 'Limit (ms)' | Out-Null
    $m.TcpTimeout = Add-Numeric -Parent $row -Value 1500 -Minimum 200 -Maximum 10000 -Width 70
    $m.TestSession = Add-CheckBox -Parent $row -Text 'Test sesji PowerShell' -Checked $true -ToolTip 'Invoke-Command na komputerze docelowym (wymaga WinRM i uprawnień)'
    $row2 = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row2 -Text 'Testuj zaznaczone' -Icon 'E72C' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $ports = @(Split-ListText ($m.Ports.Text -replace '\s+', ',') | Where-Object { $_ -match '^\d+$' -and [int]$_ -ge 1 -and [int]$_ -le 65535 } | ForEach-Object { [int]$_ } | Select-Object -Unique)
        $params = @{ Ports = $ports; TimeoutMs = (Get-Num $m.TcpTimeout); TestSession = (Test-Checked $m.TestSession) }
        $m.Data.Counts = @{ ok = 0; warn = 0; crit = 0 }
        Reset-StatTiles $m
        Start-HostOperation -Module $m -Name 'Test łączności' -Targets $targets -Local -Parameters $params -ScriptBlock {
            param($Target, $P, $Ctx)
            $row = [ordered]@{ 'Stan' = '' }
            $ip = ''
            try {
                $addresses = [System.Net.Dns]::GetHostAddresses($Target) | Where-Object { $_.AddressFamily -eq 'InterNetwork' }
                $ip = (@($addresses | ForEach-Object { $_.IPAddressToString }) -join ', ')
            }
            catch { $ip = '(brak w DNS)' }
            $row['Adres IP'] = $ip
            $pingOk = $false
            $rtt = $null
            try {
                $reply = (New-Object System.Net.NetworkInformation.Ping).Send($Target, 1000)
                if ($reply.Status -eq 'Success') { $pingOk = $true; $rtt = [int]$reply.RoundtripTime }
            }
            catch { }
            $row['Ping'] = $pingOk
            $row['Czas (ms)'] = $rtt
            $anyPort = $false
            foreach ($port in $P.Ports) {
                $open = $false
                $client = New-Object System.Net.Sockets.TcpClient
                try {
                    $async = $client.BeginConnect($Target, [int]$port, $null, $null)
                    if ($async.AsyncWaitHandle.WaitOne([int]$P.TimeoutMs) -and $client.Connected) { $open = $true }
                }
                catch { }
                finally { $client.Close() }
                if ($open) { $anyPort = $true }
                $row["TCP $port"] = $open
            }
            $psOk = $null
            if ($P.TestSession) {
                try {
                    $ic = @{ ComputerName = $Target; ErrorAction = 'Stop'; ScriptBlock = { $PSVersionTable.PSVersion.ToString() } }
                    if ($Ctx.Credential) { $ic.Credential = $Ctx.Credential }
                    if ($Ctx.SessionOption) { $ic.SessionOption = $Ctx.SessionOption }
                    $row['PowerShell zdalnie'] = 'Tak (PS ' + (Invoke-Command @ic) + ')'
                    $psOk = $true
                }
                catch {
                    $row['PowerShell zdalnie'] = 'Nie'
                    $row['Szczegóły'] = ($_.Exception.Message -split "`n")[0].Trim()
                    $psOk = $false
                }
            }
            if ($psOk -eq $true -or ($null -eq $psOk -and ($pingOk -or $anyPort))) { $row['Stan'] = 'Dostępny'; $row['__tone'] = 'ok' }
            elseif ($pingOk -or $anyPort) { $row['Stan'] = 'Częściowo'; $row['__tone'] = 'warn' }
            else { $row['Stan'] = 'Niedostępny'; $row['__tone'] = 'crit' }
            [pscustomobject]$row
        } -OnResult {
            param($m, $r)
            $d = @($r.Data | Where-Object { $_ })
            $tone = if ($d.Count -gt 0) { [string](Get-ObjectValue $d[0] '__tone') } else { 'crit' }
            if (-not $tone) { $tone = 'crit' }
            $status = switch ($tone) { 'ok' { 'online' } 'warn' { 'częściowo dostępny' } default { 'offline' } }
            $psText = if ($d.Count -gt 0) { [string](Get-ObjectValue $d[0] 'PowerShell zdalnie') } else { '' }
            if ($psText -like 'Tak*') { $status += ' • ' + ($psText -replace '^Tak \((.*)\)$', '$1') }
            Set-ComputerState -Name $r.Target -State $tone -Status $status
            $m.Data.Counts[$tone]++
        } -OnComplete {
            param($m)
            Set-StatTile -Module $m -Key 'ok' -Value ([string]$m.Data.Counts.ok) -Tone 'ok'
            Set-StatTile -Module $m -Key 'warn' -Value ([string]$m.Data.Counts.warn) -Tone $(if ($m.Data.Counts.warn) { 'warn' } else { '' })
            Set-StatTile -Module $m -Key 'crit' -Value ([string]$m.Data.Counts.crit) -Tone $(if ($m.Data.Counts.crit) { 'crit' } else { '' })
        }
    } | Out-Null
    Add-Button -Parent $row2 -Text 'Zaznacz tylko dostępne' -Icon 'E73E' -Module $m -AlwaysEnabled -ToolTip 'Na liście komputerów zostaną zaznaczone tylko komputery dostępne w ostatnim teście' -OnClick {
        param($m)
        $online = @($m.Table.Rows | Where-Object { [string]$_['__tone'] -eq 'ok' } | ForEach-Object { [string]$_['Komputer'] })
        if ($online.Count -eq 0) { Show-Warning 'Brak komputerów dostępnych w ostatnim teście.'; return }
        foreach ($r in $script:UI.HostTable.Rows) { $r['Sel'] = ($online -contains [string]$r['Name']) }
        Update-TargetCount $script:UI.ComputerPanel
        Show-Toast "Zaznaczono dostępne komputery: $($online.Count)" 'ok'
    } | Out-Null
    Add-StatTile -Module $m -Key 'ok' -Label 'Dostępne' -Icon 'E73E' | Out-Null
    Add-StatTile -Module $m -Key 'warn' -Label 'Częściowo (bez WinRM)' -Icon 'E7BA' | Out-Null
    Add-StatTile -Module $m -Key 'crit' -Label 'Niedostępne' -Icon 'E711' | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'Diagnostyka' -Key 'Inventory' -Title 'Inwentaryzacja' -Icon 'E7F8' `
    -Description 'Sprzęt, system, numer seryjny, pamięć, dysk systemowy, adresy IP i zalogowany użytkownik zaznaczonych komputerów.' -Build {
    param($m)
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row -Text 'Pobierz informacje' -Icon 'E896' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Inwentaryzacja' -Targets $targets -ScriptBlock {
            param($P)
            $cs = Get-CimInstance -ClassName Win32_ComputerSystem
            $os = Get-CimInstance -ClassName Win32_OperatingSystem
            $bios = Get-CimInstance -ClassName Win32_BIOS
            $cpu = @(Get-CimInstance -ClassName Win32_Processor)[0]
            $nics = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True')
            $ips = @($nics | ForEach-Object { $_.IPAddress } | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' }) -join ', '
            $macs = @($nics | ForEach-Object { $_.MACAddress } | Where-Object { $_ }) -join ', '
            $sysDisk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $env:SystemDrive)
            $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
            $release = ''
            $build = $os.BuildNumber
            if ($cv) {
                if ($cv.DisplayVersion) { $release = $cv.DisplayVersion } elseif ($cv.ReleaseId) { $release = $cv.ReleaseId }
                if ($null -ne $cv.UBR) { $build = '{0}.{1}' -f $os.BuildNumber, $cv.UBR }
            }
            $tpm = $null
            try { $tpm = Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction Stop } catch { }
            $secureBoot = $null
            try { $secureBoot = Confirm-SecureBootUEFI -ErrorAction Stop } catch { }
            [pscustomobject]@{
                'Producent'                = $cs.Manufacturer
                'Model'                    = $cs.Model
                'Numer seryjny'            = $bios.SerialNumber
                'BIOS'                     = $bios.SMBIOSBIOSVersion
                'System'                   = $os.Caption
                'Wydanie'                  = $release
                'Kompilacja'               = $build
                'Architektura'             = $os.OSArchitecture
                'Procesor'                 = ([string]$cpu.Name).Trim()
                'Rdzenie'                  = $cpu.NumberOfCores
                'RAM (GB)'                 = [Math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
                'Dysk systemowy (GB)'      = $(if ($sysDisk) { [Math]::Round($sysDisk.Size / 1GB, 1) } else { $null })
                'Wolne na systemowym (GB)' = $(if ($sysDisk) { [Math]::Round($sysDisk.FreeSpace / 1GB, 1) } else { $null })
                'TPM'                      = $(if ($tpm) { 'wersja ' + (([string]$tpm.SpecVersion -split ',')[0]).Trim() } else { 'brak / niedostępny' })
                'Secure Boot'              = $secureBoot
                'Zalogowany użytkownik'    = $cs.UserName
                'Domena'                   = $cs.Domain
                'Adresy IP'                = $ips
                'MAC'                      = $macs
                'Instalacja systemu'       = $os.InstallDate
                'Ostatni start'            = $os.LastBootUpTime
            }
        }
    } | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'Diagnostyka' -Key 'Performance' -Title 'Wydajność' -Icon 'E9D2' -Badge 'Nowość w wersji 4.0' `
    -Description 'Bieżące obciążenie: procesor, pamięć, dysk systemowy, kolejka dysku i procesy zużywające najwięcej zasobów. Kolor oceny wskazuje komputery wymagające uwagi.' -Build {
    param($m)
    $m.PillColumns = @('Ocena')
    $row = Add-ToolbarRow -Module $m -Title 'Parametry'
    Add-Label -Parent $row -Text 'Próbkowanie (s)' | Out-Null
    $m.Sample = Add-Numeric -Parent $row -Value 2 -Minimum 1 -Maximum 10 -Width 60
    Add-Label -Parent $row -Text 'Procesy w zestawieniu' | Out-Null
    $m.Top = Add-Numeric -Parent $row -Value 5 -Minimum 1 -Maximum 20 -Width 60
    $row2 = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row2 -Text 'Zmierz obciążenie' -Icon 'E9D2' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Wydajność' -Targets $targets -Parameters @{ Sample = (Get-Num $m.Sample); Top = (Get-Num $m.Top) } -ScriptBlock {
            param($P)
            $cores = [Math]::Max(1, [int](Get-CimInstance -ClassName Win32_ComputerSystem).NumberOfLogicalProcessors)
            # Pierwszy odczyt liczników sformatowanych bywa zerowy - dwa odczyty w odstępie
            $null = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfProc_Process -ErrorAction SilentlyContinue
            Start-Sleep -Seconds ([int]$P.Sample)
            $cpu = (Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'").PercentProcessorTime
            $procs = @(Get-CimInstance -ClassName Win32_PerfFormattedData_PerfProc_Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne '_Total' -and $_.Name -ne 'Idle' })
            $os = Get-CimInstance -ClassName Win32_OperatingSystem
            $memPct = [Math]::Round((1 - ($os.FreePhysicalMemory / [double]$os.TotalVisibleMemorySize)) * 100)
            $commitPct = [Math]::Round((1 - ($os.FreeVirtualMemory / [double]$os.TotalVirtualMemorySize)) * 100)
            $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $env:SystemDrive)
            $freePct = if ($disk -and $disk.Size) { [Math]::Round($disk.FreeSpace / [double]$disk.Size * 100) } else { $null }
            $queue = $null
            try { $queue = (Get-CimInstance -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'" -ErrorAction Stop).CurrentDiskQueueLength } catch { }
            $topCpu = @($procs | Sort-Object PercentProcessorTime -Descending | Select-Object -First ([int]$P.Top) | Where-Object { $_.PercentProcessorTime -gt 0 } |
                ForEach-Object { '{0} {1}%' -f ($_.Name -replace '#\d+$', ''), [Math]::Round($_.PercentProcessorTime / $cores) }) -join ', '
            $topMem = @($procs | Sort-Object WorkingSetPrivate -Descending | Select-Object -First ([int]$P.Top) |
                ForEach-Object { '{0} {1} MB' -f ($_.Name -replace '#\d+$', ''), [Math]::Round($_.WorkingSetPrivate / 1MB) }) -join ', '
            $issues = @()
            $tone = 'ok'
            if ($cpu -ge 90) { $issues += 'procesor'; $tone = 'crit' } elseif ($cpu -ge 75) { $issues += 'procesor'; if ($tone -eq 'ok') { $tone = 'warn' } }
            if ($memPct -ge 92) { $issues += 'pamięć'; $tone = 'crit' } elseif ($memPct -ge 80) { $issues += 'pamięć'; if ($tone -eq 'ok') { $tone = 'warn' } }
            if ($null -ne $freePct) {
                if ($freePct -le 5) { $issues += 'mało miejsca'; $tone = 'crit' } elseif ($freePct -le 15) { $issues += 'mało miejsca'; if ($tone -eq 'ok') { $tone = 'warn' } }
            }
            $boot = $os.LastBootUpTime
            [pscustomobject]@{
                'Ocena'                  = $(if ($issues.Count -eq 0) { 'OK' } else { $issues -join ', ' })
                'CPU %'                  = [int]$cpu
                'Pamięć %'               = $memPct
                'Pamięć zadeklarowana %' = $commitPct
                'Wolne na systemowym %'  = $freePct
                'Kolejka dysku'          = $queue
                'Procesy'                = $procs.Count
                'Najwięcej CPU'          = $topCpu
                'Najwięcej pamięci'      = $topMem
                'Czas pracy (dni)'       = [Math]::Round(((Get-Date) - $boot).TotalDays, 1)
                '__tone'                 = $tone
            }
        }
    } | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'Diagnostyka' -Key 'Power' -Title 'Zasilanie i uptime' -Icon 'E7E8' `
    -Description 'Czas pracy, oczekujący restart (CBS, Windows Update, zmiana nazwy, SCCM) oraz zaplanowany restart lub wyłączenie z komunikatem dla użytkownika.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.GoodWhenNo = @('Oczekuje restartu')
    $row = Add-ToolbarRow -Module $m -Title 'Sprawdzenie'
    Add-Button -Parent $row -Text 'Pokaż uptime' -Icon 'E916' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Uptime' -Targets $targets -ScriptBlock {
            param($P)
            $os = Get-CimInstance -ClassName Win32_OperatingSystem
            $boot = $os.LastBootUpTime
            $up = (Get-Date) - $boot
            $reasons = @()
            if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $reasons += 'Obsługa składników (CBS)' }
            if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $reasons += 'Windows Update' }
            $pending = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
            if ($pending) { $reasons += 'Operacje na plikach' }
            $active = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName' -ErrorAction SilentlyContinue).ComputerName
            $next = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName' -ErrorAction SilentlyContinue).ComputerName
            if ($active -and $next -and $active -ne $next) { $reasons += "Zmiana nazwy na $next" }
            try {
                $ccm = Invoke-CimMethod -Namespace 'root\ccm\ClientSDK' -ClassName CCM_ClientUtilities -MethodName DetermineIfRebootPending -ErrorAction Stop
                if ($ccm.RebootPending -or $ccm.IsHardRebootPending) { $reasons += 'SCCM' }
            }
            catch { }
            [pscustomobject]@{
                'Ostatni start'     = $boot
                'Czas pracy'        = '{0} d {1:00} h {2:00} min' -f $up.Days, $up.Hours, $up.Minutes
                'Dni pracy'         = [Math]::Round($up.TotalDays, 1)
                'Oczekuje restartu' = ($reasons.Count -gt 0)
                'Powód'             = ($reasons -join ', ')
                'Zalogowany'        = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
                '__flag'            = $(if ($reasons.Count -gt 0) { 'warn' } else { '' })
            }
        }
    } | Out-Null

    $row2 = Add-ToolbarRow -Module $m -Title 'Restart i wyłączenie'
    Add-Label -Parent $row2 -Text 'Opóźnienie (s)' | Out-Null
    $m.Delay = Add-Numeric -Parent $row2 -Value 60 -Minimum 0 -Maximum 86400 -Width 80
    Add-Label -Parent $row2 -Text 'Komunikat' | Out-Null
    $m.Message = Add-TextBox -Parent $row2 -Width 380 -Text 'Komputer zostanie uruchomiony ponownie przez administratora. Zapisz swoją pracę.'
    $m.Force = Add-CheckBox -Parent $row2 -Text 'Wymuś zamknięcie aplikacji' -Checked $true
    $row3 = Add-ToolbarRow -Module $m -Title ' '
    $powerAction = {
        param($m, $s)
        $mode = [string]$s.Tag
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $verb = @{ '/r' = 'Uruchomić ponownie'; '/s' = 'Wyłączyć'; '/a' = 'Anulować zaplanowany restart lub wyłączenie na' }[$mode]
        if (-not (Confirm-Action -Text "$verb $($targets.Count) komputer(ów)?" -Items $targets -ConfirmText $(if ($mode -eq '/a') { 'Anuluj zaplanowane' } else { 'Wykonaj' }) -Danger:($mode -ne '/a'))) { return }
        $params = @{ Mode = $mode; Delay = (Get-Num $m.Delay); Message = $m.Message.Text.Trim(); Force = (Test-Checked $m.Force) }
        Start-HostOperation -Module $m -Name "Zasilanie ($mode)" -Targets $targets -Output Log -Parameters $params -ScriptBlock {
            param($P)
            $shutdown = Join-Path $env:SystemRoot 'System32\shutdown.exe'
            if ($P.Mode -eq '/a') {
                $out = & $shutdown /a 2>&1
                if ($LASTEXITCODE -eq 0) { return 'Anulowano zaplanowane zamknięcie systemu.' }
                if ($LASTEXITCODE -eq 1116) { return 'Brak zaplanowanego zamknięcia systemu.' }
                throw ("shutdown.exe /a: kod {0} {1}" -f $LASTEXITCODE, ($out | Out-String).Trim())
            }
            $cmdArgs = @($P.Mode, '/t', [string]$P.Delay, '/d', 'p:0:0')
            if ($P.Force) { $cmdArgs += '/f' }
            if ($P.Message) { $cmdArgs += '/c'; $cmdArgs += $P.Message.Substring(0, [Math]::Min(500, $P.Message.Length)) }
            $out = & $shutdown @cmdArgs 2>&1
            if ($LASTEXITCODE -ne 0) { throw ("shutdown.exe: kod {0} {1}" -f $LASTEXITCODE, ($out | Out-String).Trim()) }
            $what = if ($P.Mode -eq '/r') { 'Restart' } else { 'Wyłączenie' }
            "$what zaplanowane za $($P.Delay) s."
        }
    }
    $b = Add-Button -Parent $row3 -Text 'Restart' -Icon 'E777' -Module $m -Danger -OnClick $powerAction
    $b.Tag = '/r'
    $b = Add-Button -Parent $row3 -Text 'Wyłącz' -Icon 'E7E8' -Module $m -Danger -OnClick $powerAction
    $b.Tag = '/s'
    $b = Add-Button -Parent $row3 -Text 'Anuluj zaplanowane' -Icon 'E711' -Module $m -OnClick $powerAction
    $b.Tag = '/a'
}
#endregion

#region Zarządzanie zdalne: Użytkownicy i dostęp
# Profile użytkowników: lista profili z komputerów + weryfikacja kont w AD (ADSI - nie wymaga RSAT).
# Kandydaci do usunięcia: konta usunięte, wyłączone lub wygasłe w AD, usunięte konta lokalne, profile tymczasowe
# i uszkodzone, opcjonalnie profile nieużywane dłużej niż N dni. Profile załadowane (zalogowany użytkownik)
# i systemowe nigdy nie są kandydatami i nie są usuwane.
$script:ProfileScanScript = {
    param($Target, $P, $Ctx)
    $ic = @{ ComputerName = $Target; ErrorAction = 'Stop'; ArgumentList = @(, @{ Size = $P.Size }) }
    if ($Ctx.Credential) { $ic.Credential = $Ctx.Credential }
    if ($Ctx.SessionOption) { $ic.SessionOption = $Ctx.SessionOption }
    $profiles = @(Invoke-Command @ic -ScriptBlock {
            param($Q)
            # Konta lokalne (SID -> włączone) i SID komputera - do rozpoznania profili kont lokalnych
            $machineSid = ''
            $localUsers = @{}
            try {
                foreach ($u in @(Get-LocalUser -ErrorAction Stop)) {
                    $localUsers[[string]$u.SID] = [bool]$u.Enabled
                    if (-not $machineSid) { $machineSid = [string]$u.SID.AccountDomainSid }
                }
            }
            catch {
                try {
                    $computer = [ADSI]"WinNT://$env:COMPUTERNAME,computer"
                    foreach ($child in $computer.Children) {
                        if ($child.SchemaClassName -ne 'User') { continue }
                        $sidObj = New-Object System.Security.Principal.SecurityIdentifier($child.objectSid.Value, 0)
                        $localUsers[$sidObj.Value] = -not ([int]$child.UserFlags.Value -band 2)
                        if (-not $machineSid) { $machineSid = $sidObj.AccountDomainSid.Value }
                    }
                }
                catch { }
            }
            $system = @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')
            foreach ($pr in @(Get-CimInstance -ClassName Win32_UserProfile | Where-Object { -not $_.Special -and $system -notcontains $_.SID })) {
                $sid = [string]$pr.SID
                # Ostatnie użycie: czasy załadowania/wyładowania profilu z ProfileList (Windows 10 1903+),
                # potem NTUSER.DAT; LastUseTime z WMI bywa zmieniany przez skanery i kopie zapasowe
                $reg = Get-ItemProperty -LiteralPath "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" -ErrorAction SilentlyContinue
                $times = @()
                foreach ($pair in @(@('LocalProfileLoadTimeHigh', 'LocalProfileLoadTimeLow'), @('LocalProfileUnloadTimeHigh', 'LocalProfileUnloadTimeLow'))) {
                    if ($reg -and $null -ne $reg.($pair[0]) -and $null -ne $reg.($pair[1])) {
                        try {
                            $ft = ([int64][uint32]$reg.($pair[0]) -shl 32) -bor [int64][uint32]$reg.($pair[1])
                            if ($ft -gt 0) { $times += [datetime]::FromFileTime($ft) }
                        }
                        catch { }
                    }
                }
                $last = $null
                $source = 'rejestr (ProfileList)'
                if ($times.Count -gt 0) { $last = @($times | Sort-Object -Descending)[0] }
                else {
                    $nt = if ($pr.LocalPath) { Get-Item -LiteralPath (Join-Path $pr.LocalPath 'NTUSER.DAT') -Force -ErrorAction SilentlyContinue } else { $null }
                    if ($nt) { $last = $nt.LastWriteTime; $source = 'NTUSER.DAT' }
                    elseif ($pr.LastUseTime) { $last = $pr.LastUseTime; $source = 'WMI LastUseTime' }
                }
                $size = $null
                if ($Q.Size -and $pr.LocalPath -and (Test-Path -LiteralPath $pr.LocalPath)) {
                    # Własne przejście katalogów: pomija dowiązania (reparse points), nie przerywa na braku dostępu
                    $total = [int64]0
                    $stack = New-Object System.Collections.Stack
                    $stack.Push((New-Object System.IO.DirectoryInfo($pr.LocalPath)))
                    while ($stack.Count -gt 0) {
                        $dir = $stack.Pop()
                        try {
                            foreach ($e in $dir.EnumerateFileSystemInfos()) {
                                if ($e.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                                if ($e -is [System.IO.DirectoryInfo]) { $stack.Push($e) } else { $total += $e.Length }
                            }
                        }
                        catch { }
                    }
                    $size = $total
                }
                $localState = ''
                if ($machineSid -and $sid.StartsWith($machineSid + '-')) {
                    $localState = if (-not $localUsers.ContainsKey($sid)) { 'missing' } elseif ($localUsers[$sid]) { 'enabled' } else { 'disabled' }
                }
                $account = ''
                try { $account = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value } catch { }
                [pscustomobject]@{
                    Sid        = $sid
                    Path       = [string]$pr.LocalPath
                    Loaded     = [bool]$pr.Loaded
                    Status     = [int]$pr.Status
                    Roaming    = [bool]$pr.RoamingConfigured
                    LastUse    = $last
                    LastSource = $source
                    Size       = $size
                    LocalState = $localState
                    Account    = $account
                }
            }
        })

    # Weryfikacja kont w AD przez ADSI (wspólna pamięć podręczna wszystkich komputerów w $P.Cache)
    $cache = $P.Cache
    $server = if ($Ctx.Server) { $Ctx.Server + '/' } else { '' }
    $user = $null
    $pass = $null
    if ($Ctx.Credential) { $user = $Ctx.Credential.UserName; $pass = $Ctx.Credential.GetNetworkCredential().Password }
    $newEntry = {
        param([string]$Path)
        if ($user) { New-Object System.DirectoryServices.DirectoryEntry($Path, $user, $pass) } else { New-Object System.DirectoryServices.DirectoryEntry($Path) }
    }
    if (-not $cache.ContainsKey('__domainSid')) {
        try {
            $rootDse = & $newEntry ('LDAP://{0}RootDSE' -f $server)
            $dnc = [string]$rootDse.Properties['defaultNamingContext'].Value
            $domain = & $newEntry ('LDAP://{0}{1}' -f $server, $dnc)
            $cache['__domainSid'] = (New-Object System.Security.Principal.SecurityIdentifier([byte[]]$domain.Properties['objectSid'].Value, 0)).Value
        }
        catch {
            $cache['__domainSid'] = ''
            $cache['__domainError'] = $_.Exception.Message
        }
    }
    $domainSid = [string]$cache['__domainSid']
    $classify = [scriptblock]::Create($P.ClassifyScript)
    $now = Get-Date
    foreach ($pr in $profiles) {
        $sid = [string]$pr.Sid
        $ad = @{ Kind = 'other'; Found = $false; Error = ''; Sam = ''; Name = ''; Uac = 0; Expires = [int64]0; Logon = [int64]0 }
        if ($sid.StartsWith('S-1-12-1-')) { $ad.Kind = 'entra' }
        elseif ($pr.LocalState) { $ad.Kind = 'local' }
        elseif ($domainSid -and $sid.StartsWith($domainSid + '-')) {
            $ad.Kind = 'domain'
            $info = $cache[$sid]
            if (-not $info) {
                $info = @{ Found = $false; Error = ''; Sam = ''; Name = ''; Uac = 0; Expires = [int64]0; Logon = [int64]0 }
                try {
                    $entry = & $newEntry ('LDAP://{0}<SID={1}>' -f $server, $sid)
                    $ds = New-Object System.DirectoryServices.DirectorySearcher($entry)
                    $ds.SearchScope = 'Base'
                    $ds.Filter = '(objectClass=*)'
                    foreach ($a in 'sAMAccountName', 'displayName', 'userAccountControl', 'accountExpires', 'lastLogonTimestamp') { [void]$ds.PropertiesToLoad.Add($a) }
                    $res = $ds.FindOne()
                    if ($res) {
                        $info.Found = $true
                        if ($res.Properties['samaccountname'].Count) { $info.Sam = [string]$res.Properties['samaccountname'][0] }
                        if ($res.Properties['displayname'].Count) { $info.Name = [string]$res.Properties['displayname'][0] }
                        if ($res.Properties['useraccountcontrol'].Count) { $info.Uac = [int]$res.Properties['useraccountcontrol'][0] }
                        if ($res.Properties['accountexpires'].Count) { $info.Expires = [int64]$res.Properties['accountexpires'][0] }
                        if ($res.Properties['lastlogontimestamp'].Count) { $info.Logon = [int64]$res.Properties['lastlogontimestamp'][0] }
                    }
                }
                catch {
                    $ex = $_.Exception
                    while ($ex.InnerException -and -not ($ex -is [System.Runtime.InteropServices.COMException])) { $ex = $ex.InnerException }
                    # 0x80072030: obiekt nie istnieje - konto usunięte
                    if (-not ($ex -is [System.Runtime.InteropServices.COMException] -and $ex.ErrorCode -eq -2147016656)) { $info.Error = $ex.Message }
                }
                $cache[$sid] = $info
            }
            foreach ($k in @($info.Keys)) { $ad[$k] = $info[$k] }
        }
        elseif (-not $domainSid -and $cache['__domainError']) { $ad.Kind = 'error'; $ad.Error = [string]$cache['__domainError'] }
        & $classify $pr $ad $P $now
    }
}

# Ocena profilu (wydzielona, żeby dało się ją testować bez AD i WinRM):
# $pr - profil z komputera, $Ad - wynik sprawdzenia konta (Kind: domain/local/entra/other/error), $P - parametry, $Now - data
$script:ProfileClassifyScript = {
    param($pr, $Ad, $P, [datetime]$Now)
    $kind = 'inny'
    $state = ''
    $stateTone = ''
    $candidate = $false
    $display = ''
    $adLogon = $null
    $account = [string]$pr.Account
    switch ($Ad.Kind) {
        'entra' { $kind = 'Entra ID'; $state = 'Konto Entra ID (nie sprawdzano)' }
        'local' {
            $kind = 'lokalny'
            switch ($pr.LocalState) {
                'missing' { $state = 'Konto lokalne usunięte'; $stateTone = 'crit'; $candidate = $true }
                'disabled' { $state = 'Konto lokalne wyłączone'; $stateTone = 'warn' }
                default { $state = 'Konto lokalne aktywne' }
            }
        }
        'domain' {
            $kind = 'domenowy'
            if ($Ad.Error) { $state = 'Nie sprawdzono: ' + $Ad.Error }
            elseif (-not $Ad.Found) { $state = 'Konto usunięte z AD'; $stateTone = 'crit'; $candidate = $true }
            else {
                if ($Ad.Sam -and -not $account) { $account = $Ad.Sam }
                $display = $Ad.Name
                if ([int64]$Ad.Logon -gt 0) { $adLogon = [datetime]::FromFileTime([int64]$Ad.Logon) }
                $expired = ([int64]$Ad.Expires -gt 0 -and [int64]$Ad.Expires -lt [int64]::MaxValue -and [datetime]::FromFileTimeUtc([int64]$Ad.Expires) -lt $Now.ToUniversalTime())
                if ([int]$Ad.Uac -band 2) { $state = 'Konto wyłączone w AD'; $stateTone = 'crit'; $candidate = $true }
                elseif ($expired) { $state = 'Konto wygasło'; $stateTone = 'crit'; $candidate = $true }
                else { $state = 'Konto aktywne w AD' }
            }
        }
        'error' { $state = 'Nie sprawdzono: ' + $Ad.Error }
        default { $state = 'Konto z innej domeny (nie sprawdzano)' }
    }
    $temp = (([int]$pr.Status -band 1) -ne 0) -or ([string]$pr.Path -match '\.TEMP(\.\d+)?$')
    $corrupt = (([int]$pr.Status -band 8) -ne 0)
    $days = if ($pr.LastUse) { [int][Math]::Floor(($Now - [datetime]$pr.LastUse).TotalDays) } else { $null }
    $inactive = ($null -ne $days -and $days -ge [int]$P.Days)
    $verdict = ''
    $tone = 'ok'
    if ($pr.Loaded) { $verdict = 'W użyciu (zalogowany)'; $tone = 'info'; $candidate = $false }
    elseif ($stateTone -eq 'crit') { $verdict = $state; $tone = 'crit' }
    elseif ($temp -or $corrupt) { $verdict = $(if ($corrupt) { 'Profil uszkodzony' } else { 'Profil tymczasowy' }); $tone = 'warn'; $candidate = $true }
    elseif ($inactive) {
        $verdict = "Nieużywany od $days dni"
        $tone = 'warn'
        if ($P.InactiveCandidates) { $candidate = $true }
    }
    elseif ($state -like 'Nie sprawdzono*') { $verdict = 'Nie sprawdzono'; $tone = '' }
    elseif ($stateTone -eq 'warn') { $verdict = $state; $tone = 'warn' }
    else { $verdict = 'Aktywny' }
    $type = $kind
    if ($pr.Roaming) { $type += ', mobilny' }
    if ($temp) { $type += ', tymczasowy' }
    [pscustomobject][ordered]@{
        'Ocena'                   = $verdict
        'Kandydat'                = $candidate
        'Konto'                   = $(if ($account) { $account } else { '(nieznane)' })
        'Nazwa'                   = $display
        'Stan konta'              = $state
        'Ostatnie użycie'         = $pr.LastUse
        'Dni'                     = $days
        'Rozmiar (GB)'            = $(if ($null -ne $pr.Size) { [Math]::Round([double]$pr.Size / 1GB, 2) } else { $null })
        'Ścieżka'                 = $pr.Path
        'Załadowany'              = [bool]$pr.Loaded
        'Ostatnie logowanie w AD' = $adLogon
        'Typ'                     = $type
        'Źródło daty'             = $pr.LastSource
        'SID'                     = [string]$pr.Sid
        '__tone'                  = $tone
        '__flag'                  = $(if ($pr.Loaded) { 'muted' } else { '' })
    }
}

function Update-ProfileStats {
    param([hashtable]$Module)
    $rows = @($Module.Table.Rows | Where-Object { [string]$_['SID'] })
    $candidates = @($rows | Where-Object { [string]$_['Kandydat'] -eq 'Tak' })
    $measured = @($candidates | Where-Object { $_['Rozmiar (GB)'] -isnot [System.DBNull] })
    $hosts = @($Module.Table.Rows | ForEach-Object { [string]$_['Komputer'] } | Select-Object -Unique)
    Set-StatTile -Module $Module -Key 'profiles' -Value ([string]$rows.Count)
    Set-StatTile -Module $Module -Key 'candidates' -Value ([string]$candidates.Count) -Tone $(if ($candidates.Count) { 'warn' } else { 'ok' })
    if ($measured.Count -gt 0) {
        $sum = 0.0
        foreach ($r in $measured) { $sum += [double]$r['Rozmiar (GB)'] }
        Set-StatTile -Module $Module -Key 'reclaim' -Value ('{0:N1} GB' -f $sum) -Tone 'info' -Label 'Do odzyskania (kandydaci)'
    }
    else { Set-StatTile -Module $Module -Key 'reclaim' -Value '–' -Label 'Do odzyskania (włącz pomiar rozmiaru)' }
    Set-StatTile -Module $Module -Key 'hosts' -Value ([string]$hosts.Count)
}

Register-Module -Workspace 'Remote' -Category 'Użytkownicy i dostęp' -Key 'Profiles' -Title 'Profile użytkowników' -Icon 'E77B' -Badge 'Nowość w wersji 4.0' `
    -Description 'Profile na komputerach porównane z Active Directory: konta usunięte, wyłączone i wygasłe, profile tymczasowe oraz nieużywane – kandydaci do usunięcia. Weryfikacja przez ADSI (bez RSAT), bezpieczne usuwanie przez Win32_UserProfile.' -Build {
    param($m)
    $m.PillColumns = @('Ocena')
    $m.GoodWhenNo = @('Kandydat', 'Załadowany')
    $m.HiddenColumns = @()
    $m.EmptyIcon = 'E77B'
    $row = Add-ToolbarRow -Module $m -Title 'Ocena'
    Add-Label -Parent $row -Text 'Nieużywany dłużej niż (dni)' | Out-Null
    $m.Days = Add-Numeric -Parent $row -Value ([int]$script:Settings.InactiveDays) -Minimum 1 -Maximum 3650 -Width 70
    $m.InactiveCand = Add-CheckBox -Parent $row -Text 'Nieużywane też są kandydatami' -Checked $true -ToolTip 'Profile aktywnych kont, nieużywane dłużej niż podana liczba dni, zostaną oznaczone jako kandydaci do usunięcia'
    $m.MeasureSize = Add-CheckBox -Parent $row -Text 'Mierz rozmiar profili (wolniej)' -ToolTip 'Sumuje rozmiar plików w folderach profili (z pominięciem dowiązań)'
    $row2 = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row2 -Text 'Sprawdź profile' -Icon 'E721' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $params = @{
            Days               = (Get-Num $m.Days)
            InactiveCandidates = (Test-Checked $m.InactiveCand)
            Size               = (Test-Checked $m.MeasureSize)
            Cache              = [hashtable]::Synchronized(@{})
            ClassifyScript     = $script:ProfileClassifyScript.ToString()
        }
        Reset-StatTiles $m
        Start-HostOperation -Module $m -Name 'Profile użytkowników' -Targets $targets -Local -Parameters $params -ScriptBlock $script:ProfileScanScript -OnComplete { param($m) Update-ProfileStats -Module $m }
    } | Out-Null
    Add-Button -Parent $row2 -Text 'Zaznacz kandydatów' -Icon 'E8B3' -Module $m -OnClick {
        param($m)
        $grid = $m.Grid
        $grid.SelectedItems.Clear()
        $count = 0
        foreach ($drv in @($m.View | ForEach-Object { $_ })) {
            if ([string]$drv['Kandydat'] -eq 'Tak') { [void]$grid.SelectedItems.Add($drv); $count++ }
        }
        if ($count -eq 0) { Show-Toast 'Brak kandydatów do usunięcia w widocznych wierszach.' 'info' }
        else { Show-Toast "Zaznaczono kandydatów: $count" 'ok' }
    } | Out-Null
    $m.OnlyCand = Add-CheckBox -Parent $row2 -Text 'Pokaż tylko kandydatów'
    Register-ControlHandler -Control $m.OnlyCand -EventName 'Checked' -Module $m -Action { param($m) $m.ExtraFilter = "[Kandydat] = 'Tak'"; Update-ResultFilter -Module $m }
    Register-ControlHandler -Control $m.OnlyCand -EventName 'Unchecked' -Module $m -Action { param($m) $m.ExtraFilter = ''; Update-ResultFilter -Module $m }
    Add-Button -Parent $row2 -Text 'Usuń zaznaczone profile' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Remove $m $null } | Out-Null

    $m.Actions.Remove = {
        param($m, $Rows)
        $source = @(if ($null -ne $Rows) { $Rows } else { Get-SelectedResultRows -Module $m })
        $selected = @($source | Where-Object { [string](Get-ObjectValue $_ 'SID') -match '^S-1-[\d-]+$' })
        if ($selected.Count -eq 0) { Show-Warning 'Zaznacz w tabeli profile do usunięcia (np. przyciskiem «Zaznacz kandydatów»).'; return }
        $loaded = @($selected | Where-Object { [string](Get-ObjectValue $_ 'Załadowany') -eq 'Tak' })
        $selected = @($selected | Where-Object { [string](Get-ObjectValue $_ 'Załadowany') -ne 'Tak' })
        if ($selected.Count -eq 0) { Show-Warning 'Zaznaczone profile są w użyciu (użytkownicy są zalogowani) – nie można ich usunąć.'; return }
        $nonCandidates = @($selected | Where-Object { [string](Get-ObjectValue $_ 'Kandydat') -ne 'Tak' })
        $items = foreach ($r in $selected) {
            $size = Get-ObjectValue $r 'Rozmiar (GB)'
            '{0}: {1}{2} – {3}' -f (Get-ObjectValue $r 'Komputer'), (Get-ObjectValue $r 'Ścieżka'), $(if ($null -ne $size) { " ($size GB)" } else { '' }), (Get-ObjectValue $r 'Ocena')
        }
        $text = "Usunąć $($selected.Count) profil(i) użytkowników? Folder profilu i jego wpis w rejestrze zostaną trwale usunięte – tej operacji nie można cofnąć."
        if ($loaded.Count -gt 0) { $text += "`r`n`r`nPominięto profile w użyciu: $($loaded.Count)." }
        if ($nonCandidates.Count -gt 0) { $text += "`r`n`r`nUWAGA: $($nonCandidates.Count) z zaznaczonych profili NIE jest kandydatem do usunięcia (konto jest aktywne i profil był niedawno używany)." }
        if (-not (Confirm-Action -Text $text -Items $items -Title 'Usuwanie profili' -ConfirmText 'Usuń profile' -Danger)) { return }
        $per = @{}
        foreach ($r in $selected) {
            $h = [string](Get-ObjectValue $r 'Komputer')
            if (-not $per.ContainsKey($h)) { $per[$h] = @{ Sids = New-Object System.Collections.ArrayList } }
            [void]$per[$h].Sids.Add([string](Get-ObjectValue $r 'SID'))
        }
        foreach ($h in @($per.Keys)) { $per[$h].Sids = @($per[$h].Sids) }
        Start-HostOperation -Module $m -Name 'Usuwanie profili' -Targets @($per.Keys) -PerTarget $per -Output Log -ScriptBlock {
            param($P)
            foreach ($sid in $P.Sids) {
                if ($sid -notmatch '^S-1-[\d-]+$') { continue }
                try {
                    $pr = Get-CimInstance -ClassName Win32_UserProfile -Filter ("SID='{0}'" -f $sid)
                    if (-not $pr) { [pscustomobject]@{ 'SID' = $sid; 'Wynik' = 'Nie znaleziono profilu' }; continue }
                    if ($pr.Special) { [pscustomobject]@{ 'SID' = $sid; 'Wynik' = 'Pominięto – profil systemowy' }; continue }
                    if ($pr.Loaded) { [pscustomobject]@{ 'SID' = $sid; 'Ścieżka' = $pr.LocalPath; 'Wynik' = 'Pominięto – profil w użyciu' }; continue }
                    $path = $pr.LocalPath
                    Remove-CimInstance -InputObject $pr -ErrorAction Stop
                    $left = $path -and (Test-Path -LiteralPath $path)
                    [pscustomobject]@{ 'SID' = $sid; 'Ścieżka' = $path; 'Wynik' = $(if ($left) { 'Usunięto (część plików pozostała w folderze)' } else { 'Usunięto' }) }
                }
                catch { [pscustomobject]@{ 'SID' = $sid; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        } -OnResult {
            param($m, $r)
            $removed = @($r.Data | Where-Object { [string](Get-ObjectValue $_ 'Wynik') -like 'Usunięto*' } | ForEach-Object { [string](Get-ObjectValue $_ 'SID') })
            if ($removed.Count -eq 0) { return }
            $rows = @($m.Table.Rows | Where-Object { [string]$_['Komputer'] -eq $r.Target -and $removed -contains [string]$_['SID'] })
            Remove-ResultRows -Module $m -Rows $rows
        } -OnComplete { param($m) Update-ProfileStats -Module $m }
    }
    Add-RowAction -Module $m -Text 'Usuń profil' -Icon 'E74D' -Danger -Action { param($m, $rows) & $m.Actions.Remove $m $rows }
    Add-RowAction -Module $m -Text 'Otwórz folder profilu (C$)' -Icon 'E838' -Action {
        param($m, $rows)
        $r = $rows[0]
        $computer = [string](Get-ObjectValue $r 'Komputer')
        $path = [string](Get-ObjectValue $r 'Ścieżka')
        if ($path -notmatch '^[A-Za-z]:\\') { return }
        Start-Tool -FilePath 'explorer.exe' -Arguments @(('\\{0}\{1}${2}' -f $computer, $path.Substring(0, 1), $path.Substring(2))) -Name $computer
    }
    Add-StatTile -Module $m -Key 'profiles' -Label 'Profile' -Icon 'E77B' | Out-Null
    Add-StatTile -Module $m -Key 'candidates' -Label 'Kandydaci do usunięcia' -Icon 'E7BA' | Out-Null
    Add-StatTile -Module $m -Key 'reclaim' -Label 'Do odzyskania' -Icon 'EDA2' | Out-Null
    Add-StatTile -Module $m -Key 'hosts' -Label 'Komputery' -Icon 'E7F4' | Out-Null
    $m.ResultHint = 'Kandydaci: konta usunięte/wyłączone/wygasłe, profile tymczasowe i nieużywane'
}

Register-Module -Workspace 'Remote' -Category 'Użytkownicy i dostęp' -Key 'Sessions' -Title 'Sesje użytkowników' -Icon 'E7EE' -Badge 'Nowość w wersji 4.0' `
    -Description 'Zalogowani użytkownicy (konsola i pulpit zdalny): stan sesji, bezczynność, czas logowania. Wylogowanie, rozłączenie, wiadomość dla użytkownika i podgląd sesji (shadow).' -Build {
    param($m)
    $m.PillColumns = @('Stan')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Sesje' -Targets $targets -ScriptBlock {
            param($P)
            $oem = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = Join-Path $env:SystemRoot 'System32\quser.exe'
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.CreateNoWindow = $true
            $psi.StandardOutputEncoding = $oem
            $psi.StandardErrorEncoding = $oem
            $proc = [System.Diagnostics.Process]::Start($psi)
            $errTask = $proc.StandardError.ReadToEndAsync()
            $out = $proc.StandardOutput.ReadToEnd()
            $proc.WaitForExit()
            $lines = @($out -split "`r?`n" | Where-Object { $_.Trim() })
            if ($lines.Count -le 1) { return [pscustomobject]@{ 'Użytkownik' = '(brak zalogowanych użytkowników)'; '__flag' = 'muted' } }
            foreach ($line in ($lines | Select-Object -Skip 1)) {
                $mt = [regex]::Match($line, '^(?<cur>>)?\s*(?<user>\S+)\s+(?:(?<session>\S+)\s+)?(?<id>\d+)\s+(?<state>\S+)\s+(?<idle>\S+)\s+(?<logon>.+?)\s*$')
                if (-not $mt.Success) { continue }
                $state = $mt.Groups['state'].Value
                $active = ($state -match '^(Active|Aktywn)')
                [pscustomobject]@{
                    'Użytkownik'   = $mt.Groups['user'].Value
                    'Stan'         = $(if ($active) { 'Aktywna' } else { 'Rozłączona' })
                    'Sesja'        = $mt.Groups['session'].Value
                    'ID'           = [int]$mt.Groups['id'].Value
                    'Bezczynność'  = $mt.Groups['idle'].Value
                    'Zalogowano'   = $mt.Groups['logon'].Value
                    '__tone'       = $(if ($active) { 'ok' } else { 'warn' })
                }
            }
        }
    }
    $m.Actions.Session = {
        param($m, [string]$Op, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('ID', 'Użytkownik') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli sesje użytkowników.'; return }
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) '{0} (sesja {1})' -f $i['Użytkownik'], $i['ID'] }
        $message = ''
        switch ($Op) {
            'Logoff' { if (-not (Confirm-Action -Text 'Wylogować wybrane sesje? Niezapisane dane użytkowników zostaną utracone.' -Items $items -ConfirmText 'Wyloguj' -Danger)) { return } }
            'Disconnect' { if (-not (Confirm-Action -Text 'Rozłączyć wybrane sesje? Programy użytkowników pozostaną uruchomione.' -Items $items -ConfirmText 'Rozłącz')) { return } }
            'Message' {
                $message = Show-InputDialog -Title 'Wiadomość dla użytkowników' -Prompt 'Treść komunikatu wyświetlanego w wybranych sesjach (maks. 255 znaków).' -Default 'Za 10 minut nastąpi restart komputera. Zapisz swoją pracę.' -Icon 'E715' -Validate { param($t) if (-not $t.Trim()) { 'Wpisz treść wiadomości.' } elseif ($t.Length -gt 255) { 'Maksymalnie 255 znaków.' } else { '' } }
                if (-not $message) { return }
            }
        }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Op = $Op; Message = $message; Ids = @($byHost[$h] | ForEach-Object { [int]$_['ID'] }) } }
        $onComplete = if ($Op -eq 'Message') { $null } else { { param($m) & $m.Actions.List $m } }
        Start-HostOperation -Module $m -Name "Sesje – $Op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete $onComplete -ScriptBlock {
            param($P)
            foreach ($id in $P.Ids) {
                $exe = switch ($P.Op) { 'Logoff' { 'logoff.exe' } 'Disconnect' { 'tsdiscon.exe' } 'Message' { 'msg.exe' } }
                $cmdArgs = @([string]$id)
                if ($P.Op -eq 'Message') { $cmdArgs += '/TIME:300'; $cmdArgs += $P.Message }
                $out = & (Join-Path $env:SystemRoot "System32\$exe") @cmdArgs 2>&1
                if ($LASTEXITCODE -eq 0) { [pscustomobject]@{ 'Sesja' = $id; 'Operacja' = $P.Op; 'Wynik' = 'OK' } }
                else { [pscustomobject]@{ 'Sesja' = $id; 'Operacja' = $P.Op; 'Wynik' = ('Błąd – kod {0} {1}' -f $LASTEXITCODE, ($out | Out-String).Trim()) } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row -Text 'Pokaż sesje' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Wiadomość…' -Icon 'E715' -Module $m -OnClick { param($m) & $m.Actions.Session $m 'Message' $null } | Out-Null
    Add-Button -Parent $row -Text 'Rozłącz' -Icon 'E8CD' -Module $m -OnClick { param($m) & $m.Actions.Session $m 'Disconnect' $null } | Out-Null
    Add-Button -Parent $row -Text 'Wyloguj' -Icon 'E7E8' -Module $m -Danger -OnClick { param($m) & $m.Actions.Session $m 'Logoff' $null } | Out-Null
    Add-RowAction -Module $m -Text 'Wyślij wiadomość…' -Icon 'E715' -Action { param($m, $rows) & $m.Actions.Session $m 'Message' $rows }
    Add-RowAction -Module $m -Text 'Podgląd sesji (shadow)' -Icon 'E7B3' -Action {
        param($m, $rows)
        $r = $rows[0]
        $computer = [string](Get-ObjectValue $r 'Komputer')
        $id = [string](Get-ObjectValue $r 'ID')
        if ($id -notmatch '^\d+$') { return }
        Start-Tool -FilePath 'mstsc.exe' -Arguments @("/v:$computer", "/shadow:$id", '/control') -Name $computer
    }
    Add-RowAction -Module $m -Text 'Rozłącz' -Icon 'E8CD' -Action { param($m, $rows) & $m.Actions.Session $m 'Disconnect' $rows }
    Add-RowAction -Module $m -Text 'Wyloguj' -Icon 'E7E8' -Danger -Separator -Action { param($m, $rows) & $m.Actions.Session $m 'Logoff' $rows }
}

$script:LocalGroupDefs = @(
    @{ Name = 'Administratorzy'; Sid = 'S-1-5-32-544' }
    @{ Name = 'Użytkownicy pulpitu zdalnego'; Sid = 'S-1-5-32-555' }
    @{ Name = 'Użytkownicy zarządzania zdalnego (WinRM)'; Sid = 'S-1-5-32-580' }
    @{ Name = 'Czytelnicy dziennika zdarzeń'; Sid = 'S-1-5-32-573' }
    @{ Name = 'Operatorzy kopii zapasowych'; Sid = 'S-1-5-32-551' }
    @{ Name = 'Użytkownicy'; Sid = 'S-1-5-32-545' }
)

Register-Module -Workspace 'Remote' -Category 'Użytkownicy i dostęp' -Key 'LocalGroups' -Title 'Grupy lokalne' -Icon 'E902' `
    -Description 'Członkowie wbudowanych grup lokalnych (Administratorzy, Pulpit zdalny, WinRM…) wyznaczanych po SID – działa w każdym języku systemu. Podgląd, dodawanie i usuwanie członków.' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $sid = $script:LocalGroupDefs[$m.Group.SelectedIndex].Sid
        Start-HostOperation -Module $m -Name 'Grupy lokalne' -Targets $targets -Parameters @{ Sid = $sid } -ScriptBlock {
            param($P)
            $groupName = (New-Object System.Security.Principal.SecurityIdentifier($P.Sid)).Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
            $rows = @()
            try {
                $rows = @(Get-LocalGroupMember -SID $P.Sid -ErrorAction Stop | ForEach-Object {
                        [pscustomobject]@{ 'Członek' = $_.Name; 'Typ' = [string]$_.ObjectClass; 'Źródło' = [string]$_.PrincipalSource; 'SID' = [string]$_.SID }
                    })
            }
            catch {
                # Get-LocalGroupMember zawodzi np. przy osieroconych SID-ach - odczyt przez ADSI
                $group = [ADSI]("WinNT://{0}/{1},group" -f $env:COMPUTERNAME, $groupName)
                foreach ($member in @($group.Invoke('Members'))) {
                    $type = $member.GetType()
                    $path = [string]$type.InvokeMember('ADsPath', 'GetProperty', $null, $member, $null)
                    $class = [string]$type.InvokeMember('Class', 'GetProperty', $null, $member, $null)
                    $sid = ''
                    try { $sid = (New-Object System.Security.Principal.SecurityIdentifier($type.InvokeMember('objectSid', 'GetProperty', $null, $member, $null), 0)).Value } catch { }
                    $parts = @($path -replace '^WinNT://', '' -split '/')
                    $name = if ($parts.Count -ge 2) { $parts[-2] + '\' + $parts[-1] } else { $parts[-1] }
                    $rows += [pscustomobject]@{ 'Członek' = $name; 'Typ' = $class; 'Źródło' = '(ADSI)'; 'SID' = $sid }
                }
            }
            if ($rows.Count -eq 0) { return [pscustomobject]@{ 'Grupa' = $groupName; 'Członek' = '(grupa jest pusta)'; '__flag' = 'muted' } }
            $rows | ForEach-Object {
                $flag = if ($_.'Członek' -match '^S-1-') { 'warn' } else { '' }
                $_ | Add-Member -NotePropertyName 'Grupa' -NotePropertyValue $groupName -PassThru | Add-Member -NotePropertyName '__flag' -NotePropertyValue $flag -PassThru
            }
        }
    }
    $m.Actions.Remove = {
        param($m, $Rows)
        $source = @(if ($null -ne $Rows) { $Rows } else { Get-SelectedResultRows -Module $m })
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Członek') -Rows $source
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli członków grupy do usunięcia.'; return }
        $sid = $script:LocalGroupDefs[$m.Group.SelectedIndex].Sid
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) $i['Członek'] }
        if (-not (Confirm-Action -Text "Usunąć wybrane konta z grupy «$($m.Group.SelectedItem)»?" -Items $items -ConfirmText 'Usuń z grupy' -Danger)) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ GroupSid = $sid; Members = @($byHost[$h] | ForEach-Object { @{ Name = [string]$_['Członek']; Sid = [string](Get-RowValue $_['__row'] 'SID') } }) } }
        Start-HostOperation -Module $m -Name 'Usuwanie z grupy' -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            $groupName = (New-Object System.Security.Principal.SecurityIdentifier($P.GroupSid)).Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
            foreach ($member in $P.Members) {
                try {
                    # Zabezpieczenie przed odcięciem dostępu: wbudowane konto Administrator i Domain Admins w grupie Administratorzy
                    if ($P.GroupSid -eq 'S-1-5-32-544' -and ($member.Sid -match '-500$' -or $member.Sid -match '^S-1-5-21-.+-512$')) { throw 'Pominięto: wbudowane konto Administrator / grupa Domain Admins.' }
                    $identity = if ($member.Sid) { $member.Sid } else { $member.Name }
                    if (Get-Command Remove-LocalGroupMember -ErrorAction SilentlyContinue) {
                        Remove-LocalGroupMember -SID $P.GroupSid -Member $identity -ErrorAction Stop
                    }
                    else {
                        $group = [ADSI]("WinNT://{0}/{1},group" -f $env:COMPUTERNAME, $groupName)
                        $group.Remove('WinNT://' + ($member.Name -replace '\\', '/'))
                    }
                    [pscustomobject]@{ 'Członek' = $member.Name; 'Wynik' = 'Usunięto' }
                }
                catch { [pscustomobject]@{ 'Członek' = $member.Name; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Grupa'
    $m.Group = Add-ComboBox -Parent $row -Items @($script:LocalGroupDefs | ForEach-Object { $_.Name }) -Width 320
    Add-Button -Parent $row -Text 'Pokaż członków' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Usuń zaznaczonych' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Remove $m $null } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Dodaj członków'
    $m.NewMember = Add-TextBox -Parent $row2 -Width 340 -Placeholder 'DOMENA\konto lub grupa (kilka – po przecinku)'
    Add-Button -Parent $row2 -Text 'Wybierz grupy z AD…' -Icon 'E902' -Module $m -OnClick {
        param($m)
        $groups = Select-AdGroups -Title 'Grupy do dodania' -Subtitle 'Wybrane grupy domenowe zostaną dopisane do pola poniżej.'
        if (-not $groups) { return }
        $domain = if ($env:USERDOMAIN) { $env:USERDOMAIN } else { '' }
        $names = @(Split-ListText $m.NewMember.Text) + @($groups | ForEach-Object { if ($domain) { "$domain\$($_.Sam)" } else { $_.Sam } })
        $m.NewMember.Text = (@($names | Select-Object -Unique) -join ', ')
    } | Out-Null
    Add-Button -Parent $row2 -Text 'Dodaj na zaznaczonych komputerach' -Icon 'E710' -Module $m -OnClick {
        param($m)
        $members = @(Split-ListText $m.NewMember.Text)
        if ($members.Count -eq 0) { Show-Warning 'Podaj konto lub grupę, np. FIRMA\Helpdesk.'; return }
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $sid = $script:LocalGroupDefs[$m.Group.SelectedIndex].Sid
        if (-not (Confirm-Action -Text ("Dodać {0} do grupy «{1}» na {2} komputer(ach)?" -f ($members -join ', '), $m.Group.SelectedItem, $targets.Count) -Items $targets -ConfirmText 'Dodaj')) { return }
        Start-HostOperation -Module $m -Name 'Dodawanie do grupy' -Targets $targets -Output Log -Parameters @{ Members = $members; GroupSid = $sid } -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            $groupName = (New-Object System.Security.Principal.SecurityIdentifier($P.GroupSid)).Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
            foreach ($member in $P.Members) {
                try {
                    if (Get-Command Add-LocalGroupMember -ErrorAction SilentlyContinue) {
                        Add-LocalGroupMember -SID $P.GroupSid -Member $member -ErrorAction Stop
                    }
                    else {
                        $group = [ADSI]("WinNT://{0}/{1},group" -f $env:COMPUTERNAME, $groupName)
                        $group.Add('WinNT://' + ($member -replace '\\', '/'))
                    }
                    [pscustomobject]@{ 'Członek' = $member; 'Grupa' = $groupName; 'Wynik' = 'Dodano' }
                }
                catch {
                    $msg = $_.Exception.Message
                    if ($_.FullyQualifiedErrorId -like 'MemberExists*') { $msg = 'już jest członkiem grupy' }
                    [pscustomobject]@{ 'Członek' = $member; 'Grupa' = $groupName; 'Wynik' = "Błąd – $msg" }
                }
            }
        }
    } | Out-Null
    Add-RowAction -Module $m -Text 'Usuń z grupy' -Icon 'E74D' -Danger -Action { param($m, $rows) & $m.Actions.Remove $m $rows }
}

Register-Module -Workspace 'Remote' -Category 'Użytkownicy i dostęp' -Key 'LocalUsers' -Title 'Konta lokalne' -Icon 'E77B' `
    -Description 'Lokalne konta użytkowników: stan, ostatnie logowanie, wiek hasła; włączanie, wyłączanie i ustawianie hasła zaznaczonych kont.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Konta lokalne' -Targets $targets -ScriptBlock {
            param($P)
            if (Get-Command Get-LocalUser -ErrorAction SilentlyContinue) {
                Get-LocalUser | ForEach-Object {
                    $age = if ($_.PasswordLastSet) { [int]((Get-Date) - $_.PasswordLastSet).TotalDays } else { $null }
                    [pscustomobject]@{
                        'Konto'              = $_.Name
                        'Pełna nazwa'        = $_.FullName
                        'Włączone'           = $_.Enabled
                        'Ostatnie logowanie' = $_.LastLogon
                        'Hasło ustawione'    = $_.PasswordLastSet
                        'Wiek hasła (dni)'   = $age
                        'Hasło wygasa'       = $_.PasswordExpires
                        'Opis'               = $_.Description
                        'SID'                = [string]$_.SID
                        '__flag'             = $(if (-not $_.Enabled) { 'muted' } else { '' })
                    }
                }
            }
            else {
                Get-CimInstance -ClassName Win32_UserAccount -Filter 'LocalAccount = True' | ForEach-Object {
                    [pscustomobject]@{ 'Konto' = $_.Name; 'Pełna nazwa' = $_.FullName; 'Włączone' = (-not $_.Disabled); 'Opis' = $_.Description; 'SID' = $_.SID }
                }
            }
        }
    }
    $m.Actions.Change = {
        param($m, [string]$Op, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Konto') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli konta lokalne.'; return }
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) $i['Konto'] }
        $password = $null
        if ($Op -eq 'Password') {
            $password = Read-NewPassword -Subtitle 'Nowe hasło zostanie ustawione dla wszystkich zaznaczonych kont lokalnych.'
            if (-not $password) { return }
        }
        $question = @{ Enable = 'Włączyć wybrane konta lokalne?'; Disable = 'Wyłączyć wybrane konta lokalne?'; Password = 'Ustawić nowe hasło dla wybranych kont lokalnych?' }[$Op]
        if (-not (Confirm-Action -Text $question -Items $items -ConfirmText 'Wykonaj' -Danger:($Op -ne 'Enable'))) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Op = $Op; Password = $password; Names = @($byHost[$h] | ForEach-Object { [string]$_['Konto'] }) } }
        Start-HostOperation -Module $m -Name "Konta lokalne – $Op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            foreach ($name in $P.Names) {
                try {
                    switch ($P.Op) {
                        'Enable' { Enable-LocalUser -Name $name -ErrorAction Stop }
                        'Disable' { Disable-LocalUser -Name $name -ErrorAction Stop }
                        'Password' { Set-LocalUser -Name $name -Password $P.Password -ErrorAction Stop }
                    }
                    [pscustomobject]@{ 'Konto' = $name; 'Operacja' = $P.Op; 'Wynik' = 'OK' }
                }
                catch { [pscustomobject]@{ 'Konto' = $name; 'Operacja' = $P.Op; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row -Text 'Pokaż konta' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Włącz' -Icon 'E73E' -Module $m -OnClick { param($m) & $m.Actions.Change $m 'Enable' $null } | Out-Null
    Add-Button -Parent $row -Text 'Wyłącz' -Icon 'E8D8' -Module $m -Danger -OnClick { param($m) & $m.Actions.Change $m 'Disable' $null } | Out-Null
    Add-Button -Parent $row -Text 'Ustaw hasło…' -Icon 'E8D7' -Module $m -Danger -OnClick { param($m) & $m.Actions.Change $m 'Password' $null } | Out-Null
    Add-RowAction -Module $m -Text 'Włącz konto' -Icon 'E73E' -Action { param($m, $rows) & $m.Actions.Change $m 'Enable' $rows }
    Add-RowAction -Module $m -Text 'Wyłącz konto' -Icon 'E8D8' -Action { param($m, $rows) & $m.Actions.Change $m 'Disable' $rows }
    Add-RowAction -Module $m -Text 'Ustaw hasło…' -Icon 'E8D7' -Danger -Action { param($m, $rows) & $m.Actions.Change $m 'Password' $rows }
}

Register-Module -Workspace 'Remote' -Category 'Użytkownicy i dostęp' -Key 'Rdp' -Title 'Pulpit zdalny' -Icon 'E8AF' -Badge 'Nowość w wersji 4.0' `
    -Description 'Stan Pulpitu zdalnego (RDP): włączenie, uwierzytelnianie na poziomie sieci (NLA), port, reguły zapory i członkowie grupy Użytkownicy pulpitu zdalnego. Włączanie i wyłączanie RDP razem z zaporą.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('RDP')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Pulpit zdalny' -Targets $targets -ScriptBlock {
            param($P)
            $ts = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -ErrorAction SilentlyContinue
            $tcp = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -ErrorAction SilentlyContinue
            $policy = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services' -ErrorAction SilentlyContinue
            $enabled = ($ts -and [int]$ts.fDenyTSConnections -eq 0)
            $byPolicy = ($policy -and $null -ne $policy.PSObject.Properties['fDenyTSConnections'])
            $fw = $null
            try {
                $rules = @(Get-NetFirewallRule -Group '@FirewallAPI.dll,-28752' -ErrorAction Stop)
                $fw = (@($rules | Where-Object { $_.Enabled -eq 'True' -and $_.Direction -eq 'Inbound' }).Count -gt 0)
            }
            catch { }
            $members = @()
            try { $members = @(Get-LocalGroupMember -SID 'S-1-5-32-555' -ErrorAction Stop | ForEach-Object { $_.Name }) } catch { }
            $svc = Get-Service -Name TermService -ErrorAction SilentlyContinue
            [pscustomobject]@{
                'RDP'                     = $(if ($enabled) { 'Włączony' } else { 'Wyłączony' })
                'NLA'                     = $(if ($tcp) { [int]$tcp.UserAuthentication -eq 1 } else { $null })
                'Port'                    = $(if ($tcp) { [int]$tcp.PortNumber } else { $null })
                'Zapora (reguły RDP)'     = $fw
                'Usługa'                  = $(if ($svc) { [string]$svc.Status } else { '' })
                'Wymuszone zasadami'      = $byPolicy
                'Użytkownicy pulpitu zdalnego' = ($members -join ', ')
                '__tone'                  = $(if ($enabled) { 'ok' } else { '' })
            }
        }
    }
    $m.Actions.Set = {
        param($m, [string]$Op)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $text = @{
            Enable     = 'Włączyć Pulpit zdalny (z NLA) i odblokować reguły zapory?'
            Disable    = 'Wyłączyć Pulpit zdalny? Aktywne sesje RDP pozostaną do wylogowania, nowe połączenia będą odrzucane.'
            NlaOn      = 'Włączyć wymaganie uwierzytelniania na poziomie sieci (NLA)?'
            NlaOff     = 'Wyłączyć wymaganie NLA? Obniża to bezpieczeństwo – używaj tylko do diagnostyki.'
        }[$Op]
        if (-not (Confirm-Action -Text $text -Items $targets -ConfirmText 'Wykonaj' -Danger:($Op -eq 'Disable' -or $Op -eq 'NlaOff'))) { return }
        Start-HostOperation -Module $m -Name "Pulpit zdalny – $Op" -Targets $targets -Output Log -Parameters @{ Op = $Op } -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            $ts = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server'
            $tcp = "$ts\WinStations\RDP-Tcp"
            switch ($P.Op) {
                'Enable' {
                    Set-ItemProperty -Path $ts -Name fDenyTSConnections -Value 0 -Type DWord
                    Set-ItemProperty -Path $tcp -Name UserAuthentication -Value 1 -Type DWord
                    try { Enable-NetFirewallRule -Group '@FirewallAPI.dll,-28752' -ErrorAction Stop } catch { 'Uwaga: nie udało się włączyć reguł zapory – ' + $_.Exception.Message }
                    'Pulpit zdalny włączony (NLA wymagane).'
                }
                'Disable' {
                    Set-ItemProperty -Path $ts -Name fDenyTSConnections -Value 1 -Type DWord
                    try { Disable-NetFirewallRule -Group '@FirewallAPI.dll,-28752' -ErrorAction Stop } catch { }
                    'Pulpit zdalny wyłączony.'
                }
                'NlaOn' { Set-ItemProperty -Path $tcp -Name UserAuthentication -Value 1 -Type DWord; 'NLA włączone.' }
                'NlaOff' { Set-ItemProperty -Path $tcp -Name UserAuthentication -Value 0 -Type DWord; 'NLA wyłączone.' }
            }
            if (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services' -Name fDenyTSConnections -ErrorAction SilentlyContinue) {
                'Uwaga: ustawienie RDP jest wymuszone zasadami grupy – zmiana może zostać nadpisana.'
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row -Text 'Sprawdź stan' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Włącz RDP' -Icon 'E73E' -Module $m -OnClick { param($m) & $m.Actions.Set $m 'Enable' } | Out-Null
    Add-Button -Parent $row -Text 'Wyłącz RDP' -Icon 'E711' -Module $m -Danger -OnClick { param($m) & $m.Actions.Set $m 'Disable' } | Out-Null
    Add-Button -Parent $row -Text 'Wymagaj NLA' -Icon 'E72E' -Module $m -OnClick { param($m) & $m.Actions.Set $m 'NlaOn' } | Out-Null
    Add-Button -Parent $row -Text 'Wyłącz NLA' -Icon 'E785' -Module $m -Danger -OnClick { param($m) & $m.Actions.Set $m 'NlaOff' } | Out-Null
    Add-RowAction -Module $m -Text 'Połącz (mstsc)' -Icon 'E8AF' -Action {
        param($m, $rows)
        $computer = [string](Get-ObjectValue $rows[0] 'Komputer')
        Start-Tool -FilePath 'mstsc.exe' -Arguments @("/v:$computer") -Name $computer
    }
}
#endregion

#region Zarządzanie zdalne: Zdalne wykonanie
Register-Module -Workspace 'Remote' -Category 'Zdalne wykonanie' -Key 'Commands' -Title 'Polecenia' -Icon 'E756' `
    -Description 'Uruchamia polecenia PowerShell lub cmd.exe na zaznaczonych komputerach (sesja WinRM, bez pulpitu użytkownika). Pełny wynik w panelu szczegółów lub po dwukliku.' -Build {
    param($m)
    $m.Templates = @(
        @{ Name = '(wybierz szablon polecenia)'; Mode = ''; Text = '' }
        @{ Name = 'Wersja systemu i czas pracy'; Mode = 'PS'; Text = 'Get-CimInstance Win32_OperatingSystem | Select-Object Caption, Version, LastBootUpTime' }
        @{ Name = 'Ostatnie poprawki (10)'; Mode = 'PS'; Text = 'Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 10 HotFixID, Description, InstalledOn' }
        @{ Name = 'Konfiguracja IP'; Mode = 'CMD'; Text = 'ipconfig /all' }
        @{ Name = 'Odśwież DNS (flushdns + registerdns)'; Mode = 'CMD'; Text = 'ipconfig /flushdns && ipconfig /registerdns' }
        @{ Name = 'Wyczyść bilety Kerberos komputera'; Mode = 'CMD'; Text = 'klist -li 0x3e7 purge' }
        @{ Name = 'Wynikowe zasady komputera (gpresult)'; Mode = 'CMD'; Text = 'gpresult /r /scope computer' }
        @{ Name = 'Procesy – top 10 pamięci'; Mode = 'PS'; Text = "Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 Name, Id, @{n='MB';e={[math]::Round(`$_.WorkingSet64/1MB)}}" }
        @{ Name = 'Test połączenia z kontrolerem domeny'; Mode = 'PS'; Text = 'Test-NetConnection -ComputerName $env:USERDNSDOMAIN -Port 389 | Select-Object ComputerName, RemoteAddress, TcpTestSucceeded' }
        @{ Name = 'Kanał zaufania z domeną'; Mode = 'PS'; Text = 'Test-ComputerSecureChannel -Verbose 4>&1' }
        @{ Name = 'Synchronizacja czasu (w32tm)'; Mode = 'CMD'; Text = 'w32tm /query /status && w32tm /resync' }
        @{ Name = 'Naprawa obrazu systemu (DISM + SFC)'; Mode = 'CMD'; Text = 'DISM /Online /Cleanup-Image /RestoreHealth && sfc /scannow' }
    )
    $row = Add-ToolbarRow -Module $m -Title 'Interpreter'
    $m.Mode = Add-Segmented -Parent $row -Items @('PowerShell', 'cmd.exe')
    Add-Label -Parent $row -Text '   Szablon' | Out-Null
    $m.Template = Add-ComboBox -Parent $row -Items @($m.Templates | ForEach-Object { $_.Name }) -Width 340
    Register-ControlHandler -Control $m.Template -EventName 'SelectionChanged' -Module $m -Action {
        param($m, $s)
        $t = $m.Templates[$s.SelectedIndex]
        if (-not $t.Mode) { return }
        Set-SegmentIndex $m.Mode $(if ($t.Mode -eq 'PS') { 0 } else { 1 })
        $m.CommandBox.Text = $t.Text
    }
    $m.CommandBox = Add-StretchTextBox -Module $m -Title 'Polecenie' -Multiline -Height 130 -Placeholder 'Wpisz polecenie lub skrypt…'
    $row3 = Add-ToolbarRow -Module $m -Title ' '
    Add-Button -Parent $row3 -Text 'Uruchom na zaznaczonych' -Icon 'E768' -Module $m -Primary -OnClick {
        param($m)
        $command = $m.CommandBox.Text.Trim()
        if (-not $command) { Show-Warning 'Wpisz polecenie do wykonania.'; return }
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $isPs = ((Get-SegmentIndex $m.Mode) -eq 0)
        $mode = if ($isPs) { 'PowerShell' } else { 'cmd.exe' }
        $preview = if ($command.Length -gt 400) { $command.Substring(0, 400) + '…' } else { $command }
        if (-not (Confirm-Action -Text "Uruchomić polecenie ($mode) na $($targets.Count) komputer(ach)?`r`n`r`n$preview" -Items $targets -ConfirmText 'Uruchom')) { return }
        Write-Log "Polecenie ($mode): $command"
        if ($isPs) {
            Start-HostOperation -Module $m -Name 'Polecenie PowerShell' -Targets $targets -Parameters @{ Command = $command } -ScriptBlock {
                param($P)
                $sb = [scriptblock]::Create($P.Command)
                $text = (& $sb 2>&1 | Out-String -Width 250).Trim()
                [pscustomobject]@{ 'Wynik' = $text }
            }
        }
        else {
            Start-HostOperation -Module $m -Name 'Polecenie cmd.exe' -Targets $targets -Parameters @{ Command = $command } -ScriptBlock {
                param($P)
                $oem = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = Join-Path $env:SystemRoot 'System32\cmd.exe'
                $psi.Arguments = '/d /s /c "' + $P.Command + '"'
                $psi.UseShellExecute = $false
                $psi.RedirectStandardOutput = $true
                $psi.RedirectStandardError = $true
                $psi.CreateNoWindow = $true
                $psi.StandardOutputEncoding = $oem
                $psi.StandardErrorEncoding = $oem
                $proc = [System.Diagnostics.Process]::Start($psi)
                $errTask = $proc.StandardError.ReadToEndAsync()
                $stdout = $proc.StandardOutput.ReadToEnd()
                $proc.WaitForExit()
                $stderr = $errTask.Result
                $text = $stdout
                if ($stderr) { $text += "`r`n" + $stderr }
                [pscustomobject]@{ 'Kod wyjścia' = $proc.ExitCode; 'Wynik' = $text.Trim(); '__flag' = $(if ($proc.ExitCode -ne 0) { 'warn' } else { '' }) }
            }
        }
    } | Out-Null
    Add-Label -Parent $row3 -Text 'Polecenia działają bez profilu i pulpitu użytkownika; zasoby sieciowe mogą być niedostępne (podwójny przeskok).' -Hint -MaxWidth 560 | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'Zdalne wykonanie' -Key 'Install' -Title 'Instalacja oprogramowania' -Icon 'E896' `
    -Description 'Kopiuje instalator (MSI, MSP, MSU, EXE) na zaznaczone komputery i uruchamia go w trybie cichym. Wynik zawiera kod wyjścia i jego znaczenie.' -Build {
    param($m)
    $m.PillColumns = @('Wynik')
    $row = Add-ToolbarRow -Module $m -Title 'Instalator'
    $m.Path = Add-TextBox -Parent $row -Width 560 -Placeholder 'Ścieżka do pliku .msi, .msp, .msu lub .exe'
    Add-Button -Parent $row -Text 'Przeglądaj…' -Icon 'E8E5' -Module $m -AlwaysEnabled -OnClick {
        param($m)
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Filter = 'Instalatory (*.msi;*.msp;*.msu;*.exe)|*.msi;*.msp;*.msu;*.exe|Wszystkie pliki (*.*)|*.*'
        if ($dlg.ShowDialog($script:UI.Window) -eq $true) { $m.Path.Text = $dlg.FileName }
    } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Opcje'
    $m.Args = Add-TextBox -Parent $row2 -Width 300 -Placeholder 'Dodatkowe argumenty (EXE: np. /S, /quiet)'
    $m.Cleanup = Add-CheckBox -Parent $row2 -Text 'Usuń instalator po zakończeniu' -Checked $true
    Add-Button -Parent $row2 -Text 'Zainstaluj na zaznaczonych' -Icon 'E896' -Module $m -Primary -OnClick {
        param($m)
        $path = $m.Path.Text.Trim().Trim('"')
        if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { Show-Warning 'Wskaż istniejący plik instalatora.'; return }
        $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
        if (@('.msi', '.msp', '.msu', '.exe') -notcontains $ext) { Show-Warning 'Obsługiwane są pliki .msi, .msp, .msu i .exe.'; return }
        $extra = $m.Args.Text.Trim()
        if ($ext -eq '.exe' -and -not $extra) {
            if (-not (Confirm-Action -Text 'Nie podano argumentów cichej instalacji dla pliku EXE. Instalator może czekać na odpowiedź użytkownika, którego nie ma – operacja zawiśnie do czasu przerwania. Kontynuować?' -ConfirmText 'Kontynuuj')) { return }
        }
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        if (-not (Confirm-Action -Text "Zainstalować $([System.IO.Path]::GetFileName($path)) na $($targets.Count) komputer(ach)?" -Items $targets -ConfirmText 'Zainstaluj')) { return }
        $params = @{ LocalPath = $path; FileName = [System.IO.Path]::GetFileName($path); Args = $extra; Cleanup = (Test-Checked $m.Cleanup) }
        Start-HostOperation -Module $m -Name 'Instalacja' -Targets $targets -Local -Parameters $params -ScriptBlock {
            param($Target, $P, $Ctx)
            $ErrorActionPreference = 'Stop'
            $sp = @{ ComputerName = $Target }
            if ($Ctx.Credential) { $sp.Credential = $Ctx.Credential }
            if ($Ctx.SessionOption) { $sp.SessionOption = $Ctx.SessionOption }
            $session = New-PSSession @sp
            try {
                $remoteFile = Invoke-Command -Session $session -ArgumentList $P.FileName -ScriptBlock {
                    param($name)
                    $dir = Join-Path $env:SystemRoot 'Temp\DomainOps'
                    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                    Join-Path $dir $name
                }
                # Najpierw szybka kopia przez udział ADMIN$ (tylko dla bieżących poświadczeń), potem przez WinRM
                $method = ''
                if (-not $Ctx.Credential) {
                    try {
                        Copy-Item -LiteralPath $P.LocalPath -Destination ('\\{0}\ADMIN$\Temp\DomainOps\{1}' -f $Target, $P.FileName) -Force
                        $method = 'SMB'
                    }
                    catch { }
                }
                if (-not $method) {
                    Copy-Item -LiteralPath $P.LocalPath -Destination $remoteFile -ToSession $session -Force
                    $method = 'WinRM'
                }
                $res = Invoke-Command -Session $session -ArgumentList $remoteFile, $P.Args, $P.Cleanup -ScriptBlock {
                    param($File, $ExtraArgs, $Cleanup)
                    $ext = [System.IO.Path]::GetExtension($File).ToLowerInvariant()
                    $log = [System.IO.Path]::ChangeExtension($File, '.log')
                    $exe = $File
                    $arguments = [string]$ExtraArgs
                    switch ($ext) {
                        '.msi' { $exe = Join-Path $env:SystemRoot 'System32\msiexec.exe'; $arguments = ('/i "{0}" /qn /norestart /l*v "{1}" {2}' -f $File, $log, $ExtraArgs) }
                        '.msp' { $exe = Join-Path $env:SystemRoot 'System32\msiexec.exe'; $arguments = ('/p "{0}" /qn /norestart /l*v "{1}" {2}' -f $File, $log, $ExtraArgs) }
                        '.msu' { $exe = Join-Path $env:SystemRoot 'System32\wusa.exe'; $arguments = ('"{0}" /quiet /norestart {1}' -f $File, $ExtraArgs); $log = '' }
                        default { $log = '' }
                    }
                    $sw = [System.Diagnostics.Stopwatch]::StartNew()
                    $spArgs = @{ FilePath = $exe; PassThru = $true; WindowStyle = 'Hidden' }
                    if ($arguments.Trim()) { $spArgs.ArgumentList = $arguments.Trim() }
                    $proc = Start-Process @spArgs
                    $null = $proc.Handle
                    $proc.WaitForExit()
                    $code = $proc.ExitCode
                    $known = @{
                        0           = 'Sukces'
                        3010        = 'Sukces – wymagany restart'
                        1641        = 'Sukces – instalator uruchomił restart'
                        1602        = 'Błąd – instalacja anulowana'
                        1603        = 'Błąd – krytyczny błąd instalacji (1603)'
                        1618        = 'Błąd – trwa inna instalacja (1618)'
                        1619        = 'Błąd – nie można otworzyć pakietu (1619)'
                        1625        = 'Błąd – instalacja zablokowana przez zasady (1625)'
                        1633        = 'Błąd – nieobsługiwana platforma (1633)'
                        1638        = 'Błąd – zainstalowana jest inna wersja produktu (1638)'
                        2359302     = 'Aktualizacja jest już zainstalowana'
                        -2145124329 = 'Aktualizacja nie dotyczy tego systemu'
                    }
                    $meaning = if ($known.ContainsKey($code)) { $known[$code] } else { "Błąd – kod wyjścia $code" }
                    if ($Cleanup) {
                        Start-Sleep -Seconds 1
                        Remove-Item -LiteralPath $File -Force -ErrorAction SilentlyContinue
                    }
                    [pscustomobject]@{ Plik = [System.IO.Path]::GetFileName($File); Kod = $code; Wynik = $meaning; Czas = [Math]::Round($sw.Elapsed.TotalSeconds); Log = $log }
                }
                [pscustomobject]@{
                    'Plik'            = $res.Plik
                    'Wynik'           = $res.Wynik
                    'Kod wyjścia'     = $res.Kod
                    'Czas (s)'        = $res.Czas
                    'Kopiowanie'      = $method
                    'Log instalatora' = $res.Log
                    '__tone'          = $(if ($res.Wynik -like 'Sukces*' -or $res.Wynik -like 'Aktualizacja jest już*') { 'ok' } elseif ($res.Wynik -like 'Aktualizacja nie dotyczy*') { 'info' } else { 'crit' })
                }
            }
            finally {
                Remove-PSSession -Session $session -ErrorAction SilentlyContinue
            }
        }
    } | Out-Null
    Add-Label -Parent (Add-ToolbarRow -Module $m -Title ' ') -Text 'MSI/MSP: automatycznie /qn /norestart i log w %SystemRoot%\Temp\DomainOps;  MSU: /quiet /norestart;  EXE: podaj przełączniki cichej instalacji.' -Hint | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'Zdalne wykonanie' -Key 'GPUpdate' -Title 'Aktualizacja zasad grupy' -Icon 'E895' `
    -Description 'Wymusza odświeżenie zasad grupy (gpupdate) na zaznaczonych komputerach.' -Build {
    param($m)
    $m.PillColumns = @('Wynik')
    $row = Add-ToolbarRow -Module $m -Title 'Parametry'
    Add-Label -Parent $row -Text 'Zakres' | Out-Null
    $m.Scope = Add-ComboBox -Parent $row -Items @('Komputer', 'Komputer i użytkownik', 'Użytkownik') -Width 200
    $m.Force = Add-CheckBox -Parent $row -Text 'Wymuś ponowne zastosowanie (/force)' -Checked $true
    Add-Button -Parent $row -Text 'Uruchom gpupdate' -Icon 'E895' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $scope = @('Computer', '', 'User')[$m.Scope.SelectedIndex]
        Start-HostOperation -Module $m -Name 'GPUpdate' -Targets $targets -Parameters @{ Target = $scope; Force = (Test-Checked $m.Force) } -ScriptBlock {
            param($P)
            $oem = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
            $arguments = @()
            if ($P.Target) { $arguments += "/target:$($P.Target)" }
            if ($P.Force) { $arguments += '/force' }
            $arguments += '/wait:600'
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = Join-Path $env:SystemRoot 'System32\gpupdate.exe'
            $psi.Arguments = $arguments -join ' '
            $psi.UseShellExecute = $false
            $psi.RedirectStandardOutput = $true
            $psi.RedirectStandardError = $true
            $psi.RedirectStandardInput = $true
            $psi.CreateNoWindow = $true
            $psi.StandardOutputEncoding = $oem
            $psi.StandardErrorEncoding = $oem
            $proc = [System.Diagnostics.Process]::Start($psi)
            # gpupdate może pytać o wylogowanie/restart - odpowiadamy "N"
            $proc.StandardInput.WriteLine('N')
            $proc.StandardInput.WriteLine('N')
            $proc.StandardInput.Close()
            $errTask = $proc.StandardError.ReadToEndAsync()
            $stdout = $proc.StandardOutput.ReadToEnd()
            $proc.WaitForExit()
            $lines = @(($stdout + "`n" + $errTask.Result) -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            [pscustomobject]@{
                'Wynik'       = $(if ($proc.ExitCode -eq 0) { 'OK' } else { "Błąd – kod $($proc.ExitCode)" })
                'Kod wyjścia' = $proc.ExitCode
                'Komunikaty'  = ($lines -join ' | ')
                '__tone'      = $(if ($proc.ExitCode -eq 0) { 'ok' } else { 'crit' })
            }
        }
    } | Out-Null
}
#endregion

#region Zarządzanie zdalne: System
function Get-HostItemList {
    # Lista "komputer: element" do okna potwierdzenia
    param([System.Collections.IDictionary]$ByHost, [scriptblock]$Format)
    return @(foreach ($h in $ByHost.Keys) { foreach ($i in $ByHost[$h]) { '{0}: {1}' -f $h, (& $Format $i) } })
}

Register-Module -Workspace 'Remote' -Category 'System' -Key 'Services' -Title 'Usługi' -Icon 'E9F5' `
    -Description 'Usługi na zaznaczonych komputerach. Akcje (przyciski lub menu pod prawym przyciskiem) dotyczą usług zaznaczonych w tabeli – na właściwych komputerach.' -Build {
    param($m)
    $m.PillColumns = @('Stan')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $state = @('All', 'Running', 'Stopped', 'AutoStopped')[$m.StateFilter.SelectedIndex]
        Start-HostOperation -Module $m -Name 'Usługi' -Targets $targets -Parameters @{ Filter = $m.Filter.Text.Trim(); State = $state } -ScriptBlock {
            param($P)
            $services = @(Get-CimInstance -ClassName Win32_Service)
            if ($P.Filter) {
                $f = "*$($P.Filter)*"
                $services = @($services | Where-Object { $_.Name -like $f -or $_.DisplayName -like $f })
            }
            switch ($P.State) {
                'Running' { $services = @($services | Where-Object { $_.State -eq 'Running' }) }
                'Stopped' { $services = @($services | Where-Object { $_.State -ne 'Running' }) }
                'AutoStopped' { $services = @($services | Where-Object { $_.StartMode -eq 'Auto' -and $_.State -ne 'Running' }) }
            }
            $services | Sort-Object DisplayName | ForEach-Object {
                $tone = if ($_.State -eq 'Running') { 'ok' } elseif ($_.StartMode -eq 'Auto') { 'crit' } elseif ($_.StartMode -eq 'Disabled') { '' } else { 'info' }
                [pscustomobject]@{
                    'Nazwa'             = $_.Name
                    'Nazwa wyświetlana' = $_.DisplayName
                    'Stan'              = $_.State
                    'Uruchamianie'      = $_.StartMode
                    'Konto'             = $_.StartName
                    'PID'               = $_.ProcessId
                    'Ścieżka'           = $_.PathName
                    '__tone'            = $tone
                }
            }
        }
    }
    $m.Actions.Change = {
        param($m, [string]$Op, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Nazwa') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli usługi, których dotyczy operacja.'; return }
        $startup = @('Automatic', 'Manual', 'Disabled')[$m.StartupType.SelectedIndex]
        $question = @{
            Start       = 'Uruchomić wybrane usługi?'
            Stop        = 'Zatrzymać wybrane usługi? Zatrzymane zostaną też usługi od nich zależne.'
            Restart     = 'Uruchomić ponownie wybrane usługi?'
            StartupType = "Ustawić typ uruchamiania «$($m.StartupType.SelectedItem)» dla wybranych usług?"
        }[$Op]
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) $i['Nazwa'] }
        if (-not (Confirm-Action -Text $question -Items $items -ConfirmText 'Wykonaj' -Danger:($Op -eq 'Stop'))) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Op = $Op; StartupType = $startup; Names = @($byHost[$h] | ForEach-Object { [string]$_['Nazwa'] }) } }
        Start-HostOperation -Module $m -Name "Usługi – $Op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            foreach ($name in $P.Names) {
                try {
                    switch ($P.Op) {
                        'Start' { Start-Service -Name $name -ErrorAction Stop }
                        'Stop' { Stop-Service -Name $name -Force -ErrorAction Stop }
                        'Restart' { Restart-Service -Name $name -Force -ErrorAction Stop }
                        'StartupType' { Set-Service -Name $name -StartupType $P.StartupType -ErrorAction Stop }
                    }
                    $svc = Get-Service -Name $name
                    [pscustomobject]@{ 'Usługa' = $name; 'Operacja' = $P.Op; 'Wynik' = 'OK'; 'Stan' = [string]$svc.Status }
                }
                catch {
                    [pscustomobject]@{ 'Usługa' = $name; 'Operacja' = $P.Op; 'Wynik' = "Błąd – $($_.Exception.Message)" }
                }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Lista'
    Add-Label -Parent $row -Text 'Nazwa' | Out-Null
    $m.Filter = Add-TextBox -Parent $row -Width 170 -Placeholder 'fragment nazwy'
    Add-Label -Parent $row -Text 'Stan' | Out-Null
    $m.StateFilter = Add-ComboBox -Parent $row -Items @('Wszystkie', 'Uruchomione', 'Zatrzymane', 'Automatyczne, ale zatrzymane') -Width 230
    Add-Button -Parent $row -Text 'Pokaż usługi' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Zaznaczone usługi'
    foreach ($a in @(@('Start', 'Start', 'E768'), @('Stop', 'Stop', 'E71A'), @('Restart', 'Restart', 'E72C'))) {
        $b = Add-Button -Parent $row2 -Text $a[0] -Icon $a[2] -Module $m -OnClick { param($m, $s) & $m.Actions.Change $m ([string]$s.Tag) $null }
        $b.Tag = $a[1]
    }
    Add-Label -Parent $row2 -Text '   Typ uruchamiania' | Out-Null
    $m.StartupType = Add-ComboBox -Parent $row2 -Items @('Automatyczny', 'Ręczny', 'Wyłączony') -Width 140
    $b = Add-Button -Parent $row2 -Text 'Ustaw' -Module $m -OnClick { param($m, $s) & $m.Actions.Change $m 'StartupType' $null }
    Add-RowAction -Module $m -Text 'Uruchom' -Icon 'E768' -Action { param($m, $rows) & $m.Actions.Change $m 'Start' $rows }
    Add-RowAction -Module $m -Text 'Zatrzymaj' -Icon 'E71A' -Action { param($m, $rows) & $m.Actions.Change $m 'Stop' $rows }
    Add-RowAction -Module $m -Text 'Uruchom ponownie' -Icon 'E72C' -Action { param($m, $rows) & $m.Actions.Change $m 'Restart' $rows }
}

Register-Module -Workspace 'Remote' -Category 'System' -Key 'Processes' -Title 'Procesy' -Icon 'E7EF' `
    -Description 'Procesy uruchomione na zaznaczonych komputerach (pamięć, właściciel, wiersz poleceń) i kończenie wybranych procesów.' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Procesy' -Targets $targets -Parameters @{ Filter = $m.Filter.Text.Trim(); WithOwner = (Test-Checked $m.WithOwner) } -ScriptBlock {
            param($P)
            $procs = @(Get-CimInstance -ClassName Win32_Process)
            if ($P.Filter) { $procs = @($procs | Where-Object { $_.Name -like "*$($P.Filter)*" }) }
            $procs | Sort-Object WorkingSetSize -Descending | ForEach-Object {
                $owner = ''
                if ($P.WithOwner) {
                    try {
                        $o = Invoke-CimMethod -InputObject $_ -MethodName GetOwner -ErrorAction Stop
                        if ($o.User) { $owner = '{0}\{1}' -f $o.Domain, $o.User }
                    }
                    catch { }
                }
                [pscustomobject]@{
                    'Proces'         = $_.Name
                    'PID'            = [int]$_.ProcessId
                    'Pamięć (MB)'    = [Math]::Round($_.WorkingSetSize / 1MB, 1)
                    'Właściciel'     = $owner
                    'Sesja'          = $_.SessionId
                    'Uruchomiony'    = $_.CreationDate
                    'Wiersz poleceń' = $_.CommandLine
                }
            }
        }
    }
    $m.Actions.Kill = {
        param($m, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('PID', 'Proces') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli procesy do zakończenia.'; return }
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) '{0} (PID {1})' -f $i['Proces'], $i['PID'] }
        if (-not (Confirm-Action -Text 'Zakończyć wybrane procesy? Niezapisane dane w tych programach zostaną utracone.' -Items $items -ConfirmText 'Zakończ procesy' -Danger)) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Ids = @($byHost[$h] | ForEach-Object { [int]$_['PID'] }) } }
        Start-HostOperation -Module $m -Name 'Kończenie procesów' -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            foreach ($id in $P.Ids) {
                try {
                    $proc = Get-Process -Id $id -ErrorAction Stop
                    Stop-Process -Id $id -Force -ErrorAction Stop
                    [pscustomobject]@{ 'Proces' = $proc.ProcessName; 'PID' = $id; 'Wynik' = 'Zakończono' }
                }
                catch { [pscustomobject]@{ 'PID' = $id; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Lista'
    Add-Label -Parent $row -Text 'Nazwa' | Out-Null
    $m.Filter = Add-TextBox -Parent $row -Width 170 -Placeholder 'np. chrome'
    $m.WithOwner = Add-CheckBox -Parent $row -Text 'Pokaż właściciela (wolniej)'
    Add-Button -Parent $row -Text 'Pokaż procesy' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Zakończ zaznaczone' -Icon 'E711' -Module $m -Danger -OnClick { param($m) & $m.Actions.Kill $m $null } | Out-Null
    Add-RowAction -Module $m -Text 'Zakończ proces' -Icon 'E711' -Danger -Action { param($m, $rows) & $m.Actions.Kill $m $rows }
}

Register-Module -Workspace 'Remote' -Category 'System' -Key 'Disks' -Title 'Dyski' -Icon 'EDA2' `
    -Description 'Zajętość dysków lokalnych oraz czyszczenie plików tymczasowych i Kosza (z raportem odzyskanego miejsca).' -Build {
    param($m)
    $m.PillColumns = @('Stan')
    $row = Add-ToolbarRow -Module $m -Title 'Zajętość'
    Add-Button -Parent $row -Text 'Pokaż dyski' -Icon 'EDA2' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Dyski' -Targets $targets -ScriptBlock {
            param($P)
            Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 3' | ForEach-Object {
                $size = [double]$_.Size
                $free = [double]$_.FreeSpace
                $freePct = if ($size -gt 0) { [Math]::Round($free / $size * 100, 1) } else { $null }
                $tone = if ($null -eq $freePct) { '' } elseif ($freePct -le 5) { 'crit' } elseif ($freePct -le 15) { 'warn' } else { 'ok' }
                [pscustomobject]@{
                    'Dysk'          = $_.DeviceID
                    'Stan'          = $(switch ($tone) { 'crit' { 'krytycznie mało' } 'warn' { 'mało miejsca' } 'ok' { 'OK' } default { '' } })
                    'Etykieta'      = $_.VolumeName
                    'System plików' = $_.FileSystem
                    'Rozmiar (GB)'  = [Math]::Round($size / 1GB, 1)
                    'Wolne (GB)'    = [Math]::Round($free / 1GB, 1)
                    'Wolne (%)'     = $freePct
                    '__tone'        = $tone
                }
            }
        }
    } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Czyszczenie'
    $m.WinTemp = Add-CheckBox -Parent $row2 -Text 'Windows\Temp' -Checked $true
    $m.UserTemp = Add-CheckBox -Parent $row2 -Text 'TEMP profili użytkowników' -Checked $true
    $m.Recycle = Add-CheckBox -Parent $row2 -Text 'Kosz (wszystkie dyski)' -Checked $true
    $m.WuCache = Add-CheckBox -Parent $row2 -Text 'Pobrane aktualizacje (SoftwareDistribution\Download)' -ToolTip 'Usługa Windows Update zostanie na chwilę zatrzymana'
    Add-Label -Parent $row2 -Text 'Starsze niż (dni)' | Out-Null
    $m.Days = Add-Numeric -Parent $row2 -Value 2 -Minimum 0 -Maximum 365 -Width 60
    Add-Button -Parent $row2 -Text 'Wyczyść na zaznaczonych' -Icon 'E74D' -Module $m -Danger -OnClick {
        param($m)
        if (-not ((Test-Checked $m.WinTemp) -or (Test-Checked $m.UserTemp) -or (Test-Checked $m.Recycle) -or (Test-Checked $m.WuCache))) { Show-Warning 'Wybierz, co ma zostać wyczyszczone.'; return }
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        if (-not (Confirm-Action -Text "Usunąć pliki tymczasowe na $($targets.Count) komputer(ach)?" -Items $targets -ConfirmText 'Wyczyść' -Danger)) { return }
        $params = @{ WindowsTemp = (Test-Checked $m.WinTemp); UserTemp = (Test-Checked $m.UserTemp); RecycleBin = (Test-Checked $m.Recycle); WuCache = (Test-Checked $m.WuCache); OlderThanDays = (Get-Num $m.Days) }
        Start-HostOperation -Module $m -Name 'Czyszczenie dysków' -Targets $targets -Parameters $params -ScriptBlock {
            param($P)
            $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $env:SystemDrive)
            $before = [double]$drive.FreeSpace
            $limit = (Get-Date).AddDays( - [int]$P.OlderThanDays)
            $roots = @()
            if ($P.WindowsTemp) { $roots += (Join-Path $env:SystemRoot 'Temp') }
            if ($P.UserTemp) {
                $profiles = Get-CimInstance -ClassName Win32_UserProfile -Filter 'Special = False' -ErrorAction SilentlyContinue
                foreach ($pr in $profiles) {
                    $t = Join-Path $pr.LocalPath 'AppData\Local\Temp'
                    if (Test-Path -LiteralPath $t) { $roots += $t }
                }
            }
            $removed = 0
            $failed = 0
            foreach ($root in $roots) {
                Get-ChildItem -LiteralPath $root -Recurse -Force -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.LastWriteTime -lt $limit -and $_.FullName -notlike '*\Temp\DomainOps\*' } |
                    ForEach-Object {
                        try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop; $removed++ } catch { $failed++ }
                    }
                Get-ChildItem -LiteralPath $root -Recurse -Force -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.FullName -notlike '*\Temp\DomainOps*' } |
                    Sort-Object { $_.FullName.Length } -Descending |
                    ForEach-Object {
                        if (-not (Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue)) {
                            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                        }
                    }
            }
            if ($P.RecycleBin) {
                foreach ($d in Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 3') {
                    $bin = $d.DeviceID + '\$Recycle.Bin'
                    if (-not (Test-Path -LiteralPath $bin)) { continue }
                    foreach ($sidDir in Get-ChildItem -LiteralPath $bin -Force -Directory -ErrorAction SilentlyContinue) {
                        Get-ChildItem -LiteralPath $sidDir.FullName -Force -ErrorAction SilentlyContinue |
                            Where-Object { $_.Name -ne 'desktop.ini' } |
                            ForEach-Object {
                                try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop; $removed++ } catch { $failed++ }
                            }
                    }
                }
            }
            if ($P.WuCache) {
                $download = Join-Path $env:SystemRoot 'SoftwareDistribution\Download'
                try {
                    Stop-Service -Name wuauserv -Force -ErrorAction Stop
                    Get-ChildItem -LiteralPath $download -Force -ErrorAction SilentlyContinue | ForEach-Object {
                        try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop; $removed++ } catch { $failed++ }
                    }
                }
                finally { Start-Service -Name wuauserv -ErrorAction SilentlyContinue }
            }
            $after = [double](Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $env:SystemDrive)).FreeSpace
            [pscustomobject]@{
                'Usunięte elementy'            = $removed
                'Nie udało się usunąć'         = $failed
                'Odzyskano na systemowym (GB)' = [Math]::Round(($after - $before) / 1GB, 2)
                'Wolne na systemowym (GB)'     = [Math]::Round($after / 1GB, 1)
            }
        }
    } | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'System' -Key 'Events' -Title 'Dziennik zdarzeń' -Icon 'E81C' `
    -Description 'Zdarzenia z wybranych dzienników z ostatnich godzin. Gotowe zestawy pomagają szybko znaleźć typowe problemy (nieoczekiwane wyłączenia, błędy dysku, nieudane logowania).' -Build {
    param($m)
    $m.PillColumns = @('Poziom')
    $m.Presets = @(
        @{ Name = '(własne ustawienia)' }
        @{ Name = 'Nieoczekiwane wyłączenia i restarty'; Logs = @('System'); Ids = '41, 6008, 1074, 1076'; Levels = @(1, 2, 3, 4) }
        @{ Name = 'Błędy dysków i systemu plików'; Logs = @('System'); Ids = '7, 11, 15, 51, 55, 98, 129, 153'; Levels = @(1, 2, 3) }
        @{ Name = 'Awarie aplikacji'; Logs = @('Application'); Ids = '1000, 1001, 1002, 1026'; Levels = @(1, 2, 4) }
        @{ Name = 'Nieudane logowania (Security 4625)'; Logs = @('Security'); Ids = '4625'; Levels = @() }
        @{ Name = 'Blokady kont (Security 4740)'; Logs = @('Security'); Ids = '4740'; Levels = @() }
        @{ Name = 'Zasady grupy – błędy'; Logs = @('System'); Ids = '1085, 1096, 1125, 1129'; Levels = @(1, 2, 3) }
        @{ Name = 'Windows Update – instalacje'; Logs = @('System'); Ids = '19, 20, 43'; Levels = @(1, 2, 3, 4) }
    )
    $row = Add-ToolbarRow -Module $m -Title 'Zestaw'
    $m.Preset = Add-ComboBox -Parent $row -Items @($m.Presets | ForEach-Object { $_.Name }) -Width 320
    Register-ControlHandler -Control $m.Preset -EventName 'SelectionChanged' -Module $m -Action {
        param($m, $s)
        $p = $m.Presets[$s.SelectedIndex]
        if (-not $p.ContainsKey('Logs')) { return }
        $m.LogSystem.IsChecked = ($p.Logs -contains 'System')
        $m.LogApp.IsChecked = ($p.Logs -contains 'Application')
        $m.LogSec.IsChecked = ($p.Logs -contains 'Security')
        $m.LvlCrit.IsChecked = ($p.Levels -contains 1)
        $m.LvlErr.IsChecked = ($p.Levels -contains 2)
        $m.LvlWarn.IsChecked = ($p.Levels -contains 3)
        $m.LvlInfo.IsChecked = ($p.Levels -contains 4)
        $m.Ids.Text = $p.Ids
    }
    $row1 = Add-ToolbarRow -Module $m -Title 'Dzienniki'
    $m.LogSystem = Add-CheckBox -Parent $row1 -Text 'System' -Checked $true
    $m.LogApp = Add-CheckBox -Parent $row1 -Text 'Application' -Checked $true
    $m.LogSec = Add-CheckBox -Parent $row1 -Text 'Security'
    Add-Label -Parent $row1 -Text '   Poziom' -Hint | Out-Null
    $m.LvlCrit = Add-CheckBox -Parent $row1 -Text 'Krytyczny' -Checked $true
    $m.LvlErr = Add-CheckBox -Parent $row1 -Text 'Błąd' -Checked $true
    $m.LvlWarn = Add-CheckBox -Parent $row1 -Text 'Ostrzeżenie' -Checked $true
    $m.LvlInfo = Add-CheckBox -Parent $row1 -Text 'Informacja'
    $row2 = Add-ToolbarRow -Module $m -Title 'Zakres'
    Add-Label -Parent $row2 -Text 'Ostatnie godziny' | Out-Null
    $m.Hours = Add-Numeric -Parent $row2 -Value 24 -Minimum 1 -Maximum 2160 -Width 70
    Add-Label -Parent $row2 -Text 'ID zdarzeń' | Out-Null
    $m.Ids = Add-TextBox -Parent $row2 -Width 200 -Placeholder 'opcjonalnie, np. 41, 6008'
    Add-Label -Parent $row2 -Text 'Maks. na komputer' | Out-Null
    $m.Max = Add-Numeric -Parent $row2 -Value 300 -Minimum 10 -Maximum 10000 -Width 70
    Add-Button -Parent $row2 -Text 'Pobierz zdarzenia' -Icon 'E896' -Module $m -Primary -OnClick {
        param($m)
        $logs = @()
        if (Test-Checked $m.LogSystem) { $logs += 'System' }
        if (Test-Checked $m.LogApp) { $logs += 'Application' }
        if (Test-Checked $m.LogSec) { $logs += 'Security' }
        if ($logs.Count -eq 0) { Show-Warning 'Wybierz co najmniej jeden dziennik.'; return }
        $levels = @()
        if (Test-Checked $m.LvlCrit) { $levels += 1 }
        if (Test-Checked $m.LvlErr) { $levels += 2 }
        if (Test-Checked $m.LvlWarn) { $levels += 3 }
        if (Test-Checked $m.LvlInfo) { $levels += 4; $levels += 0 }
        if ($levels.Count -eq 0 -and -not ($logs.Count -eq 1 -and $logs[0] -eq 'Security')) { Show-Warning 'Wybierz co najmniej jeden poziom zdarzeń.'; return }
        $ids = @(Split-ListText ($m.Ids.Text -replace '\s+', ',') | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $params = @{ Logs = $logs; Levels = $levels; Ids = $ids; Hours = (Get-Num $m.Hours); Max = (Get-Num $m.Max) }
        Start-HostOperation -Module $m -Name 'Dziennik zdarzeń' -Targets $targets -Parameters $params -ScriptBlock {
            param($P)
            $filter = @{ LogName = [string[]]@($P.Logs); StartTime = (Get-Date).AddHours( - [int]$P.Hours) }
            $ids = @($P.Ids | Where-Object { $null -ne $_ })
            if ($ids.Count -gt 0) { $filter.Id = [int[]]$ids }
            # Security nie używa poziomów (audyt) - bez filtra poziomu, gdy wybrano tylko Security
            $levels = [int[]]@($P.Levels)
            if ($levels.Count -gt 0 -and (@($P.Logs) -notcontains 'Security' -or @($P.Logs).Count -gt 1)) { $filter.Level = $levels }
            try {
                $events = @(Get-WinEvent -FilterHashtable $filter -MaxEvents ([int]$P.Max) -ErrorAction Stop)
            }
            catch {
                if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { $events = @() } else { throw }
            }
            $events | Sort-Object TimeCreated -Descending | ForEach-Object {
                $msg = $_.Message
                if (-not $msg) { $msg = '(brak opisu – brak biblioteki komunikatów dostawcy)' }
                $tone = switch ([int]$_.Level) { 1 { 'crit' } 2 { 'crit' } 3 { 'warn' } default { 'info' } }
                [pscustomobject]@{
                    'Czas'      = $_.TimeCreated
                    'Poziom'    = $(if ($_.LevelDisplayName) { $_.LevelDisplayName } else { 'Inspekcja' })
                    'ID'        = $_.Id
                    'Źródło'    = $_.ProviderName
                    'Dziennik'  = $_.LogName
                    'Komunikat' = ($msg -replace '\s+', ' ').Trim()
                    '__tone'    = $tone
                }
            }
        }
    } | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'System' -Key 'Tasks' -Title 'Harmonogram zadań' -Icon 'E787' `
    -Description 'Zadania Harmonogramu: podgląd, uruchamianie, włączanie, wyłączanie, usuwanie i tworzenie prostych zadań (konto SYSTEM).' -Build {
    param($m)
    $m.PillColumns = @('Ostatni wynik')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Harmonogram' -Targets $targets -Parameters @{ Filter = $m.Filter.Text.Trim(); HideMicrosoft = (Test-Checked $m.HideMs) } -ScriptBlock {
            param($P)
            $tasks = @(Get-ScheduledTask -ErrorAction Stop)
            if ($P.HideMicrosoft) { $tasks = @($tasks | Where-Object { $_.TaskPath -notlike '\Microsoft\*' }) }
            if ($P.Filter) {
                $f = "*$($P.Filter)*"
                $tasks = @($tasks | Where-Object { $_.TaskName -like $f -or $_.TaskPath -like $f })
            }
            foreach ($t in $tasks) {
                $info = $null
                try { $info = Get-ScheduledTaskInfo -InputObject $t -ErrorAction Stop } catch { }
                $actions = @($t.Actions | ForEach-Object { if ($_.PSObject.Properties['Execute'] -and $_.Execute) { ('{0} {1}' -f $_.Execute, $_.Arguments).Trim() } }) -join ' ; '
                $lastResult = ''
                $tone = ''
                $lastRun = $null
                $nextRun = $null
                if ($info) {
                    $lastRun = $info.LastRunTime
                    $nextRun = $info.NextRunTime
                    $code = [int64]$info.LastTaskResult
                    $lastResult = switch ($code) {
                        0 { 'Sukces' }
                        267009 { 'W trakcie' }
                        267011 { 'Jeszcze nie uruchomiono' }
                        267014 { 'Zatrzymane przez użytkownika' }
                        default { '0x{0:X8}' -f $code }
                    }
                    $tone = switch ($code) { 0 { 'ok' } 267009 { 'info' } 267011 { '' } default { 'crit' } }
                }
                [pscustomobject]@{
                    'Ścieżka'               = $t.TaskPath
                    'Nazwa'                 = $t.TaskName
                    'Stan'                  = [string]$t.State
                    'Ostatnie uruchomienie' = $lastRun
                    'Ostatni wynik'         = $lastResult
                    'Następne uruchomienie' = $nextRun
                    'Konto'                 = $t.Principal.UserId
                    'Akcja'                 = $actions
                    'Autor'                 = $t.Author
                    '__tone'                = $tone
                    '__flag'                = $(if ([string]$t.State -eq 'Disabled') { 'muted' } else { '' })
                }
            }
        }
    }
    $m.Actions.Change = {
        param($m, [string]$Op, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Ścieżka', 'Nazwa') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli zadania, których dotyczy operacja.'; return }
        $label = @{ Run = 'Uruchomić'; Stop = 'Zatrzymać'; Enable = 'Włączyć'; Disable = 'Wyłączyć'; Delete = 'Usunąć' }[$Op]
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) $i['Ścieżka'] + $i['Nazwa'] }
        if (-not (Confirm-Action -Text "$label wybrane zadania?" -Items $items -ConfirmText $(if ($Op -eq 'Delete') { 'Usuń zadania' } else { 'Wykonaj' }) -Danger:($Op -eq 'Delete'))) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) {
            $per[$h] = @{ Op = $Op; Tasks = @($byHost[$h] | ForEach-Object { @{ Path = [string]$_['Ścieżka']; Name = [string]$_['Nazwa'] } }) }
        }
        Start-HostOperation -Module $m -Name "Harmonogram – $Op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            foreach ($t in $P.Tasks) {
                try {
                    $tp = @{ TaskPath = $t.Path; TaskName = $t.Name; ErrorAction = 'Stop' }
                    switch ($P.Op) {
                        'Run' { Start-ScheduledTask @tp }
                        'Stop' { Stop-ScheduledTask @tp }
                        'Enable' { Enable-ScheduledTask @tp | Out-Null }
                        'Disable' { Disable-ScheduledTask @tp | Out-Null }
                        'Delete' { Unregister-ScheduledTask @tp -Confirm:$false }
                    }
                    [pscustomobject]@{ 'Zadanie' = $t.Path + $t.Name; 'Operacja' = $P.Op; 'Wynik' = 'OK' }
                }
                catch { [pscustomobject]@{ 'Zadanie' = $t.Path + $t.Name; 'Operacja' = $P.Op; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Lista'
    Add-Label -Parent $row -Text 'Filtr' | Out-Null
    $m.Filter = Add-TextBox -Parent $row -Width 170 -Placeholder 'nazwa lub ścieżka'
    $m.HideMs = Add-CheckBox -Parent $row -Text 'Ukryj zadania systemowe (\Microsoft\)' -Checked $true
    Add-Button -Parent $row -Text 'Pokaż zadania' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Zaznaczone zadania'
    foreach ($a in @(@('Uruchom', 'Run', 'E768'), @('Zatrzymaj', 'Stop', 'E71A'), @('Włącz', 'Enable', 'E73E'), @('Wyłącz', 'Disable', 'E8D8'))) {
        $b = Add-Button -Parent $row2 -Text $a[0] -Icon $a[2] -Module $m -OnClick { param($m, $s) & $m.Actions.Change $m ([string]$s.Tag) $null }
        $b.Tag = $a[1]
    }
    $b = Add-Button -Parent $row2 -Text 'Usuń' -Icon 'E74D' -Module $m -Danger -OnClick { param($m, $s) & $m.Actions.Change $m 'Delete' $null }
    Add-RowAction -Module $m -Text 'Uruchom teraz' -Icon 'E768' -Action { param($m, $rows) & $m.Actions.Change $m 'Run' $rows }
    Add-RowAction -Module $m -Text 'Włącz' -Icon 'E73E' -Action { param($m, $rows) & $m.Actions.Change $m 'Enable' $rows }
    Add-RowAction -Module $m -Text 'Wyłącz' -Icon 'E8D8' -Action { param($m, $rows) & $m.Actions.Change $m 'Disable' $rows }
    Add-RowAction -Module $m -Text 'Usuń zadanie' -Icon 'E74D' -Danger -Separator -Action { param($m, $rows) & $m.Actions.Change $m 'Delete' $rows }

    $row3 = Add-ToolbarRow -Module $m -Title 'Nowe zadanie'
    $m.NewName = Add-TextBox -Parent $row3 -Width 150 -Placeholder 'Nazwa zadania'
    $m.NewExe = Add-TextBox -Parent $row3 -Width 220 -Placeholder 'Program, np. powershell.exe'
    $m.NewArgs = Add-TextBox -Parent $row3 -Width 220 -Placeholder 'Argumenty'
    $m.NewTrigger = Add-ComboBox -Parent $row3 -Items @('Przy logowaniu', 'Przy uruchomieniu', 'Codziennie o', 'Jednorazowo o', 'Tylko na żądanie') -Width 160
    $m.NewTime = Add-TextBox -Parent $row3 -Width 64 -Text '07:00'
    Add-Button -Parent $row3 -Text 'Utwórz' -Icon 'E710' -Module $m -OnClick {
        param($m)
        $name = $m.NewName.Text.Trim()
        $exe = $m.NewExe.Text.Trim()
        if (-not $name -or -not $exe) { Show-Warning 'Podaj nazwę zadania i program do uruchomienia.'; return }
        if ($name -match '[\\/:*?"<>|]') { Show-Warning 'Nazwa zadania zawiera niedozwolone znaki.'; return }
        $trigger = @('Logon', 'Startup', 'Daily', 'Once', 'None')[$m.NewTrigger.SelectedIndex]
        $time = $m.NewTime.Text.Trim()
        if (($trigger -eq 'Daily' -or $trigger -eq 'Once') -and $time -notmatch '^([01]?\d|2[0-3]):[0-5]\d$') { Show-Warning 'Podaj godzinę w formacie GG:MM (np. 07:30).'; return }
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        if (-not (Confirm-Action -Text "Utworzyć zadanie «$name» (konto SYSTEM, najwyższe uprawnienia)?`r`nProgram: $exe $($m.NewArgs.Text.Trim())" -Items $targets -ConfirmText 'Utwórz')) { return }
        $params = @{ Name = $name; Execute = $exe; Arguments = $m.NewArgs.Text.Trim(); Trigger = $trigger; Time = $time }
        Start-HostOperation -Module $m -Name 'Tworzenie zadania' -Targets $targets -Output Log -Parameters $params -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            if ($P.Arguments) { $action = New-ScheduledTaskAction -Execute $P.Execute -Argument $P.Arguments }
            else { $action = New-ScheduledTaskAction -Execute $P.Execute }
            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
            $register = @{ TaskName = $P.Name; Action = $action; Principal = $principal; Settings = $settings; Force = $true; ErrorAction = 'Stop' }
            $at = $null
            if ($P.Time) {
                $at = [datetime]::Today.Add([timespan]::Parse($P.Time))
                if ($P.Trigger -eq 'Once' -and $at -lt (Get-Date)) { $at = $at.AddDays(1) }
            }
            switch ($P.Trigger) {
                'Logon' { $register.Trigger = New-ScheduledTaskTrigger -AtLogOn }
                'Startup' { $register.Trigger = New-ScheduledTaskTrigger -AtStartup }
                'Daily' { $register.Trigger = New-ScheduledTaskTrigger -Daily -At $at }
                'Once' { $register.Trigger = New-ScheduledTaskTrigger -Once -At $at }
            }
            Register-ScheduledTask @register | Out-Null
            [pscustomobject]@{ 'Zadanie' = '\' + $P.Name; 'Wynik' = 'Utworzono' }
        }
    } | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'System' -Key 'Autostart' -Title 'Autostart' -Icon 'E768' -Badge 'Nowość w wersji 4.0' `
    -Description 'Programy uruchamiane automatycznie: klucze Run/RunOnce (komputer i załadowane profile użytkowników) oraz foldery Autostart. Wybrane wpisy można usunąć.' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Autostart' -Targets $targets -ScriptBlock {
            param($P)
            $keys = @(
                @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'; Scope = 'Komputer' }
                @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce'; Scope = 'Komputer (jednorazowo)' }
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run'; Scope = 'Komputer (32-bit)' }
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\RunOnce'; Scope = 'Komputer (32-bit, jednorazowo)' }
            )
            # Załadowane gałęzie użytkowników (zalogowani) - nazwy kont z SID
            foreach ($hive in @(Get-ChildItem -Path 'Registry::HKEY_USERS' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^S-1-5-21-[\d-]+$' })) {
                $sid = $hive.PSChildName
                $user = $sid
                try { $user = (New-Object System.Security.Principal.SecurityIdentifier($sid)).Translate([System.Security.Principal.NTAccount]).Value } catch { }
                $keys += @{ Path = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\Run"; Scope = "Użytkownik $user" }
                $keys += @{ Path = "Registry::HKEY_USERS\$sid\Software\Microsoft\Windows\CurrentVersion\RunOnce"; Scope = "Użytkownik $user (jednorazowo)" }
            }
            foreach ($k in $keys) {
                $item = Get-Item -LiteralPath $k.Path -ErrorAction SilentlyContinue
                if (-not $item) { continue }
                foreach ($name in $item.GetValueNames()) {
                    if (-not $name) { continue }
                    [pscustomobject]@{
                        'Nazwa'    = $name
                        'Polecenie' = [string]$item.GetValue($name)
                        'Zakres'   = $k.Scope
                        'Typ'      = 'Rejestr'
                        'Lokalizacja' = $k.Path -replace '^Registry::', ''
                    }
                }
            }
            $folders = @(@{ Path = Join-Path $env:ProgramData 'Microsoft\Windows\Start Menu\Programs\Startup'; Scope = 'Wszyscy użytkownicy' })
            foreach ($pr in @(Get-CimInstance -ClassName Win32_UserProfile -Filter 'Special = False' -ErrorAction SilentlyContinue)) {
                $folders += @{ Path = Join-Path $pr.LocalPath 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup'; Scope = 'Profil ' + (Split-Path $pr.LocalPath -Leaf) }
            }
            foreach ($f in $folders) {
                if (-not (Test-Path -LiteralPath $f.Path)) { continue }
                foreach ($file in @(Get-ChildItem -LiteralPath $f.Path -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'desktop.ini' })) {
                    $target = ''
                    if ($file.Extension -eq '.lnk') {
                        try { $target = (New-Object -ComObject WScript.Shell).CreateShortcut($file.FullName).TargetPath } catch { }
                    }
                    [pscustomobject]@{
                        'Nazwa'       = $file.Name
                        'Polecenie'   = $(if ($target) { $target } else { $file.FullName })
                        'Zakres'      = $f.Scope
                        'Typ'         = 'Folder Autostart'
                        'Lokalizacja' = $file.FullName
                    }
                }
            }
        }
    }
    $m.Actions.Remove = {
        param($m, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Nazwa', 'Typ', 'Lokalizacja') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli wpisy autostartu do usunięcia.'; return }
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) '{0} ({1})' -f $i['Nazwa'], $i['Lokalizacja'] }
        if (-not (Confirm-Action -Text 'Usunąć wybrane wpisy autostartu? Programy nie będą uruchamiane automatycznie (same programy nie zostaną odinstalowane).' -Items $items -ConfirmText 'Usuń wpisy' -Danger)) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Items = @($byHost[$h] | ForEach-Object { @{ Name = [string]$_['Nazwa']; Type = [string]$_['Typ']; Location = [string]$_['Lokalizacja'] } }) } }
        Start-HostOperation -Module $m -Name 'Usuwanie z autostartu' -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            foreach ($i in $P.Items) {
                try {
                    if ($i.Type -eq 'Rejestr') {
                        $path = if ($i.Location -like 'HKEY_USERS\*') { 'Registry::' + $i.Location } else { $i.Location }
                        Remove-ItemProperty -LiteralPath $path -Name $i.Name -ErrorAction Stop
                    }
                    else { Remove-Item -LiteralPath $i.Location -Force -ErrorAction Stop }
                    [pscustomobject]@{ 'Wpis' = $i.Name; 'Wynik' = 'Usunięto' }
                }
                catch { [pscustomobject]@{ 'Wpis' = $i.Name; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row -Text 'Pokaż autostart' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Usuń zaznaczone wpisy' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Remove $m $null } | Out-Null
    Add-Label -Parent $row -Text 'Wpisy użytkowników są widoczne tylko dla zalogowanych (załadowane profile).' -Hint | Out-Null
    Add-RowAction -Module $m -Text 'Usuń z autostartu' -Icon 'E74D' -Danger -Action { param($m, $rows) & $m.Actions.Remove $m $rows }
}

Register-Module -Workspace 'Remote' -Category 'System' -Key 'Printers' -Title 'Drukarki' -Icon 'E749' -Badge 'Nowość w wersji 4.0' `
    -Description 'Drukarki zainstalowane na komputerze (sterownik, port, udostępnianie, zadania w kolejce), czyszczenie kolejek, usuwanie drukarek i restart bufora wydruku.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Stan')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Drukarki' -Targets $targets -ScriptBlock {
            param($P)
            $spooler = Get-Service -Name Spooler -ErrorAction SilentlyContinue
            $printers = @(Get-CimInstance -ClassName Win32_Printer)
            $jobs = @(Get-CimInstance -ClassName Win32_PrintJob -ErrorAction SilentlyContinue)
            foreach ($p in $printers) {
                $count = @($jobs | Where-Object { ($_.Name -split ',')[0] -eq $p.Name }).Count
                $status = switch ([int]$p.PrinterStatus) { 1 { 'Inny' } 2 { 'Nieznany' } 3 { 'Bezczynna' } 4 { 'Drukuje' } 5 { 'Rozgrzewanie' } 6 { 'Zatrzymana' } 7 { 'Offline' } default { [string]$p.PrinterStatus } }
                if ($p.WorkOffline) { $status = 'Offline' }
                $tone = switch ($status) { 'Bezczynna' { 'ok' } 'Drukuje' { 'info' } 'Offline' { 'crit' } 'Zatrzymana' { 'crit' } default { 'warn' } }
                [pscustomobject]@{
                    'Drukarka'     = $p.Name
                    'Stan'         = $status
                    'Zadania'      = $count
                    'Domyślna'     = [bool]$p.Default
                    'Sterownik'    = $p.DriverName
                    'Port'         = $p.PortName
                    'Udostępniona' = [bool]$p.Shared
                    'Nazwa udziału' = $p.ShareName
                    'Lokalizacja'  = $p.Location
                    'Bufor wydruku' = $(if ($spooler) { [string]$spooler.Status } else { '?' })
                    '__tone'       = $tone
                }
            }
        }
    }
    $m.Actions.Change = {
        param($m, [string]$Op, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Drukarka') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli drukarki.'; return }
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) $i['Drukarka'] }
        $text = if ($Op -eq 'Remove') { 'Usunąć wybrane drukarki z komputerów?' } else { 'Usunąć wszystkie zadania z kolejek wybranych drukarek?' }
        if (-not (Confirm-Action -Text $text -Items $items -ConfirmText $(if ($Op -eq 'Remove') { 'Usuń drukarki' } else { 'Wyczyść kolejki' }) -Danger)) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Op = $Op; Names = @($byHost[$h] | ForEach-Object { [string]$_['Drukarka'] }) } }
        Start-HostOperation -Module $m -Name "Drukarki – $Op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            foreach ($name in $P.Names) {
                try {
                    $printer = Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Name -eq $name }
                    if (-not $printer) { throw 'Nie znaleziono drukarki.' }
                    if ($P.Op -eq 'Remove') {
                        Remove-CimInstance -InputObject $printer -ErrorAction Stop
                        [pscustomobject]@{ 'Drukarka' = $name; 'Wynik' = 'Usunięto' }
                    }
                    else {
                        $r = Invoke-CimMethod -InputObject $printer -MethodName CancelAllJobs
                        [pscustomobject]@{ 'Drukarka' = $name; 'Wynik' = $(if ($r.ReturnValue -eq 0) { 'Wyczyszczono kolejkę' } else { "Błąd – kod $($r.ReturnValue)" }) }
                    }
                }
                catch { [pscustomobject]@{ 'Drukarka' = $name; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row -Text 'Pokaż drukarki' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Wyczyść kolejki zaznaczonych' -Icon 'E894' -Module $m -OnClick { param($m) & $m.Actions.Change $m 'Clear' $null } | Out-Null
    Add-Button -Parent $row -Text 'Usuń zaznaczone' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Change $m 'Remove' $null } | Out-Null
    Add-Button -Parent $row -Text 'Restart bufora wydruku' -Icon 'E777' -Module $m -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        if (-not (Confirm-Action -Text 'Uruchomić ponownie usługę bufora wydruku (Spooler)? Zadania w trakcie drukowania mogą zostać przerwane.' -Items $targets -ConfirmText 'Uruchom ponownie')) { return }
        Start-HostOperation -Module $m -Name 'Restart bufora wydruku' -Targets $targets -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            Restart-Service -Name Spooler -Force -ErrorAction Stop
            'Bufor wydruku: ' + [string](Get-Service -Name Spooler).Status
        }
    } | Out-Null
    Add-RowAction -Module $m -Text 'Wyczyść kolejkę' -Icon 'E894' -Action { param($m, $rows) & $m.Actions.Change $m 'Clear' $rows }
    Add-RowAction -Module $m -Text 'Usuń drukarkę' -Icon 'E74D' -Danger -Action { param($m, $rows) & $m.Actions.Change $m 'Remove' $rows }
}

Register-Module -Workspace 'Remote' -Category 'System' -Key 'Drivers' -Title 'Sterowniki i urządzenia' -Icon 'E772' `
    -Description 'Zainstalowane sterowniki (Win32_PnPSignedDriver) oraz urządzenia zgłaszające problem w Menedżerze urządzeń.' -Build {
    param($m)
    $m.ColorBools = $true
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Label -Parent $row -Text 'Filtr' | Out-Null
    $m.Filter = Add-TextBox -Parent $row -Width 220 -Placeholder 'urządzenie, producent lub klasa'
    Add-Button -Parent $row -Text 'Pokaż sterowniki' -Icon 'E72C' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Sterowniki' -Targets $targets -Parameters @{ Filter = $m.Filter.Text.Trim() } -ScriptBlock {
            param($P)
            $drivers = @(Get-CimInstance -ClassName Win32_PnPSignedDriver | Where-Object { $_.DeviceName })
            if ($P.Filter) {
                $f = "*$($P.Filter)*"
                $drivers = @($drivers | Where-Object { $_.DeviceName -like $f -or $_.Manufacturer -like $f -or $_.DriverProviderName -like $f -or $_.DeviceClass -like $f })
            }
            $drivers | Sort-Object DeviceClass, DeviceName | ForEach-Object {
                [pscustomobject]@{
                    'Urządzenie' = $_.DeviceName
                    'Klasa'      = $_.DeviceClass
                    'Wersja'     = $_.DriverVersion
                    'Data'       = $_.DriverDate
                    'Producent'  = $_.Manufacturer
                    'Dostawca'   = $_.DriverProviderName
                    'Plik INF'   = $_.InfName
                    'Podpisany'  = $_.IsSigned
                }
            }
        }
    } | Out-Null
    Add-Button -Parent $row -Text 'Urządzenia z problemami' -Icon 'E7BA' -Module $m -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Urządzenia z problemami' -Targets $targets -ScriptBlock {
            param($P)
            $devices = @(Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.ConfigManagerErrorCode -ne 0 })
            if ($devices.Count -eq 0) { return [pscustomobject]@{ 'Urządzenie' = '(brak urządzeń z problemami)' } }
            foreach ($d in $devices) {
                [pscustomobject]@{
                    'Urządzenie'    = $d.Name
                    'Kod błędu'     = $d.ConfigManagerErrorCode
                    'Stan'          = $d.Status
                    'Klasa'         = $d.PNPClass
                    'Producent'     = $d.Manufacturer
                    'ID urządzenia' = $d.DeviceID
                    '__flag'        = 'warn'
                }
            }
        }
    } | Out-Null
}
#endregion

#region Zarządzanie zdalne: Oprogramowanie
# Skrypt instalacji aktualizacji uruchamiany na komputerze jako zadanie SYSTEM - API Windows Update
# (pobieranie/instalacja) nie działa bezpośrednio w sesji WinRM.
$script:WuJobScript = @'
param([switch]$AutoReboot, [switch]$IncludeDrivers)
$log = Join-Path $PSScriptRoot 'WU.log'
function Write-WuLog([string]$Text) {
    Add-Content -LiteralPath $log -Value ('[{0:yyyy-MM-dd HH:mm:ss}] {1}' -f (Get-Date), $Text) -Encoding UTF8
}
Write-WuLog '==== START ===='
try {
    $session = New-Object -ComObject Microsoft.Update.Session
    $session.ClientApplicationID = 'DomainOps'
    $criteria = "IsInstalled=0 and IsHidden=0"
    if (-not $IncludeDrivers) { $criteria += " and Type='Software'" }
    Write-WuLog "Wyszukiwanie aktualizacji ($criteria)…"
    $search = $session.CreateUpdateSearcher().Search($criteria)
    if ($search.Updates.Count -eq 0) {
        Write-WuLog 'Brak aktualizacji do zainstalowania.'
    }
    else {
        $toInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($u in $search.Updates) {
            if (-not $u.EulaAccepted) { $u.AcceptEula() }
            [void]$toInstall.Add($u)
            Write-WuLog "Do instalacji: $($u.Title)"
        }
        Write-WuLog "Pobieranie ($($toInstall.Count))…"
        $downloader = $session.CreateUpdateDownloader()
        $downloader.Updates = $toInstall
        $download = $downloader.Download()
        Write-WuLog "Pobieranie zakończone (ResultCode=$($download.ResultCode))."
        $ready = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($u in $toInstall) { if ($u.IsDownloaded) { [void]$ready.Add($u) } }
        if ($ready.Count -eq 0) {
            Write-WuLog 'Żadna aktualizacja nie została pobrana.'
        }
        else {
            Write-WuLog "Instalacja ($($ready.Count))…"
            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $ready
            $result = $installer.Install()
            for ($i = 0; $i -lt $ready.Count; $i++) {
                Write-WuLog ('{0} -> ResultCode={1}' -f $ready.Item($i).Title, $result.GetUpdateResult($i).ResultCode)
            }
            Write-WuLog "Instalacja zakończona (ResultCode=$($result.ResultCode), RebootRequired=$($result.RebootRequired))."
            if ($result.RebootRequired -and $AutoReboot) {
                Write-WuLog 'Restart za 5 minut.'
                & (Join-Path $env:SystemRoot 'System32\shutdown.exe') /r /t 300 /d p:2:17 /c 'Restart po instalacji aktualizacji (Domain Ops).'
            }
        }
    }
}
catch {
    Write-WuLog "BŁĄD: $($_.Exception.Message)"
}
Write-WuLog '==== KONIEC ===='
'@

Register-Module -Workspace 'Remote' -Category 'Oprogramowanie' -Key 'Programs' -Title 'Zainstalowane programy' -Icon 'E71D' `
    -Description 'Programy z rejestru (64/32-bit i profile zalogowanych użytkowników) oraz ciche odinstalowanie zaznaczonych (MSI albo QuietUninstallString).' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Programy' -Targets $targets -Parameters @{ Filter = $m.Filter.Text.Trim(); ShowSystem = (Test-Checked $m.ShowSystem) } -ScriptBlock {
            param($P)
            $roots = @(
                @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'Komputer' },
                @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'Komputer (32-bit)' },
                @{ Path = 'Registry::HKEY_USERS\*\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'; Scope = 'Użytkownik' }
            )
            foreach ($root in $roots) {
                Get-ItemProperty -Path $root.Path -ErrorAction SilentlyContinue | ForEach-Object {
                    if (-not $_.DisplayName) { return }
                    if (-not $P.ShowSystem -and ($_.SystemComponent -eq 1 -or $_.ParentKeyName -or @('Update', 'Hotfix', 'Security Update') -contains $_.ReleaseType)) { return }
                    if ($P.Filter -and $_.DisplayName -notlike "*$($P.Filter)*" -and $_.Publisher -notlike "*$($P.Filter)*") { return }
                    $date = $null
                    if ($_.InstallDate -match '^\d{8}$') { try { $date = [datetime]::ParseExact($_.InstallDate, 'yyyyMMdd', $null) } catch { } }
                    $isMsi = ($_.WindowsInstaller -eq 1 -and $_.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$')
                    $uninstall = if ($_.QuietUninstallString) { $_.QuietUninstallString } else { $_.UninstallString }
                    [pscustomobject]@{
                        'Nazwa'           = $_.DisplayName
                        'Wersja'          = $_.DisplayVersion
                        'Wydawca'         = $_.Publisher
                        'Data instalacji' = $date
                        'Rozmiar (MB)'    = $(if ($_.EstimatedSize) { [Math]::Round($_.EstimatedSize / 1024, 1) } else { $null })
                        'Zakres'          = $root.Scope
                        'Typ'             = $(if ($isMsi) { 'MSI' } else { 'Inny' })
                        'Identyfikator'   = $_.PSChildName
                        'Odinstalowanie'  = $uninstall
                    }
                }
            }
        }
    }
    $m.Actions.Uninstall = {
        param($m, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Identyfikator', 'Nazwa', 'Zakres') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli programy do odinstalowania.'; return }
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) $i['Nazwa'] }
        if (-not (Confirm-Action -Text 'Odinstalować wybrane programy? Operacja jest wykonywana bez interakcji z użytkownikiem i bez restartu.' -Items $items -ConfirmText 'Odinstaluj' -Danger)) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) {
            $per[$h] = @{ Items = @($byHost[$h] | ForEach-Object { @{ Id = [string]$_['Identyfikator']; Name = [string]$_['Nazwa']; Scope = [string]$_['Zakres'] } }) }
        }
        Start-HostOperation -Module $m -Name 'Odinstalowanie' -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            foreach ($item in $P.Items) {
                try {
                    if ($item.Scope -eq 'Użytkownik') { throw 'Programy instalowane w profilu użytkownika trzeba odinstalować w jego sesji.' }
                    $key = $null
                    foreach ($root in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
                        $candidate = Join-Path $root $item.Id
                        if (Test-Path -LiteralPath $candidate) { $key = Get-ItemProperty -LiteralPath $candidate; break }
                    }
                    if (-not $key) { throw 'Nie znaleziono wpisu w rejestrze.' }
                    if ($key.WindowsInstaller -eq 1 -and $item.Id -match '^\{[0-9A-Fa-f-]{36}\}$') {
                        $exe = Join-Path $env:SystemRoot 'System32\msiexec.exe'
                        $arguments = '/x {0} /qn /norestart' -f $item.Id
                    }
                    elseif ($key.QuietUninstallString) {
                        $exe = Join-Path $env:SystemRoot 'System32\cmd.exe'
                        $arguments = '/d /s /c "' + $key.QuietUninstallString + '"'
                    }
                    else { throw 'Brak cichego odinstalowania (QuietUninstallString) – odinstaluj program ręcznie.' }
                    $proc = Start-Process -FilePath $exe -ArgumentList $arguments -PassThru -WindowStyle Hidden
                    $null = $proc.Handle
                    $proc.WaitForExit()
                    $code = $proc.ExitCode
                    $status = switch ($code) { 0 { 'Odinstalowano' } 3010 { 'Odinstalowano – wymagany restart' } 1605 { 'Program nie jest zainstalowany' } default { "Błąd – kod wyjścia $code" } }
                    [pscustomobject]@{ 'Program' = $item.Name; 'Kod wyjścia' = $code; 'Wynik' = $status }
                }
                catch { [pscustomobject]@{ 'Program' = $item.Name; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Lista'
    Add-Label -Parent $row -Text 'Nazwa lub wydawca' | Out-Null
    $m.Filter = Add-TextBox -Parent $row -Width 200 -Placeholder 'np. Office'
    $m.ShowSystem = Add-CheckBox -Parent $row -Text 'Pokaż aktualizacje i składniki systemowe'
    Add-Button -Parent $row -Text 'Pokaż programy' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Odinstaluj zaznaczone' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Uninstall $m $null } | Out-Null
    Add-RowAction -Module $m -Text 'Odinstaluj' -Icon 'E74D' -Danger -Action { param($m, $rows) & $m.Actions.Uninstall $m $rows }
    Add-RowAction -Module $m -Text 'Pokaż ten program na wszystkich komputerach' -Icon 'E721' -Action {
        param($m, $rows)
        $name = [string](Get-ObjectValue $rows[0] 'Nazwa')
        if ($name) { $m.FilterBox.Text = $name }
    }
}

Register-Module -Workspace 'Remote' -Category 'Oprogramowanie' -Key 'WindowsUpdate' -Title 'Windows Update' -Icon 'E777' `
    -Description 'Dostępne aktualizacje, historia, sprawdzanie konkretnych poprawek KB oraz instalacja (zadanie SYSTEM na komputerze, log w %SystemRoot%\Temp\DomainOps\WU.log). Nie wymaga modułu PSWindowsUpdate.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Status', 'Stan')
    $row = Add-ToolbarRow -Module $m -Title 'Sprawdzenie'
    $m.Drivers = Add-CheckBox -Parent $row -Text 'Uwzględnij sterowniki'
    Add-Button -Parent $row -Text 'Wyszukaj dostępne' -Icon 'E721' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Wyszukiwanie aktualizacji' -Targets $targets -Parameters @{ IncludeDrivers = (Test-Checked $m.Drivers) } -ScriptBlock {
            param($P)
            $session = New-Object -ComObject Microsoft.Update.Session
            $criteria = "IsInstalled=0 and IsHidden=0"
            if (-not $P.IncludeDrivers) { $criteria += " and Type='Software'" }
            $result = $session.CreateUpdateSearcher().Search($criteria)
            if ($result.Updates.Count -eq 0) { return [pscustomobject]@{ 'Aktualizacja' = '(brak dostępnych aktualizacji)'; 'Stan' = 'Aktualny'; '__tone' = 'ok' } }
            foreach ($u in $result.Updates) {
                [pscustomobject]@{
                    'Aktualizacja'    = $u.Title
                    'Stan'            = 'Do instalacji'
                    'KB'              = (@($u.KBArticleIDs | ForEach-Object { "KB$_" }) -join ', ')
                    'Kategoria'       = (@($u.Categories | ForEach-Object { $_.Name }) -join ', ')
                    'Ważność'         = $u.MsrcSeverity
                    'Rozmiar (MB)'    = [Math]::Round([double]$u.MaxDownloadSize / 1MB, 1)
                    'Pobrana'         = [bool]$u.IsDownloaded
                    'Wymaga restartu' = ($u.InstallationBehavior.RebootBehavior -ne 0)
                    '__tone'          = $(if ($u.MsrcSeverity -eq 'Critical') { 'crit' } else { 'warn' })
                }
            }
        }
    } | Out-Null
    Add-Button -Parent $row -Text 'Historia (ostatnie 50)' -Icon 'E81C' -Module $m -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Historia aktualizacji' -Targets $targets -ScriptBlock {
            param($P)
            $searcher = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher()
            $count = $searcher.GetTotalHistoryCount()
            if ($count -eq 0) { return [pscustomobject]@{ 'Aktualizacja' = '(historia jest pusta)' } }
            foreach ($h in $searcher.QueryHistory(0, [Math]::Min($count, 50))) {
                if (-not $h.Title) { continue }
                $status = switch ([int]$h.ResultCode) { 1 { 'W toku' } 2 { 'Sukces' } 3 { 'Sukces z błędami' } 4 { 'Błąd' } 5 { 'Przerwano' } default { 'Nieznany' } }
                $operation = switch ([int]$h.Operation) { 1 { 'Instalacja' } 2 { 'Odinstalowanie' } default { '' } }
                [pscustomobject]@{
                    'Data'         = $h.Date.ToLocalTime()
                    'Aktualizacja' = $h.Title
                    'Operacja'     = $operation
                    'Status'       = $status
                    'Kod HRESULT'  = $(if ($h.HResult) { '0x{0:X8}' -f $h.HResult } else { '' })
                    '__tone'       = $(switch ([int]$h.ResultCode) { 2 { 'ok' } 3 { 'warn' } 4 { 'crit' } 5 { 'warn' } default { 'info' } })
                }
            }
        }
    } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Poprawki KB'
    $m.Kb = Add-TextBox -Parent $row2 -Width 300 -Placeholder 'np. KB5034441, KB5005565'
    Add-Button -Parent $row2 -Text 'Sprawdź obecność' -Icon 'E73E' -Module $m -OnClick {
        param($m)
        $kbs = @(Split-ListText ($m.Kb.Text -replace '\s+', ',') | ForEach-Object { $_.ToUpperInvariant() -replace '^(KB)?(\d+)$', 'KB$2' } | Where-Object { $_ -match '^KB\d+$' } | Select-Object -Unique)
        if ($kbs.Count -eq 0) { Show-Warning 'Podaj numery poprawek, np. KB5034441.'; return }
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Sprawdzanie KB' -Targets $targets -Parameters @{ Kbs = $kbs } -ScriptBlock {
            param($P)
            $hotfixes = @{}
            foreach ($h in @(Get-CimInstance -ClassName Win32_QuickFixEngineering -ErrorAction SilentlyContinue)) { $hotfixes[[string]$h.HotFixID] = $h }
            $history = @()
            try {
                $searcher = (New-Object -ComObject Microsoft.Update.Session).CreateUpdateSearcher()
                $count = $searcher.GetTotalHistoryCount()
                if ($count -gt 0) { $history = @($searcher.QueryHistory(0, [Math]::Min($count, 500)) | Where-Object { $_.Title -and $_.ResultCode -eq 2 -and $_.Operation -eq 1 }) }
            }
            catch { }
            $os = Get-CimInstance -ClassName Win32_OperatingSystem
            foreach ($kb in $P.Kbs) {
                $found = $hotfixes[$kb]
                $wu = @($history | Where-Object { $_.Title -like "*$kb*" } | Select-Object -First 1)
                $installed = ($null -ne $found -or $wu.Count -gt 0)
                $date = $null
                if ($found -and $found.InstalledOn) { $date = $found.InstalledOn } elseif ($wu.Count -gt 0) { $date = $wu[0].Date.ToLocalTime() }
                [pscustomobject]@{
                    'KB'              = $kb
                    'Stan'            = $(if ($installed) { 'Zainstalowana' } else { 'Brak' })
                    'Data instalacji' = $date
                    'Źródło'          = $(if ($found) { 'Win32_QuickFixEngineering' } elseif ($wu.Count -gt 0) { 'Historia Windows Update' } else { '' })
                    'Opis'            = $(if ($found) { $found.Description } elseif ($wu.Count -gt 0) { $wu[0].Title } else { '' })
                    'System'          = '{0} ({1})' -f $os.Caption, $os.BuildNumber
                    '__tone'          = $(if ($installed) { 'ok' } else { 'crit' })
                }
            }
        }
    } | Out-Null
    $row3 = Add-ToolbarRow -Module $m -Title 'Instalacja'
    $m.AutoReboot = Add-CheckBox -Parent $row3 -Text 'Automatyczny restart po instalacji (za 5 min), jeśli wymagany'
    Add-Button -Parent $row3 -Text 'Zainstaluj aktualizacje' -Icon 'E896' -Module $m -Danger -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $reboot = if (Test-Checked $m.AutoReboot) { 'z automatycznym restartem' } else { 'bez restartu' }
        if (-not (Confirm-Action -Text "Zainstalować wszystkie dostępne aktualizacje ($reboot) na $($targets.Count) komputer(ach)?" -Items $targets -ConfirmText 'Zainstaluj')) { return }
        $params = @{ Script = $script:WuJobScript; AutoReboot = (Test-Checked $m.AutoReboot); IncludeDrivers = (Test-Checked $m.Drivers) }
        Start-HostOperation -Module $m -Name 'Instalacja aktualizacji' -Targets $targets -Output Log -Parameters $params -ScriptBlock {
            param($P)
            $dir = Join-Path $env:SystemRoot 'Temp\DomainOps'
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            $file = Join-Path $dir 'Invoke-DomainOpsWU.ps1'
            Set-Content -LiteralPath $file -Value $P.Script -Encoding UTF8
            $taskName = 'DomainOps-WindowsUpdate'
            $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($existing -and [string]$existing.State -eq 'Running') { return 'Instalacja już trwa (zadanie DomainOps-WindowsUpdate).' }
            $argLine = '-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $file
            if ($P.AutoReboot) { $argLine += ' -AutoReboot' }
            if ($P.IncludeDrivers) { $argLine += ' -IncludeDrivers' }
            $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argLine
            $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Hours 4)
            Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings -Force | Out-Null
            Start-ScheduledTask -TaskName $taskName
            'Zlecono instalację (zadanie SYSTEM). Postęp: przycisk «Stan instalacji».'
        }
    } | Out-Null
    Add-Button -Parent $row3 -Text 'Stan instalacji' -Icon 'E9D9' -Module $m -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Stan instalacji aktualizacji' -Targets $targets -ScriptBlock {
            param($P)
            $log = Join-Path $env:SystemRoot 'Temp\DomainOps\WU.log'
            $task = Get-ScheduledTask -TaskName 'DomainOps-WindowsUpdate' -ErrorAction SilentlyContinue
            $lines = @()
            if (Test-Path -LiteralPath $log) {
                $all = @(Get-Content -LiteralPath $log -Encoding UTF8)
                $start = 0
                for ($i = $all.Count - 1; $i -ge 0; $i--) { if ($all[$i] -like '*==== START ====*') { $start = $i; break } }
                $lines = @($all[$start..($all.Count - 1)])
            }
            $rebootRequired = $false
            try { $rebootRequired = [bool](New-Object -ComObject Microsoft.Update.SystemInfo).RebootRequired } catch { }
            $state = if ($task) { [string]$task.State } else { '(brak – instalacji nie zlecano)' }
            [pscustomobject]@{
                'Stan'            = $(if ($state -eq 'Running') { 'Trwa instalacja' } elseif ($lines.Count -and $lines[-1] -like '*KONIEC*') { 'Zakończono' } else { $state })
                'Ostatni wpis'    = $(if ($lines.Count) { $lines[-1] } else { '' })
                'Wymaga restartu' = $rebootRequired
                'Log'             = ($lines -join "`r`n")
                '__tone'          = $(if ($state -eq 'Running') { 'info' } elseif ($rebootRequired) { 'warn' } else { 'ok' })
            }
        }
    } | Out-Null
    $m.GoodWhenNo = @('Wymaga restartu')
}
#endregion

#region Zarządzanie zdalne: Bezpieczeństwo
Register-Module -Workspace 'Remote' -Category 'Bezpieczeństwo' -Key 'SecurityAudit' -Title 'Szybki audyt bezpieczeństwa' -Icon 'EA18' -Badge 'Nowość w wersji 4.0' `
    -Description 'Kilkanaście kontroli w jednym przebiegu: zapora, antywirus i sygnatury, BitLocker, SMBv1, NLA, UAC, LAPS, aktualizacje, Secure Boot, TPM, WDigest, ochrona LSA, LLMNR i konto Gość. Wynik procentowy dla każdego komputera.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Wynik')
    $row = Add-ToolbarRow -Module $m -Title 'Parametry'
    Add-Label -Parent $row -Text 'Aktualizacje nie starsze niż (dni)' | Out-Null
    $m.UpdateDays = Add-Numeric -Parent $row -Value 45 -Minimum 7 -Maximum 365 -Width 60
    Add-Label -Parent $row -Text 'Sygnatury AV (dni)' | Out-Null
    $m.SigDays = Add-Numeric -Parent $row -Value 3 -Minimum 1 -Maximum 60 -Width 60
    Add-Button -Parent $row -Text 'Uruchom audyt' -Icon 'EA18' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Reset-StatTiles $m
        Start-HostOperation -Module $m -Name 'Audyt bezpieczeństwa' -Targets $targets -Parameters @{ UpdateDays = (Get-Num $m.UpdateDays); SigDays = (Get-Num $m.SigDays) } -ScriptBlock {
            param($P)
            $r = [ordered]@{}
            $problems = New-Object System.Collections.ArrayList
            $check = {
                param([string]$Name, $Value, [string]$Problem)
                $r[$Name] = $Value
                if ($Value -eq $false -and $Problem) { [void]$problems.Add($Problem) }
            }
            $isServer = ((Get-CimInstance -ClassName Win32_OperatingSystem).ProductType -ne 1)
            # Zapora
            $fw = $null
            try { $fw = (@(Get-NetFirewallProfile -ErrorAction Stop | Where-Object { [string]$_.Enabled -ne 'True' }).Count -eq 0) } catch { }
            & $check 'Zapora włączona' $fw 'zapora wyłączona w co najmniej jednym profilu'
            # Antywirus
            $av = $null
            $sigOk = $null
            try {
                $mp = Get-MpComputerStatus -ErrorAction Stop
                $av = [bool]($mp.AntivirusEnabled -and $mp.RealTimeProtectionEnabled)
                $sigOk = ([int]$mp.AntivirusSignatureAge -le [int]$P.SigDays)
            }
            catch {
                if (-not $isServer) {
                    try {
                        $products = @(Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction Stop)
                        # productState: bajt 2 = 0x10 -> ochrona włączona, bajt 3 = 0x00 -> sygnatury aktualne
                        $active = @($products | Where-Object { (([int]$_.productState -shr 8) -band 0xFF) -band 0x10 })
                        $av = ($active.Count -gt 0)
                        if ($av) { $sigOk = (@($active | Where-Object { ([int]$_.productState -band 0xFF) -eq 0 }).Count -gt 0) }
                    }
                    catch { }
                }
            }
            & $check 'Antywirus aktywny' $av 'antywirus wyłączony lub nieobecny'
            & $check 'Sygnatury aktualne' $sigOk 'nieaktualne sygnatury antywirusa'
            # BitLocker na dysku systemowym
            $bl = $null
            try {
                $vol = Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftVolumeEncryption' -ClassName Win32_EncryptableVolume -Filter ("DriveLetter='{0}'" -f $env:SystemDrive) -ErrorAction Stop
                if ($vol) { $bl = ([int]$vol.ProtectionStatus -eq 1) }
            }
            catch { }
            & $check 'BitLocker na systemowym' $bl 'dysk systemowy bez ochrony BitLocker'
            # SMBv1
            $smb1 = $null
            try { $smb1 = -not [bool](Get-SmbServerConfiguration -ErrorAction Stop).EnableSMB1Protocol } catch { }
            & $check 'SMBv1 wyłączony' $smb1 'włączony protokół SMBv1'
            # RDP + NLA
            $ts = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' -ErrorAction SilentlyContinue
            $rdpOn = ($ts -and [int]$ts.fDenyTSConnections -eq 0)
            $nla = $null
            if ($rdpOn) { $nla = ([int](Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -ErrorAction SilentlyContinue).UserAuthentication -eq 1) }
            $r['RDP włączony'] = $rdpOn
            & $check 'NLA dla RDP' $nla 'pulpit zdalny bez NLA'
            # UAC
            $lua = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -ErrorAction SilentlyContinue).EnableLUA
            & $check 'UAC włączony' ($null -eq $lua -or [int]$lua -eq 1) 'kontrola konta użytkownika (UAC) wyłączona'
            # LAPS (Windows LAPS albo LAPS legacy)
            $laps = $false
            if (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\LAPS' -ErrorAction SilentlyContinue) { $laps = $true }
            if (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Policies\LAPS' -ErrorAction SilentlyContinue) { $laps = $true }
            $legacy = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft Services\AdmPwd' -ErrorAction SilentlyContinue
            if ($legacy -and [int]$legacy.AdmPwdEnabled -eq 1) { $laps = $true }
            if ((Get-CimInstance -ClassName Win32_ComputerSystem).DomainRole -ge 4) { $laps = $null }
            & $check 'LAPS skonfigurowany' $laps 'brak konfiguracji LAPS'
            # Aktualizacje
            $last = $null
            try { $last = @(Get-HotFix -ErrorAction Stop | Where-Object { $_.InstalledOn } | Sort-Object InstalledOn -Descending)[0].InstalledOn } catch { }
            $r['Ostatnia poprawka'] = $last
            & $check 'Aktualizacje świeże' $(if ($last) { ((Get-Date) - $last).TotalDays -le [int]$P.UpdateDays } else { $null }) "brak poprawek z ostatnich $($P.UpdateDays) dni"
            # Oczekujący restart
            $pending = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
            & $check 'Bez oczekującego restartu' (-not $pending) 'oczekuje na restart'
            # Secure Boot i TPM
            $sb = $null
            try { $sb = [bool](Confirm-SecureBootUEFI -ErrorAction Stop) } catch { }
            & $check 'Secure Boot' $sb 'Secure Boot wyłączony'
            $tpm = $null
            try {
                $t = Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftTpm' -ClassName Win32_Tpm -ErrorAction Stop
                $tpm = [bool]($t -and $t.IsEnabled_InitialValue -and $t.IsActivated_InitialValue)
            }
            catch { }
            & $check 'TPM aktywny' $tpm 'brak aktywnego modułu TPM'
            # WDigest, LSA, LLMNR
            $wd = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' -ErrorAction SilentlyContinue).UseLogonCredential
            & $check 'WDigest wyłączony' ($null -eq $wd -or [int]$wd -eq 0) 'WDigest przechowuje hasła w pamięci'
            $ppl = (Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' -ErrorAction SilentlyContinue).RunAsPPL
            & $check 'Ochrona LSA' ($null -ne $ppl -and [int]$ppl -ge 1) 'brak ochrony procesu LSA (RunAsPPL)'
            $llmnr = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' -ErrorAction SilentlyContinue).EnableMulticast
            & $check 'LLMNR wyłączony' ($null -ne $llmnr -and [int]$llmnr -eq 0) 'włączony LLMNR'
            # Konta lokalne
            $guest = $null
            $admins = $null
            try {
                $g = @(Get-LocalUser -ErrorAction Stop | Where-Object { [string]$_.SID -match '-501$' })
                if ($g.Count) { $guest = -not $g[0].Enabled }
                $admins = @(Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop).Count
            }
            catch { }
            & $check 'Konto Gość wyłączone' $guest 'włączone konto Gość'
            $r['Administratorzy lokalni'] = $admins
            # Ocena: odsetek spełnionych kontroli (pomijamy niedostępne)
            $evaluated = @($r.Keys | Where-Object { $r[$_] -is [bool] -and $_ -ne 'RDP włączony' })
            $passed = @($evaluated | Where-Object { $r[$_] -eq $true }).Count
            $score = if ($evaluated.Count) { [int][Math]::Round($passed / $evaluated.Count * 100) } else { 0 }
            $result = [ordered]@{
                'Wynik'    = '{0}% ({1}/{2})' -f $score, $passed, $evaluated.Count
                'Problemy' = (@($problems) -join '; ')
            }
            foreach ($k in $r.Keys) { $result[$k] = $r[$k] }
            $result['__tone'] = if ($score -ge 85) { 'ok' } elseif ($score -ge 60) { 'warn' } else { 'crit' }
            $result['__score'] = $score
            [pscustomobject]$result
        } -OnComplete {
            param($m)
            $rows = @($m.Table.Rows | Where-Object { $m.Table.Columns.Contains('__score') -and $_['__score'] -isnot [System.DBNull] })
            if ($rows.Count -eq 0) { return }
            $sum = 0
            foreach ($r in $rows) { $sum += [int]$r['__score'] }
            $avg = [int][Math]::Round($sum / $rows.Count)
            Set-StatTile -Module $m -Key 'avg' -Value "$avg%" -Tone $(if ($avg -ge 85) { 'ok' } elseif ($avg -ge 60) { 'warn' } else { 'crit' })
            Set-StatTile -Module $m -Key 'good' -Value ([string]@($rows | Where-Object { [int]$_['__score'] -ge 85 }).Count) -Tone 'ok'
            $bad = @($rows | Where-Object { [int]$_['__score'] -lt 60 }).Count
            Set-StatTile -Module $m -Key 'bad' -Value ([string]$bad) -Tone $(if ($bad) { 'crit' } else { '' })
        }
    } | Out-Null
    Add-StatTile -Module $m -Key 'avg' -Label 'Średni wynik' -Icon 'EA18' | Out-Null
    Add-StatTile -Module $m -Key 'good' -Label 'Komputery ≥ 85%' -Icon 'E73E' | Out-Null
    Add-StatTile -Module $m -Key 'bad' -Label 'Komputery < 60%' -Icon 'E7BA' | Out-Null
    $m.ResultHint = 'Puste pole = kontrola niedostępna na tym komputerze'
}

Register-Module -Workspace 'Remote' -Category 'Bezpieczeństwo' -Key 'Defender' -Title 'Microsoft Defender' -Icon 'E83D' `
    -Description 'Stan ochrony i sygnatur, wykryte zagrożenia, aktualizacja sygnatur oraz skanowanie (skan trwa w tle – można pracować dalej).' -Build {
    param($m)
    $m.ColorBools = $true
    $row = Add-ToolbarRow -Module $m -Title 'Stan'
    Add-Button -Parent $row -Text 'Stan ochrony' -Icon 'E83D' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Defender – stan' -Targets $targets -ScriptBlock {
            param($P)
            if (-not (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue)) { throw 'Brak modułu Defender (Get-MpComputerStatus) na komputerze.' }
            $s = Get-MpComputerStatus -ErrorAction Stop
            [pscustomobject]@{
                'Antywirus'                 = $s.AntivirusEnabled
                'Ochrona w czasie rzecz.'   = $s.RealTimeProtectionEnabled
                'Ochrona przed naruszeniem' = $s.IsTamperProtected
                'Tryb'                      = $s.AMRunningMode
                'Wersja sygnatur'           = $s.AntivirusSignatureVersion
                'Sygnatury z dnia'          = $s.AntivirusSignatureLastUpdated
                'Wiek sygnatur (dni)'       = $s.AntivirusSignatureAge
                'Ostatni szybki skan'       = $s.QuickScanEndTime
                'Ostatni pełny skan'        = $s.FullScanEndTime
                'Wersja silnika'            = $s.AMEngineVersion
                '__flag'                    = $(if (-not $s.RealTimeProtectionEnabled -or $s.AntivirusSignatureAge -gt 3) { 'warn' } else { '' })
            }
        }
    } | Out-Null
    Add-Button -Parent $row -Text 'Wykryte zagrożenia' -Icon 'E7BA' -Module $m -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Defender – zagrożenia' -Targets $targets -ScriptBlock {
            param($P)
            if (-not (Get-Command Get-MpThreatDetection -ErrorAction SilentlyContinue)) { throw 'Brak modułu Defender na komputerze.' }
            $names = @{}
            foreach ($t in @(Get-MpThreat -ErrorAction SilentlyContinue)) { $names[[string]$t.ThreatID] = $t.ThreatName }
            $detections = @(Get-MpThreatDetection -ErrorAction SilentlyContinue)
            if ($detections.Count -eq 0) { return [pscustomobject]@{ 'Zagrożenie' = '(brak wykrytych zagrożeń)'; '__flag' = 'muted' } }
            foreach ($d in $detections) {
                [pscustomobject]@{
                    'Wykryto'     = $d.InitialDetectionTime
                    'Zagrożenie'  = $names[[string]$d.ThreatID]
                    'Zasoby'      = (@($d.Resources) -join '; ')
                    'Akcja udana' = $d.ActionSuccess
                    'Proces'      = $d.ProcessName
                    'Użytkownik'  = $d.DomainUser
                    '__flag'      = 'crit'
                }
            }
        }
    } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Akcje'
    $defAction = {
        param($m, $s)
        $op = [string]$s.Tag
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $label = @{ Update = 'Zaktualizować sygnatury'; QuickScan = 'Uruchomić szybkie skanowanie'; FullScan = 'Uruchomić PEŁNE skanowanie (może trwać godzinami i obciąża dysk)' }[$op]
        if (-not (Confirm-Action -Text "$label na $($targets.Count) komputer(ach)?" -Items $targets -ConfirmText 'Wykonaj')) { return }
        Start-HostOperation -Module $m -Name "Defender – $op" -Targets $targets -Output Log -Parameters @{ Op = $op } -ScriptBlock {
            param($P)
            if (-not (Get-Command Start-MpScan -ErrorAction SilentlyContinue)) { throw 'Brak modułu Defender na komputerze.' }
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            switch ($P.Op) {
                'Update' { Update-MpSignature -ErrorAction Stop; $v = (Get-MpComputerStatus).AntivirusSignatureVersion; "Sygnatury zaktualizowane (wersja $v)." }
                'QuickScan' { Start-MpScan -ScanType QuickScan -ErrorAction Stop; "Szybkie skanowanie zakończone ($([Math]::Round($sw.Elapsed.TotalMinutes, 1)) min)." }
                'FullScan' { Start-MpScan -ScanType FullScan -ErrorAction Stop; "Pełne skanowanie zakończone ($([Math]::Round($sw.Elapsed.TotalMinutes, 1)) min)." }
            }
        }
    }
    foreach ($a in @(@('Aktualizuj sygnatury', 'Update', 'E895'), @('Szybki skan', 'QuickScan', 'E721'), @('Pełny skan', 'FullScan', 'E9D9'))) {
        $b = Add-Button -Parent $row2 -Text $a[0] -Icon $a[2] -Module $m -OnClick $defAction
        $b.Tag = $a[1]
    }
}

Register-Module -Workspace 'Remote' -Category 'Bezpieczeństwo' -Key 'BitLocker' -Title 'BitLocker' -Icon 'E72E' `
    -Description 'Stan szyfrowania woluminów i kopia zapasowa kluczy odzyskiwania do AD. Klucze zapisane w AD: przestrzeń «Komputery AD».' -Build {
    param($m)
    $m.PillColumns = @('Ochrona')
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row -Text 'Stan woluminów' -Icon 'E72E' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'BitLocker – stan' -Targets $targets -ScriptBlock {
            param($P)
            if (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue) {
                foreach ($v in @(Get-BitLockerVolume -ErrorAction Stop)) {
                    $prot = [string]$v.ProtectionStatus
                    [pscustomobject]@{
                        'Wolumin'          = $v.MountPoint
                        'Ochrona'          = $(if ($prot -eq 'On') { 'Włączona' } elseif ($prot -eq 'Off') { 'Wyłączona' } else { $prot })
                        'Typ'              = [string]$v.VolumeType
                        'Stan'             = [string]$v.VolumeStatus
                        'Zaszyfrowano (%)' = $v.EncryptionPercentage
                        'Metoda'           = [string]$v.EncryptionMethod
                        'Blokada'          = [string]$v.LockStatus
                        'Zabezpieczenia'   = (@($v.KeyProtector | ForEach-Object { [string]$_.KeyProtectorType }) -join ', ')
                        '__tone'           = $(if ($prot -eq 'On') { 'ok' } elseif ([string]$v.VolumeType -eq 'OperatingSystem') { 'crit' } else { 'warn' })
                    }
                }
                return
            }
            # Starsze systemy / brak modułu: klasa WMI (niezależna od języka systemu)
            $conversion = @{ 0 = 'Odszyfrowany'; 1 = 'Zaszyfrowany'; 2 = 'Szyfrowanie w toku'; 3 = 'Odszyfrowywanie w toku'; 4 = 'Szyfrowanie wstrzymane'; 5 = 'Odszyfrowywanie wstrzymane' }
            $volumes = @(Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftVolumeEncryption' -ClassName Win32_EncryptableVolume -ErrorAction Stop)
            foreach ($v in $volumes) {
                $status = Invoke-CimMethod -InputObject $v -MethodName GetConversionStatus -ErrorAction SilentlyContinue
                $on = ([int]$v.ProtectionStatus -eq 1)
                [pscustomobject]@{
                    'Wolumin'          = $v.DriveLetter
                    'Ochrona'          = $(if ($on) { 'Włączona' } else { 'Wyłączona' })
                    'Stan'             = $(if ($status) { $conversion[[int]$status.ConversionStatus] } else { '' })
                    'Zaszyfrowano (%)' = $(if ($status) { $status.EncryptionPercentage } else { $null })
                    '__tone'           = $(if ($on) { 'ok' } else { 'warn' })
                }
            }
        }
    } | Out-Null
    Add-Button -Parent $row -Text 'Kopia kluczy do AD' -Icon 'E74E' -Module $m -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        if (-not (Confirm-Action -Text "Zapisać w AD klucze odzyskiwania BitLocker (wszystkie woluminy) z $($targets.Count) komputer(ów)?" -Items $targets -ConfirmText 'Zapisz w AD')) { return }
        Start-HostOperation -Module $m -Name 'BitLocker – kopia do AD' -Targets $targets -Output Log -ScriptBlock {
            param($P)
            if (Get-Command Backup-BitLockerKeyProtector -ErrorAction SilentlyContinue) {
                $found = $false
                foreach ($v in @(Get-BitLockerVolume -ErrorAction Stop)) {
                    foreach ($kp in @($v.KeyProtector | Where-Object { [string]$_.KeyProtectorType -eq 'RecoveryPassword' })) {
                        $found = $true
                        try {
                            Backup-BitLockerKeyProtector -MountPoint $v.MountPoint -KeyProtectorId $kp.KeyProtectorId -ErrorAction Stop | Out-Null
                            [pscustomobject]@{ 'Wolumin' = $v.MountPoint; 'Klucz' = $kp.KeyProtectorId; 'Wynik' = 'Zapisano w AD' }
                        }
                        catch { [pscustomobject]@{ 'Wolumin' = $v.MountPoint; 'Klucz' = $kp.KeyProtectorId; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
                    }
                }
                if (-not $found) { 'Brak zabezpieczeń typu hasło odzyskiwania (RecoveryPassword) – nie ma czego zapisać.' }
                return
            }
            $bde = Join-Path $env:SystemRoot 'System32\manage-bde.exe'
            if (-not (Test-Path -LiteralPath $bde)) { throw 'Brak Get-BitLockerVolume i manage-bde.exe na komputerze.' }
            $volumes = @(Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftVolumeEncryption' -ClassName Win32_EncryptableVolume -ErrorAction Stop | Where-Object { $_.DriveLetter })
            foreach ($v in $volumes) {
                $text = (& $bde -protectors -get $v.DriveLetter -Type RecoveryPassword 2>&1) -join "`n"
                $ids = @([regex]::Matches($text, '\{[0-9A-Fa-f-]{36}\}') | ForEach-Object { $_.Value } | Select-Object -Unique)
                foreach ($id in $ids) {
                    $out = & $bde -protectors -adbackup $v.DriveLetter -id $id 2>&1
                    $res = if ($LASTEXITCODE -eq 0) { 'Zapisano w AD' } else { "Błąd – $(($out | Out-String).Trim())" }
                    [pscustomobject]@{ 'Wolumin' = $v.DriveLetter; 'Klucz' = $id; 'Wynik' = $res }
                }
            }
        }
    } | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'Bezpieczeństwo' -Key 'Firewall' -Title 'Zapora Windows' -Icon 'E785' `
    -Description 'Stan profili zapory, reguły (z portami i programami), włączanie, wyłączanie i usuwanie zaznaczonych reguł oraz tworzenie nowych.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $params = @{ Direction = @('Inbound', 'Outbound')[$m.Direction.SelectedIndex]; EnabledOnly = (Test-Checked $m.EnabledOnly); Filter = $m.Filter.Text.Trim() }
        Start-HostOperation -Module $m -Name 'Reguły zapory' -Targets $targets -Parameters $params -ScriptBlock {
            param($P)
            $rules = @(Get-NetFirewallRule -Direction $P.Direction -ErrorAction Stop)
            if ($P.EnabledOnly) { $rules = @($rules | Where-Object { [string]$_.Enabled -eq 'True' }) }
            if ($P.Filter) {
                $f = "*$($P.Filter)*"
                $rules = @($rules | Where-Object { $_.DisplayName -like $f -or $_.Name -like $f -or $_.DisplayGroup -like $f })
            }
            if ($rules.Count -eq 0) { return }
            # Filtry pobierane hurtowo i łączone po InstanceID - zapytanie per reguła trwałoby minutami
            $ports = @{}
            foreach ($pf in @(Get-NetFirewallPortFilter -ErrorAction SilentlyContinue)) { $ports[$pf.InstanceID] = $pf }
            $apps = @{}
            foreach ($af in @(Get-NetFirewallApplicationFilter -ErrorAction SilentlyContinue)) { $apps[$af.InstanceID] = $af }
            $addresses = @{}
            foreach ($ad in @(Get-NetFirewallAddressFilter -ErrorAction SilentlyContinue)) { $addresses[$ad.InstanceID] = $ad }
            foreach ($r in $rules) {
                $pf = $ports[$r.InstanceID]
                $af = $apps[$r.InstanceID]
                $adf = $addresses[$r.InstanceID]
                [pscustomobject]@{
                    'Nazwa wyświetlana' = $r.DisplayName
                    'Włączona'          = ([string]$r.Enabled -eq 'True')
                    'Akcja'             = [string]$r.Action
                    'Profil'            = [string]$r.Profile
                    'Protokół'          = $(if ($pf) { [string]$pf.Protocol } else { '' })
                    'Port lokalny'      = $(if ($pf) { @($pf.LocalPort) -join ',' } else { '' })
                    'Port zdalny'       = $(if ($pf) { @($pf.RemotePort) -join ',' } else { '' })
                    'Adres zdalny'      = $(if ($adf) { @($adf.RemoteAddress) -join ',' } else { '' })
                    'Program'           = $(if ($af) { $af.Program } else { '' })
                    'Grupa'             = $r.DisplayGroup
                    'ID reguły'         = $r.Name
                    '__flag'            = $(if ([string]$r.Enabled -ne 'True') { 'muted' } elseif ([string]$r.Action -eq 'Block') { 'warn' } else { '' })
                }
            }
        }
    }
    $m.Actions.Change = {
        param($m, [string]$Op, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('ID reguły', 'Nazwa wyświetlana') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli reguły zapory (widok «Pokaż reguły»).'; return }
        $label = @{ Enable = 'Włączyć'; Disable = 'Wyłączyć'; Delete = 'Usunąć' }[$Op]
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) $i['Nazwa wyświetlana'] }
        if (-not (Confirm-Action -Text "$label wybrane reguły zapory?" -Items $items -ConfirmText $(if ($Op -eq 'Delete') { 'Usuń reguły' } else { 'Wykonaj' }) -Danger:($Op -ne 'Enable'))) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Op = $Op; Names = @($byHost[$h] | ForEach-Object { [string]$_['ID reguły'] }) } }
        Start-HostOperation -Module $m -Name "Zapora – $Op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            foreach ($name in $P.Names) {
                try {
                    switch ($P.Op) {
                        'Enable' { Enable-NetFirewallRule -Name $name -ErrorAction Stop }
                        'Disable' { Disable-NetFirewallRule -Name $name -ErrorAction Stop }
                        'Delete' { Remove-NetFirewallRule -Name $name -ErrorAction Stop }
                    }
                    [pscustomobject]@{ 'Reguła' = $name; 'Operacja' = $P.Op; 'Wynik' = 'OK' }
                }
                catch { [pscustomobject]@{ 'Reguła' = $name; 'Operacja' = $P.Op; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Profile'
    Add-Button -Parent $row -Text 'Stan profili zapory' -Icon 'E785' -Module $m -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Profile zapory' -Targets $targets -ScriptBlock {
            param($P)
            Get-NetFirewallProfile -ErrorAction Stop | ForEach-Object {
                [pscustomobject]@{
                    'Profil'               = $_.Name
                    'Włączony'             = ([string]$_.Enabled -eq 'True')
                    'Ruch przychodzący'    = [string]$_.DefaultInboundAction
                    'Ruch wychodzący'      = [string]$_.DefaultOutboundAction
                    'Rejestrowanie blokad' = [string]$_.LogBlocked
                }
            }
        }
    } | Out-Null
    $row1 = Add-ToolbarRow -Module $m -Title 'Reguły'
    $m.Direction = Add-ComboBox -Parent $row1 -Items @('Przychodzące', 'Wychodzące') -Width 150
    $m.EnabledOnly = Add-CheckBox -Parent $row1 -Text 'Tylko włączone' -Checked $true
    $m.Filter = Add-TextBox -Parent $row1 -Width 180 -Placeholder 'nazwa, grupa lub ID'
    Add-Button -Parent $row1 -Text 'Pokaż reguły' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Zaznaczone reguły'
    foreach ($a in @(@('Włącz', 'Enable', 'E73E'), @('Wyłącz', 'Disable', 'E8D8'))) {
        $b = Add-Button -Parent $row2 -Text $a[0] -Icon $a[2] -Module $m -OnClick { param($m, $s) & $m.Actions.Change $m ([string]$s.Tag) $null }
        $b.Tag = $a[1]
    }
    Add-Button -Parent $row2 -Text 'Usuń' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Change $m 'Delete' $null } | Out-Null
    Add-RowAction -Module $m -Text 'Włącz regułę' -Icon 'E73E' -Action { param($m, $rows) & $m.Actions.Change $m 'Enable' $rows }
    Add-RowAction -Module $m -Text 'Wyłącz regułę' -Icon 'E8D8' -Action { param($m, $rows) & $m.Actions.Change $m 'Disable' $rows }
    Add-RowAction -Module $m -Text 'Usuń regułę' -Icon 'E74D' -Danger -Separator -Action { param($m, $rows) & $m.Actions.Change $m 'Delete' $rows }

    $row3 = Add-ToolbarRow -Module $m -Title 'Nowa reguła'
    $m.NewName = Add-TextBox -Parent $row3 -Width 200 -Text 'Domain Ops – nowa reguła'
    $m.NewDirection = Add-ComboBox -Parent $row3 -Items @('Przychodząca', 'Wychodząca') -Width 140
    $m.NewAction = Add-ComboBox -Parent $row3 -Items @('Zezwalaj', 'Blokuj') -Width 110
    $m.NewProtocol = Add-ComboBox -Parent $row3 -Items @('TCP', 'UDP') -Width 80
    $m.NewPorts = Add-TextBox -Parent $row3 -Width 120 -Text '5985' -Placeholder 'porty'
    $m.NewProfile = Add-ComboBox -Parent $row3 -Items @('Wszystkie profile', 'Domena', 'Prywatny', 'Publiczny') -Width 160
    $m.NewRemote = Add-TextBox -Parent $row3 -Width 140 -Text 'Any' -Placeholder 'adres zdalny'
    Add-Button -Parent $row3 -Text 'Utwórz' -Icon 'E710' -Module $m -OnClick {
        param($m)
        $name = $m.NewName.Text.Trim()
        $ports = @(Split-ListText $m.NewPorts.Text)
        if (-not $name) { Show-Warning 'Podaj nazwę reguły.'; return }
        if ($ports.Count -eq 0 -or @($ports | Where-Object { $_ -notmatch '^\d{1,5}(-\d{1,5})?$' }).Count -gt 0) { Show-Warning 'Podaj porty jako liczby lub zakresy, np. 80, 443, 8000-8010.'; return }
        $remote = @(Split-ListText $m.NewRemote.Text)
        if ($remote.Count -eq 0) { $remote = @('Any') }
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $params = @{
            DisplayName   = $name
            Direction     = @('Inbound', 'Outbound')[$m.NewDirection.SelectedIndex]
            Action        = @('Allow', 'Block')[$m.NewAction.SelectedIndex]
            Protocol      = [string]$m.NewProtocol.SelectedItem
            Ports         = $ports
            Profile       = @('Any', 'Domain', 'Private', 'Public')[$m.NewProfile.SelectedIndex]
            RemoteAddress = $remote
        }
        if (-not (Confirm-Action -Text ("Utworzyć regułę «{0}» ({1}, {2}, {3}/{4})?" -f $name, $m.NewDirection.SelectedItem, $m.NewAction.SelectedItem, $params.Protocol, ($ports -join ',')) -Items $targets -ConfirmText 'Utwórz')) { return }
        Start-HostOperation -Module $m -Name 'Nowa reguła zapory' -Targets $targets -Output Log -Parameters $params -ScriptBlock {
            param($P)
            $rp = @{
                DisplayName   = $P.DisplayName
                Group         = 'Domain Ops'
                Direction     = $P.Direction
                Action        = $P.Action
                Protocol      = $P.Protocol
                LocalPort     = [string[]]@($P.Ports)
                RemoteAddress = [string[]]@($P.RemoteAddress)
                Profile       = $P.Profile
                Enabled       = 'True'
                ErrorAction   = 'Stop'
            }
            $rule = New-NetFirewallRule @rp
            [pscustomobject]@{ 'Reguła' = $P.DisplayName; 'ID reguły' = $rule.Name; 'Wynik' = 'Utworzono (grupa «Domain Ops»)' }
        }
    } | Out-Null
}

Register-Module -Workspace 'Remote' -Category 'Bezpieczeństwo' -Key 'Certificates' -Title 'Certyfikaty komputera' -Icon 'EB95' `
    -Description 'Certyfikaty z magazynów LocalMachine, wyszukiwanie wygasających oraz eksport zaznaczonych do plików .cer.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Ważność')
    $row = Add-ToolbarRow -Module $m -Title 'Magazyn'
    $m.Store = Add-ComboBox -Parent $row -Items @('My', 'Root', 'CA', 'TrustedPublisher', 'TrustedPeople', 'WebHosting', 'Remote Desktop') -Width 170
    $m.Filter = Add-TextBox -Parent $row -Width 180 -Placeholder 'podmiot, wystawca, odcisk'
    $m.OnlyExpiring = Add-CheckBox -Parent $row -Text 'Tylko wygasające w ciągu (dni)'
    $m.Days = Add-Numeric -Parent $row -Value 30 -Minimum 1 -Maximum 3650 -Width 60
    $row2 = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row2 -Text 'Pokaż certyfikaty' -Icon 'EB95' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $params = @{ Store = [string]$m.Store.SelectedItem; Filter = $m.Filter.Text.Trim(); ExpiringDays = $(if (Test-Checked $m.OnlyExpiring) { Get-Num $m.Days } else { 0 }) }
        Start-HostOperation -Module $m -Name 'Certyfikaty' -Targets $targets -Parameters $params -ScriptBlock {
            param($P)
            $now = Get-Date
            foreach ($c in @(Get-ChildItem -Path ('Cert:\LocalMachine\' + $P.Store) -ErrorAction Stop)) {
                if ($P.Filter) {
                    $f = "*$($P.Filter)*"
                    if ($c.Subject -notlike $f -and $c.Thumbprint -notlike $f -and $c.FriendlyName -notlike $f -and $c.Issuer -notlike $f) { continue }
                }
                $days = [int][Math]::Floor(($c.NotAfter - $now).TotalDays)
                if ($P.ExpiringDays -gt 0 -and $days -gt $P.ExpiringDays) { continue }
                [pscustomobject]@{
                    'Podmiot'            = $c.Subject
                    'Ważność'            = $(if ($days -lt 0) { 'Wygasł' } elseif ($days -le 30) { "Wygasa za $days dni" } else { 'Ważny' })
                    'Wystawca'           = $c.Issuer
                    'Ważny od'           = $c.NotBefore
                    'Ważny do'           = $c.NotAfter
                    'Dni do wygaśnięcia' = $days
                    'Klucz prywatny'     = $c.HasPrivateKey
                    'Przeznaczenie'      = (@($c.EnhancedKeyUsageList | ForEach-Object { $_.FriendlyName } | Where-Object { $_ }) -join ', ')
                    'Nazwa przyjazna'    = $c.FriendlyName
                    'Odcisk palca'       = $c.Thumbprint
                    'Magazyn'            = $P.Store
                    '__tone'             = $(if ($days -lt 0) { 'crit' } elseif ($days -le 30) { 'warn' } else { 'ok' })
                }
            }
        }
    } | Out-Null
    Add-Button -Parent $row2 -Text 'Eksportuj zaznaczone (.cer)…' -Icon 'EDE1' -Module $m -OnClick {
        param($m)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Odcisk palca', 'Magazyn')
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli certyfikaty do eksportu.'; return }
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Folder docelowy dla plików .cer'
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $m.Data.ExportFolder = $dlg.SelectedPath
        $per = @{}
        foreach ($h in $byHost.Keys) {
            $per[$h] = @{ Items = @($byHost[$h] | ForEach-Object { @{ Thumbprint = [string]$_['Odcisk palca']; Store = [string]$_['Magazyn'] } }) }
        }
        Start-HostOperation -Module $m -Name 'Eksport certyfikatów' -Targets @($byHost.Keys) -PerTarget $per -Output None -ScriptBlock {
            param($P)
            foreach ($i in $P.Items) {
                $cert = Get-Item -Path ('Cert:\LocalMachine\{0}\{1}' -f $i.Store, $i.Thumbprint) -ErrorAction Stop
                [pscustomobject]@{ Thumbprint = $cert.Thumbprint; Base64 = [Convert]::ToBase64String($cert.RawData) }
            }
        } -OnResult {
            param($m, $r)
            foreach ($d in @($r.Data)) {
                $thumb = [string](Get-ObjectValue $d 'Thumbprint')
                $b64 = [string](Get-ObjectValue $d 'Base64')
                if (-not $thumb -or -not $b64) { continue }
                $file = Join-Path $m.Data.ExportFolder ('{0}_{1}.cer' -f $r.Target, $thumb)
                [System.IO.File]::WriteAllBytes($file, [Convert]::FromBase64String($b64))
                Write-Log ("[{0}] zapisano {1}" -f $r.Target, $file) 'OK'
            }
        }
    } | Out-Null
}
#endregion

#region Zarządzanie zdalne: Udostępnianie
Register-Module -Workspace 'Remote' -Category 'Udostępnianie' -Key 'Shares' -Title 'Udziały sieciowe' -Icon 'E72D' `
    -Description 'Udziały SMB na zaznaczonych komputerach: podgląd, uprawnienia, tworzenie (z uprawnieniami udziału i opcjonalnie NTFS) oraz usuwanie.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.GoodWhenNo = @('Administracyjny')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Udziały' -Targets $targets -Parameters @{ ShowSpecial = (Test-Checked $m.ShowSpecial) } -ScriptBlock {
            param($P)
            if (Get-Command Get-SmbShare -ErrorAction SilentlyContinue) {
                $shares = @(Get-SmbShare -ErrorAction Stop)
                if (-not $P.ShowSpecial) { $shares = @($shares | Where-Object { -not $_.Special }) }
                foreach ($s in $shares) {
                    [pscustomobject]@{ 'Udział' = $s.Name; 'Ścieżka' = $s.Path; 'Opis' = $s.Description; 'Administracyjny' = [bool]$s.Special; 'Połączenia' = $s.CurrentUsers }
                }
            }
            else {
                foreach ($s in @(Get-CimInstance -ClassName Win32_Share)) {
                    $special = ([int64]$s.Type -ge 2147483648)
                    if ($special -and -not $P.ShowSpecial) { continue }
                    [pscustomobject]@{ 'Udział' = $s.Name; 'Ścieżka' = $s.Path; 'Opis' = $s.Description; 'Administracyjny' = $special }
                }
            }
        }
    }
    $m.Actions.Permissions = {
        param($m, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Udział') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli udziały.'; return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Names = @($byHost[$h] | ForEach-Object { [string]$_['Udział'] }) } }
        $m.Data.PermissionRows = New-Object System.Collections.ArrayList
        Start-HostOperation -Module $m -Name 'Uprawnienia udziałów' -Targets @($byHost.Keys) -PerTarget $per -Output None -ScriptBlock {
            param($P)
            foreach ($name in $P.Names) {
                Get-SmbShareAccess -Name $name -ErrorAction Stop | ForEach-Object {
                    [pscustomobject]@{ 'Udział' = $name; 'Konto' = $_.AccountName; 'Typ' = [string]$_.AccessControlType; 'Uprawnienie' = [string]$_.AccessRight }
                }
            }
        } -OnResult {
            param($m, $r)
            foreach ($d in @($r.Data)) {
                [void]$m.Data.PermissionRows.Add([pscustomobject]@{
                        'Komputer'    = $r.Target
                        'Udział'      = Get-ObjectValue $d 'Udział'
                        'Konto'       = Get-ObjectValue $d 'Konto'
                        'Typ'         = Get-ObjectValue $d 'Typ'
                        'Uprawnienie' = Get-ObjectValue $d 'Uprawnienie'
                    })
            }
        } -OnComplete {
            param($m)
            if ($m.Data.PermissionRows.Count -eq 0) { return }
            # Okno pokazujemy poza obsługą timera, żeby nie wstrzymywać innych trwających operacji
            Invoke-Deferred -Module $m -Action { param($m) Show-GridDialog -Title 'Uprawnienia udziałów' -Subtitle 'Uprawnienia na poziomie udziału (SMB)' -Rows @($m.Data.PermissionRows) }
        }
    }
    $m.Actions.Remove = {
        param($m, $Rows)
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Udział') -Rows $Rows
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli udziały do usunięcia.'; return }
        $items = Get-HostItemList -ByHost $byHost -Format { param($i) $i['Udział'] }
        if (-not (Confirm-Action -Text 'Usunąć wybrane udziały? Dane w folderach pozostaną nienaruszone.' -Items $items -ConfirmText 'Usuń udziały' -Danger)) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Names = @($byHost[$h] | ForEach-Object { [string]$_['Udział'] }) } }
        Start-HostOperation -Module $m -Name 'Usuwanie udziałów' -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            foreach ($name in $P.Names) {
                try {
                    if (@('ADMIN$', 'IPC$', 'PRINT$') -contains $name.ToUpperInvariant() -or $name -match '^[A-Za-z]\$$') { throw 'Pominięto udział administracyjny.' }
                    if (Get-Command Remove-SmbShare -ErrorAction SilentlyContinue) { Remove-SmbShare -Name $name -Force -ErrorAction Stop }
                    else {
                        $share = Get-CimInstance -ClassName Win32_Share -Filter ("Name = '{0}'" -f $name.Replace("'", "''"))
                        if (-not $share) { throw 'Nie znaleziono udziału.' }
                        $r = Invoke-CimMethod -InputObject $share -MethodName Delete
                        if ($r.ReturnValue -ne 0) { throw "Win32_Share.Delete zwrócił $($r.ReturnValue)." }
                    }
                    [pscustomobject]@{ 'Udział' = $name; 'Wynik' = 'Usunięto' }
                }
                catch { [pscustomobject]@{ 'Udział' = $name; 'Wynik' = "Błąd – $($_.Exception.Message)" } }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Lista'
    $m.ShowSpecial = Add-CheckBox -Parent $row -Text 'Pokaż udziały administracyjne (C$, ADMIN$…)'
    Add-Button -Parent $row -Text 'Pokaż udziały' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Uprawnienia zaznaczonych' -Icon 'E8D7' -Module $m -OnClick { param($m) & $m.Actions.Permissions $m $null } | Out-Null
    Add-Button -Parent $row -Text 'Usuń zaznaczone' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Remove $m $null } | Out-Null
    Add-RowAction -Module $m -Text 'Uprawnienia udziału' -Icon 'E8D7' -Action { param($m, $rows) & $m.Actions.Permissions $m $rows }
    Add-RowAction -Module $m -Text 'Otwórz w Eksploratorze' -Icon 'E838' -Action {
        param($m, $rows)
        $computer = [string](Get-ObjectValue $rows[0] 'Komputer')
        $share = [string](Get-ObjectValue $rows[0] 'Udział')
        Start-Tool -FilePath 'explorer.exe' -Arguments @("\\$computer\$share") -Name $computer
    }
    Add-RowAction -Module $m -Text 'Usuń udział' -Icon 'E74D' -Danger -Separator -Action { param($m, $rows) & $m.Actions.Remove $m $rows }

    $row2 = Add-ToolbarRow -Module $m -Title 'Nowy udział'
    $m.NewName = Add-TextBox -Parent $row2 -Width 150 -Placeholder 'Nazwa udziału'
    $m.NewPath = Add-TextBox -Parent $row2 -Width 220 -Text 'D:\Udzial' -Placeholder 'Ścieżka na komputerze'
    $m.NewDesc = Add-TextBox -Parent $row2 -Width 200 -Placeholder 'Opis'
    $row3 = Add-ToolbarRow -Module $m -Title 'Uprawnienia'
    $m.NewFull = Add-TextBox -Parent $row3 -Width 170 -Placeholder 'Pełna kontrola'
    $m.NewChange = Add-TextBox -Parent $row3 -Width 170 -Placeholder 'Zmiana'
    $m.NewRead = Add-TextBox -Parent $row3 -Width 170 -Placeholder 'Odczyt'
    $m.NewNtfs = Add-CheckBox -Parent $row3 -Text 'Nadaj też NTFS'
    $m.NewCreate = Add-CheckBox -Parent $row3 -Text 'Utwórz folder' -Checked $true
    Add-Button -Parent $row3 -Text 'Utwórz' -Icon 'E710' -Module $m -OnClick {
        param($m)
        $name = $m.NewName.Text.Trim()
        $path = $m.NewPath.Text.Trim()
        if (-not $name -or -not $path) { Show-Warning 'Podaj nazwę udziału i ścieżkę folderu na komputerze.'; return }
        if ($path -notmatch '^[A-Za-z]:\\') { Show-Warning 'Ścieżka musi być lokalną ścieżką na komputerze, np. D:\Dane\Projekty.'; return }
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $params = @{
            Name = $name; Path = $path; Description = $m.NewDesc.Text.Trim()
            Full = @(Split-ListText $m.NewFull.Text); Change = @(Split-ListText $m.NewChange.Text); Read = @(Split-ListText $m.NewRead.Text)
            Ntfs = (Test-Checked $m.NewNtfs); CreateFolder = (Test-Checked $m.NewCreate)
        }
        if (-not (Confirm-Action -Text ("Utworzyć udział {0} -> {1}?" -f $name, $path) -Items $targets -ConfirmText 'Utwórz')) { return }
        Start-HostOperation -Module $m -Name 'Tworzenie udziału' -Targets $targets -Output Log -Parameters $params -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($P)
            if (-not (Test-Path -LiteralPath $P.Path)) {
                if ($P.CreateFolder) { New-Item -ItemType Directory -Path $P.Path -Force | Out-Null }
                else { throw "Folder $($P.Path) nie istnieje." }
            }
            $full = @($P.Full | Where-Object { $_ })
            $change = @($P.Change | Where-Object { $_ })
            $read = @($P.Read | Where-Object { $_ })
            if (Get-Command New-SmbShare -ErrorAction SilentlyContinue) {
                $sp = @{ Name = $P.Name; Path = $P.Path; ErrorAction = 'Stop' }
                if ($P.Description) { $sp.Description = $P.Description }
                if ($full.Count) { $sp.FullAccess = [string[]]$full }
                if ($change.Count) { $sp.ChangeAccess = [string[]]$change }
                if ($read.Count) { $sp.ReadAccess = [string[]]$read }
                New-SmbShare @sp | Out-Null
            }
            else {
                $r = Invoke-CimMethod -ClassName Win32_Share -MethodName Create -Arguments @{ Path = $P.Path; Name = $P.Name; Type = [uint32]0; Description = [string]$P.Description }
                if ($r.ReturnValue -ne 0) { throw "Win32_Share.Create zwrócił $($r.ReturnValue)." }
            }
            $ntfs = ''
            if ($P.Ntfs) {
                $grants = @()
                foreach ($a in $full) { $grants += "${a}:(OI)(CI)F" }
                foreach ($a in $change) { $grants += "${a}:(OI)(CI)M" }
                foreach ($a in $read) { $grants += "${a}:(OI)(CI)RX" }
                foreach ($g in $grants) {
                    $out = & icacls.exe $P.Path /grant $g 2>&1
                    if ($LASTEXITCODE -ne 0) { throw "icacls $g : $(($out | Out-String).Trim())" }
                }
                if ($grants.Count) { $ntfs = ' + NTFS' }
            }
            [pscustomobject]@{ 'Udział' = $P.Name; 'Ścieżka' = $P.Path; 'Wynik' = "Utworzono$ntfs" }
        }
    } | Out-Null
}
#endregion

#region Active Directory: wspólne
# Operacje AD wykonywane są lokalnie w puli wątków modułem ActiveDirectory. Blok modułu dostaje gotową
# hashtablę $ad (Server, Credential, ErrorAction = Stop) do rozwinięcia w poleceniach: Get-ADUser ... @ad
$script:AdPrelude = @'
Import-Module ActiveDirectory -ErrorAction Stop -Verbose:$false
$ad = @{ ErrorAction = 'Stop' }
if ($Ctx.Server) { $ad.Server = $Ctx.Server }
if ($Ctx.Credential) { $ad.Credential = $Ctx.Credential }
'@

function Test-AdAvailable {
    if (Get-Module -ListAvailable -Name ActiveDirectory) { return $true }
    Show-Warning 'Ta funkcja wymaga modułu ActiveDirectory (RSAT: Active Directory Domain Services). Zainstaluj RSAT i uruchom program ponownie.'
    return $false
}

function Start-AdOperation {
    # Start-HostOperation w trybie lokalnym z modułem ActiveDirectory i hashtablą $ad w bloku skryptu
    param(
        [Parameter(Mandatory)][hashtable]$Module,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][AllowEmptyCollection()][string[]]$Targets,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [hashtable]$Parameters = @{},
        [hashtable]$PerTarget = @{},
        [ValidateSet('Grid', 'Log', 'None')][string]$Output = 'Grid',
        [string]$TargetColumn = '',
        [switch]$Append,
        [scriptblock]$OnResult,
        [scriptblock]$OnComplete
    )
    if (-not (Test-AdAvailable)) { return }
    $text = 'param($Target, $P, $Ctx)' + "`n" + $script:AdPrelude + "`n" + '$__body = {' + $ScriptBlock.ToString() + "`n}`n" + '& $__body $Target $P $Ctx'
    $sp = @{
        Module       = $Module
        Name         = $Name
        Targets      = $Targets
        ScriptBlock  = [scriptblock]::Create($text)
        Local        = $true
        Parameters   = $Parameters
        PerTarget    = $PerTarget
        Output       = $Output
        TargetColumn = $TargetColumn
        Append       = $Append
    }
    if ($OnResult) { $sp.OnResult = $OnResult }
    if ($OnComplete) { $sp.OnComplete = $OnComplete }
    Start-HostOperation @sp
}

function Add-ResultsToTargets {
    # Przenosi obiekty z tabeli wyników na listę po lewej (i zaznacza je) - most między raportami a operacjami
    param([hashtable]$Module, [ValidateSet('User', 'Computer')][string]$Kind, [string]$Column, $Rows = $null)
    $source = @(if ($null -ne $Rows) { $Rows } else { Get-SelectedResultRows -Module $Module })
    if ($source.Count -le 1 -and $null -eq $Rows) { $source = @($Module.View | ForEach-Object { $_ }) }
    $names = @($source | ForEach-Object { [string](Get-ObjectValue $_ $Column) } | Where-Object { $_ } | Select-Object -Unique)
    if ($names.Count -eq 0) { Show-Warning 'Brak obiektów do przeniesienia na listę.'; return }
    if ($Kind -eq 'User') {
        $items = foreach ($drv in $source) {
            $login = [string](Get-ObjectValue $drv $Column)
            if (-not $login) { continue }
            [pscustomobject]@{
                Login      = $login
                Name       = Get-ObjectValue $drv 'Nazwa'
                Enabled    = $(switch ([string](Get-ObjectValue $drv 'Włączone')) { 'Tak' { $true } 'Nie' { $false } default { $null } })
                Department = Get-ObjectValue $drv 'Dział'
                DN         = Get-ObjectValue $drv 'DN'
            }
        }
        $panel = $script:UI.UserPanel
        Set-TargetCheck -Panel $panel -Mode UncheckAll
        $added = Import-TargetRows -Panel $panel -Items @($items) -Source 'raport' -Check
    }
    else {
        $items = foreach ($drv in $source) {
            $name = [string](Get-ObjectValue $drv $Column)
            if (-not $name) { continue }
            [pscustomobject]@{ Name = $name; OS = Get-ObjectValue $drv 'System'; DN = Get-ObjectValue $drv 'DN' }
        }
        $panel = $script:UI.ComputerPanel
        Set-TargetCheck -Panel $panel -Mode UncheckAll
        $added = Import-TargetRows -Panel $panel -Items @($items) -Source 'raport' -Check
    }
    Show-Toast ("Na liście zaznaczono {0} obiektów (nowych: {1})." -f $names.Count, $added) 'ok'
}

function Get-LdapDate {
    # Data w formacie generalizedTime/LDAP do filtrów (np. whenCreated>=...)
    param([datetime]$Date)
    return $Date.ToUniversalTime().ToString('yyyyMMddHHmmss.0Z')
}

function Register-GroupMembershipModule {
    # Członkostwo w grupach - wspólny moduł dla użytkowników i komputerów
    param([string]$Workspace, [ValidateSet('User', 'Computer')][string]$Kind, [string]$Key, [string]$Category)
    $title = 'Członkostwo w grupach'
    $desc = if ($Kind -eq 'User') { 'Grupy zaznaczonych kont (bezpośrednie i zagnieżdżone), dodawanie do grup, usuwanie z zaznaczonych grup i kopiowanie członkostwa z konta wzorcowego.' }
    else { 'Grupy zaznaczonych kont komputerów (bezpośrednie i zagnieżdżone), dodawanie do grup i usuwanie z zaznaczonych grup.' }
    $build = {
        param($m)
        $m.PillColumns = @('Członkostwo')
        $m.Data.Column = if ($m.Target -eq 'User') { 'Login' } else { 'Komputer' }
        $m.Actions.Targets = {
            param($m)
            if ($m.Target -eq 'User') { return @(Get-TargetUsers) }
            return @(Get-TargetComputers | ForEach-Object { ($_ -split '\.')[0] })
        }
        $m.Actions.List = {
            param($m)
            $targets = @(& $m.Actions.Targets $m)
            if (-not $targets) { return }
            Start-AdOperation -Module $m -Name 'Grupy' -Targets $targets -TargetColumn $m.Data.Column -Parameters @{ Nested = (Test-Checked $m.Nested); Kind = $m.Target } -ScriptBlock {
                $obj = if ($P.Kind -eq 'User') { Get-ADUser -Identity $Target -Properties memberOf, PrimaryGroupID @ad } else { Get-ADComputer -Identity $Target -Properties memberOf, PrimaryGroupID @ad }
                $direct = @{}
                foreach ($dn in @($obj.memberOf)) { $direct[[string]$dn] = $true }
                $groups = @()
                if ($P.Nested) { $groups = @(Get-ADGroup -LDAPFilter ("(member:1.2.840.113556.1.4.1941:={0})" -f $obj.DistinguishedName) -Properties Description @ad) }
                else { $groups = @($obj.memberOf | ForEach-Object { Get-ADGroup -Identity $_ -Properties Description @ad }) }
                # Grupa podstawowa (Domain Users / Domain Computers) nie występuje w memberOf
                try {
                    $domainSid = $obj.SID.AccountDomainSid.Value
                    $primary = Get-ADGroup -Identity ('{0}-{1}' -f $domainSid, $obj.PrimaryGroupID) -Properties Description @ad
                    if ($primary) { $groups += $primary; $direct[$primary.DistinguishedName] = $true }
                }
                catch { }
                foreach ($g in ($groups | Sort-Object Name -Unique)) {
                    $isDirect = $direct.ContainsKey([string]$g.DistinguishedName)
                    [pscustomobject]@{
                        'Grupa'       = $g.Name
                        'Członkostwo' = $(if ($g.SID.Value -match '-(513|515)$' -and $isDirect) { 'Podstawowa' } elseif ($isDirect) { 'Bezpośrednie' } else { 'Zagnieżdżone' })
                        'Zakres'      = [string]$g.GroupScope
                        'Typ'         = [string]$g.GroupCategory
                        'Opis'        = $g.Description
                        'DN grupy'    = $g.DistinguishedName
                        '__tone'      = $(if ($isDirect) { 'info' } else { '' })
                    }
                }
            }
        }
        $m.Actions.Add = {
            param($m, [object[]]$Groups)
            $targets = @(& $m.Actions.Targets $m)
            if (-not $targets) { return }
            if (-not $Groups) {
                $Groups = Select-AdGroups -Title 'Dodaj do grup' -Subtitle "Zaznaczone obiekty ($($targets.Count)) zostaną dodane do wybranych grup."
                if (-not $Groups) { return }
            }
            $names = @($Groups | ForEach-Object { $_.Name })
            if (-not (Confirm-Action -Text ("Dodać {0} obiekt(ów) do grup: {1}?" -f $targets.Count, ($names -join ', ')) -Items $targets -ConfirmText 'Dodaj do grup')) { return }
            Start-AdOperation -Module $m -Name 'Dodawanie do grup' -Targets $targets -TargetColumn $m.Data.Column -Output Log -Parameters @{ Groups = @($Groups | ForEach-Object { $_.DN }); Kind = $m.Target } -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
                $obj = if ($P.Kind -eq 'User') { Get-ADUser -Identity $Target @ad } else { Get-ADComputer -Identity $Target @ad }
                foreach ($g in $P.Groups) {
                    try {
                        Add-ADGroupMember -Identity $g -Members $obj @ad
                        [pscustomobject]@{ 'Obiekt' = $Target; 'Grupa' = ($g -replace '^CN=((?:\\.|[^,])+),.*$', '$1'); 'Wynik' = 'Dodano' }
                    }
                    catch { [pscustomobject]@{ 'Obiekt' = $Target; 'Grupa' = ($g -replace '^CN=((?:\\.|[^,])+),.*$', '$1'); 'Wynik' = "Błąd – $($_.Exception.Message)" } }
                }
            }
        }
        $m.Actions.Remove = {
            param($m, $Rows)
            $byObj = Get-SelectedRowsByHost -Module $m -Columns @('DN grupy', 'Grupa', 'Członkostwo') -Rows $Rows -TargetColumn $m.Data.Column
            $per = @{}
            $items = @()
            foreach ($o in $byObj.Keys) {
                $groups = @($byObj[$o] | Where-Object { $_['Członkostwo'] -eq 'Bezpośrednie' })
                if ($groups.Count -eq 0) { continue }
                $per[$o] = @{ Groups = @($groups | ForEach-Object { [string]$_['DN grupy'] }); Kind = $m.Target }
                foreach ($g in $groups) { $items += ('{0}: {1}' -f $o, $g['Grupa']) }
            }
            if ($per.Count -eq 0) { Show-Warning 'Zaznacz w tabeli grupy z członkostwem bezpośrednim (zagnieżdżonego i podstawowego nie da się usunąć z tego miejsca).'; return }
            if (-not (Confirm-Action -Text 'Usunąć obiekty z wybranych grup?' -Items $items -ConfirmText 'Usuń z grup' -Danger)) { return }
            Start-AdOperation -Module $m -Name 'Usuwanie z grup' -Targets @($per.Keys) -PerTarget $per -TargetColumn $m.Data.Column -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
                $obj = if ($P.Kind -eq 'User') { Get-ADUser -Identity $Target @ad } else { Get-ADComputer -Identity $Target @ad }
                foreach ($g in $P.Groups) {
                    try {
                        Remove-ADGroupMember -Identity $g -Members $obj -Confirm:$false @ad
                        [pscustomobject]@{ 'Obiekt' = $Target; 'Grupa' = ($g -replace '^CN=((?:\\.|[^,])+),.*$', '$1'); 'Wynik' = 'Usunięto' }
                    }
                    catch { [pscustomobject]@{ 'Obiekt' = $Target; 'Grupa' = ($g -replace '^CN=((?:\\.|[^,])+),.*$', '$1'); 'Wynik' = "Błąd – $($_.Exception.Message)" } }
                }
            }
        }
        $row = Add-ToolbarRow -Module $m -Title 'Lista'
        $m.Nested = Add-CheckBox -Parent $row -Text 'Uwzględnij grupy zagnieżdżone' -ToolTip 'Pełne członkowanie (także przez inne grupy) – LDAP_MATCHING_RULE_IN_CHAIN'
        Add-Button -Parent $row -Text 'Pokaż grupy' -Icon 'E902' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
        $row2 = Add-ToolbarRow -Module $m -Title 'Zmiany'
        Add-Button -Parent $row2 -Text 'Dodaj do grup…' -Icon 'E710' -Module $m -OnClick { param($m) & $m.Actions.Add $m $null } | Out-Null
        Add-Button -Parent $row2 -Text 'Usuń z zaznaczonych grup' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Remove $m $null } | Out-Null
        if ($m.Target -eq 'User') {
            Add-Button -Parent $row2 -Text 'Kopiuj grupy z konta…' -Icon 'E8C8' -Module $m -OnClick {
                param($m)
                $template = Show-InputDialog -Title 'Konto wzorcowe' -Prompt 'Login konta, którego grupy (członkostwo bezpośrednie) mają otrzymać zaznaczone konta.' -Icon 'E77B' -Validate { param($t) if ($t.Trim() -match '^[^\s"/\\\[\]:;|=,+*?<>]+$') { '' } else { 'Podaj poprawny login (sAMAccountName).' } }
                if (-not $template) { return }
                Import-AdModule
                $ad = Get-AdSplat
                $user = Invoke-WithWaitCursor { Get-ADUser -Identity $template.Trim() -Properties memberOf @ad }
                $groups = @($user.memberOf | ForEach-Object { [pscustomobject]@{ Name = (Get-RdnValue $_); DN = [string]$_ } })
                if ($groups.Count -eq 0) { Show-Warning "Konto $template nie należy bezpośrednio do żadnej grupy (poza podstawową)."; return }
                & $m.Actions.Add $m $groups
            } | Out-Null
        }
        Add-RowAction -Module $m -Text 'Usuń z grupy' -Icon 'E74D' -Danger -Action { param($m, $rows) & $m.Actions.Remove $m $rows }
        Add-RowAction -Module $m -Text 'Kopiuj nazwy grup' -Icon 'E8C8' -Action {
            param($m, $rows)
            $names = @($rows | ForEach-Object { [string](Get-ObjectValue $_ 'Grupa') } | Select-Object -Unique)
            Set-ClipboardText ($names -join [Environment]::NewLine)
            Show-Toast "Skopiowano nazw grup: $($names.Count)" 'ok'
        }
    }
    Register-Module -Workspace $Workspace -Category $Category -Key $Key -Title $title -Icon 'E902' -Description $desc -Build $build
}
#endregion

#region Użytkownicy AD
$script:UserStateScript = {
    # Wspólna ocena stanu konta (wiersz wyniku dla modułów kont)
    param($u)
    $now = Get-Date
    $expiry = $null
    $raw = $u.'msDS-UserPasswordExpiryTimeComputed'
    if ($raw -and [int64]$raw -gt 0 -and [int64]$raw -lt [int64]::MaxValue) { try { $expiry = [datetime]::FromFileTime([int64]$raw) } catch { } }
    $state = 'Aktywne'
    $tone = 'ok'
    if ($u.LockedOut) { $state = 'Zablokowane'; $tone = 'crit' }
    elseif (-not $u.Enabled) { $state = 'Wyłączone'; $tone = '' }
    elseif ($u.AccountExpirationDate -and $u.AccountExpirationDate -lt $now) { $state = 'Wygasłe'; $tone = 'crit' }
    elseif ($u.PasswordExpired) { $state = 'Hasło wygasło'; $tone = 'warn' }
    elseif ($expiry -and $expiry -lt $now.AddDays(7)) { $state = 'Hasło wkrótce wygaśnie'; $tone = 'warn' }
    return @{ State = $state; Tone = $tone; Expiry = $expiry }
}

Register-Module -Workspace 'AdUsers' -Category 'Konta' -Key 'UserDetails' -Title 'Szczegóły konta' -Icon 'E779' `
    -Description 'Pełne informacje o zaznaczonych kontach: stan, kontakt, przełożony, logowania, hasło, wygaśnięcie, profil i jednostka organizacyjna.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Stan')
    $m.GoodWhenNo = @('Zablokowane', 'Hasło nigdy nie wygasa', 'Zmiana hasła przy logowaniu')
    $row = Add-ToolbarRow -Module $m -Title 'Akcje'
    Add-Button -Parent $row -Text 'Pobierz szczegóły' -Icon 'E896' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetUsers)
        if (-not $targets) { return }
        Start-AdOperation -Module $m -Name 'Szczegóły kont' -Targets $targets -TargetColumn 'Login' -Parameters @{ StateScript = $script:UserStateScript.ToString() } -ScriptBlock {
            $u = Get-ADUser -Identity $Target -Properties * @ad
            $st = & ([scriptblock]::Create($P.StateScript)) $u
            $manager = ''
            if ($u.Manager) { $manager = (([string]$u.Manager -replace '^CN=((?:\\.|[^,])+),.*$', '$1') -replace '\\(.)', '$1') }
            [pscustomobject][ordered]@{
                'Stan'                        = $st.State
                'Nazwa'                       = $u.DisplayName
                'UPN'                         = $u.UserPrincipalName
                'E-mail'                      = $u.mail
                'Stanowisko'                  = $u.Title
                'Dział'                       = $u.Department
                'Firma'                       = $u.Company
                'Biuro'                       = $u.physicalDeliveryOfficeName
                'Telefon'                     = $u.telephoneNumber
                'Komórka'                     = $u.mobile
                'Przełożony'                  = $manager
                'Włączone'                    = [bool]$u.Enabled
                'Zablokowane'                 = [bool]$u.LockedOut
                'Ostatnie logowanie'          = $u.LastLogonDate
                'Hasło ustawione'             = $u.PasswordLastSet
                'Hasło wygasa'                = $st.Expiry
                'Hasło nigdy nie wygasa'      = [bool]$u.PasswordNeverExpires
                'Zmiana hasła przy logowaniu' = ($null -ne $u.pwdLastSet -and [int64]$u.pwdLastSet -eq 0)
                'Błędne hasła'                = $u.badPwdCount
                'Konto wygasa'                = $u.AccountExpirationDate
                'Utworzono'                   = $u.whenCreated
                'Zmieniono'                   = $u.whenChanged
                'Grupy'                       = @($u.memberOf).Count
                'Opis'                        = $u.Description
                'Skrypt logowania'            = $u.ScriptPath
                'Folder domowy'               = $u.HomeDirectory
                'Profil mobilny'              = $u.ProfilePath
                'Jednostka OU'                = ($u.DistinguishedName -replace '^CN=(?:\\.|[^,])+,', '')
                'SID'                         = [string]$u.SID
                'DN'                          = $u.DistinguishedName
                '__tone'                      = $st.Tone
            }
        }
    } | Out-Null
    Add-RowAction -Module $m -Text 'Kopiuj DN' -Icon 'E8C8' -Action {
        param($m, $rows)
        $dns = @($rows | ForEach-Object { [string](Get-ObjectValue $_ 'DN') } | Where-Object { $_ })
        Set-ClipboardText ($dns -join [Environment]::NewLine)
        Show-Toast "Skopiowano DN: $($dns.Count)" 'ok'
    }
}

Register-Module -Workspace 'AdUsers' -Category 'Konta' -Key 'UserPassword' -Title 'Hasło i blokada' -Icon 'E8D7' `
    -Description 'Stan hasła i blokady, odblokowanie, reset hasła (losowe – inne dla każdego konta – albo wpisane), wymuszenie zmiany przy logowaniu i ustawienie «hasło nigdy nie wygasa».' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Stan', 'Wynik')
    $m.SecretColumns = @('Nowe hasło')
    $m.GoodWhenNo = @('Zablokowane', 'Hasło wygasło', 'Hasło nigdy nie wygasa')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetUsers)
        if (-not $targets) { return }
        Start-AdOperation -Module $m -Name 'Stan haseł' -Targets $targets -TargetColumn 'Login' -Parameters @{ StateScript = $script:UserStateScript.ToString() } -ScriptBlock {
            $u = Get-ADUser -Identity $Target -Properties DisplayName, Enabled, LockedOut, AccountLockoutTime, badPwdCount, LastBadPasswordAttempt, PasswordLastSet, PasswordExpired, PasswordNeverExpires, CannotChangePassword, pwdLastSet, 'msDS-UserPasswordExpiryTimeComputed', AccountExpirationDate @ad
            $st = & ([scriptblock]::Create($P.StateScript)) $u
            [pscustomobject][ordered]@{
                'Stan'                        = $st.State
                'Nazwa'                       = $u.DisplayName
                'Zablokowane'                 = [bool]$u.LockedOut
                'Zablokowano'                 = $u.AccountLockoutTime
                'Błędne hasła'                = $u.badPwdCount
                'Ostatnie błędne hasło'       = $u.LastBadPasswordAttempt
                'Hasło ustawione'             = $u.PasswordLastSet
                'Hasło wygasa'                = $st.Expiry
                'Hasło wygasło'               = [bool]$u.PasswordExpired
                'Hasło nigdy nie wygasa'      = [bool]$u.PasswordNeverExpires
                'Zmiana hasła przy logowaniu' = ($null -ne $u.pwdLastSet -and [int64]$u.pwdLastSet -eq 0)
                'Nie może zmienić hasła'      = [bool]$u.CannotChangePassword
                '__tone'                      = $st.Tone
            }
        }
    }
    $m.Actions.Simple = {
        param($m, [string]$Op, [string[]]$Logins)
        $targets = if ($Logins) { $Logins } else { @(Get-TargetUsers) }
        if (-not $targets) { return }
        $text = @{
            Unlock      = 'Odblokować zaznaczone konta?'
            MustChange  = 'Wymusić zmianę hasła przy następnym logowaniu?'
            NeverOn     = 'Ustawić «hasło nigdy nie wygasa»? Obniża to bezpieczeństwo – stosuj tylko dla kont usług.'
            NeverOff    = 'Wyłączyć «hasło nigdy nie wygasa»? Hasła mogą od razu wygasnąć, jeśli są starsze niż zasady domeny.'
        }[$Op]
        if (-not (Confirm-Action -Text $text -Items $targets -ConfirmText 'Wykonaj' -Danger:($Op -eq 'NeverOn'))) { return }
        Start-AdOperation -Module $m -Name "Hasła – $Op" -Targets $targets -TargetColumn 'Login' -Output Log -Parameters @{ Op = $Op } -ScriptBlock {
            switch ($P.Op) {
                'Unlock' { Unlock-ADAccount -Identity $Target @ad; 'Odblokowano.' }
                'MustChange' { Set-ADUser -Identity $Target -ChangePasswordAtLogon $true @ad; 'Wymuszono zmianę hasła przy logowaniu.' }
                'NeverOn' { Set-ADUser -Identity $Target -PasswordNeverExpires $true @ad; 'Hasło nigdy nie wygasa: włączone.' }
                'NeverOff' { Set-ADUser -Identity $Target -PasswordNeverExpires $false @ad; 'Hasło nigdy nie wygasa: wyłączone.' }
            }
        } -OnResult {
            param($m, $r)
            if ($r.Ok -and @($r.Data | Where-Object { [string]$_ -eq 'Odblokowano.' }).Count) { Update-UserRow -Login $r.Target -Values @{ Locked = $false } }
        } -OnComplete { param($m) & $m.Actions.List $m }
    }
    $m.Actions.Reset = {
        param($m)
        $targets = @(Get-TargetUsers)
        if (-not $targets) { return }
        $opt = Show-PasswordDialog -Subtitle ("Konta: {0}{1}" -f (($targets | Select-Object -First 5) -join ', '), $(if ($targets.Count -gt 5) { " i $($targets.Count - 5) więcej" } else { '' })) -Count $targets.Count
        if (-not $opt) { return }
        if (-not (Confirm-Action -Text "Zresetować hasło $($targets.Count) kont(a)?" -Items $targets -ConfirmText 'Resetuj hasła' -Danger)) { return }
        $per = @{}
        foreach ($t in $targets) {
            $per[$t] = @{
                Plain      = $(if ($opt.Mode -eq 'Generate') { New-RandomPassword -Length $opt.Length } else { '' })
                Secure     = $(if ($opt.Mode -eq 'Manual') { $opt.Password } else { $null })
                MustChange = $opt.MustChange
                Unlock     = $opt.Unlock
            }
        }
        Write-Log ("Reset hasła: {0} kont(a), tryb {1}, zmiana przy logowaniu: {2}" -f $targets.Count, $opt.Mode, $opt.MustChange)
        Start-AdOperation -Module $m -Name 'Reset hasła' -Targets $targets -PerTarget $per -TargetColumn 'Login' -ScriptBlock {
            $secure = if ($P.Plain) { ConvertTo-SecureString -String $P.Plain -AsPlainText -Force } else { $P.Secure }
            Set-ADAccountPassword -Identity $Target -Reset -NewPassword $secure @ad
            $notes = @()
            if ($P.Unlock) { try { Unlock-ADAccount -Identity $Target @ad; $notes += 'odblokowano' } catch { $notes += 'odblokowanie: ' + $_.Exception.Message } }
            $must = $false
            if ($P.MustChange) {
                try { Set-ADUser -Identity $Target -ChangePasswordAtLogon $true @ad; $must = $true }
                catch { $notes += 'zmiana przy logowaniu: ' + $_.Exception.Message }
            }
            [pscustomobject][ordered]@{
                'Wynik'                       = 'Hasło zresetowane'
                'Nowe hasło'                  = $(if ($P.Plain) { $P.Plain } else { '(wpisane ręcznie)' })
                'Zmiana hasła przy logowaniu' = $must
                'Uwagi'                       = ($notes -join '; ')
                '__tone'                      = 'ok'
            }
        } -OnResult {
            param($m, $r)
            if ($r.Ok) { Update-UserRow -Login $r.Target -Values @{ Locked = $false } }
        } -OnComplete {
            param($m)
            Show-Toast 'Hasła zresetowane. Zaznacz wiersz i użyj «Kopiuj hasło» (schowek czyści się po 60 s).' 'ok' 6
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Stan'
    Add-Button -Parent $row -Text 'Pokaż stan haseł' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Odblokuj' -Icon 'E785' -Module $m -OnClick { param($m) & $m.Actions.Simple $m 'Unlock' $null } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Hasło'
    Add-Button -Parent $row2 -Text 'Resetuj hasło…' -Icon 'E8D7' -Module $m -Danger -OnClick $m.Actions.Reset | Out-Null
    Add-Button -Parent $row2 -Text 'Wymuś zmianę przy logowaniu' -Icon 'E777' -Module $m -OnClick { param($m) & $m.Actions.Simple $m 'MustChange' $null } | Out-Null
    Add-Button -Parent $row2 -Text 'Nigdy nie wygasa: włącz' -Module $m -OnClick { param($m) & $m.Actions.Simple $m 'NeverOn' $null } | Out-Null
    Add-Button -Parent $row2 -Text 'wyłącz' -Module $m -OnClick { param($m) & $m.Actions.Simple $m 'NeverOff' $null } | Out-Null
    Add-RowAction -Module $m -Text 'Kopiuj hasło (60 s)' -Icon 'E8C8' -Action {
        param($m, $rows)
        $secret = [string](Get-RowValue $rows[0] 'Nowe hasło')
        if (-not $secret -or $secret -like '(*') { Show-Warning 'Wiersz nie zawiera wygenerowanego hasła.'; return }
        Set-ClipboardSecret -Text $secret -Seconds 60
        Write-Log ("Skopiowano nowe hasło konta {0}; schowek zostanie wyczyszczony po 60 s." -f (Get-ObjectValue $rows[0] 'Login')) 'OK'
        Show-Toast 'Hasło w schowku – zostanie wyczyszczone po 60 s.' 'ok'
    }
    Add-RowAction -Module $m -Text 'Odblokuj konto' -Icon 'E785' -Action {
        param($m, $rows)
        & $m.Actions.Simple $m 'Unlock' @($rows | ForEach-Object { [string](Get-ObjectValue $_ 'Login') })
    }
    $m.ResultHint = 'Nowe hasła są ukryte – «Pokaż poufne» albo «Kopiuj hasło» w menu wiersza'
}

Register-Module -Workspace 'AdUsers' -Category 'Konta' -Key 'UserState' -Title 'Stan konta' -Icon 'E8D8' `
    -Description 'Włączanie i wyłączanie kont (z adnotacją w opisie i przeniesieniem do OU), data wygaśnięcia konta oraz przenoszenie do innej jednostki organizacyjnej.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Stan')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetUsers)
        if (-not $targets) { return }
        Start-AdOperation -Module $m -Name 'Stan kont' -Targets $targets -TargetColumn 'Login' -Parameters @{ StateScript = $script:UserStateScript.ToString() } -ScriptBlock {
            $u = Get-ADUser -Identity $Target -Properties DisplayName, Enabled, LockedOut, AccountExpirationDate, Description, LastLogonDate, PasswordExpired, 'msDS-UserPasswordExpiryTimeComputed' @ad
            $st = & ([scriptblock]::Create($P.StateScript)) $u
            [pscustomobject][ordered]@{
                'Stan'               = $st.State
                'Nazwa'              = $u.DisplayName
                'Włączone'           = [bool]$u.Enabled
                'Konto wygasa'       = $u.AccountExpirationDate
                'Ostatnie logowanie' = $u.LastLogonDate
                'Opis'               = $u.Description
                'Jednostka OU'       = ($u.DistinguishedName -replace '^CN=(?:\\.|[^,])+,', '')
                '__tone'             = $st.Tone
            }
        }
    }
    $m.Actions.Change = {
        param($m, [string]$Op)
        $targets = @(Get-TargetUsers)
        if (-not $targets) { return }
        $params = @{ Op = $Op; Note = ''; MoveTo = ''; Date = $null }
        switch ($Op) {
            'Disable' {
                if (Test-Checked $m.Stamp) {
                    $reason = $m.Reason.Text.Trim()
                    $params.Note = ('Wyłączone {0:yyyy-MM-dd} przez {1}{2}' -f (Get-Date), $env:USERNAME, $(if ($reason) { ": $reason" } else { '' }))
                }
                $params.MoveTo = $m.DisabledOu.Text.Trim()
                $text = "Wyłączyć $($targets.Count) kont(a)?"
                if ($params.Note) { $text += "`r`nOpis zostanie poprzedzony: «$($params.Note)»" }
                if ($params.MoveTo) { $text += "`r`nKonta zostaną przeniesione do: $($params.MoveTo)" }
                if (-not (Confirm-Action -Text $text -Items $targets -ConfirmText 'Wyłącz konta' -Danger)) { return }
            }
            'Enable' { if (-not (Confirm-Action -Text "Włączyć $($targets.Count) kont(a)?" -Items $targets -ConfirmText 'Włącz konta')) { return } }
            'Expire' {
                $date = [datetime]::MinValue
                if (-not [datetime]::TryParseExact($m.ExpireDate.Text.Trim(), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, 'None', [ref]$date)) { Show-Warning 'Podaj datę w formacie RRRR-MM-DD, np. 2026-12-31.'; return }
                # Konto wygasa z końcem wskazanego dnia (AD przechowuje początek następnego dnia)
                $params.Date = $date.AddDays(1)
                if (-not (Confirm-Action -Text ("Ustawić wygaśnięcie kont z końcem dnia {0:yyyy-MM-dd}?" -f $date) -Items $targets -ConfirmText 'Ustaw')) { return }
            }
            'ClearExpire' { if (-not (Confirm-Action -Text 'Usunąć datę wygaśnięcia kont (konta nie będą wygasać)?' -Items $targets -ConfirmText 'Usuń datę')) { return } }
            'Move' {
                $ou = Select-OrganizationalUnit -Title 'Docelowa jednostka dla kont'
                if (-not $ou) { return }
                $params.MoveTo = $ou
                if (-not (Confirm-Action -Text "Przenieść konta do:`r`n$ou ?" -Items $targets -ConfirmText 'Przenieś')) { return }
            }
        }
        Start-AdOperation -Module $m -Name "Stan konta – $Op" -Targets $targets -TargetColumn 'Login' -Output Log -Parameters $params -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            $u = Get-ADUser -Identity $Target -Properties Description @ad
            switch ($P.Op) {
                'Disable' {
                    Disable-ADAccount -Identity $u @ad
                    if ($P.Note) {
                        $desc = if ($u.Description) { "$($P.Note) | $($u.Description)" } else { $P.Note }
                        if ($desc.Length -gt 1024) { $desc = $desc.Substring(0, 1024) }
                        Set-ADUser -Identity $u -Description $desc @ad
                    }
                    if ($P.MoveTo) { Move-ADObject -Identity $u.DistinguishedName -TargetPath $P.MoveTo @ad; "Wyłączono i przeniesiono do $($P.MoveTo)." }
                    else { 'Wyłączono.' }
                }
                'Enable' { Enable-ADAccount -Identity $u @ad; 'Włączono.' }
                'Expire' { Set-ADAccountExpiration -Identity $u -DateTime $P.Date @ad; ('Konto wygasa {0:yyyy-MM-dd HH:mm}.' -f $P.Date) }
                'ClearExpire' { Clear-ADAccountExpiration -Identity $u @ad; 'Usunięto datę wygaśnięcia.' }
                'Move' { Move-ADObject -Identity $u.DistinguishedName -TargetPath $P.MoveTo @ad; "Przeniesiono do $($P.MoveTo)." }
            }
        } -OnResult {
            param($m, $r)
            if (-not $r.Ok) { return }
            $text = (@($r.Data) -join ' ')
            if ($text -like 'Wyłączono*') { Update-UserRow -Login $r.Target -Values @{ Enabled = $false } }
            elseif ($text -like 'Włączono*') { Update-UserRow -Login $r.Target -Values @{ Enabled = $true } }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Stan'
    Add-Button -Parent $row -Text 'Pokaż stan kont' -Icon 'E72C' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Włącz' -Icon 'E73E' -Module $m -OnClick { param($m) & $m.Actions.Change $m 'Enable' } | Out-Null
    Add-Button -Parent $row -Text 'Przenieś do OU…' -Icon 'E8DE' -Module $m -OnClick { param($m) & $m.Actions.Change $m 'Move' } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Wyłączanie'
    $m.Stamp = Add-CheckBox -Parent $row2 -Text 'Dopisz do opisu datę i autora' -Checked $true
    $m.Reason = Add-TextBox -Parent $row2 -Width 220 -Placeholder 'powód, np. numer zgłoszenia'
    $m.DisabledOu = Add-TextBox -Parent $row2 -Width 220 -Placeholder 'przenieś do OU (opcjonalnie)'
    Add-Button -Parent $row2 -Text '' -Icon 'E8B7' -Module $m -AlwaysEnabled -ToolTip 'Wybierz OU dla wyłączonych kont' -OnClick {
        param($m)
        $ou = Select-OrganizationalUnit -Title 'OU dla wyłączonych kont' -Selected $m.DisabledOu.Text.Trim()
        if ($ou) { $m.DisabledOu.Text = $ou }
    } | Out-Null
    Add-Button -Parent $row2 -Text 'Wyłącz konta' -Icon 'E8D8' -Module $m -Danger -OnClick { param($m) & $m.Actions.Change $m 'Disable' } | Out-Null
    $row3 = Add-ToolbarRow -Module $m -Title 'Wygaśnięcie'
    $m.ExpireDate = Add-TextBox -Parent $row3 -Width 130 -Text ((Get-Date).AddDays(30).ToString('yyyy-MM-dd')) -Placeholder 'RRRR-MM-DD'
    Add-Button -Parent $row3 -Text 'Ustaw datę wygaśnięcia' -Icon 'E787' -Module $m -OnClick { param($m) & $m.Actions.Change $m 'Expire' } | Out-Null
    Add-Button -Parent $row3 -Text 'Usuń datę wygaśnięcia' -Icon 'E711' -Module $m -OnClick { param($m) & $m.Actions.Change $m 'ClearExpire' } | Out-Null
}

$script:UserAttributes = @(
    @{ Name = 'title'; Label = 'Stanowisko (title)' }
    @{ Name = 'department'; Label = 'Dział (department)' }
    @{ Name = 'company'; Label = 'Firma (company)' }
    @{ Name = 'physicalDeliveryOfficeName'; Label = 'Biuro (physicalDeliveryOfficeName)' }
    @{ Name = 'description'; Label = 'Opis (description)' }
    @{ Name = 'telephoneNumber'; Label = 'Telefon (telephoneNumber)' }
    @{ Name = 'mobile'; Label = 'Komórka (mobile)' }
    @{ Name = 'mail'; Label = 'E-mail (mail)' }
    @{ Name = 'manager'; Label = 'Przełożony (manager – login)' }
    @{ Name = 'employeeID'; Label = 'Numer pracownika (employeeID)' }
    @{ Name = 'l'; Label = 'Miasto (l)' }
    @{ Name = 'streetAddress'; Label = 'Ulica (streetAddress)' }
    @{ Name = 'postalCode'; Label = 'Kod pocztowy (postalCode)' }
    @{ Name = 'wWWHomePage'; Label = 'Strona WWW (wWWHomePage)' }
    @{ Name = 'info'; Label = 'Uwagi (info)' }
    @{ Name = 'scriptPath'; Label = 'Skrypt logowania (scriptPath)' }
) + @(1..15 | ForEach-Object { @{ Name = "extensionAttribute$_"; Label = "extensionAttribute$_" } })

Register-Module -Workspace 'AdUsers' -Category 'Konta' -Key 'UserAttributes' -Title 'Edycja atrybutów' -Icon 'E70F' `
    -Description 'Odczyt i masowa zmiana atrybutów zaznaczonych kont. Wartość może zawierać pola {login}, {imie}, {nazwisko}, {nazwa} – np. {login}@firma.pl.' -Build {
    param($m)
    $row = Add-ToolbarRow -Module $m -Title 'Atrybut'
    $m.Attr = Add-ComboBox -Parent $row -Items @($script:UserAttributes | ForEach-Object { $_.Label }) -Width 320
    Add-Button -Parent $row -Text 'Pokaż wartości' -Icon 'E72C' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-TargetUsers)
        if (-not $targets) { return }
        $attr = $script:UserAttributes[$m.Attr.SelectedIndex].Name
        Start-AdOperation -Module $m -Name "Atrybut $attr" -Targets $targets -TargetColumn 'Login' -Parameters @{ Attr = $attr } -ScriptBlock {
            $u = Get-ADUser -Identity $Target -Properties $P.Attr, DisplayName @ad
            $v = $u.($P.Attr)
            if ($P.Attr -eq 'manager' -and $v) { $v = ([string]$v -replace '^CN=((?:\\.|[^,])+),.*$', '$1') -replace '\\(.)', '$1' }
            [pscustomobject]@{ 'Nazwa' = $u.DisplayName; 'Atrybut' = $P.Attr; 'Wartość' = $(if ($null -eq $v) { '' } else { (@($v) -join '; ') }) }
        }
    } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Nowa wartość'
    $m.Value = Add-TextBox -Parent $row2 -Width 380 -Placeholder 'wartość lub szablon, np. {login}@firma.pl'
    Add-Button -Parent $row2 -Text 'Ustaw' -Icon 'E74E' -Module $m -OnClick {
        param($m)
        $targets = @(Get-TargetUsers)
        if (-not $targets) { return }
        $attr = $script:UserAttributes[$m.Attr.SelectedIndex].Name
        $value = $m.Value.Text.Trim()
        if (-not $value) { Show-Warning 'Wpisz wartość (aby usunąć wartość, użyj «Wyczyść»).'; return }
        if (-not (Confirm-Action -Text "Ustawić atrybut $attr = «$value» dla $($targets.Count) kont(a)?" -Items $targets -ConfirmText 'Ustaw')) { return }
        Start-AdOperation -Module $m -Name "Zmiana $attr" -Targets $targets -TargetColumn 'Login' -Output Log -Parameters @{ Attr = $attr; Value = $value } -ScriptBlock {
            $u = Get-ADUser -Identity $Target -Properties GivenName, Surname, DisplayName @ad
            $v = $P.Value.Replace('{login}', $u.SamAccountName).Replace('{imie}', [string]$u.GivenName).Replace('{nazwisko}', [string]$u.Surname).Replace('{nazwa}', [string]$u.DisplayName)
            if ($P.Attr -eq 'manager') { $v = (Get-ADUser -Identity $v @ad).DistinguishedName }
            Set-ADUser -Identity $u -Replace @{ $P.Attr = $v } @ad
            "$($P.Attr) = $v"
        } -OnComplete { param($m) Show-Toast 'Zmieniono atrybuty – «Pokaż wartości», aby sprawdzić.' 'ok' }
    } | Out-Null
    Add-Button -Parent $row2 -Text 'Wyczyść' -Icon 'E894' -Module $m -Danger -OnClick {
        param($m)
        $targets = @(Get-TargetUsers)
        if (-not $targets) { return }
        $attr = $script:UserAttributes[$m.Attr.SelectedIndex].Name
        if (-not (Confirm-Action -Text "Wyczyścić atrybut $attr dla $($targets.Count) kont(a)?" -Items $targets -ConfirmText 'Wyczyść' -Danger)) { return }
        Start-AdOperation -Module $m -Name "Czyszczenie $attr" -Targets $targets -TargetColumn 'Login' -Output Log -Parameters @{ Attr = $attr } -ScriptBlock {
            Set-ADUser -Identity $Target -Clear $P.Attr @ad
            "Wyczyszczono $($P.Attr)."
        }
    } | Out-Null
}

Register-GroupMembershipModule -Workspace 'AdUsers' -Kind User -Key 'UserGroups' -Category 'Grupy'

$script:UserReports = @(
    @{ Key = 'Locked'; Name = 'Zablokowane konta' }
    @{ Key = 'Disabled'; Name = 'Wyłączone konta' }
    @{ Key = 'Inactive'; Name = 'Nieaktywne (brak logowania od N dni)' }
    @{ Key = 'NeverLogged'; Name = 'Nigdy nie logowane' }
    @{ Key = 'PwdExpiring'; Name = 'Hasło wygasa w ciągu N dni' }
    @{ Key = 'PwdExpired'; Name = 'Hasło wygasło' }
    @{ Key = 'PwdNever'; Name = 'Hasło nigdy nie wygasa' }
    @{ Key = 'AccExpiring'; Name = 'Konto wygasło lub wygasa w ciągu N dni' }
    @{ Key = 'Created'; Name = 'Utworzone w ciągu N dni' }
    @{ Key = 'Privileged'; Name = 'Konta uprzywilejowane (adminCount = 1)' }
)

Register-Module -Workspace 'AdUsers' -Category 'Raporty' -Key 'UserReports' -Title 'Raporty kont' -Icon 'E9F9' `
    -Description 'Zestawienia z całej domeny lub jednostki: zablokowane, wyłączone, nieaktywne, z wygasającym hasłem, wygasające, nowe i uprzywilejowane. Wyniki można przenieść na listę kont i wykonać na nich operacje.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Stan')
    $m.GoodWhenNo = @('Zablokowane', 'Hasło nigdy nie wygasa')
    $row = Add-ToolbarRow -Module $m -Title 'Raport'
    $m.Report = Add-ComboBox -Parent $row -Items @($script:UserReports | ForEach-Object { $_.Name }) -Width 340
    Add-Label -Parent $row -Text 'N (dni)' | Out-Null
    $m.Days = Add-Numeric -Parent $row -Value ([int]$script:Settings.InactiveDays) -Minimum 1 -Maximum 3650 -Width 70
    $m.OnlyEnabled = Add-CheckBox -Parent $row -Text 'Tylko włączone' -Checked $true -ToolTip 'Dotyczy raportów nieaktywności, haseł i nowych kont'
    $row2 = Add-ToolbarRow -Module $m -Title 'Zakres'
    $m.Ou = Add-TextBox -Parent $row2 -Width 420 -Placeholder 'Cała domena'
    Add-Button -Parent $row2 -Text '' -Icon 'E8B7' -Module $m -AlwaysEnabled -ToolTip 'Wybierz jednostkę organizacyjną' -OnClick {
        param($m)
        $ou = Select-OrganizationalUnit -Title 'Zakres raportu' -Selected $m.Ou.Text.Trim() -AllowDomainRoot
        if ($null -ne $ou) { $m.Ou.Text = $ou }
    } | Out-Null
    Add-Button -Parent $row2 -Text 'Generuj raport' -Icon 'E9F9' -Module $m -Primary -OnClick {
        param($m)
        $report = $script:UserReports[$m.Report.SelectedIndex]
        $params = @{ Report = $report.Key; Days = (Get-Num $m.Days); OnlyEnabled = (Test-Checked $m.OnlyEnabled); SearchBase = $m.Ou.Text.Trim(); StateScript = $script:UserStateScript.ToString() }
        Start-AdOperation -Module $m -Name $report.Name -Targets @('AD') -Parameters $params -ScriptBlock {
            $now = Get-Date
            $days = [int]$P.Days
            $enabledOnly = '(!(userAccountControl:1.2.840.113556.1.4.803:=2))'
            $base = '(objectCategory=person)(objectClass=user)'
            $cut = $now.AddDays(-$days)
            $ft = $cut.ToFileTimeUtc()
            $gen = $cut.ToUniversalTime().ToString('yyyyMMddHHmmss.0Z')
            $extra = switch ($P.Report) {
                'Locked' { '(lockoutTime>=1)' }
                'Disabled' { '(userAccountControl:1.2.840.113556.1.4.803:=2)' }
                'Inactive' { "(|(lastLogonTimestamp<=$ft)(&(!(lastLogonTimestamp=*))(whenCreated<=$gen)))" }
                'NeverLogged' { '(!(lastLogonTimestamp=*))' }
                'PwdExpiring' { '(!(userAccountControl:1.2.840.113556.1.4.803:=65536))' }
                'PwdExpired' { '(!(userAccountControl:1.2.840.113556.1.4.803:=65536))' }
                'PwdNever' { '(userAccountControl:1.2.840.113556.1.4.803:=65536)' }
                'AccExpiring' { '(accountExpires>=1)(!(accountExpires=9223372036854775807))' }
                'Created' { "(whenCreated>=$gen)" }
                'Privileged' { '(adminCount=1)' }
            }
            if ($P.OnlyEnabled -and @('Locked', 'Disabled', 'Privileged') -notcontains $P.Report) { $extra += $enabledOnly }
            $q = @{
                LDAPFilter = "(&$base$extra)"
                Properties = @('DisplayName', 'Enabled', 'LockedOut', 'LastLogonDate', 'PasswordLastSet', 'PasswordNeverExpires', 'PasswordExpired', 'msDS-UserPasswordExpiryTimeComputed', 'AccountExpirationDate', 'whenCreated', 'Department', 'Title', 'mail', 'Description', 'AccountLockoutTime')
            }
            if ($P.SearchBase) { $q.SearchBase = $P.SearchBase }
            $users = @(Get-ADUser @q @ad)
            $stateScript = [scriptblock]::Create($P.StateScript)
            foreach ($u in $users) {
                $st = & $stateScript $u
                # Uwaga: "continue" wewnątrz switch dotyczy switch, nie pętli - stąd osobna zmienna
                $include = switch ($P.Report) {
                    'Locked' { [bool]$u.LockedOut }
                    'PwdExpiring' { [bool]($st.Expiry -and $st.Expiry -ge $now -and $st.Expiry -le $now.AddDays($days)) }
                    'PwdExpired' { [bool]$u.PasswordExpired }
                    'AccExpiring' { [bool]($u.AccountExpirationDate -and $u.AccountExpirationDate -le $now.AddDays($days)) }
                    default { $true }
                }
                if (-not $include) { continue }
                [pscustomobject][ordered]@{
                    'Login'              = $u.SamAccountName
                    'Nazwa'              = $u.DisplayName
                    'Stan'               = $st.State
                    'Włączone'           = [bool]$u.Enabled
                    'Zablokowane'        = [bool]$u.LockedOut
                    'Ostatnie logowanie' = $u.LastLogonDate
                    'Dni bez logowania'  = $(if ($u.LastLogonDate) { [int]($now - $u.LastLogonDate).TotalDays } else { $null })
                    'Hasło ustawione'    = $u.PasswordLastSet
                    'Hasło wygasa'       = $st.Expiry
                    'Hasło nigdy nie wygasa' = [bool]$u.PasswordNeverExpires
                    'Konto wygasa'       = $u.AccountExpirationDate
                    'Utworzono'          = $u.whenCreated
                    'Dział'              = $u.Department
                    'Stanowisko'         = $u.Title
                    'E-mail'             = $u.mail
                    'Opis'               = $u.Description
                    'DN'                 = $u.DistinguishedName
                    '__tone'             = $st.Tone
                }
            }
        } -OnComplete {
            param($m)
            Set-StatTile -Module $m -Key 'count' -Value ([string]$m.Table.Rows.Count) -Tone $(if ($m.Table.Rows.Count) { 'warn' } else { 'ok' })
            Set-StatTile -Module $m -Key 'locked' -Value ([string]@($m.Table.Rows | Where-Object { $m.Table.Columns.Contains('Zablokowane') -and [string]$_['Zablokowane'] -eq 'Tak' }).Count)
            Set-StatTile -Module $m -Key 'disabled' -Value ([string]@($m.Table.Rows | Where-Object { $m.Table.Columns.Contains('Włączone') -and [string]$_['Włączone'] -eq 'Nie' }).Count)
        }
    } | Out-Null
    $row3 = Add-ToolbarRow -Module $m -Title 'Wyniki'
    Add-Button -Parent $row3 -Text 'Zaznacz na liście kont' -Icon 'E8B3' -Module $m -AlwaysEnabled -ToolTip 'Zaznaczone wiersze (albo wszystkie widoczne) trafią na listę kont po lewej – zaznaczone do dalszych operacji' -OnClick {
        param($m)
        Add-ResultsToTargets -Module $m -Kind User -Column 'Login'
    } | Out-Null
    Add-RowAction -Module $m -Text 'Zaznacz na liście kont' -Icon 'E8B3' -Action { param($m, $rows) Add-ResultsToTargets -Module $m -Kind User -Column 'Login' -Rows $rows }
    Add-RowAction -Module $m -Text 'Odblokuj' -Icon 'E785' -Separator -Action {
        param($m, $rows)
        $logins = @($rows | ForEach-Object { [string](Get-ObjectValue $_ 'Login') } | Where-Object { $_ })
        if (-not (Confirm-Action -Text 'Odblokować wybrane konta?' -Items $logins -ConfirmText 'Odblokuj')) { return }
        Start-AdOperation -Module $m -Name 'Odblokowanie' -Targets $logins -TargetColumn 'Login' -Output Log -ScriptBlock { Unlock-ADAccount -Identity $Target @ad; 'Odblokowano.' }
    }
    Add-RowAction -Module $m -Text 'Wyłącz konto' -Icon 'E8D8' -Danger -Action {
        param($m, $rows)
        $logins = @($rows | ForEach-Object { [string](Get-ObjectValue $_ 'Login') } | Where-Object { $_ })
        if (-not (Confirm-Action -Text 'Wyłączyć wybrane konta?' -Items $logins -ConfirmText 'Wyłącz' -Danger)) { return }
        Start-AdOperation -Module $m -Name 'Wyłączanie kont' -Targets $logins -TargetColumn 'Login' -Output Log -ScriptBlock { Disable-ADAccount -Identity $Target @ad; 'Wyłączono.' }
    }
    Add-StatTile -Module $m -Key 'count' -Label 'Konta w raporcie' -Icon 'E716' | Out-Null
    Add-StatTile -Module $m -Key 'locked' -Label 'Zablokowane' -Icon 'E72E' | Out-Null
    Add-StatTile -Module $m -Key 'disabled' -Label 'Wyłączone' -Icon 'E8D8' | Out-Null
    $m.EmptyHint = 'Wybierz raport i kliknij «Generuj raport» (F5). Nie trzeba zaznaczać kont na liście.'
}

Register-Module -Workspace 'AdUsers' -Category 'Raporty' -Key 'LockoutSource' -Title 'Źródło blokady konta' -Icon 'E7BA' `
    -Description 'Skąd pochodzą blokady: zdarzenia 4740 z kontrolerów domeny (komputer, z którego przyszły błędne hasła) oraz opcjonalnie nieudane uwierzytelnienia Kerberos 4771 i NTLM 4776. Bez zaznaczonych kont – wszystkie blokady.' -Build {
    param($m)
    $m.PillColumns = @('Zdarzenie')
    $row = Add-ToolbarRow -Module $m -Title 'Parametry'
    Add-Label -Parent $row -Text 'Ostatnie godziny' | Out-Null
    $m.Hours = Add-Numeric -Parent $row -Value 24 -Minimum 1 -Maximum 720 -Width 70
    $m.AllDcs = Add-CheckBox -Parent $row -Text 'Wszystkie kontrolery domeny' -ToolTip 'Domyślnie tylko emulator PDC (tam trafiają wszystkie blokady). Wszystkie DC – także zdarzenia 4771/4776 z innych kontrolerów.'
    $m.Failures = Add-CheckBox -Parent $row -Text 'Nieudane uwierzytelnienia (4771, 4776)'
    $m.Max = Add-Numeric -Parent $row -Value 500 -Minimum 10 -Maximum 10000 -Width 70
    $m.Max.ToolTip = 'Maksymalna liczba zdarzeń z jednego kontrolera'
    Add-Button -Parent $row -Text 'Szukaj źródła' -Icon 'E721' -Module $m -Primary -OnClick {
        param($m)
        if (-not (Test-AdAvailable)) { return }
        Import-AdModule
        $ad = Get-AdSplat
        $info = Invoke-WithWaitCursor {
            $domain = Get-ADDomain @ad
            $list = if (Test-Checked $m.AllDcs) { @(Get-ADDomainController -Filter * @ad | ForEach-Object { $_.HostName }) } else { @($domain.PDCEmulator) }
            @{ Dcs = $list }
        }
        $users = @(Get-TargetUsers -Quiet | Where-Object { $_ -notmatch "['""]" })
        $params = @{ Users = $users; Hours = (Get-Num $m.Hours); Failures = (Test-Checked $m.Failures); Max = (Get-Num $m.Max) }
        $scope = if ($users.Count) { "kont: $($users.Count)" } else { 'wszystkich kont' }
        Write-Log ("Szukanie blokad ({0}) na: {1}" -f $scope, ($info.Dcs -join ', '))
        Start-HostOperation -Module $m -Name 'Źródło blokady' -Targets $info.Dcs -Local -TargetColumn 'Kontroler domeny' -Parameters $params -ScriptBlock {
            param($Target, $P, $Ctx)
            $ids = @(4740)
            if ($P.Failures) { $ids += 4771; $ids += 4776 }
            $idFilter = (@($ids | ForEach-Object { "EventID=$_" }) -join ' or ')
            $ms = [int64]$P.Hours * 3600000
            $xpath = "*[System[($idFilter) and TimeCreated[timediff(@SystemTime) <= $ms]]"
            if (@($P.Users).Count -gt 0) { $xpath += ' and EventData[' + ((@($P.Users | ForEach-Object { "Data[@Name='TargetUserName']='$_'" })) -join ' or ') + ']' }
            $xpath += ']'
            $query = {
                param($XPath, $Max)
                $events = @()
                try { $events = @(Get-WinEvent -LogName Security -FilterXPath $XPath -MaxEvents $Max -ErrorAction Stop) }
                catch { if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') { throw } }
                foreach ($e in $events) {
                    $data = @{}
                    foreach ($d in ([xml]$e.ToXml()).Event.EventData.Data) { $data[[string]$d.Name] = [string]$d.'#text' }
                    [pscustomobject]@{ Time = $e.TimeCreated; Id = $e.Id; Data = $data }
                }
            }
            $ic = @{ ComputerName = $Target; ScriptBlock = $query; ArgumentList = @($xpath, [int]$P.Max); ErrorAction = 'Stop' }
            if ($Ctx.Credential) { $ic.Credential = $Ctx.Credential }
            if ($Ctx.SessionOption) { $ic.SessionOption = $Ctx.SessionOption }
            $events = @()
            try { $events = @(Invoke-Command @ic) }
            catch {
                # Bez WinRM na DC - odczyt zdalnego dziennika przez RPC
                $gp = @{ ComputerName = $Target; LogName = 'Security'; FilterXPath = $xpath; MaxEvents = [int]$P.Max; ErrorAction = 'Stop' }
                if ($Ctx.Credential) { $gp.Credential = $Ctx.Credential }
                try {
                    $events = @(Get-WinEvent @gp | ForEach-Object {
                            $data = @{}
                            foreach ($d in ([xml]$_.ToXml()).Event.EventData.Data) { $data[[string]$d.Name] = [string]$d.'#text' }
                            [pscustomobject]@{ Time = $_.TimeCreated; Id = $_.Id; Data = $data }
                        })
                }
                catch { if ($_.FullyQualifiedErrorId -notlike 'NoMatchingEventsFound*') { throw } }
            }
            $dns = @{}
            foreach ($e in ($events | Sort-Object Time -Descending)) {
                $d = $e.Data
                if ([int]$e.Id -eq 4776 -and [string]$d['Status'] -eq '0x0') { continue }
                $source = ''
                $detail = ''
                $tone = 'info'
                switch ([int]$e.Id) {
                    4740 { $kind = 'Blokada konta'; $source = $d['TargetDomainName']; $tone = 'crit' }
                    4771 {
                        $kind = 'Kerberos – błąd uwierzytelnienia'
                        $ip = ([string]$d['IpAddress']) -replace '^::ffff:', ''
                        $source = $ip
                        if ($ip -and $ip -ne '::1' -and $ip -ne '-') {
                            if (-not $dns.ContainsKey($ip)) { $dns[$ip] = ''; try { $dns[$ip] = [System.Net.Dns]::GetHostEntry($ip).HostName } catch { } }
                            if ($dns[$ip]) { $source = '{0} ({1})' -f $dns[$ip], $ip }
                        }
                        $detail = switch ([string]$d['Status']) { '0x18' { 'Błędne hasło' } '0x12' { 'Konto zablokowane / wyłączone' } '0x17' { 'Hasło wygasło' } default { "Kod $($d['Status'])" } }
                        $tone = 'warn'
                    }
                    4776 {
                        $kind = 'NTLM – błąd uwierzytelnienia'
                        $source = $d['Workstation']
                        $detail = switch ([string]$d['Status']) { '0xc000006a' { 'Błędne hasło' } '0xc0000234' { 'Konto zablokowane' } '0xc0000064' { 'Nieznane konto' } '0xc0000072' { 'Konto wyłączone' } default { "Kod $($d['Status'])" } }
                        $tone = 'warn'
                    }
                }
                [pscustomobject][ordered]@{
                    'Czas'        = $e.Time
                    'Zdarzenie'   = $kind
                    'Konto'       = $d['TargetUserName']
                    'Źródło'      = $source
                    'Szczegóły'   = $detail
                    'ID'          = $e.Id
                    '__tone'      = $tone
                }
            }
        } -OnComplete {
            param($m)
            $sources = @($m.Table.Rows | Where-Object { $m.Table.Columns.Contains('Źródło') -and [string]$_['Źródło'] } | Group-Object { [string]$_['Źródło'] } | Sort-Object Count -Descending)
            if ($sources.Count -gt 0) { Show-Toast ("Najczęstsze źródło: {0} ({1} zdarzeń)" -f $sources[0].Name, $sources[0].Count) 'info' 8 }
        }
    } | Out-Null
    $m.EmptyHint = 'Zaznacz konta po lewej (albo żadnego – wszystkie blokady) i kliknij «Szukaj źródła». Wymaga uprawnień do dziennika Security na kontrolerach domeny.'
    Add-RowAction -Module $m -Text 'Odblokuj konto' -Icon 'E785' -Action {
        param($m, $rows)
        $logins = @($rows | ForEach-Object { [string](Get-ObjectValue $_ 'Konto') } | Where-Object { $_ } | Select-Object -Unique)
        if (-not (Confirm-Action -Text 'Odblokować wybrane konta?' -Items $logins -ConfirmText 'Odblokuj')) { return }
        Start-AdOperation -Module $m -Name 'Odblokowanie' -Targets $logins -TargetColumn 'Login' -Output Log -ScriptBlock { Unlock-ADAccount -Identity $Target @ad; 'Odblokowano.' }
    }
}
#endregion

#region Komputery AD
function Get-AdComputerTargets {
    # Nazwy komputerów z listy po lewej bez części domenowej (FQDN -> nazwa)
    return @(Get-TargetComputers | ForEach-Object { ($_ -split '\.')[0] } | Select-Object -Unique)
}

Register-Module -Workspace 'AdComputers' -Category 'Konta komputerów' -Key 'ComputerAccount' -Title 'Konto komputera' -Icon 'E977' `
    -Description 'Obiekt komputera w AD: informacje, kanał zaufania (test i naprawa), włączanie, wyłączanie, opis, przenoszenie do OU, reset i usuwanie konta oraz tworzenie nowych kont (pre-staging).' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Stan')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-AdComputerTargets)
        if (-not $targets) { return }
        Start-AdOperation -Module $m -Name 'Konto komputera – informacje' -Targets $targets -TargetColumn 'Komputer' -ScriptBlock {
            $props = @('Enabled', 'LastLogonDate', 'PasswordLastSet', 'OperatingSystem', 'OperatingSystemVersion', 'whenCreated', 'Description', 'IPv4Address', 'ManagedBy', 'Location', 'DNSHostName')
            $c = Get-ADComputer -Identity $Target -Properties $props @ad
            $days = if ($c.LastLogonDate) { [int]((Get-Date) - $c.LastLogonDate).TotalDays } else { $null }
            $state = if (-not $c.Enabled) { 'Wyłączone' } elseif ($null -eq $days) { 'Nigdy nie logowane' } elseif ($days -gt 90) { "Nieaktywne ($days dni)" } else { 'Aktywne' }
            [pscustomobject][ordered]@{
                'Stan'                  = $state
                'Włączone'              = [bool]$c.Enabled
                'Ostatnie logowanie'    = $c.LastLogonDate
                'Dni bez logowania'     = $days
                'Hasło konta zmienione' = $c.PasswordLastSet
                'System'                = $c.OperatingSystem
                'Wersja'                = $c.OperatingSystemVersion
                'IPv4'                  = $c.IPv4Address
                'DNS'                   = $c.DNSHostName
                'Utworzono'             = $c.whenCreated
                'Opis'                  = $c.Description
                'Lokalizacja'           = $c.Location
                'Jednostka OU'          = ($c.DistinguishedName -replace '^CN=(?:\\.|[^,])+,', '')
                'DN'                    = $c.DistinguishedName
                '__tone'                = $(if (-not $c.Enabled) { '' } elseif ($state -eq 'Aktywne') { 'ok' } else { 'warn' })
            }
        }
    }
    $m.Actions.Change = {
        param($m, [string]$Op)
        $targets = @(Get-AdComputerTargets)
        if (-not $targets) { return }
        $params = @{ Op = $Op; TargetOU = ''; Description = '' }
        $danger = $false
        switch ($Op) {
            'Reset' { $question = 'Zresetować konta komputerów w AD? Komputery stracą relację zaufania i trzeba będzie ponownie dołączyć je do domeny lub naprawić kanał.'; $danger = $true }
            'Enable' { $question = 'Włączyć konta komputerów w AD?' }
            'Disable' { $question = 'Wyłączyć konta komputerów w AD? Użytkownicy nie zalogują się kontem domenowym na tych komputerach.'; $danger = $true }
            'Delete' { $question = 'USUNĄĆ konta komputerów z AD (razem z obiektami podrzędnymi, np. kluczami odzyskiwania BitLocker)? Tej operacji nie można cofnąć bez kosza AD.'; $danger = $true }
            'Describe' {
                $params.Description = $m.Description.Text.Trim()
                $question = "Ustawić opis «$($params.Description)»?"
            }
            'Move' {
                $ou = Select-OrganizationalUnit -Title 'Docelowa jednostka organizacyjna'
                if (-not $ou) { return }
                $params.TargetOU = $ou
                $question = "Przenieść konta komputerów do:`r`n$ou ?"
            }
        }
        if (-not (Confirm-Action -Text $question -Items $targets -ConfirmText $(if ($Op -eq 'Delete') { 'Usuń konta' } else { 'Wykonaj' }) -Danger:$danger)) { return }
        $m.Data.LastOp = $Op
        Start-AdOperation -Module $m -Name "Konto komputera – $Op" -Targets $targets -TargetColumn 'Komputer' -Output Log -Parameters $params -OnComplete { param($m) if ($m.Data.LastOp -ne 'Delete') { & $m.Actions.List $m } } -ScriptBlock {
            $c = Get-ADComputer -Identity $Target @ad
            switch ($P.Op) {
                'Reset' {
                    # Odpowiednik "Resetuj konto" z konsoli ADUC: hasło = nazwa komputera małymi literami (bez $, maks. 14 znaków)
                    $plain = $c.SamAccountName.TrimEnd('$').ToLowerInvariant()
                    if ($plain.Length -gt 14) { $plain = $plain.Substring(0, 14) }
                    Set-ADAccountPassword -Identity $c -Reset -NewPassword (ConvertTo-SecureString $plain -AsPlainText -Force) @ad
                    'Konto zresetowane – dołącz komputer ponownie do domeny lub napraw kanał zaufania.'
                }
                'Enable' { Enable-ADAccount -Identity $c @ad; 'Konto włączone.' }
                'Disable' { Disable-ADAccount -Identity $c @ad; 'Konto wyłączone.' }
                'Move' { Move-ADObject -Identity $c.DistinguishedName -TargetPath $P.TargetOU @ad; "Przeniesiono do $($P.TargetOU)." }
                'Describe' {
                    if ($P.Description) { Set-ADComputer -Identity $c -Description $P.Description @ad } else { Set-ADComputer -Identity $c -Clear description @ad }
                    'Opis ustawiony.'
                }
                'Delete' { Remove-ADObject -Identity $c.DistinguishedName -Recursive -Confirm:$false @ad; 'Konto usunięte z AD.' }
            }
        }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Informacje'
    Add-Button -Parent $row -Text 'Informacje z AD' -Icon 'E946' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Test kanału zaufania' -Icon 'E9D9' -Module $m -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Test kanału zaufania' -Targets $targets -Parameters @{ Server = [string]$script:Settings.DomainController } -ScriptBlock {
            param($P)
            $tp = @{ ErrorAction = 'Stop' }
            if ($P.Server) { $tp.Server = $P.Server }
            $ok = Test-ComputerSecureChannel @tp
            $domain = (Get-CimInstance -ClassName Win32_ComputerSystem).Domain
            $dc = ''
            try {
                $text = (& nltest.exe "/sc_query:$domain" 2>&1) -join "`n"
                if ($text -match '\\\\([^\s\\]+)') { $dc = $Matches[1] }
            }
            catch { }
            [pscustomobject]@{
                'Stan'             = $(if ($ok) { 'Kanał poprawny' } else { 'Kanał uszkodzony' })
                'Domena'           = $domain
                'Kontroler domeny' = $dc
                '__tone'           = $(if ($ok) { 'ok' } else { 'crit' })
            }
        }
    } | Out-Null
    Add-Button -Parent $row -Text 'Napraw kanał zaufania' -Icon 'E90F' -Module $m -Danger -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $cred = Get-EffectiveCredential
        if (-not $cred) {
            $cred = Show-CredentialDialog -Message 'Naprawa kanału zaufania wymaga poświadczeń domenowych z prawem resetu konta komputera (nie przechodzą przez WinRM automatycznie).'
            if (-not $cred) { return }
        }
        if (-not (Confirm-Action -Text "Naprawić kanał zaufania (Test-ComputerSecureChannel -Repair)?" -Items $targets -ConfirmText 'Napraw')) { return }
        Start-HostOperation -Module $m -Name 'Naprawa kanału zaufania' -Targets $targets -Output Log -Parameters @{ Credential = $cred; Server = [string]$script:Settings.DomainController } -ScriptBlock {
            param($P)
            $tp = @{ Repair = $true; Credential = $P.Credential; ErrorAction = 'Stop' }
            if ($P.Server) { $tp.Server = $P.Server }
            if (Test-ComputerSecureChannel @tp) { 'Kanał zaufania naprawiony.' } else { 'Błąd – naprawa nie powiodła się.' }
        }
    } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Konto w AD'
    Add-Button -Parent $row2 -Text 'Włącz' -Icon 'E73E' -Module $m -OnClick { param($m) & $m.Actions.Change $m 'Enable' } | Out-Null
    Add-Button -Parent $row2 -Text 'Wyłącz' -Icon 'E8D8' -Module $m -Danger -OnClick { param($m) & $m.Actions.Change $m 'Disable' } | Out-Null
    Add-Button -Parent $row2 -Text 'Przenieś do OU…' -Icon 'E8DE' -Module $m -OnClick { param($m) & $m.Actions.Change $m 'Move' } | Out-Null
    Add-Button -Parent $row2 -Text 'Resetuj konto' -Icon 'E777' -Module $m -Danger -OnClick { param($m) & $m.Actions.Change $m 'Reset' } | Out-Null
    Add-Button -Parent $row2 -Text 'Usuń z AD' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Change $m 'Delete' } | Out-Null
    $row3 = Add-ToolbarRow -Module $m -Title 'Opis'
    $m.Description = Add-TextBox -Parent $row3 -Width 360 -Placeholder 'opis konta komputera (puste = usuń opis)'
    Add-Button -Parent $row3 -Text 'Ustaw opis' -Icon 'E70F' -Module $m -OnClick { param($m) & $m.Actions.Change $m 'Describe' } | Out-Null
    $row4 = Add-ToolbarRow -Module $m -Title 'Nowe konto'
    Add-Button -Parent $row4 -Text 'Utwórz konto komputera…' -Icon 'E710' -Module $m -OnClick {
        param($m)
        if (-not (Test-AdAvailable)) { return }
        $text = Show-InputDialog -Title 'Nowe konta komputerów' -Prompt 'Nazwy komputerów (NetBIOS, maks. 15 znaków) – po jednej w wierszu. Konta zostaną utworzone jako włączone, gotowe do dołączenia komputerów do domeny.' -Multiline -Icon 'E977' -Validate {
            param($t)
            foreach ($n in (Read-NameList $t)) { $problem = Test-NetBiosName $n; if ($problem) { return "$n`: $problem" } }
            if (@(Read-NameList $t).Count -eq 0) { return 'Wpisz co najmniej jedną nazwę.' }
            return ''
        }
        if (-not $text) { return }
        $names = @(Read-NameList $text | ForEach-Object { $_.ToUpperInvariant() })
        $ou = Select-OrganizationalUnit -Title 'Jednostka dla nowych kont komputerów'
        if (-not $ou) { return }
        if (-not (Confirm-Action -Text "Utworzyć konta komputerów w:`r`n$ou ?" -Items $names -ConfirmText 'Utwórz')) { return }
        Start-AdOperation -Module $m -Name 'Tworzenie kont komputerów' -Targets $names -TargetColumn 'Komputer' -Output Log -Parameters @{ Path = $ou } -ScriptBlock {
            New-ADComputer -Name $Target -SamAccountName "$Target`$" -Path $P.Path -Enabled $true @ad
            "Utworzono w $($P.Path)."
        } -OnComplete {
            param($m)
            Show-Toast 'Utworzono konta – możesz dodać je do listy przyciskiem «Wczytaj z AD».' 'ok'
        }
    } | Out-Null
}

Register-Module -Workspace 'AdComputers' -Category 'Hasła i klucze' -Key 'Laps' -Title 'LAPS' -Icon 'E8D7' `
    -Description 'Hasła lokalnego administratora z AD (Windows LAPS, także szyfrowane, oraz LAPS legacy), wymuszanie zmiany hasła i bezpieczne kopiowanie do schowka (czyszczonego po 60 s).' -Build {
    param($m)
    $m.SecretColumns = @('Hasło')
    $m.PillColumns = @('Rozwiązanie')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-AdComputerTargets)
        if (-not $targets) { return }
        Start-AdOperation -Module $m -Name 'LAPS – odczyt' -Targets $targets -TargetColumn 'Komputer' -ScriptBlock {
            $toDate = {
                param($Value)
                if ($null -eq $Value -or [string]$Value -eq '' -or [string]$Value -eq '0') { return $null }
                try { return [DateTime]::FromFileTimeUtc([int64][string]$Value).ToLocalTime() } catch { return $null }
            }
            $c = Get-ADComputer -Identity $Target -Properties 'msLAPS-Password', 'msLAPS-PasswordExpirationTime', 'msLAPS-EncryptedPassword', 'ms-Mcs-AdmPwd', 'ms-Mcs-AdmPwdExpirationTime' @ad
            $row = $null
            $lapsError = ''
            if (Get-Command Get-LapsADPassword -ErrorAction SilentlyContinue) {
                try {
                    $lp = @{ Identity = $Target; AsPlainText = $true; ErrorAction = 'Stop' }
                    if ($Ctx.Server) { $lp.DomainController = $Ctx.Server }
                    if ($Ctx.Credential) { $lp.Credential = $Ctx.Credential }
                    $info = Get-LapsADPassword @lp
                    if ($info -and $info.Password) {
                        $row = [pscustomobject]@{ 'Rozwiązanie' = 'Windows LAPS'; 'Konto' = $info.Account; 'Hasło' = [string]$info.Password; 'Zmienione' = $info.PasswordUpdateTime; 'Wygasa' = $info.ExpirationTimestamp; 'Źródło' = [string]$info.Source; 'Uwagi' = ''; '__tone' = 'ok' }
                    }
                    elseif ($info) { $lapsError = "Status odszyfrowania: $($info.DecryptionStatus)" }
                }
                catch { $lapsError = $_.Exception.Message }
            }
            if (-not $row -and $c.'msLAPS-Password') {
                try {
                    $json = [string]$c.'msLAPS-Password' | ConvertFrom-Json
                    $row = [pscustomobject]@{ 'Rozwiązanie' = 'Windows LAPS'; 'Konto' = $json.n; 'Hasło' = [string]$json.p; 'Zmienione' = $null; 'Wygasa' = (& $toDate $c.'msLAPS-PasswordExpirationTime'); 'Źródło' = 'msLAPS-Password'; 'Uwagi' = ''; '__tone' = 'ok' }
                }
                catch { }
            }
            if (-not $row -and $c.'ms-Mcs-AdmPwd') {
                $row = [pscustomobject]@{ 'Rozwiązanie' = 'LAPS (legacy)'; 'Konto' = '(wg zasad GPO)'; 'Hasło' = [string]$c.'ms-Mcs-AdmPwd'; 'Zmienione' = $null; 'Wygasa' = (& $toDate $c.'ms-Mcs-AdmPwdExpirationTime'); 'Źródło' = 'ms-Mcs-AdmPwd'; 'Uwagi' = ''; '__tone' = 'info' }
            }
            if (-not $row) {
                $note = 'Brak hasła LAPS albo brak uprawnień do odczytu.'
                if ($c.'msLAPS-EncryptedPassword') { $note = 'Hasło jest zaszyfrowane – brak uprawnień do odszyfrowania lub brak modułu LAPS (RSAT).' }
                if ($lapsError) { $note += " ($lapsError)" }
                $row = [pscustomobject]@{ 'Rozwiązanie' = 'brak'; 'Konto' = ''; 'Hasło' = ''; 'Zmienione' = $null; 'Wygasa' = $null; 'Źródło' = ''; 'Uwagi' = $note; '__tone' = 'crit' }
            }
            $row
        }
    }
    $m.Actions.Copy = {
        param($m, $Rows)
        $source = @(if ($null -ne $Rows) { $Rows } else { Get-SelectedResultRows -Module $m })
        $selected = $source | Select-Object -First 1
        $secret = [string](Get-RowValue $selected 'Hasło')
        if (-not $secret) { Show-Warning 'Zaznacz wiersz z hasłem.'; return }
        Set-ClipboardSecret -Text $secret -Seconds 60
        Write-Log ("Skopiowano hasło LAPS komputera {0}; schowek zostanie wyczyszczony po 60 s." -f (Get-ObjectValue $selected 'Komputer')) 'OK'
        Show-Toast 'Hasło w schowku – zostanie wyczyszczone po 60 s.' 'ok'
    }
    $row = Add-ToolbarRow -Module $m -Title 'Hasła'
    Add-Button -Parent $row -Text 'Pokaż hasła LAPS' -Icon 'E8D7' -Module $m -Primary -OnClick $m.Actions.List | Out-Null
    Add-Button -Parent $row -Text 'Kopiuj hasło zaznaczonego' -Icon 'E8C8' -Module $m -AlwaysEnabled -OnClick { param($m) & $m.Actions.Copy $m $null } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Zmiana hasła'
    $m.ProcessNow = Add-CheckBox -Parent $row2 -Text 'Od razu przetwórz zasady na komputerze' -Checked $true
    Add-Button -Parent $row2 -Text 'Wymuś zmianę hasła' -Icon 'E777' -Module $m -Danger -OnClick {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        if (-not (Confirm-Action -Text "Wymusić zmianę hasła LAPS (ustawienie wygaśnięcia na teraz)?" -Items $targets -ConfirmText 'Wymuś zmianę')) { return }
        Start-AdOperation -Module $m -Name 'LAPS – wymuszenie zmiany' -Targets $targets -TargetColumn 'Komputer' -Output Log -Parameters @{ ProcessNow = (Test-Checked $m.ProcessNow) } -ScriptBlock {
            $name = ($Target -split '\.')[0]
            $c = Get-ADComputer -Identity $name -Properties 'msLAPS-PasswordExpirationTime', 'msLAPS-EncryptedPassword', 'msLAPS-Password', 'ms-Mcs-AdmPwdExpirationTime', 'ms-Mcs-AdmPwd' @ad
            $done = @()
            if ($c.'msLAPS-PasswordExpirationTime' -or $c.'msLAPS-EncryptedPassword' -or $c.'msLAPS-Password') {
                if (Get-Command Set-LapsADPasswordExpirationTime -ErrorAction SilentlyContinue) {
                    $sp = @{ Identity = $name; WhenEffective = (Get-Date); ErrorAction = 'Stop' }
                    if ($Ctx.Server) { $sp.DomainController = $Ctx.Server }
                    if ($Ctx.Credential) { $sp.Credential = $Ctx.Credential }
                    Set-LapsADPasswordExpirationTime @sp | Out-Null
                }
                else { Set-ADComputer -Identity $c -Replace @{ 'msLAPS-PasswordExpirationTime' = 0 } @ad }
                $done += 'Windows LAPS'
            }
            if ($c.'ms-Mcs-AdmPwdExpirationTime' -or $c.'ms-Mcs-AdmPwd') {
                Set-ADComputer -Identity $c -Replace @{ 'ms-Mcs-AdmPwdExpirationTime' = 0 } @ad
                $done += 'LAPS legacy'
            }
            if ($done.Count -eq 0) { throw 'Komputer nie ma atrybutów LAPS w AD (albo brak uprawnień do ich odczytu).' }
            $msg = 'Ustawiono natychmiastowe wygaśnięcie hasła (' + ($done -join ', ') + ')'
            if ($P.ProcessNow) {
                $ic = @{
                    ComputerName = $Target
                    ErrorAction  = 'Stop'
                    ScriptBlock  = {
                        if (Get-Command Invoke-LapsPolicyProcessing -ErrorAction SilentlyContinue) { Invoke-LapsPolicyProcessing; 'Invoke-LapsPolicyProcessing' }
                        else { & gpupdate.exe /target:computer /force | Out-Null; 'gpupdate' }
                    }
                }
                if ($Ctx.Credential) { $ic.Credential = $Ctx.Credential }
                if ($Ctx.SessionOption) { $ic.SessionOption = $Ctx.SessionOption }
                try { $msg += '; na komputerze uruchomiono ' + (Invoke-Command @ic) }
                catch { $msg += '; nie udało się wymusić przetwarzania na komputerze: ' + $_.Exception.Message }
            }
            $msg + '.'
        }
    } | Out-Null
    Add-RowAction -Module $m -Text 'Kopiuj hasło (60 s)' -Icon 'E8C8' -Action { param($m, $rows) & $m.Actions.Copy $m $rows }
    $m.RowDoubleClick = { param($m, $row) & $m.Actions.Copy $m @($row) }
    $m.ResultHint = 'Dwuklik na wierszu – kopiuje hasło do schowka (60 s)'
}

Register-Module -Workspace 'AdComputers' -Category 'Hasła i klucze' -Key 'BitLockerKeys' -Title 'Klucze BitLocker (AD)' -Icon 'E72E' `
    -Description 'Hasła odzyskiwania BitLocker zapisane w AD: dla zaznaczonych komputerów albo wyszukiwanie po identyfikatorze klucza z ekranu odzyskiwania (cała domena). Wymaga uprawnień do msFVE-RecoveryInformation.' -Build {
    param($m)
    $m.SecretColumns = @('Hasło odzyskiwania')
    $m.Actions.Copy = {
        param($m, $Rows)
        $source = @(if ($null -ne $Rows) { $Rows } else { Get-SelectedResultRows -Module $m })
        $row = $source | Select-Object -First 1
        $secret = [string](Get-RowValue $row 'Hasło odzyskiwania')
        if (-not $secret) { Show-Warning 'Zaznacz wiersz z hasłem odzyskiwania.'; return }
        Set-ClipboardSecret -Text $secret -Seconds 60
        Write-Log ("Skopiowano hasło odzyskiwania BitLocker ({0}, klucz {1}); schowek zostanie wyczyszczony po 60 s." -f (Get-ObjectValue $row 'Komputer'), (Get-ObjectValue $row 'Identyfikator klucza')) 'OK'
        Show-Toast 'Hasło odzyskiwania w schowku – zostanie wyczyszczone po 60 s.' 'ok'
    }
    $row = Add-ToolbarRow -Module $m -Title 'Komputery'
    Add-Button -Parent $row -Text 'Klucze zaznaczonych komputerów' -Icon 'E72E' -Module $m -Primary -OnClick {
        param($m)
        $targets = @(Get-AdComputerTargets)
        if (-not $targets) { return }
        Start-AdOperation -Module $m -Name 'BitLocker – klucze w AD' -Targets $targets -TargetColumn 'Komputer' -ScriptBlock {
            $computer = Get-ADComputer -Identity $Target @ad
            $keys = @(Get-ADObject -SearchBase $computer.DistinguishedName -LDAPFilter '(objectClass=msFVE-RecoveryInformation)' -Properties 'msFVE-RecoveryPassword', 'whenCreated' @ad)
            if ($keys.Count -eq 0) { return [pscustomobject]@{ 'Identyfikator klucza' = '(brak kluczy w AD albo brak uprawnień do ich odczytu)'; '__flag' = 'warn' } }
            $keys | Sort-Object whenCreated -Descending | ForEach-Object {
                $id = if ($_.Name -match '\{([0-9A-Fa-f-]{36})\}') { $Matches[1] } else { $_.Name }
                [pscustomobject]@{ 'Utworzono' = $_.whenCreated; 'Identyfikator klucza' = $id; 'Hasło odzyskiwania' = $_.'msFVE-RecoveryPassword' }
            }
        }
    } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Wyszukiwanie'
    $m.KeyId = Add-TextBox -Parent $row2 -Width 260 -Placeholder 'ID klucza (min. 8 znaków), np. 1A2B3C4D'
    Add-Button -Parent $row2 -Text 'Znajdź komputer po ID klucza' -Icon 'E721' -Module $m -OnClick {
        param($m)
        $id = ($m.KeyId.Text.Trim() -replace '[{}\s]', '')
        if ($id -notmatch '^[0-9A-Fa-f-]{8,36}$') { Show-Warning 'Podaj co najmniej 8 pierwszych znaków identyfikatora klucza (cyfry szesnastkowe).'; return }
        Start-AdOperation -Module $m -Name 'BitLocker – wyszukiwanie klucza' -Targets @('AD') -Parameters @{ Id = $id } -ScriptBlock {
            $keys = @(Get-ADObject -LDAPFilter ("(&(objectClass=msFVE-RecoveryInformation)(name=*{{{0}*))" -f $P.Id) -Properties 'msFVE-RecoveryPassword', 'whenCreated' @ad)
            if ($keys.Count -eq 0) { return [pscustomobject]@{ 'Komputer' = '(nie znaleziono klucza o tym identyfikatorze)'; '__flag' = 'warn' } }
            foreach ($k in $keys) {
                $parent = $k.DistinguishedName -replace '^CN=(?:\\.|[^,])+,', ''
                $id = if ($k.Name -match '\{([0-9A-Fa-f-]{36})\}') { $Matches[1] } else { $k.Name }
                [pscustomobject][ordered]@{
                    'Komputer'             = (($parent -replace '^CN=((?:\\.|[^,])+),.*$', '$1') -replace '\\(.)', '$1')
                    'Utworzono'            = $k.whenCreated
                    'Identyfikator klucza' = $id
                    'Hasło odzyskiwania'   = $k.'msFVE-RecoveryPassword'
                    'DN komputera'         = $parent
                }
            }
        }
    } | Out-Null
    Add-Button -Parent $row2 -Text 'Kopiuj hasło zaznaczonego' -Icon 'E8C8' -Module $m -AlwaysEnabled -OnClick { param($m) & $m.Actions.Copy $m $null } | Out-Null
    Add-RowAction -Module $m -Text 'Kopiuj hasło odzyskiwania (60 s)' -Icon 'E8C8' -Action { param($m, $rows) & $m.Actions.Copy $m $rows }
    $m.RowDoubleClick = { param($m, $row) & $m.Actions.Copy $m @($row) }
    $m.ResultHint = 'Dwuklik na wierszu – kopiuje hasło odzyskiwania (60 s)'
}

Register-Module -Workspace 'AdComputers' -Category 'Konta komputerów' -Key 'Rename' -Title 'Zmiana nazwy komputerów' -Icon 'E8AC' `
    -Description 'Wsadowa zmiana nazw z autonumeracją, mapowaniem z listy i walidacją (NetBIOS: maks. 15 znaków). Kolumnę «Nowa nazwa» można edytować bezpośrednio w tabeli.' -Build {
    param($m)
    $m.PillColumns = @('Walidacja')
    $m.Validate = {
        param($m)
        [void]$m.Grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
        $names = @{}
        foreach ($r in $m.Table.Rows) {
            $new = ([string]$r['Nowa nazwa']).Trim()
            if ($new) { $names[$new] = 1 + [int]$names[$new] }
        }
        $valid = 0
        foreach ($r in $m.Table.Rows) {
            $old = [string]$r['Komputer']
            $new = ([string]$r['Nowa nazwa']).Trim()
            $problem = Test-NetBiosName -Name $new
            if (-not $problem -and $new -eq ($old -split '\.')[0]) { $problem = 'Nazwa bez zmian' }
            if (-not $problem -and $names[$new] -gt 1) { $problem = 'Nazwa powtarza się na liście' }
            if (-not $problem -and $script:UI.HostTable.Rows.Find($new)) { $problem = 'Taki komputer jest już na liście' }
            if ($problem) { $r['Walidacja'] = $problem; $r['__tone'] = 'crit' } else { $r['Walidacja'] = 'OK'; $r['__tone'] = 'ok'; $valid++ }
        }
        return $valid
    }
    $m.Actions.Load = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Reset-ResultTable -Module $m
        foreach ($t in $targets) { Add-ResultRows -Module $m -Computer $t -Objects @([pscustomobject]@{ 'Nowa nazwa' = ''; 'Walidacja' = ''; 'Wynik' = '' }) }
        $m.Grid.IsReadOnly = $false
        foreach ($c in $m.Grid.Columns) { $c.IsReadOnly = ([string]$c.SortMemberPath -ne 'Nowa nazwa') }
    }
    $row = Add-ToolbarRow -Module $m -Title 'Lista'
    Add-Button -Parent $row -Text 'Wczytaj zaznaczone komputery' -Icon 'E896' -Module $m -Primary -OnClick $m.Actions.Load | Out-Null
    Add-Button -Parent $row -Text 'Mapowanie z listy…' -Icon 'E8A5' -Module $m -OnClick {
        param($m)
        $text = Show-InputDialog -Title 'Mapowanie nazw' -Prompt 'Wklej pary «STARA NOWA» – po jednej w wierszu (rozdzielone spacją, średnikiem, przecinkiem albo tabulatorem, np. z Excela).' -Multiline -Icon 'E8AC'
        if (-not $text) { return }
        $map = [ordered]@{}
        foreach ($line in ($text -split "`r?`n")) {
            $parts = @($line.Trim() -split '[\s;,]+' | Where-Object { $_ })
            if ($parts.Count -ge 2) { $map[$parts[0]] = $parts[1].ToUpperInvariant() }
        }
        if ($map.Count -eq 0) { Show-Warning 'Nie rozpoznano par nazw.'; return }
        if ($m.Table.Rows.Count -eq 0 -or -not $m.Table.Columns.Contains('Nowa nazwa')) {
            Reset-ResultTable -Module $m
            foreach ($k in $map.Keys) { Add-ResultRows -Module $m -Computer $k -Objects @([pscustomobject]@{ 'Nowa nazwa' = ''; 'Walidacja' = ''; 'Wynik' = '' }) }
            $m.Grid.IsReadOnly = $false
            foreach ($c in $m.Grid.Columns) { $c.IsReadOnly = ([string]$c.SortMemberPath -ne 'Nowa nazwa') }
        }
        $hit = 0
        foreach ($r in $m.Table.Rows) {
            $old = ([string]$r['Komputer'] -split '\.')[0]
            foreach ($k in $map.Keys) { if ($k -eq $old -or $k -eq [string]$r['Komputer']) { $r['Nowa nazwa'] = $map[$k]; $hit++ } }
        }
        [void](& $m.Validate $m)
        Show-Toast "Dopasowano nazw: $hit z $($map.Count)" $(if ($hit -eq $map.Count) { 'ok' } else { 'warn' })
    } | Out-Null
    $row2 = Add-ToolbarRow -Module $m -Title 'Autonumeracja'
    Add-Label -Parent $row2 -Text 'Prefiks' | Out-Null
    $m.Prefix = Add-TextBox -Parent $row2 -Width 100 -Text 'PC-'
    Add-Label -Parent $row2 -Text 'Start' | Out-Null
    $m.Start = Add-Numeric -Parent $row2 -Value 1 -Minimum 0 -Maximum 99999 -Width 70
    Add-Label -Parent $row2 -Text 'Cyfr' | Out-Null
    $m.Pad = Add-Numeric -Parent $row2 -Value 3 -Minimum 1 -Maximum 8 -Width 50
    Add-Label -Parent $row2 -Text 'Sufiks' | Out-Null
    $m.Suffix = Add-TextBox -Parent $row2 -Width 80
    Add-Button -Parent $row2 -Text 'Numeruj' -Icon 'E8EF' -Module $m -OnClick {
        param($m)
        if ($m.Table.Rows.Count -eq 0 -or -not $m.Table.Columns.Contains('Nowa nazwa')) { Show-Warning 'Najpierw wczytaj zaznaczone komputery.'; return }
        [void]$m.Grid.CommitEdit([System.Windows.Controls.DataGridEditingUnit]::Row, $true)
        $n = Get-Num $m.Start
        $format = 'D' + (Get-Num $m.Pad)
        # Numeracja w kolejności widocznej w tabeli (z uwzględnieniem sortowania i filtra)
        foreach ($drv in @($m.Grid.Items | Where-Object { $_ -is [System.Data.DataRowView] })) {
            $drv.Row['Nowa nazwa'] = ($m.Prefix.Text.Trim() + $n.ToString($format) + $m.Suffix.Text.Trim()).ToUpperInvariant()
            $n++
        }
        [void](& $m.Validate $m)
    } | Out-Null
    $row3 = Add-ToolbarRow -Module $m -Title 'Zmiana'
    Add-Button -Parent $row3 -Text 'Sprawdź nazwy' -Icon 'E73E' -Module $m -OnClick {
        param($m)
        if ($m.Table.Rows.Count -eq 0 -or -not $m.Table.Columns.Contains('Nowa nazwa')) { Show-Warning 'Najpierw wczytaj zaznaczone komputery.'; return }
        $valid = & $m.Validate $m
        Show-Toast "Poprawnych nowych nazw: $valid z $($m.Table.Rows.Count)" $(if ($valid -eq $m.Table.Rows.Count) { 'ok' } else { 'warn' })
    } | Out-Null
    $m.Restart = Add-CheckBox -Parent $row3 -Text 'Restart po zmianie (za 30 s)' -Checked $true
    Add-Button -Parent $row3 -Text 'Zmień nazwy' -Icon 'E8AC' -Module $m -Danger -OnClick {
        param($m)
        if ($m.Table.Rows.Count -eq 0 -or -not $m.Table.Columns.Contains('Nowa nazwa')) { Show-Warning 'Najpierw wczytaj zaznaczone komputery.'; return }
        $valid = & $m.Validate $m
        if ($valid -eq 0) { Show-Warning 'Brak poprawnych nowych nazw – sprawdź kolumnę «Walidacja».'; return }
        $per = @{}
        $items = @()
        foreach ($r in $m.Table.Rows) {
            if ([string]$r['Walidacja'] -ne 'OK') { continue }
            $old = [string]$r['Komputer']
            $new = ([string]$r['Nowa nazwa']).Trim()
            $per[$old] = @{ NewName = $new; Restart = (Test-Checked $m.Restart); Credential = $null }
            $items += "$old → $new"
        }
        $cred = Get-EffectiveCredential
        if (-not $cred) {
            $cred = Show-CredentialDialog -Message 'Zmiana nazwy komputera w domenie wymaga poświadczeń domenowych z prawem zmiany obiektu komputera (nie przechodzą przez WinRM automatycznie).'
            if (-not $cred) { return }
        }
        foreach ($k in @($per.Keys)) { $per[$k].Credential = $cred }
        $skipped = $m.Table.Rows.Count - $per.Count
        $question = 'Zmienić nazwy komputerów?'
        if ($skipped -gt 0) { $question += " (pominięte z powodu walidacji: $skipped)" }
        if (-not (Confirm-Action -Text $question -Items $items -ConfirmText 'Zmień nazwy' -Danger)) { return }
        $m.Grid.IsReadOnly = $true
        Start-HostOperation -Module $m -Name 'Zmiana nazw' -Targets @($per.Keys) -PerTarget $per -Output Log -ScriptBlock {
            param($P)
            $rp = @{ NewName = $P.NewName; DomainCredential = $P.Credential; Force = $true; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
            Rename-Computer @rp
            if ($P.Restart) {
                & (Join-Path $env:SystemRoot 'System32\shutdown.exe') /r /t 30 /d p:4:2 /c ('Zmiana nazwy komputera na ' + $P.NewName) | Out-Null
                "Zmieniono nazwę na $($P.NewName) – restart za 30 s."
            }
            else { "Zmieniono nazwę na $($P.NewName) – zmiana zadziała po restarcie." }
        } -OnResult {
            param($m, $r)
            $row = $null
            foreach ($x in $m.Table.Rows) { if ([string]$x['Komputer'] -eq $r.Target) { $row = $x; break } }
            if (-not $row) { return }
            if ($r.Ok) {
                Set-ResultValue -Module $m -Row $row -Column 'Wynik' -Value 'Zmieniono'
                Rename-ComputerRow -OldName $r.Target -NewName ([string]$row['Nowa nazwa']).Trim()
            }
            else {
                Set-ResultValue -Module $m -Row $row -Column 'Wynik' -Value ('Błąd – ' + ((@($r.Errors)) -join ' | '))
                $row['__flag'] = 'crit'
            }
        }
    } | Out-Null
    $m.ResultHint = 'Kliknij dwukrotnie komórkę «Nowa nazwa», aby ją edytować'
    $m.RowDoubleClick = { param($m, $row) }
}

Register-GroupMembershipModule -Workspace 'AdComputers' -Kind Computer -Key 'ComputerGroups' -Category 'Konta komputerów'

$script:ComputerReports = @(
    @{ Key = 'Inactive'; Name = 'Nieaktywne (brak logowania od N dni)' }
    @{ Key = 'Disabled'; Name = 'Wyłączone konta' }
    @{ Key = 'Created'; Name = 'Utworzone w ciągu N dni' }
    @{ Key = 'OsSummary'; Name = 'Podsumowanie systemów operacyjnych' }
    @{ Key = 'Unsupported'; Name = 'Nieobsługiwane systemy operacyjne' }
    @{ Key = 'NoLaps'; Name = 'Bez hasła LAPS w AD' }
    @{ Key = 'NoBitLocker'; Name = 'Bez klucza BitLocker w AD' }
    @{ Key = 'Servers'; Name = 'Serwery' }
    @{ Key = 'All'; Name = 'Wszystkie komputery' }
)

Register-Module -Workspace 'AdComputers' -Category 'Raporty' -Key 'ComputerReports' -Title 'Raporty komputerów' -Icon 'E9F9' `
    -Description 'Zestawienia z domeny lub jednostki: nieaktywne, wyłączone, nowe, systemy operacyjne, nieobsługiwane systemy, brak LAPS i kluczy BitLocker. Wyniki można przenieść na listę komputerów, wyłączyć, przenieść lub usunąć.' -Build {
    param($m)
    $m.ColorBools = $true
    $m.PillColumns = @('Stan')
    $row = Add-ToolbarRow -Module $m -Title 'Raport'
    $m.Report = Add-ComboBox -Parent $row -Items @($script:ComputerReports | ForEach-Object { $_.Name }) -Width 340
    Add-Label -Parent $row -Text 'N (dni)' | Out-Null
    $m.Days = Add-Numeric -Parent $row -Value ([int]$script:Settings.InactiveDays) -Minimum 1 -Maximum 3650 -Width 70
    $m.OnlyEnabled = Add-CheckBox -Parent $row -Text 'Tylko włączone' -Checked $true
    $row2 = Add-ToolbarRow -Module $m -Title 'Zakres'
    $m.Ou = Add-TextBox -Parent $row2 -Width 420 -Placeholder 'Cała domena'
    Add-Button -Parent $row2 -Text '' -Icon 'E8B7' -Module $m -AlwaysEnabled -ToolTip 'Wybierz jednostkę organizacyjną' -OnClick {
        param($m)
        $ou = Select-OrganizationalUnit -Title 'Zakres raportu' -Selected $m.Ou.Text.Trim() -AllowDomainRoot
        if ($null -ne $ou) { $m.Ou.Text = $ou }
    } | Out-Null
    Add-Button -Parent $row2 -Text 'Generuj raport' -Icon 'E9F9' -Module $m -Primary -OnClick {
        param($m)
        $report = $script:ComputerReports[$m.Report.SelectedIndex]
        $params = @{ Report = $report.Key; Days = (Get-Num $m.Days); OnlyEnabled = (Test-Checked $m.OnlyEnabled); SearchBase = $m.Ou.Text.Trim() }
        Reset-StatTiles $m
        Start-AdOperation -Module $m -Name $report.Name -Targets @('AD') -Parameters $params -ScriptBlock {
            $now = Get-Date
            $days = [int]$P.Days
            $cut = $now.AddDays(-$days)
            $ft = $cut.ToFileTimeUtc()
            $gen = $cut.ToUniversalTime().ToString('yyyyMMddHHmmss.0Z')
            $extra = switch ($P.Report) {
                'Inactive' { "(|(lastLogonTimestamp<=$ft)(&(!(lastLogonTimestamp=*))(whenCreated<=$gen)))" }
                'Disabled' { '(userAccountControl:1.2.840.113556.1.4.803:=2)' }
                'Created' { "(whenCreated>=$gen)" }
                'Servers' { '(operatingSystem=*Server*)' }
                'NoLaps' { '(!(ms-Mcs-AdmPwdExpirationTime=*))(!(msLAPS-PasswordExpirationTime=*))' }
                default { '' }
            }
            if ($P.OnlyEnabled -and $P.Report -ne 'Disabled') { $extra += '(!(userAccountControl:1.2.840.113556.1.4.803:=2))' }
            $q = @{
                LDAPFilter = "(&(objectCategory=computer)$extra)"
                Properties = @('OperatingSystem', 'OperatingSystemVersion', 'Enabled', 'LastLogonDate', 'PasswordLastSet', 'whenCreated', 'Description', 'IPv4Address')
            }
            if ($P.SearchBase) { $q.SearchBase = $P.SearchBase }
            try { $computers = @(Get-ADComputer @q @ad) }
            catch {
                # Brak schematu LAPS (atrybut nieznany) - raport bez filtra jednego z atrybutów
                if ($P.Report -eq 'NoLaps') { $q.LDAPFilter = $q.LDAPFilter -replace '\(!\(msLAPS-PasswordExpirationTime=\*\)\)', ''; $computers = @(Get-ADComputer @q @ad) } else { throw }
            }
            if ($P.Report -eq 'OsSummary') {
                $computers | Group-Object { '{0}|{1}' -f $_.OperatingSystem, $_.OperatingSystemVersion } | Sort-Object Count -Descending | ForEach-Object {
                    $parts = $_.Name -split '\|'
                    [pscustomobject][ordered]@{ 'System' = $(if ($parts[0]) { $parts[0] } else { '(nieznany)' }); 'Wersja' = $parts[1]; 'Liczba' = $_.Count; 'Włączone' = @($_.Group | Where-Object { $_.Enabled }).Count }
                }
                return
            }
            $withKey = $null
            if ($P.Report -eq 'NoBitLocker') {
                $withKey = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
                $kq = @{ LDAPFilter = '(objectClass=msFVE-RecoveryInformation)' }
                if ($P.SearchBase) { $kq.SearchBase = $P.SearchBase }
                foreach ($k in @(Get-ADObject @kq @ad)) { [void]$withKey.Add(($k.DistinguishedName -replace '^CN=(?:\\.|[^,])+,', '')) }
            }
            # Systemy bez wsparcia producenta (stan na 2026 r.): Windows 10 poza LTSC, Windows 7/8.x, Server 2003-2012 R2
            $unsupported = 'Windows (XP|Vista|7|8|8\.1)( |$)|Windows 10 (?!.*LTSC)|Windows Server (2003|2008|2012)'
            foreach ($c in $computers) {
                $include = switch ($P.Report) {
                    'Unsupported' { [string]$c.OperatingSystem -match $unsupported }
                    'NoBitLocker' { -not $withKey.Contains($c.DistinguishedName) }
                    default { $true }
                }
                if (-not $include) { continue }
                $daysIdle = if ($c.LastLogonDate) { [int]($now - $c.LastLogonDate).TotalDays } else { $null }
                $state = if (-not $c.Enabled) { 'Wyłączone' } elseif ($null -eq $daysIdle) { 'Nigdy nie logowane' } elseif ($daysIdle -gt $days) { "Nieaktywne ($daysIdle dni)" } else { 'Aktywne' }
                [pscustomobject][ordered]@{
                    'Komputer'              = $c.Name
                    'Stan'                  = $state
                    'System'                = $c.OperatingSystem
                    'Wersja'                = $c.OperatingSystemVersion
                    'Włączone'              = [bool]$c.Enabled
                    'Ostatnie logowanie'    = $c.LastLogonDate
                    'Dni bez logowania'     = $daysIdle
                    'Hasło konta zmienione' = $c.PasswordLastSet
                    'Utworzono'             = $c.whenCreated
                    'IPv4'                  = $c.IPv4Address
                    'Opis'                  = $c.Description
                    'DN'                    = $c.DistinguishedName
                    '__tone'                = $(if (-not $c.Enabled) { '' } elseif ($state -eq 'Aktywne') { 'ok' } else { 'warn' })
                }
            }
        } -OnComplete {
            param($m)
            $count = $m.Table.Rows.Count
            Set-StatTile -Module $m -Key 'count' -Value ([string]$count)
            if ($m.Table.Columns.Contains('Włączone') -and $m.Table.Columns.Contains('Komputer')) {
                Set-StatTile -Module $m -Key 'disabled' -Value ([string]@($m.Table.Rows | Where-Object { [string]$_['Włączone'] -eq 'Nie' }).Count)
                Set-StatTile -Module $m -Key 'inactive' -Value ([string]@($m.Table.Rows | Where-Object { [string]$_['Stan'] -like 'Nieaktywne*' -or [string]$_['Stan'] -eq 'Nigdy nie logowane' }).Count) -Tone 'warn'
            }
        }
    } | Out-Null
    $m.Actions.Bulk = {
        param($m, [string]$Op, $Rows)
        $source = @(if ($null -ne $Rows) { $Rows } else { Get-SelectedResultRows -Module $m })
        $names = @($source | ForEach-Object { [string](Get-ObjectValue $_ 'Komputer') } | Where-Object { $_ -and $_ -notlike '(*' } | Select-Object -Unique)
        if ($names.Count -eq 0) { Show-Warning 'Zaznacz w tabeli komputery.'; return }
        $params = @{ Op = $Op; TargetOU = '' }
        switch ($Op) {
            'Disable' { if (-not (Confirm-Action -Text 'Wyłączyć konta wybranych komputerów?' -Items $names -ConfirmText 'Wyłącz' -Danger)) { return } }
            'Delete' { if (-not (Confirm-Action -Text 'USUNĄĆ konta wybranych komputerów z AD (razem z obiektami podrzędnymi)?' -Items $names -ConfirmText 'Usuń konta' -Danger)) { return } }
            'Move' {
                $ou = Select-OrganizationalUnit -Title 'Docelowa jednostka organizacyjna'
                if (-not $ou) { return }
                $params.TargetOU = $ou
                if (-not (Confirm-Action -Text "Przenieść konta do:`r`n$ou ?" -Items $names -ConfirmText 'Przenieś')) { return }
            }
        }
        Start-AdOperation -Module $m -Name "Raport – $Op" -Targets $names -TargetColumn 'Komputer' -Output Log -Parameters $params -ScriptBlock {
            $c = Get-ADComputer -Identity $Target @ad
            switch ($P.Op) {
                'Disable' { Disable-ADAccount -Identity $c @ad; 'Konto wyłączone.' }
                'Move' { Move-ADObject -Identity $c.DistinguishedName -TargetPath $P.TargetOU @ad; "Przeniesiono do $($P.TargetOU)." }
                'Delete' { Remove-ADObject -Identity $c.DistinguishedName -Recursive -Confirm:$false @ad; 'Konto usunięte z AD.' }
            }
        } -OnResult {
            param($m, $r)
            if ($r.Ok -and @($r.Data | Where-Object { [string]$_ -like 'Konto usunięte*' }).Count) {
                Remove-ResultRows -Module $m -Rows @($m.Table.Rows | Where-Object { [string]$_['Komputer'] -eq $r.Target })
            }
        }
    }
    $row3 = Add-ToolbarRow -Module $m -Title 'Wyniki'
    Add-Button -Parent $row3 -Text 'Zaznacz na liście komputerów' -Icon 'E8B3' -Module $m -AlwaysEnabled -OnClick { param($m) Add-ResultsToTargets -Module $m -Kind Computer -Column 'Komputer' } | Out-Null
    Add-Button -Parent $row3 -Text 'Wyłącz zaznaczone' -Icon 'E8D8' -Module $m -Danger -OnClick { param($m) & $m.Actions.Bulk $m 'Disable' $null } | Out-Null
    Add-Button -Parent $row3 -Text 'Przenieś zaznaczone…' -Icon 'E8DE' -Module $m -OnClick { param($m) & $m.Actions.Bulk $m 'Move' $null } | Out-Null
    Add-Button -Parent $row3 -Text 'Usuń zaznaczone z AD' -Icon 'E74D' -Module $m -Danger -OnClick { param($m) & $m.Actions.Bulk $m 'Delete' $null } | Out-Null
    Add-RowAction -Module $m -Text 'Zaznacz na liście komputerów' -Icon 'E8B3' -Action { param($m, $rows) Add-ResultsToTargets -Module $m -Kind Computer -Column 'Komputer' -Rows $rows }
    Add-RowAction -Module $m -Text 'Wyłącz konto' -Icon 'E8D8' -Separator -Action { param($m, $rows) & $m.Actions.Bulk $m 'Disable' $rows }
    Add-RowAction -Module $m -Text 'Przenieś do OU…' -Icon 'E8DE' -Action { param($m, $rows) & $m.Actions.Bulk $m 'Move' $rows }
    Add-RowAction -Module $m -Text 'Usuń z AD' -Icon 'E74D' -Danger -Action { param($m, $rows) & $m.Actions.Bulk $m 'Delete' $rows }
    Add-StatTile -Module $m -Key 'count' -Label 'Pozycje raportu' -Icon 'E977' | Out-Null
    Add-StatTile -Module $m -Key 'inactive' -Label 'Nieaktywne / nigdy nie logowane' -Icon 'E916' | Out-Null
    Add-StatTile -Module $m -Key 'disabled' -Label 'Wyłączone' -Icon 'E8D8' | Out-Null
    $m.EmptyHint = 'Wybierz raport i kliknij «Generuj raport» (F5). Nie trzeba zaznaczać komputerów na liście.'
}
#endregion

#region Uruchomienie
function Initialize-MainWindow {
    # Buduje okno główne (bez wyświetlania) - wydzielone, aby dało się je testować
    Import-Settings
    $w = New-MainWindow
    $c = $script:UI.Controls

    $clip = New-Object System.Windows.Threading.DispatcherTimer
    $clip.Interval = [TimeSpan]::FromSeconds(60)
    $clip.add_Tick({ param($s, $e) Clear-ClipboardSecret })
    $script:Clipboard.Timer = $clip
    Initialize-EngineTimer

    $cp = Initialize-ComputerPanel
    [void]$c.targetHost.Children.Add($cp.Root)
    $c.panelComputers = $cp.Root
    $up = Initialize-UserPanel
    [void]$c.targetHost.Children.Add($up.Root)
    $c.panelUsers = $up.Root

    Initialize-Navigation
    Set-LogVisible ([bool]$script:Settings.LogVisible)
    $w.Dispatcher.add_UnhandledException({
            param($s, $e)
            $e.Handled = $true
            Write-Log "Nieobsłużony błąd interfejsu: $($e.Exception.Message)" 'ERROR' -Module ''
        })

    Write-Log ("Domain Ops {0} – PowerShell {1}, użytkownik {2}\{3}" -f $script:AppVersion, $PSVersionTable.PSVersion, $env:USERDOMAIN, $env:USERNAME) -Module ''
    if ($script:PluginFiles.Count -gt 0) {
        Write-Log ("Wczytano moduły własne: {0}" -f ((@($script:PluginFiles | ForEach-Object { $_.Name })) -join ', ')) 'OK' -Module ''
    }
    foreach ($err in $script:PluginErrors) { Write-Log "Błąd modułu własnego $err" 'ERROR' -Module '' }
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Log 'Brak modułu ActiveDirectory (RSAT) – przestrzenie AD i wczytywanie list z AD będą niedostępne. Zdalne zarządzanie działa z listą wpisaną ręcznie.' 'WARN' -Module ''
    }
    Show-Workspace -Key ([string]$script:Settings.LastWorkspace)
    return $w
}

function Start-DomainOps {
    $w = Initialize-MainWindow
    try {
        [void]$w.ShowDialog()
    }
    finally {
        Close-Engine
    }
}

# Moduły własne: pliki *.ps1 z folderu AD-ManagerDiamond.Modules obok skryptu (wczytywane w zakresie skryptu,
# więc mogą korzystać ze wszystkich funkcji programu i wywoływać Register-Workspace / Register-Module)
$script:PluginErrors = New-Object System.Collections.ArrayList
$script:PluginFiles = @()
if ($script:App.ModulesDir -and (Test-Path -LiteralPath $script:App.ModulesDir)) {
    $script:PluginFiles = @(Get-ChildItem -LiteralPath $script:App.ModulesDir -Filter '*.ps1' -File | Sort-Object Name)
}
foreach ($pluginFile in $script:PluginFiles) {
    try { . $pluginFile.FullName }
    catch { [void]$script:PluginErrors.Add(('{0}: {1}' -f $pluginFile.Name, $_.Exception.Message)) }
}

# DOMAINOPS_NOSTART=1 - tylko wczytanie funkcji (testy, osadzanie)
if ($env:DOMAINOPS_NOSTART -ne '1') { Start-DomainOps }
#endregion
