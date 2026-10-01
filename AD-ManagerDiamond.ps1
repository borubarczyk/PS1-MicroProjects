<#
.SYNOPSIS
    Domain Ops (AD-ManagerDiamond) - graficzne narzędzie do zdalnej administracji komputerami w domenie Active Directory.

.DESCRIPTION
    Okno składa się z trzech części:
      - lewy panel: lista komputerów (z Active Directory, wpisanych ręcznie lub wczytanych z pliku)
        z zaznaczaniem hostów docelowych, szybkim wyszukiwaniem i menu kontekstowym,
      - drzewo modułów pogrupowanych w kategorie (diagnostyka, zdalne wykonanie, system, oprogramowanie,
        bezpieczeństwo, udostępnianie, Active Directory),
      - aktywny moduł z wynikami w tabeli: filtrowanie, sortowanie, eksport CSV, kopiowanie, podgląd szczegółów.

    Wszystkie operacje wykonywane są w tle i równolegle (pula wątków PowerShell), więc okno nie zawiesza się,
    a wyniki pojawiają się na bieżąco dla kolejnych hostów. Trwające operacje można anulować z paska stanu.
    Operacje zdalne korzystają z PowerShell Remoting (WinRM); operacje na obiektach AD (konto komputera, LAPS,
    klucze odzyskiwania BitLocker) wykonywane są lokalnie modułem ActiveDirectory.

.NOTES
    Wymagania:
      - Windows PowerShell 5.1 lub PowerShell 7 w systemie Windows (skrypt sam uruchomi się ponownie w trybie STA),
      - RSAT: moduł ActiveDirectory (lista komputerów i moduły AD), opcjonalnie moduł LAPS (Windows LAPS),
      - włączony WinRM na hostach docelowych i uprawnienia administratora lokalnego.
    Ustawienia: %APPDATA%\AD-ManagerDiamond\settings.json
    Dziennik:   %LOCALAPPDATA%\AD-ManagerDiamond\Logs\DomainOps_RRRRMMDD.log
    Pliki robocze na hostach: %SystemRoot%\Temp\DomainOps
    Plik musi pozostać zapisany jako UTF-8 z BOM (polskie znaki w Windows PowerShell 5.1).

.EXAMPLE
    powershell.exe -STA -ExecutionPolicy Bypass -File .\AD-ManagerDiamond.ps1
#>
#Requires -Version 5.1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Start: platforma i tryb STA
if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) {
    throw 'Domain Ops działa wyłącznie w systemie Windows.'
}

# WinForms wymaga wątku STA - jeśli go nie mamy, uruchamiamy skrypt ponownie z przełącznikiem -STA
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    if ($PSCommandPath) {
        $exe = (Get-Process -Id $PID).Path
        Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-STA', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath)) | Out-Null
        return
    }
    throw 'Uruchom skrypt w trybie STA: powershell.exe -STA -File .\AD-ManagerDiamond.ps1'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

# Moduł ActiveDirectory nie musi tworzyć dysku AD: (szybszy import, mniej błędów przy braku DC)
$env:ADPS_LoadDefaultDrive = '0'
#endregion

#region Konfiguracja i stan
$script:AppVersion = '3.0'

$script:App = @{
    Name    = 'Domain Ops'
    DataDir = Join-Path $env:APPDATA 'AD-ManagerDiamond'
    LogDir  = Join-Path $env:LOCALAPPDATA 'AD-ManagerDiamond\Logs'
}
$script:App.SettingsFile = Join-Path $script:App.DataDir 'settings.json'
$script:App.LogFile = Join-Path $script:App.LogDir ('DomainOps_{0:yyyyMMdd}.log' -f (Get-Date))

# Ustawienia zapamiętywane między uruchomieniami
$script:Settings = [ordered]@{
    SearchBase       = ''
    NameFilter       = ''
    OnlyEnabled      = $true
    DomainController = ''
    ThrottleLimit    = 16
    TimeoutSec       = 20
    LastModule       = 'Connectivity'
    WindowWidth      = 1500
    WindowHeight     = 900
    WindowMaximized  = $false
}

# Poświadczenia bieżącej sesji (nie są zapisywane na dysku)
$script:State = @{
    Credential = $null
    UseCurrent = $true
}

# Kontrolki i konteksty modułów
$script:UI = @{
    Font        = New-Object System.Drawing.Font('Segoe UI', 9)
    FontBold    = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    FontMono    = New-Object System.Drawing.Font('Consolas', 9.5)
    Form        = $null
    Modules     = @{}
    ModuleDefs  = New-Object System.Collections.ArrayList
    Categories  = @('Diagnostyka', 'Zdalne wykonanie', 'System', 'Oprogramowanie', 'Bezpieczeństwo', 'Udostępnianie', 'Active Directory')
    ActiveModule = $null
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
        [string]$Module = $script:LogContext
    )
    $prefix = if ($Module) { "[$Module] " } else { '' }
    $now = Get-Date
    $box = $script:UI['LogBox']
    if ($box) {
        try {
            if ($box.TextLength -gt 500000) {
                $box.Clear()
                $box.AppendText('(dziennik w oknie został skrócony - pełny zapis znajduje się w pliku logu)' + [Environment]::NewLine)
            }
            $color = switch ($Level) {
                'OK' { [System.Drawing.Color]::ForestGreen }
                'WARN' { [System.Drawing.Color]::DarkOrange }
                'ERROR' { [System.Drawing.Color]::Firebrick }
                default { [System.Drawing.Color]::FromArgb(40, 40, 40) }
            }
            $box.SelectionStart = $box.TextLength
            $box.SelectionLength = 0
            $box.SelectionColor = $color
            $box.AppendText(('{0:HH:mm:ss}  {1,-5}  {2}{3}' -f $now, $Level, $prefix, $Message) + [Environment]::NewLine)
            $box.ScrollToCaret()
        }
        catch { }
    }
    try {
        $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}{3}' -f $now, $Level, $prefix, $Message
        [System.IO.File]::AppendAllText($script:App.LogFile, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    }
    catch { }
}

function Show-Message {
    param(
        [string]$Text,
        [string]$Title = 'Domain Ops',
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )
    $owner = $script:UI.Form
    if ($owner) { [void][System.Windows.Forms.MessageBox]::Show($owner, $Text, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, $Icon) }
    else { [void][System.Windows.Forms.MessageBox]::Show($Text, $Title, [System.Windows.Forms.MessageBoxButtons]::OK, $Icon) }
}

function Show-Warning([string]$Text) {
    Show-Message -Text $Text -Title 'Uwaga' -Icon ([System.Windows.Forms.MessageBoxIcon]::Warning)
}

function Show-Error {
    param([string]$Text, $ErrorObject = $null)
    # $ErrorObject może być ErrorRecord ($_ z bloku catch), wyjątkiem albo tekstem
    $detail = if ($ErrorObject -is [System.Management.Automation.ErrorRecord]) { $ErrorObject.Exception.Message }
    elseif ($ErrorObject -is [System.Exception]) { $ErrorObject.Message }
    elseif ($ErrorObject) { [string]$ErrorObject }
    else { '' }
    $message = if ($detail) { "$Text`r`n`r`n$detail" } else { $Text }
    Show-Message -Text $message -Title 'Błąd' -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
}

function Confirm-Action {
    param([string]$Text, [string[]]$Items = @())
    $message = $Text
    $list = @($Items | Where-Object { $_ })
    if ($list.Count -gt 0) {
        $message += "`r`n`r`n" + ((@($list | Select-Object -First 15)) -join "`r`n")
        if ($list.Count -gt 15) { $message += "`r`n… i $($list.Count - 15) więcej" }
    }
    $answer = [System.Windows.Forms.MessageBox]::Show($script:UI.Form, $message, 'Potwierdzenie',
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    return ($answer -eq [System.Windows.Forms.DialogResult]::Yes)
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
    if ($InputObject -is [System.Data.DataRowView]) {
        if (-not $InputObject.Row.Table.Columns.Contains($Name)) { return $null }
        $v = $InputObject.Row[$Name]
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

function Format-LogObject {
    param($InputObject)
    if ($null -eq $InputObject) { return '' }
    if ($InputObject -is [System.Management.Automation.PSObject]) {
        $base = $InputObject.PSObject.BaseObject
        if ($base -is [string] -or $base -is [System.ValueType]) { return [string]$base }
    }
    if ($InputObject -is [string] -or $InputObject -is [System.ValueType]) { return [string]$InputObject }
    $parts = foreach ($p in $InputObject.PSObject.Properties) {
        if ($script:HiddenProperties -contains $p.Name) { continue }
        $v = ConvertTo-CellValue $p.Value
        if ($v -is [System.DBNull] -or [string]$v -eq '') { continue }
        '{0}: {1}' -f $p.Name, $v
    }
    return (@($parts) -join '; ')
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

function Set-ClipboardSecret {
    # Kopiuje poufny tekst do schowka i czyści go po upływie czasu (o ile nadal tam jest)
    param([string]$Text, [int]$Seconds = 60)
    if ([string]::IsNullOrEmpty($Text)) { return }
    [System.Windows.Forms.Clipboard]::SetText($Text)
    $script:Clipboard.Secret = $Text
    $timer = $script:Clipboard.Timer
    if ($timer) {
        $timer.Stop()
        $timer.Interval = [Math]::Max(5, $Seconds) * 1000
        $timer.Start()
    }
}

function Clear-ClipboardSecret {
    if (-not $script:Clipboard.Secret) { return }
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsText() -and [System.Windows.Forms.Clipboard]::GetText() -eq $script:Clipboard.Secret) {
            [System.Windows.Forms.Clipboard]::Clear()
            Write-Log 'Wyczyszczono poufną wartość ze schowka.' -Module ''
        }
    }
    catch { }
    $script:Clipboard.Secret = $null
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
    $script:Settings.ThrottleLimit = [Math]::Min(64, [Math]::Max(1, [int]$script:Settings.ThrottleLimit))
    $script:Settings.TimeoutSec = [Math]::Min(300, [Math]::Max(5, [int]$script:Settings.TimeoutSec))
    $script:Settings.OnlyEnabled = [bool]$script:Settings.OnlyEnabled
    $script:Settings.WindowMaximized = [bool]$script:Settings.WindowMaximized
}

function Export-Settings {
    try {
        if (-not (Test-Path -LiteralPath $script:App.DataDir)) { New-Item -ItemType Directory -Path $script:App.DataDir -Force | Out-Null }
        $script:Settings | ConvertTo-Json | Set-Content -LiteralPath $script:App.SettingsFile -Encoding UTF8
    }
    catch { }
}
#endregion

#region Fabryka kontrolek
function Add-DockStack {
    # Układa kontrolki: -Top od góry (w podanej kolejności), -Bottom od dołu (ostatnia na samym dole), -Fill w pozostałym miejscu.
    # WinForms dokuje kontrolki w odwrotnej kolejności dodawania, stąd odwrócone pętle.
    param($Parent, [object[]]$Top = @(), $Fill = $null, [object[]]$Bottom = @())
    if ($Fill) {
        $Fill.Dock = [System.Windows.Forms.DockStyle]::Fill
        $Parent.Controls.Add($Fill)
    }
    foreach ($c in $Bottom) {
        $c.Dock = [System.Windows.Forms.DockStyle]::Bottom
        $Parent.Controls.Add($c)
    }
    for ($i = $Top.Count - 1; $i -ge 0; $i--) {
        $Top[$i].Dock = [System.Windows.Forms.DockStyle]::Top
        $Parent.Controls.Add($Top[$i])
    }
}

function New-FlowRow {
    $p = New-Object System.Windows.Forms.FlowLayoutPanel
    $p.AutoSize = $true
    $p.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $p.WrapContents = $true
    $p.Margin = New-Object System.Windows.Forms.Padding(0)
    $p.Padding = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)
    return $p
}

function New-StretchRow {
    # Wiersz: kontrolka rozciągana na całą szerokość, z kontrolkami o stałym rozmiarze przed (-Before) i za nią (-After)
    param($Stretch, [object[]]$Before = @(), [object[]]$After = @())
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.AutoSize = $true
    $t.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $t.RowCount = 1
    $t.ColumnCount = $Before.Count + 1 + $After.Count
    $t.Margin = New-Object System.Windows.Forms.Padding(0)
    [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $col = 0
    foreach ($c in $Before) {
        [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
        $t.Controls.Add($c, $col, 0)
        $col++
    }
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $Stretch.Anchor = [System.Windows.Forms.AnchorStyles]'Left,Right'
    $t.Controls.Add($Stretch, $col, 0)
    $col++
    foreach ($c in $After) {
        [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
        $t.Controls.Add($c, $col, 0)
        $col++
    }
    return $t
}

function Add-Label {
    param($Parent, [string]$Text, [switch]$Hint, [switch]$Bold)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.AutoSize = $true
    $l.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $l.Margin = New-Object System.Windows.Forms.Padding(3, 7, 3, 3)
    if ($Hint) { $l.ForeColor = [System.Drawing.Color]::DimGray }
    if ($Bold) { $l.Font = $script:UI.FontBold }
    if ($Parent) { $Parent.Controls.Add($l) }
    return $l
}

function Add-TextBox {
    param($Parent, [int]$Width = 160, [string]$Text = '')
    $t = New-Object System.Windows.Forms.TextBox
    $t.Width = $Width
    $t.Text = $Text
    $t.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    if ($Parent) { $Parent.Controls.Add($t) }
    return $t
}

function Add-ComboBox {
    param($Parent, [string[]]$Items, [int]$Width = 150, [int]$SelectedIndex = 0)
    $c = New-Object System.Windows.Forms.ComboBox
    $c.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $c.Width = $Width
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    foreach ($i in $Items) { [void]$c.Items.Add($i) }
    if ($c.Items.Count -gt $SelectedIndex) { $c.SelectedIndex = $SelectedIndex }
    if ($Parent) { $Parent.Controls.Add($c) }
    return $c
}

function Add-Numeric {
    param($Parent, [int]$Minimum = 0, [int]$Maximum = 100, [int]$Value = 0, [int]$Width = 70)
    $n = New-Object System.Windows.Forms.NumericUpDown
    $n.Minimum = $Minimum
    $n.Maximum = $Maximum
    $n.Value = [Math]::Min($Maximum, [Math]::Max($Minimum, $Value))
    $n.Width = $Width
    $n.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 3)
    if ($Parent) { $Parent.Controls.Add($n) }
    return $n
}

function Add-CheckBox {
    param($Parent, [string]$Text, [bool]$Checked = $false)
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.Checked = $Checked
    $c.AutoSize = $true
    $c.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $c.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 3)
    if ($Parent) { $Parent.Controls.Add($c) }
    return $c
}

function Add-RadioButton {
    param($Parent, [string]$Text, [bool]$Checked = $false)
    $r = New-Object System.Windows.Forms.RadioButton
    $r.Text = $Text
    $r.Checked = $Checked
    $r.AutoSize = $true
    $r.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $r.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 3)
    if ($Parent) { $Parent.Controls.Add($r) }
    return $r
}

function New-PlainButton {
    param($Parent, [string]$Text)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.AutoSize = $true
    $b.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $b.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
    $b.MinimumSize = New-Object System.Drawing.Size(0, 27)
    $b.Margin = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)
    $b.UseVisualStyleBackColor = $true
    if ($Parent) { $Parent.Controls.Add($b) }
    return $b
}

# Obsługa zdarzeń kontrolek modułów: handler dostaje kontekst modułu ($m) niezależnie od tego,
# gdzie został zdefiniowany (zmienne lokalne buildera modułu nie istnieją już w chwili kliknięcia).
$script:Dispatchers = @{
    Click                = { param($s) Invoke-ControlHandler -Source $s -EventName 'Click' }
    TextChanged          = { param($s) Invoke-ControlHandler -Source $s -EventName 'TextChanged' }
    CheckedChanged       = { param($s) Invoke-ControlHandler -Source $s -EventName 'CheckedChanged' }
    SelectedIndexChanged = { param($s) Invoke-ControlHandler -Source $s -EventName 'SelectedIndexChanged' }
}

function Register-ControlHandler {
    param($Control, [string]$EventName, [hashtable]$Module, [scriptblock]$Action)
    $tag = $Control.Tag
    if (-not ($tag -is [hashtable])) {
        $tag = @{}
        $Control.Tag = $tag
    }
    $tag['ModuleKey'] = $Module.Key
    $tag["On$EventName"] = $Action
    $Control."add_$EventName"($script:Dispatchers[$EventName])
}

function Invoke-ControlHandler {
    param($Source, [string]$EventName)
    try {
        $tag = $Source.Tag
        if (-not ($tag -is [hashtable])) { return }
        $module = $script:UI.Modules[[string]$tag['ModuleKey']]
        $action = $tag["On$EventName"]
        if ($action) { Invoke-UiAction -Module $module -Action $action -Source $Source }
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

function Add-Button {
    # Przycisk akcji modułu - wyłączany automatycznie na czas operacji tego modułu
    param(
        [Parameter(Mandatory)]$Parent,
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][hashtable]$Module,
        [Parameter(Mandatory)][scriptblock]$OnClick,
        [switch]$Primary,
        [switch]$Danger
    )
    $b = New-PlainButton -Parent $Parent -Text $Text
    if ($Primary) { $b.Font = $script:UI.FontBold }
    if ($Danger) { $b.ForeColor = [System.Drawing.Color]::DarkRed }
    Register-ControlHandler -Control $b -EventName 'Click' -Module $Module -Action $OnClick
    [void]$Module.Buttons.Add($b)
    return $b
}

function Add-ToolbarRow([hashtable]$Module) {
    $row = New-FlowRow
    [void]$Module.TopControls.Add($row)
    return $row
}

function Add-TopControl([hashtable]$Module, $Control) {
    [void]$Module.TopControls.Add($Control)
    return $Control
}
#endregion

#region Okna dialogowe
function New-DialogForm {
    param([string]$Title, [int]$Width = 520, [int]$Height = 300, [switch]$Resizable)
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title
    $f.Size = New-Object System.Drawing.Size($Width, $Height)
    $f.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $f.MinimizeBox = $false
    $f.MaximizeBox = [bool]$Resizable
    $f.ShowInTaskbar = $false
    $f.Font = $script:UI.Font
    $f.Padding = New-Object System.Windows.Forms.Padding(10)
    if (-not $Resizable) { $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog }
    return $f
}

function Add-DialogButtons {
    param([System.Windows.Forms.Form]$Form, [string]$OkText = 'OK', [string]$CancelText = 'Anuluj')
    $panel = New-Object System.Windows.Forms.FlowLayoutPanel
    $panel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $panel.AutoSize = $true
    $panel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $panel.WrapContents = $false
    $panel.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $cancel = New-PlainButton -Parent $null -Text $CancelText
    $cancel.MinimumSize = New-Object System.Drawing.Size(90, 28)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $ok = New-PlainButton -Parent $null -Text $OkText
    $ok.MinimumSize = New-Object System.Drawing.Size(90, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $panel.Controls.Add($cancel)
    $panel.Controls.Add($ok)
    $Form.AcceptButton = $ok
    $Form.CancelButton = $cancel
    return @{ Panel = $panel; Ok = $ok; Cancel = $cancel }
}

function New-FormGrid {
    # Dwukolumnowa tabela "etykieta: pole" do prostych formularzy
    param([object[]]$Rows)
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.AutoSize = $true
    $t.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $t.ColumnCount = 2
    $t.RowCount = $Rows.Count
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    for ($i = 0; $i -lt $Rows.Count; $i++) {
        [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
        $label = Add-Label -Parent $null -Text $Rows[$i][0]
        $control = $Rows[$i][1]
        $control.Anchor = [System.Windows.Forms.AnchorStyles]'Left,Right'
        $t.Controls.Add($label, 0, $i)
        $t.Controls.Add($control, 1, $i)
    }
    return $t
}

function Show-CredentialDialog {
    param([string]$Message = 'Podaj poświadczenia konta z uprawnieniami administracyjnymi.', [string]$UserName = '')
    $dlg = New-DialogForm -Title 'Poświadczenia' -Width 460 -Height 250
    $lbl = Add-Label -Parent $null -Text $Message
    $lbl.AutoSize = $false
    $lbl.Height = 40
    $txtUser = Add-TextBox -Parent $null -Width 260 -Text $UserName
    $txtPass = Add-TextBox -Parent $null -Width 260
    $txtPass.UseSystemPasswordChar = $true
    $grid = New-FormGrid -Rows @(@('Użytkownik:', $txtUser), @('Hasło:', $txtPass))
    $hint = Add-Label -Parent $null -Text 'Format: DOMENA\login albo login@domena' -Hint
    $buttons = Add-DialogButtons -Form $dlg
    $buttons.Ok.DialogResult = [System.Windows.Forms.DialogResult]::None
    $buttons.Ok.Add_Click({
            if ([string]::IsNullOrWhiteSpace($txtUser.Text)) { Show-Warning 'Podaj nazwę użytkownika.'; return }
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        })
    Add-DockStack -Parent $dlg -Top @($lbl, $grid, $hint) -Bottom @($buttons.Panel)
    $dlg.Add_Shown({ if ($txtUser.Text) { [void]$txtPass.Focus() } else { [void]$txtUser.Focus() } })
    $result = $null
    if ($dlg.ShowDialog($script:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $secure = New-Object System.Security.SecureString
        foreach ($ch in $txtPass.Text.ToCharArray()) { $secure.AppendChar($ch) }
        $secure.MakeReadOnly()
        $result = New-Object System.Management.Automation.PSCredential($txtUser.Text.Trim(), $secure)
    }
    $dlg.Dispose()
    return $result
}

function Show-PasswordDialog {
    # Nowe hasło wpisane dwukrotnie; zwraca SecureString albo $null
    param([string]$Message = 'Podaj nowe hasło.')
    $dlg = New-DialogForm -Title 'Nowe hasło' -Width 460 -Height 260
    $lbl = Add-Label -Parent $null -Text $Message
    $lbl.AutoSize = $false
    $lbl.Height = 40
    $txt1 = Add-TextBox -Parent $null -Width 260
    $txt1.UseSystemPasswordChar = $true
    $txt2 = Add-TextBox -Parent $null -Width 260
    $txt2.UseSystemPasswordChar = $true
    $grid = New-FormGrid -Rows @(@('Hasło:', $txt1), @('Powtórz hasło:', $txt2))
    $buttons = Add-DialogButtons -Form $dlg
    $buttons.Ok.DialogResult = [System.Windows.Forms.DialogResult]::None
    $buttons.Ok.Add_Click({
            if (-not $txt1.Text) { Show-Warning 'Hasło nie może być puste.'; return }
            if ($txt1.Text -cne $txt2.Text) { Show-Warning 'Hasła nie są identyczne.'; return }
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        })
    Add-DockStack -Parent $dlg -Top @($lbl, $grid) -Bottom @($buttons.Panel)
    $result = $null
    if ($dlg.ShowDialog($script:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $result = New-Object System.Security.SecureString
        foreach ($ch in $txt1.Text.ToCharArray()) { $result.AppendChar($ch) }
        $result.MakeReadOnly()
    }
    $dlg.Dispose()
    return $result
}

function Show-InputDialog {
    param([string]$Title, [string]$Prompt, [string]$Default = '', [switch]$Multiline)
    $height = if ($Multiline) { 440 } else { 190 }
    $dlg = New-DialogForm -Title $Title -Width 520 -Height $height -Resizable:$Multiline
    $lbl = Add-Label -Parent $null -Text $Prompt
    $lbl.AutoSize = $false
    $lbl.Height = 40
    $txt = Add-TextBox -Parent $null -Width 300 -Text $Default
    if ($Multiline) {
        $txt.Multiline = $true
        $txt.AcceptsReturn = $true
        $txt.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $txt.Font = $script:UI.FontMono
    }
    $buttons = Add-DialogButtons -Form $dlg
    if ($Multiline) {
        $dlg.AcceptButton = $null
        Add-DockStack -Parent $dlg -Top @($lbl) -Fill $txt -Bottom @($buttons.Panel)
    }
    else {
        Add-DockStack -Parent $dlg -Top @($lbl, $txt) -Bottom @($buttons.Panel)
    }
    $result = $null
    if ($dlg.ShowDialog($script:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) { $result = $txt.Text }
    $dlg.Dispose()
    return $result
}

function Show-TextDialog {
    param([string]$Title, [string]$Text)
    $dlg = New-DialogForm -Title $Title -Width 900 -Height 600 -Resizable
    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Multiline = $true
    $txt.ReadOnly = $true
    $txt.BackColor = [System.Drawing.SystemColors]::Window
    $txt.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $txt.WordWrap = $false
    $txt.Font = $script:UI.FontMono
    $txt.Text = ($Text -replace "`r?`n", "`r`n")
    $buttons = Add-DialogButtons -Form $dlg -OkText 'Zamknij' -CancelText 'Kopiuj'
    $buttons.Cancel.DialogResult = [System.Windows.Forms.DialogResult]::None
    $buttons.Cancel.Add_Click({ if ($txt.Text) { [System.Windows.Forms.Clipboard]::SetText($txt.Text) } })
    $dlg.CancelButton = $buttons.Ok
    Add-DockStack -Parent $dlg -Fill $txt -Bottom @($buttons.Panel)
    $dlg.Add_Shown({ $txt.SelectionLength = 0 })
    [void]$dlg.ShowDialog($script:UI.Form)
    $dlg.Dispose()
}

function Show-GridDialog {
    # Wyświetla dowolne obiekty w tabeli z filtrem, eksportem i kopiowaniem
    param([string]$Title, [object[]]$Rows, [string[]]$SecretColumns = @())
    $key = 'Dialog_' + [guid]::NewGuid().ToString('N')
    $m = New-ModuleContext -Definition @{ Key = $key; Title = $Title; Description = ''; Category = '' }
    $m.SecretColumns = @($SecretColumns)
    $script:UI.Modules[$key] = $m
    $dlg = $null
    try {
        $dlg = New-DialogForm -Title $Title -Width 960 -Height 560 -Resizable
        New-ResultView -Module $m
        foreach ($r in $Rows) {
            Add-ResultRows -Module $m -Computer ([string](Get-ObjectValue $r 'Komputer')) -Objects @($r)
        }
        Resize-ResultColumns -Module $m
        $buttons = Add-DialogButtons -Form $dlg -OkText 'Zamknij'
        $buttons.Cancel.Visible = $false
        $dlg.CancelButton = $buttons.Ok
        Add-DockStack -Parent $dlg -Top @($m.ResultBar) -Fill $m.Grid -Bottom @($buttons.Panel)
        [void]$dlg.ShowDialog($script:UI.Form)
    }
    finally {
        $script:UI.Modules.Remove($key)
        if ($dlg) { $dlg.Dispose() }
    }
}

function Select-OrganizationalUnit {
    # Wybór OU z drzewa domeny. Zwraca DN, '' (cała domena - tylko z -AllowDomainRoot) albo $null (anulowano).
    param([string]$Title = 'Wybierz jednostkę organizacyjną', [string]$Selected = '', [switch]$AllowDomainRoot)
    Import-AdModule
    $ad = Get-AdSplat
    $script:UI.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $domain = Get-ADDomain @ad
        $dns = @(Get-ADOrganizationalUnit -Filter * @ad | ForEach-Object { $_.DistinguishedName })
        $dns += [string]$domain.ComputersContainer
    }
    finally {
        $script:UI.Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }

    $dlg = New-DialogForm -Title $Title -Width 520 -Height 620 -Resizable
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.HideSelection = $false
    $tree.Font = $script:UI.Font
    $root = $tree.Nodes.Add([string]$domain.DNSRoot)
    $root.Tag = [string]$domain.DistinguishedName
    $map = @{ ([string]$domain.DistinguishedName).ToLowerInvariant() = $root }
    $sorted = $dns | Where-Object { $_ } | Sort-Object { @($_ -split '(?<!\\),').Count }, { $_ }
    foreach ($dn in $sorted) {
        $parent = $map[(Get-ParentDN $dn).ToLowerInvariant()]
        if (-not $parent) { $parent = $root }
        $node = $parent.Nodes.Add((Get-RdnValue $dn))
        $node.Tag = $dn
        $map[$dn.ToLowerInvariant()] = $node
        if ($Selected -and $dn -eq $Selected) { $tree.SelectedNode = $node }
    }
    $root.Expand()
    if (-not $tree.SelectedNode) { $tree.SelectedNode = $root }
    if ($tree.SelectedNode) { $tree.SelectedNode.EnsureVisible() }

    $hintText = if ($AllowDomainRoot) { 'Zaznacz OU lub korzeń domeny (cała domena).' } else { 'Zaznacz docelową jednostkę organizacyjną.' }
    $hint = Add-Label -Parent $null -Text $hintText -Hint
    $buttons = Add-DialogButtons -Form $dlg
    $buttons.Ok.DialogResult = [System.Windows.Forms.DialogResult]::None
    $buttons.Ok.Add_Click({
            if (-not $tree.SelectedNode) { return }
            if (-not $AllowDomainRoot -and $tree.SelectedNode -eq $root) { Show-Warning 'Wybierz jednostkę organizacyjną (nie korzeń domeny).'; return }
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
        })
    $tree.Add_NodeMouseDoubleClick({ $buttons.Ok.PerformClick() })
    Add-DockStack -Parent $dlg -Top @($hint) -Fill $tree -Bottom @($buttons.Panel)
    $result = $null
    if ($dlg.ShowDialog($script:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $result = if ($tree.SelectedNode -eq $root) { '' } else { [string]$tree.SelectedNode.Tag }
    }
    $dlg.Dispose()
    return $result
}
#endregion

#region Tabela wyników (DataTable + DataView + DataGridView)
$script:GridEvents = @{
    CellDoubleClick     = {
        param($s, $e)
        try {
            if ($e.RowIndex -lt 0) { return }
            $m = $script:UI.Modules[[string]$s.Tag]
            if ($m) { Show-RowDetails -Module $m -RowIndex $e.RowIndex }
        }
        catch { Write-Log "Nie można wyświetlić szczegółów: $($_.Exception.Message)" 'ERROR' }
    }
    CellFormatting      = {
        param($s, $e)
        if ($e.RowIndex -lt 0 -or $null -eq $e.Value -or $e.Value -is [System.DBNull]) { return }
        try {
            $m = $script:UI.Modules[[string]$s.Tag]
            if (-not $m) { return }
            $name = $s.Columns[$e.ColumnIndex].DataPropertyName
            if ($m.SecretColumns.Count -gt 0 -and $m.SecretColumns -contains $name) {
                if (-not $m.RevealSecrets -and [string]$e.Value -ne '') {
                    $e.Value = '••••••••••'
                    $e.FormattingApplied = $true
                }
                return
            }
            if ($name -eq 'Status' -or $name -eq 'Wynik') {
                if (([string]$e.Value).StartsWith('Błąd')) { $e.CellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
                return
            }
            if ($m.ColorBools) {
                $v = [string]$e.Value
                if ($v -eq 'Tak') { $e.CellStyle.ForeColor = [System.Drawing.Color]::ForestGreen }
                elseif ($v -eq 'Nie') { $e.CellStyle.ForeColor = [System.Drawing.Color]::Firebrick }
            }
        }
        catch { }
    }
    DataBindingComplete = {
        param($s, $e)
        try { if ($s.Columns.Contains('__search')) { $s.Columns['__search'].Visible = $false } } catch { }
    }
    DataError           = {
        param($s, $e)
        $e.ThrowException = $false
    }
}

function New-ResultGrid {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.ReadOnly = $true
    $g.AllowUserToAddRows = $false
    $g.AllowUserToDeleteRows = $false
    $g.AllowUserToResizeRows = $false
    $g.AllowUserToOrderColumns = $true
    $g.RowHeadersVisible = $false
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $g.MultiSelect = $true
    $g.AutoGenerateColumns = $true
    $g.BackgroundColor = [System.Drawing.SystemColors]::Window
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $g.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
    $g.ColumnHeadersDefaultCellStyle.Font = $script:UI.FontBold
    $g.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 251)
    $g.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $g.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
    $g.RowTemplate.Height = 22
    try {
        $prop = [System.Windows.Forms.DataGridView].GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance,NonPublic')
        $prop.SetValue($g, $true, $null)
    }
    catch { }
    $g.add_CellDoubleClick($script:GridEvents.CellDoubleClick)
    $g.add_CellFormatting($script:GridEvents.CellFormatting)
    $g.add_DataBindingComplete($script:GridEvents.DataBindingComplete)
    $g.add_DataError($script:GridEvents.DataError)
    return $g
}

function New-ResultView {
    # Pasek nad tabelą (filtr, licznik, eksport, kopiowanie) + sama tabela
    param([hashtable]$Module)
    $bar = New-FlowRow
    $bar.Padding = New-Object System.Windows.Forms.Padding(0, 8, 0, 2)
    [void](Add-Label -Parent $bar -Text 'Filtr wyników:')
    $filter = Add-TextBox -Parent $bar -Width 220
    Register-ControlHandler -Control $filter -EventName 'TextChanged' -Module $Module -Action { param($m) Update-ResultFilter -Module $m }
    $count = Add-Label -Parent $bar -Text '0 wierszy' -Hint
    $count.MinimumSize = New-Object System.Drawing.Size(120, 0)
    $btnExport = New-PlainButton -Parent $bar -Text 'Eksport CSV…'
    Register-ControlHandler -Control $btnExport -EventName 'Click' -Module $Module -Action { param($m) Export-ResultView -Module $m }
    $btnCopy = New-PlainButton -Parent $bar -Text 'Kopiuj'
    Register-ControlHandler -Control $btnCopy -EventName 'Click' -Module $Module -Action { param($m) Copy-ResultView -Module $m }
    if ($Module.SecretColumns.Count -gt 0) {
        $chk = Add-CheckBox -Parent $bar -Text 'Pokaż wartości poufne'
        Register-ControlHandler -Control $chk -EventName 'CheckedChanged' -Module $Module -Action {
            param($m, $s)
            $m.RevealSecrets = $s.Checked
            $m.Grid.Invalidate()
        }
    }
    [void](Add-Label -Parent $bar -Text 'Dwuklik na wierszu – szczegóły' -Hint)

    $grid = New-ResultGrid
    $grid.Tag = $Module.Key
    $Module.Grid = $grid
    $Module.ResultBar = $bar
    $Module.FilterBox = $filter
    $Module.CountLabel = $count
    Reset-ResultTable -Module $Module
}

function Reset-ResultTable {
    param([hashtable]$Module)
    $table = New-Object System.Data.DataTable 'Wyniki'
    # Ukryta kolumna z tekstem do wyszukiwania (filtr działa po wszystkich kolumnach naraz)
    [void]$table.Columns.Add('__search', [string])
    # Hidden: DataGridView nie generuje dla niej kolumny, a RowFilter nadal może z niej korzystać
    $table.Columns['__search'].ColumnMapping = [System.Data.MappingType]::Hidden
    $view = [System.Data.DataView]::new($table)
    $Module.Table = $table
    $Module.View = $view
    if ($Module.Grid) {
        $Module.Grid.DataSource = $null
        $Module.Grid.Columns.Clear()
        $Module.Grid.DataSource = $view
    }
    Update-ResultFilter -Module $Module
}

function Add-ResultRows {
    # Dodaje obiekty jako wiersze tabeli modułu; nowe właściwości tworzą nowe kolumny
    param([hashtable]$Module, [string]$Computer, [object[]]$Objects)
    $table = $Module.Table
    if ($null -eq $table) { return }
    $wasEmpty = ($table.Rows.Count -eq 0)
    $table.BeginLoadData()
    try {
        foreach ($obj in $Objects) {
            if ($null -eq $obj) { continue }
            $values = New-Object System.Collections.Specialized.OrderedDictionary
            if ($Computer) { $values['Komputer'] = $Computer }
            $base = if ($obj -is [System.Management.Automation.PSObject]) { $obj.PSObject.BaseObject } else { $obj }
            if ($base -is [string] -or $base -is [System.ValueType]) {
                $values['Wynik'] = ConvertTo-CellValue $base
            }
            else {
                foreach ($p in $obj.PSObject.Properties) {
                    if ($script:HiddenProperties -contains $p.Name) { continue }
                    if ($Computer -and $p.Name -eq 'Komputer') { continue }
                    $values[$p.Name] = ConvertTo-CellValue $p.Value
                }
            }
            foreach ($name in $values.Keys) {
                if (-not $table.Columns.Contains($name)) { [void]$table.Columns.Add($name, [object]) }
            }
            $row = $table.NewRow()
            $search = New-Object System.Text.StringBuilder
            foreach ($name in $values.Keys) {
                $v = $values[$name]
                $row[$name] = $v
                if ($v -isnot [System.DBNull] -and $Module.SecretColumns -notcontains $name) {
                    [void]$search.Append([string]$v).Append(' ')
                }
            }
            $row['__search'] = $search.ToString().ToLowerInvariant()
            $table.Rows.Add($row)
        }
    }
    finally {
        $table.EndLoadData()
    }
    if ($Module.Grid -and $Module.Grid.Columns.Contains('__search')) { $Module.Grid.Columns['__search'].Visible = $false }
    if ($wasEmpty -and $table.Rows.Count -gt 0) { Resize-ResultColumns -Module $Module }
    Update-ResultCount -Module $Module
}

function Update-ResultFilter {
    param([hashtable]$Module)
    # Uwaga: DataView jest dla PowerShella listą - pusty widok byłby fałszem, stąd porównanie z $null
    if ($null -eq $Module.View) { return }
    $text = if ($Module.FilterBox) { $Module.FilterBox.Text.Trim().ToLowerInvariant() } else { '' }
    $terms = @($text -split '\s+' | Where-Object { $_ })
    $filter = (@($terms | ForEach-Object { "__search LIKE '*{0}*'" -f (ConvertTo-LikeLiteral $_) }) -join ' AND ')
    try { $Module.View.RowFilter = $filter } catch { $Module.View.RowFilter = '' }
    Update-ResultCount -Module $Module
}

function Update-ResultCount {
    param([hashtable]$Module)
    if ($null -eq $Module.CountLabel -or $null -eq $Module.View) { return }
    $total = $Module.Table.Rows.Count
    $visible = $Module.View.Count
    $Module.CountLabel.Text = if ($visible -eq $total) { "$total wierszy" } else { "$visible z $total wierszy" }
}

function Resize-ResultColumns {
    param([hashtable]$Module)
    $g = $Module.Grid
    if (-not $g -or $g.Columns.Count -eq 0) { return }
    try {
        $g.AutoResizeColumns([System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::DisplayedCells)
        foreach ($c in $g.Columns) {
            if ($c.Width -gt 420) { $c.Width = 420 }
            if ($c.Width -lt 60) { $c.Width = 60 }
        }
    }
    catch { }
}

function Get-VisibleColumnNames {
    param([hashtable]$Module)
    $cols = @($Module.Grid.Columns | Where-Object { $_.Visible -and $_.DataPropertyName -ne '__search' } | Sort-Object DisplayIndex)
    return @($cols | ForEach-Object { $_.DataPropertyName })
}

function Get-ExportValue {
    # Wartość komórki do eksportu/kopiowania (z maskowaniem wartości poufnych)
    param([hashtable]$Module, $RowView, [string]$Column)
    $v = Get-ObjectValue $RowView $Column
    if ($null -eq $v) { return '' }
    if ($Module.SecretColumns -contains $Column -and -not $Module.RevealSecrets -and [string]$v -ne '') { return '********' }
    return $v
}

function Get-SelectedResultRows {
    # Zaznaczone wiersze tabeli (DataRowView); gdy nic nie zaznaczono - bieżący wiersz
    param([hashtable]$Module)
    $g = $Module.Grid
    $rows = @(foreach ($r in $g.SelectedRows) { if ($r.DataBoundItem) { $r.DataBoundItem } })
    if ($rows.Count -eq 0 -and $g.CurrentRow -and $g.CurrentRow.DataBoundItem) { $rows = @($g.CurrentRow.DataBoundItem) }
    return $rows
}

function Get-SelectedRowsByHost {
    # Grupuje zaznaczone wiersze po kolumnie Komputer: host -> lista hashtabel z wartościami kolumn
    param([hashtable]$Module, [string[]]$Columns)
    $result = [ordered]@{}
    foreach ($drv in @(Get-SelectedResultRows -Module $Module)) {
        $computer = [string](Get-ObjectValue $drv 'Komputer')
        if (-not $computer -or (Get-ObjectValue $drv 'Status') -eq 'Błąd') { continue }
        $item = @{}
        $valid = $true
        foreach ($c in $Columns) {
            $v = Get-ObjectValue $drv $c
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
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv'
    $dlg.FileName = '{0}_{1:yyyyMMdd_HHmm}.csv' -f ($Module.Title -replace '[\\/:*?"<>|\s]+', '_'), (Get-Date)
    if ($dlg.ShowDialog($script:UI.Form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $columns = Get-VisibleColumnNames -Module $Module
    $objects = foreach ($drv in $Module.View) {
        $o = [ordered]@{}
        foreach ($c in $columns) { $o[$c] = Get-ExportValue -Module $Module -RowView $drv -Column $c }
        [pscustomobject]$o
    }
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $objects | Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -UseCulture -Encoding $encoding
    Write-Log "Zapisano $(@($objects).Count) wierszy do pliku $($dlg.FileName)" 'OK'
}

function Copy-ResultView {
    # Kopiuje zaznaczone wiersze (albo wszystkie widoczne) jako tekst rozdzielany tabulatorami (Excel)
    param([hashtable]$Module)
    if ($null -eq $Module.View -or $Module.View.Count -eq 0) { Show-Warning 'Brak danych do skopiowania.'; return }
    $columns = Get-VisibleColumnNames -Module $Module
    $rows = @(foreach ($r in $Module.Grid.SelectedRows) { if ($r.DataBoundItem) { $r.DataBoundItem } })
    if ($rows.Count -le 1) { $rows = @($Module.View | ForEach-Object { $_ }) }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(($columns -join "`t"))
    foreach ($drv in $rows) {
        $cells = foreach ($c in $columns) { ([string](Get-ExportValue -Module $Module -RowView $drv -Column $c)) -replace '[\t\r\n]+', ' ' }
        [void]$sb.AppendLine((@($cells) -join "`t"))
    }
    [System.Windows.Forms.Clipboard]::SetText($sb.ToString())
    Write-Log "Skopiowano do schowka $($rows.Count) wierszy." 'OK'
}

function Show-RowDetails {
    param([hashtable]$Module, [int]$RowIndex)
    $drv = $Module.Grid.Rows[$RowIndex].DataBoundItem
    if (-not $drv) { return }
    $sb = New-Object System.Text.StringBuilder
    foreach ($col in $Module.Table.Columns) {
        $name = $col.ColumnName
        if ($name -eq '__search') { continue }
        $v = Get-ExportValue -Module $Module -RowView $drv -Column $name
        if ([string]$v -eq '') { continue }
        $text = [string]$v
        if ($text -match "[\r\n]") { [void]$sb.AppendLine("${name}:").AppendLine($text).AppendLine() }
        else { [void]$sb.AppendLine(('{0}: {1}' -f $name, $text)) }
    }
    $title = 'Szczegóły'
    $computer = Get-ObjectValue $drv 'Komputer'
    if ($computer) { $title = "Szczegóły – $computer" }
    Show-TextDialog -Title $title -Text $sb.ToString()
}
#endregion

#region Silnik operacji w tle
# Każdy host to osobne zadanie w puli wątków. Tryb 'Remote' wykonuje blok skryptu na hoście przez
# Invoke-Command (blok dostaje jeden parametr: hashtablę $P). Tryb 'Local' wykonuje blok lokalnie
# z parametrami ($Target, $P, $Ctx) - np. dla poleceń AD albo kopiowania plików.
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

function Set-EngineThrottle([int]$Limit) {
    $script:Settings.ThrottleLimit = $Limit
    if ($script:Engine.Pool) {
        try { [void]$script:Engine.Pool.SetMaxRunspaces($Limit) } catch { }
    }
}

function Set-ModuleBusy {
    param([hashtable]$Module, [bool]$Busy)
    $Module.Busy = $Busy
    foreach ($b in $Module.Buttons) { $b.Enabled = -not $Busy }
}

function New-SessionOption {
    # Odpowiednik New-PSSessionOption -OpenTimeout (obiekt tworzony bezpośrednio)
    $option = New-Object System.Management.Automation.Remoting.PSSessionOption
    $option.OpenTimeout = [TimeSpan]::FromSeconds([int]$script:Settings.TimeoutSec)
    return $option
}

function Start-HostOperation {
    <#
        Uruchamia blok skryptu dla każdego hosta z -Targets.
        -Output Grid : wyniki trafiają do tabeli modułu (kolumna Komputer + właściwości obiektów)
        -Output Log  : wyniki wypisywane są w dzienniku
        -Output None : tylko -OnResult/-OnComplete
        -PerTarget   : osobna hashtabla $P dla wybranych hostów (np. różne usługi na różnych komputerach)
        -OnResult { param($m, $r) }   - po zakończeniu każdego hosta ($r: Target, Ok, Data, Errors)
        -OnComplete { param($m, $op) } - po zakończeniu wszystkich hostów (nie wywoływane po anulowaniu)
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
        [switch]$Append,
        [scriptblock]$OnResult,
        [scriptblock]$OnComplete
    )
    if ($Module.Busy) {
        Show-Warning "Poprzednia operacja w module «$($Module.Title)» jeszcze trwa. Poczekaj na jej zakończenie lub anuluj ją na pasku stanu."
        return
    }
    $hosts = @($Targets | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Select-Object -Unique)
    if ($hosts.Count -eq 0) { return }

    Initialize-Engine
    if ($Output -eq 'Grid' -and -not $Append -and $Module.Grid) { Reset-ResultTable -Module $Module }

    $ctx = @{
        Credential    = Get-EffectiveCredential
        Server        = [string]$script:Settings.DomainController
        SessionOption = New-SessionOption
    }
    $mode = if ($Local) { 'Local' } else { 'Remote' }
    $scriptText = $ScriptBlock.ToString()

    $op = @{
        Id         = $script:Engine.NextId
        Name       = $Name
        Module     = $Module
        Output     = $Output
        OnResult   = $OnResult
        OnComplete = $OnComplete
        Items      = New-Object System.Collections.ArrayList
        Total      = $hosts.Count
        Done       = 0
        Failed     = 0
        Cancelled  = $false
        Started    = Get-Date
    }
    $script:Engine.NextId++

    foreach ($t in $hosts) {
        $p = $Parameters
        if ($PerTarget.ContainsKey($t)) { $p = $PerTarget[$t] }
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.RunspacePool = $script:Engine.Pool
        [void]$ps.AddScript($script:WorkerScript)
        [void]$ps.AddArgument($t).AddArgument($mode).AddArgument($scriptText).AddArgument($p).AddArgument($ctx)
        $handle = $ps.BeginInvoke()
        [void]$op.Items.Add(@{ Target = $t; PS = $ps; Handle = $handle; Finished = $false })
    }

    Set-ModuleBusy -Module $Module -Busy $true
    [void]$script:Engine.Operations.Add($op)
    $list = (@($hosts | Select-Object -First 8) -join ', ')
    if ($hosts.Count -gt 8) { $list += ", … (+$($hosts.Count - 8))" }
    Write-Log ("{0} – start dla {1} host(ów): {2}" -f $Name, $hosts.Count, $list) -Module $Module.Title
    $script:Engine.Timer.Start()
    Update-StatusBar
}

function Update-Operations {
    # Wywoływane przez timer w wątku okna
    if ($script:Engine.InTick) { return }
    $script:Engine.InTick = $true
    try {
        foreach ($op in @($script:Engine.Operations)) {
            foreach ($item in $op.Items) {
                if ($item.Finished -or -not $item.Handle.IsCompleted) { continue }
                $item.Finished = $true
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
            if ($op.Done -ge $op.Total) { Complete-Operation -Operation $op }
        }
    }
    catch {
        Write-Log "Błąd silnika operacji: $($_.Exception.Message)" 'ERROR' -Module ''
    }
    finally {
        $script:Engine.InTick = $false
        if ($script:Engine.Operations.Count -eq 0) { $script:Engine.Timer.Stop() }
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
                Add-ResultRows -Module $m -Computer $Result.Target -Objects @([pscustomobject]@{ 'Status' = 'Błąd'; 'Szczegóły' = $errorText })
            }
        }
        else {
            if ($errorText) { Write-Log ("[{0}] ostrzeżenia: {1}" -f $Result.Target, $errorText) 'WARN' }
            $data = @($Result.Data | Where-Object { $null -ne $_ })
            switch ($Operation.Output) {
                'Grid' {
                    if ($data.Count -gt 0) { Add-ResultRows -Module $m -Computer $Result.Target -Objects $data }
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
        Write-Log ("{0} – anulowano (zakończone: {1}, czas {2} s)" -f $Operation.Name, $okCount, $seconds) 'WARN' -Module $m.Title
    }
    else {
        $level = if ($Operation.Failed -gt 0) { 'WARN' } else { 'OK' }
        Write-Log ("{0} – zakończono: {1} OK, {2} z błędem, czas {3} s" -f $Operation.Name, $okCount, $Operation.Failed, $seconds) $level -Module $m.Title
    }
    if ($Operation.Output -eq 'Grid' -and $m.Grid) { Resize-ResultColumns -Module $m }
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
    Write-Log 'Anulowanie trwających operacji…' 'WARN' -Module ''
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
    $label = $script:UI['StatusLabel']
    if (-not $label) { return }
    $ops = @($script:Engine.Operations)
    if ($ops.Count -eq 0) {
        $label.Text = 'Gotowe'
        $script:UI.StatusProgress.Visible = $false
        $script:UI.StatusCancel.Enabled = $false
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
    $bar = $script:UI.StatusProgress
    $bar.Maximum = [Math]::Max(1, $total)
    $bar.Value = [Math]::Min($done, $bar.Maximum)
    $bar.Visible = $true
    $script:UI.StatusCancel.Enabled = $true
}

function Get-TargetComputers {
    # Komputery zaznaczone na liście po lewej
    param([switch]$Quiet)
    $rows = @($script:UI.HostTable.Select('Sel = true', 'Name ASC'))
    $names = @($rows | ForEach-Object { [string]$_['Name'] } | Where-Object { $_ } | Select-Object -Unique)
    if ($names.Count -eq 0 -and -not $Quiet) { Show-Warning 'Zaznacz komputery na liście po lewej stronie.' }
    return $names
}
#endregion

#region Moduły (rejestracja i budowa)
function Register-Module {
    param(
        [Parameter(Mandatory)][string]$Key,
        [Parameter(Mandatory)][string]$Category,
        [Parameter(Mandatory)][string]$Title,
        [string]$Description = '',
        [Parameter(Mandatory)][scriptblock]$Build
    )
    [void]$script:UI.ModuleDefs.Add(@{ Key = $Key; Category = $Category; Title = $Title; Description = $Description; Build = $Build })
}

function New-ModuleContext {
    param([hashtable]$Definition)
    return @{
        Key           = $Definition.Key
        Title         = $Definition.Title
        Description   = $Definition.Description
        Category      = $Definition.Category
        Busy          = $false
        Buttons       = New-Object System.Collections.ArrayList
        TopControls   = New-Object System.Collections.ArrayList
        Actions       = @{}
        SecretColumns = @()
        RevealSecrets = $false
        ColorBools    = $false
        Root          = $null
        Grid          = $null
        Table         = $null
        View          = $null
        ResultBar     = $null
        FilterBox     = $null
        CountLabel    = $null
    }
}

function Initialize-Module {
    param([hashtable]$Definition)
    $m = New-ModuleContext -Definition $Definition
    $script:UI.Modules[$Definition.Key] = $m

    $root = New-Object System.Windows.Forms.Panel
    $root.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 8)
    $root.Visible = $false
    $m.Root = $root

    $title = New-Object System.Windows.Forms.Label
    $title.Text = $Definition.Title
    $title.Font = New-Object System.Drawing.Font('Segoe UI', 13, [System.Drawing.FontStyle]::Bold)
    $title.ForeColor = [System.Drawing.Color]::FromArgb(30, 60, 110)
    $title.AutoSize = $false
    $title.Height = 32
    $title.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    [void]$m.TopControls.Add($title)
    if ($Definition.Description) {
        $desc = New-Object System.Windows.Forms.Label
        $desc.Text = $Definition.Description
        $desc.ForeColor = [System.Drawing.Color]::DimGray
        $desc.AutoSize = $false
        $desc.Height = 36
        [void]$m.TopControls.Add($desc)
    }

    $null = & $Definition.Build $m

    New-ResultView -Module $m
    [void]$m.TopControls.Add($m.ResultBar)
    Add-DockStack -Parent $root -Top @($m.TopControls) -Fill $m.Grid
    $script:UI.ContentHost.Controls.Add($root)
    $root.Dock = [System.Windows.Forms.DockStyle]::Fill
    return $m
}

function Show-Module {
    param([string]$Key)
    $definition = $null
    foreach ($d in $script:UI.ModuleDefs) { if ($d.Key -eq $Key) { $definition = $d; break } }
    if (-not $definition) { return }
    $m = $script:UI.Modules[$Key]
    if (-not $m) {
        $script:UI.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try { $m = Initialize-Module -Definition $definition }
        finally { $script:UI.Form.Cursor = [System.Windows.Forms.Cursors]::Default }
    }
    $active = $script:UI.ActiveModule
    if ($active -and -not [object]::ReferenceEquals($active, $m)) { $active.Root.Visible = $false }
    $m.Root.Visible = $true
    $m.Root.BringToFront()
    $script:UI.ActiveModule = $m
    $script:Settings.LastModule = $Key
}
#endregion

#region Lista komputerów (lewy panel)
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

function New-HostTable {
    $t = New-Object System.Data.DataTable 'Hosts'
    [void]$t.Columns.Add('Sel', [bool])
    [void]$t.Columns.Add('Name', [string])
    [void]$t.Columns.Add('OS', [string])
    [void]$t.Columns.Add('LastLogon', [datetime])
    [void]$t.Columns.Add('Enabled', [string])
    [void]$t.Columns.Add('DNSHostName', [string])
    [void]$t.Columns.Add('DN', [string])
    [void]$t.Columns.Add('Source', [string])
    $t.Columns['Sel'].DefaultValue = $false
    $t.PrimaryKey = [System.Data.DataColumn[]]@($t.Columns['Name'])
    # Przecinek: PowerShell rozwijałby DataTable na wiersze (pusta tabela dałaby $null)
    return , $t
}

function ConvertTo-DbValue($Value) {
    if ($null -eq $Value -or [string]$Value -eq '') { return [System.DBNull]::Value }
    return $Value
}

function Import-HostRows {
    # Dodaje/aktualizuje komputery na liście; zachowuje zaznaczenia istniejących pozycji
    param([object[]]$Items, [string]$Source, [switch]$Check, [switch]$ReplaceSource)
    $t = $script:UI.HostTable
    $grid = $script:UI.HostGrid
    $checked = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $t.Select('Sel = true')) { [void]$checked.Add([string]$r['Name']) }
    $added = 0
    # Odpięcie tabeli od siatki na czas importu - tysiące wierszy dodają się wtedy błyskawicznie
    $grid.DataSource = $null
    try {
        if ($ReplaceSource) {
            foreach ($r in @($t.Select(("Source = '{0}'" -f $Source.Replace("'", "''"))))) { $t.Rows.Remove($r) }
        }
        $existing = @{}
        foreach ($r in $t.Rows) { $existing[[string]$r['Name']] = $r }
        foreach ($it in $Items) {
            $name = ([string](Get-ObjectValue $it 'Name')).Trim()
            if (-not $name) { continue }
            $row = $existing[$name]
            $isNew = ($null -eq $row)
            if ($isNew) {
                $row = $t.NewRow()
                $row['Name'] = $name
                $row['Sel'] = $false
                $row['Source'] = $Source
            }
            elseif ($Source -eq 'AD') { $row['Source'] = 'AD' }
            foreach ($col in @('OS', 'DNSHostName', 'DN')) {
                $v = Get-ObjectValue $it $col
                if ($null -ne $v) { $row[$col] = ConvertTo-DbValue ([string]$v) }
            }
            $lastLogon = Get-ObjectValue $it 'LastLogon'
            if ($lastLogon -is [datetime]) { $row['LastLogon'] = $lastLogon }
            $enabled = Get-ObjectValue $it 'Enabled'
            if ($null -ne $enabled) { $row['Enabled'] = $(if ([bool]$enabled) { 'Tak' } else { 'Nie' }) }
            if ($Check -or $checked.Contains($name)) { $row['Sel'] = $true }
            if ($isNew) {
                $t.Rows.Add($row)
                $existing[$name] = $row
                $added++
            }
        }
        $t.AcceptChanges()
    }
    finally {
        $grid.DataSource = $script:UI.HostView
    }
    Update-HostCount
    return $added
}

function Update-HostCount {
    $t = $script:UI.HostTable
    $selected = @($t.Select('Sel = true')).Count
    $script:UI.HostCountLabel.Text = 'Zaznaczone: {0} z {1}   (widoczne: {2})' -f $selected, $t.Rows.Count, $script:UI.HostView.Count
}

function Set-HostCheck {
    param([ValidateSet('CheckVisible', 'UncheckAll', 'InvertVisible', 'CheckSelected', 'UncheckSelected')][string]$Mode)
    $g = $script:UI.HostGrid
    [void]$g.EndEdit()
    $t = $script:UI.HostTable
    $t.BeginLoadData()
    try {
        switch ($Mode) {
            'CheckVisible' { foreach ($drv in @($script:UI.HostView | ForEach-Object { $_ })) { $drv.Row['Sel'] = $true } }
            'UncheckAll' { foreach ($r in $t.Rows) { $r['Sel'] = $false } }
            'InvertVisible' { foreach ($drv in @($script:UI.HostView | ForEach-Object { $_ })) { $drv.Row['Sel'] = -not [bool]$drv.Row['Sel'] } }
            'CheckSelected' { foreach ($gr in @($g.SelectedRows)) { if ($gr.DataBoundItem) { $gr.DataBoundItem.Row['Sel'] = $true } } }
            'UncheckSelected' { foreach ($gr in @($g.SelectedRows)) { if ($gr.DataBoundItem) { $gr.DataBoundItem.Row['Sel'] = $false } } }
        }
    }
    finally {
        $t.EndLoadData()
    }
    $g.Invalidate()
    Update-HostCount
}

function Update-HostQuickFilter {
    $text = $script:UI.HostSearch.Text.Trim()
    if ($text) {
        $lit = ConvertTo-LikeLiteral $text
        $script:UI.HostView.RowFilter = "Name LIKE '*{0}*' OR OS LIKE '*{0}*' OR Source LIKE '*{0}*'" -f $lit
    }
    else {
        $script:UI.HostView.RowFilter = ''
    }
    Update-HostCount
}

function Start-AdHostLoad {
    param([hashtable]$Module)
    $script:Settings.SearchBase = $script:UI.HostSearchBase.Text.Trim()
    $script:Settings.NameFilter = $script:UI.HostNameFilter.Text.Trim()
    $script:Settings.OnlyEnabled = $script:UI.HostOnlyEnabled.Checked
    $params = @{
        LdapFilter = New-ComputerLdapFilter -NamePattern $script:Settings.NameFilter -OnlyEnabled $script:Settings.OnlyEnabled
        SearchBase = $script:Settings.SearchBase
    }
    Start-HostOperation -Module $Module -Name 'Pobieranie komputerów z AD' -Targets @('Active Directory') -Local -Output None -Parameters $params -ScriptBlock {
        param($Target, $P, $Ctx)
        Import-Module ActiveDirectory -ErrorAction Stop
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
            $m.LastError = (@($r.Errors)) -join "`r`n"
            Invoke-Deferred -Module $m -Action { param($m) Show-Error 'Nie udało się pobrać komputerów z Active Directory.' $m.LastError }
            return
        }
        $items = @($r.Data)
        $added = Import-HostRows -Items $items -Source 'AD' -ReplaceSource
        Write-Log ("Wczytano z AD {0} komputerów (nowych na liście: {1})." -f $items.Count, $added) 'OK'
    }
}

function Add-ManualHosts {
    $text = Show-InputDialog -Title 'Dodaj komputery' -Prompt 'Wpisz lub wklej nazwy komputerów (po jednej w wierszu albo rozdzielone przecinkiem/spacją):' -Multiline
    if ($null -eq $text) { return }
    $names = @($text -split '[\s,;]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if ($names.Count -eq 0) { return }
    $added = Import-HostRows -Items @($names | ForEach-Object { [pscustomobject]@{ Name = $_ } }) -Source 'Ręcznie' -Check
    Write-Log ("Dodano ręcznie {0} komputerów (nowych: {1}); zostały zaznaczone." -f $names.Count, $added) 'OK' -Module 'Komputery'
}

function Import-HostFile {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Pliki tekstowe i CSV (*.txt;*.csv)|*.txt;*.csv|Wszystkie pliki (*.*)|*.*'
    if ($dlg.ShowDialog($script:UI.Form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $names = foreach ($line in (Get-Content -LiteralPath $dlg.FileName -Encoding UTF8)) {
        $first = (($line -split '[,;\t]')[0]).Trim().Trim('"')
        if ($first -and $first -notmatch '^(name|computer|computername|hostname|nazwa|komputer)$') { $first }
    }
    $names = @($names | Select-Object -Unique)
    if ($names.Count -eq 0) { Show-Warning 'Plik nie zawiera nazw komputerów.'; return }
    $added = Import-HostRows -Items @($names | ForEach-Object { [pscustomobject]@{ Name = $_ } }) -Source 'Plik' -Check
    Write-Log ("Wczytano z pliku {0} komputerów (nowych: {1}); zostały zaznaczone." -f $names.Count, $added) 'OK' -Module 'Komputery'
}

function Remove-SelectedHosts {
    $g = $script:UI.HostGrid
    $rows = @(foreach ($gr in $g.SelectedRows) { if ($gr.DataBoundItem) { $gr.DataBoundItem.Row } })
    if ($rows.Count -eq 0) { return }
    foreach ($r in $rows) { $script:UI.HostTable.Rows.Remove($r) }
    $script:UI.HostTable.AcceptChanges()
    Update-HostCount
}

function Rename-HostRow {
    # Po zmianie nazwy komputera aktualizuje pozycję na liście
    param([string]$OldName, [string]$NewName)
    $t = $script:UI.HostTable
    $row = $t.Rows.Find($OldName)
    if (-not $row -or $t.Rows.Find($NewName)) { return }
    $row['Name'] = $NewName
    $t.AcceptChanges()
}

function New-HostPanel {
    $hm = New-ModuleContext -Definition @{ Key = 'Hosts'; Title = 'Komputery'; Description = ''; Category = '' }
    $script:UI.Modules['Hosts'] = $hm

    $group = New-Object System.Windows.Forms.GroupBox
    $group.Text = 'Komputery docelowe'
    $group.Padding = New-Object System.Windows.Forms.Padding(8, 6, 8, 6)

    # OU / SearchBase
    $lblOu = Add-Label -Parent $null -Text 'Jednostka OU (puste = cała domena):'
    $lblOu.Dock = [System.Windows.Forms.DockStyle]::Top
    $txtBase = Add-TextBox -Parent $null -Width 200 -Text ([string]$script:Settings.SearchBase)
    $btnOu = New-PlainButton -Parent $null -Text 'Wybierz…'
    Register-ControlHandler -Control $btnOu -EventName 'Click' -Module $hm -Action {
        param($m)
        $dn = Select-OrganizationalUnit -Title 'Zakres wyszukiwania komputerów' -Selected $script:UI.HostSearchBase.Text.Trim() -AllowDomainRoot
        if ($null -ne $dn) { $script:UI.HostSearchBase.Text = $dn }
    }
    $rowOu = New-StretchRow -Stretch $txtBase -After @($btnOu)

    # Filtr nazwy + tylko włączone
    $rowFilter = New-FlowRow
    [void](Add-Label -Parent $rowFilter -Text 'Nazwa:')
    $txtName = Add-TextBox -Parent $rowFilter -Width 150 -Text ([string]$script:Settings.NameFilter)
    $chkEnabled = Add-CheckBox -Parent $rowFilter -Text 'Tylko włączone konta' -Checked ([bool]$script:Settings.OnlyEnabled)

    # Źródła listy
    $rowLoad = New-FlowRow
    [void](Add-Button -Parent $rowLoad -Text 'Wczytaj z AD' -Module $hm -Primary -OnClick { param($m) Start-AdHostLoad -Module $m })
    $btnManual = New-PlainButton -Parent $rowLoad -Text 'Dodaj ręcznie…'
    Register-ControlHandler -Control $btnManual -EventName 'Click' -Module $hm -Action { Add-ManualHosts }
    $btnFile = New-PlainButton -Parent $rowLoad -Text 'Z pliku…'
    Register-ControlHandler -Control $btnFile -EventName 'Click' -Module $hm -Action { Import-HostFile }

    # Szybkie wyszukiwanie na liście
    $txtSearch = Add-TextBox -Parent $null -Width 200
    $lblSearch = Add-Label -Parent $null -Text 'Szukaj na liście:'
    $rowSearch = New-StretchRow -Stretch $txtSearch -Before @($lblSearch)
    $txtSearch.Add_TextChanged({ try { Update-HostQuickFilter } catch { } })

    # Tabela komputerów
    $table = New-HostTable
    $view = [System.Data.DataView]::new($table)
    $view.Sort = 'Name ASC'
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.AutoGenerateColumns = $false
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
    $grid.MultiSelect = $true
    $grid.BackgroundColor = [System.Drawing.SystemColors]::Window
    $grid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $grid.ColumnHeadersDefaultCellStyle.Font = $script:UI.FontBold
    $grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 251)
    $grid.RowTemplate.Height = 22
    try {
        $prop = [System.Windows.Forms.DataGridView].GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance,NonPublic')
        $prop.SetValue($grid, $true, $null)
    }
    catch { }
    $colSel = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colSel.Name = 'Sel'
    $colSel.DataPropertyName = 'Sel'
    $colSel.HeaderText = '✔'
    $colSel.Width = 30
    $colSel.ToolTipText = 'Kliknij nagłówek, aby zaznaczyć/odznaczyć wszystkie widoczne'
    [void]$grid.Columns.Add($colSel)
    foreach ($c in @(
            @{ Name = 'Name'; Header = 'Nazwa'; Width = 130 },
            @{ Name = 'OS'; Header = 'System'; Width = 150 },
            @{ Name = 'LastLogon'; Header = 'Ostatnie logowanie'; Width = 120 },
            @{ Name = 'Enabled'; Header = 'Aktywne'; Width = 60 },
            @{ Name = 'Source'; Header = 'Źródło'; Width = 65 }
        )) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name = $c.Name
        $col.DataPropertyName = $c.Name
        $col.HeaderText = $c.Header
        $col.Width = $c.Width
        $col.ReadOnly = $true
        $col.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Automatic
        [void]$grid.Columns.Add($col)
    }
    $grid.Columns['LastLogon'].DefaultCellStyle.Format = 'yyyy-MM-dd HH:mm'
    $grid.DataSource = $view

    $grid.add_CurrentCellDirtyStateChanged({
            param($s, $e)
            if ($s.IsCurrentCellDirty) { [void]$s.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) }
        })
    $grid.add_CellValueChanged({
            param($s, $e)
            if ($e.RowIndex -ge 0 -and $e.ColumnIndex -eq 0) { try { Update-HostCount } catch { } }
        })
    $grid.add_ColumnHeaderMouseClick({
            param($s, $e)
            if ($e.ColumnIndex -ne 0) { return }
            try {
                $allChecked = $true
                foreach ($drv in @($script:UI.HostView | ForEach-Object { $_ })) { if (-not [bool]$drv.Row['Sel']) { $allChecked = $false; break } }
                if ($allChecked) { Set-HostCheck -Mode 'UncheckAll' } else { Set-HostCheck -Mode 'CheckVisible' }
            }
            catch { }
        })
    $grid.add_KeyDown({
            param($s, $e)
            if ($e.KeyCode -ne [System.Windows.Forms.Keys]::Space) { return }
            try {
                $rows = @($s.SelectedRows)
                if ($rows.Count -eq 0 -or -not $rows[0].DataBoundItem) { return }
                # Spacja przełącza zaznaczenie wszystkich wybranych wierszy (także gdy fokus jest w innej kolumnie)
                $target = -not [bool]$rows[0].DataBoundItem.Row['Sel']
                if ($target) { Set-HostCheck -Mode 'CheckSelected' } else { Set-HostCheck -Mode 'UncheckSelected' }
                $e.Handled = $true
                $e.SuppressKeyPress = $true
            }
            catch { }
        })
    $grid.add_DataError({ param($s, $e) $e.ThrowException = $false })

    # Menu kontekstowe listy
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $miCheck = $menu.Items.Add('Zaznacz wybrane wiersze')
    $miCheck.add_Click({ try { Set-HostCheck -Mode 'CheckSelected' } catch { } })
    $miUncheck = $menu.Items.Add('Odznacz wybrane wiersze')
    $miUncheck.add_Click({ try { Set-HostCheck -Mode 'UncheckSelected' } catch { } })
    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
    $miCopy = $menu.Items.Add('Kopiuj nazwy wybranych')
    $miCopy.add_Click({
            try {
                $names = @(foreach ($gr in $script:UI.HostGrid.SelectedRows) { if ($gr.DataBoundItem) { [string]$gr.DataBoundItem.Row['Name'] } }) | Sort-Object
                if ($names) { [System.Windows.Forms.Clipboard]::SetText(($names -join [Environment]::NewLine)) }
            }
            catch { }
        })
    $miRemove = $menu.Items.Add('Usuń wybrane z listy')
    $miRemove.add_Click({ try { Remove-SelectedHosts } catch { } })
    $grid.ContextMenuStrip = $menu

    # Przyciski zaznaczania
    $rowCheck = New-FlowRow
    $b1 = New-PlainButton -Parent $rowCheck -Text 'Zaznacz widoczne'
    $b1.add_Click({ try { Set-HostCheck -Mode 'CheckVisible' } catch { } })
    $b2 = New-PlainButton -Parent $rowCheck -Text 'Odznacz wszystkie'
    $b2.add_Click({ try { Set-HostCheck -Mode 'UncheckAll' } catch { } })
    $b3 = New-PlainButton -Parent $rowCheck -Text 'Odwróć'
    $b3.add_Click({ try { Set-HostCheck -Mode 'InvertVisible' } catch { } })
    $lblCount = Add-Label -Parent $null -Text ''
    $lblCount.AutoSize = $false
    $lblCount.Height = 22
    $lblCount.ForeColor = [System.Drawing.Color]::FromArgb(30, 60, 110)
    $lblCount.Font = $script:UI.FontBold

    $script:UI.HostTable = $table
    $script:UI.HostView = $view
    $script:UI.HostGrid = $grid
    $script:UI.HostSearch = $txtSearch
    $script:UI.HostSearchBase = $txtBase
    $script:UI.HostNameFilter = $txtName
    $script:UI.HostOnlyEnabled = $chkEnabled
    $script:UI.HostCountLabel = $lblCount

    Add-DockStack -Parent $group -Top @($lblOu, $rowOu, $rowFilter, $rowLoad, $rowSearch) -Fill $grid -Bottom @($rowCheck, $lblCount)
    Update-HostCount
    return $group
}
#endregion

#region Moduły: Diagnostyka
Register-Module -Key 'Connectivity' -Category 'Diagnostyka' -Title 'Łączność' -Description 'Test DNS, ping, portów TCP i sesji PowerShell Remoting dla zaznaczonych komputerów. Wykonywany lokalnie – działa także dla hostów bez WinRM.' -Build {
    param($m)
    $m.ColorBools = $true
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Porty TCP:')
    $m.Ports = Add-TextBox $row 200 '5985, 5986, 445, 3389, 135'
    [void](Add-Label $row 'Limit (ms):')
    $m.TcpTimeout = Add-Numeric $row 200 10000 1500 70
    $m.TestSession = Add-CheckBox $row 'Test sesji PowerShell (Invoke-Command)' $true
    [void](Add-Button $row 'Testuj zaznaczone' $m -Primary {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            $ports = @(Split-ListText ($m.Ports.Text -replace '\s+', ',') | Where-Object { $_ -match '^\d+$' -and [int]$_ -ge 1 -and [int]$_ -le 65535 } | ForEach-Object { [int]$_ } | Select-Object -Unique)
            $params = @{ Ports = $ports; TimeoutMs = [int]$m.TcpTimeout.Value; TestSession = $m.TestSession.Checked }
            Start-HostOperation -Module $m -Name 'Test łączności' -Targets $targets -Local -Parameters $params -ScriptBlock {
                param($Target, $P, $Ctx)
                $row = [ordered]@{}
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
                foreach ($port in $P.Ports) {
                    $open = $false
                    $client = New-Object System.Net.Sockets.TcpClient
                    try {
                        $async = $client.BeginConnect($Target, [int]$port, $null, $null)
                        if ($async.AsyncWaitHandle.WaitOne([int]$P.TimeoutMs) -and $client.Connected) { $open = $true }
                    }
                    catch { }
                    finally { $client.Close() }
                    $row["TCP $port"] = $open
                }
                if ($P.TestSession) {
                    try {
                        $ic = @{ ComputerName = $Target; ErrorAction = 'Stop'; ScriptBlock = { $PSVersionTable.PSVersion.ToString() } }
                        if ($Ctx.Credential) { $ic.Credential = $Ctx.Credential }
                        if ($Ctx.SessionOption) { $ic.SessionOption = $Ctx.SessionOption }
                        $row['PowerShell zdalnie'] = 'Tak (PS ' + (Invoke-Command @ic) + ')'
                    }
                    catch {
                        $row['PowerShell zdalnie'] = 'Nie'
                        $row['Szczegóły'] = ($_.Exception.Message -split "`n")[0].Trim()
                    }
                }
                [pscustomobject]$row
            }
        })
}

Register-Module -Key 'Inventory' -Category 'Diagnostyka' -Title 'Inwentaryzacja' -Description 'Sprzęt, system, numer seryjny, pamięć, adresy IP i zalogowany użytkownik zaznaczonych komputerów.' -Build {
    param($m)
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Pobierz informacje' $m -Primary {
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
                if ($cv) {
                    if ($cv.DisplayVersion) { $release = $cv.DisplayVersion } elseif ($cv.ReleaseId) { $release = $cv.ReleaseId }
                    if ($null -ne $cv.UBR) { $build = '{0}.{1}' -f $os.BuildNumber, $cv.UBR } else { $build = $os.BuildNumber }
                }
                else { $build = $os.BuildNumber }
                [pscustomobject]@{
                    'Producent'              = $cs.Manufacturer
                    'Model'                  = $cs.Model
                    'Numer seryjny'          = $bios.SerialNumber
                    'BIOS'                   = $bios.SMBIOSBIOSVersion
                    'System'                 = $os.Caption
                    'Wydanie'                = $release
                    'Kompilacja'             = $build
                    'Architektura'           = $os.OSArchitecture
                    'Procesor'               = ([string]$cpu.Name).Trim()
                    'Rdzenie'                = $cpu.NumberOfCores
                    'RAM (GB)'               = [Math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
                    'Dysk systemowy (GB)'    = $(if ($sysDisk) { [Math]::Round($sysDisk.Size / 1GB, 1) } else { $null })
                    'Wolne na systemowym (GB)' = $(if ($sysDisk) { [Math]::Round($sysDisk.FreeSpace / 1GB, 1) } else { $null })
                    'Zalogowany użytkownik'  = $cs.UserName
                    'Domena'                 = $cs.Domain
                    'Adresy IP'              = $ips
                    'MAC'                    = $macs
                    'Instalacja systemu'     = $os.InstallDate
                    'Ostatni start'          = $os.LastBootUpTime
                }
            }
        })
}

Register-Module -Key 'Power' -Category 'Diagnostyka' -Title 'Zasilanie i uptime' -Description 'Czas pracy, oczekujący restart (CBS, Windows Update, zmiana nazwy, SCCM) oraz zaplanowany restart/wyłączenie z komunikatem dla użytkownika.' -Build {
    param($m)
    $m.Actions.List = {
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
                'Ostatni start'       = $boot
                'Czas pracy'          = '{0} d {1:00} h {2:00} min' -f $up.Days, $up.Hours, $up.Minutes
                'Dni pracy'           = [Math]::Round($up.TotalDays, 1)
                'Oczekuje restartu'   = ($reasons.Count -gt 0)
                'Powód'               = ($reasons -join ', ')
                'Zalogowany'          = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
            }
        }
    }
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Pokaż uptime' $m -Primary $m.Actions.List)

    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'Opóźnienie (s):')
    $m.Delay = Add-Numeric $row2 0 86400 60 80
    [void](Add-Label $row2 'Komunikat:')
    $m.Message = Add-TextBox $row2 330 'Komputer zostanie uruchomiony ponownie przez administratora. Zapisz swoją pracę.'
    $m.Force = Add-CheckBox $row2 'Wymuś zamknięcie aplikacji' $true

    $powerAction = {
        param($m, $s)
        $mode = [string]$s.Tag['Mode']
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $verb = @{ '/r' = 'Uruchomić ponownie'; '/s' = 'Wyłączyć'; '/a' = 'Anulować zaplanowany restart/wyłączenie na' }[$mode]
        if (-not (Confirm-Action "$verb $($targets.Count) komputer(ów)?" $targets)) { return }
        $params = @{ Mode = $mode; Delay = [int]$m.Delay.Value; Message = $m.Message.Text.Trim(); Force = $m.Force.Checked }
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
    $b = Add-Button $row2 'Restart' $m -Danger $powerAction
    $b.Tag['Mode'] = '/r'
    $b = Add-Button $row2 'Wyłącz' $m -Danger $powerAction
    $b.Tag['Mode'] = '/s'
    $b = Add-Button $row2 'Anuluj zaplanowane' $m $powerAction
    $b.Tag['Mode'] = '/a'
}
#endregion

#region Moduły: Zdalne wykonanie
Register-Module -Key 'Commands' -Category 'Zdalne wykonanie' -Title 'Polecenia' -Description 'Uruchamia polecenia PowerShell lub cmd.exe na zaznaczonych komputerach (sesja WinRM, bez pulpitu użytkownika). Dwuklik na wierszu pokazuje pełny wynik.' -Build {
    param($m)
    $m.Templates = @(
        @{ Name = '(wybierz szablon polecenia)'; Mode = ''; Text = '' },
        @{ Name = 'Wersja systemu i czas pracy'; Mode = 'PS'; Text = "Get-CimInstance Win32_OperatingSystem | Select-Object Caption, Version, LastBootUpTime" },
        @{ Name = 'Ostatnie poprawki (10)'; Mode = 'PS'; Text = "Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 10 HotFixID, Description, InstalledOn" },
        @{ Name = 'Konfiguracja IP'; Mode = 'CMD'; Text = 'ipconfig /all' },
        @{ Name = 'Odśwież DNS (flushdns + registerdns)'; Mode = 'CMD'; Text = 'ipconfig /flushdns && ipconfig /registerdns' },
        @{ Name = 'Wyczyść bilety Kerberos komputera'; Mode = 'CMD'; Text = 'klist -li 0x3e7 purge' },
        @{ Name = 'Sesje użytkowników (quser)'; Mode = 'CMD'; Text = 'quser' },
        @{ Name = 'Wynikowe zasady komputera (gpresult)'; Mode = 'CMD'; Text = 'gpresult /r /scope computer' },
        @{ Name = 'Procesy – top 10 pamięci'; Mode = 'PS'; Text = "Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 10 Name, Id, @{n='MB';e={[math]::Round(`$_.WorkingSet64/1MB)}}" },
        @{ Name = 'Test połączenia z kontrolerem domeny'; Mode = 'PS'; Text = "Test-NetConnection -ComputerName `$env:USERDNSDOMAIN -Port 389 | Select-Object ComputerName, RemoteAddress, TcpTestSucceeded" }
    )
    $row = Add-ToolbarRow $m
    $m.RbPs = Add-RadioButton $row 'PowerShell' $true
    $m.RbCmd = Add-RadioButton $row 'cmd.exe' $false
    [void](Add-Label $row '   Szablon:')
    $m.Template = Add-ComboBox $row @($m.Templates | ForEach-Object { $_.Name }) 300 0
    Register-ControlHandler -Control $m.Template -EventName 'SelectedIndexChanged' -Module $m -Action {
        param($m, $s)
        $t = $m.Templates[$s.SelectedIndex]
        if (-not $t.Mode) { return }
        $m.RbPs.Checked = ($t.Mode -eq 'PS')
        $m.RbCmd.Checked = ($t.Mode -eq 'CMD')
        $m.CommandBox.Text = $t.Text
    }

    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true
    $box.AcceptsReturn = $true
    $box.AcceptsTab = $true
    $box.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $box.WordWrap = $false
    $box.Font = $script:UI.FontMono
    $box.Height = 130
    $m.CommandBox = Add-TopControl $m $box

    $row2 = Add-ToolbarRow $m
    [void](Add-Button $row2 'Uruchom na zaznaczonych' $m -Primary {
            param($m)
            $command = $m.CommandBox.Text.Trim()
            if (-not $command) { Show-Warning 'Wpisz polecenie do wykonania.'; return }
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            $mode = if ($m.RbPs.Checked) { 'PowerShell' } else { 'cmd.exe' }
            $preview = if ($command.Length -gt 300) { $command.Substring(0, 300) + '…' } else { $command }
            if (-not (Confirm-Action "Uruchomić polecenie ($mode) na $($targets.Count) komputer(ach)?`r`n`r`n$preview" $targets)) { return }
            Write-Log "Polecenie ($mode): $command"
            if ($m.RbPs.Checked) {
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
                    [pscustomobject]@{ 'Kod wyjścia' = $proc.ExitCode; 'Wynik' = $text.Trim() }
                }
            }
        })
    [void](Add-Label $row2 'Uwaga: polecenia działają bez profilu i pulpitu użytkownika; zasoby sieciowe mogą być niedostępne (podwójny przeskok).' -Hint)
}

Register-Module -Key 'Install' -Category 'Zdalne wykonanie' -Title 'Instalacja oprogramowania' -Description 'Kopiuje instalator (MSI, MSP, MSU, EXE) na zaznaczone komputery i uruchamia go w trybie cichym. Wynik zawiera kod wyjścia i jego znaczenie.' -Build {
    param($m)
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Instalator:')
    $m.Path = Add-TextBox $row 460
    [void](Add-Button $row 'Przeglądaj…' $m {
            param($m)
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Filter = 'Instalatory (*.msi;*.msp;*.msu;*.exe)|*.msi;*.msp;*.msu;*.exe|Wszystkie pliki (*.*)|*.*'
            if ($dlg.ShowDialog($script:UI.Form) -eq [System.Windows.Forms.DialogResult]::OK) { $m.Path.Text = $dlg.FileName }
        })
    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'Dodatkowe argumenty:')
    $m.Args = Add-TextBox $row2 300
    $m.Cleanup = Add-CheckBox $row2 'Usuń instalator po zakończeniu' $true
    [void](Add-Button $row2 'Zainstaluj na zaznaczonych' $m -Primary {
            param($m)
            $path = $m.Path.Text.Trim().Trim('"')
            if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { Show-Warning 'Wskaż istniejący plik instalatora.'; return }
            $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()
            if (@('.msi', '.msp', '.msu', '.exe') -notcontains $ext) { Show-Warning 'Obsługiwane są pliki .msi, .msp, .msu i .exe.'; return }
            $extra = $m.Args.Text.Trim()
            if ($ext -eq '.exe' -and -not $extra) {
                if (-not (Confirm-Action 'Nie podano argumentów cichej instalacji dla pliku EXE. Instalator może czekać na odpowiedź użytkownika, którego nie ma – operacja zawiśnie do czasu anulowania. Kontynuować?')) { return }
            }
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            if (-not (Confirm-Action "Zainstalować $([System.IO.Path]::GetFileName($path)) na $($targets.Count) komputer(ach)?" $targets)) { return }
            $params = @{ LocalPath = $path; FileName = [System.IO.Path]::GetFileName($path); Args = $extra; Cleanup = $m.Cleanup.Checked }
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
                        [pscustomobject]@{
                            'Plik'            = [System.IO.Path]::GetFileName($File)
                            'Kod wyjścia'     = $code
                            'Wynik'           = $meaning
                            'Czas (s)'        = [Math]::Round($sw.Elapsed.TotalSeconds)
                            'Log instalatora' = $log
                        }
                    }
                    [pscustomobject]@{
                        'Plik'            = $res.Plik
                        'Kod wyjścia'     = $res.'Kod wyjścia'
                        'Wynik'           = $res.Wynik
                        'Czas (s)'        = $res.'Czas (s)'
                        'Kopiowanie'      = $method
                        'Log instalatora' = $res.'Log instalatora'
                    }
                }
                finally {
                    Remove-PSSession -Session $session -ErrorAction SilentlyContinue
                }
            }
        })
    [void](Add-Label (Add-ToolbarRow $m) 'MSI/MSP: automatycznie /qn /norestart i log w %SystemRoot%\Temp\DomainOps;  MSU: /quiet /norestart;  EXE: podaj przełączniki cichej instalacji (np. /S, /quiet, /silent).' -Hint)
}

Register-Module -Key 'GPUpdate' -Category 'Zdalne wykonanie' -Title 'Aktualizacja zasad grupy' -Description 'Wymusza odświeżenie zasad grupy (gpupdate) na zaznaczonych komputerach.' -Build {
    param($m)
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Zakres:')
    $m.Target = Add-ComboBox $row @('Komputer', 'Komputer i użytkownik', 'Użytkownik') 180 0
    $m.Force = Add-CheckBox $row 'Wymuś ponowne zastosowanie (/force)' $true
    [void](Add-Button $row 'Uruchom gpupdate' $m -Primary {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            $scope = @('Computer', '', 'User')[$m.Target.SelectedIndex]
            Start-HostOperation -Module $m -Name 'GPUpdate' -Targets $targets -Parameters @{ Target = $scope; Force = $m.Force.Checked } -ScriptBlock {
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
                    'Kod wyjścia' = $proc.ExitCode
                    'Wynik'       = $(if ($proc.ExitCode -eq 0) { 'OK' } else { "Błąd – kod $($proc.ExitCode)" })
                    'Komunikaty'  = ($lines -join ' | ')
                }
            }
        })
}
#endregion

#region Moduły: System
Register-Module -Key 'Services' -Category 'System' -Title 'Usługi' -Description 'Lista usług na zaznaczonych komputerach. Akcje dotyczą usług zaznaczonych w tabeli (na właściwych hostach).' -Build {
    param($m)
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
                [pscustomobject]@{
                    'Nazwa'             = $_.Name
                    'Nazwa wyświetlana' = $_.DisplayName
                    'Stan'              = $_.State
                    'Uruchamianie'      = $_.StartMode
                    'Konto'             = $_.StartName
                    'PID'               = $_.ProcessId
                    'Ścieżka'           = $_.PathName
                }
            }
        }
    }
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Filtr nazwy:')
    $m.Filter = Add-TextBox $row 160
    [void](Add-Label $row 'Stan:')
    $m.StateFilter = Add-ComboBox $row @('Wszystkie', 'Uruchomione', 'Zatrzymane', 'Automatyczne, ale zatrzymane') 220 0
    [void](Add-Button $row 'Pokaż usługi' $m -Primary $m.Actions.List)

    $serviceAction = {
        param($m, $s)
        $op = [string]$s.Tag['Op']
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Nazwa')
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli usługi, których dotyczy operacja.'; return }
        $startup = @('Automatic', 'Manual', 'Disabled')[$m.StartupType.SelectedIndex]
        $question = @{
            Start       = 'Uruchomić wybrane usługi?'
            Stop        = 'Zatrzymać wybrane usługi? Zatrzymane zostaną też usługi od nich zależne.'
            Restart     = 'Uruchomić ponownie wybrane usługi?'
            StartupType = "Ustawić typ uruchamiania «$($m.StartupType.Text)» dla wybranych usług?"
        }[$op]
        $items = foreach ($h in $byHost.Keys) { foreach ($i in $byHost[$h]) { '{0}: {1}' -f $h, $i['Nazwa'] } }
        if (-not (Confirm-Action $question @($items))) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Op = $op; StartupType = $startup; Names = @($byHost[$h] | ForEach-Object { [string]$_['Nazwa'] }) } }
        Start-HostOperation -Module $m -Name "Usługi – $op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
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
    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'Zaznaczone usługi:')
    foreach ($a in @(@('Start', 'Start'), @('Stop', 'Stop'), @('Restart', 'Restart'))) {
        $b = Add-Button $row2 $a[0] $m $serviceAction
        $b.Tag['Op'] = $a[1]
    }
    [void](Add-Label $row2 '   Typ uruchamiania:')
    $m.StartupType = Add-ComboBox $row2 @('Automatyczny', 'Ręczny', 'Wyłączony') 130 0
    $b = Add-Button $row2 'Ustaw' $m $serviceAction
    $b.Tag['Op'] = 'StartupType'
}

Register-Module -Key 'Processes' -Category 'System' -Title 'Procesy' -Description 'Procesy uruchomione na zaznaczonych komputerach (pamięć, właściciel, wiersz poleceń) i kończenie wybranych procesów.' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Procesy' -Targets $targets -Parameters @{ Filter = $m.Filter.Text.Trim(); WithOwner = $m.WithOwner.Checked } -ScriptBlock {
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
                    'Uruchomiony'    = $_.CreationDate
                    'Wiersz poleceń' = $_.CommandLine
                }
            }
        }
    }
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Filtr nazwy:')
    $m.Filter = Add-TextBox $row 160
    $m.WithOwner = Add-CheckBox $row 'Pokaż właściciela (wolniej)' $false
    [void](Add-Button $row 'Pokaż procesy' $m -Primary $m.Actions.List)
    [void](Add-Button $row 'Zakończ zaznaczone procesy' $m -Danger {
            param($m)
            $byHost = Get-SelectedRowsByHost -Module $m -Columns @('PID', 'Proces')
            if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli procesy do zakończenia.'; return }
            $items = foreach ($h in $byHost.Keys) { foreach ($i in $byHost[$h]) { '{0}: {1} (PID {2})' -f $h, $i['Proces'], $i['PID'] } }
            if (-not (Confirm-Action 'Zakończyć wybrane procesy? Niezapisane dane w tych programach zostaną utracone.' @($items))) { return }
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
        })
}

Register-Module -Key 'Disks' -Category 'System' -Title 'Dyski' -Description 'Zajętość dysków lokalnych oraz czyszczenie plików tymczasowych i Kosza (z raportem odzyskanego miejsca).' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Dyski' -Targets $targets -ScriptBlock {
            param($P)
            Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType = 3' | ForEach-Object {
                $size = [double]$_.Size
                $free = [double]$_.FreeSpace
                [pscustomobject]@{
                    'Dysk'          = $_.DeviceID
                    'Etykieta'      = $_.VolumeName
                    'System plików' = $_.FileSystem
                    'Rozmiar (GB)'  = [Math]::Round($size / 1GB, 1)
                    'Wolne (GB)'    = [Math]::Round($free / 1GB, 1)
                    'Zajęte (%)'    = $(if ($size -gt 0) { [Math]::Round(($size - $free) / $size * 100, 1) } else { $null })
                }
            }
        }
    }
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Pokaż dyski' $m -Primary $m.Actions.List)

    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'Czyszczenie:')
    $m.WinTemp = Add-CheckBox $row2 'Windows\Temp' $true
    $m.UserTemp = Add-CheckBox $row2 'TEMP profili użytkowników' $true
    $m.Recycle = Add-CheckBox $row2 'Kosz (wszystkie dyski)' $true
    [void](Add-Label $row2 'Pliki starsze niż (dni):')
    $m.Days = Add-Numeric $row2 0 365 2 60
    [void](Add-Button $row2 'Wyczyść na zaznaczonych' $m -Danger {
            param($m)
            if (-not ($m.WinTemp.Checked -or $m.UserTemp.Checked -or $m.Recycle.Checked)) { Show-Warning 'Wybierz, co ma zostać wyczyszczone.'; return }
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            if (-not (Confirm-Action "Usunąć pliki tymczasowe na $($targets.Count) komputer(ach)?" $targets)) { return }
            $params = @{ WindowsTemp = $m.WinTemp.Checked; UserTemp = $m.UserTemp.Checked; RecycleBin = $m.Recycle.Checked; OlderThanDays = [int]$m.Days.Value }
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
                    if ($root -like '*\Temp\DomainOps*') { continue }
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
                $after = [double](Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $env:SystemDrive)).FreeSpace
                [pscustomobject]@{
                    'Usunięte elementy'          = $removed
                    'Nie udało się usunąć'       = $failed
                    'Odzyskano na systemowym (GB)' = [Math]::Round(($after - $before) / 1GB, 2)
                    'Wolne na systemowym (GB)'   = [Math]::Round($after / 1GB, 1)
                }
            }
        })
}

Register-Module -Key 'Events' -Category 'System' -Title 'Dziennik zdarzeń' -Description 'Błędy i ostrzeżenia z wybranych dzienników z ostatnich godzin. Filtr wyników działa także po treści komunikatu.' -Build {
    param($m)
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Ostatnie godziny:')
    $m.Hours = Add-Numeric $row 1 720 24 60
    $m.LogSystem = Add-CheckBox $row 'System' $true
    $m.LogApp = Add-CheckBox $row 'Application' $true
    $m.LogSec = Add-CheckBox $row 'Security' $false
    [void](Add-Label $row '  Poziom:')
    $m.LvlCrit = Add-CheckBox $row 'Krytyczny' $true
    $m.LvlErr = Add-CheckBox $row 'Błąd' $true
    $m.LvlWarn = Add-CheckBox $row 'Ostrzeżenie' $true
    $m.LvlInfo = Add-CheckBox $row 'Informacja' $false
    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'ID zdarzeń (opcjonalnie, np. 41, 6008):')
    $m.Ids = Add-TextBox $row2 160
    [void](Add-Label $row2 'Maks. na host:')
    $m.Max = Add-Numeric $row2 10 5000 200 70
    [void](Add-Button $row2 'Pobierz zdarzenia' $m -Primary {
            param($m)
            $logs = @()
            if ($m.LogSystem.Checked) { $logs += 'System' }
            if ($m.LogApp.Checked) { $logs += 'Application' }
            if ($m.LogSec.Checked) { $logs += 'Security' }
            if ($logs.Count -eq 0) { Show-Warning 'Wybierz co najmniej jeden dziennik.'; return }
            $levels = @()
            if ($m.LvlCrit.Checked) { $levels += 1 }
            if ($m.LvlErr.Checked) { $levels += 2 }
            if ($m.LvlWarn.Checked) { $levels += 3 }
            if ($m.LvlInfo.Checked) { $levels += 4; $levels += 0 }
            if ($levels.Count -eq 0) { Show-Warning 'Wybierz co najmniej jeden poziom zdarzeń.'; return }
            $ids = @(Split-ListText ($m.Ids.Text -replace '\s+', ',') | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            $params = @{ Logs = $logs; Levels = $levels; Ids = $ids; Hours = [int]$m.Hours.Value; Max = [int]$m.Max.Value }
            Start-HostOperation -Module $m -Name 'Dziennik zdarzeń' -Targets $targets -Parameters $params -ScriptBlock {
                param($P)
                $filter = @{ LogName = [string[]]@($P.Logs); StartTime = (Get-Date).AddHours( - [int]$P.Hours) }
                $ids = @($P.Ids | Where-Object { $null -ne $_ })
                if ($ids.Count -gt 0) { $filter.Id = [int[]]$ids }
                # Security nie używa poziomów (wszystko to "Informacje"/audyt) - bez filtra poziomu, gdy wybrano tylko Security
                $levels = [int[]]@($P.Levels)
                if (@($P.Logs) -notcontains 'Security' -or @($P.Logs).Count -gt 1) { $filter.Level = $levels }
                try {
                    $events = @(Get-WinEvent -FilterHashtable $filter -MaxEvents ([int]$P.Max) -ErrorAction Stop)
                }
                catch {
                    if ($_.FullyQualifiedErrorId -like 'NoMatchingEventsFound*') { $events = @() } else { throw }
                }
                $events | Sort-Object TimeCreated -Descending | ForEach-Object {
                    $msg = $_.Message
                    if (-not $msg) { $msg = '(brak opisu – brak biblioteki komunikatów dostawcy)' }
                    [pscustomobject]@{
                        'Czas'      = $_.TimeCreated
                        'Poziom'    = $_.LevelDisplayName
                        'ID'        = $_.Id
                        'Źródło'    = $_.ProviderName
                        'Dziennik'  = $_.LogName
                        'Komunikat' = ($msg -replace '\s+', ' ').Trim()
                    }
                }
            }
        })
}

Register-Module -Key 'Tasks' -Category 'System' -Title 'Harmonogram zadań' -Description 'Zadania Harmonogramu na zaznaczonych komputerach: podgląd, uruchamianie, włączanie, wyłączanie, usuwanie i tworzenie prostych zadań (konto SYSTEM).' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Harmonogram' -Targets $targets -Parameters @{ Filter = $m.Filter.Text.Trim(); HideMicrosoft = $m.HideMs.Checked } -ScriptBlock {
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
                }
            }
        }
    }
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Filtr:')
    $m.Filter = Add-TextBox $row 160
    $m.HideMs = Add-CheckBox $row 'Ukryj zadania systemowe (\Microsoft\)' $true
    [void](Add-Button $row 'Pokaż zadania' $m -Primary $m.Actions.List)

    $taskAction = {
        param($m, $s)
        $op = [string]$s.Tag['Op']
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Ścieżka', 'Nazwa')
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli zadania, których dotyczy operacja.'; return }
        $label = @{ Run = 'Uruchomić'; Stop = 'Zatrzymać'; Enable = 'Włączyć'; Disable = 'Wyłączyć'; Delete = 'USUNĄĆ' }[$op]
        $items = foreach ($h in $byHost.Keys) { foreach ($i in $byHost[$h]) { '{0}: {1}{2}' -f $h, $i['Ścieżka'], $i['Nazwa'] } }
        if (-not (Confirm-Action "$label wybrane zadania?" @($items))) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) {
            $per[$h] = @{ Op = $op; Tasks = @($byHost[$h] | ForEach-Object { @{ Path = [string]$_['Ścieżka']; Name = [string]$_['Nazwa'] } }) }
        }
        Start-HostOperation -Module $m -Name "Harmonogram – $op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
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
    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'Zaznaczone zadania:')
    foreach ($a in @(@('Uruchom', 'Run'), @('Zatrzymaj', 'Stop'), @('Włącz', 'Enable'), @('Wyłącz', 'Disable'))) {
        $b = Add-Button $row2 $a[0] $m $taskAction
        $b.Tag['Op'] = $a[1]
    }
    $b = Add-Button $row2 'Usuń' $m -Danger $taskAction
    $b.Tag['Op'] = 'Delete'

    $row3 = Add-ToolbarRow $m
    [void](Add-Label $row3 'Nowe zadanie:' -Bold)
    [void](Add-Label $row3 'Nazwa:')
    $m.NewName = Add-TextBox $row3 150
    [void](Add-Label $row3 'Program:')
    $m.NewExe = Add-TextBox $row3 200
    [void](Add-Label $row3 'Argumenty:')
    $m.NewArgs = Add-TextBox $row3 180
    [void](Add-Label $row3 'Wyzwalacz:')
    $m.NewTrigger = Add-ComboBox $row3 @('Przy logowaniu', 'Przy uruchomieniu', 'Codziennie o', 'Jednorazowo o', 'Tylko na żądanie') 140 0
    $m.NewTime = Add-TextBox $row3 50 '07:00'
    [void](Add-Button $row3 'Utwórz na zaznaczonych' $m {
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
            if (-not (Confirm-Action "Utworzyć zadanie «$name» (konto SYSTEM, najwyższe uprawnienia) na $($targets.Count) komputer(ach)?`r`nProgram: $exe $($m.NewArgs.Text.Trim())" $targets)) { return }
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
        })
}

Register-Module -Key 'Drivers' -Category 'System' -Title 'Sterowniki i urządzenia' -Description 'Zainstalowane sterowniki (Win32_PnPSignedDriver) oraz urządzenia zgłaszające problem w Menedżerze urządzeń.' -Build {
    param($m)
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Filtr (urządzenie/producent/klasa):')
    $m.Filter = Add-TextBox $row 200
    [void](Add-Button $row 'Pokaż sterowniki' $m -Primary {
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
        })
    [void](Add-Button $row 'Urządzenia z problemami' $m {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            Start-HostOperation -Module $m -Name 'Urządzenia z problemami' -Targets $targets -ScriptBlock {
                param($P)
                $devices = @(Get-CimInstance -ClassName Win32_PnPEntity | Where-Object { $_.ConfigManagerErrorCode -ne 0 })
                if ($devices.Count -eq 0) { return [pscustomobject]@{ 'Urządzenie' = '(brak urządzeń z problemami)' } }
                foreach ($d in $devices) {
                    [pscustomobject]@{
                        'Urządzenie'  = $d.Name
                        'Kod błędu'   = $d.ConfigManagerErrorCode
                        'Stan'        = $d.Status
                        'Klasa'       = $d.PNPClass
                        'Producent'   = $d.Manufacturer
                        'ID urządzenia' = $d.DeviceID
                    }
                }
            }
        })
}
#endregion

#region Moduły: Oprogramowanie
# Skrypt instalacji aktualizacji uruchamiany na hoście jako zadanie SYSTEM - API Windows Update
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

Register-Module -Key 'Programs' -Category 'Oprogramowanie' -Title 'Zainstalowane programy' -Description 'Programy z rejestru (64/32-bit i profile zalogowanych użytkowników) oraz ciche odinstalowanie zaznaczonych (MSI albo QuietUninstallString).' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Programy' -Targets $targets -Parameters @{ Filter = $m.Filter.Text.Trim(); ShowSystem = $m.ShowSystem.Checked } -ScriptBlock {
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
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Filtr (nazwa/wydawca):')
    $m.Filter = Add-TextBox $row 180
    $m.ShowSystem = Add-CheckBox $row 'Pokaż aktualizacje i składniki systemowe' $false
    [void](Add-Button $row 'Pokaż programy' $m -Primary $m.Actions.List)
    [void](Add-Button $row 'Odinstaluj zaznaczone' $m -Danger {
            param($m)
            $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Identyfikator', 'Nazwa', 'Zakres')
            if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli programy do odinstalowania.'; return }
            $items = foreach ($h in $byHost.Keys) { foreach ($i in $byHost[$h]) { '{0}: {1}' -f $h, $i['Nazwa'] } }
            if (-not (Confirm-Action 'Odinstalować wybrane programy? Operacja jest wykonywana bez interakcji z użytkownikiem i bez restartu.' @($items))) { return }
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
        })
}

Register-Module -Key 'WindowsUpdate' -Category 'Oprogramowanie' -Title 'Windows Update' -Description 'Wyszukiwanie dostępnych aktualizacji, historia oraz instalacja (zadanie SYSTEM na hoście, log w %SystemRoot%\Temp\DomainOps\WU.log). Nie wymaga modułu PSWindowsUpdate.' -Build {
    param($m)
    $row = Add-ToolbarRow $m
    $m.Drivers = Add-CheckBox $row 'Uwzględnij sterowniki' $false
    [void](Add-Button $row 'Wyszukaj dostępne' $m -Primary {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            Start-HostOperation -Module $m -Name 'Wyszukiwanie aktualizacji' -Targets $targets -Parameters @{ IncludeDrivers = $m.Drivers.Checked } -ScriptBlock {
                param($P)
                $session = New-Object -ComObject Microsoft.Update.Session
                $criteria = "IsInstalled=0 and IsHidden=0"
                if (-not $P.IncludeDrivers) { $criteria += " and Type='Software'" }
                $result = $session.CreateUpdateSearcher().Search($criteria)
                if ($result.Updates.Count -eq 0) { return [pscustomobject]@{ 'Aktualizacja' = '(brak dostępnych aktualizacji)' } }
                foreach ($u in $result.Updates) {
                    [pscustomobject]@{
                        'Aktualizacja'    = $u.Title
                        'KB'              = (@($u.KBArticleIDs | ForEach-Object { "KB$_" }) -join ', ')
                        'Kategoria'       = (@($u.Categories | ForEach-Object { $_.Name }) -join ', ')
                        'Ważność'         = $u.MsrcSeverity
                        'Rozmiar (MB)'    = [Math]::Round([double]$u.MaxDownloadSize / 1MB, 1)
                        'Pobrana'         = [bool]$u.IsDownloaded
                        'Wymaga restartu' = ($u.InstallationBehavior.RebootBehavior -ne 0)
                    }
                }
            }
        })
    [void](Add-Button $row 'Historia (ostatnie 50)' $m {
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
                    }
                }
            }
        })
    $row2 = Add-ToolbarRow $m
    $m.AutoReboot = Add-CheckBox $row2 'Automatyczny restart po instalacji (za 5 min), jeśli wymagany' $false
    [void](Add-Button $row2 'Zainstaluj aktualizacje' $m -Danger {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            $reboot = if ($m.AutoReboot.Checked) { 'z automatycznym restartem' } else { 'bez restartu' }
            if (-not (Confirm-Action "Zainstalować wszystkie dostępne aktualizacje ($reboot) na $($targets.Count) komputer(ach)?" $targets)) { return }
            $params = @{ Script = $script:WuJobScript; AutoReboot = $m.AutoReboot.Checked; IncludeDrivers = $m.Drivers.Checked }
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
        })
    [void](Add-Button $row2 'Stan instalacji' $m {
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
                [pscustomobject]@{
                    'Zadanie'         = $(if ($task) { [string]$task.State } else { '(brak – instalacji nie zlecano)' })
                    'Ostatni wpis'    = $(if ($lines.Count) { $lines[-1] } else { '' })
                    'Wymaga restartu' = $rebootRequired
                    'Log'             = ($lines -join "`r`n")
                }
            }
        })
}
#endregion

#region Moduły: Bezpieczeństwo
Register-Module -Key 'Defender' -Category 'Bezpieczeństwo' -Title 'Microsoft Defender' -Description 'Stan ochrony i sygnatur, wykryte zagrożenia, aktualizacja sygnatur oraz skanowanie (skan trwa w tle – można pracować dalej).' -Build {
    param($m)
    $m.ColorBools = $true
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Stan ochrony' $m -Primary {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            Start-HostOperation -Module $m -Name 'Defender – stan' -Targets $targets -ScriptBlock {
                param($P)
                if (-not (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue)) { throw 'Brak modułu Defender (Get-MpComputerStatus) na hoście.' }
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
                }
            }
        })
    [void](Add-Button $row 'Wykryte zagrożenia' $m {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            Start-HostOperation -Module $m -Name 'Defender – zagrożenia' -Targets $targets -ScriptBlock {
                param($P)
                if (-not (Get-Command Get-MpThreatDetection -ErrorAction SilentlyContinue)) { throw 'Brak modułu Defender na hoście.' }
                $names = @{}
                foreach ($t in @(Get-MpThreat -ErrorAction SilentlyContinue)) { $names[[string]$t.ThreatID] = $t.ThreatName }
                $detections = @(Get-MpThreatDetection -ErrorAction SilentlyContinue)
                if ($detections.Count -eq 0) { return [pscustomobject]@{ 'Zagrożenie' = '(brak wykrytych zagrożeń)' } }
                foreach ($d in $detections) {
                    [pscustomobject]@{
                        'Wykryto'       = $d.InitialDetectionTime
                        'Zagrożenie'    = $names[[string]$d.ThreatID]
                        'Zasoby'        = (@($d.Resources) -join '; ')
                        'Akcja udana'   = $d.ActionSuccess
                        'Proces'        = $d.ProcessName
                        'Użytkownik'    = $d.DomainUser
                    }
                }
            }
        })
    $defAction = {
        param($m, $s)
        $op = [string]$s.Tag['Op']
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $label = @{ Update = 'Zaktualizować sygnatury'; QuickScan = 'Uruchomić szybkie skanowanie'; FullScan = 'Uruchomić PEŁNE skanowanie (może trwać godzinami i obciąża dysk)' }[$op]
        if (-not (Confirm-Action "$label na $($targets.Count) komputer(ach)?" $targets)) { return }
        Start-HostOperation -Module $m -Name "Defender – $op" -Targets $targets -Output Log -Parameters @{ Op = $op } -ScriptBlock {
            param($P)
            if (-not (Get-Command Start-MpScan -ErrorAction SilentlyContinue)) { throw 'Brak modułu Defender na hoście.' }
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            switch ($P.Op) {
                'Update' { Update-MpSignature -ErrorAction Stop; $v = (Get-MpComputerStatus).AntivirusSignatureVersion; "Sygnatury zaktualizowane (wersja $v)." }
                'QuickScan' { Start-MpScan -ScanType QuickScan -ErrorAction Stop; "Szybkie skanowanie zakończone ($([Math]::Round($sw.Elapsed.TotalMinutes, 1)) min)." }
                'FullScan' { Start-MpScan -ScanType FullScan -ErrorAction Stop; "Pełne skanowanie zakończone ($([Math]::Round($sw.Elapsed.TotalMinutes, 1)) min)." }
            }
        }
    }
    foreach ($a in @(@('Aktualizuj sygnatury', 'Update'), @('Szybki skan', 'QuickScan'), @('Pełny skan', 'FullScan'))) {
        $b = Add-Button $row $a[0] $m $defAction
        $b.Tag['Op'] = $a[1]
    }
}

Register-Module -Key 'BitLocker' -Category 'Bezpieczeństwo' -Title 'BitLocker' -Description 'Stan szyfrowania woluminów, kopia zapasowa kluczy odzyskiwania do AD oraz odczyt kluczy zapisanych w AD (wymaga uprawnień do msFVE-RecoveryInformation).' -Build {
    param($m)
    $m.SecretColumns = @('Hasło odzyskiwania')
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Stan woluminów' $m -Primary {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            Start-HostOperation -Module $m -Name 'BitLocker – stan' -Targets $targets -ScriptBlock {
                param($P)
                if (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue) {
                    foreach ($v in @(Get-BitLockerVolume -ErrorAction Stop)) {
                        [pscustomobject]@{
                            'Wolumin'          = $v.MountPoint
                            'Typ'              = [string]$v.VolumeType
                            'Ochrona'          = [string]$v.ProtectionStatus
                            'Stan'             = [string]$v.VolumeStatus
                            'Zaszyfrowano (%)' = $v.EncryptionPercentage
                            'Metoda'           = [string]$v.EncryptionMethod
                            'Blokada'          = [string]$v.LockStatus
                            'Zabezpieczenia'   = (@($v.KeyProtector | ForEach-Object { [string]$_.KeyProtectorType }) -join ', ')
                        }
                    }
                    return
                }
                # Starsze systemy / brak modułu: klasa WMI (niezależna od języka systemu)
                $conversion = @{ 0 = 'Odszyfrowany'; 1 = 'Zaszyfrowany'; 2 = 'Szyfrowanie w toku'; 3 = 'Odszyfrowywanie w toku'; 4 = 'Szyfrowanie wstrzymane'; 5 = 'Odszyfrowywanie wstrzymane' }
                $protection = @{ 0 = 'Off'; 1 = 'On'; 2 = 'Unknown' }
                $volumes = @(Get-CimInstance -Namespace 'root\cimv2\Security\MicrosoftVolumeEncryption' -ClassName Win32_EncryptableVolume -ErrorAction Stop)
                foreach ($v in $volumes) {
                    $status = Invoke-CimMethod -InputObject $v -MethodName GetConversionStatus -ErrorAction SilentlyContinue
                    [pscustomobject]@{
                        'Wolumin'          = $v.DriveLetter
                        'Ochrona'          = $protection[[int]$v.ProtectionStatus]
                        'Stan'             = $(if ($status) { $conversion[[int]$status.ConversionStatus] } else { '' })
                        'Zaszyfrowano (%)' = $(if ($status) { $status.EncryptionPercentage } else { $null })
                    }
                }
            }
        })
    [void](Add-Button $row 'Kopia kluczy do AD' $m {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            if (-not (Confirm-Action "Zapisać w AD klucze odzyskiwania BitLocker (wszystkie woluminy) z $($targets.Count) komputer(ów)?" $targets)) { return }
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
                if (-not (Test-Path -LiteralPath $bde)) { throw 'Brak Get-BitLockerVolume i manage-bde.exe na hoście.' }
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
        })
    [void](Add-Button $row 'Klucze odzyskiwania z AD' $m {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            Start-HostOperation -Module $m -Name 'BitLocker – klucze w AD' -Targets $targets -Local -ScriptBlock {
                param($Target, $P, $Ctx)
                Import-Module ActiveDirectory -ErrorAction Stop
                $ad = @{}
                if ($Ctx.Server) { $ad.Server = $Ctx.Server }
                if ($Ctx.Credential) { $ad.Credential = $Ctx.Credential }
                $computer = Get-ADComputer -Identity ($Target -split '\.')[0] @ad -ErrorAction Stop
                $keys = @(Get-ADObject -SearchBase $computer.DistinguishedName -LDAPFilter '(objectClass=msFVE-RecoveryInformation)' -Properties 'msFVE-RecoveryPassword', 'whenCreated' @ad -ErrorAction Stop)
                if ($keys.Count -eq 0) { return [pscustomobject]@{ 'Identyfikator klucza' = '(brak kluczy w AD albo brak uprawnień do ich odczytu)' } }
                $keys | Sort-Object whenCreated -Descending | ForEach-Object {
                    $id = if ($_.Name -match '\{([0-9A-Fa-f-]{36})\}') { $Matches[1] } else { $_.Name }
                    [pscustomobject]@{
                        'Utworzono'            = $_.whenCreated
                        'Identyfikator klucza' = $id
                        'Hasło odzyskiwania'   = $_.'msFVE-RecoveryPassword'
                    }
                }
            }
        })
    [void](Add-Button $row 'Kopiuj hasło odzyskiwania' $m {
            param($m)
            $row = @(Get-SelectedResultRows -Module $m) | Select-Object -First 1
            $secret = [string](Get-ObjectValue $row 'Hasło odzyskiwania')
            if (-not $secret) { Show-Warning 'Zaznacz wiersz z hasłem odzyskiwania (przycisk «Klucze odzyskiwania z AD»).'; return }
            Set-ClipboardSecret -Text $secret -Seconds 60
            Write-Log ("Skopiowano hasło odzyskiwania BitLocker ({0}, klucz {1}); schowek zostanie wyczyszczony po 60 s." -f (Get-ObjectValue $row 'Komputer'), (Get-ObjectValue $row 'Identyfikator klucza')) 'OK'
        })
}

Register-Module -Key 'Firewall' -Category 'Bezpieczeństwo' -Title 'Zapora Windows' -Description 'Stan profili zapory, reguły (z portami i programami), włączanie/wyłączanie/usuwanie zaznaczonych reguł oraz tworzenie nowych.' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $params = @{ Direction = @('Inbound', 'Outbound')[$m.Direction.SelectedIndex]; EnabledOnly = $m.EnabledOnly.Checked; Filter = $m.Filter.Text.Trim() }
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
                }
            }
        }
    }
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Profile zapory' $m {
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
        })
    [void](Add-Label $row '  Reguły:')
    $m.Direction = Add-ComboBox $row @('Przychodzące', 'Wychodzące') 120 0
    $m.EnabledOnly = Add-CheckBox $row 'Tylko włączone' $true
    [void](Add-Label $row 'Filtr:')
    $m.Filter = Add-TextBox $row 150
    [void](Add-Button $row 'Pokaż reguły' $m -Primary $m.Actions.List)

    $ruleAction = {
        param($m, $s)
        $op = [string]$s.Tag['Op']
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('ID reguły', 'Nazwa wyświetlana')
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli reguły zapory (widok «Pokaż reguły»).'; return }
        $label = @{ Enable = 'Włączyć'; Disable = 'Wyłączyć'; Delete = 'USUNĄĆ' }[$op]
        $items = foreach ($h in $byHost.Keys) { foreach ($i in $byHost[$h]) { '{0}: {1}' -f $h, $i['Nazwa wyświetlana'] } }
        if (-not (Confirm-Action "$label wybrane reguły zapory?" @($items))) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Op = $op; Names = @($byHost[$h] | ForEach-Object { [string]$_['ID reguły'] }) } }
        Start-HostOperation -Module $m -Name "Zapora – $op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
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
    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'Zaznaczone reguły:')
    $b = Add-Button $row2 'Włącz' $m $ruleAction
    $b.Tag['Op'] = 'Enable'
    $b = Add-Button $row2 'Wyłącz' $m $ruleAction
    $b.Tag['Op'] = 'Disable'
    $b = Add-Button $row2 'Usuń' $m -Danger $ruleAction
    $b.Tag['Op'] = 'Delete'

    $row3 = Add-ToolbarRow $m
    [void](Add-Label $row3 'Nowa reguła:' -Bold)
    $m.NewName = Add-TextBox $row3 170 'Domain Ops – nowa reguła'
    $m.NewDirection = Add-ComboBox $row3 @('Przychodząca', 'Wychodząca') 110 0
    $m.NewAction = Add-ComboBox $row3 @('Zezwalaj', 'Blokuj') 90 0
    $m.NewProtocol = Add-ComboBox $row3 @('TCP', 'UDP') 60 0
    [void](Add-Label $row3 'Porty:')
    $m.NewPorts = Add-TextBox $row3 110 '5985'
    [void](Add-Label $row3 'Profil:')
    $m.NewProfile = Add-ComboBox $row3 @('Wszystkie', 'Domena', 'Prywatny', 'Publiczny') 100 0
    [void](Add-Label $row3 'Adres zdalny:')
    $m.NewRemote = Add-TextBox $row3 110 'Any'
    [void](Add-Button $row3 'Utwórz na zaznaczonych' $m {
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
                Protocol      = $m.NewProtocol.Text
                Ports         = $ports
                Profile       = @('Any', 'Domain', 'Private', 'Public')[$m.NewProfile.SelectedIndex]
                RemoteAddress = $remote
            }
            if (-not (Confirm-Action ("Utworzyć regułę «{0}» ({1} {2} {3}/{4}) na {5} komputer(ach)?" -f $name, $m.NewDirection.Text, $m.NewAction.Text, $params.Protocol, ($ports -join ','), $targets.Count) $targets)) { return }
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
        })
}

Register-Module -Key 'Certificates' -Category 'Bezpieczeństwo' -Title 'Certyfikaty komputera' -Description 'Certyfikaty z magazynów LocalMachine, wyszukiwanie wygasających oraz eksport zaznaczonych do plików .cer.' -Build {
    param($m)
    $row = Add-ToolbarRow $m
    [void](Add-Label $row 'Magazyn:')
    $m.Store = Add-ComboBox $row @('My', 'Root', 'CA', 'TrustedPublisher', 'TrustedPeople', 'WebHosting', 'Remote Desktop') 140 0
    [void](Add-Label $row 'Filtr:')
    $m.Filter = Add-TextBox $row 150
    $m.OnlyExpiring = Add-CheckBox $row 'Tylko wygasające w ciągu (dni):' $false
    $m.Days = Add-Numeric $row 1 3650 30 60
    [void](Add-Button $row 'Pokaż certyfikaty' $m -Primary {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            $params = @{ Store = $m.Store.Text; Filter = $m.Filter.Text.Trim(); ExpiringDays = $(if ($m.OnlyExpiring.Checked) { [int]$m.Days.Value } else { 0 }) }
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
                        'Wystawca'           = $c.Issuer
                        'Ważny od'           = $c.NotBefore
                        'Ważny do'           = $c.NotAfter
                        'Dni do wygaśnięcia' = $days
                        'Klucz prywatny'     = $c.HasPrivateKey
                        'Przeznaczenie'      = (@($c.EnhancedKeyUsageList | ForEach-Object { $_.FriendlyName } | Where-Object { $_ }) -join ', ')
                        'Nazwa przyjazna'    = $c.FriendlyName
                        'Odcisk palca'       = $c.Thumbprint
                        'Magazyn'            = $P.Store
                    }
                }
            }
        })
    [void](Add-Button $row 'Eksportuj zaznaczone (.cer)…' $m {
            param($m)
            $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Odcisk palca', 'Magazyn')
            if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli certyfikaty do eksportu.'; return }
            $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
            $dlg.Description = 'Folder docelowy dla plików .cer'
            if ($dlg.ShowDialog($script:UI.Form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $m.ExportFolder = $dlg.SelectedPath
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
                    $file = Join-Path $m.ExportFolder ('{0}_{1}.cer' -f $r.Target, $thumb)
                    [System.IO.File]::WriteAllBytes($file, [Convert]::FromBase64String($b64))
                    Write-Log ("[{0}] zapisano {1}" -f $r.Target, $file) 'OK'
                }
            }
        })
}

Register-Module -Key 'LocalAdmins' -Category 'Bezpieczeństwo' -Title 'Lokalni administratorzy' -Description 'Członkowie lokalnej grupy Administratorzy (wyznaczanej po SID S-1-5-32-544, więc działa w każdym języku systemu): podgląd, dodawanie i usuwanie.' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Lokalni administratorzy' -Targets $targets -ScriptBlock {
            param($P)
            $groupName = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')).Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
            $rows = @()
            try {
                $rows = @(Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop | ForEach-Object {
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
            $rows | ForEach-Object { $_ | Add-Member -NotePropertyName 'Grupa' -NotePropertyValue $groupName -PassThru }
        }
    }
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Pokaż członków' $m -Primary $m.Actions.List)
    [void](Add-Button $row 'Usuń zaznaczonych' $m -Danger {
            param($m)
            $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Członek')
            if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli członków grupy do usunięcia.'; return }
            $items = foreach ($h in $byHost.Keys) { foreach ($i in $byHost[$h]) { '{0}: {1}' -f $h, $i['Członek'] } }
            if (-not (Confirm-Action 'Usunąć wybrane konta z lokalnej grupy Administratorzy?' @($items))) { return }
            $per = @{}
            foreach ($h in $byHost.Keys) {
                $sids = @{}
                foreach ($drv in @(Get-SelectedResultRows -Module $m)) {
                    if ((Get-ObjectValue $drv 'Komputer') -eq $h) { $sids[[string](Get-ObjectValue $drv 'Członek')] = [string](Get-ObjectValue $drv 'SID') }
                }
                $per[$h] = @{ Members = @($byHost[$h] | ForEach-Object { @{ Name = [string]$_['Członek']; Sid = $sids[[string]$_['Członek']] } }) }
            }
            Start-HostOperation -Module $m -Name 'Usuwanie administratorów' -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
                param($P)
                $groupName = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')).Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
                foreach ($member in $P.Members) {
                    try {
                        # Zabezpieczenie przed odcięciem dostępu: wbudowane konto Administrator i Domain Admins
                        if ($member.Sid -match '-500$' -or $member.Sid -match '^S-1-5-21-.+-512$') { throw 'Pominięto: wbudowane konto Administrator / grupa Domain Admins.' }
                        $identity = if ($member.Sid) { $member.Sid } else { $member.Name }
                        if (Get-Command Remove-LocalGroupMember -ErrorAction SilentlyContinue) {
                            Remove-LocalGroupMember -SID 'S-1-5-32-544' -Member $identity -ErrorAction Stop
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
        })
    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'Dodaj konto/grupę (DOMENA\nazwa, kilka – rozdziel przecinkiem):')
    $m.NewMember = Add-TextBox $row2 280
    [void](Add-Button $row2 'Dodaj na zaznaczonych komputerach' $m {
            param($m)
            $members = @(Split-ListText $m.NewMember.Text)
            if ($members.Count -eq 0) { Show-Warning 'Podaj konto lub grupę, np. FIRMA\Helpdesk.'; return }
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            if (-not (Confirm-Action ("Dodać {0} do lokalnej grupy Administratorzy na {1} komputer(ach)?" -f ($members -join ', '), $targets.Count) $targets)) { return }
            Start-HostOperation -Module $m -Name 'Dodawanie administratorów' -Targets $targets -Output Log -Parameters @{ Members = $members } -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
                param($P)
                $groupName = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-544')).Translate([System.Security.Principal.NTAccount]).Value.Split('\')[-1]
                foreach ($member in $P.Members) {
                    try {
                        if (Get-Command Add-LocalGroupMember -ErrorAction SilentlyContinue) {
                            Add-LocalGroupMember -SID 'S-1-5-32-544' -Member $member -ErrorAction Stop
                        }
                        else {
                            $group = [ADSI]("WinNT://{0}/{1},group" -f $env:COMPUTERNAME, $groupName)
                            $group.Add('WinNT://' + ($member -replace '\\', '/'))
                        }
                        [pscustomobject]@{ 'Członek' = $member; 'Wynik' = 'Dodano' }
                    }
                    catch {
                        $msg = $_.Exception.Message
                        if ($_.FullyQualifiedErrorId -like 'MemberExists*') { $msg = 'już jest członkiem grupy' }
                        [pscustomobject]@{ 'Członek' = $member; 'Wynik' = "Błąd – $msg" }
                    }
                }
            }
        })
}

Register-Module -Key 'LocalUsers' -Category 'Bezpieczeństwo' -Title 'Konta lokalne' -Description 'Lokalne konta użytkowników: stan, ostatnie logowanie, wiek hasła; włączanie, wyłączanie i ustawianie hasła zaznaczonych kont.' -Build {
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
                    [pscustomobject]@{
                        'Konto'              = $_.Name
                        'Pełna nazwa'        = $_.FullName
                        'Włączone'           = $_.Enabled
                        'Ostatnie logowanie' = $_.LastLogon
                        'Hasło ustawione'    = $_.PasswordLastSet
                        'Hasło wygasa'       = $_.PasswordExpires
                        'Opis'               = $_.Description
                        'SID'                = [string]$_.SID
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
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Pokaż konta' $m -Primary $m.Actions.List)
    $userAction = {
        param($m, $s)
        $op = [string]$s.Tag['Op']
        $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Konto')
        if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli konta lokalne.'; return }
        $items = foreach ($h in $byHost.Keys) { foreach ($i in $byHost[$h]) { '{0}: {1}' -f $h, $i['Konto'] } }
        $password = $null
        if ($op -eq 'Password') {
            $password = Show-PasswordDialog -Message 'Nowe hasło zostanie ustawione dla wszystkich zaznaczonych kont.'
            if (-not $password) { return }
        }
        $question = @{ Enable = 'Włączyć wybrane konta lokalne?'; Disable = 'Wyłączyć wybrane konta lokalne?'; Password = 'Ustawić nowe hasło dla wybranych kont lokalnych?' }[$op]
        if (-not (Confirm-Action $question @($items))) { return }
        $per = @{}
        foreach ($h in $byHost.Keys) { $per[$h] = @{ Op = $op; Password = $password; Names = @($byHost[$h] | ForEach-Object { [string]$_['Konto'] }) } }
        Start-HostOperation -Module $m -Name "Konta lokalne – $op" -Targets @($byHost.Keys) -PerTarget $per -Output Log -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
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
    [void](Add-Label $row '   Zaznaczone konta:')
    $b = Add-Button $row 'Włącz' $m $userAction
    $b.Tag['Op'] = 'Enable'
    $b = Add-Button $row 'Wyłącz' $m -Danger $userAction
    $b.Tag['Op'] = 'Disable'
    $b = Add-Button $row 'Ustaw hasło…' $m -Danger $userAction
    $b.Tag['Op'] = 'Password'
}
#endregion

#region Moduły: Udostępnianie
Register-Module -Key 'Shares' -Category 'Udostępnianie' -Title 'Udziały sieciowe' -Description 'Udziały SMB na zaznaczonych komputerach: podgląd, uprawnienia, tworzenie (z uprawnieniami udziału i opcjonalnie NTFS) oraz usuwanie.' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Udziały' -Targets $targets -Parameters @{ ShowSpecial = $m.ShowSpecial.Checked } -ScriptBlock {
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
    $row = Add-ToolbarRow $m
    $m.ShowSpecial = Add-CheckBox $row 'Pokaż udziały administracyjne (C$, ADMIN$…)' $false
    [void](Add-Button $row 'Pokaż udziały' $m -Primary $m.Actions.List)
    [void](Add-Button $row 'Uprawnienia zaznaczonych' $m {
            param($m)
            $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Udział')
            if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli udziały.'; return }
            $per = @{}
            foreach ($h in $byHost.Keys) { $per[$h] = @{ Names = @($byHost[$h] | ForEach-Object { [string]$_['Udział'] }) } }
            $m.PermissionRows = New-Object System.Collections.ArrayList
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
                    [void]$m.PermissionRows.Add([pscustomobject]@{
                            'Komputer'    = $r.Target
                            'Udział'      = Get-ObjectValue $d 'Udział'
                            'Konto'       = Get-ObjectValue $d 'Konto'
                            'Typ'         = Get-ObjectValue $d 'Typ'
                            'Uprawnienie' = Get-ObjectValue $d 'Uprawnienie'
                        })
                }
            } -OnComplete {
                param($m)
                if ($m.PermissionRows.Count -eq 0) { return }
                # Okno pokazujemy poza obsługą timera, żeby nie wstrzymywać innych trwających operacji
                Invoke-Deferred -Module $m -Action { param($m) Show-GridDialog -Title 'Uprawnienia udziałów' -Rows @($m.PermissionRows) }
            }
        })
    [void](Add-Button $row 'Usuń zaznaczone' $m -Danger {
            param($m)
            $byHost = Get-SelectedRowsByHost -Module $m -Columns @('Udział')
            if ($byHost.Count -eq 0) { Show-Warning 'Zaznacz w tabeli udziały do usunięcia.'; return }
            $items = foreach ($h in $byHost.Keys) { foreach ($i in $byHost[$h]) { '{0}: {1}' -f $h, $i['Udział'] } }
            if (-not (Confirm-Action 'Usunąć wybrane udziały? Dane w folderach pozostaną nienaruszone.' @($items))) { return }
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
        })

    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'Nowy udział:' -Bold)
    [void](Add-Label $row2 'Nazwa:')
    $m.NewName = Add-TextBox $row2 120
    [void](Add-Label $row2 'Ścieżka na hoście:')
    $m.NewPath = Add-TextBox $row2 200 'D:\Udzial'
    [void](Add-Label $row2 'Opis:')
    $m.NewDesc = Add-TextBox $row2 160
    $row3 = Add-ToolbarRow $m
    [void](Add-Label $row3 'Pełna kontrola:')
    $m.NewFull = Add-TextBox $row3 150
    [void](Add-Label $row3 'Zmiana:')
    $m.NewChange = Add-TextBox $row3 150
    [void](Add-Label $row3 'Odczyt:')
    $m.NewRead = Add-TextBox $row3 150
    $m.NewNtfs = Add-CheckBox $row3 'Nadaj też NTFS' $false
    $m.NewCreate = Add-CheckBox $row3 'Utwórz folder' $true
    [void](Add-Button $row3 'Utwórz na zaznaczonych' $m {
            param($m)
            $name = $m.NewName.Text.Trim()
            $path = $m.NewPath.Text.Trim()
            if (-not $name -or -not $path) { Show-Warning 'Podaj nazwę udziału i ścieżkę folderu na hoście.'; return }
            if ($path -notmatch '^[A-Za-z]:\\') { Show-Warning 'Ścieżka musi być lokalną ścieżką na hoście, np. D:\Dane\Projekty.'; return }
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            $params = @{
                Name = $name; Path = $path; Description = $m.NewDesc.Text.Trim()
                Full = @(Split-ListText $m.NewFull.Text); Change = @(Split-ListText $m.NewChange.Text); Read = @(Split-ListText $m.NewRead.Text)
                Ntfs = $m.NewNtfs.Checked; CreateFolder = $m.NewCreate.Checked
            }
            if (-not (Confirm-Action ("Utworzyć udział {0} -> {1} na {2} komputer(ach)?" -f $name, $path, $targets.Count) $targets)) { return }
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
        })
}
#endregion

#region Moduły: Active Directory
Register-Module -Key 'ComputerAccount' -Category 'Active Directory' -Title 'Konto komputera' -Description 'Informacje o obiekcie komputera w AD, test i naprawa kanału zaufania, reset konta, włączanie/wyłączanie i przenoszenie do innej jednostki OU.' -Build {
    param($m)
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'Konto komputera – informacje' -Targets $targets -Local -ScriptBlock {
            param($Target, $P, $Ctx)
            Import-Module ActiveDirectory -ErrorAction Stop
            $ad = @{}
            if ($Ctx.Server) { $ad.Server = $Ctx.Server }
            if ($Ctx.Credential) { $ad.Credential = $Ctx.Credential }
            $props = @('Enabled', 'LastLogonDate', 'PasswordLastSet', 'OperatingSystem', 'OperatingSystemVersion', 'whenCreated', 'Description', 'IPv4Address', 'ManagedBy', 'Location')
            $c = Get-ADComputer -Identity ($Target -split '\.')[0] -Properties $props @ad -ErrorAction Stop
            [pscustomobject]@{
                'Włączone'               = $c.Enabled
                'Ostatnie logowanie'     = $c.LastLogonDate
                'Hasło konta zmienione'  = $c.PasswordLastSet
                'Dni od zmiany hasła'    = $(if ($c.PasswordLastSet) { [int]((Get-Date) - $c.PasswordLastSet).TotalDays } else { $null })
                'System'                 = $c.OperatingSystem
                'Wersja'                 = $c.OperatingSystemVersion
                'IPv4'                   = $c.IPv4Address
                'Utworzono'              = $c.whenCreated
                'Opis'                   = $c.Description
                'Lokalizacja'            = $c.Location
                'Jednostka OU'           = ($c.DistinguishedName -replace '^CN=(?:\\.|[^,])+,', '')
            }
        }
    }
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Informacje z AD' $m -Primary $m.Actions.List)
    [void](Add-Button $row 'Test kanału zaufania' $m {
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
                    'Kanał zaufania'   = $(if ($ok) { 'Poprawny' } else { 'Błąd – kanał uszkodzony' })
                    'Domena'           = $domain
                    'Kontroler domeny' = $dc
                }
            }
        })
    [void](Add-Button $row 'Napraw kanał zaufania' $m -Danger {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            $cred = Get-EffectiveCredential
            if (-not $cred) {
                $cred = Show-CredentialDialog -Message 'Naprawa kanału zaufania wymaga poświadczeń domenowych z prawem resetu konta komputera (nie przechodzą przez WinRM automatycznie).'
                if (-not $cred) { return }
            }
            if (-not (Confirm-Action "Naprawić kanał zaufania (Test-ComputerSecureChannel -Repair) na $($targets.Count) komputer(ach)?" $targets)) { return }
            Start-HostOperation -Module $m -Name 'Naprawa kanału zaufania' -Targets $targets -Output Log -Parameters @{ Credential = $cred; Server = [string]$script:Settings.DomainController } -ScriptBlock {
                param($P)
                $tp = @{ Repair = $true; Credential = $P.Credential; ErrorAction = 'Stop' }
                if ($P.Server) { $tp.Server = $P.Server }
                if (Test-ComputerSecureChannel @tp) { 'Kanał zaufania naprawiony.' } else { 'Błąd – naprawa nie powiodła się.' }
            }
        })
    $adAction = {
        param($m, $s)
        $op = [string]$s.Tag['Op']
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        $params = @{ Op = $op; TargetOU = '' }
        switch ($op) {
            'Reset' { $question = 'Zresetować konta komputerów w AD? Komputery stracą relację zaufania i trzeba będzie ponownie dołączyć je do domeny lub naprawić kanał.' }
            'Enable' { $question = 'Włączyć konta komputerów w AD?' }
            'Disable' { $question = 'Wyłączyć konta komputerów w AD? Użytkownicy nie zalogują się kontem domenowym na tych komputerach.' }
            'Move' {
                $ou = Select-OrganizationalUnit -Title 'Docelowa jednostka organizacyjna'
                if (-not $ou) { return }
                $params.TargetOU = $ou
                $question = "Przenieść konta komputerów do:`r`n$ou ?"
            }
        }
        if (-not (Confirm-Action $question $targets)) { return }
        Start-HostOperation -Module $m -Name "Konto komputera – $op" -Targets $targets -Local -Output Log -Parameters $params -OnComplete { param($m) & $m.Actions.List $m } -ScriptBlock {
            param($Target, $P, $Ctx)
            Import-Module ActiveDirectory -ErrorAction Stop
            $ad = @{}
            if ($Ctx.Server) { $ad.Server = $Ctx.Server }
            if ($Ctx.Credential) { $ad.Credential = $Ctx.Credential }
            $c = Get-ADComputer -Identity ($Target -split '\.')[0] @ad -ErrorAction Stop
            switch ($P.Op) {
                'Reset' {
                    # Odpowiednik "Resetuj konto" z konsoli ADUC: hasło domyślne = nazwa komputera małymi literami (bez $, maks. 14 znaków)
                    $plain = $c.SamAccountName.TrimEnd('$').ToLowerInvariant()
                    if ($plain.Length -gt 14) { $plain = $plain.Substring(0, 14) }
                    Set-ADAccountPassword -Identity $c -Reset -NewPassword (ConvertTo-SecureString $plain -AsPlainText -Force) @ad -ErrorAction Stop
                    'Konto zresetowane – dołącz komputer ponownie do domeny lub napraw kanał zaufania.'
                }
                'Enable' { Enable-ADAccount -Identity $c @ad -ErrorAction Stop; 'Konto włączone.' }
                'Disable' { Disable-ADAccount -Identity $c @ad -ErrorAction Stop; 'Konto wyłączone.' }
                'Move' { Move-ADObject -Identity $c.DistinguishedName -TargetPath $P.TargetOU @ad -ErrorAction Stop; "Przeniesiono do $($P.TargetOU)." }
            }
        }
    }
    $row2 = Add-ToolbarRow $m
    [void](Add-Label $row2 'Operacje w AD:')
    foreach ($a in @(@('Włącz konto', 'Enable', $false), @('Wyłącz konto', 'Disable', $true), @('Przenieś do OU…', 'Move', $false), @('Resetuj konto', 'Reset', $true))) {
        $b = Add-Button $row2 $a[0] $m -Danger:$a[2] $adAction
        $b.Tag['Op'] = $a[1]
    }
}

Register-Module -Key 'Laps' -Category 'Active Directory' -Title 'LAPS' -Description 'Hasła lokalnego administratora z AD (Windows LAPS, także szyfrowane, oraz LAPS legacy), wymuszanie zmiany hasła i bezpieczne kopiowanie do schowka (czyszczonego po 60 s).' -Build {
    param($m)
    $m.SecretColumns = @('Hasło')
    $m.Actions.List = {
        param($m)
        $targets = @(Get-TargetComputers)
        if (-not $targets) { return }
        Start-HostOperation -Module $m -Name 'LAPS – odczyt' -Targets $targets -Local -ScriptBlock {
            param($Target, $P, $Ctx)
            Import-Module ActiveDirectory -ErrorAction Stop
            $ad = @{}
            if ($Ctx.Server) { $ad.Server = $Ctx.Server }
            if ($Ctx.Credential) { $ad.Credential = $Ctx.Credential }
            function ConvertFrom-FileTimeValue($Value) {
                if ($null -eq $Value -or [string]$Value -eq '' -or [string]$Value -eq '0') { return $null }
                try { return [DateTime]::FromFileTimeUtc([int64][string]$Value).ToLocalTime() } catch { return $null }
            }
            $name = ($Target -split '\.')[0]
            $c = Get-ADComputer -Identity $name -Properties * @ad -ErrorAction Stop
            $row = $null
            $lapsError = ''
            if (Get-Command Get-LapsADPassword -ErrorAction SilentlyContinue) {
                try {
                    $lp = @{ Identity = $name; AsPlainText = $true; ErrorAction = 'Stop' }
                    if ($Ctx.Server) { $lp.DomainController = $Ctx.Server }
                    if ($Ctx.Credential) { $lp.Credential = $Ctx.Credential }
                    $info = Get-LapsADPassword @lp
                    if ($info -and $info.Password) {
                        $row = [pscustomobject]@{
                            'Rozwiązanie' = 'Windows LAPS'
                            'Konto'       = $info.Account
                            'Hasło'       = [string]$info.Password
                            'Zmienione'   = $info.PasswordUpdateTime
                            'Wygasa'      = $info.ExpirationTimestamp
                            'Źródło'      = [string]$info.Source
                            'Uwagi'       = ''
                        }
                    }
                    elseif ($info) { $lapsError = "Status odszyfrowania: $($info.DecryptionStatus)" }
                }
                catch { $lapsError = $_.Exception.Message }
            }
            if (-not $row -and $c.'msLAPS-Password') {
                try {
                    $json = [string]$c.'msLAPS-Password' | ConvertFrom-Json
                    $row = [pscustomobject]@{
                        'Rozwiązanie' = 'Windows LAPS'
                        'Konto'       = $json.n
                        'Hasło'       = [string]$json.p
                        'Zmienione'   = $null
                        'Wygasa'      = ConvertFrom-FileTimeValue $c.'msLAPS-PasswordExpirationTime'
                        'Źródło'      = 'msLAPS-Password'
                        'Uwagi'       = ''
                    }
                }
                catch { }
            }
            if (-not $row -and $c.'ms-Mcs-AdmPwd') {
                $row = [pscustomobject]@{
                    'Rozwiązanie' = 'LAPS (legacy)'
                    'Konto'       = '(wg zasad GPO)'
                    'Hasło'       = [string]$c.'ms-Mcs-AdmPwd'
                    'Zmienione'   = $null
                    'Wygasa'      = ConvertFrom-FileTimeValue $c.'ms-Mcs-AdmPwdExpirationTime'
                    'Źródło'      = 'ms-Mcs-AdmPwd'
                    'Uwagi'       = ''
                }
            }
            if (-not $row) {
                $note = 'Brak hasła LAPS albo brak uprawnień do odczytu.'
                if ($c.'msLAPS-EncryptedPassword') { $note = 'Hasło jest zaszyfrowane – brak uprawnień do odszyfrowania lub brak modułu LAPS (RSAT).' }
                if ($lapsError) { $note += " ($lapsError)" }
                $row = [pscustomobject]@{ 'Rozwiązanie' = 'brak'; 'Konto' = ''; 'Hasło' = ''; 'Zmienione' = $null; 'Wygasa' = $null; 'Źródło' = ''; 'Uwagi' = $note }
            }
            $row
        }
    }
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Pokaż hasła LAPS' $m -Primary $m.Actions.List)
    [void](Add-Button $row 'Kopiuj hasło zaznaczonego' $m {
            param($m)
            $selected = @(Get-SelectedResultRows -Module $m) | Select-Object -First 1
            $secret = [string](Get-ObjectValue $selected 'Hasło')
            if (-not $secret) { Show-Warning 'Zaznacz wiersz z hasłem.'; return }
            Set-ClipboardSecret -Text $secret -Seconds 60
            Write-Log ("Skopiowano hasło LAPS komputera {0}; schowek zostanie wyczyszczony po 60 s." -f (Get-ObjectValue $selected 'Komputer')) 'OK'
        })
    $m.ProcessNow = Add-CheckBox $row 'Od razu przetwórz zasady na hoście' $true
    [void](Add-Button $row 'Wymuś zmianę hasła' $m -Danger {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            if (-not (Confirm-Action "Wymusić zmianę hasła LAPS (ustawienie wygaśnięcia na teraz) dla $($targets.Count) komputer(ów)?" $targets)) { return }
            Start-HostOperation -Module $m -Name 'LAPS – wymuszenie zmiany' -Targets $targets -Local -Output Log -Parameters @{ ProcessNow = $m.ProcessNow.Checked } -ScriptBlock {
                param($Target, $P, $Ctx)
                Import-Module ActiveDirectory -ErrorAction Stop
                $ad = @{}
                if ($Ctx.Server) { $ad.Server = $Ctx.Server }
                if ($Ctx.Credential) { $ad.Credential = $Ctx.Credential }
                $name = ($Target -split '\.')[0]
                $c = Get-ADComputer -Identity $name -Properties * @ad -ErrorAction Stop
                $done = @()
                if ($c.'msLAPS-PasswordExpirationTime' -or $c.'msLAPS-EncryptedPassword' -or $c.'msLAPS-Password') {
                    if (Get-Command Set-LapsADPasswordExpirationTime -ErrorAction SilentlyContinue) {
                        $sp = @{ Identity = $name; WhenEffective = (Get-Date); ErrorAction = 'Stop' }
                        if ($Ctx.Server) { $sp.DomainController = $Ctx.Server }
                        if ($Ctx.Credential) { $sp.Credential = $Ctx.Credential }
                        Set-LapsADPasswordExpirationTime @sp | Out-Null
                    }
                    else {
                        Set-ADComputer -Identity $c -Replace @{ 'msLAPS-PasswordExpirationTime' = 0 } @ad -ErrorAction Stop
                    }
                    $done += 'Windows LAPS'
                }
                if ($c.'ms-Mcs-AdmPwdExpirationTime' -or $c.'ms-Mcs-AdmPwd') {
                    Set-ADComputer -Identity $c -Replace @{ 'ms-Mcs-AdmPwdExpirationTime' = 0 } @ad -ErrorAction Stop
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
                    try { $msg += '; na hoście uruchomiono ' + (Invoke-Command @ic) }
                    catch { $msg += '; nie udało się wymusić przetwarzania na hoście: ' + $_.Exception.Message }
                }
                $msg + '.'
            }
        })
}

Register-Module -Key 'Rename' -Category 'Active Directory' -Title 'Zmiana nazwy komputerów' -Description 'Wsadowa zmiana nazw z autonumeracją i walidacją (NetBIOS: maks. 15 znaków). Kolumnę «Nowa nazwa» można edytować bezpośrednio w tabeli.' -Build {
    param($m)
    $m.Validate = {
        param($m)
        [void]$m.Grid.EndEdit()
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
            if ($problem) { $r['Walidacja'] = $problem } else { $r['Walidacja'] = 'OK'; $valid++ }
        }
        return $valid
    }
    $row = Add-ToolbarRow $m
    [void](Add-Button $row 'Wczytaj zaznaczone komputery' $m -Primary {
            param($m)
            $targets = @(Get-TargetComputers)
            if (-not $targets) { return }
            Reset-ResultTable -Module $m
            foreach ($t in $targets) { Add-ResultRows -Module $m -Computer $t -Objects @([pscustomobject]@{ 'Nowa nazwa' = ''; 'Walidacja' = ''; 'Wynik' = '' }) }
            $m.Grid.ReadOnly = $false
            foreach ($c in $m.Grid.Columns) { $c.ReadOnly = ($c.DataPropertyName -ne 'Nowa nazwa') }
            Resize-ResultColumns -Module $m
        })
    [void](Add-Label $row '   Prefiks:')
    $m.Prefix = Add-TextBox $row 90 'PC-'
    [void](Add-Label $row 'Start:')
    $m.Start = Add-Numeric $row 0 99999 1 70
    [void](Add-Label $row 'Cyfr:')
    $m.Pad = Add-Numeric $row 1 8 3 50
    [void](Add-Label $row 'Sufiks:')
    $m.Suffix = Add-TextBox $row 70
    [void](Add-Button $row 'Autonumeracja' $m {
            param($m)
            if ($m.Table.Rows.Count -eq 0) { Show-Warning 'Najpierw wczytaj zaznaczone komputery.'; return }
            $n = [int]$m.Start.Value
            $format = 'D' + [int]$m.Pad.Value
            # Numeracja w kolejności widocznej w tabeli (z uwzględnieniem sortowania i filtra)
            foreach ($drv in @($m.View | ForEach-Object { $_ })) {
                $drv.Row['Nowa nazwa'] = ($m.Prefix.Text.Trim() + $n.ToString($format) + $m.Suffix.Text.Trim()).ToUpperInvariant()
                $n++
            }
            [void](& $m.Validate $m)
        })
    $row2 = Add-ToolbarRow $m
    [void](Add-Button $row2 'Sprawdź nazwy' $m {
            param($m)
            if ($m.Table.Rows.Count -eq 0) { Show-Warning 'Najpierw wczytaj zaznaczone komputery.'; return }
            $valid = & $m.Validate $m
            Write-Log "Poprawnych nowych nazw: $valid z $($m.Table.Rows.Count)."
        })
    $m.Restart = Add-CheckBox $row2 'Restart po zmianie (za 30 s)' $true
    [void](Add-Button $row2 'Zmień nazwy' $m -Danger {
            param($m)
            if ($m.Table.Rows.Count -eq 0) { Show-Warning 'Najpierw wczytaj zaznaczone komputery.'; return }
            $valid = & $m.Validate $m
            if ($valid -eq 0) { Show-Warning 'Brak poprawnych nowych nazw – sprawdź kolumnę «Walidacja».'; return }
            $per = @{}
            $items = @()
            foreach ($r in $m.Table.Rows) {
                if ([string]$r['Walidacja'] -ne 'OK') { continue }
                $old = [string]$r['Komputer']
                $new = ([string]$r['Nowa nazwa']).Trim()
                $per[$old] = @{ NewName = $new; Restart = $m.Restart.Checked; Credential = $null }
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
            if (-not (Confirm-Action $question $items)) { return }
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
                    $row['Wynik'] = 'Zmieniono'
                    Rename-HostRow -OldName $r.Target -NewName ([string]$row['Nowa nazwa']).Trim()
                }
                else { $row['Wynik'] = 'Błąd – ' + ((@($r.Errors)) -join ' | ') }
            }
        })
    [void](Add-Label $row2 'Komputery muszą być włączone; po zmianie nazwy lista po lewej jest aktualizowana.' -Hint)
}
#endregion

#region Okno główne
function Update-CredentialLabel {
    $label = $script:UI.CredLabel
    if ($script:State.UseCurrent -or -not $script:State.Credential) {
        $label.Text = "Działam jako: $env:USERDOMAIN\$env:USERNAME"
        $label.ForeColor = [System.Drawing.Color]::FromArgb(40, 40, 40)
    }
    else {
        $label.Text = "Działam jako: $($script:State.Credential.UserName) (poświadczenia alternatywne)"
        $label.ForeColor = [System.Drawing.Color]::DarkRed
    }
}

function New-TopBar {
    $hm = $script:UI.Modules['Hosts']
    $bar = New-FlowRow
    $bar.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
    $bar.BackColor = [System.Drawing.Color]::FromArgb(236, 241, 248)

    $chk = Add-CheckBox -Parent $bar -Text 'Bieżący użytkownik' -Checked $true
    $script:UI.CredCurrent = $chk
    $btnCred = New-PlainButton -Parent $bar -Text 'Inne poświadczenia…'
    $lbl = Add-Label -Parent $bar -Text ''
    $lbl.Margin = New-Object System.Windows.Forms.Padding(6, 7, 24, 3)
    $script:UI.CredLabel = $lbl

    Register-ControlHandler -Control $btnCred -EventName 'Click' -Module $hm -Action {
        $cred = Show-CredentialDialog -UserName $(if ($script:State.Credential) { $script:State.Credential.UserName } else { '' })
        if (-not $cred) { return }
        $script:State.Credential = $cred
        $script:State.UseCurrent = $false
        $script:UI.CredCurrent.Checked = $false
        Update-CredentialLabel
        Write-Log "Ustawiono poświadczenia alternatywne: $($cred.UserName)" -Module ''
    }
    Register-ControlHandler -Control $chk -EventName 'CheckedChanged' -Module $hm -Action {
        param($m, $s)
        if (-not $s.Checked -and -not $script:State.Credential) {
            $cred = Show-CredentialDialog
            if (-not $cred) { $s.Checked = $true; return }
            $script:State.Credential = $cred
            Write-Log "Ustawiono poświadczenia alternatywne: $($cred.UserName)" -Module ''
        }
        $script:State.UseCurrent = $s.Checked
        Update-CredentialLabel
    }

    [void](Add-Label -Parent $bar -Text 'Równolegle:')
    $numThrottle = Add-Numeric -Parent $bar -Minimum 1 -Maximum 64 -Value ([int]$script:Settings.ThrottleLimit) -Width 55
    $numThrottle.add_ValueChanged({ param($s, $e) try { Set-EngineThrottle -Limit ([int]$s.Value) } catch { } })
    [void](Add-Label -Parent $bar -Text 'Limit połączenia (s):')
    $numTimeout = Add-Numeric -Parent $bar -Minimum 5 -Maximum 300 -Value ([int]$script:Settings.TimeoutSec) -Width 55
    $numTimeout.add_ValueChanged({ param($s, $e) $script:Settings.TimeoutSec = [int]$s.Value })
    [void](Add-Label -Parent $bar -Text 'Kontroler domeny:')
    $txtDc = Add-TextBox -Parent $bar -Width 150 -Text ([string]$script:Settings.DomainController)
    $txtDc.add_TextChanged({ param($s, $e) $script:Settings.DomainController = $s.Text.Trim() })
    $tip = New-Object System.Windows.Forms.ToolTip
    $tip.SetToolTip($txtDc, 'Opcjonalnie: kontroler domeny dla operacji AD (puste = wybór automatyczny).')
    $tip.SetToolTip($numThrottle, 'Liczba hostów obsługiwanych jednocześnie.')
    $tip.SetToolTip($numTimeout, 'Czas oczekiwania na nawiązanie połączenia WinRM z hostem.')
    $script:UI.ToolTip = $tip
    Update-CredentialLabel
    return $bar
}

function New-LogPanel {
    $panel = New-Object System.Windows.Forms.Panel
    $head = New-FlowRow
    [void](Add-Label -Parent $head -Text 'Dziennik operacji' -Bold)
    $btnClear = New-PlainButton -Parent $head -Text 'Wyczyść'
    $btnClear.add_Click({ try { $script:UI.LogBox.Clear() } catch { } })
    $btnOpen = New-PlainButton -Parent $head -Text 'Otwórz plik dziennika'
    $btnOpen.add_Click({
            try {
                if (Test-Path -LiteralPath $script:App.LogFile) { Start-Process -FilePath 'notepad.exe' -ArgumentList ('"{0}"' -f $script:App.LogFile) }
                else { Show-Warning 'Plik dziennika jeszcze nie istnieje.' }
            }
            catch { Show-Error 'Nie można otworzyć pliku dziennika.' $_ }
        })
    [void](Add-Label -Parent $head -Text $script:App.LogFile -Hint)
    $box = New-Object System.Windows.Forms.RichTextBox
    $box.ReadOnly = $true
    $box.BackColor = [System.Drawing.Color]::FromArgb(252, 252, 252)
    $box.Font = $script:UI.FontMono
    $box.WordWrap = $false
    $box.DetectUrls = $false
    $box.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $script:UI.LogBox = $box
    Add-DockStack -Parent $panel -Top @($head) -Fill $box
    return $panel
}

function New-StatusStrip {
    $strip = New-Object System.Windows.Forms.StatusStrip
    $label = New-Object System.Windows.Forms.ToolStripStatusLabel
    $label.Spring = $true
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $label.Text = 'Gotowe'
    $progress = New-Object System.Windows.Forms.ToolStripProgressBar
    $progress.Size = New-Object System.Drawing.Size(220, 16)
    $progress.Visible = $false
    $cancel = New-Object System.Windows.Forms.ToolStripButton
    $cancel.Text = 'Anuluj operacje'
    $cancel.ForeColor = [System.Drawing.Color]::DarkRed
    $cancel.Enabled = $false
    $cancel.add_Click({ try { Stop-AllOperations } catch { } })
    [void]$strip.Items.Add($label)
    [void]$strip.Items.Add($progress)
    [void]$strip.Items.Add($cancel)
    $script:UI.StatusLabel = $label
    $script:UI.StatusProgress = $progress
    $script:UI.StatusCancel = $cancel
    return $strip
}

function New-ModuleTree {
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.HideSelection = $false
    $tree.FullRowSelect = $true
    $tree.ShowLines = $false
    $tree.ShowNodeToolTips = $true
    $tree.ItemHeight = 24
    $tree.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $tree.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 251)
    # Czcionka drzewa pogrubiona (kategorie), moduły zwykłą - inaczej pogrubione etykiety byłyby przycinane
    $tree.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
    $regular = New-Object System.Drawing.Font('Segoe UI', 9.5)
    foreach ($category in $script:UI.Categories) {
        $defs = @($script:UI.ModuleDefs | Where-Object { $_.Category -eq $category })
        if ($defs.Count -eq 0) { continue }
        $catNode = $tree.Nodes.Add($category)
        $catNode.ForeColor = [System.Drawing.Color]::FromArgb(30, 60, 110)
        foreach ($d in $defs) {
            $node = $catNode.Nodes.Add($d.Title)
            $node.Tag = $d.Key
            $node.NodeFont = $regular
            $node.ToolTipText = $d.Description
        }
    }
    $tree.ExpandAll()
    $tree.add_AfterSelect({
            param($s, $e)
            try {
                if ($e.Node.Tag) { Show-Module -Key ([string]$e.Node.Tag) }
                elseif ($e.Node.Nodes.Count -gt 0) { $s.SelectedNode = $e.Node.Nodes[0] }
            }
            catch {
                Write-Log "Nie udało się otworzyć modułu: $($_.Exception.Message)" 'ERROR' -Module ''
                Show-Error 'Nie udało się otworzyć modułu.' $_
            }
        })
    $script:UI.Tree = $tree
    return $tree
}

function Select-ModuleNode([string]$Key) {
    foreach ($cat in $script:UI.Tree.Nodes) {
        foreach ($node in $cat.Nodes) {
            if ([string]$node.Tag -eq $Key) { $script:UI.Tree.SelectedNode = $node; return $true }
        }
    }
    return $false
}

function New-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Domain Ops $($script:AppVersion) – zdalna administracja komputerami w domenie"
    $form.Font = $script:UI.Font
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.MinimumSize = New-Object System.Drawing.Size(1100, 650)
    $form.Size = New-Object System.Drawing.Size([Math]::Max(1100, [int]$script:Settings.WindowWidth), [Math]::Max(650, [int]$script:Settings.WindowHeight))
    if ($script:Settings.WindowMaximized) { $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized }
    try { $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Join-Path $env:SystemRoot 'System32\mmc.exe')) } catch { }
    $script:UI.Form = $form

    $hostPanel = New-HostPanel
    $topBar = New-TopBar
    $status = New-StatusStrip
    $logPanel = New-LogPanel
    $tree = New-ModuleTree

    $content = New-Object System.Windows.Forms.Panel
    $content.BackColor = [System.Drawing.SystemColors]::Window
    $script:UI.ContentHost = $content

    $splitNav = New-Object System.Windows.Forms.SplitContainer
    $splitNav.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $splitNav.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel1
    $splitNav.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tree.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitNav.Panel1.Controls.Add($tree)
    $content.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitNav.Panel2.Controls.Add($content)

    $splitMain = New-Object System.Windows.Forms.SplitContainer
    $splitMain.Orientation = [System.Windows.Forms.Orientation]::Vertical
    $splitMain.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel1
    $splitMain.Dock = [System.Windows.Forms.DockStyle]::Fill
    $hostPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitMain.Panel1.Controls.Add($hostPanel)
    $splitMain.Panel2.Controls.Add($splitNav)

    $splitLog = New-Object System.Windows.Forms.SplitContainer
    $splitLog.Orientation = [System.Windows.Forms.Orientation]::Horizontal
    $splitLog.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel2
    $splitLog.Panel1.Controls.Add($splitMain)
    $logPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $splitLog.Panel2.Controls.Add($logPanel)

    Add-DockStack -Parent $form -Top @($topBar) -Fill $splitLog -Bottom @($status)

    $script:UI.SplitMain = $splitMain
    $script:UI.SplitNav = $splitNav
    $script:UI.SplitLog = $splitLog

    # Timery: odbiór wyników operacji w tle i czyszczenie schowka z haseł
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 200
    $timer.add_Tick({ try { Update-Operations } catch { } })
    $script:Engine.Timer = $timer
    $clipTimer = New-Object System.Windows.Forms.Timer
    $clipTimer.add_Tick({ try { $script:Clipboard.Timer.Stop(); Clear-ClipboardSecret } catch { } })
    $script:Clipboard.Timer = $clipTimer

    $form.add_Shown({
            param($s, $e)
            # Rozmiary paneli ustawiane dopiero po pokazaniu okna (wcześniej kontenery mają rozmiar domyślny)
            try {
                $script:UI.SplitLog.SplitterDistance = [Math]::Max(200, $script:UI.SplitLog.Height - 190)
                $script:UI.SplitLog.Panel2MinSize = 80
            }
            catch { }
            try {
                $script:UI.SplitMain.SplitterDistance = 430
                $script:UI.SplitMain.Panel1MinSize = 320
            }
            catch { }
            try {
                $script:UI.SplitNav.SplitterDistance = 220
                $script:UI.SplitNav.Panel1MinSize = 160
            }
            catch { }
            Write-Log ("Uruchomiono {0} {1} jako {2}\{3} (PowerShell {4})." -f $script:App.Name, $script:AppVersion, $env:USERDOMAIN, $env:USERNAME, $PSVersionTable.PSVersion) -Module ''
            if (-not (Select-ModuleNode -Key ([string]$script:Settings.LastModule))) { [void](Select-ModuleNode -Key 'Connectivity') }
            if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
                Write-Log 'Brak modułu ActiveDirectory (RSAT) – listę komputerów dodaj ręcznie lub z pliku; moduły AD będą niedostępne.' 'WARN' -Module ''
            }
        })

    $form.add_FormClosing({
            param($s, $e)
            try {
                if (@($script:Engine.Operations).Count -gt 0) {
                    if (-not (Confirm-Action 'Trwają operacje w tle. Przerwać je i zamknąć program?')) { $e.Cancel = $true; return }
                    Stop-AllOperations
                }
                $script:Settings.WindowMaximized = ($s.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)
                if ($s.WindowState -eq [System.Windows.Forms.FormWindowState]::Normal) {
                    $script:Settings.WindowWidth = $s.Width
                    $script:Settings.WindowHeight = $s.Height
                }
                $script:Settings.SearchBase = $script:UI.HostSearchBase.Text.Trim()
                $script:Settings.NameFilter = $script:UI.HostNameFilter.Text.Trim()
                $script:Settings.OnlyEnabled = $script:UI.HostOnlyEnabled.Checked
                Export-Settings
            }
            catch { }
        })
    return $form
}
#endregion

#region Uruchomienie
Import-Settings
try {
    if (-not (Test-Path -LiteralPath $script:App.LogDir)) { New-Item -ItemType Directory -Path $script:App.LogDir -Force | Out-Null }
}
catch { }

$mainForm = New-MainForm
try {
    [void]$mainForm.ShowDialog()
}
finally {
    Close-Engine
    Clear-ClipboardSecret
    $mainForm.Dispose()
}
#endregion
