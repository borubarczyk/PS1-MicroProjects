Add-Type -AssemblyName PresentationFramework
# --- KONFIGURACJA I HISTORIA ---
# 1. Konfiguracja ląduje w stałym AppData (nigdy nie zniknie)
$appDataPath = Join-Path $env:APPDATA "CyberGenPro"
if (-not (Test-Path $appDataPath)) { New-Item -ItemType Directory -Path $appDataPath | Out-Null }
$configFile = Join-Path $appDataPath "CyberGenConfig.json"

# 2. Historia ląduje w ulotnym Tempie (bezpieczeństwo po restarcie)
$historyFile = Join-Path $env:TEMP "CyberGenHistory.json"

$historyList = @()
if (Test-Path $historyFile) {
    try {
        $loadedHist = @(Get-Content $historyFile -Raw | ConvertFrom-Json)
        $cutoff = (Get-Date).AddHours(-24)
        $historyList = @($loadedHist | Where-Object { [datetime]$_.Date -ge $cutoff })
    } catch {}
}

# --- KRYPTOGRAFIA I LOGIKA ---
function Get-RandomIndex([int]$Max) {
    $bytes = [Byte[]]::new(4)
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
    return ([System.BitConverter]::ToUInt32($bytes, 0) % $Max)
}

function Shuffle-String([string]$str) {
    $chars = $str.ToCharArray()
    for ($i = $chars.Length - 1; $i -gt 0; $i--) {
        $j = Get-RandomIndex ($i + 1)
        $temp = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $temp
    }
    return -join $chars
}

function Generate-Password($cfg) {
    $upper = if ($cfg.ExcludeAmbiguous) { "ABCDEFGHJKLMNPQRSTUVWXYZ" } else { "ABCDEFGHIJKLMNOPQRSTUVWXYZ" }
    $lower = if ($cfg.ExcludeAmbiguous) { "abcdefghijkmnopqrstuvwxyz" } else { "abcdefghijklmnopqrstuvwxyz" }
    $digits = if ($cfg.ExcludeAmbiguous) { "23456789" } else { "0123456789" }
    $charPool = ""
    
    if ($cfg.Upper) { $charPool += $upper }
    if ($cfg.Lower) { $charPool += $lower }
    if ($cfg.Digits) { $charPool += $digits }
    if ($cfg.Special -and $cfg.SpecialChars.Length -gt 0) { $charPool += $cfg.SpecialChars }
    if ([string]::IsNullOrEmpty($charPool)) { $charPool = "abcdefghijklmnopqrstuvwxyz" }

    $res = ""
    $charsNeeded = $cfg.Length

    if ($cfg.ForceRules -and $cfg.Length -ge 3) {
        $res += $digits[$(Get-RandomIndex $digits.Length)]
        $res += $digits[$(Get-RandomIndex $digits.Length)]
        $specPool = if ($cfg.SpecialChars) { $cfg.SpecialChars } else { "!@#$%^&*" }
        $res += $specPool[$(Get-RandomIndex $specPool.Length)]
        $charsNeeded -= 3
    }

    for ($i = 0; $i -lt $charsNeeded; $i++) {
        $res += $charPool[$(Get-RandomIndex $charPool.Length)]
    }

    if ($cfg.ForceRules) { return Shuffle-String $res }
    return $res
}

# --- XAML UI ---
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="CyberGen Pro (PS7)" Height="620" Width="850" 
        Background="{DynamicResource BgMainBrush}" Foreground="{DynamicResource TextBrush}" 
        WindowStartupLocation="CenterScreen" FontFamily="Segoe UI">
    
    <Window.Resources>
        <Style TargetType="Border">
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Background" Value="{DynamicResource BgPanelBrush}"/>
            <Setter Property="Padding" Value="15"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="{DynamicResource AccentBrush}"/>
            <Setter Property="Foreground" Value="#FAFAFA"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="8"/>
            <Setter Property="MinHeight" Value="35"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="6" Padding="0" Margin="0">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{DynamicResource InputBrush}"/>
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="6"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Margin" Value="0,0,0,5"/>
        </Style>
    </Window.Resources>

    <Grid Margin="15">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="1.1*"/>
            <ColumnDefinition Width="15"/>
            <ColumnDefinition Width="1*"/>
        </Grid.ColumnDefinitions>

        <!-- LEWA KOLUMNA -->
        <StackPanel Grid.Column="0">
            <Grid Margin="0,0,0,10">
                <TextBlock Text="Kreator Hasła" FontSize="20" FontWeight="Bold" VerticalAlignment="Center"/>
                <CheckBox Name="ChkTheme" Content="💡 Tryb Jasny" HorizontalAlignment="Right" VerticalAlignment="Center" FontWeight="Bold" Margin="0"/>
            </Grid>
            
            <Border>
                <StackPanel>
                    <TextBlock Text="Długość (użyj scrolla):" Foreground="{DynamicResource TextDimBrush}" Margin="0,0,0,5"/>
                    <Grid Margin="0,0,0,10">
                        <Slider Name="SldLength" Minimum="8" Maximum="128" Value="24" TickFrequency="1" IsSnapToTickEnabled="True" Margin="0,0,35,0" VerticalAlignment="Center"/>
                        <TextBlock Name="TxtLengthVal" Text="24" HorizontalAlignment="Right" FontWeight="Bold" FontSize="16"/>
                    </Grid>
                    
                    <WrapPanel Margin="0,0,0,5">
                        <CheckBox Name="ChkUp" Content="Wielkie (A-Z)" IsChecked="True" Margin="0,0,10,5"/>
                        <CheckBox Name="ChkLow" Content="Małe (a-z)" IsChecked="True" Margin="0,0,10,5"/>
                        <CheckBox Name="ChkDig" Content="Cyfry (0-9)" IsChecked="True" Margin="0,0,10,5"/>
                        <CheckBox Name="ChkAmb" Content="Bez mylących (l,1,O)" IsChecked="True" Margin="0,0,10,5"/>
                    </WrapPanel>
                    
                    <CheckBox Name="ChkForce" Content="Wymuś min. 2 cyfry i 1 znak spec." Foreground="{DynamicResource AccentBrush}" FontWeight="Bold" IsChecked="True"/>
                    
                    <StackPanel Orientation="Horizontal" Margin="0,5,0,0">
                        <CheckBox Name="ChkSpec" Content="Specjalne:" IsChecked="True" VerticalAlignment="Center" Margin="0,0,10,0"/>
                        <TextBox Name="TxtSpec" Text="!@#$%^&amp;*()-_=+[]{}" Width="150"/>
                    </StackPanel>
                </StackPanel>
            </Border>

            <Border Background="{DynamicResource InputBrush}">
                <StackPanel>
                    <TextBox Name="TxtResult" FontSize="18" FontWeight="Bold" FontFamily="Consolas" TextAlignment="Center" IsReadOnly="True" Foreground="{DynamicResource SuccessBrush}" Margin="0,0,0,10" BorderThickness="0" Background="Transparent"/>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="8"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Button Name="BtnRefresh" Content="ODŚWIEŻ" Background="{DynamicResource BorderBrush}" Grid.Column="0"/>
                        <Button Name="BtnCopy" Content="KOPIUJ HASŁO" Grid.Column="2"/>
                    </Grid>
                </StackPanel>
            </Border>

            <Border Margin="0">
                <StackPanel>
                    <TextBlock Text="Bramka SMS / E-mail" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                    <Grid Margin="0,0,0,5">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="60"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                        <TextBlock Text="Adres:" Foreground="{DynamicResource TextDimBrush}" VerticalAlignment="Center"/>
                        <TextBox Name="TxtSmsEmail" Grid.Column="1"/>
                    </Grid>
                    <Grid Margin="0,0,0,10">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="60"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                        <TextBlock Text="Temat:" Foreground="{DynamicResource TextDimBrush}" VerticalAlignment="Center"/>
                        <TextBox Name="TxtSmsSubj" Text="Nowe hasło" Grid.Column="1"/>
                    </Grid>
                    <Button Name="BtnSms" Content="WYŚLIJ VIA MAILTO" Background="{DynamicResource SuccessBrush}"/>
                </StackPanel>
            </Border>
        </StackPanel>

        <!-- PRAWA KOLUMNA -->
        <StackPanel Grid.Column="2">
            <Border Margin="0,0,0,10" BorderBrush="{DynamicResource AccentBrush}" BorderThickness="1">
                <StackPanel>
                    <TextBlock Text="Generowanie Masowe" FontSize="16" FontWeight="Bold" Margin="0,0,0,10"/>
                    <TextBlock Text="Ilość haseł (scroll):" Foreground="{DynamicResource TextDimBrush}" Margin="0,0,0,5"/>
                    <Grid Margin="0,0,0,10">
                        <Slider Name="SldMass" Minimum="1" Maximum="100" Value="10" TickFrequency="1" IsSnapToTickEnabled="True" Margin="0,0,35,0" VerticalAlignment="Center"/>
                        <TextBlock Name="TxtMassVal" Text="10" HorizontalAlignment="Right" FontWeight="Bold" FontSize="16"/>
                    </Grid>
                    <Button Name="BtnMass" Content="GENERUJ I KOPIUJ PACZKĘ" Background="#8B5CF6"/>
                </StackPanel>
            </Border>

            <Border VerticalAlignment="Stretch" Margin="0">
                <StackPanel>
                    <TextBlock Text="Historia (ostatnie 24h)" FontSize="16" FontWeight="Bold" Margin="0,0,0,5"/>
                    <TextBlock Text="Dwuklik na haśle wybiórczo je kopiuje." Foreground="{DynamicResource TextDimBrush}" FontSize="11" Margin="0,0,0,5"/>
                    <ListBox Name="LstHistory" Height="220" Background="{DynamicResource InputBrush}" Foreground="{DynamicResource TextBrush}" FontFamily="Consolas" BorderThickness="1" BorderBrush="{DynamicResource BorderBrush}" Margin="0,0,0,10"/>
                    <Button Name="BtnClearHist" Content="WYCZYŚĆ HISTORIĘ" Background="#EF4444"/>
                </StackPanel>
            </Border>
        </StackPanel>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# --- BINDING ---
$SldLength = $window.FindName("SldLength"); $TxtLengthVal = $window.FindName("TxtLengthVal")
$ChkUp = $window.FindName("ChkUp"); $ChkLow = $window.FindName("ChkLow"); $ChkDig = $window.FindName("ChkDig")
$ChkAmb = $window.FindName("ChkAmb"); $ChkForce = $window.FindName("ChkForce"); $ChkSpec = $window.FindName("ChkSpec")
$TxtSpec = $window.FindName("TxtSpec"); $TxtResult = $window.FindName("TxtResult")
$BtnRefresh = $window.FindName("BtnRefresh"); $BtnCopy = $window.FindName("BtnCopy")
$SldMass = $window.FindName("SldMass"); $TxtMassVal = $window.FindName("TxtMassVal"); $BtnMass = $window.FindName("BtnMass")
$LstHistory = $window.FindName("LstHistory"); $BtnClearHist = $window.FindName("BtnClearHist")
$TxtSmsEmail = $window.FindName("TxtSmsEmail"); $TxtSmsSubj = $window.FindName("TxtSmsSubj"); $BtnSms = $window.FindName("BtnSms")
$ChkTheme = $window.FindName("ChkTheme")

$script:isUpdating = $true

function Update-Theme([bool]$IsLight) {
    $themeColors = if ($IsLight) {
        @{
            "BgMainBrush" = "#F3F4F6"; "BgPanelBrush" = "#FFFFFF"; "InputBrush" = "#E5E7EB"
            "TextBrush" = "#111827"; "TextDimBrush" = "#4B5563"; "BorderBrush" = "#D1D5DB"
            "AccentBrush" = "#2563EB"; "SuccessBrush" = "#059669"
        }
    } else {
        @{
            "BgMainBrush" = "#09090B"; "BgPanelBrush" = "#18181B"; "InputBrush" = "#27272A"
            "TextBrush" = "#FAFAFA"; "TextDimBrush" = "#D4D4D8"; "BorderBrush" = "#3F3F46"
            "AccentBrush" = "#3B82F6"; "SuccessBrush" = "#10B981"
        }
    }

    foreach ($key in $themeColors.Keys) {
        $colorObj = [System.Windows.Media.ColorConverter]::ConvertFromString($themeColors[$key])
        $brushObj = [System.Windows.Media.SolidColorBrush]::new($colorObj)
        if ($window.Resources.Contains($key)) { $window.Resources.Remove($key) }
        $window.Resources.Add($key, $brushObj)
    }
}

if (Test-Path $configFile) {
    try {
        $cfg = Get-Content $configFile -Raw | ConvertFrom-Json
        if ($null -ne $cfg.Length) { $SldLength.Value = $cfg.Length }
        if ($null -ne $cfg.Upper) { $ChkUp.IsChecked = $cfg.Upper }
        if ($null -ne $cfg.Lower) { $ChkLow.IsChecked = $cfg.Lower }
        if ($null -ne $cfg.Digits) { $ChkDig.IsChecked = $cfg.Digits }
        if ($null -ne $cfg.ExcludeAmbiguous) { $ChkAmb.IsChecked = $cfg.ExcludeAmbiguous }
        if ($null -ne $cfg.ForceRules) { $ChkForce.IsChecked = $cfg.ForceRules }
        if ($null -ne $cfg.Special) { $ChkSpec.IsChecked = $cfg.Special }
        if ($null -ne $cfg.SpecialChars) { $TxtSpec.Text = $cfg.SpecialChars }
        if ($null -ne $cfg.EmailAddress) { $TxtSmsEmail.Text = $cfg.EmailAddress }
        if ($null -ne $cfg.EmailSubject) { $TxtSmsSubj.Text = $cfg.EmailSubject }
        if ($null -ne $cfg.IsLightTheme) { $ChkTheme.IsChecked = $cfg.IsLightTheme }
    } catch {}
}

$TxtLengthVal.Text = [math]::Round($SldLength.Value)
$TxtMassVal.Text = [math]::Round($SldMass.Value)
Update-Theme $ChkTheme.IsChecked

function Get-CurrentConfig {
    return @{
        Length = [int]$SldLength.Value; Upper = [bool]$ChkUp.IsChecked
        Lower = [bool]$ChkLow.IsChecked; Digits = [bool]$ChkDig.IsChecked
        ExcludeAmbiguous = [bool]$ChkAmb.IsChecked; ForceRules = [bool]$ChkForce.IsChecked
        Special = [bool]$ChkSpec.IsChecked; SpecialChars = $TxtSpec.Text
    }
}

function Add-ToHistory($pass) {
    $entry = @{ Pass = $pass; Date = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
    $global:historyList = @($entry) + $global:historyList
    $LstHistory.Items.Insert(0, "[$($entry.Date)] $($entry.Pass)")
}

function Do-Refresh {
    if ($script:isUpdating) { return }
    $cfg = Get-CurrentConfig
    $newPass = Generate-Password $cfg
    $TxtResult.Text = $newPass
    Add-ToHistory $newPass
}

foreach ($h in $historyList) { $LstHistory.Items.Add("[$($h.Date)] $($h.Pass)") | Out-Null }

$script:isUpdating = $false

# ZDARZENIA
$ChkTheme.add_Click({ Update-Theme $ChkTheme.IsChecked })

$SldLength.add_PreviewMouseWheel({
    param($sender, $e)
    if ($e.Delta -gt 0) { $SldLength.Value++ } else { $SldLength.Value-- }
    $e.Handled = $true
})
$SldMass.add_PreviewMouseWheel({
    param($sender, $e)
    if ($e.Delta -gt 0) { $SldMass.Value++ } else { $SldMass.Value-- }
    $e.Handled = $true
})

$SldLength.add_ValueChanged({ $TxtLengthVal.Text = [math]::Round($SldLength.Value); Do-Refresh })
$SldMass.add_ValueChanged({ $TxtMassVal.Text = [math]::Round($SldMass.Value) })

$ChkUp.add_Click({ Do-Refresh }); $ChkLow.add_Click({ Do-Refresh }); $ChkDig.add_Click({ Do-Refresh })
$ChkAmb.add_Click({ Do-Refresh }); $ChkForce.add_Click({ Do-Refresh })
$ChkSpec.add_Click({ $TxtSpec.IsEnabled = $ChkSpec.IsChecked; Do-Refresh })
$TxtSpec.add_TextChanged({ Do-Refresh })

$BtnRefresh.add_Click({ Do-Refresh })
$BtnCopy.add_Click({ [System.Windows.Clipboard]::SetText($TxtResult.Text) })

# NOWOCZESNY ZEGAR DLA PS7 (bez zmiennych globalnych)
$BtnMass.add_Click({
    $script:isUpdating = $true
    $cfg = Get-CurrentConfig
    $newBatch = @()
    
    for ($i = 0; $i -lt [int]$SldMass.Value; $i++) {
        $p = Generate-Password $cfg
        Add-ToHistory $p
        $newBatch += $p
    }
    $TxtResult.Text = $p
    
    [System.Windows.Clipboard]::SetText(($newBatch -join [Environment]::NewLine))
    $BtnMass.Content = "SKOPIOWANO PACZKĘ!"
    
    $localTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $localTimer.Interval = [TimeSpan]::FromSeconds(1.5)
    $localTimer.add_Tick({
        param($sender, $e)
        $BtnMass.Content = "GENERUJ I KOPIUJ PACZKĘ"
        $sender.Stop()
    })
    $localTimer.Start()

    $script:isUpdating = $false
})

$LstHistory.add_MouseDoubleClick({
    if ($LstHistory.SelectedItem) {
        $pass = ($LstHistory.SelectedItem -split '] ')[1]
        [System.Windows.Clipboard]::SetText($pass)
    }
})

$BtnClearHist.add_Click({ $global:historyList = @(); $LstHistory.Items.Clear() })

$BtnSms.add_Click({
    $email = $TxtSmsEmail.Text.Trim(); $subj = [uri]::EscapeDataString($TxtSmsSubj.Text); $body = [uri]::EscapeDataString($TxtResult.Text)
    if ($email) { Start-Process "mailto:$email`?subject=$subj&body=$body" }
})

$window.add_Closing({
    $cfgToSave = Get-CurrentConfig
    $cfgToSave.EmailAddress = $TxtSmsEmail.Text
    $cfgToSave.EmailSubject = $TxtSmsSubj.Text
    $cfgToSave.IsLightTheme = [bool]$ChkTheme.IsChecked
    
    $cfgToSave | ConvertTo-Json | Set-Content -Path $configFile -Encoding UTF8
    $global:historyList | ConvertTo-Json | Set-Content -Path $historyFile -Encoding UTF8
})

Do-Refresh
$window.ShowDialog() | Out-Null