#requires -Version 5.1
<#
NTFS-Audit-GUI.ps1 (clean sync build)
Cel: Audyt NTFS rekursywnie, wykrywanie wpisów nadanych użytkownikom (zamiast grup), GUI WPF z filtrami i eksportem.
Brak zewnętrznych modułów.
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Set-StrictMode -Version Latest
if (-not ('CodexNativeLookup' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;

public enum CodexSidNameUse
{
    User = 1,
    Group = 2,
    Domain = 3,
    Alias = 4,
    WellKnownGroup = 5,
    DeletedAccount = 6,
    Invalid = 7,
    Unknown = 8,
    Computer = 9,
    Label = 10
}

public static class CodexNativeLookup
{
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool LookupAccountName(
        string systemName,
        string accountName,
        byte[] sid,
        ref uint sidSize,
        StringBuilder referencedDomainName,
        ref uint domainNameSize,
        out CodexSidNameUse accountType);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool LookupAccountSid(
        string systemName,
        byte[] sid,
        StringBuilder name,
        ref uint cchName,
        StringBuilder referencedDomainName,
        ref uint cchReferencedDomainName,
        out CodexSidNameUse accountType);
}
'@ -ErrorAction Stop
}
$ErrorActionPreference = 'Stop'

# ========================= Helpers =========================
$PrincipalCache = @{}
$script:CancellationTokenSource = $null
$script:IsDebugLogging = $false

function Write-LogUI {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string]$Level = 'INFO'
    )
    if ($Level -eq 'DEBUG' -and -not $script:IsDebugLogging) { return }
    $ts = (Get-Date).ToString('HH:mm:ss')
    $text = '[{0}] {1}: {2}' -f $ts, $Level, $Message

    try {
        if ($script:rbLog) {
            $appendAction = [System.Action]{
                $script:rbLog.AppendText($text + [Environment]::NewLine)
                $script:rbLog.ScrollToEnd()
            }
            $dispatcher = $script:rbLog.Dispatcher
            if ($dispatcher -and -not $dispatcher.CheckAccess()) {
                $null = $dispatcher.BeginInvoke($appendAction)
            } else {
                $appendAction.Invoke()
            }
        }
    } catch {
        Write-Host "[LOG][WARN] UI append failed: $($_.Exception.Message)"
    }

    Write-Host $text
}
function Write-LogDebug {
    param([string]$Message)
    Write-LogUI $Message 'DEBUG'
}

function Test-ExcelAvailable {
    try { [void][type]::GetTypeFromProgID('Excel.Application'); return $true } catch { return $false }
}

function Get-PrincipalInfo {
    <#
      Wejscie: NTAccount/SID string lub obiekt; Wyjscie: PSCustomObject z Name/Domain/Type/Scope
      LookupAccountName/LookupAccountSid (Win32) zapewnia szybkie rozpoznawanie bez kosztow ADSI.
    #>
    param([Parameter(Mandatory)][object]$Identity)

    $acct = switch ($Identity) {
        { $_ -is [System.Security.Principal.NTAccount] } { $_; break }
        { $_ -is [System.Security.Principal.SecurityIdentifier] } { $_; break }
        { $_ -is [string] } {
            if ($_ -like 'S-1-*') {
                try { [System.Security.Principal.SecurityIdentifier]::new($_) }
                catch { [System.Security.Principal.NTAccount]::new($_) }
            } else {
                [System.Security.Principal.NTAccount]::new($_)
            }
            break
        }
        default { return [pscustomobject]@{ Name=$Identity.ToString(); Domain=''; Type='Unknown'; Scope='Unknown' } }
    }

    $key = $acct.Value.ToUpperInvariant()
    if ($PrincipalCache.ContainsKey($key)) { return $PrincipalCache[$key] }

    $name = ''
    $domain = ''
    $type = 'Unknown'
    $scope = 'Unknown'
    $lookupSucceeded = $false
    $stringComparer = [System.StringComparer]::OrdinalIgnoreCase

    $getScope = {
        param([string]$Dom)
        if ([string]::IsNullOrEmpty($Dom)) { return 'Local' }
        if ($stringComparer.Equals($Dom, 'BUILTIN') -or $stringComparer.Equals($Dom, 'NT AUTHORITY')) { return 'Builtin' }
        if ($stringComparer.Equals($Dom, $env:COMPUTERNAME)) { return 'Local' }
        return 'Domain'
    }

    $getTypeFromUse = {
        param([CodexSidNameUse]$Use)
        switch ($Use) {
            ([CodexSidNameUse]::User)           { 'User' }
            ([CodexSidNameUse]::Computer)       { 'Computer' }
            ([CodexSidNameUse]::Group)          { 'Group' }
            ([CodexSidNameUse]::Alias)          { 'Group' }
            ([CodexSidNameUse]::WellKnownGroup) { 'Group' }
            default                             { 'Unknown' }
        }
    }

    if ($acct -is [System.Security.Principal.SecurityIdentifier]) {
        $sidBytes = New-Object byte[] $acct.BinaryLength
        $acct.GetBinaryForm($sidBytes, 0)
        $nameBuilder = New-Object System.Text.StringBuilder 256
        $domainBuilder = New-Object System.Text.StringBuilder 256
        $cchName = [uint32]$nameBuilder.Capacity
        $cchDomain = [uint32]$domainBuilder.Capacity
        $use = [CodexSidNameUse]::Unknown
        $lookupSucceeded = [CodexNativeLookup]::LookupAccountSid($null, $sidBytes, $nameBuilder, [ref]$cchName, $domainBuilder, [ref]$cchDomain, [ref]$use)
        if (-not $lookupSucceeded -and [System.Runtime.InteropServices.Marshal]::GetLastWin32Error() -eq 122) {
            $nameBuilder = New-Object System.Text.StringBuilder([int][Math]::Max($cchName, 1))
            $domainBuilder = New-Object System.Text.StringBuilder([int][Math]::Max($cchDomain, 1))
            $cchName = [uint32]$nameBuilder.Capacity
            $cchDomain = [uint32]$domainBuilder.Capacity
            $use = [CodexSidNameUse]::Unknown
            $lookupSucceeded = [CodexNativeLookup]::LookupAccountSid($null, $sidBytes, $nameBuilder, [ref]$cchName, $domainBuilder, [ref]$cchDomain, [ref]$use)
        }
        if ($lookupSucceeded) {
            $name = if ($nameBuilder.Length -gt 0) { $nameBuilder.ToString() } else { $acct.Value }
            $domain = $domainBuilder.ToString()
            $type = & $getTypeFromUse $use
            $scope = & $getScope $domain
        }
    } else {
        $acctValue = $acct.Value
        $parts = $acctValue -split '\\', 2
        if ($parts.Count -eq 2) { $domain = $parts[0]; $name = $parts[1] } else { $name = $acctValue }

        $sidSize = 0u
        $domainSize = 0u
        $use = [CodexSidNameUse]::Unknown
        $domainBuilder = New-Object System.Text.StringBuilder 0
        $lookupSucceeded = [CodexNativeLookup]::LookupAccountName($null, $acctValue, $null, [ref]$sidSize, $domainBuilder, [ref]$domainSize, [ref]$use)
        if (-not $lookupSucceeded) {
            $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
            if ($err -eq 122 -and $sidSize -gt 0) {
                $sidBuffer = New-Object byte[] $sidSize
                $domainBuilder = New-Object System.Text.StringBuilder([int][Math]::Max($domainSize, 1))
                $use = [CodexSidNameUse]::Unknown
                $lookupSucceeded = [CodexNativeLookup]::LookupAccountName($null, $acctValue, $sidBuffer, [ref]$sidSize, $domainBuilder, [ref]$domainSize, [ref]$use)
            }
        }
        if ($lookupSucceeded) {
            if ($domainBuilder.Length -gt 0) { $domain = $domainBuilder.ToString() }
            $type = & $getTypeFromUse $use
            $scope = & $getScope $domain
        }
    }

    if (-not $lookupSucceeded) {
        if (-not $name) { $name = $acct.Value }
        $scope = & $getScope $domain
        Write-LogDebug ("LookupAccountName nie powiodlo sie dla: {0}" -f $acct.Value)
    }

    $obj = [pscustomobject]@{
        Name   = $name
        Domain = $domain
        Type   = $type
        Scope  = $scope
    }
    $PrincipalCache[$key] = $obj
    return $obj
}

function Get-NTFSAclReport {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$RootPath,
        [Parameter()]
        [psobject]$Options,
        [System.Threading.CancellationToken]$CancellationToken = [System.Threading.CancellationToken]::None
    )

    Write-LogUI "Rozpoczynanie skanowania: $RootPath"
    Write-LogDebug ("Skanowanie - root: {0} (Thread: {1})" -f $RootPath, [System.Threading.Thread]::CurrentThread.ManagedThreadId)

    $results = New-Object System.Collections.ObjectModel.ObservableCollection[object]
    $queue = New-Object System.Collections.Generic.Queue[System.IO.DirectoryInfo]
    $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

    $includeInherited = $true
    $onlyUsers = $false
    $onlyGroups = $false
    if ($null -ne $Options) {
        if ($Options.PSObject.Properties.Match('IncludeInherited')) { $includeInherited = [bool]$Options.IncludeInherited }
        if ($Options.PSObject.Properties.Match('OnlyUsers')) { $onlyUsers = [bool]$Options.OnlyUsers }
        if ($Options.PSObject.Properties.Match('OnlyGroups')) { $onlyGroups = [bool]$Options.OnlyGroups }
    }

    try {
        Write-LogDebug "Pobieranie katalogu root: $RootPath"
        $rootItem = Get-Item -LiteralPath $RootPath -ErrorAction Stop
        if (-not $rootItem.PSIsContainer) { throw "Sciezka nie jest katalogiem: $RootPath" }
        $normalizedRoot = $rootItem.FullName.TrimEnd('\\').ToLowerInvariant()
        if (-not $visited.Add($normalizedRoot)) {
            Write-LogUI "Katalog root juz przetworzony: $RootPath" 'WARN'
            return $results
        }
        $queue.Enqueue($rootItem)
        Write-LogDebug ("Dodano katalog root do kolejki: {0}" -f $rootItem.FullName)
    } catch {
        Write-LogUI ("Blad dostepu do katalogu root: {0} - {1}" -f $RootPath, $_.Exception.Message) 'ERROR'
        return $results
    }

    $processedFolders = 0
    while ($queue.Count -gt 0) {
        if ($CancellationToken.IsCancellationRequested) {
            Write-LogUI 'Skanowanie przerwane przez uzytkownika' 'WARN'
            break
        }

        $current = $queue.Dequeue()
        Write-LogDebug ("Przetwarzanie katalogu: {0} (Queue: {1})" -f $current.FullName, $queue.Count)
        $processedFolders++

        $isReparsePoint = [bool]($current.Attributes -band [System.IO.FileAttributes]::ReparsePoint)
        if (-not $isReparsePoint) {
            try {
                Write-LogDebug ("Pobieranie podkatalogow: {0}" -f $current.FullName)
                $children = Get-ChildItem -LiteralPath $current.FullName -Directory -Force -ErrorAction Stop
                $addedChildren = 0
                foreach ($child in $children) {
                    $key = $child.FullName.TrimEnd('\\').ToLowerInvariant()
                    $childIsReparse = [bool]($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint)
                    if (-not $childIsReparse -and $visited.Add($key)) {
                        $queue.Enqueue($child)
                        $addedChildren++
                    } else {
                        Write-LogDebug ("Pominieto reparse point lub juz przetworzony: {0}" -f $child.FullName)
                    }
                }
                Write-LogDebug ("Dodano {0} podkatalogow z: {1}" -f $addedChildren, $current.FullName)
            } catch {
                Write-LogUI ("Blad dostepu do podkatalogow: {0} - {1}" -f $current.FullName, $_.Exception.Message) 'WARN'
                continue
            }
        } else {
            Write-LogDebug ("Pominieto reparse point: {0}" -f $current.FullName)
        }

        try {
            Write-LogDebug ("Pobieranie ACL dla: {0}" -f $current.FullName)
            $acl = Get-Acl -LiteralPath $current.FullName -ErrorAction Stop
            $aceCount = if ($null -ne $acl.Access) { $acl.Access.Count } else { 0 }
            Write-LogDebug ("Analiza ACL - wpisow: {0} | Katalog: {1}" -f $aceCount, $current.FullName)
        } catch {
            Write-LogUI ("Blad Get-Acl: {0} - {1}" -f $current.FullName, $_.Exception.Message) 'WARN'
            continue
        }

        foreach ($ace in $acl.Access) {
            if (-not $includeInherited -and $ace.IsInherited) { continue }

            $principal = Get-PrincipalInfo -Identity $ace.IdentityReference
            if ($onlyUsers -and $principal.PrincipalType -ne 'User') { continue }
            if ($onlyGroups -and $principal.PrincipalType -ne 'Group') { continue }

            Write-LogDebug ("Przetwarzanie ACE: {0}" -f $ace.IdentityReference.Value)

            $results.Add([pscustomobject]@{
                FolderPath       = $current.FullName
                Identity         = $ace.IdentityReference.Value
                PrincipalName    = $principal.Name
                PrincipalDomain  = $principal.Domain
                PrincipalType    = $principal.Type
                PrincipalScope   = $principal.Scope
                Rights           = [string]$ace.FileSystemRights
                AccessType       = [string]$ace.AccessControlType
                IsInherited      = [bool]$ace.IsInherited
                IsExplicit       = -not [bool]$ace.IsInherited
                InheritanceFlags = [string]$ace.InheritanceFlags
                PropagationFlags = [string]$ace.PropagationFlags
            })
        }
    }

    Write-LogDebug ("Skanowanie zakonczone - katalogow: {0}, wpisow ACL: {1}" -f $processedFolders, $results.Count)
    Write-LogUI ("Koniec. Przeskanowano katalogow: {0}. Wpisow ACL: {1}" -f $processedFolders, $results.Count)

    return $results
}

function Export-ReportCSV {
    param([Parameter(Mandatory)][System.Collections.IEnumerable]$Data)

    $path = Join-Path ([Environment]::GetFolderPath('Desktop')) ("NTFS_Audit_" + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.csv')
    $Data | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    Write-LogUI "Zapisano CSV: $path"
    [System.Windows.MessageBox]::Show("Zapisano CSV:`n$path", 'Eksport CSV', 'OK', 'Information') | Out-Null
}

function Export-ReportXLSX {
    param([Parameter(Mandatory)][System.Collections.IEnumerable]$Data)

    if (-not (Test-ExcelAvailable)) {
        [System.Windows.MessageBox]::Show('Brak zainstalowanego Microsoft Excel. Uzyj CSV.', 'Eksport XLSX', 'OK', 'Warning') | Out-Null
        return
    }

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $cols = 'FolderPath','Identity','PrincipalName','PrincipalDomain','PrincipalType','PrincipalScope','Rights','AccessType','IsInherited','IsExplicit','InheritanceFlags','PropagationFlags'
    for ($i = 0; $i -lt $cols.Count; $i++) { $ws.Cells.Item(1, $i + 1) = $cols[$i] }
    $rowIndex = 2
    foreach ($row in $Data) {
        for ($i = 0; $i -lt $cols.Count; $i++) {
            $ws.Cells.Item($rowIndex, $i + 1) = [string]($row.($cols[$i]))
        }
        $rowIndex++
    }

    $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
    $path = Join-Path ([Environment]::GetFolderPath('Desktop')) ("NTFS_Audit_" + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.xlsx')
    $wb.SaveAs($path)
    $wb.Close($true)
    $xl.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

    Write-LogUI "Zapisano XLSX: $path"
    [System.Windows.MessageBox]::Show("Zapisano XLSX:`n$path", 'Eksport XLSX', 'OK', 'Information') | Out-Null
}

# Whitelist użytkowników dopuszczalnych wyjątków (regex)
$script:AllowedUserSamPatterns = @('^administrator$','^admin_\w+$','^admin-l\d+$')

function Test-IsProblemACE {
    param([object]$Row)
    # Problem: ACE niedziedziczony nadany USEROWI, który nie pasuje do whitelisty
    if ($Row.PrincipalType -eq 'User' -and $Row.IsExplicit) {
        $sam = ($Row.PrincipalName ?? '').ToString()
        foreach ($rx in $script:AllowedUserSamPatterns) { if ($sam -match $rx) { return $false } }
        return $true
    }
    return $false
}

function Build-TreeView {
    $tv.Items.Clear()
    $groups = $AllItems | Group-Object FolderPath | Sort-Object Name
    foreach ($g in $groups) {
        $folder = $g.Name
        $rows = $g.Group
        $userExp = @($rows | Where-Object { $_.PrincipalType -eq 'User' -and $_.IsExplicit })
        $grpAny  = @($rows | Where-Object { $_.PrincipalType -eq 'Group' })
        $hasBadUser = ($userExp | Where-Object { Test-IsProblemACE $_ } | Select-Object -First 1)
        # Folder PROBLEM: ma user-explicit i brak grup LUB ma >1 user-explicit LUB user-explicit nie-whitelisted
        $hasProblem = ($userExp.Count -gt 0 -and $grpAny.Count -eq 0) -or ($userExp.Count -gt 1) -or $hasBadUser

        $ti = New-Object System.Windows.Controls.TreeViewItem
        $ti.Header = ("{0}  [ACE:{1} | Uexp:{2} | Grp:{3}]" -f $folder, $rows.Count, $userExp.Count, $grpAny.Count)
        $ti.Foreground = if ($hasProblem) { [System.Windows.Media.Brushes]::Red } else { [System.Windows.Media.Brushes]::Green }

        foreach ($r in $rows) {
            $child = New-Object System.Windows.Controls.TreeViewItem
            $child.Header = ("{0} | {1} | {2} | {3}" -f $r.PrincipalType, $r.Identity, $r.AccessType, $r.Rights)
            if (Test-IsProblemACE $r) { $child.Foreground = [System.Windows.Media.Brushes]::Red }
            elseif ($r.PrincipalType -eq 'Group') { $child.Foreground = [System.Windows.Media.Brushes]::Green }
            $ti.Items.Add($child) | Out-Null
        }
        $tv.Items.Add($ti) | Out-Null
    }
}


# ========================= GUI (WPF) =========================
$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="NTFS Audit GUI" Height="720" Width="1240" MinWidth="1000" WindowStartupLocation="CenterScreen">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="170"/>
    </Grid.RowDefinitions>

    <!-- Top: command area -->
    <Grid Grid.Row="0" Margin="0,0,0,10">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <StackPanel Orientation="Horizontal" Grid.Row="0">
        <Label Content="Katalog:" VerticalAlignment="Center"/>
        <TextBox x:Name="tbPath" Width="520" Margin="6,0"/>
        <Button x:Name="btnBrowse" Content="Wybierz..." Width="100"/>
        <Button x:Name="btnScan" Content="Skanuj" Width="100" Margin="6,0"/>
      </StackPanel>
      <StackPanel Orientation="Horizontal" Grid.Row="1" Margin="0,6,0,0">
        <ProgressBar x:Name="pb" Width="220" Height="20"/>
        <TextBox x:Name="tbSearch" Width="220" Margin="12,0,0,0" ToolTip="Filtruj po sciezce lub tozsamosci"/>
        <CheckBox x:Name="cbOnlyUsers" Content="Tylko Uzytkownicy" Margin="12,0,0,0"/>
        <CheckBox x:Name="cbOnlyGroups" Content="Tylko Grupy" Margin="6,0,0,0"/>
        <CheckBox x:Name="cbOnlyExplicit" Content="Tylko Niedziedziczone" Margin="6,0,0,0"/>
        <Button x:Name="btnToggleDebug" Content="Debug log OFF" Width="130" Margin="12,0,0,0"/>
      </StackPanel>
    </Grid>

    <!-- Middle: tabbed results -->
    <TabControl Grid.Row="1">
      <TabItem Header="Tabela">
        <Grid>
          <DataGrid x:Name="dg" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False" ItemsSource="{Binding}" Margin="0,6,0,0">
            <DataGrid.Columns>
              <DataGridTextColumn Header="Folder" Binding="{Binding FolderPath}" Width="*"/>
              <DataGridTextColumn Header="Identity" Binding="{Binding Identity}" Width="220"/>
              <DataGridTextColumn Header="Typ" Binding="{Binding PrincipalType}" Width="80"/>
              <DataGridTextColumn Header="Zakres" Binding="{Binding PrincipalScope}" Width="80"/>
              <DataGridTextColumn Header="Prawa" Binding="{Binding Rights}" Width="200"/>
              <DataGridTextColumn Header="Dostęp" Binding="{Binding AccessType}" Width="80"/>
              <DataGridCheckBoxColumn Header="Niedziedz." Binding="{Binding IsExplicit}" Width="90"/>
              <DataGridTextColumn Header="Inherit" Binding="{Binding InheritanceFlags}" Width="100"/>
              <DataGridTextColumn Header="Propagate" Binding="{Binding PropagationFlags}" Width="100"/>
            </DataGrid.Columns>
          </DataGrid>
        </Grid>
      </TabItem>
      <TabItem Header="Drzewo">
        <Grid>
          <TreeView x:Name="tv" Margin="0,6,0,0"/>
        </Grid>
      </TabItem>
    </TabControl>

    <!-- Bottom: actions + log -->
    <Grid Grid.Row="2" Margin="0,10,0,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="2*"/>
      </Grid.ColumnDefinitions>
      <StackPanel Orientation="Horizontal" Grid.Column="0" VerticalAlignment="Top">
        <Button x:Name="btnExportCsv" Content="Eksport CSV" Width="120" Margin="0,0,10,0"/>
        <Button x:Name="btnExportXlsx" Content="Eksport XLSX" Width="120" Margin="0,0,10,0"/>
        <Button x:Name="btnClear" Content="Wyczyść wyniki" Width="140"/>
      </StackPanel>
      <RichTextBox x:Name="rbLog" Grid.Column="1" Margin="10,0,0,0" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
    </Grid>
  </Grid>
</Window>

"@
$window = [Windows.Markup.XamlReader]::Parse($Xaml)
$tbPath         = $window.FindName('tbPath')
$btnBrowse      = $window.FindName('btnBrowse')
$btnScan        = $window.FindName('btnScan')
$pb             = $window.FindName('pb')
$script:rbLog   = $window.FindName('rbLog')
$cbOnlyUsers    = $window.FindName('cbOnlyUsers')
$cbOnlyGroups   = $window.FindName('cbOnlyGroups')
$cbOnlyExplicit = $window.FindName('cbOnlyExplicit')
$tbSearch       = $window.FindName('tbSearch')
$btnToggleDebug = $window.FindName('btnToggleDebug')
$dg             = $window.FindName('dg')
$tv             = $window.FindName('tv')
$btnExportCsv   = $window.FindName('btnExportCsv')
$btnExportXlsx  = $window.FindName('btnExportXlsx')
$btnClear       = $window.FindName('btnClear')

# Data containers
$AllItems = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$View = [System.Windows.Data.CollectionViewSource]::GetDefaultView($AllItems)
$dg.ItemsSource = $View

function Set-GridFilter {
    $View.Filter = {
        param($row)
        $onlyUsers    = $cbOnlyUsers.IsChecked
        $onlyGroups   = $cbOnlyGroups.IsChecked
        $onlyExplicit = $cbOnlyExplicit.IsChecked
        $query        = ($tbSearch.Text ?? "").Trim()
        if ($onlyUsers -and $row.PrincipalType -ne 'User') { return $false }
        if ($onlyGroups -and $row.PrincipalType -ne 'Group') { return $false }
        if ($onlyExplicit -and -not $row.IsExplicit) { return $false }
        if ($query) {
            $hit = ($row.FolderPath -like "*${query}*") -or ($row.Identity -like "*${query}*")
            if (-not $hit) { return $false }
        }
        return $true
    }
    $View.Refresh()
}

function Set-UiBusyState {
    param([bool]$IsBusy)

    $btnScan.IsEnabled        = -not $IsBusy
    $btnBrowse.IsEnabled      = -not $IsBusy
    $btnExportCsv.IsEnabled   = -not $IsBusy
    $btnExportXlsx.IsEnabled  = -not $IsBusy
    $btnClear.IsEnabled       = -not $IsBusy
    $tbPath.IsEnabled         = -not $IsBusy
    $tbSearch.IsEnabled       = -not $IsBusy
    $cbOnlyUsers.IsEnabled    = -not $IsBusy
    $cbOnlyGroups.IsEnabled   = -not $IsBusy
    $cbOnlyExplicit.IsEnabled = -not $IsBusy
    $btnToggleDebug.IsEnabled = -not $IsBusy
}


# Browse folder
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Wybierz katalog do audytu'
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq 'OK') { $tbPath.Text = $dlg.SelectedPath }
})

# Scan (async; progress = indeterminate)
$btnScan.Add_Click({
    try {
        [string]$scanPath = $tbPath.Text
        if (-not $scanPath) {
            [System.Windows.MessageBox]::Show('Podaj sciezke do katalogu.', 'Blad', 'OK', 'Error') | Out-Null
            return
        }
        if (-not (Test-Path -LiteralPath $scanPath)) {
            [System.Windows.MessageBox]::Show('Podaj istniejaca sciezke.', 'Blad', 'OK', 'Error') | Out-Null
            return
        }
        try {
            $rootTest = Get-Item -LiteralPath $scanPath -ErrorAction Stop
            if (-not $rootTest.PSIsContainer) {
                [System.Windows.MessageBox]::Show('Sciezka nie jest katalogiem.', 'Blad', 'OK', 'Error') | Out-Null
                return
            }
        } catch {
            [System.Windows.MessageBox]::Show("Blad dostepu do sciezki: $($_.Exception.Message)", 'Blad', 'OK', 'Error') | Out-Null
            return
        }

        Set-UiBusyState $true
        $pb.IsIndeterminate = $false
        $pb.Value = 0
        $AllItems.Clear()
        $tv.Items.Clear()
        $PrincipalCache.Clear()
        Write-LogDebug 'Wyczyszczono poprzednie wyniki i cache.'
        Write-LogUI 'Start zadania...'
        Write-LogDebug ("UI thread {0} - przygotowanie skanowania dla: {1}" -f [System.Threading.Thread]::CurrentThread.ManagedThreadId, $scanPath)

        $opts = [pscustomobject]@{
            OnlyUsers        = [bool]$cbOnlyUsers.IsChecked
            OnlyGroups       = [bool]$cbOnlyGroups.IsChecked
            IncludeInherited = -not [bool]$cbOnlyExplicit.IsChecked
        }

        Write-LogDebug ("Task start (thread {0})" -f [System.Threading.Thread]::CurrentThread.ManagedThreadId)
        $result = Get-NTFSAclReport -RootPath $scanPath -Options $opts
        $itemsCount = if ($null -ne $result) { $result.Count } else { 0 }
        Write-LogDebug ("Get-NTFSAclReport zwrocil {0} elementow" -f $itemsCount)

        if ($result) {
            foreach ($r in $result) { $AllItems.Add($r) }
        }

        Set-GridFilter
        Build-TreeView
        $pb.Value = 100
        Write-LogUI 'Skanowanie zakonczone.'

    } catch {
        Write-LogUI ("Blad skanowania: {0}" -f $_.Exception.Message) 'ERROR'
    } finally {
        Set-UiBusyState $false
        $pb.IsIndeterminate = $false
        if (-not $AllItems.Count) { $pb.Value = 0 }
    }
})
$btnToggleDebug.Add_Click({
    $script:IsDebugLogging = -not $script:IsDebugLogging
    $btnToggleDebug.Content = if ($script:IsDebugLogging) { 'Debug log ON' } else { 'Debug log OFF' }
    $state = if ($script:IsDebugLogging) { 'wlaczony' } else { 'wylaczony' }
    Write-LogUI "Tryb logowania debug: $state"
})

# Filters events
$cbOnlyUsers.Add_Click({ Set-GridFilter })
$cbOnlyGroups.Add_Click({ Set-GridFilter })
$cbOnlyExplicit.Add_Click({ Set-GridFilter })
$tbSearch.Add_TextChanged({ Set-GridFilter })

$btnClear.Add_Click({
    $AllItems.Clear(); $tv.Items.Clear(); $pb.Value = 0
    Write-LogUI 'Wyczyszczono wyniki.'
})

$btnExportCsv.Add_Click({
    if ($AllItems.Count -gt 0) { Export-ReportCSV -Data $View }
    else { [System.Windows.MessageBox]::Show('Brak danych.', 'Eksport', 'OK', 'Information') | Out-Null }
})
$btnExportXlsx.Add_Click({
    if ($AllItems.Count -gt 0) { Export-ReportXLSX -Data $View }
    else { [System.Windows.MessageBox]::Show('Brak danych.', 'Eksport', 'OK', 'Information') | Out-Null }
})


# Start window
[void]$window.ShowDialog()




