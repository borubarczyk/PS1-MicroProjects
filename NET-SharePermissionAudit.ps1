#requires -Version 5.1
<#
.SYNOPSIS
    Kompleksowy audyt uprawnień udziałów sieciowych (SMB): uprawnienia udziału + NTFS, wykrywanie ryzyk,
    efektywny dostęp użytkowników, raport HTML/CSV/JSON/XLSX.

.DESCRIPTION
    Skrypt wyszukuje udziały dyskowe na wskazanych serwerach (lub na serwerach z Active Directory) i odczytuje:
      * uprawnienia udziału (share permissions) przez NetShareEnum (poziom 502) - bez WMI i WinRM, wystarczy port 445,
      * właściwości udziału: Access-Based Enumeration, szyfrowanie SMB, tryb pamięci podręcznej (offline), DFS,
      * uprawnienia NTFS folderu głównego udziału i podfolderów do zadanej głębokości (kilka udziałów równolegle),
    a następnie:
      * rozwiązuje SID-y w kontekście serwera (konta domenowe, lokalne, wbudowane, osierocone SID-y),
      * sprawdza w AD, czy konta posiadające uprawnienia nie są wyłączone,
      * wykrywa typowe błędy konfiguracji: zapis dla Everyone/Domain Users, dostęp anonimowy, pełna kontrola dla
        zwykłych kont, uprawnienia nadane bezpośrednio użytkownikom, osierocone SID-y, wpisy Deny, wyłączone
        dziedziczenie, właściciel-użytkownik, brak dostępu SYSTEM/Administrators, udziały bez DACL,
      * opcjonalnie rozwija członkostwo grup (rekurencyjnie w AD oraz lokalne grupy serwera),
      * opcjonalnie liczy efektywny dostęp wskazanych użytkowników (udział ∩ NTFS, z uwzględnieniem Deny,
        kolejności wpisów ACL oraz członkostwa w grupach domenowych i lokalnych grupach serwera).

    Wynik trafia do folderu ShareAudit_<data>: interaktywny raport HTML (wyszukiwanie, filtry, sortowanie,
    eksport widoku), pliki CSV (separator ';'), opcjonalnie JSON i XLSX oraz log przebiegu.

    Skrypt wyłącznie ODCZYTUJE konfigurację - niczego nie zmienia na serwerach.

.PARAMETER ComputerName
    Serwery do audytu (nazwa NetBIOS, FQDN lub IP). Domyślnie komputer lokalny.

.PARAMETER FromAD
    Pobiera listę serwerów z Active Directory (włączone konta komputerów z systemem serwerowym).

.PARAMETER SearchBase
    Z -FromAD: DN jednostki organizacyjnej, np. 'OU=Serwery,DC=firma,DC=local'.

.PARAMETER ADComputerFilter
    Z -FromAD: filtr nazwy komputera z symbolem wieloznacznym *, np. 'FS*'. Domyślnie '*'.

.PARAMETER IncludeWorkstations
    Z -FromAD: uwzględnia także stacje robocze (domyślnie tylko systemy serwerowe).

.PARAMETER Path
    Konkretne lokalizacje do audytu: ścieżki UNC (\\serwer\udział\folder) lub lokalne (D:\Dane).
    Dla ścieżek UNC odczytywane są też uprawnienia udziału, w którym leży ścieżka.

.PARAMETER IncludeShare
    Nazwy udziałów do uwzględnienia (wildcard). Domyślnie '*'.

.PARAMETER ExcludeShare
    Nazwy udziałów do pominięcia (wildcard), np. 'Backup*','Skany$'.

.PARAMETER IncludeAdminShares
    Uwzględnia udziały administracyjne (C$, ADMIN$, print$ itp.), które domyślnie są pomijane.

.PARAMETER Depth
    Głębokość skanowania NTFS: 0 = tylko folder główny udziału, -1 = bez limitu. Domyślnie 2.

.PARAMETER IncludeInherited
    W tabeli NTFS pokazuje także wpisy dziedziczone w podfolderach. Domyślnie podfoldery pokazują tylko
    wpisy jawne (folder główny zawsze w całości).

.PARAMETER AllFolders
    Raportuje każdy przeskanowany folder, także bez jawnych uprawnień (duży raport).

.PARAMETER ShareOnly
    Tylko uprawnienia i właściwości udziałów, bez skanowania NTFS (szybki przegląd wielu serwerów).

.PARAMETER ScanMode
    Unc (domyślnie) - NTFS czytany przez ścieżkę UNC (wystarczy port 445).
    Remote - skanowanie wykonywane na serwerze przez WinRM (Invoke-Command), znacznie szybsze dla dużych drzew.
    Gdy WinRM jest niedostępny, skrypt automatycznie wraca do trybu UNC.

.PARAMETER Credential
    Poświadczenia do serwerów (połączenie IPC$ dla SMB/RPC oraz WinRM). Zapytania do AD używają bieżącego konta.

.PARAMETER ThrottleLimit
    Liczba lokalizacji skanowanych równolegle. Domyślnie 8.

.PARAMETER ExpandGroups
    Rozwija członkostwo grup z uprawnień: grupy domenowe rekurencyjnie (AD), grupy lokalne serwera bezpośrednio.
    Szerokie grupy (Everyone, Domain Users, Users...) nie są rozwijane.

.PARAMETER MaxGroupMembers
    Maksymalna liczba członków pobieranych dla jednej grupy. Domyślnie 500.

.PARAMETER UserAccess
    Konta (sAMAccountName, DOMENA\login lub UPN), dla których policzyć efektywny dostęp do audytowanych folderów.

.PARAMETER TrustedPrincipal
    Dodatkowe konta/grupy administracyjne (nazwa z wildcard lub SID), które nie mają być zgłaszane jako
    "pełna kontrola dla zwykłego konta", np. 'FIRMA\ADM-Pliki'.

.PARAMETER SkipADLookup
    Nie odpytuje AD (stan kont, członkostwo grup, -UserAccess) - np. na komputerze spoza domeny.

.PARAMETER OutputPath
    Folder, w którym powstanie katalog raportu ShareAudit_<data>. Domyślnie Pulpit.

.PARAMETER Format
    Formaty raportu: Html, Csv, Json, Xlsx (Xlsx wymaga modułu ImportExcel). Domyślnie Html i Csv.

.PARAMETER CsvDelimiter
    Separator pól CSV. Domyślnie ';' (polski Excel).

.PARAMETER HtmlRowLimit
    Maksymalna liczba wierszy jednej tabeli w raporcie HTML (pełne dane są zawsze w CSV/JSON). Domyślnie 20000.

.PARAMETER NoOpen
    Nie otwiera raportu HTML po zakończeniu audytu.

.PARAMETER PassThru
    Zwraca obiekt z wynikami do dalszej obróbki w PowerShell.

.EXAMPLE
    .\NET-SharePermissionAudit.ps1
    Audyt udziałów komputera lokalnego (NTFS do głębokości 2), raport na Pulpicie.

.EXAMPLE
    .\NET-SharePermissionAudit.ps1 -ComputerName FS01, FS02 -Depth 4 -ExpandGroups
    Audyt dwóch serwerów plików z rozwinięciem członkostwa grup.

.EXAMPLE
    .\NET-SharePermissionAudit.ps1 -FromAD -SearchBase 'OU=Serwery,DC=firma,DC=local' -ShareOnly -Format Html, Csv, Xlsx
    Szybki przegląd uprawnień udziałów na wszystkich serwerach z OU (bez skanowania NTFS).

.EXAMPLE
    .\NET-SharePermissionAudit.ps1 -Path '\\FS01\Dzialy\Kadry', '\\FS01\Projekty' -Depth -1 -UserAccess jkowalski, anowak
    Pełne drzewo wskazanych lokalizacji i mapa efektywnego dostępu dwóch użytkowników.

.EXAMPLE
    .\NET-SharePermissionAudit.ps1 -ComputerName FS01 -ScanMode Remote -Credential (Get-Credential) -Depth -1
    Skanowanie całych drzew po stronie serwera (WinRM) z innymi poświadczeniami.

.EXAMPLE
    $audit = .\NET-SharePermissionAudit.ps1 -ComputerName FS01 -NoOpen -PassThru
    $audit.Findings | Where-Object SeverityRank -ge 4 | Format-Table Severity, Title, Path, Principal

.NOTES
    Wymagania: Windows, PowerShell 5.1 lub 7. Odczyt uprawnień udziału wymaga uprawnień administratora na serwerze
    (bez nich skrypt i tak wyliczy udziały i przeskanuje NTFS tam, gdzie ma dostęp). Rozwijanie grup, stan kont
    i -UserAccess wymagają komputera w domenie.
    Ograniczenia analizy efektywnego dostępu: nie uwzględnia niejawnych praw właściciela, Dynamic Access Control
    ani uprawnień nadanych na pojedynczych plikach (analizowane są foldery).
#>
[CmdletBinding(DefaultParameterSetName = 'Computer')]
param(
    [Parameter(ParameterSetName = 'Computer', Position = 0)]
    [Alias('Server', 'CN')]
    [ValidateNotNullOrEmpty()]
    [string[]]$ComputerName = @($env:COMPUTERNAME),

    [Parameter(ParameterSetName = 'AD', Mandatory = $true)]
    [switch]$FromAD,

    [Parameter(ParameterSetName = 'AD')]
    [string]$SearchBase,

    [Parameter(ParameterSetName = 'AD')]
    [string]$ADComputerFilter = '*',

    [Parameter(ParameterSetName = 'AD')]
    [switch]$IncludeWorkstations,

    [Parameter(ParameterSetName = 'Path', Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [string[]]$IncludeShare = @('*'),
    [string[]]$ExcludeShare = @(),
    [switch]$IncludeAdminShares,

    [ValidateRange(-1, 1000)]
    [int]$Depth = 2,
    [switch]$IncludeInherited,
    [switch]$AllFolders,
    [switch]$ShareOnly,

    [ValidateSet('Unc', 'Remote')]
    [string]$ScanMode = 'Unc',
    [System.Management.Automation.PSCredential]$Credential,
    [ValidateRange(1, 64)]
    [int]$ThrottleLimit = 8,

    [switch]$ExpandGroups,
    [ValidateRange(1, 100000)]
    [int]$MaxGroupMembers = 500,
    [string[]]$UserAccess,
    [string[]]$TrustedPrincipal = @(),
    [switch]$SkipADLookup,

    [string]$OutputPath,
    [ValidateSet('Html', 'Csv', 'Json', 'Xlsx')]
    [string[]]$Format = @('Html', 'Csv'),
    [ValidateNotNullOrEmpty()]
    [string]$CsvDelimiter = ';',
    [ValidateRange(100, 1000000)]
    [int]$HtmlRowLimit = 20000,
    [switch]$NoOpen,
    [switch]$PassThru
)

if ($PSVersionTable.PSVersion.Major -ge 6 -and -not $IsWindows) {
    throw 'Ten skrypt działa wyłącznie w systemie Windows.'
}

#region Typy natywne (netapi32 / advapi32 / mpr)
if (-not ('ShareAudit.Native' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace ShareAudit
{
    public class ShareEntry
    {
        public string Name;
        public uint Type;
        public string Remark;
        public string Path;
        public uint CurrentUses;
        public int InfoLevel;
        public byte[] SecurityDescriptor;

        public uint BaseType { get { return Type & 0xFF; } }
        public bool IsSpecial { get { return (Type & 0x80000000) != 0; } }
        public bool IsTemporary { get { return (Type & 0x40000000) != 0; } }
    }

    public class SidLookupResult
    {
        public bool Success;
        public string Name;
        public string Domain;
        public int Use;
        public int Error;
    }

    public static class Native
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHARE_INFO_1
        {
            public string shi1_netname;
            public uint shi1_type;
            public string shi1_remark;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHARE_INFO_502
        {
            public string shi502_netname;
            public uint shi502_type;
            public string shi502_remark;
            public uint shi502_permissions;
            public uint shi502_max_uses;
            public uint shi502_current_uses;
            public string shi502_path;
            public string shi502_passwd;
            public uint shi502_reserved;
            public IntPtr shi502_security_descriptor;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private class NETRESOURCE
        {
            public int dwScope;
            public int dwType;
            public int dwDisplayType;
            public int dwUsage;
            public string lpLocalName;
            public string lpRemoteName;
            public string lpComment;
            public string lpProvider;
        }

        [DllImport("netapi32.dll", CharSet = CharSet.Unicode)]
        private static extern int NetShareEnum(string serverName, int level, out IntPtr bufPtr, int prefMaxLen, out int entriesRead, out int totalEntries, ref int resumeHandle);

        [DllImport("netapi32.dll", CharSet = CharSet.Unicode)]
        private static extern int NetShareGetInfo(string serverName, string netName, int level, out IntPtr bufPtr);

        [DllImport("netapi32.dll")]
        private static extern int NetApiBufferFree(IntPtr buffer);

        [DllImport("advapi32.dll")]
        private static extern bool IsValidSecurityDescriptor(IntPtr sd);

        [DllImport("advapi32.dll")]
        private static extern uint GetSecurityDescriptorLength(IntPtr sd);

        [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool LookupAccountSid(string systemName, byte[] sid, StringBuilder name, ref uint cchName, StringBuilder domain, ref uint cchDomain, out int use);

        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetAddConnection2(NETRESOURCE netResource, string password, string userName, int flags);

        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetCancelConnection2(string name, int flags, bool force);

        private const int ERROR_ACCESS_DENIED = 5;
        private const int ERROR_INSUFFICIENT_BUFFER = 122;
        private const int ERROR_MORE_DATA = 234;
        private const int MAX_PREFERRED_LENGTH = -1;

        // PowerShell przekazuje $null do parametru string jako "", a API oczekuje NULL dla komputera lokalnego.
        private static string NullIfEmpty(string value)
        {
            return string.IsNullOrEmpty(value) ? null : value;
        }

        public static int EnumShares(string server, out ShareEntry[] shares, out int level)
        {
            List<ShareEntry> list = new List<ShareEntry>();
            level = 502;
            int rc = EnumSharesAtLevel(NullIfEmpty(server), 502, list);
            if (rc == ERROR_ACCESS_DENIED)
            {
                // Poziom 502 (z deskryptorem zabezpieczeń) wymaga administratora - poziom 1 nie.
                list.Clear();
                level = 1;
                rc = EnumSharesAtLevel(NullIfEmpty(server), 1, list);
            }
            shares = list.ToArray();
            return rc;
        }

        private static int EnumSharesAtLevel(string server, int level, List<ShareEntry> list)
        {
            int resume = 0;
            int rc;
            do
            {
                IntPtr buffer = IntPtr.Zero;
                int read;
                int total;
                rc = NetShareEnum(server, level, out buffer, MAX_PREFERRED_LENGTH, out read, out total, ref resume);
                try
                {
                    if (rc != 0 && rc != ERROR_MORE_DATA) { break; }
                    int size = level == 502 ? Marshal.SizeOf(typeof(SHARE_INFO_502)) : Marshal.SizeOf(typeof(SHARE_INFO_1));
                    for (int i = 0; i < read; i++)
                    {
                        IntPtr item = new IntPtr(buffer.ToInt64() + (long)i * size);
                        ShareEntry entry = new ShareEntry();
                        entry.InfoLevel = level;
                        if (level == 502)
                        {
                            SHARE_INFO_502 si = (SHARE_INFO_502)Marshal.PtrToStructure(item, typeof(SHARE_INFO_502));
                            entry.Name = si.shi502_netname;
                            entry.Type = si.shi502_type;
                            entry.Remark = si.shi502_remark;
                            entry.Path = si.shi502_path;
                            entry.CurrentUses = si.shi502_current_uses;
                            if (si.shi502_security_descriptor != IntPtr.Zero && IsValidSecurityDescriptor(si.shi502_security_descriptor))
                            {
                                uint length = GetSecurityDescriptorLength(si.shi502_security_descriptor);
                                entry.SecurityDescriptor = new byte[length];
                                Marshal.Copy(si.shi502_security_descriptor, entry.SecurityDescriptor, 0, (int)length);
                            }
                        }
                        else
                        {
                            SHARE_INFO_1 si = (SHARE_INFO_1)Marshal.PtrToStructure(item, typeof(SHARE_INFO_1));
                            entry.Name = si.shi1_netname;
                            entry.Type = si.shi1_type;
                            entry.Remark = si.shi1_remark;
                        }
                        list.Add(entry);
                    }
                }
                finally
                {
                    if (buffer != IntPtr.Zero) { NetApiBufferFree(buffer); }
                }
            } while (rc == ERROR_MORE_DATA);
            return rc;
        }

        // SHARE_INFO_1005 - flagi udziału (ABE, szyfrowanie, cache, DFS). Nie wymaga uprawnień administratora.
        public static int GetShareFlags(string server, string share, out uint flags)
        {
            flags = 0;
            IntPtr buffer = IntPtr.Zero;
            int rc = NetShareGetInfo(NullIfEmpty(server), share, 1005, out buffer);
            try
            {
                if (rc == 0 && buffer != IntPtr.Zero) { flags = (uint)Marshal.ReadInt32(buffer); }
            }
            finally
            {
                if (buffer != IntPtr.Zero) { NetApiBufferFree(buffer); }
            }
            return rc;
        }

        public static SidLookupResult LookupSid(string systemName, byte[] sid)
        {
            SidLookupResult result = new SidLookupResult();
            string system = NullIfEmpty(systemName);
            StringBuilder name = new StringBuilder(256);
            StringBuilder domain = new StringBuilder(256);
            uint cchName = (uint)name.Capacity;
            uint cchDomain = (uint)domain.Capacity;
            int use;
            bool ok = LookupAccountSid(system, sid, name, ref cchName, domain, ref cchDomain, out use);
            int error = ok ? 0 : Marshal.GetLastWin32Error();
            if (!ok && error == ERROR_INSUFFICIENT_BUFFER)
            {
                name = new StringBuilder((int)cchName + 1);
                domain = new StringBuilder((int)cchDomain + 1);
                cchName = (uint)name.Capacity;
                cchDomain = (uint)domain.Capacity;
                ok = LookupAccountSid(system, sid, name, ref cchName, domain, ref cchDomain, out use);
                error = ok ? 0 : Marshal.GetLastWin32Error();
            }
            result.Success = ok;
            result.Error = error;
            if (ok)
            {
                result.Name = name.ToString();
                result.Domain = domain.ToString();
                result.Use = use;
            }
            return result;
        }

        public static int AddConnection(string remoteName, string userName, string password)
        {
            NETRESOURCE resource = new NETRESOURCE();
            resource.dwType = 0; // RESOURCETYPE_ANY
            resource.lpRemoteName = remoteName;
            return WNetAddConnection2(resource, password, userName, 0);
        }

        public static int CancelConnection(string remoteName)
        {
            return WNetCancelConnection2(remoteName, 0, true);
        }
    }
}
'@
}
#endregion

#region Stan i stałe
$script:ScriptVersion   = '1.0'
$script:StartTime       = Get-Date
$script:LogLines        = New-Object System.Collections.Generic.List[string]
$script:ErrorRows       = New-Object System.Collections.Generic.List[object]
$script:Findings        = New-Object System.Collections.Generic.List[object]
$script:FindingKeys     = New-Object 'System.Collections.Generic.HashSet[string]'
$script:Shares          = New-Object System.Collections.Generic.List[object]
$script:Folders         = New-Object System.Collections.Generic.List[object]
$script:ServerRows      = New-Object System.Collections.Generic.List[object]
$script:IpcConnections  = New-Object System.Collections.Generic.List[string]
$script:PrincipalCache  = @{}
$script:ADEnabledCache  = @{}
$script:ServerCache     = @{}
$script:LocalGroupCache = @{}
$script:UserSidSetCache = @{}
$script:UsedGroups      = @{}
$script:ADAvailable     = $false

$script:SeverityRank = [ordered]@{ 'Krytyczne' = 5; 'Wysokie' = 4; 'Średnie' = 3; 'Niskie' = 2; 'Info' = 1 }
$script:SeverityColor = @{ 'Krytyczne' = 'Magenta'; 'Wysokie' = 'Red'; 'Średnie' = 'Yellow'; 'Niskie' = 'Cyan'; 'Info' = 'Gray' }
$script:AccessLabels = @{ 4 = 'Pełna kontrola'; 3 = 'Modyfikacja'; 2 = 'Zapis'; 1 = 'Odczyt'; 0 = 'Specjalne' }
$script:AclStatusText = @{
    OK           = 'OK'
    NullDacl     = 'Brak DACL - pełny dostęp dla wszystkich'
    EmptyDacl    = 'Pusta DACL - brak dostępu'
    AccessDenied = 'Brak uprawnień do odczytu (wymagany administrator)'
    Default      = 'Domyślne (udział administracyjny)'
    NotShare     = 'Ścieżka lokalna (bez udziału)'
    Unknown      = 'Nie znaleziono udziału (np. ścieżka DFS)'
    Error        = 'Błąd odczytu deskryptora'
}
$script:AdminShareNames = @('ADMIN$', 'IPC$', 'PRINT$', 'FAX$')

# Szerokie grupy: Everyone, Anonymous, Authenticated Users, NETWORK, INTERACTIVE, This Organization, Users, Guests,
# Guest, Domain Users, Domain Guests, Domain Computers
$script:BroadSidRegex = '^(S-1-1-0|S-1-5-7|S-1-5-11|S-1-5-2|S-1-5-4|S-1-5-15|S-1-5-32-545|S-1-5-32-546|S-1-5-21-\d+-\d+-\d+-(501|513|514|515))$'
# Dostęp anonimowy / gościa
$script:AnonSidRegex = '^(S-1-5-7|S-1-5-32-546|S-1-5-21-\d+-\d+-\d+-(501|514))$'
# SYSTEM, Administrators, Server/Backup Operators, Enterprise DCs, TrustedInstaller, Administrator, Domain Admins,
# Domain Controllers, Schema/Enterprise Admins, Enterprise RODCs
$script:AdminSidRegex = '^(S-1-5-18|S-1-5-32-544|S-1-5-32-549|S-1-5-32-551|S-1-5-9|S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464|S-1-5-21-\d+-\d+-\d+-(500|512|516|518|519|498))$'
# CREATOR OWNER, CREATOR GROUP, OWNER RIGHTS
$script:CreatorSidRegex = '^(S-1-3-0|S-1-3-1|S-1-3-4)$'

$script:LocalNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($n in @('.', 'localhost', '127.0.0.1', '::1', $env:COMPUTERNAME)) { if ($n) { [void]$script:LocalNames.Add($n) } }
try { [void]$script:LocalNames.Add([System.Net.Dns]::GetHostEntry([string]::Empty).HostName) } catch { }
if ($env:USERDNSDOMAIN) { [void]$script:LocalNames.Add(('{0}.{1}' -f $env:COMPUTERNAME, $env:USERDNSDOMAIN)) }
#endregion

#region Funkcje pomocnicze
function Write-AuditLog {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP')]
        [string]$Level = 'INFO'
    )
    $colors = @{ INFO = 'Gray'; OK = 'Green'; WARN = 'Yellow'; ERROR = 'Red'; STEP = 'Cyan' }
    $line = '[{0:HH:mm:ss}] {1,-5} {2}' -f (Get-Date), $Level, $Message
    Write-Host $line -ForegroundColor $colors[$Level]
    $script:LogLines.Add($line)
}

function Add-AuditError {
    param(
        [string]$Server,
        [string]$Share = '',
        [string]$Path = '',
        [string]$Stage,
        [string]$Message,
        [switch]$Quiet
    )
    $script:ErrorRows.Add([pscustomobject]@{
        Server  = $Server
        Share   = $Share
        Path    = $Path
        Stage   = $Stage
        Message = $Message
    })
    $text = '{0} {1} [{2}] {3}' -f $Server, $Path, $Stage, $Message
    if ($Quiet) {
        Write-Verbose $text
        $script:LogLines.Add(('[{0:HH:mm:ss}] WARN  {1}' -f (Get-Date), $text))
    } else {
        Write-AuditLog $text 'WARN'
    }
}

function Test-IsLocalComputer {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $true }
    return $script:LocalNames.Contains($Name.Trim().TrimEnd('.'))
}

function Get-Win32ErrorText {
    param([int]$Code)
    $known = @{
        5    = 'Odmowa dostępu'
        53   = 'Nie znaleziono ścieżki sieciowej'
        67   = 'Nie można odnaleźć nazwy sieciowej'
        1219 = 'Istnieje już połączenie z serwerem z innymi poświadczeniami'
        1326 = 'Nieprawidłowa nazwa użytkownika lub hasło'
        1722 = 'Serwer RPC jest niedostępny'
        2114 = 'Usługa Serwer (LanmanServer) nie jest uruchomiona'
        2310 = 'Udział nie istnieje'
    }
    if ($known.ContainsKey($Code)) { return '{0} (kod {1})' -f $known[$Code], $Code }
    return '{0} (kod {1})' -f ([System.ComponentModel.Win32Exception]::new($Code)).Message, $Code
}

function ConvertTo-LdapFilterValue {
    param([string]$Value, [switch]$AllowWildcard)
    $escaped = $Value -replace '\\', '\5c' -replace '\(', '\28' -replace '\)', '\29' -replace "`0", '\00'
    if (-not $AllowWildcard) { $escaped = $escaped -replace '\*', '\2a' }
    return $escaped
}

function Test-TcpPortBulk {
    param([string[]]$HostNames, [int]$Port = 445, [int]$TimeoutMs = 2000)
    $result = @{}
    $batchSize = 64
    for ($i = 0; $i -lt $HostNames.Count; $i += $batchSize) {
        $last = [Math]::Min($i + $batchSize, $HostNames.Count) - 1
        $probes = foreach ($h in $HostNames[$i..$last]) {
            $client = New-Object System.Net.Sockets.TcpClient
            $task = $null
            try { $task = $client.ConnectAsync($h, $Port) } catch { }
            [pscustomobject]@{ Host = $h; Client = $client; Task = $task }
        }
        $tasks = [System.Threading.Tasks.Task[]]@($probes | Where-Object { $_.Task } | ForEach-Object { $_.Task })
        if ($tasks.Count -gt 0) {
            try { [void][System.Threading.Tasks.Task]::WaitAll($tasks, $TimeoutMs) } catch { }
        }
        foreach ($probe in $probes) {
            $ok = $false
            if ($probe.Task -and $probe.Task.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion) { $ok = $probe.Client.Connected }
            $result[$probe.Host] = [bool]$ok
            try { $probe.Client.Close() } catch { }
        }
    }
    return $result
}

function Get-ADServerList {
    $searcher = New-Object System.DirectoryServices.DirectorySearcher
    if ($SearchBase) { $searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry ('LDAP://{0}' -f $SearchBase) }
    $osFilter = if ($IncludeWorkstations) { '' } else { '(operatingSystem=*Server*)' }
    $nameFilter = ConvertTo-LdapFilterValue -Value $ADComputerFilter -AllowWildcard
    $searcher.Filter = '(&(objectCategory=computer){0}(!(userAccountControl:1.2.840.113556.1.4.803:=2))(name={1}))' -f $osFilter, $nameFilter
    $searcher.PageSize = 1000
    foreach ($prop in 'name', 'dnshostname') { [void]$searcher.PropertiesToLoad.Add($prop) }
    $found = $searcher.FindAll()
    try {
        foreach ($r in $found) {
            if ($r.Properties['dnshostname'].Count -gt 0) { [string]$r.Properties['dnshostname'][0] }
            else { [string]$r.Properties['name'][0] }
        }
    } finally {
        $found.Dispose()
    }
}

function Connect-AuditIpc {
    param([string]$Server)
    if (-not $Credential -or (Test-IsLocalComputer $Server)) { return }
    $remote = '\\{0}\IPC$' -f $Server
    if ($script:IpcConnections.Contains($remote)) { return }
    $rc = [ShareAudit.Native]::AddConnection($remote, $Credential.UserName, $Credential.GetNetworkCredential().Password)
    if ($rc -eq 0) {
        $script:IpcConnections.Add($remote)
        Write-AuditLog ('{0}: połączono jako {1}' -f $Server, $Credential.UserName) 'INFO'
    } elseif ($rc -eq 1219) {
        Write-AuditLog ('{0}: istnieje już połączenie z innymi poświadczeniami - zostanie użyte istniejące.' -f $Server) 'WARN'
    } else {
        Write-AuditLog ('{0}: nie udało się połączyć z IPC$ - {1}' -f $Server, (Get-Win32ErrorText $rc)) 'WARN'
    }
}

function Disconnect-AuditIpc {
    foreach ($remote in $script:IpcConnections) {
        try { [void][ShareAudit.Native]::CancelConnection($remote) } catch { }
    }
    $script:IpcConnections.Clear()
}
#endregion

#region Maski uprawnień
function ConvertTo-UInt32Mask {
    param([int64]$Value)
    if ($Value -lt 0) { $Value += 4294967296 }
    return $Value
}

# Zamienia prawa ogólne (GENERIC_*) na konkretne prawa plikowe i obcina bity spoza FullControl.
function ConvertTo-NormalizedMask {
    param([int64]$Mask)
    $m = $Mask
    if ($m -band 0x10000000) { $m = $m -bor 0x1F01FF }  # GENERIC_ALL
    if ($m -band 2147483648) { $m = $m -bor 0x120089 }  # GENERIC_READ
    if ($m -band 0x40000000) { $m = $m -bor 0x120116 }  # GENERIC_WRITE
    if ($m -band 0x20000000) { $m = $m -bor 0x1200A0 }  # GENERIC_EXECUTE
    return ($m -band 0x1F01FF)
}

# 4 = pełna kontrola lub prawo zmiany uprawnień/właściciela, 3 = modyfikacja, 2 = zapis, 1 = odczyt, 0 = specjalne
function Get-AccessRank {
    param([int64]$Mask)
    $m = ConvertTo-NormalizedMask $Mask
    if (($m -band 0x1F01FF) -eq 0x1F01FF) { return 4 }
    if ($m -band 0xC0000) { return 4 }                    # WRITE_DAC / WRITE_OWNER
    if (($m -band 0x1301BF) -eq 0x1301BF) { return 3 }
    if ($m -band 0x10046) { return 2 }                    # WriteData, AppendData, DeleteChild, Delete
    if ($m -band 0x1) { return 1 }                        # ReadData / ListDirectory
    return 0
}

function Get-AccessLabel {
    param([int64]$Mask)
    $m = ConvertTo-NormalizedMask $Mask
    if ($m -eq 0) { return 'Brak' }
    if ((($m -band 0x1F01FF) -ne 0x1F01FF) -and ($m -band 0xC0000)) { return 'Zmiana uprawnień/właściciela' }
    return $script:AccessLabels[(Get-AccessRank $Mask)]
}

function Get-RightsText {
    param([int64]$Mask)
    $parts = New-Object System.Collections.Generic.List[string]
    if ($Mask -band 0x10000000) { $parts.Add('GENERIC_ALL') }
    if ($Mask -band 2147483648) { $parts.Add('GENERIC_READ') }
    if ($Mask -band 0x40000000) { $parts.Add('GENERIC_WRITE') }
    if ($Mask -band 0x20000000) { $parts.Add('GENERIC_EXECUTE') }
    $specific = [int]($Mask -band 0x0FFFFFFF)
    if ($specific -ne 0) {
        $text = ([Enum]::ToObject([System.Security.AccessControl.FileSystemRights], $specific)).ToString()
        if ($text -ne 'Synchronize') { $text = $text -replace ',\s*Synchronize', '' }
        $parts.Add($text)
    }
    if ($parts.Count -eq 0) { return '0x0' }
    return ($parts -join ', ')
}

function Get-AppliesToText {
    param([int]$InheritanceFlags, [int]$PropagationFlags)
    $ci = ($InheritanceFlags -band 1) -ne 0
    $oi = ($InheritanceFlags -band 2) -ne 0
    $np = ($PropagationFlags -band 1) -ne 0
    $io = ($PropagationFlags -band 2) -ne 0
    $text = if (-not $ci -and -not $oi) { 'Tylko ten folder' }
    elseif ($io) {
        if ($ci -and $oi) { 'Tylko podfoldery i pliki' } elseif ($ci) { 'Tylko podfoldery' } else { 'Tylko pliki' }
    } else {
        if ($ci -and $oi) { 'Ten folder, podfoldery i pliki' } elseif ($ci) { 'Ten folder i podfoldery' } else { 'Ten folder i pliki' }
    }
    if ($np -and ($ci -or $oi)) { $text += ' (tylko 1 poziom)' }
    return $text
}

# Efektywna maska dla zbioru SID-ów - wpisy oceniane w kolejności ACL (jak AccessCheck: pierwszy pasujący wpis wygrywa).
function Get-EffectiveMask {
    param([object[]]$Aces, $SidSet, [switch]$FolderOnly)
    [int64]$granted = 0
    [int64]$denied = 0
    foreach ($ace in $Aces) {
        if ($null -eq $ace) { continue }
        if ($FolderOnly -and (([int]$ace.PF -band 2) -ne 0)) { continue }   # InheritOnly - nie dotyczy samego folderu
        if (-not $SidSet.Contains([string]$ace.Sid)) { continue }
        $m = ConvertTo-NormalizedMask ([int64]$ace.Mask)
        if ($ace.Type -eq 'Deny') {
            $denied = $denied -bor ($m -band (-bnot $granted))
        } else {
            $granted = $granted -bor ($m -band (-bnot $denied))
        }
    }
    return $granted
}

function ConvertFrom-ShareSecurityDescriptor {
    param([byte[]]$Bytes)
    $sd = [System.Security.AccessControl.RawSecurityDescriptor]::new($Bytes, 0)
    if (-not $sd.ControlFlags.HasFlag([System.Security.AccessControl.ControlFlags]::DiscretionaryAclPresent) -or $null -eq $sd.DiscretionaryAcl) {
        return [pscustomobject]@{ Status = 'NullDacl'; Aces = @() }
    }
    $aces = New-Object System.Collections.Generic.List[object]
    foreach ($ace in $sd.DiscretionaryAcl) {
        if ($ace -isnot [System.Security.AccessControl.CommonAce]) { continue }
        $type = switch ([string]$ace.AceQualifier) {
            'AccessAllowed' { 'Allow' }
            'AccessDenied'  { 'Deny' }
            default         { $null }
        }
        if (-not $type) { continue }
        $aces.Add([pscustomobject]@{
            Sid       = $ace.SecurityIdentifier.Value
            Mask      = ConvertTo-UInt32Mask $ace.AccessMask
            Type      = $type
            Inherited = $false
            IF        = 0
            PF        = 0
        })
    }
    $status = if ($aces.Count -eq 0) { 'EmptyDacl' } else { 'OK' }
    return [pscustomobject]@{ Status = $status; Aces = $aces.ToArray() }
}
#endregion

#region Podmioty zabezpieczeń (SID)
function Get-ADAccountEnabled {
    param([string]$Sid)
    if (-not $script:ADAvailable) { return $null }
    if ($script:ADEnabledCache.ContainsKey($Sid)) { return $script:ADEnabledCache[$Sid] }
    $enabled = $null
    try {
        $entry = [ADSI]('LDAP://<SID={0}>' -f $Sid)
        $uac = $entry.Properties['userAccountControl'].Value
        if ($null -ne $uac) { $enabled = -not ([int]$uac -band 2) }
    } catch { }
    $script:ADEnabledCache[$Sid] = $enabled
    return $enabled
}

function Resolve-AuditPrincipal {
    param([string]$Server, [string]$Sid)
    if ([string]::IsNullOrEmpty($Sid)) { $Sid = 'S-1-0-0' }
    $globalKey = '*|' + $Sid
    if ($script:PrincipalCache.ContainsKey($globalKey)) { return $script:PrincipalCache[$globalKey] }
    $serverKey = $Server + '|' + $Sid
    if ($script:PrincipalCache.ContainsKey($serverKey)) { return $script:PrincipalCache[$serverKey] }

    $isLocal = Test-IsLocalComputer $Server
    $account = $null
    $domain = $null
    $use = 0
    $lookupError = 0
    $sidObject = $null
    try { $sidObject = [System.Security.Principal.SecurityIdentifier]::new($Sid) } catch { }
    if ($sidObject) {
        $bytes = New-Object byte[] ($sidObject.BinaryLength)
        $sidObject.GetBinaryForm($bytes, 0)
        # Rozwiązywanie w kontekście serwera - tylko tak poprawnie rozpoznamy jego konta i grupy lokalne.
        $lookup = [ShareAudit.Native]::LookupSid($(if ($isLocal) { '' } else { $Server }), $bytes)
        if (-not $lookup.Success -and -not $isLocal) {
            $fallback = [ShareAudit.Native]::LookupSid('', $bytes)
            if ($fallback.Success) { $lookup = $fallback }
        }
        if ($lookup.Success) {
            $account = $lookup.Name
            $domain = $lookup.Domain
            $use = $lookup.Use
        } else {
            $lookupError = $lookup.Error
        }
    }

    $isAccountSid = $Sid -like 'S-1-5-21-*'
    $type = 'Unknown'
    if ($account) {
        switch ($use) {
            1       { $type = if ($account.EndsWith('$')) { 'Computer' } else { 'User' } }
            2       { $type = 'Group' }
            4       { $type = 'Group' }
            5       { $type = 'WellKnown' }
            6       { $type = 'Orphaned' }
            9       { $type = 'Computer' }
            default { $type = 'Unknown' }
        }
    } elseif ($lookupError -eq 1332 -and $isAccountSid) {
        $type = 'Orphaned'   # ERROR_NONE_MAPPED - konto usunięte
    }

    $shortServer = if ($isLocal) { $env:COMPUTERNAME } else { ($Server -split '\.')[0] }
    $scope = if ($Sid -like 'S-1-5-32-*') { 'Builtin' }
    elseif (-not $isAccountSid) { 'WellKnown' }
    elseif ($domain -and $domain -ieq $shortServer) { 'Local' }
    elseif ($domain) { 'Domain' }
    else { 'Unknown' }

    $name = if ($account) { if ($domain) { '{0}\{1}' -f $domain, $account } else { $account } } else { $Sid }

    $isTrusted = $false
    foreach ($pattern in $TrustedPrincipal) {
        if (-not $pattern) { continue }
        if ($Sid -ieq $pattern -or $name -like $pattern -or ($account -and $account -like $pattern)) { $isTrusted = $true; break }
    }

    $enabled = $null
    if (($type -eq 'User' -or $type -eq 'Computer') -and $scope -eq 'Domain') { $enabled = Get-ADAccountEnabled $Sid }

    $principal = [pscustomobject]@{
        Sid            = $Sid
        Name           = $name
        Account        = $account
        Domain         = $domain
        Type           = $type
        Scope          = $scope
        Enabled        = $enabled
        IsBroad        = $Sid -match $script:BroadSidRegex
        IsAnonymous    = $Sid -match $script:AnonSidRegex
        IsAdmin        = $Sid -match $script:AdminSidRegex
        IsCreatorOwner = $Sid -match $script:CreatorSidRegex
        IsTrusted      = $isTrusted
    }
    # Konta domenowe i dobrze znane SID-y są takie same na każdym serwerze - lokalne zależą od serwera.
    $cacheKey = if ($scope -eq 'WellKnown' -or ($scope -eq 'Domain' -and $account)) { $globalKey } else { $serverKey }
    $script:PrincipalCache[$cacheKey] = $principal
    return $principal
}

function Register-UsedGroup {
    param([string]$Server, $Principal)
    if ($Principal.Type -ne 'Group') { return }
    if ($Principal.Scope -eq 'Domain') { $key = 'D|' + $Principal.Sid }
    elseif ($Principal.Scope -eq 'Local' -or $Principal.Scope -eq 'Builtin') { $key = $Server + '|' + $Principal.Sid }
    else { return }
    if (-not $script:UsedGroups.ContainsKey($key)) {
        $script:UsedGroups[$key] = [pscustomobject]@{ Server = $Server; Principal = $Principal }
    }
}
#endregion

#region Udziały
function New-AuditShareObject {
    param(
        [string]$Server,
        [string]$Name,
        [string]$UncPath,
        [string]$LocalPath,
        [string]$Description = '',
        [string]$AclStatus,
        [object[]]$Aces = @(),
        [bool]$IsAdminShare = $false,
        [bool]$IsTemporary = $false,
        [int]$CurrentUses = 0
    )
    [pscustomobject]@{
        Server          = $Server
        Name            = $Name
        UncPath         = $UncPath
        LocalPath       = $LocalPath
        Description     = $Description
        IsAdminShare    = $IsAdminShare
        IsHidden        = $Name.EndsWith('$')
        IsTemporary     = $IsTemporary
        CurrentUses     = $CurrentUses
        AclStatus       = $AclStatus
        Aces            = @($Aces)
        Flags           = $null
        FlagsRead       = $false
        BroadRank       = $null
        FoldersScanned  = 0
        FoldersReported = 0
        NtfsErrors      = 0
    }
}

function Get-ServerShareInfo {
    param([string]$Server)
    $key = $Server.ToUpperInvariant()
    if ($script:ServerCache.ContainsKey($key)) { return $script:ServerCache[$key] }

    $isLocal = Test-IsLocalComputer $Server
    $info = [pscustomobject]@{ Server = $Server; Error = $null; InfoLevel = 0; Shares = @() }
    $entries = $null
    $level = 0
    $rc = [ShareAudit.Native]::EnumShares($(if ($isLocal) { '' } else { $Server }), [ref]$entries, [ref]$level)
    if ($rc -ne 0) {
        $info.Error = 'NetShareEnum: ' + (Get-Win32ErrorText $rc)
        $script:ServerCache[$key] = $info
        return $info
    }
    $info.InfoLevel = $level
    $uncHost = if ($isLocal) { $env:COMPUTERNAME } else { $Server }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($e in $entries) {
        if ($e.BaseType -ne 0) { continue }   # tylko udziały dyskowe (bez drukarek, IPC, urządzeń)
        $isAdminShare = $e.IsSpecial -or ($script:AdminShareNames -contains $e.Name) -or ($e.Name -match '^[A-Za-z]\$$')
        $aclStatus = 'OK'
        $aces = @()
        if ($e.InfoLevel -ne 502) {
            $aclStatus = 'AccessDenied'
        } elseif ($null -ne $e.SecurityDescriptor -and $e.SecurityDescriptor.Length -gt 0) {
            try {
                $parsed = ConvertFrom-ShareSecurityDescriptor -Bytes $e.SecurityDescriptor
                $aclStatus = $parsed.Status
                $aces = $parsed.Aces
            } catch {
                $aclStatus = 'Error'
            }
        } elseif ($isAdminShare) {
            $aclStatus = 'Default'
        } else {
            $aclStatus = 'NullDacl'
        }
        $list.Add((New-AuditShareObject -Server $Server -Name $e.Name -UncPath ('\\{0}\{1}' -f $uncHost, $e.Name) `
                    -LocalPath $e.Path -Description $e.Remark -AclStatus $aclStatus -Aces $aces `
                    -IsAdminShare ([bool]$isAdminShare) -IsTemporary $e.IsTemporary -CurrentUses ([int]$e.CurrentUses)))
    }
    $info.Shares = $list.ToArray()
    $script:ServerCache[$key] = $info
    return $info
}

function Update-ShareFlags {
    param($Share)
    if ($Share.FlagsRead) { return }
    $Share.FlagsRead = $true
    $apiServer = if (Test-IsLocalComputer $Share.Server) { '' } else { $Share.Server }
    [uint32]$flags = 0
    try {
        $rc = [ShareAudit.Native]::GetShareFlags($apiServer, $Share.Name, [ref]$flags)
        if ($rc -eq 0) { $Share.Flags = [int64]$flags }
    } catch { }
}

function Test-ShareSelected {
    param($Share)
    if ($Share.IsAdminShare -and -not $IncludeAdminShares) { return $false }
    $included = $false
    foreach ($pattern in $IncludeShare) { if ($Share.Name -like $pattern) { $included = $true; break } }
    if (-not $included) { return $false }
    foreach ($pattern in $ExcludeShare) { if ($Share.Name -like $pattern) { return $false } }
    return $true
}

function Get-ShareBroadRank {
    param($Share)
    switch ($Share.AclStatus) {
        'NullDacl'  { return 4 }
        'Default'   { return 0 }
        'EmptyDacl' { return 0 }
        'OK' {
            $max = 0
            foreach ($ace in $Share.Aces) {
                if ($ace.Type -ne 'Allow') { continue }
                $p = Resolve-AuditPrincipal -Server $Share.Server -Sid $ace.Sid
                if ($p.IsBroad) {
                    $rank = Get-AccessRank $ace.Mask
                    if ($rank -gt $max) { $max = $rank }
                }
            }
            return $max
        }
        default { return $null }   # nieznane uprawnienia udziału
    }
}

function Get-ShareFlagText {
    param($Share, [string]$What)
    if ($null -eq $Share.Flags) { return '' }
    $f = [int64]$Share.Flags
    switch ($What) {
        'ABE'     { if ($f -band 0x800) { 'Tak' } else { 'Nie' } }
        'Encrypt' { if ($f -band 0x8000) { 'Tak' } else { 'Nie' } }
        'DFS'     { if ($f -band 0x3) { 'Tak' } else { 'Nie' } }
        'Caching' {
            switch ($f -band 0x30) {
                0x00 { 'Ręczne' }
                0x10 { 'Automatyczne (dokumenty)' }
                0x20 { 'Automatyczne (programy)' }
                0x30 { 'Wyłączone' }
            }
        }
    }
}
#endregion

#region Skanowanie NTFS (runspace'y / WinRM)
# Samodzielny blok - wykonywany w runspace lub zdalnie przez Invoke-Command, dlatego nie korzysta z funkcji skryptu.
# Wpisy ACL zwracane są jako tekst "SID|maska|typ|dziedziczony|IF|PF;..." - serializacja zdalna (głębokość 1)
# spłaszczyłaby zagnieżdżone obiekty.
$script:NtfsScanScript = {
    param([string]$RootPath, [int]$MaxDepth, [bool]$AllFolders)
    $sidType = [System.Security.Principal.SecurityIdentifier]
    $sections = [System.Security.AccessControl.AccessControlSections]'Access, Owner'
    $queue = New-Object 'System.Collections.Generic.Queue[object]'
    $queue.Enqueue([pscustomobject]@{ P = $RootPath; D = 0 })
    $scanned = 0
    $errors = 0
    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $scanned++
        $aceParts = New-Object System.Collections.Generic.List[string]
        $owner = $null
        $protected = $false
        $hasExplicit = $false
        $aclError = $null
        $childError = $null
        try {
            $security = New-Object System.Security.AccessControl.DirectorySecurity -ArgumentList $current.P, $sections
            $protected = $security.AreAccessRulesProtected
            try { $owner = $security.GetOwner($sidType).Value } catch { }
            foreach ($rule in $security.GetAccessRules($true, $true, $sidType)) {
                $mask = [int64][int]$rule.FileSystemRights
                if ($mask -lt 0) { $mask += 4294967296 }
                if (-not $rule.IsInherited) { $hasExplicit = $true }
                $aceParts.Add(('{0}|{1}|{2}|{3}|{4}|{5}' -f $rule.IdentityReference.Value, $mask, $rule.AccessControlType,
                        [int][bool]$rule.IsInherited, [int]$rule.InheritanceFlags, [int]$rule.PropagationFlags))
            }
        } catch {
            $aclError = $_.Exception.Message
            $errors++
        }
        if ($MaxDepth -lt 0 -or $current.D -lt $MaxDepth) {
            try {
                $dir = New-Object System.IO.DirectoryInfo -ArgumentList $current.P
                foreach ($child in $dir.EnumerateDirectories()) {
                    if (([int]$child.Attributes -band 0x400) -ne 0) { continue }   # junction / symlink - pomijamy
                    $queue.Enqueue([pscustomobject]@{ P = $child.FullName; D = $current.D + 1 })
                }
            } catch {
                $childError = $_.Exception.Message
                $errors++
            }
        }
        if ($AllFolders -or $current.D -eq 0 -or $hasExplicit -or $protected -or $aclError -or $childError) {
            [pscustomobject]@{
                Kind       = 'Folder'
                Path       = $current.P
                Depth      = $current.D
                Protected  = $protected
                Owner      = $owner
                Aces       = ($aceParts -join ';')
                Error      = $aclError
                ChildError = $childError
            }
        }
    }
    [pscustomobject]@{ Kind = 'Summary'; Scanned = $scanned; Errors = $errors }
}

$script:WorkerScript = {
    param($Job, [string]$ScanText, [int]$MaxDepth, [bool]$AllFolders, [string]$Mode, [System.Management.Automation.PSCredential]$Credential)
    $scan = [scriptblock]::Create($ScanText)
    $out = [pscustomobject]@{ Id = $Job.Id; Records = @(); Summary = $null; Error = $null; Warning = $null; UsedMode = 'UNC' }
    try {
        $raw = $null
        if ($Job.UseLocal) {
            $out.UsedMode = 'Lokalnie'
            $raw = @(& $scan $Job.LocalStart $MaxDepth $AllFolders)
        } else {
            if ($Mode -eq 'Remote' -and $Job.LocalStart) {
                try {
                    $icm = @{
                        ComputerName = $Job.Server
                        ScriptBlock  = $scan
                        ArgumentList = @($Job.LocalStart, $MaxDepth, $AllFolders)
                        ErrorAction  = 'Stop'
                    }
                    if ($Credential) { $icm['Credential'] = $Credential }
                    $raw = @(Invoke-Command @icm)
                    $out.UsedMode = 'WinRM'
                } catch {
                    $out.Warning = 'WinRM niedostępny ({0}) - skanowanie przez UNC.' -f $_.Exception.Message
                    $raw = $null
                }
            }
            if ($null -eq $raw) { $raw = @(& $scan $Job.UncStart $MaxDepth $AllFolders) }
        }
        $out.Records = @($raw | Where-Object { $_.Kind -eq 'Folder' })
        $out.Summary = @($raw | Where-Object { $_.Kind -eq 'Summary' }) | Select-Object -First 1
    } catch {
        $out.Error = $_.Exception.Message
    }
    $out
}

function Invoke-NtfsScan {
    param([object[]]$Targets)
    $results = @{}
    $pool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
    $pool.Open()
    $workers = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($target in $Targets) {
            $ps = [powershell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($script:WorkerScript.ToString())
            [void]$ps.AddParameter('Job', $target)
            [void]$ps.AddParameter('ScanText', $script:NtfsScanScript.ToString())
            [void]$ps.AddParameter('MaxDepth', $Depth)
            [void]$ps.AddParameter('AllFolders', [bool]$AllFolders)
            [void]$ps.AddParameter('Mode', $ScanMode)
            [void]$ps.AddParameter('Credential', $Credential)
            $workers.Add([pscustomobject]@{ Target = $target; PS = $ps; Handle = $ps.BeginInvoke() })
        }
        $total = $workers.Count
        while ($true) {
            $done = @($workers | Where-Object { $_.Handle.IsCompleted }).Count
            $running = @($workers | Where-Object { -not $_.Handle.IsCompleted } | Select-Object -First 3 | ForEach-Object { $_.Target.UncStart })
            Write-Progress -Id 1 -Activity 'Skanowanie NTFS' -Status ('Ukończono {0} z {1}' -f $done, $total) `
                -CurrentOperation ($running -join ' | ') -PercentComplete ([int](100 * $done / [Math]::Max(1, $total)))
            if ($done -ge $total) { break }
            Start-Sleep -Milliseconds 300
        }
        foreach ($w in $workers) {
            try {
                $output = $w.PS.EndInvoke($w.Handle)
                if ($output.Count -gt 0) { $results[$w.Target.Id] = $output[$output.Count - 1] }
                elseif ($w.PS.Streams.Error.Count -gt 0) {
                    $results[$w.Target.Id] = [pscustomobject]@{ Records = @(); Summary = $null; Warning = $null; UsedMode = ''; Error = [string]$w.PS.Streams.Error[0] }
                }
            } catch {
                $results[$w.Target.Id] = [pscustomobject]@{ Records = @(); Summary = $null; Warning = $null; UsedMode = ''; Error = $_.Exception.Message }
            }
        }
    } finally {
        Write-Progress -Id 1 -Activity 'Skanowanie NTFS' -Completed
        foreach ($w in $workers) {
            if (-not $w.Handle.IsCompleted) { try { $w.PS.Stop() } catch { } }
            $w.PS.Dispose()
        }
        $pool.Close()
        $pool.Dispose()
    }
    return $results
}

function ConvertFrom-AceText {
    param([string]$Text)
    $list = New-Object System.Collections.Generic.List[object]
    if ($Text) {
        foreach ($part in $Text.Split(';')) {
            if (-not $part) { continue }
            $f = $part.Split('|')
            if ($f.Count -lt 6) { continue }
            $list.Add([pscustomobject]@{
                Sid       = $f[0]
                Mask      = [int64]$f[1]
                Type      = $f[2]
                Inherited = ($f[3] -eq '1')
                IF        = [int]$f[4]
                PF        = [int]$f[5]
            })
        }
    }
    return , $list.ToArray()
}

function ConvertTo-AuditUncPath {
    param([string]$Path, [string]$LocalRoot, [string]$UncRoot)
    if (-not $LocalRoot -or -not $UncRoot -or $LocalRoot -eq $UncRoot) { return $Path }
    $root = $LocalRoot.TrimEnd('\')
    if (-not $Path.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { return $Path }
    $relative = $Path.Substring($root.Length).TrimStart('\')
    if ($relative) { return '{0}\{1}' -f $UncRoot.TrimEnd('\'), $relative }
    return $UncRoot.TrimEnd('\')
}
#endregion

#region Ustalenia (reguły ryzyka)
function Add-Finding {
    param(
        [Parameter(Position = 0)][string]$Severity,
        [Parameter(Position = 1)][string]$Code,
        [Parameter(Position = 2)][string]$Category,
        [Parameter(Position = 3)][string]$Server,
        [Parameter(Position = 4)][string]$Share,
        [Parameter(Position = 5)][string]$Path,
        [string]$Principal = '',
        [string]$Access = '',
        [string]$Title,
        [string]$Details = '',
        [string]$Recommendation = ''
    )
    $key = '{0}|{1}|{2}|{3}' -f $Code, $Server, $Path, $Principal
    if (-not $script:FindingKeys.Add($key)) { return }
    $script:Findings.Add([pscustomobject]@{
        Severity       = $Severity
        SeverityRank   = $script:SeverityRank[$Severity]
        Code           = $Code
        Category       = $Category
        Server         = $Server
        Share          = $Share
        Path           = $Path
        Principal      = $Principal
        Access         = $Access
        Title          = $Title
        Details        = $Details
        Recommendation = $Recommendation
    })
}

function Add-ShareFindings {
    param($Share)
    $srv = $Share.Server
    $name = $Share.Name
    $unc = $Share.UncPath
    switch ($Share.AclStatus) {
        'NullDacl' {
            Add-Finding 'Wysokie' 'SHR-NULL-DACL' 'Udział' $srv $name $unc -Title 'Udział bez listy uprawnień (brak DACL)' `
                -Details 'Udział nie ma deskryptora zabezpieczeń - na poziomie udziału każdy ma pełny dostęp, ochronę zapewnia wyłącznie NTFS.' `
                -Recommendation 'Nadaj jawne uprawnienia udziału, np. Authenticated Users: Zmiana, Administrators: Pełna kontrola.'
        }
        'EmptyDacl' {
            Add-Finding 'Info' 'SHR-EMPTY-DACL' 'Udział' $srv $name $unc -Title 'Pusta lista uprawnień udziału' `
                -Details 'Nikt nie ma dostępu przez ten udział.' -Recommendation 'Usuń nieużywany udział lub nadaj właściwe uprawnienia.'
        }
        'AccessDenied' {
            Add-Finding 'Info' 'SHR-ACL-UNREADABLE' 'Udział' $srv $name $unc -Title 'Nie odczytano uprawnień udziału' `
                -Details 'Odczyt uprawnień udziału (NetShareEnum, poziom 502) wymaga uprawnień administratora na serwerze.' `
                -Recommendation 'Uruchom audyt kontem z uprawnieniami administratora serwera lub użyj -Credential.'
        }
        'Error' {
            Add-Finding 'Info' 'SHR-ACL-UNREADABLE' 'Udział' $srv $name $unc -Title 'Nie odczytano uprawnień udziału' `
                -Details 'Nie udało się zinterpretować deskryptora zabezpieczeń udziału.'
        }
    }

    foreach ($ace in $Share.Aces) {
        $p = Resolve-AuditPrincipal -Server $srv -Sid $ace.Sid
        $rank = Get-AccessRank $ace.Mask
        $access = '{0} ({1})' -f (Get-AccessLabel $ace.Mask), $ace.Type
        if ($p.Type -eq 'Orphaned') {
            Add-Finding 'Niskie' 'SHR-ORPHANED-SID' 'Udział' $srv $name $unc -Principal $p.Name -Access $access -Title 'Osierocony SID w uprawnieniach udziału' `
                -Details 'Konto zostało usunięte, a wpis pozostał.' -Recommendation 'Usuń nieaktualny wpis.'
        }
        if ($p.Enabled -eq $false) {
            Add-Finding 'Niskie' 'SHR-DISABLED-ACCOUNT' 'Udział' $srv $name $unc -Principal $p.Name -Access $access -Title 'Wyłączone konto w uprawnieniach udziału' `
                -Details 'Konto jest wyłączone w AD, ale nadal widnieje w uprawnieniach.' -Recommendation 'Usuń wpis lub konto, jeśli nie jest już potrzebne.'
        }
        if ($ace.Type -eq 'Deny') {
            Add-Finding 'Info' 'SHR-DENY' 'Udział' $srv $name $unc -Principal $p.Name -Access $access -Title 'Wpis Deny w uprawnieniach udziału' `
                -Details 'Wpisy odmowy są trudne w utrzymaniu i łatwo o nieoczekiwane skutki (Deny dla grupy obejmuje wszystkich jej członków).' `
                -Recommendation 'Zamiast Deny ogranicz listę grup z dostępem.'
            continue
        }
        if ($p.IsAnonymous -and $rank -ge 1) {
            Add-Finding 'Wysokie' 'SHR-ANONYMOUS' 'Udział' $srv $name $unc -Principal $p.Name -Access $access -Title 'Dostęp anonimowy/gościa do udziału' `
                -Details ('{0} ma dostęp do udziału.' -f $p.Name) -Recommendation 'Usuń dostęp dla kont anonimowych i gości.'
        } elseif ($p.IsBroad -and $rank -ge 3) {
            Add-Finding 'Niskie' 'SHR-BROAD-WRITE' 'Udział' $srv $name $unc -Principal $p.Name -Access $access -Title 'Szeroka grupa z prawem zapisu na poziomie udziału' `
                -Details ('{0} ma "{1}" na udziale - zapis ogranicza wyłącznie NTFS.' -f $p.Name, (Get-AccessLabel $ace.Mask)) `
                -Recommendation 'To akceptowalny model tylko przy restrykcyjnym NTFS. Rozważ Authenticated Users: Zmiana zamiast Everyone: Pełna kontrola.'
        }
        if ($p.Type -eq 'User' -and -not ($p.IsAdmin -or $p.IsTrusted)) {
            Add-Finding 'Niskie' 'SHR-DIRECT-USER' 'Udział' $srv $name $unc -Principal $p.Name -Access $access -Title 'Uprawnienie udziału nadane bezpośrednio użytkownikowi' `
                -Details 'Uprawnienia nadawane pojedynczym kontom są trudne w utrzymaniu i audycie.' -Recommendation 'Nadawaj uprawnienia grupom (model AGDLP).'
        }
    }

    if ($Share.IsHidden -and -not $Share.IsAdminShare) {
        Add-Finding 'Info' 'SHR-HIDDEN' 'Udział' $srv $name $unc -Title 'Ukryty udział' `
            -Details 'Nazwa kończy się znakiem $ - udział nie jest widoczny w przeglądaniu sieci, ale ukrycie nie jest zabezpieczeniem.' `
            -Recommendation 'Upewnij się, że uprawnienia udziału i NTFS są właściwe.'
    }
    if ($null -ne $Share.Flags -and -not ([int64]$Share.Flags -band 0x800) -and -not $Share.IsAdminShare) {
        Add-Finding 'Info' 'SHR-NO-ABE' 'Udział' $srv $name $unc -Title 'Access-Based Enumeration wyłączone' `
            -Details 'Użytkownicy widzą foldery, do których nie mają dostępu (ujawnianie struktury i nazw).' `
            -Recommendation 'Włącz ABE: Set-SmbShare -Name <udział> -FolderEnumerationMode AccessBased.'
    }
}

function Add-NtfsFindings {
    param($Folder)
    $srv = $Folder.Server
    $shareName = $Folder.ShareName
    $path = $Folder.Path
    $shareBroad = $Folder.Share.BroadRank

    if ($Folder.Error) {
        if ($Folder.Depth -eq 0) {
            Add-Finding 'Info' 'NTFS-ACL-UNREADABLE' 'NTFS' $srv $shareName $path -Title 'Nie odczytano uprawnień NTFS folderu głównego' `
                -Details $Folder.Error -Recommendation 'Sprawdź, czy konto audytu ma prawo odczytu uprawnień (Read permissions).'
        }
        return
    }

    foreach ($ace in $Folder.Aces) {
        if ($Folder.Depth -gt 0 -and $ace.Inherited) { continue }   # w podfolderach oceniamy tylko wpisy jawne
        $p = Resolve-AuditPrincipal -Server $srv -Sid $ace.Sid
        $rank = Get-AccessRank $ace.Mask
        $origin = if ($ace.Inherited) { 'dziedziczone' } else { 'jawne' }
        $access = '{0} ({1}; {2}; {3})' -f (Get-AccessLabel $ace.Mask), $ace.Type, (Get-AppliesToText $ace.IF $ace.PF), $origin

        if ($p.Type -eq 'Orphaned') {
            Add-Finding 'Niskie' 'NTFS-ORPHANED-SID' 'NTFS' $srv $shareName $path -Principal $p.Name -Access $access -Title 'Osierocony SID w NTFS' `
                -Details 'Konto zostało usunięte, a wpis pozostał w ACL.' -Recommendation 'Usuń nieaktualny wpis z listy uprawnień.'
        }
        if ($p.Enabled -eq $false) {
            Add-Finding 'Niskie' 'NTFS-DISABLED-ACCOUNT' 'NTFS' $srv $shareName $path -Principal $p.Name -Access $access -Title 'Wyłączone konto w uprawnieniach NTFS' `
                -Details 'Konto jest wyłączone w AD, ale nadal ma przypisane uprawnienia.' -Recommendation 'Usuń wpis podczas porządkowania uprawnień.'
        }
        if ($ace.Type -eq 'Deny') {
            if (-not $ace.Inherited) {
                Add-Finding 'Info' 'NTFS-DENY' 'NTFS' $srv $shareName $path -Principal $p.Name -Access $access -Title 'Jawny wpis Deny w NTFS' `
                    -Details 'Wpisy odmowy komplikują model uprawnień i utrudniają diagnozę problemów z dostępem.' `
                    -Recommendation 'Rozważ przebudowę uprawnień tak, aby obyć się bez Deny.'
            }
            continue
        }

        if ($p.IsAnonymous -and $rank -ge 1) {
            Add-Finding 'Wysokie' 'NTFS-ANONYMOUS' 'NTFS' $srv $shareName $path -Principal $p.Name -Access $access -Title 'Dostęp anonimowy/gościa w NTFS' `
                -Details ('{0} ma dostęp do folderu.' -f $p.Name) -Recommendation 'Usuń uprawnienia kont anonimowych i gości.'
        } elseif ($p.IsBroad -and $rank -ge 2) {
            # Waga zależy od tego, co szeroka grupa może zrobić efektywnie (NTFS ∩ udział).
            $what = if ($rank -ge 3) { 'modyfikować i usuwać dane' } else { 'tworzyć lub dopisywać pliki/foldery' }
            if ($null -eq $shareBroad) {
                $severity = 'Wysokie'
                $details = '{0} może {1} w tym miejscu. Nie ustalono, czy uprawnienia udziału to ograniczają.' -f $p.Name, $what
            } elseif ($shareBroad -ge 2) {
                $effectiveRank = [Math]::Min($rank, [int]$shareBroad)
                $severity = if ($effectiveRank -ge 3) { 'Krytyczne' } else { 'Wysokie' }
                $what = if ($effectiveRank -ge 3) { 'modyfikować i usuwać dane' } else { 'tworzyć lub dopisywać pliki/foldery' }
                $details = 'Praktycznie każdy użytkownik może {1} ({0} w NTFS, a uprawnienia udziału też dopuszczają zapis szerokiej grupie) - ryzyko ransomware, wycieku i utraty danych.' -f $p.Name, $what
            } else {
                $severity = 'Średnie'
                $details = '{0} może {1} według NTFS; obecnie blokuje to wyłącznie konfiguracja udziału (zmiana udziału, dostęp lokalny lub przez inny udział otworzy zapis dla wszystkich).' -f $p.Name, $what
            }
            Add-Finding $severity 'NTFS-BROAD-WRITE' 'NTFS' $srv $shareName $path -Principal $p.Name -Access $access -Title 'Szeroka grupa z prawem zapisu w NTFS' `
                -Details $details -Recommendation 'Zastąp szeroką grupę dedykowaną grupą uprawnień (np. FS_<folder>_RW) i ogranicz zapis do osób, które go potrzebują.'
        } elseif ($rank -eq 4 -and -not ($p.IsAdmin -or $p.IsTrusted -or $p.IsCreatorOwner -or $p.IsBroad)) {
            Add-Finding 'Średnie' 'NTFS-FULL-CONTROL' 'NTFS' $srv $shareName $path -Principal $p.Name -Access $access -Title 'Pełna kontrola dla konta nieadministracyjnego' `
                -Details 'Pełna kontrola (lub WRITE_DAC/WRITE_OWNER) pozwala zmieniać uprawnienia i przejmować własność - podmiot może nadać dostęp komukolwiek.' `
                -Recommendation 'Zamień na Modyfikację; pełną kontrolę pozostaw administratorom.'
        }

        if (-not $ace.Inherited -and $p.Type -eq 'User' -and -not ($p.IsAdmin -or $p.IsTrusted)) {
            Add-Finding 'Niskie' 'NTFS-DIRECT-USER' 'NTFS' $srv $shareName $path -Principal $p.Name -Access $access -Title 'Uprawnienie NTFS nadane bezpośrednio użytkownikowi' `
                -Details 'Uprawnienia nadawane pojedynczym kontom są trudne w utrzymaniu, audycie i odbieraniu przy zmianie stanowiska.' `
                -Recommendation 'Przenieś użytkownika do grupy uprawnień i nadaj prawa grupie (model AGDLP).'
        }
    }

    if ($Folder.Depth -gt 0 -and $Folder.Protected) {
        Add-Finding 'Info' 'NTFS-INHERITANCE-OFF' 'NTFS' $srv $shareName $path -Title 'Wyłączone dziedziczenie uprawnień' `
            -Details 'Folder ma własną, niezależną listę uprawnień - zmiany na folderze nadrzędnym nie będą tu stosowane.' `
            -Recommendation 'Sprawdź, czy przerwanie dziedziczenia jest zamierzone i udokumentowane.'
    }

    if ($Folder.Depth -eq 0) {
        if ($Folder.Owner) {
            $owner = Resolve-AuditPrincipal -Server $srv -Sid $Folder.Owner
            if ($owner.Type -eq 'User' -and -not ($owner.IsAdmin -or $owner.IsTrusted)) {
                Add-Finding 'Niskie' 'NTFS-OWNER-USER' 'NTFS' $srv $shareName $path -Principal $owner.Name -Title 'Właścicielem folderu głównego jest zwykły użytkownik' `
                    -Details 'Właściciel ma niejawne prawo zmiany uprawnień folderu, niezależnie od wpisów ACL.' `
                    -Recommendation 'Ustaw właściciela na grupę Administrators.'
            }
        }
        $hasAdminAccess = $false
        foreach ($ace in $Folder.Aces) {
            if ($ace.Type -eq 'Allow' -and ($ace.Sid -eq 'S-1-5-18' -or $ace.Sid -eq 'S-1-5-32-544') -and (([int]$ace.PF -band 2) -eq 0)) {
                $hasAdminAccess = $true
                break
            }
        }
        if (-not $hasAdminAccess -and $Folder.Aces.Count -gt 0) {
            Add-Finding 'Info' 'NTFS-NO-ADMIN-ACCESS' 'NTFS' $srv $shareName $path -Title 'Brak dostępu SYSTEM/Administrators do folderu głównego' `
                -Details 'Utrudnia to kopie zapasowe, skanowanie antywirusowe i administrację.' `
                -Recommendation 'Dodaj SYSTEM i Administrators z pełną kontrolą.'
        }
    }
}
#endregion

#region Członkostwo grup
function Get-LdapValue {
    param($Result, [string]$Name)
    $values = $Result.Properties[$Name]
    if ($values -and $values.Count -gt 0) { return $values[0] }
    return $null
}

function Get-ADGroupMemberList {
    param([string]$Sid)
    $result = [pscustomobject]@{ Members = New-Object System.Collections.Generic.List[object]; Truncated = $false; Error = $null }
    try {
        $group = [ADSI]('LDAP://<SID={0}>' -f $Sid)
        $dn = [string]$group.Properties['distinguishedName'].Value
        if (-not $dn) { throw 'Nie znaleziono grupy w AD.' }
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        # LDAP_MATCHING_RULE_IN_CHAIN - członkostwo rekurencyjne w jednym zapytaniu
        $searcher.Filter = '(&(objectClass=user)(memberOf:1.2.840.113556.1.4.1941:={0}))' -f (ConvertTo-LdapFilterValue $dn)
        $searcher.PageSize = 500
        foreach ($prop in 'samaccountname', 'displayname', 'useraccountcontrol', 'objectsid') { [void]$searcher.PropertiesToLoad.Add($prop) }
        $found = $searcher.FindAll()
        try {
            foreach ($r in $found) {
                if ($result.Members.Count -ge $MaxGroupMembers) { $result.Truncated = $true; break }
                $uac = Get-LdapValue $r 'useraccountcontrol'
                $sam = [string](Get-LdapValue $r 'samaccountname')
                $sidBytes = Get-LdapValue $r 'objectsid'
                $result.Members.Add([pscustomobject]@{
                    Name        = $sam
                    DisplayName = [string](Get-LdapValue $r 'displayname')
                    Type        = if ($sam.EndsWith('$')) { 'Computer' } else { 'User' }
                    Enabled     = if ($null -ne $uac) { -not ([int]$uac -band 2) } else { $null }
                    Sid         = if ($sidBytes) { ([System.Security.Principal.SecurityIdentifier]::new([byte[]]$sidBytes, 0)).Value } else { '' }
                })
            }
        } finally {
            $found.Dispose()
        }
    } catch {
        $result.Error = $_.Exception.Message
    }
    return $result
}

function Get-LocalGroupMemberList {
    param([string]$Server, $Principal)
    $key = '{0}|{1}' -f $Server, $Principal.Sid
    if ($script:LocalGroupCache.ContainsKey($key)) { return $script:LocalGroupCache[$key] }
    $result = [pscustomobject]@{ Members = New-Object System.Collections.Generic.List[object]; Truncated = $false; Error = $null }
    try {
        if (-not $Principal.Account) { throw 'Nie znamy nazwy grupy lokalnej.' }
        $hostName = if (Test-IsLocalComputer $Server) { $env:COMPUTERNAME } else { $Server }
        $group = [ADSI]('WinNT://{0}/{1},group' -f $hostName, $Principal.Account)
        foreach ($member in @($group.psbase.Invoke('Members'))) {
            if ($result.Members.Count -ge $MaxGroupMembers) { $result.Truncated = $true; break }
            $sidBytes = $member.GetType().InvokeMember('objectSid', 'GetProperty', $null, $member, $null)
            $sid = ([System.Security.Principal.SecurityIdentifier]::new([byte[]]$sidBytes, 0)).Value
            $mp = Resolve-AuditPrincipal -Server $Server -Sid $sid
            $result.Members.Add([pscustomobject]@{ Name = $mp.Name; DisplayName = ''; Type = $mp.Type; Enabled = $mp.Enabled; Sid = $sid })
        }
    } catch {
        $result.Error = $_.Exception.Message
    }
    $script:LocalGroupCache[$key] = $result
    return $result
}

function Get-GroupMemberRows {
    $rows = New-Object System.Collections.Generic.List[object]
    $groups = @($script:UsedGroups.Values | Where-Object { -not $_.Principal.IsBroad } | Sort-Object { $_.Principal.Name })
    $i = 0
    foreach ($g in $groups) {
        $i++
        $p = $g.Principal
        Write-Progress -Id 1 -Activity 'Rozwijanie grup' -Status $p.Name -PercentComplete ([int](100 * $i / [Math]::Max(1, $groups.Count)))
        if ($p.Scope -eq 'Domain') {
            if (-not $script:ADAvailable) { continue }
            $list = Get-ADGroupMemberList -Sid $p.Sid
            $source = 'AD (rekurencyjnie)'
            $groupScope = 'Domenowa'
            $server = ''
        } else {
            $list = Get-LocalGroupMemberList -Server $g.Server -Principal $p
            $source = 'Grupa lokalna serwera (bezpośrednio)'
            $groupScope = 'Lokalna'
            $server = $g.Server
        }
        if ($list.Error) {
            Add-AuditError -Server $server -Stage 'Członkowie grupy' -Message ('{0}: {1}' -f $p.Name, $list.Error) -Quiet
        }
        $note = if ($list.Error) { 'Błąd: ' + $list.Error } elseif ($list.Truncated) { 'Obcięto do {0} członków' -f $MaxGroupMembers } else { '' }
        if ($list.Members.Count -eq 0) {
            $rows.Add([pscustomobject]@{ Group = $p.Name; GroupScope = $groupScope; Server = $server; Member = ''; DisplayName = ''; MemberType = ''; Enabled = $null; Source = $source; Note = $(if ($note) { $note } else { 'Grupa pusta' }) })
            continue
        }
        foreach ($m in $list.Members) {
            $rows.Add([pscustomobject]@{ Group = $p.Name; GroupScope = $groupScope; Server = $server; Member = $m.Name; DisplayName = $m.DisplayName; MemberType = $m.Type; Enabled = $m.Enabled; Source = $source; Note = $note })
        }
    }
    Write-Progress -Id 1 -Activity 'Rozwijanie grup' -Completed
    return , $rows
}
#endregion

#region Efektywny dostęp użytkowników
function Get-UserSecurityContext {
    param([string]$Identity)
    $context = [pscustomobject]@{ Identity = $Identity; Name = $Identity; Sid = $null; Sids = $null; Error = $null }
    try {
        $sam = $Identity
        if ($sam -match '^[^\\]+\\(.+)$') { $sam = $Matches[1] }
        $searcher = New-Object System.DirectoryServices.DirectorySearcher
        $searcher.Filter = '(&(objectClass=user)(|(sAMAccountName={0})(userPrincipalName={1})))' -f (ConvertTo-LdapFilterValue $sam), (ConvertTo-LdapFilterValue $Identity)
        $hit = $searcher.FindOne()
        if (-not $hit) { throw ('Nie znaleziono konta "{0}" w AD.' -f $Identity) }
        $entry = $hit.GetDirectoryEntry()
        # tokenGroups = wszystkie grupy zabezpieczeń użytkownika (rekurencyjnie, łącznie z grupą podstawową)
        $entry.RefreshCache([string[]]@('tokenGroups', 'objectSid', 'sAMAccountName', 'displayName'))
        $userSid = ([System.Security.Principal.SecurityIdentifier]::new([byte[]]$entry.Properties['objectSid'].Value, 0)).Value
        $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        [void]$set.Add($userSid)
        foreach ($tokenGroup in $entry.Properties['tokenGroups']) {
            [void]$set.Add(([System.Security.Principal.SecurityIdentifier]::new([byte[]]$tokenGroup, 0)).Value)
        }
        # Everyone, Authenticated Users, NETWORK (dostęp przez sieć), This Organization
        foreach ($wellKnown in 'S-1-1-0', 'S-1-5-11', 'S-1-5-2', 'S-1-5-15') { [void]$set.Add($wellKnown) }
        $sam = [string]$entry.Properties['sAMAccountName'].Value
        $display = [string]$entry.Properties['displayName'].Value
        $context.Name = if ($display) { '{0} ({1})' -f $sam, $display } else { $sam }
        $context.Sid = $userSid
        $context.Sids = $set
    } catch {
        $context.Error = $_.Exception.Message
    }
    return $context
}

# SID-y użytkownika na danym serwerze = SID-y domenowe + lokalne grupy serwera, do których należy on lub jego grupy.
function Get-UserServerSidSet {
    param($Context, [string]$Server)
    $key = '{0}|{1}' -f $Server, $Context.Sid
    if ($script:UserSidSetCache.ContainsKey($key)) { return , $script:UserSidSetCache[$key] }
    $set = [System.Collections.Generic.HashSet[string]]::new($Context.Sids, [System.StringComparer]::OrdinalIgnoreCase)
    $lookupFailed = $false
    foreach ($g in @($script:UsedGroups.Values)) {
        if ($g.Server -ne $Server) { continue }
        $p = $g.Principal
        if ($p.Scope -ne 'Local' -and $p.Scope -ne 'Builtin') { continue }
        $members = Get-LocalGroupMemberList -Server $Server -Principal $p
        if ($members.Error) { $lookupFailed = $true; continue }
        foreach ($m in $members.Members) {
            if ($Context.Sids.Contains([string]$m.Sid)) { [void]$set.Add($p.Sid); break }
        }
    }
    if ($lookupFailed) {
        # Nie udało się odczytać grup lokalnych - przyjmujemy domyślną konfigurację serwera członkowskiego.
        [void]$set.Add('S-1-5-32-545')
        foreach ($s in $Context.Sids) { if ($s -match '-512$') { [void]$set.Add('S-1-5-32-544'); break } }
    }
    $script:UserSidSetCache[$key] = $set
    return , $set
}

function Get-UserAccessRows {
    param([object[]]$Contexts)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($ctx in $Contexts) {
        $i = 0
        foreach ($folder in $script:Folders) {
            $i++
            if ($i % 200 -eq 1) {
                Write-Progress -Id 1 -Activity ('Efektywny dostęp: {0}' -f $ctx.Name) -Status $folder.Path -PercentComplete ([int](100 * $i / [Math]::Max(1, $script:Folders.Count)))
            }
            if ($folder.Error) { continue }
            $set = Get-UserServerSidSet -Context $ctx -Server $folder.Server
            $share = $folder.Share
            $shareMask = $null
            switch ($share.AclStatus) {
                'NullDacl' { $shareMask = [int64]0x1F01FF }
                'OK'       { $shareMask = Get-EffectiveMask -Aces $share.Aces -SidSet $set }
                'EmptyDacl' { $shareMask = [int64]0 }
                'Default'  { $shareMask = if ($set.Contains('S-1-5-32-544')) { [int64]0x1F01FF } else { [int64]0 } }
            }
            $ntfsMask = Get-EffectiveMask -Aces $folder.Aces -SidSet $set -FolderOnly
            $effective = if ($null -ne $shareMask) { $ntfsMask -band $shareMask } else { $ntfsMask }
            $effectiveRank = Get-AccessRank $effective
            if ($effective -eq 0 -or $effectiveRank -lt 1) { continue }

            $via = New-Object System.Collections.Generic.List[string]
            $denyPresent = $false
            foreach ($ace in $folder.Aces) {
                if ((([int]$ace.PF -band 2) -ne 0) -or -not $set.Contains([string]$ace.Sid)) { continue }
                if ($ace.Type -eq 'Deny') { $denyPresent = $true; continue }
                $n = (Resolve-AuditPrincipal -Server $folder.Server -Sid $ace.Sid).Name
                if (-not $via.Contains($n)) { $via.Add($n) }
            }
            $shareText = if ($share.AclStatus -eq 'NotShare') { 'n/d' } elseif ($null -eq $shareMask) { 'nieznane' } else { Get-AccessLabel $shareMask }
            $rows.Add([pscustomobject]@{
                User            = $ctx.Name
                Server          = $folder.Server
                Share           = $folder.ShareName
                Path            = $folder.Path
                Depth           = $folder.Depth
                ShareAccess     = $shareText
                NtfsAccess      = Get-AccessLabel $ntfsMask
                EffectiveAccess = Get-AccessLabel $effective
                EffectiveRank   = $effectiveRank
                Via             = ($via -join ', ')
                DenyPresent     = $denyPresent
            })
        }
    }
    Write-Progress -Id 1 -Activity 'Efektywny dostęp' -Completed
    return , $rows
}
#endregion

#region Raport HTML
function New-AuditHtmlReport {
    param($Data, [string]$FilePath)
    $json = $Data | ConvertTo-Json -Depth 6 -Compress
    # '<' zakodowane jako \u003c - dane nie mogą zamknąć znacznika <script>
    $json = $json.Replace('<', '\u003c')
    $template = @'
<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Audyt udziałów sieciowych</title>
<style>
* { box-sizing: border-box; }
body { margin: 0; font-family: "Segoe UI", Roboto, Arial, sans-serif; font-size: 14px; line-height: 1.45; background: #f4f5f7; color: #1f2328; }
header { background: #1f2a44; color: #fff; padding: 18px 24px 14px; }
header h1 { margin: 0 0 4px; font-size: 20px; font-weight: 600; }
header .meta { font-size: 12px; color: #c9d1e3; }
nav { background: #fff; border-bottom: 1px solid #d8dee4; padding: 0 16px; position: sticky; top: 0; z-index: 5; }
nav button { background: none; border: 0; border-bottom: 3px solid transparent; padding: 11px 12px 9px; margin: 0 2px; font: inherit; color: #57606a; cursor: pointer; }
nav button:hover { color: #1f2328; }
nav button.active { color: #0b5cad; border-bottom-color: #0b5cad; font-weight: 600; }
nav .cnt { display: inline-block; min-width: 18px; padding: 0 6px; margin-left: 4px; border-radius: 9px; background: #eaeef2; color: #57606a; font-size: 11px; font-weight: 600; text-align: center; }
main { padding: 20px 24px 40px; }
section { display: none; }
section.active { display: block; }
h2 { font-size: 16px; margin: 26px 0 10px; font-weight: 600; }
h2:first-child { margin-top: 0; }
.muted { color: #6e7781; }
.cards { margin: 0 -6px; }
.card { display: inline-block; vertical-align: top; min-width: 170px; margin: 0 6px 12px; padding: 12px 16px; background: #fff; border: 1px solid #d8dee4; border-radius: 8px; text-decoration: none; color: inherit; }
.card .v { font-size: 24px; font-weight: 600; }
.card .l { font-size: 12px; color: #6e7781; }
a.card { cursor: pointer; }
a.card:hover { border-color: #0b5cad; }
.sevcard { border-left-width: 6px; }
.sevcard.crit { border-left-color: #a4161a; } .sevcard.high { border-left-color: #d9480f; } .sevcard.med { border-left-color: #e0a800; } .sevcard.low { border-left-color: #1c64d6; } .sevcard.info { border-left-color: #8c959f; }
.badge { display: inline-block; padding: 1px 8px; border-radius: 10px; font-size: 12px; font-weight: 600; color: #fff; white-space: nowrap; }
.badge.crit { background: #a4161a; } .badge.high { background: #d9480f; } .badge.med { background: #e0a800; color: #1f2328; } .badge.low { background: #1c64d6; } .badge.info { background: #8c959f; }
.panel { background: #fff; border: 1px solid #d8dee4; border-radius: 8px; padding: 14px 16px; }
.bars .row { margin: 6px 0; }
.bars .lbl { margin-bottom: 3px; }
.bars .track { background: #eaeef2; border-radius: 4px; height: 10px; position: relative; }
.bars .fill { height: 10px; border-radius: 4px; }
.fill.crit { background: #a4161a; } .fill.high { background: #d9480f; } .fill.med { background: #e0a800; } .fill.low { background: #1c64d6; } .fill.info { background: #8c959f; }
.twocol { overflow: hidden; margin: 0 -8px; }
.twocol > div { float: left; width: 50%; padding: 0 8px; }
@media (max-width: 900px) { .twocol > div { float: none; width: 100%; } }
.toolbar { margin-bottom: 8px; }
.toolbar input, .toolbar select, .toolbar button { font: inherit; padding: 5px 8px; margin: 0 6px 6px 0; border: 1px solid #c7ced6; border-radius: 6px; background: #fff; color: inherit; vertical-align: middle; }
.toolbar input.q { width: 280px; max-width: 100%; }
.toolbar button { cursor: pointer; }
.toolbar button:hover { border-color: #0b5cad; }
.toolbar .count { color: #6e7781; font-size: 12px; margin-left: 4px; }
.tablewrap { overflow: auto; max-height: 72vh; background: #fff; border: 1px solid #d8dee4; border-radius: 8px; }
table { border-collapse: collapse; width: 100%; }
th, td { padding: 6px 9px; border-bottom: 1px solid #eaeef2; text-align: left; vertical-align: top; }
th { position: sticky; top: 0; background: #f6f8fa; font-weight: 600; font-size: 12px; color: #424a53; cursor: pointer; white-space: nowrap; z-index: 1; }
th:hover { color: #0b5cad; }
tr:hover td { background: #f6f8fa; }
td.empty { text-align: center; color: #6e7781; padding: 24px; }
td .long { display: block; min-width: 260px; max-width: 480px; }
.mono { font-family: Consolas, "Cascadia Mono", monospace; font-size: 12.5px; word-break: break-all; }
td .mono { display: block; min-width: 190px; max-width: 460px; }
.tag { display: inline-block; padding: 0 6px; border-radius: 4px; background: #eaeef2; font-size: 12px; }
.lv4 { color: #a4161a; font-weight: 600; } .lv3 { color: #c2410c; font-weight: 600; } .lv2 { color: #9a6700; } .lv1 { color: #1a7f37; } .lv0 { color: #6e7781; }
.pager { margin-top: 8px; }
.pager button { font: inherit; padding: 3px 10px; margin-right: 4px; border: 1px solid #c7ced6; border-radius: 6px; background: #fff; color: inherit; cursor: pointer; }
.pager button[disabled] { opacity: .4; cursor: default; }
.note { padding: 8px 12px; margin: 0 0 10px; border-radius: 6px; background: #fff8c5; border: 1px solid #e3cf7a; color: #3b2e00; }
dl.params { margin: 0; }
dl.params dt { font-weight: 600; float: left; clear: left; width: 190px; }
dl.params dd { margin: 0 0 4px 200px; overflow-wrap: break-word; word-wrap: break-word; }
.panel { overflow-x: auto; }
@media (max-width: 640px) {
  header { padding: 14px 16px 12px; }
  nav { padding: 0 6px; }
  nav button { padding: 9px 8px 7px; }
  main { padding: 16px 16px 32px; }
  .card { min-width: 0; width: calc(50% - 12px); }
  .toolbar input.q { width: 100%; }
  dl.params dt { float: none; width: auto; }
  dl.params dd { margin: 0 0 8px 0; }
}
@media (prefers-color-scheme: dark) {
  body { background: #0d1117; color: #e6edf3; }
  header { background: #161b22; border-bottom: 1px solid #30363d; }
  nav { background: #161b22; border-bottom-color: #30363d; }
  nav button { color: #8d96a0; } nav button:hover { color: #e6edf3; } nav button.active { color: #58a6ff; border-bottom-color: #58a6ff; }
  nav .cnt { background: #30363d; color: #c9d1d9; }
  .card, .panel, .tablewrap { background: #161b22; border-color: #30363d; }
  .card .l, .muted, .toolbar .count, td.empty { color: #8d96a0; }
  .bars .track, .tag { background: #30363d; }
  .toolbar input, .toolbar select, .toolbar button, .pager button { background: #0d1117; border-color: #30363d; }
  th { background: #1c2128; color: #c9d1d9; }
  th, td { border-bottom-color: #21262d; }
  tr:hover td { background: #1c2128; }
  .lv4 { color: #ff7b72; } .lv3 { color: #ffa657; } .lv2 { color: #e3b341; } .lv1 { color: #56d364; }
  .note { background: #2d2400; border-color: #6b5800; color: #f2e3a0; }
}
</style>
</head>
<body>
<header>
  <h1>Audyt uprawnień udziałów sieciowych</h1>
  <div class="meta" id="meta"></div>
</header>
<nav id="tabs"></nav>
<main id="main"></main>
<script type="application/json" id="audit-data">__AUDIT_DATA__</script>
<script>
(function () {
  'use strict';
  var node = document.getElementById('audit-data');
  var D = JSON.parse(node.textContent || node.innerHTML);
  var SEV = { 'Krytyczne': 'crit', 'Wysokie': 'high', '\u015arednie': 'med', 'Niskie': 'low', 'Info': 'info' };
  var SEV_ORDER = ['Krytyczne', 'Wysokie', '\u015arednie', 'Niskie', 'Info'];

  function arr(v) { if (v === null || v === undefined) { return []; } return Object.prototype.toString.call(v) === '[object Array]' ? v : [v]; }
  function esc(v) {
    if (v === null || v === undefined) { return ''; }
    return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function num(n) { return String(n === null || n === undefined ? 0 : n).replace(/\B(?=(\d{3})+(?!\d))/g, ' '); }
  function byId(id) { return document.getElementById(id); }
  function closest(el, tag) { while (el && el.tagName !== tag) { el = el.parentNode; } return el; }

  var COLS = {
    findings: [
      { key: 'Severity', label: 'Waga', type: 'sev', sort: 'SeverityRank' },
      { key: 'Title', label: 'Ustalenie' },
      { key: 'Server', label: 'Serwer' },
      { key: 'Share', label: 'Udzia\u0142' },
      { key: 'Path', label: '\u015acie\u017cka', type: 'path' },
      { key: 'Principal', label: 'Podmiot' },
      { key: 'Access', label: 'Dost\u0119p' },
      { key: 'Details', label: 'Szczeg\u00f3\u0142y', type: 'long' },
      { key: 'Recommendation', label: 'Zalecenie', type: 'long' },
      { key: 'Code', label: 'Kod' }
    ],
    shares: [
      { key: 'Server', label: 'Serwer' },
      { key: 'Share', label: 'Udzia\u0142' },
      { key: 'MaxSeverity', label: 'Najwy\u017csza waga', type: 'sev', sort: 'MaxSeverityRank' },
      { key: 'Findings', label: 'Ustalenia' },
      { key: 'UncPath', label: '\u015acie\u017cka UNC', type: 'path' },
      { key: 'LocalPath', label: '\u015acie\u017cka lokalna', type: 'path' },
      { key: 'Description', label: 'Opis' },
      { key: 'ShareAclStatus', label: 'Uprawnienia udzia\u0142u' },
      { key: 'ShareAces', label: 'Wpisy' },
      { key: 'ABE', label: 'ABE' },
      { key: 'Encryption', label: 'Szyfrowanie SMB' },
      { key: 'Caching', label: 'Offline' },
      { key: 'Hidden', label: 'Ukryty', type: 'bool' },
      { key: 'CurrentUsers', label: 'Sesje' },
      { key: 'FoldersScanned', label: 'Foldery' },
      { key: 'FoldersReported', label: 'Foldery z jawnymi upr.' },
      { key: 'NtfsErrors', label: 'B\u0142\u0119dy NTFS' }
    ],
    shareAcl: [
      { key: 'Server', label: 'Serwer' },
      { key: 'Share', label: 'Udzia\u0142' },
      { key: 'Principal', label: 'Podmiot' },
      { key: 'PrincipalType', label: 'Typ', type: 'tag' },
      { key: 'Scope', label: 'Zakres' },
      { key: 'AccessType', label: 'Rodzaj' },
      { key: 'AccessLevel', label: 'Poziom', type: 'level', rank: 'AccessRank', sort: 'AccessRank' },
      { key: 'Rights', label: 'Prawa' },
      { key: 'AccountEnabled', label: 'Konto w\u0142\u0105czone', type: 'bool' },
      { key: 'Sid', label: 'SID', type: 'path' }
    ],
    ntfsAcl: [
      { key: 'Server', label: 'Serwer' },
      { key: 'Share', label: 'Udzia\u0142' },
      { key: 'Path', label: '\u015acie\u017cka', type: 'path' },
      { key: 'Depth', label: 'Poz.' },
      { key: 'Principal', label: 'Podmiot' },
      { key: 'PrincipalType', label: 'Typ', type: 'tag' },
      { key: 'AccessType', label: 'Rodzaj' },
      { key: 'AccessLevel', label: 'Poziom', type: 'level', rank: 'AccessRank', sort: 'AccessRank' },
      { key: 'Rights', label: 'Prawa' },
      { key: 'AppliesTo', label: 'Dotyczy' },
      { key: 'IsInherited', label: 'Dziedziczone', type: 'bool' },
      { key: 'InheritanceDisabled', label: 'Dziedziczenie wy\u0142.', type: 'bool' },
      { key: 'Owner', label: 'W\u0142a\u015bciciel' },
      { key: 'AccountEnabled', label: 'Konto w\u0142\u0105czone', type: 'bool' }
    ],
    groupMembers: [
      { key: 'Group', label: 'Grupa' },
      { key: 'GroupScope', label: 'Rodzaj grupy' },
      { key: 'Server', label: 'Serwer' },
      { key: 'Member', label: 'Cz\u0142onek' },
      { key: 'DisplayName', label: 'Nazwa wy\u015bwietlana' },
      { key: 'MemberType', label: 'Typ', type: 'tag' },
      { key: 'Enabled', label: 'W\u0142\u0105czone', type: 'bool' },
      { key: 'Source', label: '\u0179r\u00f3d\u0142o' },
      { key: 'Note', label: 'Uwagi' }
    ],
    userAccess: [
      { key: 'User', label: 'U\u017cytkownik' },
      { key: 'Server', label: 'Serwer' },
      { key: 'Share', label: 'Udzia\u0142' },
      { key: 'Path', label: '\u015acie\u017cka', type: 'path' },
      { key: 'EffectiveAccess', label: 'Dost\u0119p efektywny', type: 'level', rank: 'EffectiveRank', sort: 'EffectiveRank' },
      { key: 'ShareAccess', label: 'Udzia\u0142' },
      { key: 'NtfsAccess', label: 'NTFS' },
      { key: 'Via', label: 'Przez' },
      { key: 'DenyPresent', label: 'Deny', type: 'bool' }
    ],
    servers: [
      { key: 'Server', label: 'Serwer' },
      { key: 'Reachable', label: 'Osi\u0105galny', type: 'bool' },
      { key: 'SharesFound', label: 'Udzia\u0142y dyskowe' },
      { key: 'SharesAudited', label: 'Obj\u0119te audytem' },
      { key: 'ShareAclReadable', label: 'Odczyt upr. udzia\u0142\u00f3w', type: 'bool' },
      { key: 'Error', label: 'B\u0142\u0105d' }
    ],
    errors: [
      { key: 'Server', label: 'Serwer' },
      { key: 'Share', label: 'Udzia\u0142' },
      { key: 'Path', label: '\u015acie\u017cka', type: 'path' },
      { key: 'Stage', label: 'Etap' },
      { key: 'Message', label: 'Komunikat', type: 'long' }
    ]
  };

  var TABS = [
    { id: 'summary', label: 'Podsumowanie' },
    { id: 'findings', label: 'Ustalenia', filters: ['Severity', 'Code', 'Server', 'Share'], sort: 'Severity', desc: true },
    { id: 'shares', label: 'Udzia\u0142y', filters: ['Server', 'ShareAclStatus', 'MaxSeverity'], sort: 'MaxSeverity', desc: true },
    { id: 'shareAcl', label: 'Uprawnienia udzia\u0142\u00f3w', filters: ['Server', 'Share', 'PrincipalType', 'AccessType', 'AccessLevel'] },
    { id: 'ntfsAcl', label: 'Uprawnienia NTFS', filters: ['Server', 'Share', 'PrincipalType', 'AccessType', 'AccessLevel'] },
    { id: 'groupMembers', label: 'Cz\u0142onkowie grup', filters: ['Group', 'GroupScope'], optional: true },
    { id: 'userAccess', label: 'Dost\u0119p u\u017cytkownik\u00f3w', filters: ['User', 'Server', 'Share', 'EffectiveAccess'], optional: true, sort: 'EffectiveAccess', desc: true },
    { id: 'servers', label: 'Serwery', filters: [] },
    { id: 'errors', label: 'B\u0142\u0119dy', filters: ['Server', 'Stage'] }
  ];

  function cell(col, row) {
    var v = row[col.key];
    switch (col.type) {
      case 'sev': return v ? '<span class="badge ' + (SEV[v] || 'info') + '">' + esc(v) + '</span>' : '';
      case 'bool': return v === true ? 'Tak' : (v === false ? 'Nie' : '');
      case 'path': return '<span class="mono">' + esc(v) + '</span>';
      case 'long': return '<span class="long">' + esc(v) + '</span>';
      case 'tag': return v ? '<span class="tag">' + esc(v) + '</span>' : '';
      case 'level':
        var r = row[col.rank];
        return '<span class="lv' + (r === null || r === undefined ? 'x' : r) + '">' + esc(v) + '</span>';
      default: return esc(v);
    }
  }
  function sortValue(col, row) { var v = row[col.sort || col.key]; return (v === null || v === undefined) ? '' : v; }
  function compare(a, b) {
    if (typeof a === 'number' && typeof b === 'number') { return a - b; }
    if (typeof a === 'boolean') { a = a ? 1 : 0; }
    if (typeof b === 'boolean') { b = b ? 1 : 0; }
    if (typeof a === 'number' && typeof b === 'number') { return a - b; }
    return String(a).localeCompare(String(b), 'pl', { numeric: true, sensitivity: 'base' });
  }
  function uniqueValues(rows, key) {
    var seen = {}, out = [];
    for (var i = 0; i < rows.length; i++) {
      var v = rows[i][key];
      if (v === null || v === undefined || v === '') { continue; }
      v = String(v);
      if (!seen.hasOwnProperty(v)) { seen[v] = true; out.push(v); }
    }
    if (key === 'Severity' || key === 'MaxSeverity') {
      out.sort(function (a, b) { return SEV_ORDER.indexOf(a) - SEV_ORDER.indexOf(b); });
    } else {
      out.sort(function (a, b) { return compare(a, b); });
    }
    return out;
  }
  function colLabel(cols, key) { for (var i = 0; i < cols.length; i++) { if (cols[i].key === key) { return cols[i].label; } } return key; }
  function csvQuote(v) {
    v = (v === null || v === undefined) ? '' : String(v);
    if (/^[=+\-@]/.test(v)) { v = "'" + v; }
    return '"' + v.replace(/"/g, '""') + '"';
  }

  function DataTable(host, rows, cols, opts) {
    this.host = host; this.rows = arr(rows); this.cols = cols; this.opts = opts || {};
    this.q = ''; this.filters = {}; this.sortCol = null; this.sortDir = 1; this.page = 0; this.size = 50;
    for (var i = 0; i < cols.length; i++) { if (cols[i].key === this.opts.sort) { this.sortCol = cols[i]; this.sortDir = this.opts.desc ? -1 : 1; } }
    this.build();
  }
  DataTable.prototype.build = function () {
    var self = this, h = [], i, j;
    if (this.opts.truncated) {
      h.push('<div class="note">Tabela w raporcie HTML jest obci\u0119ta do ' + num(this.rows.length) + ' z ' + num(this.opts.total) + ' wierszy \u2013 pe\u0142ne dane znajdziesz w plikach CSV/JSON.</div>');
    }
    h.push('<div class="toolbar"><input type="search" class="q" placeholder="Szukaj we wszystkich kolumnach\u2026">');
    var filters = this.opts.filters || [];
    for (i = 0; i < filters.length; i++) {
      var values = uniqueValues(this.rows, filters[i]);
      if (values.length < 2) { continue; }
      h.push('<select class="flt" data-key="' + esc(filters[i]) + '"><option value="">' + esc(colLabel(this.cols, filters[i])) + ': wszystkie</option>');
      for (j = 0; j < values.length; j++) { h.push('<option value="' + esc(values[j]) + '">' + esc(values[j]) + '</option>'); }
      h.push('</select>');
    }
    h.push('<select class="size"><option value="25">25 / str.</option><option value="50" selected>50 / str.</option><option value="100">100 / str.</option><option value="500">500 / str.</option></select>');
    h.push('<button type="button" class="csv">Eksport widoku do CSV</button><span class="count"></span></div>');
    h.push('<div class="tablewrap"></div><div class="pager"></div>');
    this.host.innerHTML = h.join('');
    this.wrap = this.host.querySelector('.tablewrap');
    this.pager = this.host.querySelector('.pager');
    this.counter = this.host.querySelector('.count');

    var q = this.host.querySelector('.q');
    var onSearch = function () { if (self.q !== q.value) { self.q = q.value; self.page = 0; self.render(); } };
    q.addEventListener('input', onSearch);
    q.addEventListener('keyup', onSearch);
    var selects = this.host.querySelectorAll('select.flt');
    for (i = 0; i < selects.length; i++) {
      selects[i].addEventListener('change', function (e) {
        var sel = e.target || e.srcElement;
        self.filters[sel.getAttribute('data-key')] = sel.value; self.page = 0; self.render();
      });
    }
    this.host.querySelector('select.size').addEventListener('change', function (e) {
      self.size = parseInt((e.target || e.srcElement).value, 10) || 50; self.page = 0; self.render();
    });
    this.host.querySelector('button.csv').addEventListener('click', function () { self.exportCsv(); });
    this.wrap.addEventListener('click', function (e) {
      var th = closest(e.target || e.srcElement, 'TH');
      if (!th) { return; }
      var col = self.cols[parseInt(th.getAttribute('data-i'), 10)];
      if (self.sortCol === col) { self.sortDir = -self.sortDir; } else { self.sortCol = col; self.sortDir = 1; }
      self.render();
    });
    this.pager.addEventListener('click', function (e) {
      var b = closest(e.target || e.srcElement, 'BUTTON');
      if (!b || b.disabled) { return; }
      self.page = parseInt(b.getAttribute('data-p'), 10) || 0; self.render();
    });
    this.render();
  };
  DataTable.prototype.setFilter = function (key, value) {
    this.filters[key] = value;
    var sel = this.host.querySelector('select.flt[data-key="' + key + '"]');
    if (sel) { sel.value = value; }
    this.page = 0; this.render();
  };
  DataTable.prototype.view = function () {
    var q = this.q.toLowerCase(), f = this.filters, cols = this.cols, out = [], i, k, c;
    for (i = 0; i < this.rows.length; i++) {
      var r = this.rows[i], ok = true;
      for (k in f) { if (f.hasOwnProperty(k) && f[k] !== '' && String(r[k]) !== f[k]) { ok = false; break; } }
      if (ok && q) {
        ok = false;
        for (c = 0; c < cols.length; c++) {
          var v = r[cols[c].key];
          if (v !== null && v !== undefined && String(v).toLowerCase().indexOf(q) !== -1) { ok = true; break; }
        }
      }
      if (ok) { out.push(r); }
    }
    if (this.sortCol) {
      var col = this.sortCol, dir = this.sortDir;
      out.sort(function (a, b) { return dir * compare(sortValue(col, a), sortValue(col, b)); });
    }
    return out;
  };
  DataTable.prototype.render = function () {
    var rows = this.view(), total = rows.length, pages = Math.max(1, Math.ceil(total / this.size)), h = [], i, c;
    if (this.page >= pages) { this.page = pages - 1; }
    var start = this.page * this.size, end = Math.min(total, start + this.size);
    h.push('<table><thead><tr>');
    for (c = 0; c < this.cols.length; c++) {
      var arrow = this.sortCol === this.cols[c] ? (this.sortDir > 0 ? ' \u25b2' : ' \u25bc') : '';
      h.push('<th data-i="' + c + '">' + esc(this.cols[c].label) + arrow + '</th>');
    }
    h.push('</tr></thead><tbody>');
    for (i = start; i < end; i++) {
      h.push('<tr>');
      for (c = 0; c < this.cols.length; c++) { h.push('<td>' + cell(this.cols[c], rows[i]) + '</td>'); }
      h.push('</tr>');
    }
    if (!total) { h.push('<tr><td class="empty" colspan="' + this.cols.length + '">Brak wierszy do wy\u015bwietlenia</td></tr>'); }
    h.push('</tbody></table>');
    this.wrap.innerHTML = h.join('');
    this.counter.innerHTML = 'Wiersze: ' + num(total) + (total !== this.rows.length ? ' z ' + num(this.rows.length) : '');
    this.pager.innerHTML = pages > 1
      ? '<button type="button" data-p="0"' + (this.page === 0 ? ' disabled' : '') + '>\u00ab</button>' +
        '<button type="button" data-p="' + (this.page - 1) + '"' + (this.page === 0 ? ' disabled' : '') + '>\u2039 Poprzednia</button>' +
        '<span class="muted"> Strona ' + (this.page + 1) + ' z ' + pages + ' </span>' +
        '<button type="button" data-p="' + (this.page + 1) + '"' + (this.page >= pages - 1 ? ' disabled' : '') + '>Nast\u0119pna \u203a</button>' +
        '<button type="button" data-p="' + (pages - 1) + '"' + (this.page >= pages - 1 ? ' disabled' : '') + '>\u00bb</button>'
      : '';
  };
  DataTable.prototype.exportCsv = function () {
    var rows = this.view(), cols = this.cols, lines = [], i, c, line;
    line = [];
    for (c = 0; c < cols.length; c++) { line.push(csvQuote(cols[c].label)); }
    lines.push(line.join(';'));
    for (i = 0; i < rows.length; i++) {
      line = [];
      for (c = 0; c < cols.length; c++) { line.push(csvQuote(rows[i][cols[c].key])); }
      lines.push(line.join(';'));
    }
    var blob = new Blob(['\ufeff' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var name = (this.opts.name || 'eksport') + '.csv';
    if (window.navigator.msSaveOrOpenBlob) { window.navigator.msSaveOrOpenBlob(blob, name); return; }
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.parentNode.removeChild(a); }, 100);
  };

  var S = D.summary || {}, M = D.meta || {}, T = D.truncated || {}, TOTALS = S.totals || {};
  var tables = {};

  function card(value, label, cls, attrs) {
    return '<' + (attrs ? 'a' : 'div') + ' class="card ' + (cls || '') + '"' + (attrs || '') + '><div class="v">' + value + '</div><div class="l">' + esc(label) + '</div></' + (attrs ? 'a' : 'div') + '>';
  }
  function renderSummary(host) {
    var h = [], i, sev = S.severity || {};
    h.push('<h2>Zakres audytu</h2><div class="cards">');
    h.push(card(num(S.serversReachable) + ' / ' + num(S.servers), 'serwery osi\u0105galne / wszystkie'));
    h.push(card(num(S.shares), 'udzia\u0142y obj\u0119te audytem'));
    h.push(card(num(S.foldersScanned), 'foldery przeskanowane (NTFS)'));
    h.push(card(num(S.shareAces), 'wpisy uprawnie\u0144 udzia\u0142\u00f3w'));
    h.push(card(num(S.ntfsAces), 'wpisy NTFS w raporcie'));
    h.push(card(num(S.errors), 'b\u0142\u0119dy odczytu / \u0142\u0105czno\u015bci'));
    h.push('</div><h2>Ustalenia wed\u0142ug wagi</h2><div class="cards">');
    for (i = 0; i < SEV_ORDER.length; i++) {
      h.push(card(num(sev[SEV_ORDER[i]] || 0), SEV_ORDER[i], 'sevcard ' + SEV[SEV_ORDER[i]], ' href="#findings" data-sev="' + esc(SEV_ORDER[i]) + '"'));
    }
    h.push('</div><div class="twocol"><div><h2>Rodzaje ustale\u0144</h2><div class="panel bars">');
    var codes = arr(S.topCodes), max = 1;
    for (i = 0; i < codes.length; i++) { if (codes[i].Count > max) { max = codes[i].Count; } }
    if (!codes.length) { h.push('<div class="muted">Brak ustale\u0144 \u2013 nie wykryto problem\u00f3w.</div>'); }
    for (i = 0; i < codes.length; i++) {
      var cls = SEV[codes[i].Severity] || 'info';
      h.push('<div class="row"><div class="lbl"><span class="badge ' + cls + '">' + esc(codes[i].Severity) + '</span> ' + esc(codes[i].Title) +
        ' <span class="muted">\u2013 ' + num(codes[i].Count) + ' (' + esc(codes[i].Code) + ')</span></div>' +
        '<div class="track"><div class="fill ' + cls + '" style="width:' + Math.max(1, Math.round(100 * codes[i].Count / max)) + '%"></div></div></div>');
    }
    h.push('</div></div><div><h2>Podmioty z najwi\u0119ksz\u0105 liczb\u0105 jawnych wpis\u00f3w NTFS</h2><div class="panel">');
    var tops = arr(S.topPrincipals);
    if (!tops.length) { h.push('<div class="muted">Brak danych NTFS.</div>'); }
    else {
      h.push('<table><thead><tr><th>Podmiot</th><th>Typ</th><th>Jawne wpisy</th></tr></thead><tbody>');
      for (i = 0; i < tops.length; i++) { h.push('<tr><td>' + esc(tops[i].Principal) + '</td><td><span class="tag">' + esc(tops[i].Type) + '</span></td><td>' + num(tops[i].Count) + '</td></tr>'); }
      h.push('</tbody></table>');
    }
    h.push('</div></div></div><h2>Parametry</h2><div class="panel"><dl class="params">');
    var params = arr(M.parameters);
    for (i = 0; i < params.length; i++) { h.push('<dt>' + esc(params[i].Name) + '</dt><dd>' + esc(params[i].Value) + '</dd>'); }
    h.push('</dl></div>');
    host.innerHTML = h.join('');
    var links = host.querySelectorAll('a[data-sev]');
    for (i = 0; i < links.length; i++) {
      links[i].addEventListener('click', function (e) {
        var el = closest(e.target || e.srcElement, 'A');
        if (e.preventDefault) { e.preventDefault(); }
        show('findings');
        if (tables.findings) { tables.findings.setFilter('Severity', el.getAttribute('data-sev')); }
      });
    }
  }

  function show(id) {
    for (var i = 0; i < TABS.length; i++) {
      var t = TABS[i], sec = byId('sec-' + t.id), btn = byId('btn-' + t.id);
      if (!sec) { continue; }
      var on = t.id === id;
      sec.className = on ? 'active' : '';
      btn.className = on ? 'active' : '';
      if (on && !sec.getAttribute('data-ready')) {
        sec.setAttribute('data-ready', '1');
        if (t.id === 'summary') { renderSummary(sec); }
        else {
          tables[t.id] = new DataTable(sec, D[t.id], COLS[t.id], {
            filters: t.filters, sort: t.sort, desc: t.desc, name: t.id,
            truncated: !!T[t.id], total: TOTALS[t.id]
          });
        }
      }
    }
    if (window.history && window.history.replaceState) { window.history.replaceState(null, '', '#' + id); }
  }

  byId('meta').innerHTML = esc('Wygenerowano: ' + (M.generated || '') + ' \u00b7 Audytor: ' + (M.auditor || '') +
    ' \u00b7 Komputer: ' + (M.computer || '') + ' \u00b7 Czas trwania: ' + (M.duration || '') + ' \u00b7 Wersja skryptu: ' + (M.version || ''));
  var nav = [], secs = [];
  for (var i = 0; i < TABS.length; i++) {
    var t = TABS[i];
    if (t.optional && !arr(D[t.id]).length) { continue; }
    var count = t.id === 'summary' ? '' : '<span class="cnt">' + num(TOTALS[t.id] !== undefined ? TOTALS[t.id] : arr(D[t.id]).length) + '</span>';
    nav.push('<button type="button" id="btn-' + t.id + '" data-tab="' + t.id + '">' + esc(t.label) + count + '</button>');
    secs.push('<section id="sec-' + t.id + '"></section>');
  }
  byId('tabs').innerHTML = nav.join('');
  byId('main').innerHTML = secs.join('');
  byId('tabs').addEventListener('click', function (e) {
    var b = closest(e.target || e.srcElement, 'BUTTON');
    if (b) { show(b.getAttribute('data-tab')); }
  });
  var start = (window.location.hash || '').replace('#', '');
  show(byId('sec-' + start) ? start : 'summary');
})();
</script>
</body>
</html>
'@
    $html = $template.Replace('__AUDIT_DATA__', $json)
    [System.IO.File]::WriteAllText($FilePath, $html, (New-Object System.Text.UTF8Encoding($true)))
}
#endregion

#region Eksport
function Get-LimitedRows {
    param($List)
    if ($List.Count -gt $HtmlRowLimit) { return $List.GetRange(0, $HtmlRowLimit).ToArray() }
    return $List.ToArray()
}

function Export-AuditCsv {
    param($Rows, [string]$Name)
    if ($null -eq $Rows -or $Rows.Count -eq 0) { return }
    $file = Join-Path $script:ReportDir ('{0}.csv' -f $Name)
    $encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $Rows | Export-Csv -LiteralPath $file -NoTypeInformation -Delimiter $CsvDelimiter -Encoding $encoding
}
#endregion

#region Główny przebieg
$reportRoot = $OutputPath
if (-not $reportRoot) {
    $reportRoot = [Environment]::GetFolderPath('Desktop')
    if (-not $reportRoot -or -not (Test-Path -LiteralPath $reportRoot)) { $reportRoot = (Get-Location).ProviderPath }
}
$script:ReportDir = Join-Path $reportRoot ('ShareAudit_{0:yyyyMMdd_HHmmss}' -f $script:StartTime)
try {
    $script:ReportDir = (New-Item -ItemType Directory -Path $script:ReportDir -Force -ErrorAction Stop).FullName
} catch {
    throw ('Nie można utworzyć folderu raportu {0}: {1}' -f $script:ReportDir, $_.Exception.Message)
}

Write-Host ''
Write-Host '  Audyt uprawnień udziałów sieciowych' -ForegroundColor Cyan
Write-Host ('  Audytor: {0}\{1} na {2}  |  raport: {3}' -f $env:USERDOMAIN, $env:USERNAME, $env:COMPUTERNAME, $script:ReportDir) -ForegroundColor DarkGray
Write-Host ''

$needAD = -not $SkipADLookup
if ($needAD) {
    try {
        $rootDse = [ADSI]'LDAP://RootDSE'
        $namingContext = [string]$rootDse.Properties['defaultNamingContext'].Value
        if ($namingContext) {
            $script:ADAvailable = $true
            Write-AuditLog ('Active Directory: {0}' -f $namingContext) 'OK'
        }
    } catch { }
    if (-not $script:ADAvailable) {
        Write-AuditLog 'Active Directory niedostępne - pomijam stan kont, rozwijanie grup domenowych i -UserAccess.' 'WARN'
    }
}
if ($PSCmdlet.ParameterSetName -eq 'AD' -and -not $script:ADAvailable) {
    throw 'Parametr -FromAD wymaga dostępu do Active Directory (komputer w domenie, bez -SkipADLookup).'
}

try {
    #region Lista serwerów
    $pathSpecs = New-Object System.Collections.Generic.List[object]
    $serverNames = @()
    switch ($PSCmdlet.ParameterSetName) {
        'AD' {
            Write-AuditLog 'Pobieranie listy serwerów z AD...' 'STEP'
            $serverNames = @(Get-ADServerList)
            Write-AuditLog ('Znaleziono komputerów: {0}' -f $serverNames.Count) 'INFO'
        }
        'Path' {
            foreach ($p in $Path) {
                $raw = $p.Trim()
                if ($raw.StartsWith('\\')) {
                    $clean = $raw.TrimEnd('\')
                    if ($clean -match '^\\\\([^\\]+)\\([^\\]+)(\\.*)?$') {
                        $pathSpecs.Add([pscustomobject]@{ Raw = $raw; Server = $Matches[1]; ShareName = $Matches[2]; Rest = [string]$Matches[3]; Local = $false })
                        $serverNames += $Matches[1]
                    } else {
                        Add-AuditError -Server '' -Path $raw -Stage 'Parametry' -Message 'Niepoprawna ścieżka UNC (oczekiwano \\serwer\udział\...).'
                    }
                } else {
                    $pathSpecs.Add([pscustomobject]@{ Raw = $raw; Server = $env:COMPUTERNAME; ShareName = $null; Rest = ''; Local = $true })
                }
            }
        }
        default { $serverNames = @($ComputerName) }
    }
    $serverNames = @($serverNames | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Sort-Object -Unique)
    #endregion

    #region Dostępność i lista udziałów
    $remoteServers = @($serverNames | Where-Object { -not (Test-IsLocalComputer $_) })
    $reach = @{}
    if ($remoteServers.Count -gt 0) {
        Write-AuditLog ('Sprawdzanie portu 445/TCP na {0} serwerach...' -f $remoteServers.Count) 'STEP'
        $reach = Test-TcpPortBulk -HostNames $remoteServers -Port 445 -TimeoutMs 2000
    }
    $reachable = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $orderedReachable = New-Object System.Collections.Generic.List[string]
    foreach ($srv in $serverNames) {
        if ((Test-IsLocalComputer $srv) -or $reach[$srv]) {
            [void]$reachable.Add($srv)
            $orderedReachable.Add($srv)
        } else {
            Add-AuditError -Server $srv -Stage 'Łączność' -Message 'Port 445/TCP niedostępny (host wyłączony, zapora lub brak wpisu DNS).'
            $script:ServerRows.Add([pscustomobject]@{ Server = $srv; Reachable = $false; SharesFound = 0; SharesAudited = 0; ShareAclReadable = $null; Error = 'Port 445/TCP niedostępny' })
        }
    }

    Write-AuditLog ('Odczyt udziałów z {0} serwerów...' -f $orderedReachable.Count) 'STEP'
    $i = 0
    foreach ($srv in $orderedReachable) {
        $i++
        Write-Progress -Id 1 -Activity 'Odczyt udziałów' -Status $srv -PercentComplete ([int](100 * $i / [Math]::Max(1, $orderedReachable.Count)))
        Connect-AuditIpc -Server $srv
        $info = Get-ServerShareInfo -Server $srv
        $row = [pscustomobject]@{ Server = $srv; Reachable = $true; SharesFound = 0; SharesAudited = 0; ShareAclReadable = $null; Error = '' }
        if ($info.Error) {
            Add-AuditError -Server $srv -Stage 'Lista udziałów' -Message $info.Error
            $row.Error = $info.Error
            $script:ServerRows.Add($row)
            continue
        }
        $row.SharesFound = $info.Shares.Count
        $row.ShareAclReadable = ($info.InfoLevel -eq 502)
        if ($info.InfoLevel -eq 1) {
            Write-AuditLog ('{0}: brak uprawnień administratora - uprawnienia udziałów nie zostaną odczytane.' -f $srv) 'WARN'
        }
        if ($PSCmdlet.ParameterSetName -ne 'Path') {
            foreach ($share in $info.Shares) {
                if (-not (Test-ShareSelected $share)) { continue }
                Update-ShareFlags $share
                $script:Shares.Add($share)
                $row.SharesAudited++
            }
            Write-AuditLog ('{0}: udziały dyskowe {1}, do audytu {2}' -f $srv, $row.SharesFound, $row.SharesAudited) 'INFO'
        }
        $script:ServerRows.Add($row)
    }
    Write-Progress -Id 1 -Activity 'Odczyt udziałów' -Completed
    #endregion

    #region Lokalizacje do skanowania NTFS
    $targets = New-Object System.Collections.Generic.List[object]
    $targetKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $shareKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($share in $script:Shares) { [void]$shareKeys.Add($share.Server + '|' + $share.Name) }

    $addTarget = {
        param($Share, [string]$UncStart, [string]$LocalStart, [bool]$UseLocal)
        if (-not $targetKeys.Add($UncStart)) { return }
        $targets.Add([pscustomobject]@{
            Id         = $targets.Count
            Server     = $Share.Server
            ShareName  = $Share.Name
            Share      = $Share
            UncStart   = $UncStart
            LocalStart = $LocalStart
            UseLocal   = $UseLocal
        })
    }

    if ($PSCmdlet.ParameterSetName -eq 'Path') {
        foreach ($spec in $pathSpecs) {
            if ($spec.Local) {
                try {
                    $full = (Resolve-Path -LiteralPath $spec.Raw -ErrorAction Stop).ProviderPath
                } catch {
                    Add-AuditError -Server $env:COMPUTERNAME -Path $spec.Raw -Stage 'Parametry' -Message 'Ścieżka nie istnieje.'
                    continue
                }
                if (-not (Test-Path -LiteralPath $full -PathType Container)) {
                    Add-AuditError -Server $env:COMPUTERNAME -Path $full -Stage 'Parametry' -Message 'Ścieżka nie jest folderem.'
                    continue
                }
                $share = New-AuditShareObject -Server $env:COMPUTERNAME -Name '-' -UncPath $full -LocalPath $full -AclStatus 'NotShare'
                $script:Shares.Add($share)
                & $addTarget $share $full $full $true
                continue
            }
            if (-not $reachable.Contains($spec.Server)) { continue }
            $info = Get-ServerShareInfo -Server $spec.Server
            $share = $null
            if (-not $info.Error) { $share = $info.Shares | Where-Object { $_.Name -eq $spec.ShareName } | Select-Object -First 1 }
            if (-not $share) {
                $share = New-AuditShareObject -Server $spec.Server -Name $spec.ShareName -UncPath ('\\{0}\{1}' -f $spec.Server, $spec.ShareName) -LocalPath $null -AclStatus 'Unknown'
            } else {
                Update-ShareFlags $share
            }
            if ($shareKeys.Add($share.Server + '|' + $share.Name)) {
                $script:Shares.Add($share)
                foreach ($row in $script:ServerRows) { if ($row.Server -eq $spec.Server) { $row.SharesAudited++ } }
            }
            $rest = $spec.Rest.TrimEnd('\')
            $unc = '\\{0}\{1}{2}' -f $spec.Server, $spec.ShareName, $rest
            $local = $null
            if ($share.LocalPath) { $local = if ($rest) { $share.LocalPath.TrimEnd('\') + $rest } else { $share.LocalPath } }
            & $addTarget $share $unc $local ((Test-IsLocalComputer $spec.Server) -and [bool]$local)
        }
    } else {
        foreach ($share in $script:Shares) {
            & $addTarget $share $share.UncPath $share.LocalPath ((Test-IsLocalComputer $share.Server) -and [bool]$share.LocalPath)
        }
    }
    #endregion

    #region Skanowanie NTFS
    if ($ShareOnly) {
        Write-AuditLog 'Tryb -ShareOnly: pomijam skanowanie NTFS.' 'INFO'
    } elseif ($targets.Count -eq 0) {
        Write-AuditLog 'Brak lokalizacji do skanowania NTFS.' 'WARN'
    } else {
        $depthText = if ($Depth -lt 0) { 'bez limitu' } else { [string]$Depth }
        Write-AuditLog ('Skanowanie NTFS: {0} lokalizacji, głębokość {1}, równolegle {2}, tryb {3}...' -f $targets.Count, $depthText, $ThrottleLimit, $ScanMode) 'STEP'
        $scanResults = Invoke-NtfsScan -Targets $targets.ToArray()
        foreach ($t in $targets) {
            $res = $scanResults[$t.Id]
            $share = $t.Share
            if ($null -eq $res) {
                Add-AuditError -Server $t.Server -Share $t.ShareName -Path $t.UncStart -Stage 'Skanowanie NTFS' -Message 'Brak wyniku skanowania.'
                continue
            }
            if ($res.Warning) { Add-AuditError -Server $t.Server -Share $t.ShareName -Path $t.UncStart -Stage 'WinRM' -Message $res.Warning }
            if ($res.Error) {
                Add-AuditError -Server $t.Server -Share $t.ShareName -Path $t.UncStart -Stage 'Skanowanie NTFS' -Message $res.Error
                $share.NtfsErrors++
                continue
            }
            if ($res.Summary) { $share.FoldersScanned += [int]$res.Summary.Scanned }
            foreach ($rec in $res.Records) {
                $folderPath = if ($res.UsedMode -eq 'UNC') { [string]$rec.Path } else { ConvertTo-AuditUncPath -Path ([string]$rec.Path) -LocalRoot $t.LocalStart -UncRoot $t.UncStart }
                if ($rec.Error) {
                    Add-AuditError -Server $t.Server -Share $t.ShareName -Path $folderPath -Stage 'Odczyt ACL' -Message ([string]$rec.Error) -Quiet
                    $share.NtfsErrors++
                }
                if ($rec.ChildError) {
                    Add-AuditError -Server $t.Server -Share $t.ShareName -Path $folderPath -Stage 'Lista podfolderów' -Message ([string]$rec.ChildError) -Quiet
                    $share.NtfsErrors++
                }
                $script:Folders.Add([pscustomobject]@{
                    Server    = $t.Server
                    ShareName = $t.ShareName
                    Share     = $share
                    Path      = $folderPath
                    Depth     = [int]$rec.Depth
                    Protected = [bool]$rec.Protected
                    Owner     = [string]$rec.Owner
                    Aces      = ConvertFrom-AceText ([string]$rec.Aces)
                    Error     = [string]$rec.Error
                })
            }
        }
        $scannedTotal = 0
        foreach ($share in $script:Shares) { $scannedTotal += $share.FoldersScanned }
        Write-AuditLog ('Przeskanowano folderów: {0}, do raportu: {1}' -f $scannedTotal, $script:Folders.Count) 'OK'
    }
    #endregion

    #region Rozwiązywanie SID-ów
    $pairs = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($share in $script:Shares) { foreach ($ace in $share.Aces) { [void]$pairs.Add($share.Server + '|' + $ace.Sid) } }
    foreach ($folder in $script:Folders) {
        foreach ($ace in $folder.Aces) { [void]$pairs.Add($folder.Server + '|' + $ace.Sid) }
        if ($folder.Owner) { [void]$pairs.Add($folder.Server + '|' + $folder.Owner) }
    }
    Write-AuditLog ('Rozwiązywanie {0} podmiotów zabezpieczeń...' -f $pairs.Count) 'STEP'
    $i = 0
    foreach ($pair in $pairs) {
        $i++
        if ($i % 25 -eq 1) { Write-Progress -Id 1 -Activity 'Rozwiązywanie SID-ów' -Status $pair -PercentComplete ([int](100 * $i / [Math]::Max(1, $pairs.Count))) }
        $sep = $pair.IndexOf('|')
        $srv = $pair.Substring(0, $sep)
        $principal = Resolve-AuditPrincipal -Server $srv -Sid $pair.Substring($sep + 1)
        Register-UsedGroup -Server $srv -Principal $principal
    }
    Write-Progress -Id 1 -Activity 'Rozwiązywanie SID-ów' -Completed
    #endregion

    #region Wiersze raportu i ustalenia
    Write-AuditLog 'Analiza uprawnień i wykrywanie ryzyk...' 'STEP'
    $shareAclRows = New-Object System.Collections.Generic.List[object]
    $ntfsRows = New-Object System.Collections.Generic.List[object]

    foreach ($share in $script:Shares) {
        $share.BroadRank = Get-ShareBroadRank $share
        foreach ($ace in $share.Aces) {
            $p = Resolve-AuditPrincipal -Server $share.Server -Sid $ace.Sid
            $shareAclRows.Add([pscustomobject]@{
                Server         = $share.Server
                Share          = $share.Name
                UncPath        = $share.UncPath
                Principal      = $p.Name
                PrincipalType  = $p.Type
                Scope          = $p.Scope
                AccessType     = $ace.Type
                AccessLevel    = Get-AccessLabel $ace.Mask
                AccessRank     = Get-AccessRank $ace.Mask
                Rights         = Get-RightsText $ace.Mask
                AccountEnabled = $p.Enabled
                Sid            = $p.Sid
            })
        }
        Add-ShareFindings -Share $share
    }

    $n = 0
    foreach ($folder in $script:Folders) {
        $n++
        if ($n % 200 -eq 1) { Write-Progress -Id 1 -Activity 'Analiza NTFS' -Status $folder.Path -PercentComplete ([int](100 * $n / [Math]::Max(1, $script:Folders.Count))) }
        $folder.Share.FoldersReported++
        $ownerName = if ($folder.Owner) { (Resolve-AuditPrincipal -Server $folder.Server -Sid $folder.Owner).Name } else { '' }
        foreach ($ace in $folder.Aces) {
            if ($folder.Depth -gt 0 -and $ace.Inherited -and -not $IncludeInherited) { continue }
            $p = Resolve-AuditPrincipal -Server $folder.Server -Sid $ace.Sid
            $ntfsRows.Add([pscustomobject]@{
                Server              = $folder.Server
                Share               = $folder.ShareName
                Path                = $folder.Path
                Depth               = $folder.Depth
                Principal           = $p.Name
                PrincipalType       = $p.Type
                Scope               = $p.Scope
                AccessType          = $ace.Type
                AccessLevel         = Get-AccessLabel $ace.Mask
                AccessRank          = Get-AccessRank $ace.Mask
                Rights              = Get-RightsText $ace.Mask
                AppliesTo           = Get-AppliesToText $ace.IF $ace.PF
                IsInherited         = [bool]$ace.Inherited
                InheritanceDisabled = $folder.Protected
                Owner               = $ownerName
                AccountEnabled      = $p.Enabled
                Sid                 = $p.Sid
            })
        }
        Add-NtfsFindings -Folder $folder
    }
    Write-Progress -Id 1 -Activity 'Analiza NTFS' -Completed
    #endregion

    #region Grupy i efektywny dostęp
    $memberRows = New-Object System.Collections.Generic.List[object]
    if ($ExpandGroups) {
        Write-AuditLog ('Rozwijanie członkostwa {0} grup...' -f @($script:UsedGroups.Values | Where-Object { -not $_.Principal.IsBroad }).Count) 'STEP'
        $memberRows = Get-GroupMemberRows
    }

    $userAccessRows = New-Object System.Collections.Generic.List[object]
    if ($UserAccess) {
        if (-not $script:ADAvailable) {
            Write-AuditLog '-UserAccess wymaga dostępu do AD - pomijam.' 'WARN'
        } elseif ($script:Folders.Count -eq 0) {
            Write-AuditLog '-UserAccess wymaga danych NTFS (bez -ShareOnly) - pomijam.' 'WARN'
        } else {
            $contexts = New-Object System.Collections.Generic.List[object]
            foreach ($identity in $UserAccess) {
                $ctx = Get-UserSecurityContext -Identity $identity
                if ($ctx.Error) {
                    Add-AuditError -Server '' -Stage 'UserAccess' -Message ('{0}: {1}' -f $identity, $ctx.Error)
                    continue
                }
                Write-AuditLog ('Efektywny dostęp: {0} ({1} SID-ów w tokenie)' -f $ctx.Name, $ctx.Sids.Count) 'STEP'
                $contexts.Add($ctx)
            }
            if ($contexts.Count -gt 0) { $userAccessRows = Get-UserAccessRows -Contexts $contexts.ToArray() }
        }
    }
    #endregion

    #region Podsumowanie i eksport
    $sortedFindings = New-Object System.Collections.Generic.List[object]
    foreach ($f in ($script:Findings | Sort-Object @{ Expression = 'SeverityRank'; Descending = $true }, Server, Share, Path, Principal)) { $sortedFindings.Add($f) }

    $findingsByShare = @{}
    foreach ($f in $sortedFindings) {
        $key = $f.Server + '|' + $f.Share
        if (-not $findingsByShare.ContainsKey($key)) { $findingsByShare[$key] = [pscustomobject]@{ Count = 0; MaxRank = 0; MaxSeverity = '' } }
        $entry = $findingsByShare[$key]
        $entry.Count++
        if ($f.SeverityRank -gt $entry.MaxRank) { $entry.MaxRank = $f.SeverityRank; $entry.MaxSeverity = $f.Severity }
    }

    $serverRowsSorted = New-Object System.Collections.Generic.List[object]
    foreach ($row in ($script:ServerRows | Sort-Object Server)) { $serverRowsSorted.Add($row) }
    $script:ServerRows = $serverRowsSorted

    $shareRows = New-Object System.Collections.Generic.List[object]
    foreach ($share in $script:Shares) {
        $stats = $findingsByShare[$share.Server + '|' + $share.Name]
        $shareRows.Add([pscustomobject]@{
            Server          = $share.Server
            Share           = $share.Name
            UncPath         = $share.UncPath
            LocalPath       = $share.LocalPath
            Description     = $share.Description
            ShareAclStatus  = $script:AclStatusText[$share.AclStatus]
            ShareAces       = @($share.Aces).Count
            ABE             = Get-ShareFlagText $share 'ABE'
            Encryption      = Get-ShareFlagText $share 'Encrypt'
            Caching         = Get-ShareFlagText $share 'Caching'
            DFS             = Get-ShareFlagText $share 'DFS'
            Hidden          = $share.IsHidden
            AdminShare      = $share.IsAdminShare
            CurrentUsers    = $share.CurrentUses
            FoldersScanned  = $share.FoldersScanned
            FoldersReported = $share.FoldersReported
            NtfsErrors      = $share.NtfsErrors
            Findings        = if ($stats) { $stats.Count } else { 0 }
            MaxSeverity     = if ($stats) { $stats.MaxSeverity } else { '' }
            MaxSeverityRank = if ($stats) { $stats.MaxRank } else { 0 }
        })
    }

    $severityCounts = [ordered]@{}
    foreach ($s in $script:SeverityRank.Keys) { $severityCounts[$s] = 0 }
    foreach ($f in $sortedFindings) { $severityCounts[$f.Severity] = $severityCounts[$f.Severity] + 1 }

    $topCodes = @($sortedFindings | Group-Object Code | ForEach-Object {
        $first = $_.Group[0]
        [pscustomobject]@{ Code = $_.Name; Title = $first.Title; Severity = $first.Severity; SeverityRank = $first.SeverityRank; Count = $_.Count }
    } | Sort-Object @{ Expression = 'SeverityRank'; Descending = $true }, @{ Expression = 'Count'; Descending = $true })

    $principalCounts = @{}
    $principalTypes = @{}
    foreach ($r in $ntfsRows) {
        if ($r.IsInherited) { continue }
        $principalCounts[$r.Principal] = 1 + [int]$principalCounts[$r.Principal]
        $principalTypes[$r.Principal] = $r.PrincipalType
    }
    $topPrincipals = @($principalCounts.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 15 | ForEach-Object {
        [pscustomobject]@{ Principal = $_.Key; Type = $principalTypes[$_.Key]; Count = $_.Value }
    })

    $scannedTotal = 0
    foreach ($share in $script:Shares) { $scannedTotal += $share.FoldersScanned }
    $duration = (Get-Date) - $script:StartTime
    $durationText = '{0:00}:{1:00}:{2:00}' -f [int][Math]::Floor($duration.TotalHours), $duration.Minutes, $duration.Seconds

    $parameterRows = New-Object System.Collections.Generic.List[object]
    $scopeText = switch ($PSCmdlet.ParameterSetName) {
        'AD'    { 'Serwery z AD' + $(if ($SearchBase) { ' (' + $SearchBase + ')' } else { '' }) }
        'Path'  { 'Ścieżki: ' + ($Path -join ', ') }
        default { 'Serwery: ' + ($ComputerName -join ', ') }
    }
    $parameterRows.Add([pscustomobject]@{ Name = 'Zakres'; Value = $scopeText })
    $parameterRows.Add([pscustomobject]@{ Name = 'Głębokość NTFS'; Value = $(if ($ShareOnly) { 'pominięto (-ShareOnly)' } elseif ($Depth -lt 0) { 'bez limitu' } else { [string]$Depth }) })
    $parameterRows.Add([pscustomobject]@{ Name = 'Tryb skanowania'; Value = $ScanMode })
    $parameterRows.Add([pscustomobject]@{ Name = 'Udziały'; Value = ('uwzględnij: {0}; pomiń: {1}; administracyjne: {2}' -f ($IncludeShare -join ', '), $(if ($ExcludeShare) { $ExcludeShare -join ', ' } else { '-' }), $(if ($IncludeAdminShares) { 'tak' } else { 'nie' })) })
    $parameterRows.Add([pscustomobject]@{ Name = 'Wpisy dziedziczone'; Value = $(if ($IncludeInherited) { 'pokazywane' } else { 'tylko folder główny' }) })
    $parameterRows.Add([pscustomobject]@{ Name = 'Active Directory'; Value = $(if ($script:ADAvailable) { 'dostępne' } else { 'niedostępne / pominięte' }) })
    if ($TrustedPrincipal) { $parameterRows.Add([pscustomobject]@{ Name = 'Zaufane podmioty'; Value = ($TrustedPrincipal -join ', ') }) }
    if ($UserAccess) { $parameterRows.Add([pscustomobject]@{ Name = 'Efektywny dostęp dla'; Value = ($UserAccess -join ', ') }) }
    $parameterRows.Add([pscustomobject]@{ Name = 'Pliki raportu'; Value = $script:ReportDir })

    $totals = [ordered]@{
        findings     = $sortedFindings.Count
        shares       = $shareRows.Count
        shareAcl     = $shareAclRows.Count
        ntfsAcl      = $ntfsRows.Count
        groupMembers = $memberRows.Count
        userAccess   = $userAccessRows.Count
        servers      = $script:ServerRows.Count
        errors       = $script:ErrorRows.Count
    }

    $csvFiles = [ordered]@{
        'Ustalenia'           = $sortedFindings
        'Udzialy'             = $shareRows
        'UprawnieniaUdzialow' = $shareAclRows
        'UprawnieniaNTFS'     = $ntfsRows
        'CzlonkowieGrup'      = $memberRows
        'DostepUzytkownikow'  = $userAccessRows
        'Serwery'             = $script:ServerRows
        'Bledy'               = $script:ErrorRows
    }

    Write-AuditLog 'Zapisywanie raportu...' 'STEP'
    if ($Format -contains 'Csv') {
        foreach ($name in $csvFiles.Keys) { Export-AuditCsv -Rows $csvFiles[$name] -Name $name }
    }

    if ($Format -contains 'Json') {
        $full = [ordered]@{
            generated    = $script:StartTime.ToString('yyyy-MM-dd HH:mm:ss')
            findings     = $sortedFindings.ToArray()
            shares       = $shareRows.ToArray()
            shareAcl     = $shareAclRows.ToArray()
            ntfsAcl      = $ntfsRows.ToArray()
            groupMembers = $memberRows.ToArray()
            userAccess   = $userAccessRows.ToArray()
            servers      = $script:ServerRows.ToArray()
            errors       = $script:ErrorRows.ToArray()
        }
        $jsonPath = Join-Path $script:ReportDir 'Raport.json'
        [System.IO.File]::WriteAllText($jsonPath, ($full | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($false)))
    }

    if ($Format -contains 'Xlsx') {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel -ErrorAction Stop
            $xlsxPath = Join-Path $script:ReportDir 'Raport.xlsx'
            foreach ($name in $csvFiles.Keys) {
                $rows = $csvFiles[$name]
                if ($null -eq $rows -or $rows.Count -eq 0) { continue }
                $rows.ToArray() | Export-Excel -Path $xlsxPath -WorksheetName $name -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
            }
        } else {
            Write-AuditLog 'Format Xlsx wymaga modułu ImportExcel (Install-Module ImportExcel -Scope CurrentUser) - pomijam.' 'WARN'
        }
    }

    $htmlPath = $null
    if ($Format -contains 'Html') {
        $payload = [ordered]@{
            meta         = [ordered]@{
                generated  = $script:StartTime.ToString('yyyy-MM-dd HH:mm:ss')
                auditor    = '{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME
                computer   = $env:COMPUTERNAME
                duration   = $durationText
                version    = $script:ScriptVersion
                parameters = $parameterRows.ToArray()
            }
            summary      = [ordered]@{
                servers          = $serverNames.Count
                serversReachable = $reachable.Count
                shares           = $shareRows.Count
                foldersScanned   = $scannedTotal
                shareAces        = $shareAclRows.Count
                ntfsAces         = $ntfsRows.Count
                errors           = $script:ErrorRows.Count
                severity         = $severityCounts
                topCodes         = $topCodes
                topPrincipals    = $topPrincipals
                totals           = $totals
            }
            truncated    = [ordered]@{
                findings     = $sortedFindings.Count -gt $HtmlRowLimit
                shareAcl     = $shareAclRows.Count -gt $HtmlRowLimit
                ntfsAcl      = $ntfsRows.Count -gt $HtmlRowLimit
                groupMembers = $memberRows.Count -gt $HtmlRowLimit
                userAccess   = $userAccessRows.Count -gt $HtmlRowLimit
                errors       = $script:ErrorRows.Count -gt $HtmlRowLimit
            }
            findings     = @(Get-LimitedRows $sortedFindings)
            shares       = $shareRows.ToArray()
            shareAcl     = @(Get-LimitedRows $shareAclRows)
            ntfsAcl      = @(Get-LimitedRows $ntfsRows)
            groupMembers = @(Get-LimitedRows $memberRows)
            userAccess   = @(Get-LimitedRows $userAccessRows)
            servers      = $script:ServerRows.ToArray()
            errors       = @(Get-LimitedRows $script:ErrorRows)
        }
        $htmlPath = Join-Path $script:ReportDir 'Raport.html'
        if ($ntfsRows.Count -gt 5000) { Write-AuditLog 'Generowanie raportu HTML (duża liczba wierszy - to może chwilę potrwać)...' 'INFO' }
        New-AuditHtmlReport -Data $payload -FilePath $htmlPath
    }
    #endregion

    #region Wynik w konsoli
    Write-Host ''
    Write-Host '  === Podsumowanie audytu ===' -ForegroundColor Cyan
    Write-Host ('  Serwery: {0} (osiągalne: {1})  |  udziały: {2}  |  foldery: {3}  |  wpisy udziałów/NTFS: {4}/{5}  |  błędy: {6}' -f `
            $serverNames.Count, $reachable.Count, $shareRows.Count, $scannedTotal, $shareAclRows.Count, $ntfsRows.Count, $script:ErrorRows.Count)
    Write-Host '  Ustalenia: ' -NoNewline
    foreach ($s in $severityCounts.Keys) { Write-Host ('{0}: {1}   ' -f $s, $severityCounts[$s]) -NoNewline -ForegroundColor $script:SeverityColor[$s] }
    Write-Host ''
    $top = @($sortedFindings | Where-Object { $_.SeverityRank -ge 3 } | Select-Object -First 10)
    if ($top.Count -gt 0) {
        Write-Host ''
        Write-Host '  Najważniejsze ustalenia:' -ForegroundColor Cyan
        foreach ($f in $top) {
            Write-Host ('  [{0}] ' -f $f.Severity) -NoNewline -ForegroundColor $script:SeverityColor[$f.Severity]
            Write-Host ('{0} - {1}{2}' -f $f.Title, $f.Path, $(if ($f.Principal) { ' (' + $f.Principal + ')' } else { '' }))
        }
    }
    Write-Host ''
    Write-AuditLog ('Raport zapisany w: {0} (czas: {1})' -f $script:ReportDir, $durationText) 'OK'
    #endregion
} finally {
    Disconnect-AuditIpc
    Write-Progress -Id 1 -Activity 'Audyt' -Completed
    if ($script:ReportDir -and (Test-Path -LiteralPath $script:ReportDir)) {
        try { [System.IO.File]::WriteAllLines((Join-Path $script:ReportDir 'Audyt.log'), $script:LogLines.ToArray(), (New-Object System.Text.UTF8Encoding($true))) } catch { }
    }
}

if ($htmlPath -and -not $NoOpen) {
    try { Invoke-Item -LiteralPath $htmlPath } catch { }
}

if ($PassThru) {
    [pscustomobject]@{
        ReportDirectory = $script:ReportDir
        Findings        = $sortedFindings.ToArray()
        Shares          = $shareRows.ToArray()
        ShareAcl        = $shareAclRows.ToArray()
        NtfsAcl         = $ntfsRows.ToArray()
        GroupMembers    = $memberRows.ToArray()
        UserAccess      = $userAccessRows.ToArray()
        Servers         = $script:ServerRows.ToArray()
        Errors          = $script:ErrorRows.ToArray()
    }
}
#endregion
