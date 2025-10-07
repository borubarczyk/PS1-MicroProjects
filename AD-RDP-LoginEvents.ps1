#Requires -Version 5.1
<#
 .SYNOPSIS
  Raport logowań RDP (4624/4634) z korelacją czasu trwania sesji.

 .PARAMETER DaysBack
  Ile dni wstecz analizować (domyślnie 7).

 .PARAMETER OutputDir
  Katalog na pliki wyjściowe (domyślnie C:\Logs\RDP).

 .NOTES
  - Szuka 4624 (LogonType=10 – RemoteInteractive) i paruje z 4634 (logoff) po LogonId.
  - Gdy brak 4634 (np. restart), Duration = null, Status="Open/Unknown".
#>

param(
    [int]$DaysBack = 7,
    [string]$OutputDir = "C:\Logs\RDP"
)

# --- Ustawienia podstawowe ---
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
$startTime = (Get-Date).Date.AddDays(-[math]::Abs($DaysBack))
$now = Get-Date

# --- Helper: bezpieczne pobieranie pola EventData z nazwą ---
function Get-ED {
    param($EventRecord, [string]$Name)
    try {
        return ($EventRecord.Properties |
            ForEach-Object { $_ } |
            Select-Object -ExpandProperty Value -ErrorAction SilentlyContinue) | Out-Null
    } catch {}

    # Fallback – idziemy po XML
    try {
        $xml = [xml]$EventRecord.ToXml()
        return $xml.Event.EventData.Data | Where-Object { $_.Name -eq $Name } | Select-Object -ExpandProperty '#text' -ErrorAction SilentlyContinue
    } catch {
        return $null
    }
}

# --- Szybki parser EventData po nazwie (stabilniej niż po indeksach) ---
function Parse-Event4624 {
    param($e)
    $xml = [xml]$e.ToXml()
    $ed = @{}
    foreach ($d in $xml.Event.EventData.Data) { $ed[$d.Name] = $d.'#text' }
    [pscustomobject]@{
        TimeCreated    = $e.TimeCreated
        TargetUser     = "$($ed['TargetDomainName'])\$($ed['TargetUserName'])".TrimStart('\')
        TargetDomain   = $ed['TargetDomainName']
        UserName       = $ed['TargetUserName']
        LogonType      = [int]$ed['LogonType']
        IpAddress      = if ($ed['IpAddress'] -and $ed['IpAddress'] -ne '-') { $ed['IpAddress'] } else { $null }
        Workstation    = $ed['WorkstationName']
        LogonId        = $ed['TargetLogonId']
        ProcessName    = $ed['ProcessName']
        Sid            = $ed['TargetUserSid']
        Computer       = $e.MachineName
        EventRecordId  = $e.RecordId
        EventId        = 4624
    }
}

function Parse-Event4634 {
    param($e)
    $xml = [xml]$e.ToXml()
    $ed = @{}
    foreach ($d in $xml.Event.EventData.Data) { $ed[$d.Name] = $d.'#text' }
    [pscustomobject]@{
        TimeCreated    = $e.TimeCreated
        LogonId        = $ed['TargetLogonId']
        TargetUser     = "$($ed['TargetDomainName'])\$($ed['TargetUserName'])".TrimStart('\')
        Computer       = $e.MachineName
        EventRecordId  = $e.RecordId
        EventId        = 4634
    }
}

Write-Host "[INFO] Zbieram zdarzenia od $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))..."

# Tylko Security – to najpewniejsze i wszędzie dostępne
$raw4624 = Get-WinEvent -FilterHashtable @{ LogName='Security'; Id=4624; StartTime=$startTime; EndTime=$now } -ErrorAction Stop
$raw4634 = Get-WinEvent -FilterHashtable @{ LogName='Security'; Id=4634; StartTime=$startTime; EndTime=$now } -ErrorAction Stop

# Parsujemy i filtrujemy tylko RDP (LogonType = 10 -> RemoteInteractive)
$logons = foreach ($e in $raw4624) {
    $obj = Parse-Event4624 $e
    if ($obj.LogonType -eq 10) { $obj }
}

# Logoffy do korelacji
$logoffs = foreach ($e in $raw4634) { Parse-Event4634 $e }

# Indeksujemy logoffy po LogonId dla szybkiego dopasowania (może być wiele – bierzemy najbliższy po logon)
$logoffByLogonId = $logoffs | Group-Object LogonId -AsHashTable -AsString

# Korelacja 4624 -> 4634
$result = foreach ($l in $logons | Sort-Object TimeCreated) {
    $logoffTime = $null
    $status = "Open/Unknown"

    if ($l.LogonId -and $logoffByLogonId.ContainsKey($l.LogonId)) {
        $candidates = $logoffByLogonId[$l.LogonId] | Where-Object { $_.TimeCreated -ge $l.TimeCreated } | Sort-Object TimeCreated
        if ($candidates) {
            $logoffTime = $candidates[0].TimeCreated
            $status = "Closed"
        }
    }

    $duration = if ($logoffTime) { [timespan]::FromSeconds( [int]($logoffTime - $l.TimeCreated).TotalSeconds ) } else { $null }

    [pscustomobject]@{
        Server           = $l.Computer
        User             = $l.TargetUser
        UserName         = $l.UserName
        Domain           = $l.TargetDomain
        ClientIP         = $l.IpAddress
        Workstation      = $l.Workstation
        LogonTime        = $l.TimeCreated
        LogoffTime       = $logoffTime
        Duration_hhmmss  = if ($duration) { $duration.ToString() } else { $null }
        Duration_Min     = if ($duration) { [math]::Round($duration.TotalMinutes,1) } else { $null }
        LogonId          = $l.LogonId
        LogonRecordId    = $l.EventRecordId
        LogoffRecordId   = if ($logoffTime) { $candidates[0].EventRecordId } else { $null }
        Status           = $status
        LogonType        = $l.LogonType
        ProcessName      = $l.ProcessName
        SID              = $l.Sid
    }
}

# Porządkujemy, dokładamy sygnaturę i zapisujemy
$stamp = (Get-Date -Format 'yyyy-MM-dd_HHmmss')
$csvPath  = Join-Path $OutputDir "RDP_Logons_${stamp}.csv"
$result | Sort-Object LogonTime | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "[OK] Zapisano CSV: $csvPath"

# Opcjonalnie: XLSX jeśli jest Excel
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)

    # Nagłówki
    $cols = @(
        'Server','User','UserName','Domain','ClientIP','Workstation',
        'LogonTime','LogoffTime','Duration_hhmmss','Duration_Min',
        'LogonId','LogonRecordId','LogoffRecordId','Status','LogonType','ProcessName','SID'
    )
    for ($i=0; $i -lt $cols.Count; $i++) { $ws.Cells.Item(1,$i+1) = $cols[$i] }

    # Dane
    $r = 2
    foreach ($row in $result | Sort-Object LogonTime) {
        $c = 1
        foreach ($k in $cols) {
            $ws.Cells.Item($r,$c) = $row.$k
            $c++
        }
        $r++
    }

    $ws.Range("A1:$([char](64 + $cols.Count))1").Font.Bold = $true
    $ws.Columns.AutoFit() | Out-Null

    $xlsxPath = Join-Path $OutputDir "RDP_Logons_${stamp}.xlsx"
    $wb.SaveAs($xlsxPath, 51)  # 51 = xlOpenXMLWorkbook
    $wb.Close($true)
    $excel.Quit()

    # cleanup COM
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()

    Write-Host "[OK] Zapisano XLSX: $xlsxPath"
} catch {
    Write-Warning "Excel COM niedostępny – pomijam XLSX. Błąd: $($_.Exception.Message)"
}

# Opcjonalna rotacja: zostaw ostatnie 60 dni CSV/XLSX, resztę kasuj
$keepDays = 60
Get-ChildItem $OutputDir -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$keepDays) } | Remove-Item -Force -ErrorAction SilentlyContinue
