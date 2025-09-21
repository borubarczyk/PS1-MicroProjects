<#  Usuwa użytkowników WYŁĄCZONYCH z niekrytycznych grup AD.
    - Domyślnie tylko podgląd (-WhatIf). Usuń -WhatIf żeby wykonać.
    - Można użyć wyboru w GUI (Out-GridView), jeśli jest dostępny.
#>

param(
    [switch]$GuiPick = $true,       # ustaw $false na serwerach bez GUI
    [string]$LogPath = "$env:TEMP\DisabledUsersGroupsBackup.csv"
)

# --- wyjątki (nazwy użytkowników), porównanie case-insensitive
$ExceptionUsers = @(
    'Guest','krbtgt','Gość','Admin2','Konto domyślne','any connect'
)

# --- grupy, z których NIGDY nie usuwamy
$ProtectedGroups = @(
    'Domain Users','Administrators','Domain Admins','Enterprise Admins',
    'Schema Admins','Protected Users','DnsAdmins','Backup Operators',
    'Account Operators','Server Operators','Print Operators',
    'Read-only Domain Controllers'
)

# --- pobierz wyłączonych użytkowników z członkostwami
$users = Get-ADUser -Filter 'Enabled -eq $false' -Properties MemberOf,Name,SamAccountName,DistinguishedName |
    Where-Object {
        $_.MemberOf -and
        ($ExceptionUsers -notcontains $_.Name) -and
        ($ExceptionUsers -notcontains $_.SamAccountName)
    } |
    Select-Object Name,SamAccountName,DistinguishedName,MemberOf

if(-not $users){
    Write-Host 'Brak pasujących użytkowników.' -ForegroundColor Yellow
    return
}

# --- opcjonalny wybór w GUI
if($GuiPick -and (Get-Command Out-GridView -ErrorAction SilentlyContinue)){
    $users = $users | Out-GridView -OutputMode Multiple -Title 'Wybierz użytkowników do usunięcia z grup'
    if(-not $users){ Write-Host 'Nic nie wybrano.'; return }
}

# --- przygotuj kopię członkostw (backup/log)
$backup = foreach($u in $users){
    foreach($gDN in $u.MemberOf){
        # zamień DN na nazwę/przyjazne pola (o ile możliwe)
        $g = Get-ADGroup -Identity $gDN -ErrorAction SilentlyContinue
        [pscustomobject]@{
            TimeStamp        = (Get-Date).ToString('s')
            UserName         = $u.Name
            SamAccountName   = $u.SamAccountName
            UserDN           = $u.DistinguishedName
            GroupDN          = $gDN
            GroupName        = $g.Name
            GroupSam         = $g.SamAccountName
        }
    }
}
$backup | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $LogPath
Write-Host "Backup zapisany: $LogPath" -ForegroundColor Cyan

# --- zbuduj zestaw nazw grup chronionych (porównanie case-insensitive)
$protSet = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$ProtectedGroups | ForEach-Object { [void]$protSet.Add($_) }

# --- usuwanie z pominięciem grup krytycznych i Primary Group
foreach($u in $users){
    Write-Host "Przetwarzam: $($u.Name) [$($u.SamAccountName)]" -ForegroundColor Green

    # Primary Group (np. Domain Users) nie siedzi w MemberOf, ale i tak pilnujemy listy chronionej
    foreach($gDN in $u.MemberOf){
        $g = Get-ADGroup -Identity $gDN -ErrorAction SilentlyContinue
        $gName = if($g){ $g.Name } else { $null }

        # pomiń, jeśli to grupa chroniona (po Name) albo nie udało się pobrać
        if($gName -and $protSet.Contains($gName)){
            Write-Host "  Pomijam grupę chronioną: $gName" -ForegroundColor Yellow
            continue
        }

        Write-Host "  Usuwam z grupy: $($gName ?? $gDN)"
        try{
            # -WhatIf chroni przed faktycznym usunięciem, usuń -WhatIf po weryfikacji
            Remove-ADGroupMember -Identity $gDN -Members $u.DistinguishedName -Confirm:$false -ErrorAction Stop -WhatIf
        }
        catch{
            Write-Warning "  Błąd usuwania z $($gName ?? $gDN): $($_.Exception.Message)"
        }
    }
}
