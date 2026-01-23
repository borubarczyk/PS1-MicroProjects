Import-Module ActiveDirectory

# === USTAWIENIA ===
$CsvPath    = "C:\Temp\users.csv"
$LogPath    = "C:\Temp\update_dept_title.log"
$WhatIfMode = $false   # ustaw na $false żeby wykonać zmiany

"=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') START ===" | Out-File -FilePath $LogPath -Append -Encoding utf8

$rows = Import-Csv -Path $CsvPath

foreach ($r in $rows) {
    $sam   = $r.samAccountName
    $newDp = $r.department
    $newTi = $r.title

    if ([string]::IsNullOrWhiteSpace($sam)) {
        "SKIP: brak samAccountName w wierszu: $($r | ConvertTo-Json -Compress)" | Out-File $LogPath -Append -Encoding utf8
        continue
    }

    try {
        $u = Get-ADUser -Identity $sam -Properties Department, Title -ErrorAction Stop

        $curDp = $u.Department
        $curTi = $u.Title

        # jeśli CSV ma puste pola, to nie ruszamy tych atrybutów
        $dpChange = -not [string]::IsNullOrWhiteSpace($newDp) -and ($curDp -ne $newDp)
        $tiChange = -not [string]::IsNullOrWhiteSpace($newTi) -and ($curTi -ne $newTi)

        if (-not ($dpChange -or $tiChange)) {
            "OK: $sam bez zmian (Department='$curDp', Title='$curTi')" | Out-File $LogPath -Append -Encoding utf8
            continue
        }

        # przygotuj Replace tylko dla pól które mają się zmienić
        $replace = @{}
        if ($dpChange) { $replace["Department"] = $newDp }
        if ($tiChange) { $replace["Title"]      = $newTi }

        # LOG: co jest teraz i co będzie
        $logLine = "CHANGE: $sam | Department: '$curDp' -> '" + ($(if($dpChange){$newDp}else{$curDp})) + "' | Title: '$curTi' -> '" + ($(if($tiChange){$newTi}else{$curTi})) + "'"
        if ($WhatIfMode) {
            "WHATIF: $logLine" | Out-File $LogPath -Append -Encoding utf8
        } else {
            Set-ADUser -Identity $u.DistinguishedName -Replace $replace -ErrorAction Stop
            $logLine | Out-File $LogPath -Append -Encoding utf8
        }
    }
    catch {
        "ERROR: $sam - $($_.Exception.Message)" | Out-File $LogPath -Append -Encoding utf8
        continue
    }
}

"=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') END ===" | Out-File -FilePath $LogPath -Append -Encoding utf8

Write-Host "Gotowe. Log: $LogPath"
if ($WhatIfMode) { Write-Host "To był tryb testowy. Ustaw `$WhatIfMode = `$false żeby wykonać zmiany." }
