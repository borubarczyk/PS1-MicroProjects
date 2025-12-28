
# Skrypt do generowania raportu uprawnień do folderów i plików
# Napisał: Borys Kaleta

$Exceptions = @(
    "NT AUTHORITY\SYSTEM",
    "CREATOR OWNER",
    "BUILTIN\Administrators",
    "NT SERVICE\TrustedInstaller",
    "APPLICATION PACKAGE AUTHORITY\ALL APPLICATION PACKAGES",
    "APPLICATION PACKAGE AUTHORITY\ALL RESTRICTED APPLICATION PACKAGES",
    "NT VIRTUAL MACHINE\Virtual Machines",
    "NT AUTHORITY\LOCAL SERVICE",
    "NT AUTHORITY\NETWORK SERVICE",
    "AD\Domain Admins"
)
$Header = '"Sciezka";"Rodzaj uprawnien";"Typ Uprawnien";"Nazwa grupy/uzytkownika";"Dziedziczenie";"Flagi dziedziczenia";"Flagi propagacji"'

function Get-RandomID($length) {
    $characters = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.Length }
    $randomString = ""
    foreach ($index in $random) {
        $randomString += $characters[$index]
    }
    return $randomString
}

$ID = Get-RandomID 4
$date = Get-Date -Format "dd_MM_yyyy_HH_mm"
$FileFormat = "\" + $ID + "_ACL_Report_$env:computername-$date.csv"
$CleandFileFormat = "\" + $ID + "_ACL_Report_Clean_$env:computername-$date.csv"

Function Get-FolderName($Description) {
    $initialDirectory = [System.Environment+SpecialFolder]::MyComputer
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog

    $Topmost = New-Object System.Windows.Forms.Form
    $Topmost.TopMost = $True
    $Topmost.MinimizeBox = $True

    $OpenFolderDialog.Description = $Description
    $OpenFolderDialog.RootFolder = $initialDirectory
    $OpenFolderDialog.ShowDialog($Topmost) | Out-Null
    if ($OpenFolderDialog.SelectedPath -eq "") {
        return $null
    }
    else {
        return $OpenFolderDialog.SelectedPath
    }
}


function Get-AclPremmisionReport($Path, $ReportPath) {
    try {
        Write-Progress -Activity "Trwa generowanie raportu..." -Status "Czekaj" -PercentComplete -1
        Get-ChildItem -Recurse $path | Where-Object { $_.PsIsContainer } | ForEach-Object { $path1 = $_.fullname; Get-Acl $_.Fullname | ForEach-Object { $_.access | Add-Member -MemberType NoteProperty 'Path' -Value $path1 -passthru } } | Export-Csv -Path $ReportPath -Encoding UTF8 -NoTypeInformation -Delimiter ";"
        Write-Progress -Activity "Trwa generowanie raportu..." -Completed
    }
    catch {
        Write-Host "An error occurred while generating the report: $_" -ForegroundColor Red
    }
}

function Get-CleanAclPremmisionReport($ReportLocation, $NewReportLocation) {
    try {
        Write-Host "Czyszczenie raportu z wyjatkow" -ForegroundColor Green
        Import-Csv $ReportLocation -Delimiter ";" -Encoding UTF8  | Where-Object { ($Exceptions -notcontains $_.IdentityReference) -and (-not $_.IdentityReference.StartsWith("S-1-")) } | Export-Csv -Path $NewReportLocation -Encoding UTF8 -NoTypeInformation -Delimiter ";" 
        Write-Host "Raport zostal wyczyszczony pomyslnie" -ForegroundColor Green
    }
    catch {
        Write-Host "An error occurred while cleaning the report: $_" -ForegroundColor Red
    }
}

function Set-Headers($ReportPath, $Header){
    $content = Get-Content $ReportPath
    $content[0] = $Header
    $content | Set-Content $ReportPath
}

function Get-PremissionReport() {
    $SaveReportPath = Get-FolderName("Wybierz lokalizacje zapisania raportu")
    $ReportPath = Get-FolderName("Wybierz lokalizacje do weryfikacji")

    if ($null -ne $SaveReportPath) {
        if ($null -ne $ReportPath) {
            $CleandReportPath = $SaveReportPath + $CleandFileFormat
            $SaveReportPath = $SaveReportPath + $FileFormat

            Get-AclPremmisionReport $ReportPath $SaveReportPath
            Write-Host "Lokalizacja raportu: $SaveReportPath" -ForegroundColor Green
            Get-CleanAclPremmisionReport $SaveReportPath $CleandReportPath
            Write-Host "Lokalizacja wyczyszczonego raportu: $CleandReportPath" -ForegroundColor Green
            Set-Headers $CleandReportPath $Header
        }
        else {
            Write-Host "Nie wybrano lokalizacji do weryfikacji!" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Nie wybrano lokazlizacji do zapisu pliku!" -ForegroundColor Red
    }
}

Get-PremissionReport
#Read-Host "Nacisnij Enter aby zakonczyc..."

