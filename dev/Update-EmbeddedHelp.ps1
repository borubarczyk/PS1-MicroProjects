<#
.SYNOPSIS
    Osadza dokumentację z docs/*.md w AD-ManagerDiamond.ps1 – pomoc «Jak to działa» w modułach (przycisk i F1).
.DESCRIPTION
    Uruchom po każdej zmianie plików w folderze docs. Skrypt czyta:
      - docs/0N-*.md: każda sekcja «### Nazwa modułu» trafia pod klucz '<przestrzeń>|<nazwa modułu>',
        a wstęp przestrzeni (od początku pliku do «Spis modułów») pod '<przestrzeń>|*';
      - docs/README.md: rozdział «## Podstawy pracy z programem» pod kluczem '*'.
    Wynik zastępuje region «Pomoc modułów – generowane…» w AD-ManagerDiamond.ps1 (gdy go nie ma – wstawia go przed
    regionem «Uruchomienie»). Plik programu zostaje zapisany jako UTF-8 z BOM. Moduły bez opisu i opisy bez modułu
    są wypisywane jako ostrzeżenia.
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File .\dev\Update-EmbeddedHelp.ps1
#>
[CmdletBinding()]
param([string]$Root = (Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference = 'Stop'

$workspaces = [ordered]@{ '01' = 'Remote'; '02' = 'AdUsers'; '03' = 'AdGroups'; '04' = 'AdComputers'; '05' = 'Files'; '06' = 'Domain' }
$utf8 = New-Object System.Text.UTF8Encoding($true)
$help = [ordered]@{}

function Get-DocLines([string]$Path) {
    return ([System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) -replace "`r", '') -split "`n"
}

function Add-HelpEntry([string]$Key, [object[]]$Lines) {
    $l = New-Object System.Collections.ArrayList
    foreach ($x in $Lines) { if ([string]$x -ne '---') { [void]$l.Add([string]$x) } }
    while ($l.Count -and -not $l[0].Trim()) { $l.RemoveAt(0) }
    while ($l.Count -and -not $l[$l.Count - 1].Trim()) { $l.RemoveAt($l.Count - 1) }
    if ($l.Count) { $script:help[$Key] = ($l -join "`n") }
}

foreach ($file in @(Get-ChildItem -LiteralPath (Join-Path $Root 'docs') -Filter '0*.md' | Sort-Object Name)) {
    $ws = $workspaces[$file.Name.Substring(0, 2)]
    if (-not $ws) { continue }
    $intro = New-Object System.Collections.ArrayList
    $inIntro = $true
    $current = $null
    $buffer = $null
    foreach ($line in (Get-DocLines $file.FullName)) {
        if ($line -match '^# ' -or $line -match '^\[← Spis treści\]') { continue }
        if ($inIntro) {
            if ($line -match '^\*\*Spis modułów\*\*') { $inIntro = $false; Add-HelpEntry "$ws|*" $intro }
            else { [void]$intro.Add($line) }
            continue
        }
        if ($line -match '^### (.+)$') {
            if ($current) { Add-HelpEntry "$ws|$current" $buffer }
            $current = $Matches[1].Trim()
            $buffer = New-Object System.Collections.ArrayList
            continue
        }
        if ($line -match '^## ') {
            if ($current) { Add-HelpEntry "$ws|$current" $buffer }
            $current = $null
            continue
        }
        if ($current) { [void]$buffer.Add($line) }
    }
    if ($current) { Add-HelpEntry "$ws|$current" $buffer }
}

$basics = New-Object System.Collections.ArrayList
$inBasics = $false
foreach ($line in (Get-DocLines (Join-Path (Join-Path $Root 'docs') 'README.md'))) {
    if ($line -match '^## Podstawy pracy z programem') { $inBasics = $true; continue }
    if ($inBasics) { [void]$basics.Add($line) }
}
Add-HelpEntry '*' $basics

# Region do wstawienia w programie (here-string bez rozwijania: treść wstawiana dosłownie)
$sb = New-Object System.Text.StringBuilder
$begin = '#region Pomoc modułów – generowane z docs/*.md przez dev/Update-EmbeddedHelp.ps1 (nie edytuj ręcznie)'
[void]$sb.Append($begin + "`n")
[void]$sb.Append('$script:HelpTexts = @{' + "`n")
foreach ($k in $help.Keys) {
    $text = [string]$help[$k]
    if ($text -match "(?m)^['‘’‚‛]@") { throw "Sekcja «$k»: wiersz zaczynający się od '@ zamknąłby tekst w skrypcie." }
    [void]$sb.Append(("'{0}' = @'" -f $k.Replace("'", "''")) + "`n")
    [void]$sb.Append($text + "`n")
    [void]$sb.Append("'@`n")
}
[void]$sb.Append("}`n#endregion")
$region = $sb.ToString()

$scriptPath = Join-Path $Root 'AD-ManagerDiamond.ps1'
$program = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$newline = if ($program.Contains("`r`n")) { "`r`n" } else { "`n" }
$program = $program -replace "`r`n", "`n"
$start = $program.IndexOf($begin)
if ($start -ge 0) {
    $end = $program.IndexOf('#endregion', $start)
    $program = $program.Substring(0, $start) + $region + $program.Substring($end + '#endregion'.Length)
}
else {
    $anchor = $program.IndexOf("#region Uruchomienie`n")
    if ($anchor -lt 0) { throw 'Nie znaleziono regionu «Uruchomienie» w AD-ManagerDiamond.ps1.' }
    $program = $program.Substring(0, $anchor) + $region + "`n`n" + $program.Substring($anchor)
}

# Kontrola: moduły programu bez opisu i opisy bez modułu
$modules = @([regex]::Matches($program, "(?m)^Register-Module -Workspace '(\w+)' -Category '[^']+' -Key '\w+' -Title '([^']+)'") | ForEach-Object { '{0}|{1}' -f $_.Groups[1].Value, $_.Groups[2].Value })
foreach ($gm in [regex]::Matches($program, "(?m)^Register-GroupMembershipModule -Workspace '(\w+)'.*$")) {
    $title = if ($gm.Value -match "-Title '([^']+)'") { $Matches[1] } else { 'Członkostwo w grupach' }
    $modules += ('{0}|{1}' -f $gm.Groups[1].Value, $title)
}
foreach ($m in $modules) { if (-not $help.Contains($m)) { Write-Warning "Moduł bez opisu w docs: $m" } }
foreach ($k in $help.Keys) { if ($k -ne '*' -and $k -notlike '*|`*' -and $modules -notcontains $k) { Write-Warning "Opis bez modułu w programie: $k" } }

[System.IO.File]::WriteAllText($scriptPath, ($program -replace "`n", $newline), $utf8)
'Osadzono {0} opisów ({1} modułów programu) w {2}.' -f $help.Count, $modules.Count, $scriptPath
