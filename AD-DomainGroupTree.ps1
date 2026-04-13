function Get-ADDomainGroupTree {
    [CmdletBinding()]
    param (
        [Parameter(HelpMessage="Pokaż użytkowników i komputery (bez tego pokaże TYLKO relacje grupa-grupa, co jest 100x szybsze)")]
        [switch]$ShowMembers,

        [Parameter(Mandatory=$true, HelpMessage="Ścieżka zapisu, np. C:\temp\CaleAD.html")]
        [string]$ExportHtmlPath,

        [Parameter(HelpMessage="Pobiera opisy (Description). UWAGA: BARDZO spowalnia skrypt przy całym AD!")]
        [switch]$IncludeDescription
    )

    if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
        Write-Warning "Moduł ActiveDirectory nie jest zainstalowany."
        return
    }
    Import-Module ActiveDirectory

    $script:visitedGroups = @()
    $script:htmlContent = @()

    Write-Host "🔍 Pobieranie wszystkich grup z Active Directory (To może potrwać)..." -ForegroundColor Cyan
    # Szukamy grup, które nie mają rodziców (Top-Level)
    $allGroups = Get-ADGroup -Filter * -Properties MemberOf, Description
    $topLevelGroups = $allGroups | Where-Object { $_.MemberOf.Count -eq 0 } | Sort-Object Name
    
    $totalTopGroups = $topLevelGroups.Count
    Write-Host "✅ Znaleziono $totalTopGroups grup najwyższego poziomu (korzeni)." -ForegroundColor Green

    # Start HTML
    $script:htmlContent += "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Raport Całego AD</title>"
    $script:htmlContent += "<style>"
    $script:htmlContent += "body { font-family: 'Segoe UI', sans-serif; background: #fff; padding: 20px; color: #333; }"
    $script:htmlContent += ".controls { position: sticky; top: 0; margin-bottom: 20px; padding: 15px; background: #f8f9fa; border-radius: 5px; border: 1px solid #ddd; z-index: 100; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }"
    $script:htmlContent += ".search-box { padding: 8px; width: 300px; margin-right: 15px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }"
    $script:htmlContent += ".btn { padding: 8px 15px; margin-right: 10px; border: none; background: #007bff; color: white; border-radius: 4px; cursor: pointer; font-size: 14px; }"
    $script:htmlContent += ".btn:hover { background: #0056b3; }"
    $script:htmlContent += "details { margin-left: 20px; padding-left: 5px; margin-top: 5px; }"
    $script:htmlContent += "summary { cursor: pointer; font-weight: bold; padding: 4px; border-radius: 4px; list-style: none; display: flex; align-items: center; }"
    $script:htmlContent += "summary::-webkit-details-marker { display: none; }"
    $script:htmlContent += "summary::before { content: '▶'; display: inline-block; width: 20px; font-size: 12px; color: #666; transition: transform 0.2s ease-in-out; }"
    $script:htmlContent += "details[open] > summary::before { transform: rotate(90deg); color: #007bff; }"
    $script:htmlContent += "summary:hover { background-color: #f0f0f0; }"
    $script:htmlContent += ".item { margin-left: 45px; padding: 4px 0; color: #555; font-size: 14px; display: flex; align-items: center; gap: 8px; border-radius: 4px; }"
    $script:htmlContent += ".group-icon { color: #e6a817; font-size: 18px; margin-right: 5px; }"
    $script:htmlContent += ".member-icon { font-size: 16px; }"
    $script:htmlContent += ".highlight { background-color: #fff3cd; border-left: 4px solid #ffc107; padding-left: 4px; font-weight: bold; }"
    $script:htmlContent += ".root-group { font-size: 18px; border-bottom: 2px solid #eee; padding-bottom: 5px; margin-top: 15px; }"
    $script:htmlContent += "</style>"
    
    # Skrypty JS
    $script:htmlContent += "<script>"
    $script:htmlContent += "function toggleAll(state) { document.querySelectorAll('details').forEach(d => d.open = state); }"
    $script:htmlContent += "function searchTree() {"
    $script:htmlContent += "  let input = document.getElementById('searchInput').value.toLowerCase();"
    $script:htmlContent += "  let items = document.querySelectorAll('summary, .item');"
    $script:htmlContent += "  items.forEach(el => el.classList.remove('highlight'));"
    $script:htmlContent += "  if (!input) return;"
    $script:htmlContent += "  toggleAll(false);" 
    $script:htmlContent += "  items.forEach(el => {"
    $script:htmlContent += "    if (el.textContent.toLowerCase().includes(input)) {"
    $script:htmlContent += "      el.classList.add('highlight');"
    $script:htmlContent += "      let parent = el.parentElement;"
    $script:htmlContent += "      while (parent && parent.tagName === 'DETAILS') {"
    $script:htmlContent += "        parent.open = true;"
    $script:htmlContent += "        parent = parent.parentElement;"
    $script:htmlContent += "      }"
    $script:htmlContent += "    }"
    $script:htmlContent += "  });"
    $script:htmlContent += "}"
    $script:htmlContent += "</script>"
    $script:htmlContent += "</head><body>"
    $script:htmlContent += "<h2>Struktura Wszystkich Grup w AD</h2>"
    $script:htmlContent += "<div class='controls'>"
    $script:htmlContent += "<input type='text' id='searchInput' class='search-box' placeholder='🔍 Wpisz nazwę, by wyszukać...' onkeyup='searchTree()'>"
    $script:htmlContent += "<button class='btn' onclick='toggleAll(true)'>&#128194; Rozwiń wszystkie</button>"
    $script:htmlContent += "<button class='btn' onclick='toggleAll(false)'>&#128193; Zwiń wszystkie</button>"
    $script:htmlContent += "</div>"

    # Funkcja rekurencyjna
    function Process-GroupNode {
        param([string]$Name)

        if ($script:visitedGroups -contains $Name) {
            $script:htmlContent += "<div class='item'><span class='member-icon'>&#9888;</span> Zapętlenie ($Name)</div>"
            return
        }
        $script:visitedGroups += $Name

        try {
            $members = @(Get-ADGroupMember -Identity $Name -ErrorAction Stop)
        } catch {
            $script:htmlContent += "<div class='item'><span class='member-icon'>&#10060;</span> Brak dostępu</div>"
            return
        }

        if (-not $ShowMembers) {
            $members = $members | Where-Object { $_.objectClass -eq 'group' }
        }
        $members = $members | Sort-Object objectClass, name

        foreach ($member in $members) {
            $descHtml = ""
            if ($IncludeDescription) {
                $adObj = Get-ADObject -Identity $member.distinguishedName -Properties Description -ErrorAction SilentlyContinue
                if ($adObj.Description) {
                    $descHtml = " <span style='color:#888; font-style:italic; font-weight:normal;'>- $($adObj.Description)</span>"
                }
            }

            if ($member.objectClass -eq 'group') {
                $script:htmlContent += "<details><summary><span class='group-icon'>&#128193;</span> $($member.Name)$descHtml</summary>"
                Process-GroupNode -Name $member.samAccountName
                $script:htmlContent += "</details>"
            } else {
                $icon = switch ($member.objectClass) { 'user' { "&#128100;" }; 'computer' { "&#128187;" }; default { "&#128196;" } }
                $script:htmlContent += "<div class='item'><span class='member-icon'>$icon</span> $($member.Name) <small style='color:#999; margin-left:5px;'>($($member.objectClass))</small>$descHtml</div>"
            }
        }
    }

    # Główna pętla budująca strukturę
    $counter = 0
    foreach ($rootGroup in $topLevelGroups) {
        $counter++
        Write-Progress -Activity "Generowanie drzewa AD" -Status "Przetwarzanie grupy: $($rootGroup.Name)" -PercentComplete (($counter / $totalTopGroups) * 100)

        $rootDescHtml = ""
        if ($IncludeDescription -and $rootGroup.Description) {
            $rootDescHtml = " <span style='color:#888; font-weight:normal; font-style:italic;'>- $($rootGroup.Description)</span>"
        }

        # Zaczynamy każdy korzeń jako zamknięty szczegół (inaczej strona HTML by się ładowała pół minuty)
        $script:htmlContent += "<details class='root-group'><summary><span class='group-icon'>&#128193;</span> $($rootGroup.Name)$rootDescHtml</summary>"
        
        $script:visitedGroups = @() # Resetujemy ochronę przed zapętleniem dla każdego nowego drzewa
        Process-GroupNode -Name $rootGroup.samAccountName
        
        $script:htmlContent += "</details>"
    }

    $script:htmlContent += "</body></html>"

    $script:htmlContent | Out-File -FilePath $ExportHtmlPath -Encoding UTF8
    Write-Host "`n✅ Sukces! Plik wygenerowany w: $ExportHtmlPath" -ForegroundColor Green
}