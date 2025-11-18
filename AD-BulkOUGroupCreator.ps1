<# 
GUI do tworzenia wielu OU oraz (tych samych) grup w kazdym OU.
- Wspiera podglad (WhatIf) i wykonanie.
- Transliteration do sAMAccountName: ASCII, max 20, zamiana spacji/symboli na '-'.
- Zapewnienie unikalnosci sAM w domenie (sufiks -1, -2, ...).
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param()

# ======= Modul AD =======
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "Brak modulu ActiveDirectory. Zainstaluj RSAT AD PowerShell (ActiveDirectory)."
}
Import-Module ActiveDirectory -ErrorAction Stop

# ======= WinForms =======
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ======= Pomocnicze =======
$script:DoWhatIf = $false
$script:LogBox   = $null

function Write-Log {
    param([Parameter(Mandatory)][string] $Text)
    $ts = (Get-Date).ToString('HH:mm:ss')
    $line = "$ts  $Text"
    if ($script:LogBox -and -not $script:LogBox.IsDisposed) {
        $script:LogBox.AppendText("$line`r`n")
        $script:LogBox.SelectionStart = $script:LogBox.Text.Length
        $script:LogBox.ScrollToCaret()
    } else {
        Write-Host $line
    }
}

function Test-ADPathExists {
    param([Parameter(Mandatory)][string] $DistinguishedName)
    try {
        Get-ADObject -Identity $DistinguishedName -ErrorAction Stop | Out-Null
        return $true
    } catch { return $false }
}

# Budowa nazwy i opisu grupy
function Build-GroupName {
    param(
        [Parameter(Mandatory)][string] $Prefix,
        [Parameter(Mandatory)][string] $CityCode,
        [string] $RoomCode,
        [Parameter(Mandatory)][string] $Role,
        [Parameter(Mandatory)][char] $Separator
    )
    $parts = @($Prefix, $CityCode)
    if (-not [string]::IsNullOrWhiteSpace($RoomCode)) { $parts += $RoomCode }
    $parts += $Role
    return ($parts -join $Separator)
}

function Expand-DescriptionTemplate {
    param(
        [Parameter(Mandatory)][string] $Template,
        [Parameter(Mandatory)][string] $Prefix,
        [Parameter(Mandatory)][string] $CityFull,
        [Parameter(Mandatory)][string] $CityCode,
        [string] $RoomFull,
        [string] $RoomCode,
        [Parameter(Mandatory)][string] $Role,
        [string] $RoleCode,
        [Parameter(Mandatory)][char] $Separator
    )
    $map = @{
        '{PREFIX}'     = $Prefix
        '{CITY_FULL}'  = $CityFull
        '{CITY_CODE}'  = $CityCode
        '{ROOM_FULL}'  = $RoomFull
        '{ROOM_CODE}'  = $RoomCode
        '{ROLE}'       = $Role
        '{ROLE_CODE}'  = $RoleCode
        '{SEP}'        = [string]$Separator
    }
    $out = $Template
    foreach ($k in $map.Keys) { $out = $out.Replace($k, [string]$map[$k]) }
    return $out
}

# Prosty selektor OU (lista OU z wyszukiwarką)
function Select-OrganizationalUnit {
    try {
        $domainNC = (Get-ADRootDSE).defaultNamingContext
        $ous = Get-ADOrganizationalUnit -Filter * -SearchBase $domainNC -SearchScope Subtree -Properties CanonicalName | 
               Select-Object @{n='Display';e={$_.CanonicalName}}, DistinguishedName | Sort-Object Display
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Nie moge pobrac listy OU: $($_.Exception.Message)", "Blad",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        return $null
    }

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = 'Wybierz OU'
    $dlg.Size = New-Object System.Drawing.Size(600, 500)
    $dlg.StartPosition = 'CenterParent'

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = 'Filtr:'
    $lbl.Location = '10,12'
    $lbl.AutoSize = $true

    $txt = New-Object System.Windows.Forms.TextBox
    $txt.Location = '50,10'
    $txt.Size = '520,22'

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = '10,40'
    $lst.Size = '560,360'
    $lst.IntegralHeight = $false
    $lst.DisplayMember = 'Display'
    $lst.ValueMember = 'DistinguishedName'

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Location = '380,415'
    $btnOk.Size = '90,28'

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Anuluj'
    $btnCancel.Location = '480,415'
    $btnCancel.Size = '90,28'

    $all = [System.Collections.ArrayList]::new()
    [void]$all.AddRange($ous)
    function Refresh-List { param($filter)
        $lst.Items.Clear()
        $items = if ([string]::IsNullOrWhiteSpace($filter)) { $all } else { $all | Where-Object { $_.Display -like "*${filter}*" } }
        foreach ($it in $items) { [void]$lst.Items.Add($it) }
    }
    Refresh-List ''

    $txt.Add_TextChanged({ Refresh-List $txt.Text })
    $btnOk.Add_Click({ $dlg.Tag = if ($lst.SelectedItem) { $lst.SelectedItem.DistinguishedName } else { $null }; $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK })
    $btnCancel.Add_Click({ $dlg.Tag = $null; $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel })
    $lst.Add_MouseDoubleClick({ if ($lst.SelectedItem) { $dlg.Tag = $lst.SelectedItem.DistinguishedName; $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK } })

    $dlg.Controls.AddRange(@($lbl,$txt,$lst,$btnOk,$btnCancel))
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return [string]$dlg.Tag } else { return $null }
}

# transliteracja do sAM: ASCII, '-', '_.', max 20
function Convert-ToSamSafe {
    param([Parameter(Mandatory)][string] $Text)
    # Jawna transliteracja polskich znaków (bez hashtabla, unikamy duplikatów kluczy)
    $from = @('ą','ć','ę','ł','ń','ó','ś','ź','ż','Ą','Ć','Ę','Ł','Ń','Ó','Ś','Ź','Ż')
    $to   = @('a','c','e','l','n','o','s','z','z','A','C','E','L','N','O','S','Z','Z')
    for ($i = 0; $i -lt $from.Count; $i++) {
        $Text = $Text -replace [regex]::Escape($from[$i]), $to[$i]
    }
    $t = $Text -replace '[^A-Za-z0-9\-_.]','-'        # spacje i inne -> '-'
    if ($t.Length -gt 20) { $t = $t.Substring(0,20) } # limit sAM
    return $t
}

# sAM unikalny w domenie: jesli zajety, doklej -1 / -2 / ...
function Get-UniqueSam {
    param([Parameter(Mandatory)][string] $Candidate)
    $sam = $Candidate
    $i = 1
    while (Get-ADObject -LDAPFilter "(sAMAccountName=$sam)" -ErrorAction SilentlyContinue) {
        $suffix = "-$i"
        $maxBase = 20 - $suffix.Length
        if ($maxBase -lt 1) { throw "Nie mozna wygenerowac unikalnego sAMAccountName dla '$Candidate'." }
        $base = $Candidate.Substring(0, [Math]::Min($Candidate.Length, $maxBase))
        $sam = "$base$suffix"
        $i++
    }
    return $sam
}

function Ensure-OU {
    param(
        [Parameter(Mandatory)][string] $Name, 
        [Parameter(Mandatory)][string] $ParentDN,
        [Parameter()][bool] $Protect = $true
    )
    $targetDN = "OU=$Name,$ParentDN"
    if (Test-ADPathExists $targetDN) { 
        Write-Log "[=] OU istnieje: $targetDN"
        return $targetDN 
    }

    if ($PSCmdlet.ShouldProcess($targetDN, "Utw�rz OU")) {
        try {
            if ($script:DoWhatIf) {
                Write-Log "[WHATIF] Utworzylbym OU: $targetDN (Protected=$Protect)"
            } else {
                New-ADOrganizationalUnit -Name $Name -Path $ParentDN `
                    -ProtectedFromAccidentalDeletion:$Protect `
                    -ErrorAction Stop | Out-Null
                Write-Log "[OK] OU: $targetDN"
            }
        } catch {
            Write-Log "[ERR] Nie udalo sie utworzyc OU '$targetDN': $($_.Exception.Message)"
            throw
        }
    }
    return $targetDN
}

function Ensure-Group {
    param(
        [Parameter(Mandatory)][string] $Name,      # nazwa wyswietlana (CN/Name)
        [Parameter(Mandatory)][string] $Sam,       # sAMAccountName (ASCII)
        [Parameter(Mandatory)][ValidateSet('Global','Universal','DomainLocal')] [string] $Scope,
        [Parameter(Mandatory)][ValidateSet('Security','Distribution')] [string] $Category,
        [Parameter(Mandatory)][string] $Path,      # OU docelowe (DN)
        [string] $Description
    )
    $nameEsc = $Name.Replace("'", "'")
    $existing = Get-ADGroup -Filter "Name -eq '$nameEsc'" -SearchBase $Path -ErrorAction SilentlyContinue
    if ($existing) { 
        Write-Log "[=] Grupa istnieje: $($existing.DistinguishedName)"
        return $existing.DistinguishedName 
    }

    if ($PSCmdlet.ShouldProcess("$Name @ $Path", "Utw�rz grupe")) {
        try {
            if ($script:DoWhatIf) {
                Write-Log "[WHATIF] Utworzylbym grupe: $Name (sAM=$Sam, $Scope/$Category) w $Path"
            } else {
                New-ADGroup -Name $Name `
                            -SamAccountName $Sam `
                            -GroupScope $Scope `
                            -GroupCategory $Category `
                            -Path $Path `
                            -Description $Description `
                            -ErrorAction Stop | Out-Null
                Write-Log "[OK] Grupa: $Name (sAM=$Sam)"
                $created = Get-ADGroup -LDAPFilter "(sAMAccountName=$Sam)" -SearchBase $Path -ErrorAction SilentlyContinue
                if ($created) { return $created.DistinguishedName }
            }
        } catch {
            Write-Log "[ERR] Nie udalo sie utworzyc grupy '$Name': $($_.Exception.Message)"
            throw
        }
    }
    return "CN=$Name,$Path"
}

function Add-GroupMembers {
    param(
        [Parameter(Mandatory)][string] $GroupDN,
        [Parameter()][string[]] $Members
    )

    if ([string]::IsNullOrWhiteSpace($GroupDN)) { return }
    if (-not $Members -or $Members.Count -eq 0) { return }
    $normalized = @($Members | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($normalized.Count -eq 0) { return }

    if ($script:DoWhatIf) {
        $whatIfList = [string]::Join(', ', $normalized)
        Write-Log ("[WHATIF] Dodalbym do grupy {0} czlonkow: {1}" -f $GroupDN, $whatIfList)
        return
    }

    try {
        $existingDns = @()
        try {
            $existingDns = Get-ADGroupMember -Identity $GroupDN -ErrorAction Stop | Select-Object -ExpandProperty DistinguishedName
        } catch {
            $existingDns = @()
        }

        $toAdd = @($normalized | Where-Object { $existingDns -notcontains $_ })
        if ($toAdd.Count -eq 0) {
            Write-Log ("[=] Grupa {0} juz zawiera podane elementy" -f $GroupDN)
            return
        }

        Add-ADGroupMember -Identity $GroupDN -Members $toAdd -ErrorAction Stop
        $addedList = [string]::Join(', ', $toAdd)
        Write-Log ("[OK] Dodano do grupy {0}: {1}" -f $GroupDN, $addedList)
    } catch {
        Write-Log ("[ERR] Dodawanie czlonkow do grupy {0}: {1}" -f $GroupDN, $_.Exception.Message)
        throw
    }
}

function Invoke-Create {
    param(
        [Parameter(Mandatory)][string] $BaseDN,
        [Parameter(Mandatory)][string[]] $Cities,
        [Parameter(Mandatory)][string[]] $CityCodes,
        [Parameter(Mandatory)][string[]] $ComputerRoles,
        [Parameter(Mandatory)][string[]] $RoleCodes,
        [Parameter(Mandatory)][ValidateSet('Global','Universal','DomainLocal')] [string] $GroupScope,
        [Parameter(Mandatory)][ValidateSet('Security','Distribution')] [string] $GroupCategory,
        [Parameter(Mandatory)][string] $GroupNamePrefix,
        [Parameter()][char] $SeparatorChar = '-',
        [Parameter()][string] $GroupNameTemplate = '{PREFIX}{SEP}{CITY_CODE}{SEP}{ROLE_CODE}',
        [Parameter()][string] $DescriptionTemplate = 'Grupa komputerow: {ROLE} ({CITY_FULL})',
        [Parameter()][bool] $ProtectOUs = $true,
        [Parameter()][bool] $CreateCityAll = $true,
        [Parameter()][bool] $CreateRoleAll = $true,
        [Parameter()][bool] $CreateGlobalAll = $true
    )

    if (-not (Test-ADPathExists $BaseDN)) {
        throw "Base DN nie istnieje: $BaseDN"
    }

    Write-Log "=== START (WhatIf=$script:DoWhatIf) ==="
    Write-Log "Base DN: $BaseDN"
    Write-Log "OU (miasta): $($Cities -join ', ')"
    Write-Log "Skroty miast: $($CityCodes -join ', ')"
    Write-Log "Role (grupy): $($ComputerRoles -join ', ')"
    Write-Log "Skroty dzialow: $($RoleCodes -join ', ')"
    Write-Log "Prefiks grup: $GroupNamePrefix | Scope/Category: $GroupScope/$GroupCategory | Separator: '$SeparatorChar'"
    Write-Log "Ochrona OU przed usunieciem: $ProtectOUs"
    Write-Log ("Grupy ALL (miasto/rola/globalna): {0}/{1}/{2}" -f $CreateCityAll, $CreateRoleAll, $CreateGlobalAll)

    $globalMemberGroups = @()
    $roleAggregates = @{}

    for ($i = 0; $i -lt $Cities.Count; $i++) {
        $city = $Cities[$i]
        if ([string]::IsNullOrWhiteSpace($city)) { continue }
        $city = $city.Trim()

        if ($i -ge $CityCodes.Count) { throw "Brak skrotu miasta dla '$city'." }
        $cityCode = $CityCodes[$i]
        if ([string]::IsNullOrWhiteSpace($cityCode)) { throw "Brak skrotu miasta dla '$city'." }
        $cityCode = $cityCode.Trim().ToUpperInvariant()

        # 1) OU miasta
        $cityDN = Ensure-OU -Name $city -ParentDN $BaseDN -Protect:$ProtectOUs

        # 2) Grupy w OU miasta (bez pod-OU)
        $grpOUdn = $cityDN

        $cityRoleGroupDns = @()

        for ($j = 0; $j -lt $ComputerRoles.Count; $j++) {
            $role = $ComputerRoles[$j]
            if ([string]::IsNullOrWhiteSpace($role)) { continue }
            $role = $role.Trim()

            if ($j -ge $RoleCodes.Count) { throw "Brak skrotu roli dla '$role'." }
            $roleCode = $RoleCodes[$j]
            if ([string]::IsNullOrWhiteSpace($roleCode)) { throw "Brak skrotu roli dla '$role'." }
            $roleCode = $roleCode.Trim().ToUpperInvariant()

            # CN wg szablonu i sAM (ASCII, <=20) + unikalnosc
            $cnName   = Expand-DescriptionTemplate -Template $GroupNameTemplate -Prefix $GroupNamePrefix -CityFull $city -CityCode $cityCode -RoomFull $null -RoomCode $null -Role $role -RoleCode $roleCode -Separator $SeparatorChar
            $samName0 = Convert-ToSamSafe -Text $cnName
            $samName  = Get-UniqueSam -Candidate $samName0

            $desc = Expand-DescriptionTemplate -Template $DescriptionTemplate -Prefix $GroupNamePrefix -CityFull $city -CityCode $cityCode -RoomFull $null -RoomCode $null -Role $role -RoleCode $roleCode -Separator $SeparatorChar
            $groupDn = Ensure-Group -Name $cnName -Sam $samName -Scope $GroupScope -Category $GroupCategory -Path $grpOUdn -Description $desc
            if ($groupDn) {
                $cityRoleGroupDns += $groupDn
                if (-not $roleAggregates.ContainsKey($roleCode)) {
                    $roleAggregates[$roleCode] = [PSCustomObject]@{
                        Role = $role
                        RoleCode = $roleCode
                        Members = New-Object System.Collections.Generic.List[string]
                    }
                }
                $null = $roleAggregates[$roleCode].Members.Add($groupDn)
            }
        }

        if ($CreateCityAll -and $cityRoleGroupDns.Count -gt 0) {
            $cityAllName = Expand-DescriptionTemplate -Template $GroupNameTemplate -Prefix $GroupNamePrefix -CityFull $city -CityCode $cityCode -RoomFull $null -RoomCode $null -Role 'ALL' -RoleCode 'ALL' -Separator $SeparatorChar
            $cityAllSam0 = Convert-ToSamSafe -Text $cityAllName
            $cityAllSam  = Get-UniqueSam -Candidate $cityAllSam0
            $cityAllDesc = Expand-DescriptionTemplate -Template $DescriptionTemplate -Prefix $GroupNamePrefix -CityFull $city -CityCode $cityCode -RoomFull $null -RoomCode $null -Role 'ALL' -RoleCode 'ALL' -Separator $SeparatorChar
            $cityAllDn   = Ensure-Group -Name $cityAllName -Sam $cityAllSam -Scope $GroupScope -Category $GroupCategory -Path $grpOUdn -Description $cityAllDesc
            if ($cityAllDn) {
                Add-GroupMembers -GroupDN $cityAllDn -Members $cityRoleGroupDns
                $globalMemberGroups += $cityAllDn
                Write-Log ("[=] Grupa ALL miasta: {0}" -f $cityAllDn)
            } else {
                $globalMemberGroups += $cityRoleGroupDns
            }
        } else {
            $globalMemberGroups += $cityRoleGroupDns
        }
    }

    if ($CreateRoleAll -and $roleAggregates.Count -gt 0) {
        foreach ($entry in $roleAggregates.GetEnumerator()) {
            $roleInfo = $entry.Value
            $roleMembers = @($roleInfo.Members | Where-Object { $_ } | Select-Object -Unique)
            if ($roleMembers.Count -eq 0) { continue }

            $roleAllName = Expand-DescriptionTemplate -Template $GroupNameTemplate -Prefix $GroupNamePrefix -CityFull 'WSZYSTKIE' -CityCode 'ALL' -RoomFull $null -RoomCode $null -Role $roleInfo.Role -RoleCode $roleInfo.RoleCode -Separator $SeparatorChar
            $roleAllSam0 = Convert-ToSamSafe -Text $roleAllName
            $roleAllSam  = Get-UniqueSam -Candidate $roleAllSam0
            $roleAllDesc = Expand-DescriptionTemplate -Template $DescriptionTemplate -Prefix $GroupNamePrefix -CityFull 'WSZYSTKIE' -CityCode 'ALL' -RoomFull $null -RoomCode $null -Role $roleInfo.Role -RoleCode $roleInfo.RoleCode -Separator $SeparatorChar
            $roleAllDn   = Ensure-Group -Name $roleAllName -Sam $roleAllSam -Scope $GroupScope -Category $GroupCategory -Path $BaseDN -Description $roleAllDesc
            if ($roleAllDn) {
                Add-GroupMembers -GroupDN $roleAllDn -Members $roleMembers
                Write-Log ("[=] Grupa ALL roli ({0}): {1}" -f $roleInfo.RoleCode, $roleAllDn)
            }
        }
    }

    if ($CreateGlobalAll -and $globalMemberGroups.Count -gt 0) {
        $uniqueMembers = @($globalMemberGroups | Where-Object { $_ } | Select-Object -Unique)
        if ($uniqueMembers.Count -gt 0) {
            $globalName = Expand-DescriptionTemplate -Template $GroupNameTemplate -Prefix $GroupNamePrefix -CityFull 'WSZYSTKIE' -CityCode 'ALL' -RoomFull $null -RoomCode $null -Role 'ALL' -RoleCode 'ALL' -Separator $SeparatorChar
            $globalSam0 = Convert-ToSamSafe -Text $globalName
            $globalSam  = Get-UniqueSam -Candidate $globalSam0
            $globalDesc = Expand-DescriptionTemplate -Template $DescriptionTemplate -Prefix $GroupNamePrefix -CityFull 'WSZYSTKIE' -CityCode 'ALL' -RoomFull $null -RoomCode $null -Role 'ALL' -RoleCode 'ALL' -Separator $SeparatorChar
            $globalDn   = Ensure-Group -Name $globalName -Sam $globalSam -Scope $GroupScope -Category $GroupCategory -Path $BaseDN -Description $globalDesc
            if ($globalDn) {
                Add-GroupMembers -GroupDN $globalDn -Members $uniqueMembers
                Write-Log ("[=] Grupa ALL globalna: {0}" -f $globalDn)
            }
        }
    }

    Write-Log "=== KONIEC (WhatIf=$script:DoWhatIf) ===`r`n"
}

# ======= GUI =======
[System.Windows.Forms.Application]::EnableVisualStyles()
$fontUI = New-Object System.Drawing.Font('Segoe UI', 9)

$form = New-Object System.Windows.Forms.Form
$form.Text = "OU + Grupy (GUI) - Active Directory"
$form.Size = New-Object System.Drawing.Size(980, 760)
$form.MinimumSize = New-Object System.Drawing.Size(980, 760)
$form.StartPosition = 'CenterScreen'
$form.Font = $fontUI
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.Padding = New-Object System.Windows.Forms.Padding(10)

# Base DN
$lblBase = New-Object System.Windows.Forms.Label
$lblBase.Text = "Base DN (np. OU=Company,DC=contoso,DC=local)"
$lblBase.AutoSize = $true

$txtBase = New-Object System.Windows.Forms.TextBox
$txtBase.MinimumSize = New-Object System.Drawing.Size(200, 0)
$txtBase.Dock = 'Fill'
$txtBase.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)

$btnPickOU = New-Object System.Windows.Forms.Button
$btnPickOU.Text = "Wybierz OU..."
$btnPickOU.AutoSize = $true
$btnPickOU.AutoSizeMode = 'GrowAndShrink'
$btnPickOU.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)

$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = "Test DN"
$btnTest.AutoSize = $true
$btnTest.AutoSizeMode = 'GrowAndShrink'
$btnTest.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)

# OU (miasta)
$lblCities = New-Object System.Windows.Forms.Label
$lblCities.Text = "OU (miasta) - kazdy w nowej linii"
$lblCities.AutoSize = $true
$lblCities.Margin = New-Object System.Windows.Forms.Padding(0,10,0,0)

$txtCities = New-Object System.Windows.Forms.TextBox
$txtCities.Multiline = $true
$txtCities.ScrollBars = 'Vertical'
$txtCities.MinimumSize = New-Object System.Drawing.Size(0, 180)
$txtCities.Dock = 'Fill'
$txtCities.Margin = New-Object System.Windows.Forms.Padding(0,4,0,8)

# Skroty miast
$lblCityCodes = New-Object System.Windows.Forms.Label
$lblCityCodes.Text = "Skroty miast - ta sama kolejnosc"
$lblCityCodes.AutoSize = $true
$lblCityCodes.Margin = New-Object System.Windows.Forms.Padding(0,6,0,0)

$txtCityCodes = New-Object System.Windows.Forms.TextBox
$txtCityCodes.Multiline = $true
$txtCityCodes.ScrollBars = 'Vertical'
$txtCityCodes.MinimumSize = New-Object System.Drawing.Size(0,120)
$txtCityCodes.Dock = 'Fill'
$txtCityCodes.Margin = New-Object System.Windows.Forms.Padding(0,4,0,8)

# Role (grupy)
$lblRoles = New-Object System.Windows.Forms.Label
$lblRoles.Text = "Role (grupy) - kazda w nowej linii"
$lblRoles.AutoSize = $true
$lblRoles.Margin = New-Object System.Windows.Forms.Padding(0,10,0,0)

$txtRoles = New-Object System.Windows.Forms.TextBox
$txtRoles.Multiline = $true
$txtRoles.ScrollBars = 'Vertical'
$txtRoles.MinimumSize = New-Object System.Drawing.Size(0, 180)
$txtRoles.Dock = 'Fill'
$txtRoles.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)
$txtRoles.Lines = @("Rejestracja","Polozne","Diagnostyka","Embriologia","Inne")

# Skroty dzialow (role)
$lblRoleCodes = New-Object System.Windows.Forms.Label
$lblRoleCodes.Text = "Skroty dzialow (Role) - ta sama kolejnosc"
$lblRoleCodes.AutoSize = $true
$lblRoleCodes.Margin = New-Object System.Windows.Forms.Padding(0,6,0,0)

$txtRoleCodes = New-Object System.Windows.Forms.TextBox
$txtRoleCodes.Multiline = $true
$txtRoleCodes.ScrollBars = 'Vertical'
$txtRoleCodes.MinimumSize = New-Object System.Drawing.Size(0,120)
$txtRoleCodes.Dock = 'Fill'
$txtRoleCodes.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)

# Kontenery pomocnicze dla pol tekstowych (lewa kolumna)
$panelCities = New-Object System.Windows.Forms.TableLayoutPanel
$panelCities.ColumnCount = 1
$panelCities.RowCount = 2
$panelCities.Dock = 'Fill'
$panelCities.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
[void]$panelCities.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$panelCities.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$panelCities.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$lblCities.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)
[void]$panelCities.Controls.Add($lblCities, 0, 0)
[void]$panelCities.Controls.Add($txtCities, 0, 1)

$panelCityCodes = New-Object System.Windows.Forms.TableLayoutPanel
$panelCityCodes.ColumnCount = 1
$panelCityCodes.RowCount = 2
$panelCityCodes.Dock = 'Fill'
$panelCityCodes.Padding = New-Object System.Windows.Forms.Padding(0,8,0,0)
$panelCityCodes.Margin = New-Object System.Windows.Forms.Padding(0,0,8,0)
[void]$panelCityCodes.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$panelCityCodes.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$panelCityCodes.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$panelCityCodes.Controls.Add($lblCityCodes, 0, 0)
[void]$panelCityCodes.Controls.Add($txtCityCodes, 0, 1)

$panelRoles = New-Object System.Windows.Forms.TableLayoutPanel
$panelRoles.ColumnCount = 1
$panelRoles.RowCount = 2
$panelRoles.Dock = 'Fill'
$panelRoles.Padding = New-Object System.Windows.Forms.Padding(0,0,0,0)
$panelRoles.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)
[void]$panelRoles.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$panelRoles.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$panelRoles.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$panelRoles.Controls.Add($lblRoles, 0, 0)
[void]$panelRoles.Controls.Add($txtRoles, 0, 1)

$panelRoleCodes = New-Object System.Windows.Forms.TableLayoutPanel
$panelRoleCodes.ColumnCount = 1
$panelRoleCodes.RowCount = 2
$panelRoleCodes.Dock = 'Fill'
$panelRoleCodes.Padding = New-Object System.Windows.Forms.Padding(0,0,0,0)
$panelRoleCodes.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)
[void]$panelRoleCodes.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$panelRoleCodes.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$panelRoleCodes.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$baseRow = New-Object System.Windows.Forms.TableLayoutPanel
$baseRow.ColumnCount = 3
$baseRow.RowCount = 1
$baseRow.Dock = 'Fill'
[void]$baseRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$baseRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$baseRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$baseRow.Controls.Add($txtBase, 0, 0)
[void]$baseRow.Controls.Add($btnPickOU, 1, 0)
[void]$baseRow.Controls.Add($btnTest, 2, 0)

[void]$panelRoleCodes.Controls.Add($lblRoleCodes, 0, 0)
[void]$panelRoleCodes.Controls.Add($txtRoleCodes, 0, 1)

# Układ lewego panelu (miasta/role)
$leftLayout = New-Object System.Windows.Forms.TableLayoutPanel
$leftLayout.ColumnCount = 2
$leftLayout.RowCount = 4
$leftLayout.Dock = 'Fill'
$leftLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$leftLayout.ColumnStyles.Clear()
[void]$leftLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$leftLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$leftLayout.RowStyles.Clear()
[void]$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
[void]$leftLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))

$panelRoles.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)
$panelRoleCodes.Margin = New-Object System.Windows.Forms.Padding(8,0,0,0)

[void]$leftLayout.Controls.Add($lblBase, 0, 0)
$leftLayout.SetColumnSpan($lblBase, 2)
[void]$leftLayout.Controls.Add($baseRow, 0, 1)
$leftLayout.SetColumnSpan($baseRow, 2)
[void]$leftLayout.Controls.Add($panelCities, 0, 2)
[void]$leftLayout.Controls.Add($panelRoles, 1, 2)
[void]$leftLayout.Controls.Add($panelCityCodes, 0, 3)
[void]$leftLayout.Controls.Add($panelRoleCodes, 1, 3)

$leftContainer = New-Object System.Windows.Forms.GroupBox
$leftContainer.Text = "Lokalizacje i role"
$leftContainer.Dock = 'Fill'
$leftContainer.Padding = New-Object System.Windows.Forms.Padding(10)
$leftContainer.Margin = New-Object System.Windows.Forms.Padding(0,0,10,0)
[void]$leftContainer.Controls.Add($leftLayout)

# Parametry grup
$groupBox = New-Object System.Windows.Forms.GroupBox
$groupBox.Text = "Parametry grup / OU"
$groupBox.Dock = 'Fill'
$groupBox.Padding = New-Object System.Windows.Forms.Padding(10)
$groupBox.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)

$groupLayout = New-Object System.Windows.Forms.TableLayoutPanel
$groupLayout.Dock = 'Fill'
$groupLayout.ColumnCount = 2
$groupLayout.RowCount = 8
[void]$groupLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$groupLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
for ($i = 0; $i -lt 8; $i++) {
    [void]$groupLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
}

$lblPrefix = New-Object System.Windows.Forms.Label
$lblPrefix.Text = "Prefiks grup"
$lblPrefix.AutoSize = $true

$txtPrefix = New-Object System.Windows.Forms.TextBox
$txtPrefix.Text = "GG-COMP"
$txtPrefix.Dock = 'Fill'
$txtPrefix.Margin = New-Object System.Windows.Forms.Padding(5,0,0,0)

$lblScope = New-Object System.Windows.Forms.Label
$lblScope.Text = "Group Scope"
$lblScope.AutoSize = $true
$lblScope.Margin = New-Object System.Windows.Forms.Padding(0,10,0,0)

$cmbScope = New-Object System.Windows.Forms.ComboBox
$cmbScope.DropDownStyle = 'DropDownList'
[void]$cmbScope.Items.AddRange(@('Global','Universal','DomainLocal'))
$cmbScope.SelectedItem = 'Global'
$cmbScope.Dock = 'Fill'
$cmbScope.Margin = New-Object System.Windows.Forms.Padding(5,10,0,0)

$lblCat = New-Object System.Windows.Forms.Label
$lblCat.Text = "Group Category"
$lblCat.AutoSize = $true
$lblCat.Margin = New-Object System.Windows.Forms.Padding(0,10,0,0)

$cmbCat = New-Object System.Windows.Forms.ComboBox
$cmbCat.DropDownStyle = 'DropDownList'
[void]$cmbCat.Items.AddRange(@('Security','Distribution'))
$cmbCat.SelectedItem = 'Security'
$cmbCat.Dock = 'Fill'
$cmbCat.Margin = New-Object System.Windows.Forms.Padding(5,10,0,0)

$lblSep = New-Object System.Windows.Forms.Label
$lblSep.Text = "Separator"
$lblSep.AutoSize = $true
$lblSep.Margin = New-Object System.Windows.Forms.Padding(0,10,0,0)

$cmbSep = New-Object System.Windows.Forms.ComboBox
$cmbSep.DropDownStyle = 'DropDownList'
[void]$cmbSep.Items.AddRange(@('-','_') )
$cmbSep.SelectedItem = '-'
$cmbSep.Dock = 'Fill'
$cmbSep.Margin = New-Object System.Windows.Forms.Padding(5,10,0,0)

$chkProtect = New-Object System.Windows.Forms.CheckBox
$chkProtect.Text = "Chron OU przed usunieciem"
$chkProtect.Checked = $true
$chkProtect.AutoSize = $true
$chkProtect.Margin = New-Object System.Windows.Forms.Padding(0,15,0,0)

$chkCityAll = New-Object System.Windows.Forms.CheckBox
$chkCityAll.Text = "Grupa ALL dla miasta (wszystkie role)"
$chkCityAll.Checked = $true
$chkCityAll.AutoSize = $true
$chkCityAll.Margin = New-Object System.Windows.Forms.Padding(0,10,0,0)

$chkRoleAll = New-Object System.Windows.Forms.CheckBox
$chkRoleAll.Text = "Grupa ALL dla roli (wszystkie lokalizacje)"
$chkRoleAll.Checked = $true
$chkRoleAll.AutoSize = $true
$chkRoleAll.Margin = New-Object System.Windows.Forms.Padding(0,6,0,0)

$chkGlobalAll = New-Object System.Windows.Forms.CheckBox
$chkGlobalAll.Text = "Grupa ALL globalna"
$chkGlobalAll.Checked = $true
$chkGlobalAll.AutoSize = $true
$chkGlobalAll.Margin = New-Object System.Windows.Forms.Padding(0,6,0,0)

[void]$groupLayout.Controls.Add($lblPrefix, 0, 0)
[void]$groupLayout.Controls.Add($txtPrefix, 1, 0)
[void]$groupLayout.Controls.Add($lblScope, 0, 1)
[void]$groupLayout.Controls.Add($cmbScope, 1, 1)
[void]$groupLayout.Controls.Add($lblCat, 0, 2)
[void]$groupLayout.Controls.Add($cmbCat, 1, 2)
[void]$groupLayout.Controls.Add($lblSep, 0, 3)
[void]$groupLayout.Controls.Add($cmbSep, 1, 3)
[void]$groupLayout.Controls.Add($chkProtect, 0, 4)
$groupLayout.SetColumnSpan($chkProtect, 2)
[void]$groupLayout.Controls.Add($chkCityAll, 0, 5)
$groupLayout.SetColumnSpan($chkCityAll, 2)
[void]$groupLayout.Controls.Add($chkRoleAll, 0, 6)
$groupLayout.SetColumnSpan($chkRoleAll, 2)
[void]$groupLayout.Controls.Add($chkGlobalAll, 0, 7)
$groupLayout.SetColumnSpan($chkGlobalAll, 2)
[void]$groupBox.Controls.Add($groupLayout)

# Szablony
$lblNameTpl = New-Object System.Windows.Forms.Label
$lblNameTpl.Text = "Szablon nazwy grupy (uzyj: {PREFIX},{CITY_FULL},{CITY_CODE},{ROLE},{ROLE_CODE},{SEP})"
$lblNameTpl.AutoSize = $true
$lblNameTpl.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)

$txtNameTpl = New-Object System.Windows.Forms.TextBox
$txtNameTpl.Text = '{PREFIX}{SEP}{CITY_CODE}{SEP}{ROLE_CODE}'
$txtNameTpl.Dock = 'Fill'
$txtNameTpl.Margin = New-Object System.Windows.Forms.Padding(0,6,0,0)

$lblDescTpl = New-Object System.Windows.Forms.Label
$lblDescTpl.Text = "Szablon opisu (uzyj: {PREFIX},{CITY_FULL},{CITY_CODE},{ROLE},{ROLE_CODE},{SEP})"
$lblDescTpl.AutoSize = $true
$lblDescTpl.Margin = New-Object System.Windows.Forms.Padding(0,12,0,0)

$txtDescTpl = New-Object System.Windows.Forms.TextBox
$txtDescTpl.Text = 'Grupa komputerow: {ROLE} ({CITY_FULL})'
$txtDescTpl.Multiline = $true
$txtDescTpl.ScrollBars = 'Vertical'
$txtDescTpl.Dock = 'Fill'
$txtDescTpl.Margin = New-Object System.Windows.Forms.Padding(0,6,0,0)
$txtDescTpl.MinimumSize = New-Object System.Drawing.Size(0, 140)

# Log
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log / wynik:"
$lblLog.AutoSize = $true
$lblLog.Margin = New-Object System.Windows.Forms.Padding(0,10,0,4)

$rtb = New-Object System.Windows.Forms.RichTextBox
$rtb.ReadOnly = $true
$rtb.Font = New-Object System.Drawing.Font('Consolas', 9)
$rtb.Dock = 'Fill'
$rtb.Margin = New-Object System.Windows.Forms.Padding(0,4,0,0)
$script:LogBox = $rtb

# Układ szablonów po prawej
$templateLayout = New-Object System.Windows.Forms.TableLayoutPanel
$templateLayout.ColumnCount = 1
$templateLayout.RowCount = 4
$templateLayout.Dock = 'Fill'
$templateLayout.Padding = New-Object System.Windows.Forms.Padding(0)
[void]$templateLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$templateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$templateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$templateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$templateLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$templateLayout.Controls.Add($lblNameTpl, 0, 0)
[void]$templateLayout.Controls.Add($txtNameTpl, 0, 1)
[void]$templateLayout.Controls.Add($lblDescTpl, 0, 2)
[void]$templateLayout.Controls.Add($txtDescTpl, 0, 3)

$templateBox = New-Object System.Windows.Forms.GroupBox
$templateBox.Text = "Szablony nazwy i opisu"
$templateBox.Dock = 'Fill'
$templateBox.Padding = New-Object System.Windows.Forms.Padding(10)
$templateBox.Margin = New-Object System.Windows.Forms.Padding(0,10,0,0)
[void]$templateBox.Controls.Add($templateLayout)

# Prawy panel: parametry + szablony
$rightLayout = New-Object System.Windows.Forms.TableLayoutPanel
$rightLayout.ColumnCount = 1
$rightLayout.RowCount = 2
$rightLayout.Dock = 'Fill'
$rightLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$rightLayout.Margin = New-Object System.Windows.Forms.Padding(0)
[void]$rightLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$rightLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$rightLayout.Controls.Add($groupBox, 0, 0)
[void]$rightLayout.Controls.Add($templateBox, 0, 1)

# Glowny layout ciala
$bodyLayout = New-Object System.Windows.Forms.TableLayoutPanel
$bodyLayout.ColumnCount = 2
$bodyLayout.RowCount = 1
$bodyLayout.Dock = 'Fill'
$bodyLayout.Padding = New-Object System.Windows.Forms.Padding(0,0,0,8)
[void]$bodyLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 55)))
[void]$bodyLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 45)))
[void]$bodyLayout.Controls.Add($leftContainer, 0, 0)
[void]$bodyLayout.Controls.Add($rightLayout, 1, 0)

# Layout calosci (cala forma bez panelu przyciskow)
$contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
$contentLayout.ColumnCount = 1
$contentLayout.RowCount = 3
$contentLayout.Dock = 'Fill'
[void]$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 70)))
[void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$contentLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 30)))
[void]$contentLayout.Controls.Add($bodyLayout, 0, 0)
[void]$contentLayout.Controls.Add($lblLog, 0, 1)
[void]$contentLayout.Controls.Add($rtb, 0, 2)
# Layout kontenerow


# Przyciski
$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = "Podglad (WhatIf)"
$btnPreview.AutoSize = $true
$btnPreview.AutoSizeMode = 'GrowAndShrink'
$btnPreview.Padding = New-Object System.Windows.Forms.Padding(12,6,12,6)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Utworz"
$btnRun.AutoSize = $true
$btnRun.AutoSizeMode = 'GrowAndShrink'
$btnRun.Padding = New-Object System.Windows.Forms.Padding(12,6,12,6)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Zamknij"
$btnClose.AutoSize = $true
$btnClose.AutoSizeMode = 'GrowAndShrink'
$btnClose.Padding = New-Object System.Windows.Forms.Padding(12,6,12,6)

$btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$btnPanel.Dock = 'Bottom'
$btnPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
$btnPanel.Padding = New-Object System.Windows.Forms.Padding(0,8,0,0)
$btnPanel.AutoSize = $true
$btnPanel.WrapContents = $false

$btnClose.Margin = New-Object System.Windows.Forms.Padding(10,0,0,0)
$btnRun.Margin = New-Object System.Windows.Forms.Padding(10,0,0,0)
$btnPreview.Margin = New-Object System.Windows.Forms.Padding(0,0,0,0)

[void]$btnPanel.Controls.Add($btnClose)
[void]$btnPanel.Controls.Add($btnRun)
[void]$btnPanel.Controls.Add($btnPreview)

[void]$form.Controls.Add($contentLayout)
[void]$form.Controls.Add($btnPanel)

# brak dodatkowego splitera – layout statyczny
# Zdarzenia
$btnTest.Add_Click({
    $dn = $txtBase.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($dn)) {
        Write-Log "[!] Podaj Base DN."
        return
    }
    if (Test-ADPathExists -DistinguishedName $dn) {
        Write-Log "[OK] Base DN istnieje: $dn"
    } else {
        Write-Log "[!] Base DN nie istnieje: $dn"
    }
})

# Wybierz OU (picker)
$btnPickOU.Add_Click({
    try {
        $sel = Select-OrganizationalUnit
        if ($sel) { $txtBase.Text = $sel; Write-Log "[=] Wybrano OU: $sel" }
    } catch {
        Write-Log "[ERR] Picker OU: $($_.Exception.Message)"
    }
})

function Get-UiValues {
    $baseDN = $txtBase.Text.Trim()
    $cities = @($txtCities.Lines | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $roles  = @($txtRoles.Lines  | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $cityCodes = @($txtCityCodes.Lines | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $roleCodes = @($txtRoleCodes.Lines | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $scope  = [string]$cmbScope.SelectedItem
    $cat    = [string]$cmbCat.SelectedItem
    $prefix = $txtPrefix.Text.Trim()
    $protect = $chkProtect.Checked
    $createCityAll = $chkCityAll.Checked
    $createRoleAll = $chkRoleAll.Checked
    $createGlobalAll = $chkGlobalAll.Checked
    $sep    = [char][string]$cmbSep.SelectedItem
    $nameTpl = $txtNameTpl.Text
    $descTpl = $txtDescTpl.Text

    if ([string]::IsNullOrWhiteSpace($baseDN)) { throw "Podaj Base DN." }
    if ($cities.Count -eq 0) { throw "Podaj co najmniej jedna nazwe OU (miasta)." }
    if ($roles.Count -eq 0) { throw "Podaj co najmniej jedna role/grupe." }
    if ($cityCodes.Count -ne $cities.Count) { throw "Podaj skroty miast (ta sama liczba wierszy co miast)." }
    if ($roleCodes.Count -ne $roles.Count) { throw "Podaj skroty dzialow (ta sama liczba wierszy co rol)." }
    if ([string]::IsNullOrWhiteSpace($prefix)) { throw "Podaj prefiks nazw grup." }
    if ([string]::IsNullOrWhiteSpace($nameTpl)) { $nameTpl = '{PREFIX}{SEP}{CITY_CODE}{SEP}{ROLE_CODE}' }

    $cityCodes = @($cityCodes | ForEach-Object { $_.ToUpperInvariant() })
    $roleCodes = @($roleCodes | ForEach-Object { $_.ToUpperInvariant() })

    return @{
        BaseDN = $baseDN
        Cities = $cities
        Roles  = $roles
        CityCodes = $cityCodes
        RoleCodes = $roleCodes
        Scope = $scope
        Category = $cat
        Prefix = $prefix
        Sep    = $sep
        NameTpl = $nameTpl
        DescTpl = $descTpl
        Protect = $protect
        CreateCityAll = $createCityAll
        CreateRoleAll = $createRoleAll
        CreateGlobalAll = $createGlobalAll
    }
}

$btnPreview.Add_Click({
    try {
        $vals = Get-UiValues
        $script:DoWhatIf = $true
        Invoke-Create -BaseDN $vals.BaseDN -Cities $vals.Cities -CityCodes $vals.CityCodes -ComputerRoles $vals.Roles -RoleCodes $vals.RoleCodes `
                      -GroupScope $vals.Scope -GroupCategory $vals.Category `
                      -GroupNamePrefix $vals.Prefix -SeparatorChar $vals.Sep `
                      -GroupNameTemplate $vals.NameTpl -DescriptionTemplate $vals.DescTpl `
                      -ProtectOUs:$vals.Protect -CreateCityAll:$vals.CreateCityAll -CreateRoleAll:$vals.CreateRoleAll -CreateGlobalAll:$vals.CreateGlobalAll
    } catch {
        Write-Log "[ERR] $($_.Exception.Message)"
    } finally {
        $script:DoWhatIf = $false
    }
})

$btnRun.Add_Click({
    try {
        $vals = Get-UiValues
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Czy na pewno utworzyc obiekty w AD?`r`nBase DN: $($vals.BaseDN)`r`nOU: $($vals.Cities -join ', ')",
            "Potwierdz",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

        $script:DoWhatIf = $false
        Invoke-Create -BaseDN $vals.BaseDN -Cities $vals.Cities -CityCodes $vals.CityCodes -ComputerRoles $vals.Roles -RoleCodes $vals.RoleCodes `
                      -GroupScope $vals.Scope -GroupCategory $vals.Category `
                      -GroupNamePrefix $vals.Prefix -SeparatorChar $vals.Sep `
                      -GroupNameTemplate $vals.NameTpl -DescriptionTemplate $vals.DescTpl `
                      -ProtectOUs:$vals.Protect -CreateCityAll:$vals.CreateCityAll -CreateRoleAll:$vals.CreateRoleAll -CreateGlobalAll:$vals.CreateGlobalAll
        [System.Windows.Forms.MessageBox]::Show("Operacja zakonczona. Sprawdz log.", "Gotowe",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    } catch {
        Write-Log "[ERR] $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Blad: $($_.Exception.Message)", "Blad",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } finally {
        $script:DoWhatIf = $false
    }
})

$btnClose.Add_Click({ $form.Close() })

# Sugestywne wartosci przykladowe - mozesz podmienic lub usunac:
#$txtBase.Text   = "OU=Company,DC=contoso,DC=local"
#$txtCities.Lines = @("Poznan","Warszawa")

[void]$form.ShowDialog()






