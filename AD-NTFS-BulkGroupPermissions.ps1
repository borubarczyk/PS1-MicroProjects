#Requires -Modules ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Ensure-ADModule {
    if (-not (Get-Module -ListAvailable ActiveDirectory)) { throw "Brak modułu ActiveDirectory (RSAT)." }
    Import-Module ActiveDirectory -ErrorAction Stop
}
function Test-Admin {
    $wi = [Security.Principal.WindowsIdentity]::GetCurrent()
    $wp = New-Object Security.Principal.WindowsPrincipal($wi)
    return $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# ===== LOG =====
$script:log = $null
function Write-LogUI([string]$Text){
    if ($script:log -is [System.Windows.Forms.RichTextBox]) {
        $script:log.AppendText("$(Get-Date -Format 'HH:mm:ss')  $Text`r`n")
        $script:log.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    } else { Write-Host $Text }
}

# ===== AD helpers =====
function Resolve-Group([string]$id){
    if([string]::IsNullOrWhiteSpace($id)){ return $null }
    $id = $id.Trim()
    $g = Get-ADGroup -LDAPFilter "(sAMAccountName=$id)" -ErrorAction SilentlyContinue
    if(-not $g){ $g = Get-ADGroup -LDAPFilter "(name=$id)" -ErrorAction SilentlyContinue }
    if(-not $g){ try{ $g = Get-ADGroup -Identity $id -ErrorAction Stop } catch{} }
    return $g
}
function Paste-IntoGrid([System.Windows.Forms.DataGridView]$Grid){
    $txt = [Windows.Forms.Clipboard]::GetText()
    if([string]::IsNullOrWhiteSpace($txt)){ return 0 }
    $lines = $txt -split "(`r`n|`n|`r)"
    $valid = 0
    foreach($ln in $lines){
        if([string]::IsNullOrWhiteSpace($ln)){ continue }
        $firstCol = ($ln -split "`t")[0].Trim()
        if([string]::IsNullOrWhiteSpace($firstCol)){ continue }
        $idx = $Grid.Rows.Add()
        $Grid.Rows[$idx].Cells['GroupId'].Value   = $firstCol
        $Grid.Rows[$idx].Cells['Perm'].Value      = 'Modyfikacja'
        $Grid.Rows[$idx].Cells['AppliesTo'].Value = 'Ten folder, podfoldery i pliki'
        $valid++
    }
    return $valid
}

# ===== mapy uprawnień i zakresów =====
$rightsMap = @{
    'Odczyt'                = [System.Security.AccessControl.FileSystemRights]::Read
    'Odczyt i wykonywanie'  = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute
    'Zapis'                 = [System.Security.AccessControl.FileSystemRights]::Write
    'Modyfikacja'           = [System.Security.AccessControl.FileSystemRights]::Modify
    'Pełna kontrola'        = [System.Security.AccessControl.FileSystemRights]::FullControl
}

# wartości liczbowe flag (bez bitowego OR na enumach w hash)
$IF = [System.Security.AccessControl.InheritanceFlags]
$PF = [System.Security.AccessControl.PropagationFlags]
$appliesMap = @{
    'Ten folder'                     = @([int]$IF::None,                [int]$PF::None)
    'Ten folder i podfoldery'        = @([int]$IF::ContainerInherit,    [int]$PF::None)
    'Ten folder i pliki'             = @([int]$IF::ObjectInherit,       [int]$PF::None)
    'Ten folder, podfoldery i pliki' = @([int](([int]$IF::ContainerInherit) -bor ([int]$IF::ObjectInherit)), [int]$PF::None)
    'Tylko podfoldery'               = @([int]$IF::ContainerInherit,    [int]$PF::InheritOnly)
    'Tylko pliki'                    = @([int]$IF::ObjectInherit,       [int]$PF::InheritOnly)
}

function Add-NTFSPerms {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Identity,
        [Parameter(Mandatory)][string]$RightsName,
        [Parameter(Mandatory)][string]$AppliesName,
        [switch]$PurgeExisting,
        [switch]$ProtectInheritance,
        [switch]$WhatIf
    )
    if(-not (Test-Path -LiteralPath $Path)){ throw "Ścieżka nie istnieje: $Path" }
    $rights = $rightsMap[$RightsName]; if(-not $rights){ throw "Nieznane uprawnienie: $RightsName" }
    $pair = $appliesMap[$AppliesName]; if(-not $pair){ throw "Nieznany zakres: $AppliesName" }

    $inh  = [System.Security.AccessControl.InheritanceFlags]$pair[0]
    $prop = [System.Security.AccessControl.PropagationFlags]$pair[1]

    if($WhatIf){
        Write-LogUI "[WhatIf] $Identity → '$Path' : $RightsName ; $AppliesName$(if($PurgeExisting){' ; purge'})$(if($ProtectInheritance){' ; protect'})"
        return
    }

    $acl = Get-Acl -LiteralPath $Path
    if($ProtectInheritance){ $acl.SetAccessRuleProtection($true, $true) }

    $nt = New-Object System.Security.Principal.NTAccount($Identity)
    if($PurgeExisting){ $acl.PurgeAccessRules($nt) | Out-Null }

    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule($nt, $rights, $inh, $prop, [System.Security.AccessControl.AccessControlType]::Allow)
    [void]$acl.AddAccessRule($rule)
    Set-Acl -LiteralPath $Path -AclObject $acl
}

# ===== GUI =====
function Build-Form {
    $form = New-Object Windows.Forms.Form
    $form.Text = "NTFS – nadawanie uprawnień wielu grupom"
    $form.Size = [Drawing.Size]::new(1200,860)       # szerzej i wyżej
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $font = New-Object Drawing.Font('Segoe UI',10)

    $lblPath = New-Object Windows.Forms.Label
    $lblPath.Text="Folder docelowy:"; $lblPath.Font=$font; $lblPath.AutoSize=$true; $lblPath.Location='10,12'
    $form.Controls.Add($lblPath)

    $script:tbPath = New-Object Windows.Forms.TextBox
    $script:tbPath.Location='130,10'; $script:tbPath.Width=960; $script:tbPath.Font=$font
    $form.Controls.Add($script:tbPath)

    $btnBrowse = New-Object Windows.Forms.Button
    $btnBrowse.Text='Wybierz...'; $btnBrowse.Font=$font; $btnBrowse.Location='1100,8'; $btnBrowse.Size=[Drawing.Size]::new(80,28)
    $form.Controls.Add($btnBrowse)

    $script:grid = New-Object Windows.Forms.DataGridView
    $script:grid.Location='10,50'; $script:grid.Size=[Drawing.Size]::new(1170,580); $script:grid.Font=$font
    $script:grid.AllowUserToAddRows=$true; $script:grid.RowHeadersVisible=$false; $script:grid.AutoSizeColumnsMode='Fill'
    $col1 = New-Object Windows.Forms.DataGridViewTextBoxColumn; $col1.HeaderText='Id grupy (Name/sAM/DN/SID)'; $col1.Name='GroupId'; $col1.FillWeight=220
    $col2 = New-Object Windows.Forms.DataGridViewTextBoxColumn; $col2.HeaderText='Rozpoznana (sAM / DN)'; $col2.Name='Resolved'; $col2.ReadOnly=$true; $col2.FillWeight=380
    $col3 = New-Object Windows.Forms.DataGridViewComboBoxColumn; $col3.HeaderText='Uprawnienie'; $col3.Name='Perm'; $col3.Items.AddRange(@('Odczyt','Odczyt i wykonywanie','Zapis','Modyfikacja','Pełna kontrola')); $col3.FillWeight=150
    $col4 = New-Object Windows.Forms.DataGridViewComboBoxColumn; $col4.HeaderText='Zakres (applies-to)'; $col4.Name='AppliesTo'
    $col4.Items.AddRange(@('Ten folder','Ten folder i podfoldery','Ten folder i pliki','Ten folder, podfoldery i pliki','Tylko podfoldery','Tylko pliki')); $col4.FillWeight=260
    $col5 = New-Object Windows.Forms.DataGridViewTextBoxColumn; $col5.HeaderText='Status'; $col5.Name='Result'; $col5.ReadOnly=$true; $col5.FillWeight=160
    $script:grid.Columns.AddRange([Windows.Forms.DataGridViewColumn[]]@($col1,$col2,$col3,$col4,$col5))
    $form.Controls.Add($script:grid)

    # ——— Masowe ustawienia ———
    $lblBulkPerm = New-Object Windows.Forms.Label
    $lblBulkPerm.Text='Ustaw wszystkim: Uprawnienie'; $lblBulkPerm.Font=$font; $lblBulkPerm.AutoSize=$true; $lblBulkPerm.Location='10,645'
    $form.Controls.Add($lblBulkPerm)

    $script:cmbBulkPerm = New-Object Windows.Forms.ComboBox
    $script:cmbBulkPerm.Location='220,642'; $script:cmbBulkPerm.Width=220; $script:cmbBulkPerm.Font=$font
    $script:cmbBulkPerm.DropDownStyle='DropDownList'
    $script:cmbBulkPerm.Items.AddRange(@('Odczyt','Odczyt i wykonywanie','Zapis','Modyfikacja','Pełna kontrola'))
    $script:cmbBulkPerm.SelectedItem = 'Modyfikacja'
    $form.Controls.Add($script:cmbBulkPerm)

    $btnBulkPerm = New-Object Windows.Forms.Button
    $btnBulkPerm.Text='Zastosuj'; $btnBulkPerm.Font=$font; $btnBulkPerm.Location='450,640'; $btnBulkPerm.Size=[Drawing.Size]::new(100,30)
    $form.Controls.Add($btnBulkPerm)

    $lblBulkAppl = New-Object Windows.Forms.Label
    $lblBulkAppl.Text='Ustaw wszystkim: Zakres'; $lblBulkAppl.Font=$font; $lblBulkAppl.AutoSize=$true; $lblBulkAppl.Location='570,645'
    $form.Controls.Add($lblBulkAppl)

    $script:cmbBulkAppl = New-Object Windows.Forms.ComboBox
    $script:cmbBulkAppl.Location='760,642'; $script:cmbBulkAppl.Width=320; $script:cmbBulkAppl.Font=$font
    $script:cmbBulkAppl.DropDownStyle='DropDownList'
    $script:cmbBulkAppl.Items.AddRange(@('Ten folder','Ten folder i podfoldery','Ten folder i pliki','Ten folder, podfoldery i pliki','Tylko podfoldery','Tylko pliki'))
    $script:cmbBulkAppl.SelectedItem = 'Ten folder, podfoldery i pliki'
    $form.Controls.Add($script:cmbBulkAppl)

    $btnBulkAppl = New-Object Windows.Forms.Button
    $btnBulkAppl.Text='Zastosuj'; $btnBulkAppl.Font=$font; $btnBulkAppl.Location='1090,640'; $btnBulkAppl.Size=[Drawing.Size]::new(90,30)   # teraz się mieści
    $form.Controls.Add($btnBulkAppl)

    # ——— Pozostałe przyciski/checkboxy ———
    $btnPaste = New-Object Windows.Forms.Button
    $btnPaste.Text='Wklej grupy'; $btnPaste.Font=$font; $btnPaste.Location='10,680'; $btnPaste.Size=[Drawing.Size]::new(150,32)
    $form.Controls.Add($btnPaste)

    $btnClear = New-Object Windows.Forms.Button
    $btnClear.Text='Wyczyść listę'; $btnClear.Font=$font; $btnClear.Location='170,680'; $btnClear.Size=[Drawing.Size]::new(150,32)
    $form.Controls.Add($btnClear)

    $btnResolve = New-Object Windows.Forms.Button
    $btnResolve.Text='Sprawdź w AD'; $btnResolve.Font=$font; $btnResolve.Location='330,680'; $btnResolve.Size=[Drawing.Size]::new(150,32)
    $form.Controls.Add($btnResolve)

    $script:cbPurge = New-Object Windows.Forms.CheckBox
    $script:cbPurge.Text='Purge (usuń istniejące ACE tej grupy)'; $script:cbPurge.Font=$font; $script:cbPurge.AutoSize=$true; $script:cbPurge.Location='500,684'
    $form.Controls.Add($script:cbPurge)

    $script:cbProtect = New-Object Windows.Forms.CheckBox
    $script:cbProtect.Text='Zablokuj dziedziczenie (zachowaj odziedziczone)'; $script:cbProtect.Font=$font; $script:cbProtect.AutoSize=$true; $script:cbProtect.Location='780,684'
    $form.Controls.Add($script:cbProtect)

    $script:cbWhatIf = New-Object Windows.Forms.CheckBox
    $script:cbWhatIf.Text='WhatIf (symulacja, bez zmian)'; $script:cbWhatIf.Font=$font; $script:cbWhatIf.AutoSize=$true; $script:cbWhatIf.Location='1070,684'
    $form.Controls.Add($script:cbWhatIf)

    $script:btnApply = New-Object Windows.Forms.Button
    $script:btnApply.Text='Nadaj uprawnienia'; $script:btnApply.Font=$font; $script:btnApply.Location='970,720'; $script:btnApply.Size=[Drawing.Size]::new(210,38)
    $script:btnApply.Enabled=$false
    $form.Controls.Add($script:btnApply)

    # Log – wyższy
    $script:log = New-Object Windows.Forms.RichTextBox
    $script:log.Location='10,720'; $script:log.Size=[Drawing.Size]::new(950,120); $script:log.Font=$font; $script:log.ReadOnly=$true
    $form.Controls.Add($script:log)

    $btnClearLog = New-Object Windows.Forms.Button
    $btnClearLog.Text='Wyczyść log'; $btnClearLog.Font=$font; $btnClearLog.Location='970,765'; $btnClearLog.Size=[Drawing.Size]::new(210,28)
    $form.Controls.Add($btnClearLog)

    # ==== Handlers ====
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Wybierz folder"
        if($script:tbPath.Text -and (Test-Path -LiteralPath $script:tbPath.Text)){ $dlg.SelectedPath = $script:tbPath.Text }  # start z poprzedniej lokalizacji
        if($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
            $script:tbPath.Text = $dlg.SelectedPath
            $script:btnApply.Enabled = $true
            Write-LogUI "Wybrano folder: $($script:tbPath.Text)"
        }
    })
    $btnPaste.Add_Click({
        $n = Paste-IntoGrid -Grid $script:grid
        if($n -gt 0){ Write-LogUI "Wklejono $n grup." ; $script:btnApply.Enabled=$true }
    })
    $btnClear.Add_Click({ $script:grid.Rows.Clear(); Write-LogUI "Wyczyszczono listę grup." })
    $btnClearLog.Add_Click({ $script:log.Clear() })

    $btnResolve.Add_Click({
        for($i=0;$i -lt $script:grid.Rows.Count;$i++){
            $r = $script:grid.Rows[$i]; if($r.IsNewRow){ continue }
            $id = "$($r.Cells['GroupId'].Value)".Trim()
            if(-not $id){ continue }
            $g = Resolve-Group $id
            if($g){
                $r.Cells['Resolved'].Value = "$($g.SamAccountName)  [$($g.DistinguishedName)]"
                if(-not $r.Cells['Perm'].Value){ $r.Cells['Perm'].Value = 'Modyfikacja' }
                if(-not $r.Cells['AppliesTo'].Value){ $r.Cells['AppliesTo'].Value = 'Ten folder, podfoldery i pliki' }
                $r.DefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(220,255,220)
            } else {
                $r.Cells['Resolved'].Value = "NIE ZNALEZIONO"
                $r.DefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(255,230,230)
            }
        }
        Write-LogUI "Sprawdzono grupy w AD."
        $script:btnApply.Enabled = $true
    })

    # ——— Masowe zastosowanie ———
    $btnBulkPerm.Add_Click({
        $val = "$($script:cmbBulkPerm.SelectedItem)"; if(-not $val){ return }
        for($i=0;$i -lt $script:grid.Rows.Count;$i++){ $r=$script:grid.Rows[$i]; if($r.IsNewRow){continue}; $r.Cells['Perm'].Value=$val }
        Write-LogUI "Ustawiono wszystkim uprawnienie: $val."
    })
    $btnBulkAppl.Add_Click({
        $val = "$($script:cmbBulkAppl.SelectedItem)"; if(-not $val){ return }
        for($i=0;$i -lt $script:grid.Rows.Count;$i++){ $r=$script:grid.Rows[$i]; if($r.IsNewRow){continue}; $r.Cells['AppliesTo'].Value=$val }
        Write-LogUI "Ustawiono wszystkim zakres: $val."
    })

    $script:btnApply.Add_Click({
        try{
            if(-not (Test-Admin) -and -not $script:cbWhatIf.Checked){
                [System.Windows.Forms.MessageBox]::Show("Uruchom PowerShell jako Administrator.","Brak uprawnień") | Out-Null; return
            }
            $path = $script:tbPath.Text.Trim()
            if(-not $path){ [System.Windows.Forms.MessageBox]::Show("Wybierz folder.","Brak ścieżki") | Out-Null; return }

            $ok=0; $err=0
            for($i=0;$i -lt $script:grid.Rows.Count;$i++){
                $r = $script:grid.Rows[$i]; if($r.IsNewRow){ continue }
                $id = "$($r.Cells['GroupId'].Value)".Trim()
                if(-not $id){ continue }
                $perm = "$($r.Cells['Perm'].Value)"
                $appl = "$($r.Cells['AppliesTo'].Value)"
                if(-not $perm -or -not $appl){ $r.Cells['Result'].Value="Brak uprawnień/zakresu"; $err++; continue }

                $g = Resolve-Group $id
                if(-not $g){ $r.Cells['Result'].Value="Grupa nieznaleziona"; $err++; continue }

                try{
                    Add-NTFSPerms -Path $path -Identity $g.SamAccountName -RightsName $perm -AppliesName $appl -PurgeExisting:$script:cbPurge.Checked -ProtectInheritance:$script:cbProtect.Checked -WhatIf:$script:cbWhatIf.Checked
                    $r.Cells['Result'].Value = $(if($script:cbWhatIf.Checked){'WhatIf'}else{'OK'})
                    if(-not $script:cbWhatIf.Checked){ $ok++ }
                } catch {
                    $r.Cells['Result'].Value = "Błąd: $($_.Exception.Message)"; $err++
                }
            }
            Write-LogUI "Zakończono: OK=$ok, Błędy=$err na ścieżce '$path'."
            if(-not $script:cbWhatIf.Checked -and $err -eq 0){
                [System.Windows.Forms.MessageBox]::Show("Uprawnienia nadane.","Gotowe",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
            }
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message,"Błąd",[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    return $form
}

# ===== MAIN =====
try{
    Ensure-ADModule
    $form = Build-Form
    [void]$form.ShowDialog()
} catch {
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,"Błąd krytyczny",[System.Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    throw
}
