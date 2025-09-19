#Requires -Modules ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Ensure-ADModule {
    if (-not (Get-Module -ListAvailable ActiveDirectory)) {
        throw "Brak modułu ActiveDirectory (RSAT)."
    }
    Import-Module ActiveDirectory -ErrorAction Stop
}

# === GLOBAL: LOG ===
$script:log = $null
function Write-LogUI([string]$Text){
    if ($script:log -is [System.Windows.Forms.RichTextBox]) {
        $script:log.AppendText("$(Get-Date -Format 'HH:mm:ss')  $Text`r`n")
        $script:log.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    } else {
        Write-Host $Text
    }
}

# === GLOBAL: AD RESOLVERS ===
function Resolve-User([string]$id){
    if([string]::IsNullOrWhiteSpace($id)){ return $null }
    $id = $id.Trim()
    $u = Get-ADUser -LDAPFilter "(sAMAccountName=$id)" -ErrorAction SilentlyContinue
    if(-not $u){ $u = Get-ADUser -LDAPFilter "(userPrincipalName=$id)" -ErrorAction SilentlyContinue }
    if(-not $u){ try{ $u = Get-ADUser -Identity $id -ErrorAction Stop } catch{} }
    return $u
}
function Resolve-Group([string]$id){
    if([string]::IsNullOrWhiteSpace($id)){ return $null }
    $id = $id.Trim()
    $g = Get-ADGroup -LDAPFilter "(sAMAccountName=$id)" -ErrorAction SilentlyContinue
    if(-not $g){ $g = Get-ADGroup -LDAPFilter "(name=$id)" -ErrorAction SilentlyContinue }
    if(-not $g){ try{ $g = Get-ADGroup -Identity $id -ErrorAction Stop } catch{} }
    return $g
}

# === GLOBAL: WKLEJANIE Z EXCELA (1. kolumna) ===
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
        $Grid.Rows[$idx].Cells[0].Value = $firstCol
        $valid++
    }
    return $valid
}

function Build-Form {
    $form = New-Object Windows.Forms.Form
    $form.Text = "Dodawanie użytkowników do grup (AD)"
    $form.Size = [Drawing.Size]::new(1100,760)     # wyższe okno
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $font = New-Object Drawing.Font('Segoe UI',10)

    $lblU = New-Object Windows.Forms.Label
    $lblU.Text="Użytkownicy (UPN / sAM / DN / SID)"; $lblU.Font=$font; $lblU.AutoSize=$true; $lblU.Location='10,10'
    $form.Controls.Add($lblU)

    $lblG = New-Object Windows.Forms.Label
    $lblG.Text="Grupy (Name / sAM / DN / SID)"; $lblG.Font=$font; $lblG.AutoSize=$true; $lblG.Location='560,10'
    $form.Controls.Add($lblG)

    # === GRIDS ===
    $script:gridU = New-Object Windows.Forms.DataGridView
    $script:gridU.Location='10,35'; $script:gridU.Size=[Drawing.Size]::new(520,470); $script:gridU.Font=$font
    $script:gridU.AllowUserToAddRows=$true; $script:gridU.RowHeadersVisible=$false; $script:gridU.AutoSizeColumnsMode='Fill'
    $colU1 = New-Object Windows.Forms.DataGridViewTextBoxColumn; $colU1.HeaderText='Id użytkownika'; $colU1.Name='UserId'; $colU1.FillWeight=160
    $colU2 = New-Object Windows.Forms.DataGridViewTextBoxColumn; $colU2.HeaderText='Rozpoznany (sAM / DN)'; $colU2.Name='Resolved'; $colU2.ReadOnly=$true
    $script:gridU.Columns.AddRange([Windows.Forms.DataGridViewColumn[]]@($colU1,$colU2))
    $form.Controls.Add($script:gridU)

    $script:gridG = New-Object Windows.Forms.DataGridView
    $script:gridG.Location='560,35'; $script:gridG.Size=[Drawing.Size]::new(520,470); $script:gridG.Font=$font
    $script:gridG.AllowUserToAddRows=$true; $script:gridG.RowHeadersVisible=$false; $script:gridG.AutoSizeColumnsMode='Fill'
    $colG1 = New-Object Windows.Forms.DataGridViewTextBoxColumn; $colG1.HeaderText='Id grupy'; $colG1.Name='GroupId'; $colG1.FillWeight=160
    $colG2 = New-Object Windows.Forms.DataGridViewTextBoxColumn; $colG2.HeaderText='Rozpoznana (sAM / DN)'; $colG2.Name='Resolved'; $colG2.ReadOnly=$true
    $script:gridG.Columns.AddRange([Windows.Forms.DataGridViewColumn[]]@($colG1,$colG2))
    $form.Controls.Add($script:gridG)

    # === PRZYCISKI ===
    $btnPasteU = New-Object Windows.Forms.Button
    $btnPasteU.Text='Wklej użytkowników'; $btnPasteU.Font=$font; $btnPasteU.Location='10,515'; $btnPasteU.Size=[Drawing.Size]::new(200,32)
    $form.Controls.Add($btnPasteU)

    $btnClearU = New-Object Windows.Forms.Button
    $btnClearU.Text='Wyczyść użytkowników'; $btnClearU.Font=$font; $btnClearU.Location='220,515'; $btnClearU.Size=[Drawing.Size]::new(200,32)
    $form.Controls.Add($btnClearU)

    $btnPasteG = New-Object Windows.Forms.Button
    $btnPasteG.Text='Wklej grupy'; $btnPasteG.Font=$font; $btnPasteG.Location='560,515'; $btnPasteG.Size=[Drawing.Size]::new(200,32)
    $form.Controls.Add($btnPasteG)

    $btnClearG = New-Object Windows.Forms.Button
    $btnClearG.Text='Wyczyść grupy'; $btnClearG.Font=$font; $btnClearG.Location='770,515'; $btnClearG.Size=[Drawing.Size]::new(200,32)
    $form.Controls.Add($btnClearG)

    $btnResolve = New-Object Windows.Forms.Button
    $btnResolve.Text='Sprawdź w AD'; $btnResolve.Font=$font; $btnResolve.Location='10,555'; $btnResolve.Size=[Drawing.Size]::new(200,34)
    $form.Controls.Add($btnResolve)

    $script:rbAll = New-Object Windows.Forms.RadioButton
    $script:rbAll.Text="Każdy użytkownik → każda grupa (A×B)"; $script:rbAll.Font=$font; $script:rbAll.AutoSize=$true; $script:rbAll.Location='220,560'; $script:rbAll.Checked=$true
    $form.Controls.Add($script:rbAll)

    $script:rbPair = New-Object Windows.Forms.RadioButton
    $script:rbPair.Text="Wiersz-do-wiersza (U1→G1, U2→G2, …)"; $script:rbPair.Font=$font; $script:rbPair.AutoSize=$true; $script:rbPair.Location='520,560'
    $form.Controls.Add($script:rbPair)

    $script:cbWhatIf = New-Object Windows.Forms.CheckBox
    $script:cbWhatIf.Text='Dry-run (WhatIf)'; $script:cbWhatIf.Font=$font; $script:cbWhatIf.AutoSize=$true; $script:cbWhatIf.Location='800,560'
    $form.Controls.Add($script:cbWhatIf)

    $script:btnRun = New-Object Windows.Forms.Button
    $script:btnRun.Text='Dodaj do grup'; $script:btnRun.Font=$font; $script:btnRun.Location='920,555'; $script:btnRun.Size=[Drawing.Size]::new(160,38)
    $script:btnRun.Enabled = $false
    $form.Controls.Add($script:btnRun)

    # === LOG – dużo większy ===
    $script:log = New-Object Windows.Forms.RichTextBox
    $script:log.Location='10,600'
    $script:log.Size=[Drawing.Size]::new(1070,110)   # było 35 → teraz 110
    $script:log.Font=$font; $script:log.ReadOnly=$true
    $form.Controls.Add($script:log)

    $btnClearLog = New-Object Windows.Forms.Button
    $btnClearLog.Text='Wyczyść log'; $btnClearLog.Font=$font; $btnClearLog.Location='10,715'; $btnClearLog.Size=[Drawing.Size]::new(120,28)
    $form.Controls.Add($btnClearLog)

    # === HANDLERY ===
    $btnClearU.Add_Click({
        $script:gridU.Rows.Clear()
        Write-LogUI "Wyczyszczono listę użytkowników."
    })
    $btnClearG.Add_Click({
        $script:gridG.Rows.Clear()
        Write-LogUI "Wyczyszczono listę grup."
    })
    $btnClearLog.Add_Click({ $script:log.Clear() })

    $btnPasteU.Add_Click({
        $n = Paste-IntoGrid -Grid $script:gridU
        if($n -gt 0){ $script:btnRun.Enabled = $true; Write-LogUI "Wklejono $n użytkowników." }
    })
    $btnPasteG.Add_Click({
        $n = Paste-IntoGrid -Grid $script:gridG
        if($n -gt 0){ $script:btnRun.Enabled = $true; Write-LogUI "Wklejono $n grup." }
    })

    $btnResolve.Add_Click({
        try{
            for($i=0;$i -lt $script:gridU.Rows.Count;$i++){
                $r = $script:gridU.Rows[$i]; if($r.IsNewRow){ continue }
                $id = "$($r.Cells['UserId'].Value)".Trim()
                if(-not $id){ continue }
                $u = Resolve-User $id
                if($u){
                    $r.Cells['Resolved'].Value = "$($u.SamAccountName)  [$($u.DistinguishedName)]"
                    $r.DefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(220,255,220)
                } else {
                    $r.Cells['Resolved'].Value = "NIE ZNALEZIONO"
                    $r.DefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(255,230,230)
                }
            }
            for($i=0;$i -lt $script:gridG.Rows.Count;$i++){
                $r = $script:gridG.Rows[$i]; if($r.IsNewRow){ continue }
                $id = "$($r.Cells['GroupId'].Value)".Trim()
                if(-not $id){ continue }
                $g = Resolve-Group $id
                if($g){
                    $r.Cells['Resolved'].Value = "$($g.SamAccountName)  [$($g.DistinguishedName)]"
                    $r.DefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(220,255,220)
                } else {
                    $r.Cells['Resolved'].Value = "NIE ZNALEZIONO"
                    $r.DefaultCellStyle.BackColor = [Drawing.Color]::FromArgb(255,230,230)
                }
            }
            Write-LogUI "Sprawdzono obiekty w AD."
            $script:btnRun.Enabled = $true
        } catch {
            [Windows.Forms.MessageBox]::Show($_.Exception.Message,"Błąd sprawdzania",[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        }
    })

    $script:btnRun.Add_Click({
        $script:btnRun.Enabled = $false
        try{
            $users = @()
            foreach($r in $script:gridU.Rows){
                if($r.IsNewRow){ continue }
                $id = "$($r.Cells['UserId'].Value)".Trim()
                if($id){ $u = Resolve-User $id; if($u){ $users += $u } else { Write-LogUI "Użytkownik '$id' nie znaleziony – pomijam." } }
            }
            $groups = @()
            foreach($r in $script:gridG.Rows){
                if($r.IsNewRow){ continue }
                $id = "$($r.Cells['GroupId'].Value)".Trim()
                if($id){ $g = Resolve-Group $id; if($g){ $groups += $g } else { Write-LogUI "Grupa '$id' nie znaleziona – pomijam." } }
            }

            if(-not $users.Count){ Write-LogUI "Brak poprawnych użytkowników."; return }
            if(-not $groups.Count){ Write-LogUI "Brak poprawnych grup."; return }

            $whatIf = $script:cbWhatIf.Checked

            if($script:rbAll.Checked){
                foreach($g in $groups){
                    $ids = $users | ForEach-Object { $_.DistinguishedName }
                    try{
                        if($whatIf){
                            Write-LogUI "[WhatIf] Add-ADGroupMember '$($g.SamAccountName)' ← $($users.Count) użytk."
                        } else {
                            Add-ADGroupMember -Identity $g -Members $ids -ErrorAction Stop
                            Write-LogUI "OK: dodano $($users.Count) użytkowników do '$($g.SamAccountName)'."
                        }
                    } catch {
                        Write-LogUI "BŁĄD dla grupy '$($g.SamAccountName)': $($_.Exception.Message)"
                    }
                }
            } else {
                $n = [Math]::Min($users.Count,$groups.Count)
                for($i=0;$i -lt $n; $i++){
                    $u = $users[$i]; $g = $groups[$i]
                    try{
                        if($whatIf){
                            Write-LogUI "[WhatIf] Add-ADGroupMember '$($g.SamAccountName)' ← '$($u.SamAccountName)'."
                        } else {
                            Add-ADGroupMember -Identity $g -Members $u.DistinguishedName -ErrorAction Stop
                            Write-LogUI "OK: '$($u.SamAccountName)' → '$($g.SamAccountName)'."
                        }
                    } catch {
                        Write-LogUI "BŁĄD: '$($u.SamAccountName)' → '$($g.SamAccountName)': $($_.Exception.Message)"
                    }
                }
                if($users.Count -ne $groups.Count){
                    Write-LogUI "Uwaga: różna liczba użytkowników i grup. Sparowano $n wierszy."
                }
            }
            [Windows.Forms.MessageBox]::Show("Zakończono operację.","Gotowe",[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } finally {
            $script:btnRun.Enabled = $true
        }
    })

    return $form
}

# === MAIN ===
try{
    Ensure-ADModule
    $form = Build-Form
    [void]$form.ShowDialog()
} catch {
    [Windows.Forms.MessageBox]::Show($_.Exception.Message,"Błąd",[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    throw
}
