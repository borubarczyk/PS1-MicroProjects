#Requires -Modules ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===================== PICKER OU =====================
function Show-OUChooser {
    param(
        [string]$Title = "Wybierz OU",
        [string]$SearchBase,
        [switch]$RequireUnderBase,
        [string]$PreselectByName
    )

    $script:OUChooser_LastError = $null
    $script:OUChooser_LastDialogResult = $null
    $form               = New-Object System.Windows.Forms.Form
    $form.Text          = $Title
    $form.StartPosition = 'CenterParent'
    $form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.Font          = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.MinimumSize   = New-Object System.Drawing.Size(640,700)
    $form.Size          = New-Object System.Drawing.Size(720,760)

    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'; $root.ColumnCount = 1; $root.RowCount = 4
    [void]$root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))
    [void]$root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::Percent,100))
    [void]$root.RowStyles.Add([System.Windows.Forms.RowStyle]::new([System.Windows.Forms.SizeType]::AutoSize))

    $lblFilter = New-Object System.Windows.Forms.Label
    $lblFilter.Text = "Filtr nazwy OU:"; $lblFilter.AutoSize = $true; $lblFilter.Padding = '8,8,8,0'

    $tbFilter = New-Object System.Windows.Forms.TextBox
    $tbFilter.Dock = 'Top'; $tbFilter.Margin = '8,0,8,8'; $tbFilter.Width = 400

    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Dock = 'Fill'; $tree.HideSelection = $false
    $tree.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    $panelButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $panelButtons.Dock='Top'; $panelButtons.AutoSize=$true
    $panelButtons.FlowDirection=[System.Windows.Forms.FlowDirection]::RightToLeft
    $panelButtons.Padding='8,8,8,8'; $panelButtons.WrapContents=$false

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text="OK"; $btnOK.AutoSize=$false; $btnOK.Width=120; $btnOK.Height=32; $btnOK.Margin='6,0,0,0'
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text="Anuluj"; $btnCancel.AutoSize=$false; $btnCancel.Width=120; $btnCancel.Height=32

    $panelButtons.Controls.Add($btnOK) | Out-Null
    $panelButtons.Controls.Add($btnCancel) | Out-Null

    $root.Controls.Add($lblFilter,0,0)
    $root.Controls.Add($tbFilter,0,1)
    $root.Controls.Add($tree,0,2)
    $root.Controls.Add($panelButtons,0,3)
    $form.Controls.Add($root)
    $form.AcceptButton = $btnOK; $form.CancelButton = $btnCancel

    function Get-ParentDn([string]$dn) { if ($dn -and $dn.Contains(',')) { return $dn.Substring($dn.IndexOf(',')+1) } else { return $null } }

    try {
        $ouList = @()
        $baseOu = $null

        if ($SearchBase) {
            $baseOu = Get-ADOrganizationalUnit -Identity $SearchBase -Properties Name,DistinguishedName -ErrorAction Stop
            $ouList += $baseOu
            $ouList += Get-ADOrganizationalUnit -Filter * -SearchBase $SearchBase -SearchScope Subtree -ErrorAction Stop
        } else {
            $ouList = Get-ADOrganizationalUnit -Filter * -SearchScope Subtree -ErrorAction Stop
        }

        if (-not $ouList -or $ouList.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Nie znaleziono żadnych OU do wyświetlenia.","Uwaga",'OK','Information') | Out-Null
            return $null
        }

        $nodes = @{}
        foreach ($ou in $ouList) {
            if (-not $ou.DistinguishedName) { continue }
            $nodes[$ou.DistinguishedName] = New-Object System.Windows.Forms.TreeNode -Property @{
                Text = $ou.Name
                Tag  = $ou.DistinguishedName
            }
        }

        foreach ($ou in $ouList) {
            $dn = [string]$ou.DistinguishedName
            $parentDn = Get-ParentDn $dn
            if ($parentDn -and $nodes.ContainsKey($parentDn)) {
                [void]$nodes[$parentDn].Nodes.Add($nodes[$dn])
            }
        }

        $tree.BeginUpdate()
        $tree.Nodes.Clear()

        if ($SearchBase -and $baseOu -and $nodes.ContainsKey($baseOu.DistinguishedName)) {
            [void]$tree.Nodes.Add($nodes[$baseOu.DistinguishedName])
            if ($tree.Nodes.Count -gt 0) { $tree.Nodes[0].Expand() }
        } else {
            foreach ($kv in $nodes.GetEnumerator()) {
                $dn = $kv.Key
                $parentDn = Get-ParentDn $dn
                if (-not $nodes.ContainsKey($parentDn)) { [void]$tree.Nodes.Add($kv.Value) }
            }
            foreach ($n in $tree.Nodes) { $n.Expand() }
        }
        $tree.EndUpdate()
    } catch {
        $script:OUChooser_LastError = $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show("Nie udało się pobrać OU: $($_.Exception.Message)","Błąd",'OK','Error') | Out-Null
        return $null
    }

    $tree.Add_NodeMouseDoubleClick({ param($s,$e) if ($e -and $e.Node) { $tree.SelectedNode = $e.Node; $btnOK.PerformClick() } })
    $tree.Add_DoubleClick({ if ($tree.SelectedNode) { $btnOK.PerformClick() } })
    $tree.Add_KeyDown({ param($s,$e) if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter -and $tree.SelectedNode) { $e.Handled=$true; $btnOK.PerformClick() } })

    if ($PreselectByName) {
        $stack = New-Object System.Collections.Stack
        foreach ($rootNode in $tree.Nodes) { $stack.Push($rootNode) }
        $found = $null
        while ($stack.Count -gt 0 -and -not $found) {
            $node = [System.Windows.Forms.TreeNode]$stack.Pop()
            if ($node.Text -ieq $PreselectByName) { $found = $node; break }
            foreach ($c in $node.Nodes) { $stack.Push($c) }
        }
        if ($found) { $tree.SelectedNode = $found; $found.EnsureVisible() }
    }

    $tbFilter.Add_TextChanged({
        $q = $tbFilter.Text
        if ([string]::IsNullOrWhiteSpace($q)) {
            foreach ($n in $tree.Nodes) { $n.ExpandAll() }
        } else {
            function ShowMatches($node){
                $isMatch = $node.Text -like "*$q*"
                $childHas = $false
                foreach ($c in $node.Nodes) { if (ShowMatches $c) { $childHas = $true } }
                if ($isMatch -or $childHas) { $node.Expand(); return $true } else { $node.Collapse(); return $false }
            }
            foreach ($n in $tree.Nodes) { ShowMatches $n | Out-Null }
        }
    })

    $selectedDN = $null
    $btnOK.Add_Click({
        if (-not $tree.SelectedNode -or -not $tree.SelectedNode.Tag) {
            [System.Windows.Forms.MessageBox]::Show("Wybierz konkretne OU z drzewa.","Uwaga",'OK','Information') | Out-Null
            return
        }
        $dn = [string]$tree.SelectedNode.Tag

        if ($RequireUnderBase -and $SearchBase) {
            $escBase = [regex]::Escape($SearchBase)
            $isDesc = ($dn -imatch ",$escBase$") -and ($dn -ine $SearchBase)
            if (-not $isDesc) {
                [System.Windows.Forms.MessageBox]::Show("Wybierz OU, które jest POTOMKIEM bazowego OU.","Uwaga",'OK','Information') | Out-Null
                return
            }
        }

        $selectedDN = $dn
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $btnCancel.Add_Click({ $script:OUChooser_LastError = $null;  $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.Close() })

    $dialogResult = $form.ShowDialog()
    $script:OUChooser_LastDialogResult = $dialogResult
    if (-not $selectedDN -and $dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($tree.SelectedNode -and $tree.SelectedNode.Tag) {
            $selectedDN = [string]$tree.SelectedNode.Tag
        }
    }
    return $selectedDN
}

# ===================== GŁÓWNE OKNO =====================
function Show-MoveDisabledUsersGUI {
    # ===================== LISTY OCHRONNE (DOMYŚLNE) =====================
    $script:ExceptionUsers = @(
        'Guest','krbtgt','Gość','Admin2','Konto domyślne','any connect'
    )
    $script:ProtectedGroups = @(
        'Domain Users','Administrators','Domain Admins','Enterprise Admins',
        'Schema Admins','Protected Users','DnsAdmins','Backup Operators',
        'Account Operators','Server Operators','Print Operators',
        'Read-only Domain Controllers'
    )

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        [System.Windows.Forms.MessageBox]::Show("Brak modułu ActiveDirectory. Zainstaluj RSAT / AD PowerShell Tools.","Błąd",'OK','Error') | Out-Null
        return
    }
    Import-Module ActiveDirectory -ErrorAction Stop

    $form              = New-Object System.Windows.Forms.Form
    $form.Text         = "Przenoszenie zablokowanych użytkowników AD"
    $form.StartPosition= 'CenterScreen'
    $form.AutoScaleMode= [System.Windows.Forms.AutoScaleMode]::Dpi
    $form.Font         = New-Object System.Drawing.Font("Segoe UI", 10)
    $form.MinimumSize  = New-Object System.Drawing.Size(1000,760)
    $form.Size         = New-Object System.Drawing.Size(1150,820)

    # ===================== ZARZĄDZANIE KONFIGURACJĄ =====================
    function Get-ConfigPath {
        try {
            $scriptPath = if ($PSCommandPath) { $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
            $scriptDir  = Split-Path -Parent $scriptPath
            return (Join-Path $scriptDir '.move-disabled-users-ad.config.json')
        } catch { return (Join-Path (Get-Location).Path '.move-disabled-users-ad.config.json') }
    }

    function Load-Config {
        $path = Get-ConfigPath
        if (Test-Path -LiteralPath $path) {
            try {
                $config = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
                if ($config.ExceptionUsers) { $script:ExceptionUsers = @($config.ExceptionUsers) }
                if ($config.ProtectedGroups) { $script:ProtectedGroups = @($config.ProtectedGroups) }
                return $config
            } catch {
                # Błąd odczytu lub parsowania, użyj domyślnych
            }
        }
        return $null
    }

    function Save-Config {
        param(
            [string]$BaseOU,
            [string]$TargetOU
        )
        $obj = [PSCustomObject]@{
            BaseOU = $BaseOU
            TargetOU = $TargetOU
            ExceptionUsers = $script:ExceptionUsers
            ProtectedGroups = $script:ProtectedGroups
            LastUsed = (Get-Date).ToString('o')
            Version = '2.0'
        }
        $json = $obj | ConvertTo-Json -Depth 3
        try {
            Set-Content -Path (Get-ConfigPath) -Value $json -Encoding UTF8 -Force
            return $true
        } catch {
            return $false
        }
    }
    
    # Inicjalizacja konfiguracji na starcie
    $cfg = Load-Config

    # ===================== GŁÓWNY LAYOUT (TABCONTROL) =====================
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = 'Fill'
    
    $tabOperations = New-Object System.Windows.Forms.TabPage
    $tabOperations.Text = "Operacje"
    $tabOperations.Padding = '8,8,8,8'

    $tabSettings = New-Object System.Windows.Forms.TabPage
    $tabSettings.Text = "Ustawienia"
    $tabSettings.Padding = '8,8,8,8'

    $tabControl.Controls.Add($tabOperations)
    $tabControl.Controls.Add($tabSettings)
    $form.Controls.Add($tabControl)

    # ===================== ZAKŁADKA "OPERACJE" =====================
    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = 'Fill'
    $root.ColumnCount = 1
    $root.RowCount = 6
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,60)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    [void]$root.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent,40)) )
    $tabOperations.Controls.Add($root)

    # ---- BAZA (wiersz 0) ----
    $pBase = New-Object System.Windows.Forms.TableLayoutPanel
    $pBase.Dock='Fill'; $pBase.ColumnCount=3; $pBase.AutoSize=$true
    $pBase.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    $pBase.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)) )
    $pBase.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )

    $lblBase = New-Object System.Windows.Forms.Label; $lblBase.Text = "Bazowe OU (drzewo NIEAKTYWNI):"; $lblBase.AutoSize = $true
    $tbBase = New-Object System.Windows.Forms.TextBox; $tbBase.ReadOnly = $true; $tbBase.Dock = 'Fill'; $tbBase.Margin = '0,0,8,0'
    $btnPickBase = New-Object System.Windows.Forms.Button; $btnPickBase.Text = "Wybierz bazowe…"; $btnPickBase.AutoSize = $false; $btnPickBase.Width = 140; $btnPickBase.Height = 32

    $pBase.Controls.Add($lblBase,0,0); $pBase.Controls.Add($tbBase,1,0); $pBase.Controls.Add($btnPickBase,2,0)

    $lblStatus = New-Object System.Windows.Forms.Label; $lblStatus.Text = "Status: brak bazowego OU"; $lblStatus.AutoSize = $true; $lblStatus.Padding = '0,4,0,8'
    $pBaseWrap = New-Object System.Windows.Forms.FlowLayoutPanel; $pBaseWrap.Dock='Fill'; $pBaseWrap.AutoSize=$false; $pBaseWrap.FlowDirection='TopDown'; $pBaseWrap.WrapContents = $false
    $pBaseWrap.Controls.Add($pBase); $pBaseWrap.Controls.Add($lblStatus)

    # ---- SKAN (wiersz 1) ----
    $btnScan = New-Object System.Windows.Forms.Button; $btnScan.Text = "Skanuj zablokowanych (spoza bazowego OU)"; $btnScan.AutoSize = $false; $btnScan.Width = 360; $btnScan.Height = 32; $btnScan.Margin = '8,4,0,8'; $btnScan.Enabled = $false

    # ---- GRID (wiersz 2) ----
    $grid = New-Object System.Windows.Forms.DataGridView; $grid.Dock = 'Fill'; $grid.ReadOnly = $false; $grid.AllowUserToAddRows = $false; $grid.SelectionMode = 'FullRowSelect'; $grid.MultiSelect = $true; $grid.AutoGenerateColumns = $false; $grid.RowHeadersVisible = $false; $grid.Font = New-Object System.Drawing.Font("Segoe UI", 10); $grid.AutoSizeColumnsMode = 'AllCells'
    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $colCheck.HeaderText = "Wybierz"; $colCheck.Width = 70; $grid.Columns.Add($colCheck)
    $colName = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colName.HeaderText = "Name"; $colName.DataPropertyName = "Name"; $colName.AutoSizeMode = 'Fill'; $grid.Columns.Add($colName)
    $colSam = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colSam.HeaderText = "SamAccountName"; $colSam.DataPropertyName = "SamAccountName"; $colSam.Width = 180; $grid.Columns.Add($colSam)
    $colDn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDn.HeaderText = "DistinguishedName"; $colDn.DataPropertyName = "DistinguishedName"; $colDn.AutoSizeMode = 'Fill'; $grid.Columns.Add($colDn)

    # ---- TARGET (wiersz 3) ----
    $pTarget = New-Object System.Windows.Forms.TableLayoutPanel; $pTarget.Dock='Top'; $pTarget.ColumnCount=3; $pTarget.AutoSize=$true
    $pTarget.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    $pTarget.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,100)) )
    $pTarget.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    $lblTarget = New-Object System.Windows.Forms.Label; $lblTarget.Text = "Docelowe OU:"; $lblTarget.AutoSize = $true
    $tbTarget = New-Object System.Windows.Forms.TextBox; $tbTarget.ReadOnly = $true; $tbTarget.Dock = 'Fill'
    $btnPickTarget = New-Object System.Windows.Forms.Button; $btnPickTarget.Text= "Wybierz docelowe…"; $btnPickTarget.AutoSize = $false; $btnPickTarget.Width = 140; $btnPickTarget.Height = 32; $btnPickTarget.Enabled = $false
    $pTarget.Controls.Add($lblTarget,0,0); $pTarget.Controls.Add($tbTarget,1,0); $pTarget.Controls.Add($btnPickTarget,2,0)

    # ---- AKCJE (wiersz 4) ----
    $pActions = New-Object System.Windows.Forms.FlowLayoutPanel; $pActions.Dock='Top'; $pActions.AutoSize = $true; $pActions.WrapContents = $false
    $btnSelectAll = New-Object System.Windows.Forms.Button; $btnSelectAll.Text = "Zaznacz wszystkie"; $btnSelectAll.AutoSize = $false; $btnSelectAll.Width = 150; $btnSelectAll.Height = 32
    $btnClearAll = New-Object System.Windows.Forms.Button; $btnClearAll.Text = "Odznacz wszystkie"; $btnClearAll.AutoSize = $false; $btnClearAll.Width = 150; $btnClearAll.Height = 32; $btnClearAll.Margin = '6,0,0,0'
    $chkWhatIf = New-Object System.Windows.Forms.CheckBox; $chkWhatIf.Text = "WhatIf (bez zmian)"; $chkWhatIf.AutoSize = $true; $chkWhatIf.Checked = $true
    $btnMove = New-Object System.Windows.Forms.Button; $btnMove.Text = "Tylko przenieś"; $btnMove.AutoSize = $false; $btnMove.Width = 180; $btnMove.Height = 32; $btnMove.Margin = '12,0,0,0'; $btnMove.BackColor = [System.Drawing.Color]::LightYellow
    $btnMoveAndClean = New-Object System.Windows.Forms.Button; $btnMoveAndClean.Text = "Przenieś i usuń z grup"; $btnMoveAndClean.AutoSize = $false; $btnMoveAndClean.Width = 200; $btnMoveAndClean.Height = 32; $btnMoveAndClean.Margin = '6,0,0,0'; $btnMoveAndClean.BackColor = [System.Drawing.Color]::LightGreen; $btnMoveAndClean.Font = New-Object System.Drawing.Font($form.Font, [System.Drawing.FontStyle]::Bold)
    $pActions.Controls.Add($btnSelectAll); $pActions.Controls.Add($btnClearAll); $pActions.Controls.Add($chkWhatIf); $pActions.Controls.Add($btnMove); $pActions.Controls.Add($btnMoveAndClean)
    $btnSelectAll.Add_Click({ foreach ($r in $grid.Rows) { try { $r.Cells[0].Value = $true } catch {} } })
    $btnClearAll.Add_Click({ foreach ($r in $grid.Rows) { try { $r.Cells[0].Value = $false } catch {} } })

    # ---- LOG (wiersz 5) ----
    $tbLog = New-Object System.Windows.Forms.TextBox; $tbLog.Multiline = $true; $tbLog.ScrollBars = 'Both'; $tbLog.ReadOnly = $true; $tbLog.Dock = 'Fill'; $tbLog.Font = New-Object System.Drawing.Font("Consolas",10)

    $root.Controls.Add($pBaseWrap,0,0); $root.Controls.Add($btnScan,0,1); $root.Controls.Add($grid,0,2); $root.Controls.Add($pTarget,0,3); $root.Controls.Add($pActions,0,4); $root.Controls.Add($tbLog,0,5)

    # ===================== ZAKŁADKA "USTAWIENIA" =====================
    $settingsRoot = New-Object System.Windows.Forms.TableLayoutPanel
    $settingsRoot.Dock = 'Fill'; $settingsRoot.ColumnCount = 2; $settingsRoot.RowCount = 3
    [void]$settingsRoot.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)) )
    [void]$settingsRoot.ColumnStyles.Add( (New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)) )
    [void]$settingsRoot.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    [void]$settingsRoot.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)) )
    [void]$settingsRoot.RowStyles.Add( (New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)) )
    $tabSettings.Controls.Add($settingsRoot)

    $lblExceptionUsers = New-Object System.Windows.Forms.Label; $lblExceptionUsers.Text = "Wyjątki użytkowników (jeden na linię):"; $lblExceptionUsers.Dock = 'Top'
    $tbExceptionUsers = New-Object System.Windows.Forms.TextBox; $tbExceptionUsers.Multiline = $true; $tbExceptionUsers.ScrollBars = 'Both'; $tbExceptionUsers.Dock = 'Fill'; $tbExceptionUsers.Font = New-Object System.Drawing.Font("Consolas",10)
    $tbExceptionUsers.Text = $script:ExceptionUsers -join [Environment]::NewLine
    
    $lblProtectedGroups = New-Object System.Windows.Forms.Label; $lblProtectedGroups.Text = "Grupy chronione (jedna na linię):"; $lblProtectedGroups.Dock = 'Top'
    $tbProtectedGroups = New-Object System.Windows.Forms.TextBox; $tbProtectedGroups.Multiline = $true; $tbProtectedGroups.ScrollBars = 'Both'; $tbProtectedGroups.Dock = 'Fill'; $tbProtectedGroups.Font = New-Object System.Drawing.Font("Consolas",10)
    $tbProtectedGroups.Text = $script:ProtectedGroups -join [Environment]::NewLine

    $btnSaveSettings = New-Object System.Windows.Forms.Button; $btnSaveSettings.Text = "Zapisz ustawienia"; $btnSaveSettings.Width = 150; $btnSaveSettings.Height = 32; $btnSaveSettings.Anchor = 'Right'
    $pSave = New-Object System.Windows.Forms.FlowLayoutPanel; $pSave.Dock='Fill'; $pSave.FlowDirection='RightToLeft'; $pSave.Controls.Add($btnSaveSettings)

    $settingsRoot.Controls.Add($lblExceptionUsers, 0, 0); $settingsRoot.Controls.Add($lblProtectedGroups, 1, 0)
    $settingsRoot.Controls.Add($tbExceptionUsers, 0, 1);  $settingsRoot.Controls.Add($tbProtectedGroups, 1, 1)
    $settingsRoot.SetColumnSpan($pSave, 2); $settingsRoot.Controls.Add($pSave, 0, 2)

    $btnSaveSettings.Add_Click({
        $script:ExceptionUsers = $tbExceptionUsers.Text.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
        $script:ProtectedGroups = $tbProtectedGroups.Text.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
        if (Save-Config -BaseOU $tbBase.Text -TargetOU $tbTarget.Text) {
            [System.Windows.Forms.MessageBox]::Show("Ustawienia zapisane pomyślnie.", "Sukces", "OK", "Information") | Out-Null
        } else {
            [System.Windows.Forms.MessageBox]::Show("Nie udało się zapisać ustawień.", "Błąd", "OK", "Error") | Out-Null
        }
    })

    # ===== Logika =====
    function Get-DisabledUsersOutsideOU([string]$BaseOU){
        Get-ADUser -Filter 'Enabled -eq $false' -Properties DistinguishedName,SamAccountName,Name |
            Where-Object { $_.DistinguishedName -notlike "*$BaseOU*" } |
            Select-Object Name,SamAccountName,DistinguishedName |
            Sort-Object Name
    }

    if ($cfg -and $cfg.BaseOU) {
        $tbBase.Text = [string]$cfg.BaseOU
        $lblStatus.Text = "Bazowe OU: $($tbBase.Text)"
        $btnScan.Enabled = $true
        $btnPickTarget.Enabled = $true
        if ($cfg.TargetOU) { $tbTarget.Text = [string]$cfg.TargetOU }
        $tbLog.AppendText("Wczytano konfigurację.`r`n")
    } else {
        try {
            $maybe = Get-ADOrganizationalUnit -LDAPFilter '(ou=NIEAKTYWNI)' -SearchScope Subtree -ErrorAction Stop | Select-Object -First 1 -ExpandProperty DistinguishedName
            if ($maybe) {
                $tbBase.Text = $maybe
                $lblStatus.Text = "Bazowe OU: $maybe"
                $btnScan.Enabled = $true
                $btnPickTarget.Enabled = $true
                $tbLog.AppendText("Autowybrano bazowe OU: $maybe`r`n")
                Save-Config -BaseOU $tbBase.Text -TargetOU $tbTarget.Text
            }
        } catch {}
    }

    $btnPickBase.Add_Click({
        $sel = Show-OUChooser -Title "Wybierz BAZOWE OU (np. NIEAKTYWNI)" -PreselectByName "NIEAKTYWNI"
        if ($sel) {
            $tbBase.Text = $sel; $lblStatus.Text = "Bazowe OU: $sel"; $tbTarget.Text = ""
            $btnScan.Enabled = $true; $btnPickTarget.Enabled = $true
            $tbLog.AppendText("Wybrano bazowe OU: $sel`r`n")
            Save-Config -BaseOU $tbBase.Text -TargetOU $tbTarget.Text
        }
    })

    $btnPickTarget.Add_Click({
        if ([string]::IsNullOrWhiteSpace($tbBase.Text)) { [System.Windows.Forms.MessageBox]::Show("Najpierw wybierz bazowe OU.","Uwaga",'OK','Information') | Out-Null; return }
        $sel = Show-OUChooser -Title "Wybierz DOCELOWE OU (dowolne)"
        if ($sel) {
            $tbTarget.Text = $sel
            $tbLog.AppendText("Wybrano docelowe OU: $sel`r`n")
            Save-Config -BaseOU $tbBase.Text -TargetOU $tbTarget.Text
        }
    })

    $btnScan.Add_Click({
        $grid.Rows.Clear()
        $tbLog.AppendText("Skanowanie…`r`n")
        $base = $tbBase.Text
        if ([string]::IsNullOrWhiteSpace($base)) { [System.Windows.Forms.MessageBox]::Show("Wybierz najpierw BAZOWE OU.","Uwaga",'OK','Information') | Out-Null; return }
        try {
            $users = Get-DisabledUsersOutsideOU -BaseOU $base
            foreach ($u in $users) {
                $idx = $grid.Rows.Add()
                $grid.Rows[$idx].Cells[0].Value = $false
                $grid.Rows[$idx].Cells[1].Value = $u.Name
                $grid.Rows[$idx].Cells[2].Value = $u.SamAccountName
                $grid.Rows[$idx].Cells[3].Value = $u.DistinguishedName
            }
            $tbLog.AppendText("Znaleziono: $($grid.Rows.Count) rekordów.`r`n")
            if ($grid.Rows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Brak zablokowanych użytkowników SPOZA bazowego OU.","Info",'OK','Information') | Out-Null }
        } catch {
            $tbLog.AppendText("Błąd skanowania: $($_.Exception.Message)`r`n")
            [System.Windows.Forms.MessageBox]::Show("Błąd skanowania: $($_.Exception.Message)","Błąd",'OK','Error') | Out-Null
        }
    })

    function Execute-Process([bool]$CleanGroups) {
        $base   = $tbBase.Text
        $target = $tbTarget.Text
        if ([string]::IsNullOrWhiteSpace($base))   { [System.Windows.Forms.MessageBox]::Show("Wybierz bazowe OU.","Uwaga",'OK','Information') | Out-Null; return }
        if ([string]::IsNullOrWhiteSpace($target)) { [System.Windows.Forms.MessageBox]::Show("Wybierz docelowe OU.","Uwaga",'OK','Information') | Out-Null; return }

        $selected = @()
        foreach ($r in $grid.Rows) {
            if ($r.Cells[0].Value -eq $true) {
                $selected += [PSCustomObject]@{ Name = $r.Cells[1].Value; Sam = $r.Cells[2].Value; DN = $r.Cells[3].Value }
            }
        }
        if (-not $selected) {
            [System.Windows.Forms.MessageBox]::Show("Nie zaznaczono żadnych użytkowników.","Uwaga",'OK','Information') | Out-Null; return
        }

        $whatIf = $chkWhatIf.Checked
        $actionMsg = if ($CleanGroups) { "przeniesienie i usunięcie z grup" } else { "przeniesienie" }
        $ok = [System.Windows.Forms.MessageBox]::Show(
            ("Potwierdź $actionMsg {0} użytkownik(ów){1}." -f $selected.Count, $(if($whatIf){" (WhatIf)"}else{""})),
            "Potwierdzenie",'OKCancel','Question'
        )
        if ($ok -ne 'OK') { $tbLog.AppendText("Przerwano przez użytkownika.`r`n"); return }

        $protSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        $script:ProtectedGroups | ForEach-Object { [void]$protSet.Add($_) }
        
        $moved = 0; $errors = 0
        
        foreach ($u in $selected) {
            $tbLog.AppendText("--- Przetwarzam: $($u.Name) ---`r`n")
            
            # --- 1. PRZENOSZENIE ---
            try {
                if ($u.DN -like "*$target*") {
                    $tbLog.AppendText("  Pominięto (już w docelowym OU).`r`n")
                } else {
                    if ($whatIf) {
                        $tbLog.AppendText("  [WhatIf] Przeniesienie -> $target`r`n")
                    } else {
                        Move-ADObject -Identity $u.DN -TargetPath $target -ErrorAction Stop
                        $tbLog.AppendText("  [OK] Przeniesiono -> $target`r`n")
                    }
                }
                $moved++
            } catch {
                $errors++
                $tbLog.AppendText("  Błąd przenoszenia: $($_.Exception.Message)`r`n")
            }

            # --- 2. CZYSZCZENIE GRUP ---
            if ($CleanGroups) {
                if ($script:ExceptionUsers -contains $u.Name -or $script:ExceptionUsers -contains $u.Sam) {
                    $tbLog.AppendText("  [Pominięto] Użytkownik na liście wyjątków grup.`r`n")
                    continue
                }

                try {
                    $adUser = Get-ADUser -Identity $u.Sam -Properties MemberOf -ErrorAction Stop
                    foreach ($gDN in $adUser.MemberOf) {
                        $g = Get-ADGroup -Identity $gDN -ErrorAction SilentlyContinue
                        $gName = if($g){ $g.Name } else { $null }

                        if ($gName -and $protSet.Contains($gName)) {
                            $tbLog.AppendText("  [Pominięto] Grupa chroniona: $gName`r`n")
                            continue
                        }

                        $displayName = if($gName) { $gName } else { $gDN }
                        if ($whatIf) {
                            $tbLog.AppendText("  [WhatIf] Usunięcie z grupy: $($displayName)`r`n")
                        } else {
                            try {
                                Remove-ADGroupMember -Identity $gDN -Members $adUser.DistinguishedName -Confirm:$false -ErrorAction Stop
                                $tbLog.AppendText("  [OK] Usunięto z grupy: $($displayName)`r`n")
                            } catch {
                                $tbLog.AppendText("  [Błąd usuwania z grupy] $($displayName): $($_.Exception.Message)`r`n")
                            }
                        }
                    }
                } catch {
                    $tbLog.AppendText("  Błąd pobierania członkostw: $($_.Exception.Message)`r`n")
                }
            }
        }
        [System.Windows.Forms.MessageBox]::Show(("Zakończono. Przetworzono: {0}, błędów głównych: {1}{2}" -f $moved,$errors, $(if($whatIf){" (WhatIf)"}else{""})),"Raport",'OK','Information') | Out-Null
    }

    $btnMove.Add_Click({ Execute-Process -CleanGroups $false })
    $btnMoveAndClean.Add_Click({ Execute-Process -CleanGroups $true })

    $form.Add_Shown({ $form.Activate() }) | Out-Null
    $form.Add_FormClosing({ Save-Config -BaseOU $tbBase.Text -TargetOU $tbTarget.Text }) | Out-Null
    [void]$form.ShowDialog()
}

Show-MoveDisabledUsersGUI