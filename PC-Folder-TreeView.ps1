Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# -----------------------------
# Pomocnicze: hidden / ascii tree
# -----------------------------
function Test-Hidden {
    param([string]$Path)
    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        return ($item.Attributes -band [IO.FileAttributes]::Hidden) -ne 0
    } catch {
        return $false
    }
}

function Get-AsciiTree {
    param(
        [Parameter(Mandatory)]
        [string]$RootPath,
        [switch]$ShowHidden
    )

    $lines = New-Object System.Collections.Generic.List[string]

    function Add-Branch {
        param([string]$Path, [string]$Prefix)

        try {
            $entries = Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop
        } catch {
            $lines.Add("$Prefix[BŁĄD: $($_.Exception.Message)]")
            return
        }

        # Filtr ukrytych
        if (-not $ShowHidden) {
            $entries = $entries | Where-Object {
                -not (($_.Attributes -band [IO.FileAttributes]::Hidden) -ne 0) -and
                -not $_.Name.StartsWith('.')
            }
        }

        # Sort: foldery -> pliki, alfabetycznie
        $entries = $entries | Sort-Object @{e={-not $_.PSIsContainer}}, @{e={$_.Name.ToLower()}}

        for ($i=0; $i -lt $entries.Count; $i++) {
            $e = $entries[$i]
            $isLast = ($i -eq $entries.Count - 1)
            $joint = $(if ($isLast) {'└── '} else {'├── '})
            $lines.Add("$Prefix$joint$($e.Name)")
            if ($e.PSIsContainer) {
                $childPrefix = $Prefix + $(if ($isLast) {'    '} else {'│   '})
                Add-Branch -Path $e.FullName -Prefix $childPrefix
            }
        }
    }

    $rootName = Split-Path -Path (Resolve-Path -LiteralPath $RootPath) -Leaf
    if (-not $rootName) { $rootName = $RootPath }
    $lines.Add($rootName)
    Add-Branch -Path $RootPath -Prefix ''
    return ($lines -join [Environment]::NewLine)
}

# -----------------------------
# GUI
# -----------------------------
$form                = New-Object System.Windows.Forms.Form
$form.Text           = "Podgląd drzewa folderu → zapis do TXT (PowerShell)"
$form.Size           = New-Object System.Drawing.Size(1000, 700)
$form.StartPosition  = "CenterScreen"
$form.Text = "Podglad drzewa folderu - zapis do TXT (PowerShell)"

$panelTop = New-Object System.Windows.Forms.Panel
$panelTop.Dock = [System.Windows.Forms.DockStyle]::Top
$panelTop.Height = 70
## Dodane później, po dodaniu TreeView i StatusStrip, aby prawidłowo zadziałało Dockowanie

$btnChoose = New-Object System.Windows.Forms.Button
$btnChoose.Text = "Wybierz folder…"
$btnChoose.Width = 130
$btnChoose.Location = New-Object System.Drawing.Point(8, 8)
$panelTop.Controls.Add($btnChoose)

$chkHidden = New-Object System.Windows.Forms.CheckBox
$chkHidden.Text = "Pokaż pliki ukryte"
$chkHidden.AutoSize = $true
$chkHidden.Location = New-Object System.Drawing.Point(150, 12)
$panelTop.Controls.Add($chkHidden)

$btnExpand = New-Object System.Windows.Forms.Button
$btnExpand.Text = "Rozwiń wszystko"
$btnExpand.Width = 140
$btnExpand.Location = New-Object System.Drawing.Point(320, 8)
$btnExpand.Enabled = $false
$panelTop.Controls.Add($btnExpand)

$btnCollapse = New-Object System.Windows.Forms.Button
$btnCollapse.Text = "Zwiń wszystko"
$btnCollapse.Width = 140
$btnCollapse.Location = New-Object System.Drawing.Point(470, 8)
$btnCollapse.Enabled = $false
$panelTop.Controls.Add($btnCollapse)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Zapisz jako TXT…"
$btnSave.Width = 140
$btnSave.Location = New-Object System.Drawing.Point(620, 8)
$btnSave.Enabled = $false
$panelTop.Controls.Add($btnSave)

$tree = New-Object System.Windows.Forms.TreeView
$tree.Dock = [System.Windows.Forms.DockStyle]::Fill
$tree.HideSelection = $false
$tree.PathSeparator = [IO.Path]::DirectorySeparatorChar.ToString()
$tree.ShowNodeToolTips = $true
$tree.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($tree)

$status = New-Object System.Windows.Forms.StatusStrip
$status.Dock = [System.Windows.Forms.DockStyle]::Bottom
$lblStatus = New-Object System.Windows.Forms.ToolStripStatusLabel
$status.Items.Add($lblStatus) | Out-Null
$form.Controls.Add($status)

# Dodaj panelTop na końcu, aby Dock Top nie zachodził na TreeView
$form.Controls.Add($panelTop)

# -----------------------------
# Logika drzewa
# -----------------------------
$global:CurrentRootPath = $null

function Add-PlaceholderNode {
    param([System.Windows.Forms.TreeNode]$Node)
    $ph = New-Object System.Windows.Forms.TreeNode("(wczytywanie…)")
    $ph.Tag = "__placeholder__"
    [void]$Node.Nodes.Add($ph)
}

function Populate-Children {
    param([System.Windows.Forms.TreeNode]$ParentNode)

    $path = $ParentNode.Tag
    if (-not (Test-Path -LiteralPath $path)) { return }

    try {
        $entries = Get-ChildItem -LiteralPath $path -Force -ErrorAction Stop
    } catch {
        $ParentNode.Nodes.Clear()
        [void]$ParentNode.Nodes.Add("(BŁĄD: $($_.Exception.Message))")
        return
    }

    # Filtr ukrytych
    if (-not $chkHidden.Checked) {
        $entries = $entries | Where-Object {
            -not (($_.Attributes -band [IO.FileAttributes]::Hidden) -ne 0) -and
            -not $_.Name.StartsWith('.')
        }
    }

    # Sort: foldery -> pliki
    $entries = $entries | Sort-Object @{e={-not $_.PSIsContainer}}, @{e={$_.Name.ToLower()}}

    $ParentNode.Nodes.Clear()
    foreach ($e in $entries) {
        $child = New-Object System.Windows.Forms.TreeNode($e.Name)
        $child.Tag = $e.FullName
        $child.ToolTipText = $e.FullName
        [void]$ParentNode.Nodes.Add($child)
        if ($e.PSIsContainer) {
            Add-PlaceholderNode -Node $child
        }
    }
}

function Load-Root {
    $tree.BeginUpdate()
    try {
        $tree.Nodes.Clear()
        if (-not $global:CurrentRootPath) { return }

        $rootName = Split-Path -Path (Resolve-Path -LiteralPath $global:CurrentRootPath) -Leaf
        if (-not $rootName) { $rootName = $global:CurrentRootPath }

        $root = New-Object System.Windows.Forms.TreeNode($rootName)
        $root.Tag = $global:CurrentRootPath
        $root.ToolTipText = $global:CurrentRootPath
        [void]$tree.Nodes.Add($root)

        Add-PlaceholderNode -Node $root
        Populate-Children -ParentNode $root

        # Upewnij się, że pierwszy poziom jest widoczny zaraz po wczytaniu
        $root.Expand() | Out-Null
        $tree.SelectedNode = $root
        $root.EnsureVisible()

        $btnExpand.Enabled   = $true
        $btnCollapse.Enabled = $true
        $btnSave.Enabled     = $true
    } finally {
        $tree.EndUpdate()
    }
}

function Expand-All {
    param([System.Windows.Forms.TreeNode]$Node)
    foreach ($n in $Node.Nodes) {
        # Leniwe dociąganie
        if ($n.Nodes.Count -gt 0 -and $n.Nodes[0].Tag -eq "__placeholder__") {
            Populate-Children -ParentNode $n
        }
        Expand-All -Node $n
        $n.Expand()
    }
}

function Collapse-All {
    param([System.Windows.Forms.TreeNode]$Node)
    foreach ($n in $Node.Nodes) {
        Collapse-All -Node $n
        $n.Collapse()
    }
}

# -----------------------------
# Zdarzenia
# -----------------------------
$tree.add_BeforeExpand({
    $node = $_.Node
    if ($node.Nodes.Count -gt 0 -and $node.Nodes[0].Tag -eq "__placeholder__") {
        Populate-Children -ParentNode $node
    }
})

$btnChoose.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Wybierz folder startowy"
    $dlg.ShowNewFolderButton = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:CurrentRootPath = $dlg.SelectedPath
        $lblStatus.Text = "Załadowano: $($global:CurrentRootPath)"
        Load-Root
    }
})

$chkHidden.Add_CheckedChanged({
    if ($global:CurrentRootPath) {
        Load-Root
    }
})

$btnExpand.Add_Click({
    if ($tree.Nodes.Count -gt 0) {
        $form.Cursor = 'WaitCursor'
        try {
            Expand-All -Node $tree.Nodes[0]
            $tree.Nodes[0].Expand()
        } finally {
            $form.Cursor = 'Default'
        }
    }
})

$btnCollapse.Add_Click({
    if ($tree.Nodes.Count -gt 0) {
        $form.Cursor = 'WaitCursor'
        try {
            Collapse-All -Node $tree.Nodes[0]
            $tree.Nodes[0].Collapse()
            $tree.Nodes[0].Expand()  # zostaw pierwszy poziom otwarty
        } finally {
            $form.Cursor = 'Default'
        }
    }
})

$btnSave.Add_Click({
    if (-not $global:CurrentRootPath) { return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter = "Plik tekstowy|*.txt"
    $sfd.OverwritePrompt = $true
    $sfd.FileName = (Split-Path -Leaf $global:CurrentRootPath) + "_tree.txt"

    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $form.Cursor = 'WaitCursor'
            $txt = Get-AsciiTree -RootPath $global:CurrentRootPath -ShowHidden:$chkHidden.Checked
            [IO.File]::WriteAllText($sfd.FileName, $txt, [Text.Encoding]::UTF8)
            $lblStatus.Text = "Zapisano: $($sfd.FileName)"
            [System.Windows.Forms.MessageBox]::Show("Zapisano strukturę do:`n$($sfd.FileName)", "Sukces",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Błąd zapisu: $($_.Exception.Message)", "Błąd",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) | Out-Null
        } finally {
            $form.Cursor = 'Default'
        }
    }
})

# -----------------------------
# Start GUI
# -----------------------------
[void]$form.ShowDialog()
