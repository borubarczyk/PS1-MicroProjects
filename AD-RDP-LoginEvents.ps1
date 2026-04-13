#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    A GUI application to generate a report of Windows logon/logoff events.
.DESCRIPTION
    This script provides a graphical user interface for analyzing security events from local or remote computers, including domain controllers. 
    It allows specifying event IDs, a time frame, and target computers. The analysis runs in the background to keep the UI responsive, 
    and the results are saved to a CSV file. Configuration can be saved and loaded.
.NOTES
    - Author: Gemini
    - Version: 3.0 (GUI Edition)
#>

# --- ASSEMBLY AND FORM SETUP ---
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = 'Windows Event Log Analyzer'
$mainForm.Size = New-Object System.Drawing.Size(700, 600)
$mainForm.MinimumSize = New-Object System.Drawing.Size(600, 500)
$mainForm.StartPosition = 'CenterScreen'
$mainForm.FormBorderStyle = 'Sizable'
$mainForm.AutoScaleMode = 'Dpi'

# --- HELPER FUNCTION TO ADD CONTROLS ---
function Add-Control {
    param($type, $properties, $parent = $mainForm)
    $control = New-Object $type
    $properties.GetEnumerator() | ForEach-Object { $control.$($_.Name) = $_.Value }
    $parent.Controls.Add($control)
    return $control
}

# --- GLOBAL VARIABLES FOR STATE ---
$script:job = $null
$script:logTimer = New-Object System.Windows.Forms.Timer

# --- GUI CONTROLS ---

# GroupBox for Configuration
$configGroupBox = Add-Control 'System.Windows.Forms.GroupBox' @{
    Text   = '1. Analysis Configuration'
    Location = New-Object System.Drawing.Point(10, 10)
    Size     = New-Object System.Drawing.Size(660, 120)
    Anchor   = 'Top, Left, Right'
}

# DaysBack
Add-Control 'System.Windows.Forms.Label' @{ Text = 'Days Back:'; Location = New-Object System.Drawing.Point(15, 30) } $configGroupBox
$daysBackTextBox = Add-Control 'System.Windows.Forms.TextBox' @{
    Text     = '7'
    Location = New-Object System.Drawing.Point(120, 27)
    Size     = New-Object System.Drawing.Size(50, 20)
} $configGroupBox

# Event IDs
Add-Control 'System.Windows.Forms.Label' @{ Text = 'Event IDs (CSV):'; Location = New-Object System.Drawing.Point(15, 60) } $configGroupBox
$eventIdsTextBox = Add-Control 'System.Windows.Forms.TextBox' @{
    Text     = '4624, 4634'
    Location = New-Object System.Drawing.Point(120, 57)
    Size     = New-Object System.Drawing.Size(150, 20)
} $configGroupBox

# Output Directory
Add-Control 'System.Windows.Forms.Label' @{ Text = 'Output Folder:'; Location = New-Object System.Drawing.Point(15, 90) } $configGroupBox
$outputDirTextBox = Add-Control 'System.Windows.Forms.TextBox' @{
    Text     = 'C:\Logs\Events'
    Location = New-Object System.Drawing.Point(120, 87)
    Size     = New-Object System.Drawing.Size(430, 20)
    Anchor   = 'Top, Left, Right'
} $configGroupBox
$browseButton = Add-Control 'System.Windows.Forms.Button' @{
    Text     = '...'
    Location = New-Object System.Drawing.Point(560, 86)
    Size     = New-Object System.Drawing.Size(30, 23)
    Anchor   = 'Top, Right'
} $configGroupBox

# GroupBox for Target Computers
$targetGroupBox = Add-Control 'System.Windows.Forms.GroupBox' @{
    Text   = '2. Target Computers'
    Location = New-Object System.Drawing.Point(10, 140)
    Size     = New-Object System.Drawing.Size(660, 90)
    Anchor   = 'Top, Left, Right'
}

# Query All DCs CheckBox
$queryDCsCheckBox = Add-Control 'System.Windows.Forms.CheckBox' @{
    Text     = 'Query all Domain Controllers (requires AD Module)'
    Location = New-Object System.Drawing.Point(15, 30)
    Size     = New-Object System.Drawing.Size(400, 20)
    Checked  = $false # Default to off
} $targetGroupBox

# Computer Names TextBox
Add-Control 'System.Windows.Forms.Label' @{ Text = 'Or specify (CSV):'; Location = New-Object System.Drawing.Point(15, 60) } $targetGroupBox
$computersTextBox = Add-Control 'System.Windows.Forms.TextBox' @{
    Text     = $env:COMPUTERNAME
    Location = New-Object System.Drawing.Point(120, 57)
    Size     = New-Object System.Drawing.Size(510, 20)
    Anchor   = 'Top, Left, Right'
} $targetGroupBox

# GroupBox for Actions
$actionGroupBox = Add-Control 'System.Windows.Forms.GroupBox' @{
    Text   = '3. Actions'
    Location = New-Object System.Drawing.Point(10, 240)
    Size     = New-Object System.Drawing.Size(660, 60)
    Anchor   = 'Top, Left, Right'
}

# Action Buttons
$startButton = Add-Control 'System.Windows.Forms.Button' @{ Text = 'Start Analysis'; Location = New-Object System.Drawing.Point(15, 20); Size = New-Object System.Drawing.Size(120, 30) } $actionGroupBox
$saveButton = Add-Control 'System.Windows.Forms.Button' @{ Text = 'Save Config'; Location = New-Object System.Drawing.Point(150, 20); Size = New-Object System.Drawing.Size(100, 30) } $actionGroupBox
$loadButton = Add-Control 'System.Windows.Forms.Button' @{ Text = 'Load Config'; Location = New-Object System.Drawing.Point(260, 20); Size = New-Object System.Drawing.Size(100, 30) } $actionGroupBox


# Log Output TextBox
$logTextBox = Add-Control 'System.Windows.Forms.TextBox' @{
    Text          = "Welcome! Configure the analysis and click 'Start'."
    Location      = New-Object System.Drawing.Point(10, 310)
    Size          = New-Object System.Drawing.Size(660, 240)
    Multiline     = $true
    ScrollBars    = 'Vertical'
    ReadOnly      = $true
    Font          = New-Object System.Drawing.Font('Consolas', 9)
    Anchor        = 'Top, Bottom, Left, Right'
}

# --- EVENT HANDLERS ---

$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select an output folder"
    if ($folderBrowser.ShowDialog() -eq 'OK') {
        $outputDirTextBox.Text = $folderBrowser.SelectedPath
    }
})

$queryDCsCheckBox.Add_CheckedChanged({
    $computersTextBox.Enabled = -not $queryDCsCheckBox.Checked
})

$saveButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = 'JSON files (*.json)|*.json'
    $saveFileDialog.Title = 'Save Configuration'
    if ($saveFileDialog.ShowDialog() -eq 'OK') {
        $config = @{
            DaysBack      = $daysBackTextBox.Text
            EventIDs      = $eventIdsTextBox.Text
            OutputDir     = $outputDirTextBox.Text
            QueryAllDCs   = $queryDCsCheckBox.Checked
            Computers     = $computersTextBox.Text
        }
        $config | ConvertTo-Json | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
        $logTextBox.AppendText("`r`n[INFO] Configuration saved to $($saveFileDialog.FileName)")
    }
})

$loadButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = 'JSON files (*.json)|*.json'
    $openFileDialog.Title = 'Load Configuration'
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        try {
            $config = Get-Content -Path $openFileDialog.FileName | ConvertFrom-Json
            $daysBackTextBox.Text = $config.DaysBack
            $eventIdsTextBox.Text = $config.EventIDs
            $outputDirTextBox.Text = $config.OutputDir
            $queryDCsCheckBox.Checked = $config.QueryAllDCs
            $computersTextBox.Text = $config.Computers
            $logTextBox.AppendText("`r`n[INFO] Configuration loaded from $($openFileDialog.FileName)")
        } catch {
            $logTextBox.AppendText("`r`n[ERROR] Failed to load or parse configuration file. Error: $($_.Exception.Message)")
        }
    }
})

$logTimer.Add_Tick({
    if ($script:job) {
        while ($script:job.HasMoreData) {
            $message = $script:job | Receive-Job
            $logTextBox.AppendText("`r`n$message")
        }

        if ($script:job.State -in ('Completed', 'Failed', 'Stopped')) {
            $logTextBox.AppendText("`r`n`r`n[INFO] Analysis finished with state: $($script:job.State).")
            # Clean up the job
            Remove-Job $script:job
            $script:job = $null
            # Stop the timer and re-enable controls
            $script:logTimer.Stop()
            $actionGroupBox.Enabled = $true
            $configGroupBox.Enabled = $true
            $targetGroupBox.Enabled = $true
        }
    }
})

$startButton.Add_Click({
    $logTextBox.Clear()
    $logTextBox.AppendText("[INFO] Starting analysis... The UI will remain responsive.")

    # Disable controls during analysis
    $actionGroupBox.Enabled = $false
    $configGroupBox.Enabled = $false
    $targetGroupBox.Enabled = $false

    # Gather parameters from form
    $params = @{
        DaysBack  = [int]$daysBackTextBox.Text
        EventIDs  = $eventIdsTextBox.Text -split ',' | ForEach-Object { $_.Trim() }
        OutputDir = $outputDirTextBox.Text
        QueryDCs  = $queryDCsCheckBox.Checked
        Computers = $computersTextBox.Text -split ',' | ForEach-Object { $_.Trim() }
    }

    # This scriptblock runs in the background
    $scriptBlock = {
        param($params)

        # --- CORE ANALYSIS LOGIC ---
        $ErrorActionPreference = 'Stop'
        
        # 1. Determine target computers
        $targetComputers = @()
        if ($params.QueryDCs) {
            Write-Output "[INFO] Discovering Domain Controllers..."
            try {
                $targetComputers = (Get-ADDomainController -Filter * | Select-Object -ExpandProperty HostName | Sort-Object)
                Write-Output "[INFO] Found $($targetComputers.Count) Domain Controllers."
            } catch {
                Write-Output "[ERROR] Failed to get Domain Controllers. The ActiveDirectory module might be missing or there was a network issue."
                return
            }
        } else {
            $targetComputers = $params.Computers
            Write-Output "[INFO] Targeting specified computers: $($targetComputers -join ', ')"
        }

        # 2. Create output directory
        try {
            New-Item -ItemType Directory -Path $params.OutputDir -Force | Out-Null
        } catch {
            Write-Output "[ERROR] Failed to create output directory '$($params.OutputDir)'. Please check permissions."
            return
        }

        # 3. Collect Events
        $startTime = (Get-Date).Date.AddDays(-[math]::Abs($params.DaysBack))
        $allEvents = [System.Collections.Generic.List[object]]::new()
        
        $total = $targetComputers.Count
        $current = 0
        foreach ($computer in $targetComputers) {
            $current++
            Write-Output "[INFO] Querying $($computer) ($current of $total)..."
            $filter = @{
                LogName   = 'Security'
                Id        = $params.EventIDs
                StartTime = $startTime
            }
            try {
                $events = Get-WinEvent -FilterHashtable $filter -ComputerName $computer -ErrorAction Stop
                $allEvents.AddRange($events)
                Write-Output "[INFO] Found $($events.Count) events on $computer."
            } catch {
                if ($_.FullyQualifiedErrorId -ne 'NoEventsFound,Microsoft.PowerShell.Commands.GetWinEventCommand') {
                    Write-Output "[WARN] Failed to get events from '$computer'. It might be offline or access denied. Error: $($_.Exception.Message)"
                }
            }
        }

        if ($allEvents.Count -eq 0) {
            Write-Output "[INFO] No matching events found in the specified timeframe."
            return
        }

        # 4. Process events into a readable format
        Write-Output "[INFO] Processing $($allEvents.Count) total events..."
        $results = foreach ($event in $allEvents) {
            $properties = @{}
            for ($i = 0; $i -lt $event.Properties.Count; $i++) {
                # Attempt to get a meaningful name, otherwise use index
                $propName = try { ($event.ToXml() | Select-Xml -XPath "/*/*[local-name()='EventData']/*[@Name][position()=$($i+1)]").Node.Name } catch { "Property_$i" }
                $properties[$propName] = $event.Properties[$i].Value
            }
            
            [pscustomobject]@{
                TimeCreated = $event.TimeCreated
                EventID     = $event.Id
                Computer    = $event.MachineName
                Message     = $event.Message -replace "`r|`n"," "
                Properties  = $properties | ConvertTo-Json -Compress
            }
        }
        
        # 5. Save report
        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $csvPath = Join-Path $params.OutputDir "Event_Report_$stamp.csv"
        try {
            $results | Sort-Object TimeCreated -Descending | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Output "[SUCCESS] Report with $($results.Count) events saved to: $csvPath"
        }
        catch {
            Write-Output "[ERROR] Failed to save CSV file. Error: $($_.Exception.Message)"
        }
    }

    # Start the job and the timer to monitor it
    $script:job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $params
    $script:logTimer.Interval = 500 # 0.5 seconds
    $script:logTimer.Start()
})


# --- SHOW THE FORM ---
$mainForm.ShowDialog()

# --- CLEANUP ---
$mainForm.Dispose()
if ($script:job) {
    Stop-Job $script:job
    Remove-Job $script:job
}
$logTimer.Dispose()
