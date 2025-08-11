#Requires -RunAsAdministrator # Recommended for Restart-Service functionality

#region Assembly Loading and Initial Setup
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define paths and settings file
$scriptName = "BigFixLogViewer"
$appDataPath = "$env:APPDATA\$scriptName"
if (-not (Test-Path $appDataPath)) {
    New-Item -Path $appDataPath -ItemType Directory | Out-Null
}
$settingsFile = "$appDataPath\settings.json"

# Default global keywords (will be overridden by settings file if it exists)
$globalKeywords = @{
    Success = @{Patterns = @('SUCCESS', 'SUCCESSFULLY'); Color = 'LightGreen'}
    Warning = @{Patterns = @('WARNING'); Color = 'LightYellow'}
    Fail    = @{Patterns = @('FAIL', 'FAILURE', 'FAILED'); Color = 'LightCoral'}
}

# Function to save settings to JSON
function Save-Settings {
    $globalKeywords | ConvertTo-Json -Depth 3 | Out-File -FilePath $settingsFile -Encoding UTF8
}

# Function to load settings from JSON
function Load-Settings {
    if (Test-Path $settingsFile) {
        try {
            $loadedSettings = Get-Content -Path $settingsFile -Raw | ConvertFrom-Json
            # Basic validation
            if ($loadedSettings.Success -and $loadedSettings.Warning -and $loadedSettings.Fail) {
                $global:globalKeywords = $loadedSettings
            }
        }
        catch {
            Write-Warning "Could not load settings file. Using defaults. Error: $_"
        }
    }
}

# Load settings on startup
Load-Settings

# Get machine info
$machineName = [System.Net.Dns]::GetHostName()
$ipAddress = @([System.Net.Dns]::GetHostEntry($machineName).AddressList | Where-Object { $_.AddressFamily -eq 'InterNetwork' }).IPAddressToString[0]

# Check for Admin rights
$isAdmin = ([System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)

#endregion

#region GUI Creation
# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Log Viewer"
$form.Size = New-Object System.Drawing.Size(850, 650)
$form.StartPosition = "CenterScreen"

# Create TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$form.Controls.Add($tabControl)

# --- Menu Strip ---
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem("File")
$menuOpen = New-Object System.Windows.Forms.ToolStripMenuItem("Open Log...")
$menuExport = New-Object System.Windows.Forms.ToolStripMenuItem("Export View...")
$menuPause = New-Object System.Windows.Forms.ToolStripMenuItem("Pause Tailing")
$menuPause.CheckOnClick = $true
$menuExit = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
$menuFile.DropDownItems.AddRange(@($menuOpen, $menuExport, $menuPause, $menuExit))

$menuTools = New-Object System.Windows.Forms.ToolStripMenuItem("Tools")
$menuRestartBES = New-Object System.Windows.Forms.ToolStripMenuItem("Restart BESClient")
$menuRestartBES.Enabled = $isAdmin # Disable if not running as admin
$menuTools.DropDownItems.Add($menuRestartBES)

$menuView = New-Object System.Windows.Forms.ToolStripMenuItem("View")
$menuWrap = New-Object System.Windows.Forms.ToolStripMenuItem("Word Wrap")
$menuWrap.CheckOnClick = $true
$menuWrap.Checked = $true
$menuZoomIn = New-Object System.Windows.Forms.ToolStripMenuItem("Zoom In (+)")
$menuZoomOut = New-Object System.Windows.Forms.ToolStripMenuItem("Zoom Out (-)")
$menuStats = New-Object System.Windows.Forms.ToolStripMenuItem("Show Stats")
$menuView.DropDownItems.AddRange(@($menuWrap, $menuZoomIn, $menuZoomOut, $menuStats))

$menuFilter = New-Object System.Windows.Forms.ToolStripMenuItem("Filter")
$menuShowAll = New-Object System.Windows.Forms.ToolStripMenuItem("Show All")
$menuShowErrors = New-Object System.Windows.Forms.ToolStripMenuItem("Show Errors Only")
$menuShowWarnings = New-Object System.Windows.Forms.ToolStripMenuItem("Show Warnings Only")
$menuShowSuccess = New-Object System.Windows.Forms.ToolStripMenuItem("Show Success Only")
$menuCustomizeKeywords = New-Object System.Windows.Forms.ToolStripMenuItem("Customize Keywords...")
$menuCustomizeColors = New-Object System.Windows.Forms.ToolStripMenuItem("Customize Colors...")
$menuFilter.DropDownItems.AddRange(@($menuShowAll, $menuShowErrors, $menuShowWarnings, $menuShowSuccess, $menuCustomizeKeywords, $menuCustomizeColors))

$menuStrip.Items.AddRange(@($menuFile, $menuTools, $menuView, $menuFilter))
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

# --- Search Panel ---
$searchPanel = New-Object System.Windows.Forms.Panel
$searchPanel.Dock = "Top"
$searchPanel.Height = 30
$searchPanel.Visible = $false # Hidden by default
$searchBox = New-Object System.Windows.Forms.TextBox -Property @{ Location = New-Object System.Drawing.Point(5, 5); Width = 200 }
$btnSearchNext = New-Object System.Windows.Forms.Button -Property @{ Text = "Find Next"; Location = New-Object System.Drawing.Point(210, 3) }
$btnSearchPrev = New-Object System.Windows.Forms.Button -Property @{ Text = "Find Prev"; Location = New-Object System.Drawing.Point(290, 3) }
$btnHighlightAll = New-Object System.Windows.Forms.Button -Property @{ Text = "Highlight All"; Location = New-Object System.Drawing.Point(370, 3) }
$btnCloseSearch = New-Object System.Windows.Forms.Button -Property @{ Text = "X"; Location = New-Object System.Drawing.Point(450, 3); Width = 30 }
$searchPanel.Controls.AddRange(@($searchBox, $btnSearchNext, $btnSearchPrev, $btnHighlightAll, $btnCloseSearch))
$form.Controls.Add($searchPanel)
$form.KeyPreview = $true

# --- Status Strip ---
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusProgress = New-Object System.Windows.Forms.ToolStripProgressBar
$statusProgress.Visible = $false
$statusStrip.Items.AddRange(@($statusLabel, $statusProgress))
$form.Controls.Add($statusStrip)

# Tab context menu for closing
$tabContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$closeTabMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Close Tab")
$tabContextMenu.Items.Add($closeTabMenu)
$tabControl.ContextMenuStrip = $tabContextMenu

# Timers
$tailingTimer = New-Object System.Windows.Forms.Timer -Property @{ Interval = 1000 }
$jobCheckTimer = New-Object System.Windows.Forms.Timer -Property @{ Interval = 500 }

#endregion

#region Core Functions

# Function to add a new tab for a log file
function Add-LogTab {
    param ([string]$filePath)

    if (-not (Test-Path $filePath)) {
        [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Error", "OK", "Error")
        return
    }

    $tabPage = New-Object System.Windows.Forms.TabPage
    $tabPage.Text = (Get-Item $filePath).Name
    $tabPage.ToolTipText = $filePath

    $richTextBox = New-Object System.Windows.Forms.RichTextBox
    $richTextBox.Dock = "Fill"
    $richTextBox.ReadOnly = $true
    $richTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $richTextBox.WordWrap = $menuWrap.Checked
    $richTextBox.ScrollBars = "Both"
    $richTextBox.HideSelection = $false
    
    $tabPage.Controls.Add($richTextBox)

    try {
        $initialContent = Get-Content -Path $filePath -Raw -Encoding UTF8
        Add-ColoredText -rtb $richTextBox -text $initialContent

        $lastLength = (Get-Item $filePath).Length
        # Use StringBuilder for efficient string manipulation
        $stringBuilder = New-Object System.Text.StringBuilder
        $stringBuilder.Append($initialContent) | Out-Null
        
        $tabPage.Tag = @{
            FilePath    = $filePath
            LastLength  = $lastLength
            Filter      = "All"
            FullContent = $stringBuilder
        }

        $tabControl.TabPages.Add($tabPage)
        $tabControl.SelectedTab = $tabPage
        Update-Status -tab $tabPage
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading file '$filePath':`n$($_)", "File Load Error", "OK", "Error")
        return
    }
}

# Function to append text with row highlighting and optional filter
function Add-ColoredText {
    param (
        $rtb,
        [string]$text,
        [string]$filter = "All"
    )
    $rtb.BeginUpdate()
    $lines = $text -split '\r?\n'
    foreach ($line in $lines) {
        $matchType = Get-MatchType -line $line
        if ($filter -ne "All" -and $matchType -ne $filter) { continue }

        $startPos = $rtb.TextLength
        $rtb.AppendText($line + "`n")
        $rtb.Select($startPos, $line.Length)
        
        $backColor = [System.Drawing.Color]::White
        if ($matchType) {
            $backColor = [System.Drawing.Color]::$($globalKeywords[$matchType].Color)
        }
        $rtb.SelectionBackColor = $backColor
    }
    $rtb.SelectionStart = $rtb.TextLength
    $rtb.ScrollToCaret()
    $rtb.EndUpdate()
}

# Helper to get match type
function Get-MatchType {
    param ([string]$line)
    foreach ($key in $globalKeywords.Fail.Patterns) { if ($line -imatch "\b$key\b") { return "Fail" } }
    foreach ($key in $globalKeywords.Warning.Patterns) { if ($line -imatch "\b$key\b") { return "Warning" } }
    foreach ($key in $globalKeywords.Success.Patterns) { if ($line -imatch "\b$key\b") { return "Success" } }
    return $null
}

# Function to refresh all tabs with new colors or content
function Refresh-Tabs {
    foreach ($tab in $tabControl.TabPages) {
        $rtb = $tab.Controls[0]
        $ht = $tab.Tag
        $rtb.Clear()
        Add-ColoredText -rtb $rtb -text $ht.FullContent.ToString() -filter $ht.Filter
        Update-Status -tab $tab
    }
}

# Function to update the status bar for the current tab
function Update-Status {
    param($tab)
    if (-not $tab) { $statusLabel.Text = "Ready"; return }
    $ht = $tab.Tag
    $fileSizeKB = [math]::Round($ht.LastLength / 1024, 2)
    $statusLabel.Text = "Machine: $machineName | IP: $ipAddress | Path: $($ht.FilePath) | Size: $fileSizeKB KB | Filter: $($ht.Filter)"
}

# Apply filter to current tab
function Apply-Filter {
    param ([string]$filterType)
    $tab = $tabControl.SelectedTab
    if (-not $tab) { return }
    $rtb = $tab.Controls[0]
    $ht = $tab.Tag
    $ht.Filter = $filterType
    $rtb.Clear()
    Add-ColoredText -rtb $rtb -text $ht.FullContent.ToString() -filter $filterType
    Update-Status -tab $tab
}

#endregion

#region Event Handlers

# --- File Menu Events ---
$menuOpen.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        Add-LogTab -filePath $openFileDialog.FileName
    }
})

$menuExport.Add_Click({
    $tab = $tabControl.SelectedTab
    if (-not $tab) { return }
    $rtb = $tab.Controls[0]
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
    $saveFileDialog.FileName = "Exported_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        try {
            $rtb.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Exported to $($saveFileDialog.FileName)", "Success", "OK", "Information")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting file: $($_)", "Error", "OK", "Error")
        }
    }
})

$menuExit.Add_Click({ $form.Close() })

# --- Tools Menu Events ---
$menuRestartBES.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("Are you sure you want to restart the BESClient service?", "Confirm Restart", "YesNo", "Warning") -ne "Yes") {
        return
    }

    $menuRestartBES.Enabled = $false
    $statusProgress.Visible = $true
    $statusProgress.Value = 10
    $statusLabel.Text = "Restarting BESClient service..."

    $scriptBlock = {
        try {
            Write-Progress -Activity "Restarting BESClient" -Status "Stopping service..." -PercentComplete 30
            Stop-Service -Name "BESClient" -Force -ErrorAction Stop
            
            Start-Sleep -Seconds 2 # Give it a moment

            Write-Progress -Activity "Restarting BESClient" -Status "Starting service..." -PercentComplete 70
            Start-Service -Name "BESClient" -ErrorAction Stop

            return "Success: BESClient restarted."
        }
        catch {
            return "Error: $($_.Exception.Message)"
        }
    }

    $job = Start-Job -ScriptBlock $scriptBlock
    $jobCheckTimer.Tag = $job # Store job in timer's tag
    $jobCheckTimer.Start()
})

$jobCheckTimer.Add_Tick({
    $job = $jobCheckTimer.Tag
    if ($job.State -in @("Completed", "Failed", "Stopped")) {
        $jobCheckTimer.Stop()
        $result = Receive-Job -Job $job
        Remove-Job -Job $job

        if ($result -like "Success*") {
            $statusLabel.Text = "BESClient restarted successfully!"
            $statusProgress.Value = 100
        } else {
            $statusLabel.Text = "Error restarting BESClient. Check logs."
            [System.Windows.Forms.MessageBox]::Show("Failed to restart BESClient.`n$result", "Service Error", "OK", "Error")
            $statusProgress.Value = 0
        }
        
        Start-Sleep -Seconds 3 # Show status for a moment
        $statusProgress.Visible = $false
        Update-Status -tab $tabControl.SelectedTab
        $menuRestartBES.Enabled = $true
    } else {
        $statusProgress.Value = ($statusProgress.Value + 5) % 80 # Simple progress animation
    }
})


# --- View Menu Events ---
$menuWrap.Add_Click({
    foreach ($tab in $tabControl.TabPages) {
        $tab.Controls[0].WordWrap = $menuWrap.Checked
    }
})

$menuZoomIn.Add_Click({
    $rtb = $tabControl.SelectedTab.Controls[0]
    if ($rtb.Font.Size -lt 30) {
        $newSize = $rtb.Font.Size + 1
        $rtb.Font = New-Object System.Drawing.Font("Consolas", $newSize)
    }
})

$menuZoomOut.Add_Click({
    $rtb = $tabControl.SelectedTab.Controls[0]
    if ($rtb.Font.Size -gt 8) {
        $newSize = $rtb.Font.Size - 1
        $rtb.Font = New-Object System.Drawing.Font("Consolas", $newSize)
    }
})

$form.Add_KeyDown({
    param($sender, $e)
    if ($e.Control -and $e.KeyCode -eq 'F') {
        $searchPanel.Visible = $true
        $searchBox.Focus()
    }
    if ($e.KeyCode -eq 'Escape') {
        $searchPanel.Visible = $false
    }
})

# --- Filter Menu Events ---
$menuShowAll.Add_Click({ Apply-Filter -filterType "All" })
$menuShowErrors.Add_Click({ Apply-Filter -filterType "Fail" })
$menuShowWarnings.Add_Click({ Apply-Filter -filterType "Warning" })
$menuShowSuccess.Add_Click({ Apply-Filter -filterType "Success" })

# --- Search Panel Events ---
$searchLogic = {
    param($direction) # 'Next' or 'Prev'
    $tab = $tabControl.SelectedTab
    if(-not $tab) { return }
    $rtb = $tab.Controls[0]
    $searchText = $searchBox.Text
    if ([string]::IsNullOrEmpty($searchText)) { return }
    
    $start = if ($direction -eq 'Next') { $rtb.SelectionStart + $rtb.SelectionLength } else { $rtb.SelectionStart }
    $options = if ($direction -eq 'Next') { [System.Windows.Forms.RichTextBoxFinds]::None } else { [System.Windows.Forms.RichTextBoxFinds]::Reverse }
    
    $pos = $rtb.Find($searchText, $start, $rtb.TextLength, $options)
    
    if ($pos -ge 0) {
        $rtb.Focus()
    } else {
        [System.Windows.Forms.MessageBox]::Show("No more matches found for '$searchText'.", "Search Finished", "OK", "Information")
    }
}
$btnSearchNext.Add_Click({ & $searchLogic 'Next' })
$btnSearchPrev.Add_Click({ & $searchLogic 'Prev' })
$searchBox.Add_KeyDown({ if($_.KeyCode -eq 'Enter'){ & $searchLogic 'Next' }})
$btnCloseSearch.Add_Click({ $searchPanel.Visible = $false })

# --- Tab Events ---
$tabControl.Add_SelectedIndexChanged({ Update-Status -tab $tabControl.SelectedTab })
$closeTabMenu.Add_Click({
    if ($tabControl.SelectedTab) {
        $tabControl.TabPages.Remove($tabControl.SelectedTab)
    }
})

#endregion

#region Tailing and Watcher

$tailingTimer.Add_Tick({
    if ($menuPause.Checked) { return }
    foreach ($tab in $tabControl.TabPages) {
        $rtb = $tab.Controls[0]
        $ht = $tab.Tag
        $fp = $ht.FilePath
        
        try {
            $currentLength = (Get-Item $fp).Length
            # Detect log rotation (file got smaller)
            if ($currentLength -lt $ht.LastLength) {
                $ht.FullContent.Clear()
                $rtb.Clear()
                $ht.LastLength = 0
                # Fall through to read the new content from the start
            }

            if ($currentLength -gt $ht.LastLength) {
                $fs = New-Object System.IO.FileStream($fp, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $fs.Seek($ht.LastLength, [System.IO.SeekOrigin]::Begin)
                $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
                
                $newPart = $sr.ReadToEnd()
                $sr.Close()
                $fs.Close()

                if (-not [string]::IsNullOrEmpty($newPart)) {
                    $ht.FullContent.Append($newPart)
                    Add-ColoredText -rtb $rtb -text $newPart -filter $ht.Filter
                    $ht.LastLength = $currentLength
                    if ($tab -eq $tabControl.SelectedTab) {
                        Update-Status -tab $tab
                    }
                }
            }
        } catch {
            # Could be a temporary lock, just try again next tick.
            Write-Warning "Error tailing file $fp: $_"
        }
    }
})

#endregion

#region Form Load and Close
# Load the most recent BigFix log by default
$form.Add_Load({
    $bigFixLogPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
    if (Test-Path $bigFixLogPath) {
        $logFile = Get-ChildItem -Path $bigFixLogPath -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($logFile) {
            Add-LogTab -filePath $logFile.FullName
        } else {
            $statusLabel.Text = "No log files found in BigFix path. Open logs manually via File -> Open."
        }
    } else {
         $statusLabel.Text = "BigFix log path not found. Open logs manually via File -> Open."
    }
    $tailingTimer.Start()
})

# Form closing cleanup
$form.Add_Closing({
    $tailingTimer.Stop()
    # Any other cleanup
})

# Show the form
$form.ShowDialog() | Out-Null

# Clean up timers on exit
$tailingTimer.Dispose()
$jobCheckTimer.Dispose()
#endregion
