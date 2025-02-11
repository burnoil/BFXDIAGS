# Set this flag to $true to enable debug messages; $false to disable.
$EnableDebug = $false

# Load required .NET assemblies.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

###############################################################
# Section 1: Setup – Configuration, Log File, System Info, and Global Variables
###############################################################

# Set configuration file path.
$global:configFile = Join-Path $PSScriptRoot "BigFixLogViewerSettings.json"

# Define default values.
$defaultLogDir = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
$defaultRefreshInterval = 1000
$defaultWindowSize = @{ Width = 1000; Height = 750 }
$defaultWindowLocation = @{ X = 100; Y = 100 }
$defaultCustomHighlightRules = @()

# Load saved settings if available.
$global:savedSettings = $null
if (Test-Path $global:configFile) {
    try {
        $global:savedSettings = Get-Content $global:configFile -Raw | ConvertFrom-Json
    }
    catch {
        if ($EnableDebug) { Write-Host "Error reading config file: $_" }
    }
}

# Always get the newest log file from the default log directory.
if (-not (Test-Path $defaultLogDir)) {
    [System.Windows.Forms.MessageBox]::Show("Log directory not found:`n$defaultLogDir", "Error", 'OK', 'Error')
    exit
}
$newestLog = Get-ChildItem -Path $defaultLogDir -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $newestLog) {
    [System.Windows.Forms.MessageBox]::Show("No log files found in:`n$defaultLogDir", "Error", 'OK', 'Error')
    exit
}

# Determine which log file to open:
# If a saved log file exists and is still present, compare its LastWriteTime with the newest log.
# If the newest log is more recent, use it.
if ($global:savedSettings -and $global:savedSettings.LastLogFile -and (Test-Path $global:savedSettings.LastLogFile)) {
    $savedLog = Get-Item $global:savedSettings.LastLogFile
    if ($newestLog.LastWriteTime -gt $savedLog.LastWriteTime) {
        if ($EnableDebug) { Write-Host "Using newer log: $($newestLog.FullName)" }
        $global:logFilePath = $newestLog.FullName
    } else {
        if ($EnableDebug) { Write-Host "Using saved log: $($global:savedSettings.LastLogFile)" }
        $global:logFilePath = $global:savedSettings.LastLogFile
    }
} else {
    if ($EnableDebug) { Write-Host "No saved log file found. Using newest log: $($newestLog.FullName)" }
    $global:logFilePath = $newestLog.FullName
}

# Load system information.
$computerName = $env:COMPUTERNAME
try {
    # Exclude loopback and link-local addresses.
    $ipAddresses = (Get-NetIPAddress -AddressFamily IPv4 |
                    Where-Object { $_.IPAddress -notmatch '^127\.0\.0\.1' -and $_.IPAddress -notmatch '^169\.254\.' } |
                    Select-Object -ExpandProperty IPAddress) -join ', '
}
catch {
    $ipAddresses = ([System.Net.Dns]::GetHostAddresses($computerName) |
                    Where-Object { $_.AddressFamily -eq 'InterNetwork' -and $_.IPAddressToString -notmatch '^127\.0\.0\.1' -and $_.IPAddressToString -notmatch '^169\.254\.' } |
                    ForEach-Object { $_.IPAddressToString }) -join ', '
}
$disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
$diskSizeGB  = [math]::Round($disk.Size / 1GB, 2)
$freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
try {
    $relayServerKeyPath = "HKLM:\SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client\__RelayServer1"
    $relayServer = (Get-ItemProperty -Path $relayServerKeyPath -ErrorAction Stop).value
}
catch {
    $relayServer = "Not Found"
}
# Correct BESClient.exe path.
$clientExePath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClient.exe"
if (Test-Path $clientExePath) {
    $clientVersion = (Get-Item $clientExePath).VersionInfo.FileVersion
}
else {
    $clientVersion = "Not Found"
}

$systemInfoText = "Machine Name: $computerName`n" +
                  "IP Addresses: $ipAddresses`n" +
                  "Disk Size (C:): $diskSizeGB GB`n" +
                  "Free Space (C:): $freeSpaceGB GB`n" +
                  "Relay Server: $relayServer"

# Global variables.
$global:lastPos = 0
$global:errorCount   = 0
$global:warningCount = 0
$global:successCount = 0
$global:IsEvtx = $false

if ($global:savedSettings -and $global:savedSettings.CustomHighlightRules) {
    $global:customHighlightRules = foreach ($rule in $global:savedSettings.CustomHighlightRules) {
        [PSCustomObject]@{
            Keyword = $rule.Keyword
            Color   = [System.Drawing.Color]::FromName($rule.ColorName)
        }
    }
} else {
    $global:customHighlightRules = $defaultCustomHighlightRules
}

###############################################################
# Section 2: Build the Main Form and TabControl
###############################################################

$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Log Viewer"
$form.Width = if ($global:savedSettings -and $global:savedSettings.WindowSize -and $global:savedSettings.WindowSize.Width) { [int]$global:savedSettings.WindowSize.Width } else { $defaultWindowSize.Width }
$form.Height = if ($global:savedSettings -and $global:savedSettings.WindowSize -and $global:savedSettings.WindowSize.Height) { [int]$global:savedSettings.WindowSize.Height } else { $defaultWindowSize.Height }
$form.Location = if ($global:savedSettings -and $global:savedSettings.WindowLocation) { New-Object System.Drawing.Point($global:savedSettings.WindowLocation.X, $global:savedSettings.WindowLocation.Y) } else { New-Object System.Drawing.Point($defaultWindowLocation.X, $defaultWindowLocation.Y) }
$form.StartPosition = 'Manual'

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$null = $form.Controls.Add($tabControl)

# ----- Tab Page 1: Log Viewer -----
$logViewerTab = New-Object System.Windows.Forms.TabPage
$logViewerTab.Text = "Log Viewer"
$null = $tabControl.TabPages.Add($logViewerTab)

$logViewerPanel = New-Object System.Windows.Forms.Panel
$logViewerPanel.Dock = 'Fill'
$null = $logViewerTab.Controls.Add($logViewerPanel)

# -- System Info Panel (within Log Viewer Tab) --
$systemInfoPanel = New-Object System.Windows.Forms.Panel
$systemInfoPanel.Dock = 'Top'
$systemInfoPanel.Height = 120
$systemInfoPanel.BackColor = [System.Drawing.Color]::Black
$systemInfoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$null = $logViewerPanel.Controls.Add($systemInfoPanel)

$flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flowPanel.Dock = 'Fill'
$flowPanel.FlowDirection = 'TopDown'
$flowPanel.WrapContents = $false
$null = $systemInfoPanel.Controls.Add($flowPanel)

$systemInfoLabel = New-Object System.Windows.Forms.Label
$systemInfoLabel.AutoSize = $true
$systemInfoLabel.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
$systemInfoLabel.ForeColor = [System.Drawing.Color]::Cyan
$systemInfoLabel.Text = $systemInfoText
$null = $flowPanel.Controls.Add($systemInfoLabel)

$besInfoLabel = New-Object System.Windows.Forms.Label
$besInfoLabel.AutoSize = $true
$besInfoLabel.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
$besInfoLabel.ForeColor = [System.Drawing.Color]::Cyan
$besInfoLabel.Text = "BESClient Version: $clientVersion | BESClient Service Status: Unknown"
$null = $flowPanel.Controls.Add($besInfoLabel)

# -- Log File Panel (within Log Viewer Tab) --
$logFilePanel = New-Object System.Windows.Forms.Panel
$logFilePanel.Dock = 'Top'
$logFilePanel.Height = 30
$null = $logViewerPanel.Controls.Add($logFilePanel)

function Update-LogFileLabel {
    $fileItem = Get-Item $global:logFilePath
    $lastWrite = $fileItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
    $logFileLabel.Text = "Log File: $global:logFilePath (Last Modified: $lastWrite)"
}
$logFileLabel = New-Object System.Windows.Forms.Label
$logFileLabel.AutoSize = $true
$logFileLabel.Font = New-Object System.Drawing.Font("Consolas", 9)
Update-LogFileLabel
$logFileLabel.Location = New-Object System.Drawing.Point(10, 5)
$null = $logFilePanel.Controls.Add($logFileLabel)

# -- Control Panel (within Log Viewer Tab) --
$controlPanel = New-Object System.Windows.Forms.Panel
$controlPanel.Dock = 'Top'
$controlPanel.Height = 100
$controlPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
$null = $logViewerPanel.Controls.Add($controlPanel)

# Row 1: Primary Buttons.
$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Search:"
$searchLabel.Location = New-Object System.Drawing.Point(10, 5)
$searchLabel.AutoSize = $true
$null = $controlPanel.Controls.Add($searchLabel)

$searchTextBox = New-Object System.Windows.Forms.TextBox
$searchTextBox.Location = New-Object System.Drawing.Point(70, 2)
$searchTextBox.Width = 150
$null = $controlPanel.Controls.Add($searchTextBox)

$findNextButton = New-Object System.Windows.Forms.Button
$findNextButton.Text = "Find Next"
$findNextButton.Location = New-Object System.Drawing.Point(230, 2)
$findNextButton.AutoSize = $true
$null = $controlPanel.Controls.Add($findNextButton)

$pauseResumeButton = New-Object System.Windows.Forms.Button
$pauseResumeButton.Text = "Pause"
$pauseResumeButton.Location = New-Object System.Drawing.Point(320, 2)
$pauseResumeButton.AutoSize = $true
$null = $controlPanel.Controls.Add($pauseResumeButton)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export Log"
$exportButton.Location = New-Object System.Drawing.Point(400, 2)
$exportButton.AutoSize = $true
$null = $controlPanel.Controls.Add($exportButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Clear Log"
$clearButton.Location = New-Object System.Drawing.Point(490, 2)
$clearButton.AutoSize = $true
$null = $controlPanel.Controls.Add($clearButton)

$chooseFileButton = New-Object System.Windows.Forms.Button
$chooseFileButton.Text = "Choose Log File"
$chooseFileButton.AutoSize = $true
$chooseFileButton.Location = New-Object System.Drawing.Point(580, 2)
$null = $controlPanel.Controls.Add($chooseFileButton)

$manageHighlightsButton = New-Object System.Windows.Forms.Button
$manageHighlightsButton.Text = "Manage Highlights"
$manageHighlightsButton.AutoSize = $true
$manageHighlightsButton.Location = New-Object System.Drawing.Point(680, 2)
$null = $controlPanel.Controls.Add($manageHighlightsButton)

$openEventViewerButton = New-Object System.Windows.Forms.Button
$openEventViewerButton.Text = "Open Event Viewer"
$openEventViewerButton.AutoSize = $true
$openEventViewerButton.Location = New-Object System.Drawing.Point(800, 2)
$null = $controlPanel.Controls.Add($openEventViewerButton)

$restartBESClientButton = New-Object System.Windows.Forms.Button
$restartBESClientButton.Text = "Restart BESClient"
$restartBESClientButton.AutoSize = $true
$restartBESClientButton.Location = New-Object System.Drawing.Point(920, 2)
$restartBESClientButton.BackColor = [System.Drawing.Color]::LightGreen
$null = $controlPanel.Controls.Add($restartBESClientButton)

# Row 2: Refresh Interval and Stats.
$refreshLabel = New-Object System.Windows.Forms.Label
$refreshLabel.Text = "Refresh Interval (ms):"
$refreshLabel.Location = New-Object System.Drawing.Point(10, 60)
$refreshLabel.AutoSize = $true
$null = $controlPanel.Controls.Add($refreshLabel)

$refreshNumeric = New-Object System.Windows.Forms.NumericUpDown
$refreshNumeric.Location = New-Object System.Drawing.Point(140, 58)
$refreshNumeric.Minimum = 100
$refreshNumeric.Maximum = 5000
$refreshNumeric.Value = if ($global:savedSettings -and $global:savedSettings.RefreshInterval) { $global:savedSettings.RefreshInterval } else { $defaultRefreshInterval }
$null = $controlPanel.Controls.Add($refreshNumeric)

$statsLabel = New-Object System.Windows.Forms.Label
$statsLabel.Text = "Errors: 0, Warnings: 0, Successes: 0"
$statsLabel.Location = New-Object System.Drawing.Point(320, 60)
$statsLabel.AutoSize = $true
$null = $controlPanel.Controls.Add($statsLabel)

# -- Log Content Panel (within Log Viewer Tab) --
$richTextBox = New-Object System.Windows.Forms.RichTextBox
$richTextBox.Multiline = $true
$richTextBox.ReadOnly = $true
$richTextBox.ScrollBars = 'Vertical'
$richTextBox.Dock = 'Fill'
$richTextBox.Font = New-Object System.Drawing.Font("Consolas",10)
$richTextBox.BackColor = [System.Drawing.Color]::White
$null = $logViewerPanel.Controls.Add($richTextBox)

# ----- Tab Page 2: Installed Apps -----
$installedAppsTab = New-Object System.Windows.Forms.TabPage
$installedAppsTab.Text = "Installed Apps"
$null = $tabControl.TabPages.Add($installedAppsTab)

$installedAppsPanel = New-Object System.Windows.Forms.Panel
$installedAppsPanel.Dock = 'Fill'
$null = $installedAppsTab.Controls.Add($installedAppsPanel)

$appsSearchPanel = New-Object System.Windows.Forms.Panel
$appsSearchPanel.Height = 30
$appsSearchPanel.Dock = 'Top'
$null = $installedAppsPanel.Controls.Add($appsSearchPanel)

$appsSearchLabel = New-Object System.Windows.Forms.Label
$appsSearchLabel.Text = "Search:"
$appsSearchLabel.Location = New-Object System.Drawing.Point(10, 5)
$appsSearchLabel.AutoSize = $true
$null = $appsSearchPanel.Controls.Add($appsSearchLabel)

$appsSearchTextBox = New-Object System.Windows.Forms.TextBox
$appsSearchTextBox.Location = New-Object System.Drawing.Point(70, 2)
$appsSearchTextBox.Width = 200
$null = $appsSearchPanel.Controls.Add($appsSearchTextBox)

$appsSearchButton = New-Object System.Windows.Forms.Button
$appsSearchButton.Text = "Search"
$appsSearchButton.Location = New-Object System.Drawing.Point(280, 0)
$appsSearchButton.AutoSize = $true
$null = $appsSearchPanel.Controls.Add($appsSearchButton)

$appsRefreshButton = New-Object System.Windows.Forms.Button
$appsRefreshButton.Text = "Refresh"
$appsRefreshButton.AutoSize = $true
$appsRefreshButton.Location = New-Object System.Drawing.Point(360, 0)
$null = $appsSearchPanel.Controls.Add($appsRefreshButton)
$appsRefreshButton.Add_Click({ Load-InstalledApps })

# Use a TableLayoutPanel to host the custom header and ListView.
$appsTable = New-Object System.Windows.Forms.TableLayoutPanel
$appsTable.Dock = 'Fill'
$appsTable.RowCount = 2
$appsTable.ColumnCount = 1
$appsTable.RowStyles.Clear()
$appsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))
$appsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $installedAppsPanel.Controls.Add($appsTable)

# Create header panel for Installed Apps.
$appsHeaderPanel = New-Object System.Windows.Forms.Panel
$appsHeaderPanel.Dock = 'Fill'
$appsHeaderPanel.BackColor = [System.Drawing.Color]::LightGray
$null = $appsTable.Controls.Add($appsHeaderPanel, 0, 0)

$appsNameHeader = New-Object System.Windows.Forms.Label
$appsNameHeader.Text = "Name"
$appsNameHeader.Width = 400
$appsNameHeader.Dock = 'Left'
$appsNameHeader.TextAlign = 'MiddleLeft'
$null = $appsHeaderPanel.Controls.Add($appsNameHeader)

$appsVersionHeader = New-Object System.Windows.Forms.Label
$appsVersionHeader.Text = "Version"
$appsVersionHeader.Width = 100
$appsVersionHeader.Dock = 'Left'
$appsVersionHeader.TextAlign = 'MiddleLeft'
$null = $appsHeaderPanel.Controls.Add($appsVersionHeader)

$appsPublisherHeader = New-Object System.Windows.Forms.Label
$appsPublisherHeader.Text = "Publisher"
$appsPublisherHeader.Width = 200
$appsPublisherHeader.Dock = 'Fill'
$appsPublisherHeader.TextAlign = 'MiddleLeft'
$null = $appsHeaderPanel.Controls.Add($appsPublisherHeader)

# Create the Installed Apps ListView.
$appsListView = New-Object System.Windows.Forms.ListView
$appsListView.View = [System.Windows.Forms.View]::Details
$appsListView.FullRowSelect = $true
$appsListView.GridLines = $true
$appsListView.Dock = 'Fill'
$null = $appsListView.Columns.Add("Name", 400)
$null = $appsListView.Columns.Add("Version", 100)
$null = $appsListView.Columns.Add("Publisher", 200)
# Optionally, hide the default header:
#$appsListView.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::None
$null = $appsTable.Controls.Add($appsListView, 0, 1)

function Load-InstalledApps {
    $appsListView.Items.Clear()
    $installedApps = @()
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    foreach ($path in $regPaths) {
        try {
            $keys = Get-ChildItem $path -ErrorAction SilentlyContinue
            foreach ($key in $keys) {
                try {
                    $app = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
                    if ($app.DisplayName) {
                        $installedApps += [PSCustomObject]@{
                            Name = $app.DisplayName
                            Version = $app.DisplayVersion
                            Publisher = $app.Publisher
                        }
                    }
                }
                catch { }
            }
        }
        catch { }
    }
    $installedApps = $installedApps | Sort-Object Name
    foreach ($app in $installedApps) {
        $item = New-Object System.Windows.Forms.ListViewItem($app.Name)
        if ($app.Version) { $null = $item.SubItems.Add($app.Version) } else { $null = $item.SubItems.Add("") }
        if ($app.Publisher) { $null = $item.SubItems.Add($app.Publisher) } else { $null = $item.SubItems.Add("") }
        $null = $appsListView.Items.Add($item)
    }
}
Load-InstalledApps

$appsSearchButton.Add_Click({
    $searchText = $appsSearchTextBox.Text
    if (-not [string]::IsNullOrEmpty($searchText)) {
        foreach ($item in $appsListView.Items) { $item.BackColor = [System.Drawing.Color]::White }
        $foundItem = $appsListView.FindItemWithText($searchText)
        if ($foundItem) {
            $appsListView.SelectedItems.Clear()
            $foundItem.Selected = $true
            $foundItem.BackColor = [System.Drawing.Color]::LightYellow
            $appsListView.EnsureVisible($foundItem.Index)
        } else {
            [System.Windows.Forms.MessageBox]::Show("No installed application found containing '$searchText'.", "Search", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    }
})

# ----- Tab Page 3: Running Processes -----
$processesTab = New-Object System.Windows.Forms.TabPage
$processesTab.Text = "Running Processes"
$null = $tabControl.TabPages.Add($processesTab)

$processesPanel = New-Object System.Windows.Forms.Panel
$processesPanel.Dock = 'Fill'
$null = $processesTab.Controls.Add($processesPanel)

# Create a control panel for the Processes tab.
$procControlPanel = New-Object System.Windows.Forms.Panel
$procControlPanel.Height = 40
$procControlPanel.Dock = 'Top'
$null = $processesPanel.Controls.Add($procControlPanel)

$refreshProcessesButton = New-Object System.Windows.Forms.Button
$refreshProcessesButton.Text = "Refresh Processes"
$refreshProcessesButton.AutoSize = $true
$refreshProcessesButton.Location = New-Object System.Drawing.Point(10, 5)
$null = $procControlPanel.Controls.Add($refreshProcessesButton)

$killProcessButton = New-Object System.Windows.Forms.Button
$killProcessButton.Text = "Kill Process"
$killProcessButton.AutoSize = $true
$killProcessButton.Location = New-Object System.Drawing.Point(150, 5)
$null = $procControlPanel.Controls.Add($killProcessButton)

# Use a TableLayoutPanel for the Running Processes header and ListView.
$procTable = New-Object System.Windows.Forms.TableLayoutPanel
$procTable.Dock = 'Fill'
$procTable.RowCount = 2
$procTable.ColumnCount = 1
$procTable.RowStyles.Clear()
# Increase header row height to 30 to ensure full visibility.
$procTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
$procTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$null = $processesPanel.Controls.Add($procTable)

# Create header panel for Running Processes.
$procHeaderPanel = New-Object System.Windows.Forms.Panel
$procHeaderPanel.Dock = 'Fill'
$procHeaderPanel.BackColor = [System.Drawing.Color]::LightGray
$null = $procTable.Controls.Add($procHeaderPanel, 0, 0)

$procNameHeader = New-Object System.Windows.Forms.Label
$procNameHeader.Text = "Process Name"
$procNameHeader.Width = 300
$procNameHeader.Dock = 'Left'
$procNameHeader.TextAlign = 'MiddleLeft'
$null = $procHeaderPanel.Controls.Add($procNameHeader)

$procIDHeader = New-Object System.Windows.Forms.Label
$procIDHeader.Text = "ID"
$procIDHeader.Width = 80
$procIDHeader.Dock = 'Left'
$procIDHeader.TextAlign = 'MiddleLeft'
$null = $procHeaderPanel.Controls.Add($procIDHeader)

$procMemoryHeader = New-Object System.Windows.Forms.Label
$procMemoryHeader.Text = "Memory (MB)"
$procMemoryHeader.Width = 100
$procMemoryHeader.Dock = 'Left'
$procMemoryHeader.TextAlign = 'MiddleLeft'
$null = $procHeaderPanel.Controls.Add($procMemoryHeader)

$procCPUHeader = New-Object System.Windows.Forms.Label
$procCPUHeader.Text = "CPU (s)"
$procCPUHeader.Width = 80
$procCPUHeader.Dock = 'Fill'
$procCPUHeader.TextAlign = 'MiddleLeft'
$null = $procHeaderPanel.Controls.Add($procCPUHeader)

# Create the Running Processes ListView.
$procListView = New-Object System.Windows.Forms.ListView
$procListView.View = [System.Windows.Forms.View]::Details
$procListView.FullRowSelect = $true
$procListView.GridLines = $true
$procListView.Dock = 'Fill'
$null = $procListView.Columns.Add("Process Name", 300)
$null = $procListView.Columns.Add("ID", 80)
$null = $procListView.Columns.Add("Memory (MB)", 100)
$null = $procListView.Columns.Add("CPU (s)", 80)
# Optionally hide the default header:
#$procListView.HeaderStyle = [System.Windows.Forms.ColumnHeaderStyle]::None
$null = $procTable.Controls.Add($procListView, 0, 1)

# Add a ContextMenuStrip to $procListView for right-click functionality.
$procContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$killMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem "Kill Process"
$null = $procContextMenu.Items.Add($killMenuItem)
$procListView.ContextMenuStrip = $procContextMenu

# Optional: When right-clicking, select the item under the mouse.
$procListView.Add_MouseDown({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
        $hitTestInfo = $sender.HitTest($e.X, $e.Y)
        if ($hitTestInfo.Item -ne $null) {
            $sender.SelectedItems.Clear()
            $hitTestInfo.Item.Selected = $true
        }
    }
})

# When the "Kill Process" context menu item is clicked, perform the kill.
$killMenuItem.Add_Click({
    if ($procListView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No process is selected.", "Kill Process", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        $selectedItem = $procListView.SelectedItems[0]
        $procId = [int]$selectedItem.SubItems[1].Text
        $procName = $selectedItem.Text
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to kill process '$procName' (ID: $procId)?", "Confirm Kill", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Stop-Process -Id $procId -Force
                [System.Windows.Forms.MessageBox]::Show("Process '$procName' (ID: $procId) has been killed.", "Kill Process", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to kill process '$procName': $_", "Kill Process Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            Load-Processes
        }
    }
})

function Load-Processes {
    $procListView.Items.Clear()
    try {
        $procs = Get-Process | Sort-Object -Property ProcessName
        foreach ($proc in $procs) {
            $item = New-Object System.Windows.Forms.ListViewItem($proc.ProcessName)
            $null = $item.SubItems.Add($proc.Id.ToString())
            $memMB = [math]::Round($proc.WorkingSet64 / 1MB, 2)
            $null = $item.SubItems.Add($memMB.ToString())
            try {
                $cpuSec = [math]::Round($proc.TotalProcessorTime.TotalSeconds, 2)
                $null = $item.SubItems.Add($cpuSec.ToString())
            } catch {
                $null = $item.SubItems.Add("N/A")
            }
            $null = $procListView.Items.Add($item)
        }
    }
    catch {
        if ($EnableDebug) { Write-Host "Error loading processes: $_" }
    }
}
Load-Processes

$refreshProcessesButton.Add_Click({ Load-Processes })

$killProcessButton.Add_Click({
    if ($procListView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a process to kill.", "Kill Process", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        $selectedItem = $procListView.SelectedItems[0]
        $procId = [int]$selectedItem.SubItems[1].Text
        $procName = $selectedItem.Text
        $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to kill process '$procName' (ID: $procId)?", "Confirm Kill", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Stop-Process -Id $procId -Force
                [System.Windows.Forms.MessageBox]::Show("Process '$procName' (ID: $procId) has been killed.", "Kill Process", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Failed to kill process '$procName': $_", "Kill Process Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
            Load-Processes
        }
    }
})

###############################################################
# Section 3: Log Tailing, Rehighlight Function, and Helper Functions (for Log Viewer)
###############################################################

function Append-LogLine {
    param(
        [string]$line
    )
    $start = $richTextBox.TextLength
    $richTextBox.AppendText($line + "`n")
    $richTextBox.Select($start, $line.Length)
    if ($line -match '(?i)error') {
        $richTextBox.SelectionBackColor = [System.Drawing.Color]::LightSalmon
        $global:errorCount++
    }
    elseif ($line -match '(?i)warning') {
        $richTextBox.SelectionBackColor = [System.Drawing.Color]::Yellow
        $global:warningCount++
    }
    elseif ($line -match '(?i)success') {
        $richTextBox.SelectionBackColor = [System.Drawing.Color]::LightGreen
        $global:successCount++
    }
    else {
        $customMatched = $false
        foreach ($rule in $global:customHighlightRules) {
            if ($line -match "(?i)$($rule.Keyword)") {
                $richTextBox.SelectionBackColor = $rule.Color
                $customMatched = $true
                break
            }
        }
        if (-not $customMatched) {
            $richTextBox.SelectionBackColor = [System.Drawing.Color]::White
        }
    }
    $richTextBox.SelectionLength = 0
    Update-Stats
}

function Update-Stats {
    $statsLabel.Text = "Errors: $global:errorCount, Warnings: $global:warningCount, Successes: $global:successCount"
}

function Rehighlight-ExistingText {
    $currentText = $richTextBox.Text
    $richTextBox.Clear()
    $global:errorCount = 0
    $global:warningCount = 0
    $global:successCount = 0
    $lines = $currentText -split "`r?`n"
    foreach ($line in $lines) {
        if ($line.Trim() -ne "") {
            Append-LogLine -line $line
        }
    }
}

function Load-LogFile {
    param (
        [string]$filePath
    )
    if ($filePath.ToLower().EndsWith(".evtx")) {
        $global:IsEvtx = $true
        try {
            $events = Get-WinEvent -Path $filePath -MaxEvents 1000
            $output = ""
            foreach ($event in $events) {
                $time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                $id = $event.Id
                $level = $event.LevelDisplayName
                $message = $event.Message
                $output += "$time  [$level] (ID: $id) - $message`n"
            }
            $richTextBox.Text = $output
        }
        catch {
            if ($EnableDebug) { Write-Host "Error reading EVTX file: $_" }
        }
    }
    else {
        $global:IsEvtx = $false
        try {
            $fs = [System.IO.File]::Open($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $sr = New-Object System.IO.StreamReader($fs)
            $initialContent = $sr.ReadToEnd()
            $global:lastPos = $fs.Position
            $sr.Close()
            $fs.Close()
            $global:errorCount = 0; $global:warningCount = 0; $global:successCount = 0
            $richTextBox.Clear()
            $lines = $initialContent -split "`r?`n"
            foreach ($line in $lines) {
                Append-LogLine -line $line
            }
            $richTextBox.SelectionStart = $richTextBox.TextLength
            $richTextBox.ScrollToCaret()
        }
        catch {
            if ($EnableDebug) { Write-Host "Error reading log content: $_" }
        }
    }
    Update-LogFileLabel
}

function Update-Log {
    if ($global:IsEvtx) {
        try {
            $events = Get-WinEvent -Path $global:logFilePath -MaxEvents 1000
            $output = ""
            foreach ($event in $events) {
                $time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                $id = $event.Id
                $level = $event.LevelDisplayName
                $message = $event.Message
                $output += "$time  [$level] (ID: $id) - $message`n"
            }
            $richTextBox.Invoke([System.Action]{ $richTextBox.Text = $output; $richTextBox.ScrollToCaret() })
            Update-LogFileLabel
        }
        catch {
            if ($EnableDebug) { Write-Host "Error refreshing EVTX file: $_" }
        }
        return
    }
    try {
        $fs = [System.IO.File]::Open($global:logFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $currentLength = $fs.Length
        if ($currentLength -lt $global:lastPos) {
            $global:lastPos = 0
            $richTextBox.Invoke([System.Action]{ $richTextBox.Clear() })
        }
        if ($currentLength -gt $global:lastPos) {
            $fs.Seek($global:lastPos, [System.IO.SeekOrigin]::Begin) | Out-Null
            $sr = New-Object System.IO.StreamReader($fs)
            $newText = $sr.ReadToEnd()
            $global:lastPos = $fs.Position
            $sr.Close()
            $fs.Close()
            if ($newText -and $newText.Length -gt 0) {
                $lines = $newText -split "`r?`n"
                $richTextBox.Invoke([System.Action]{
                    foreach ($line in $lines) {
                        if ($line.Trim() -ne "") { Append-LogLine -line $line }
                    }
                    $richTextBox.ScrollToCaret()
                })
                Update-LogFileLabel
            }
        }
        else {
            $fs.Close()
        }
    }
    catch {
        if ($EnableDebug) { Write-Host "Error reading log file: $_" }
    }
}

###############################################################
# Section 4: Set Up Timer and Control Event Handlers
###############################################################

function Update-BESClientStatus {
    try {
        $svc = Get-Service -Name "BESClient" -ErrorAction SilentlyContinue
        if ($svc) {
            $status = $svc.Status
        } else {
            $status = "Not Installed"
        }
        $besInfoText = "BESClient Version: $clientVersion | BESClient Service Status: $status"
        $besInfoLabel.Invoke([System.Action]{ $besInfoLabel.Text = $besInfoText })
    }
    catch {
        if ($EnableDebug) { Write-Host "Error updating BESClient status: $_" }
        $besInfoLabel.Invoke([System.Action]{ $besInfoLabel.Text = "BESClient Version: $clientVersion | BESClient Service Status: Error" })
    }
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $refreshNumeric.Value
$timer.Add_Tick({ Update-Log; Update-BESClientStatus })
$timer.Start()

$chooseFileButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = Split-Path $global:logFilePath
    $openFileDialog.Filter = "Log Files (*.log;*.txt;*.evtx)|*.log;*.txt;*.evtx|All Files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:logFilePath = $openFileDialog.FileName
        Update-LogFileLabel
        $form.Text = "BigFix Log Viewer - $(Split-Path $global:logFilePath -Leaf)"
        $global:lastPos = 0
        $richTextBox.Clear()
        $global:errorCount = 0; $global:warningCount = 0; $global:successCount = 0
        Load-LogFile -filePath $global:logFilePath
    }
})

$pauseResumeButton.Add_Click({
    if ($timer.Enabled) {
        $timer.Stop()
        $pauseResumeButton.Text = "Resume"
    }
    else {
        $timer.Start()
        $pauseResumeButton.Text = "Pause"
    }
})

$exportButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $richTextBox.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
    }
})

$clearButton.Add_Click({
    $richTextBox.Clear()
    $global:errorCount = 0; $global:warningCount = 0; $global:successCount = 0
    Update-Stats
})

$refreshNumeric.Add_ValueChanged({ $timer.Interval = $refreshNumeric.Value })

$findNextButton.Add_Click({
    $searchTerm = $searchTextBox.Text
    if ([string]::IsNullOrEmpty($searchTerm)) { return }
    $startPos = $richTextBox.SelectionStart + $richTextBox.SelectionLength
    $index = $richTextBox.Find($searchTerm, $startPos, [System.Windows.Forms.RichTextBoxFinds]::None)
    if ($index -eq -1) {
        [System.Windows.Forms.MessageBox]::Show("No further occurrences found.", "Search", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    else {
        $richTextBox.Focus()
    }
})

$openEventViewerButton.Add_Click({ Start-Process "eventvwr.exe" })

$restartBESClientButton.Add_Click({
    try {
        Restart-Service -Name "BESClient" -Force -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("BESClient service restarted successfully.","Service Restart",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to restart BESClient service: $_","Service Restart Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

function Show-ManageHighlightsDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Manage Custom Highlight Rules"
    $dialog.Size = New-Object System.Drawing.Size(400,300)
    $dialog.StartPosition = "CenterParent"
    
    $rulesListBox = New-Object System.Windows.Forms.ListBox
    $rulesListBox.Location = New-Object System.Drawing.Point(10,10)
    $rulesListBox.Size = New-Object System.Drawing.Size(360,100)
    foreach ($rule in $global:customHighlightRules) {
         $null = $rulesListBox.Items.Add("$($rule.Keyword) : $($rule.Color.Name)")
    }
    $null = $dialog.Controls.Add($rulesListBox)
    
    $keywordLabel = New-Object System.Windows.Forms.Label
    $keywordLabel.Text = "Keyword:"
    $keywordLabel.Location = New-Object System.Drawing.Point(10, 120)
    $keywordLabel.AutoSize = $true
    $null = $dialog.Controls.Add($keywordLabel)
    
    $keywordTextBox = New-Object System.Windows.Forms.TextBox
    $keywordTextBox.Location = New-Object System.Drawing.Point(80, 117)
    $keywordTextBox.Width = 150
    $null = $dialog.Controls.Add($keywordTextBox)
    
    $chooseColorButton = New-Object System.Windows.Forms.Button
    $chooseColorButton.Text = "Choose Color"
    $chooseColorButton.Location = New-Object System.Drawing.Point(240, 115)
    $chooseColorButton.AutoSize = $true
    $null = $dialog.Controls.Add($chooseColorButton)
    
    $colorLabel = New-Object System.Windows.Forms.Label
    $colorLabel.Text = "No color selected"
    $colorLabel.Location = New-Object System.Drawing.Point(10, 150)
    $colorLabel.AutoSize = $true
    $null = $dialog.Controls.Add($colorLabel)
    
    $selectedColor = [ref]([System.Drawing.Color]::Empty)
    
    $chooseColorButton.Add_Click({
      $colorDialog = New-Object System.Windows.Forms.ColorDialog
      if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
         $selectedColor.Value = $colorDialog.Color
         $colorLabel.Text = "Selected: " + $selectedColor.Value.Name
      }
    })
    
    $addRuleButton = New-Object System.Windows.Forms.Button
    $addRuleButton.Text = "Add Rule"
    $addRuleButton.Location = New-Object System.Drawing.Point(10, 180)
    $addRuleButton.AutoSize = $true
    $null = $dialog.Controls.Add($addRuleButton)
    
    $addRuleButton.Add_Click({
      if (-not [string]::IsNullOrEmpty($keywordTextBox.Text) -and (-not $selectedColor.Value.IsEmpty)) {
         $rule = [PSCustomObject]@{
             Keyword = $keywordTextBox.Text
             Color   = $selectedColor.Value
         }
         $global:customHighlightRules += $rule
         $null = $rulesListBox.Items.Add("$($rule.Keyword) : $($rule.Color.Name)")
         $keywordTextBox.Clear()
         $selectedColor.Value = [System.Drawing.Color]::Empty
         $colorLabel.Text = "No color selected"
         Rehighlight-ExistingText
      } else {
         [System.Windows.Forms.MessageBox]::Show("Please enter a keyword and select a color.","Input Required",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
      }
    })
    
    $removeRuleButton = New-Object System.Windows.Forms.Button
    $removeRuleButton.Text = "Remove Selected Rule"
    $removeRuleButton.Location = New-Object System.Drawing.Point(120, 180)
    $removeRuleButton.AutoSize = $true
    $null = $dialog.Controls.Add($removeRuleButton)
    
    $removeRuleButton.Add_Click({
      if ($rulesListBox.SelectedIndex -ge 0) {
         $index = $rulesListBox.SelectedIndex
         $global:customHighlightRules = $global:customHighlightRules | Where-Object { $global:customHighlightRules.IndexOf($_) -ne $index }
         $null = $rulesListBox.Items.RemoveAt($index)
         Rehighlight-ExistingText
      }
    })
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(10, 220)
    $okButton.AutoSize = $true
    $null = $dialog.Controls.Add($okButton)
    
    $okButton.Add_Click({ $dialog.Close() })
    
    $dialog.ShowDialog() | Out-Null
}

$manageHighlightsButton.Add_Click({ Show-ManageHighlightsDialog })

###############################################################
# Section 5: Initialize and Run the Application
###############################################################

try {
    Load-LogFile -filePath $global:logFilePath
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to load log file: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
$tabControl.SelectedTab = $logViewerTab

$form.Add_Shown({
    if ($EnableDebug) { Write-Host "Form shown. Loading log file from $global:logFilePath" }
    try {
        Load-LogFile -filePath $global:logFilePath
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to load log file: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$form.Add_FormClosing({
    $newSettings = @{
       RefreshInterval = $refreshNumeric.Value
       WindowSize = @{ Width = $form.Width; Height = $form.Height }
       WindowLocation = @{ X = $form.Location.X; Y = $form.Location.Y }
       LastLogFile = $global:logFilePath
       CustomHighlightRules = @()
    }
    foreach ($rule in $global:customHighlightRules) {
       $newSettings.CustomHighlightRules += @{
            Keyword = $rule.Keyword
            ColorName = $rule.Color.Name
       }
    }
    $newSettings | ConvertTo-Json -Depth 5 | Out-File -FilePath $global:configFile -Encoding UTF8
})

###############################################################
# Section 6: Scale Fonts for Buttons and Labels (Scaling Factor Adjusted)
###############################################################

function Scale-Font {
    param (
        [System.Windows.Forms.Control]$control,
        [double]$scaleFactor
    )
    if ($control.Font -ne $null) {
        $newSize = ([double]$control.Font.Size) * $scaleFactor
        $control.Font = New-Object System.Drawing.Font($control.Font.FontFamily, $newSize, $control.Font.Style)
    }
    foreach ($child in $control.Controls) {
        Scale-Font -control $child -scaleFactor $scaleFactor
    }
}
Scale-Font -control $systemInfoPanel -scaleFactor 1.0
Scale-Font -control $logFilePanel -scaleFactor 1.0
Scale-Font -control $controlPanel -scaleFactor 1.0

[void][System.Windows.Forms.Application]::Run($form)
[void]$timer.Stop()
[void]$timer.Dispose()
