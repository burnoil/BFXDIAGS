# Set this flag to $true to enable debug messages; $false to disable.
$EnableDebug = $true # Kept true for debugging

# Define debug log file path early
$debugLogFile = Join-Path $PSScriptRoot "BFXDIAGS_Debug.log"

function Write-DebugLog {
    param (
        [string]$Message,
        [string]$Level = "Error" # Support Info, Warning, Error levels
    )
    if ($EnableDebug) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        try {
            "$timestamp [$Level] : $Message" | Out-File -FilePath $debugLogFile -Append -Encoding UTF8
        } catch {
            # Fallback to console if file write fails
            Write-Host "$timestamp [$Level] : Debug Log Write Error: $_"
        }
    }
}

# Log environment details for diagnostics
Write-DebugLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)" -Level "Info"
Write-DebugLog -Message ".NET Framework Version: $(try { (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction Stop).Version } catch { 'Unknown' })" -Level "Info"
Write-DebugLog -Message "Script Path: $PSScriptRoot" -Level "Info"
Write-DebugLog -Message "Profile Script: $(if (Test-Path $PROFILE) { $PROFILE } else { 'None' })" -Level "Info"
Write-DebugLog -Message "Loaded Assemblies: $(([System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -like '*Windows*' } | ForEach-Object { $_.GetName().Name + ' ' + $_.GetName().Version } | Select-Object -First 10) -join ', ')" -Level "Info"

# Wrap the entire script in a try-catch for robust error handling
try {
    # Initialize Windows Forms settings
    function Initialize-WindowsForms {
        try {
            # Load assemblies
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            Write-DebugLog -Message "Windows Forms assemblies loaded successfully" -Level "Info"

            # Enable visual styles
            [System.Windows.Forms.Application]::EnableVisualStyles()
            Write-DebugLog -Message "Visual styles enabled successfully" -Level "Info"

            # Skip SetCompatibleTextRenderingDefault to avoid persistent errors
            Write-DebugLog -Message "Skipping SetCompatibleTextRenderingDefault to prevent initialization errors" -Level "Info"
            return $true
        } catch {
            Write-DebugLog -Message "Error initializing Windows Forms: $_" -Level "Error"
            # Use Write-Host for critical errors if MessageBox fails
            Write-Host "Failed to initialize Windows Forms: $_"
            return $false
        }
    }

    # Call initialization
    if (-not (Initialize-WindowsForms)) {
        Write-DebugLog -Message "Windows Forms initialization failed, exiting" -Level "Error"
        exit
    }

    ###############################################################
    # Section 1: Setup – Configuration, Log File, System Info
    ###############################################################

    # Initialize global configuration object.
    $global:AppConfig = @{
        ConfigFile = Join-Path $PSScriptRoot "BigFixLogViewerSettings.json"
        LogFilePath = $null
        LastPos = 0
        ErrorCount = 0
        WarningCount = 0
        SuccessCount = 0
        IsEvtx = $false
        CustomHighlightRules = @()
        LastModified = $null
    }

    # Define default values.
    $defaultLogDir = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
    $defaultRefreshInterval = 1000
    $defaultWindowSize = @{ Width = 1000; Height = 750 }
    $defaultWindowLocation = @{ X = 100; Y = 100 }
    $defaultCustomHighlightRules = @()

    # Load saved settings with validation.
    $savedSettings = $null
    if (Test-Path $global:AppConfig.ConfigFile) {
        try {
            $jsonContent = Get-Content $global:AppConfig.ConfigFile -Raw -ErrorAction Stop
            if (-not $jsonContent.Trim()) {
                throw "Config file is empty"
            }
            $savedSettings = $jsonContent | ConvertFrom-Json -ErrorAction Stop
            Write-DebugLog -Message "Successfully loaded settings from $($global:AppConfig.ConfigFile)" -Level "Info"
        } catch {
            Write-DebugLog -Message "Error reading or parsing config file: $_" -Level "Error"
            $savedSettings = $null # Ensure null to use defaults
        }
    } else {
        Write-DebugLog -Message "Config file not found at $($global:AppConfig.ConfigFile), using defaults" -Level "Info"
    }

    # Fix: Simplified Validate-FilePath to focus on existence and accessibility
    function Validate-FilePath {
        param ([string]$filePath)
        if (-not $filePath) {
            Write-DebugLog -Message "File path validation failed: Path is empty" -Level "Warning"
            return $false
        }
        try {
            if (Test-Path $filePath -PathType Leaf -ErrorAction Stop) {
                Write-DebugLog -Message "File path validated: $filePath" -Level "Info"
                return $true
            } else {
                Write-DebugLog -Message "File path validation failed: $filePath does not exist or is not a file" -Level "Warning"
                return $false
            }
        } catch {
            Write-DebugLog -Message "File path validation error for '$filePath': $_" -Level "Error"
            return $false
        }
    }

    # Fix: Improved log file selection to ensure most recent log is chosen
    if ($savedSettings -and $savedSettings.PSObject.Properties.Name -contains "LastLogFile" -and (Validate-FilePath $savedSettings.LastLogFile)) {
        $global:AppConfig.LogFilePath = $savedSettings.LastLogFile
        Write-DebugLog -Message "Using saved log file: $($global:AppConfig.LogFilePath)" -Level "Info"
    } else {
        Write-DebugLog -Message "No valid saved log file, selecting most recent log from $defaultLogDir" -Level "Info"
        if (-not (Test-Path $defaultLogDir)) {
            [System.Windows.Forms.MessageBox]::Show("Log directory not found:`n$defaultLogDir", "Error", 'OK', 'Error')
            Write-DebugLog -Message "Log directory not found: $defaultLogDir" -Level "Error"
            exit
        }
        try {
            $latestLogFile = Get-ChildItem -Path $defaultLogDir -File -ErrorAction Stop |
                             Where-Object { $_.Extension -match '\.log$|\.txt$|\.evtx$' } | # Filter for common log extensions
                             Sort-Object LastWriteTime -Descending |
                             Select-Object -First 1
            if (-not $latestLogFile) {
                [System.Windows.Forms.MessageBox]::Show("No log files found in:`n$defaultLogDir", "Error", 'OK', 'Error')
                Write-DebugLog -Message "No log files found in: $defaultLogDir" -Level "Error"
                exit
            }
            $global:AppConfig.LogFilePath = $latestLogFile.FullName
            Write-DebugLog -Message "Selected most recent log file: $($global:AppConfig.LogFilePath)" -Level "Info"
        } catch {
            Write-DebugLog -Message "Error finding latest log file: $_" -Level "Error"
            [System.Windows.Forms.MessageBox]::Show("Error accessing log directory: $_", "Error", 'OK', 'Error')
            exit
        }
    }
    try {
        $global:AppConfig.LastModified = (Get-Item $global:AppConfig.LogFilePath -ErrorAction Stop).LastWriteTime
        Write-DebugLog -Message "Log file last modified: $($global:AppConfig.LastModified)" -Level "Info"
    } catch {
        Write-DebugLog -Message "Error getting last modified time for '$($global:AppConfig.LogFilePath)': $_" -Level "Error"
        $global:AppConfig.LastModified = [DateTime]::Now
    }

    # Load system information with error handling.
    $computerName = $env:COMPUTERNAME
    $ipAddresses = "Unknown"
    try {
        $ipAddresses = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
                        Where-Object { $_.IPAddress -notmatch '^127\.0\.0\.1' -and $_.IPAddress -notmatch '^169\.254\.' } |
                        Select-Object -ExpandProperty IPAddress) -join ', '
    } catch {
        try {
            $ipAddresses = ([System.Net.Dns]::GetHostAddresses($computerName) |
                            Where-Object { $_.AddressFamily -eq 'InterNetwork' -and $_.IPAddressToString -notmatch '^127\.0\.0\.1' -and $_.IPAddressToString -notmatch '^169\.254\.' } |
                            ForEach-Object { $_.IPAddressToString }) -join ', '
        } catch {
            Write-DebugLog -Message "Error retrieving IP addresses: $_" -Level "Error"
        }
    }
    $diskSizeGB = 0
    $freeSpaceGB = 0
    try {
        $disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" -ErrorAction Stop
        if ($disk) {
            $diskSizeGB = [math]::Round($disk.Size / 1GB, 2)
            $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        }
    } catch {
        Write-DebugLog -Message "Error retrieving disk information: $_" -Level "Error"
    }
    $relayServer = "Not Found"
    try {
        $relayServerKeyPath = "HKLM:\SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client\__RelayServer1"
        if (Test-Path $relayServerKeyPath -ErrorAction Stop) {
            $regKey = Get-ItemProperty -Path $relayServerKeyPath -ErrorAction Stop
            if ($regKey.PSObject.Properties.Name -contains "value") {
                $relayServer = $regKey.value
            }
        } else {
            Write-DebugLog -Message "Relay server registry key not found: $relayServerKeyPath" -Level "Info"
        }
    } catch {
        Write-DebugLog -Message "Unexpected error retrieving relay server from registry: $_" -Level "Error"
    }
    $clientExePath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\BESClient.exe"
    $clientVersion = "Not Found"
    if (Test-Path $clientExePath) {
        try {
            $clientVersion = (Get-Item $clientExePath -ErrorAction Stop).VersionInfo.FileVersion
        } catch {
            Write-DebugLog -Message "Error retrieving BESClient version: $_" -Level "Error"
        }
    }

    $systemInfoText = "Machine Name: $computerName`n" +
                      "IP Addresses: $ipAddresses`n" +
                      "Disk Size (C:): $diskSizeGB GB`n" +
                      "Free Space (C:): $freeSpaceGB GB`n" +
                      "Relay Server: $relayServer"

    ###############################################################
    # Section 2: Build the Main Form and TabControl
    ###############################################################

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "BigFix Log Viewer - $(Split-Path $global:AppConfig.LogFilePath -Leaf)"
    $form.Width = if ($savedSettings -and $savedSettings.PSObject.Properties.Name -contains "WindowSize" -and $savedSettings.WindowSize.Width) { $savedSettings.WindowSize.Width } else { $defaultWindowSize.Width }
    $form.Height = if ($savedSettings -and $savedSettings.PSObject.Properties.Name -contains "WindowSize" -and $savedSettings.WindowSize.Height) { $savedSettings.WindowSize.Height } else { $defaultWindowSize.Height }
    $form.Location = if ($savedSettings -and $savedSettings.PSObject.Properties.Name -contains "WindowLocation") { New-Object System.Drawing.Point($savedSettings.WindowLocation.X, $savedSettings.WindowLocation.Y) } else { New-Object System.Drawing.Point($defaultWindowLocation.X, $defaultWindowLocation.Y) }
    $form.StartPosition = 'Manual'

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Dock = 'Fill'
    $form.Controls.Add($tabControl)

    # ----- Tab Page 1: Log Viewer -----
    $logViewerTab = New-Object System.Windows.Forms.TabPage
    $logViewerTab.Text = "Log Viewer"
    $tabControl.TabPages.Add($logViewerTab)

    $logViewerPanel = New-Object System.Windows.Forms.Panel
    $logViewerPanel.Dock = 'Fill'
    $logViewerTab.Controls.Add($logViewerPanel)

    # -- System Info Panel --
    $systemInfoPanel = New-Object System.Windows.Forms.Panel
    $systemInfoPanel.Dock = 'Top'
    $systemInfoPanel.Height = 120
    $systemInfoPanel.BackColor = [System.Drawing.Color]::Black
    $systemInfoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $logViewerPanel.Controls.Add($systemInfoPanel)

    $flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $flowPanel.Dock = 'Fill'
    $flowPanel.FlowDirection = 'TopDown'
    $flowPanel.WrapContents = $false
    $systemInfoPanel.Controls.Add($flowPanel)

    $systemInfoLabel = New-Object System.Windows.Forms.Label
    $systemInfoLabel.AutoSize = $true
    $systemInfoLabel.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $systemInfoLabel.ForeColor = [System.Drawing.Color]::Cyan
    $systemInfoLabel.Text = $systemInfoText
    $flowPanel.Controls.Add($systemInfoLabel)

    $besInfoLabel = New-Object System.Windows.Forms.Label
    $besInfoLabel.AutoSize = $true
    $besInfoLabel.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
    $besInfoLabel.ForeColor = [System.Drawing.Color]::Cyan
    $besInfoLabel.Text = "BESClient Version: $clientVersion | BESClient Service Status: Unknown"
    $flowPanel.Controls.Add($besInfoLabel)

    # -- Log File Panel --
    $logFilePanel = New-Object System.Windows.Forms.Panel
    $logFilePanel.Dock = 'Top'
    $logFilePanel.Height = 30
    $logViewerPanel.Controls.Add($logFilePanel)

    function Update-LogFileLabel {
        try {
            $fileItem = Get-Item $global:AppConfig.LogFilePath -ErrorAction Stop
            $lastWrite = $fileItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
            if ($logFileLabel.InvokeRequired) {
                $logFileLabel.Invoke([System.Action]{ $logFileLabel.Text = "Log File: $($global:AppConfig.LogFilePath) (Last Modified: $lastWrite)" })
            } else {
                $logFileLabel.Text = "Log File: $($global:AppConfig.LogFilePath) (Last Modified: $lastWrite)"
            }
        } catch {
            Write-DebugLog -Message "Error updating log file label: $_" -Level "Error"
            if ($logFileLabel.InvokeRequired) {
                $logFileLabel.Invoke([System.Action]{ $logFileLabel.Text = "Log File: $($global:AppConfig.LogFilePath) (Last Modified: Unknown)" })
            } else {
                $logFileLabel.Text = "Log File: $($global:AppConfig.LogFilePath) (Last Modified: Unknown)"
            }
        }
    }
    $logFileLabel = New-Object System.Windows.Forms.Label
    $logFileLabel.AutoSize = $true
    $logFileLabel.Font = New-Object System.Drawing.Font("Consolas", 9)
    Update-LogFileLabel
    $logFileLabel.Location = New-Object System.Drawing.Point(10, 5)
    $logFilePanel.Controls.Add($logFileLabel)

    # -- Control Panel --
    $controlPanel = New-Object System.Windows.Forms.Panel
    $controlPanel.Dock = 'Top'
    $controlPanel.Height = 100
    $controlPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
    $logViewerPanel.Controls.Add($controlPanel)

    # Row 1: Search and Primary Buttons
    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "Search:"
    $searchLabel.Location = New-Object System.Drawing.Point(10, 5)
    $searchLabel.AutoSize = $true
    $controlPanel.Controls.Add($searchLabel)

    $searchTextBox = New-Object System.Windows.Forms.TextBox
    $searchTextBox.Location = New-Object System.Drawing.Point(70, 2)
    $searchTextBox.Width = 150
    $controlPanel.Controls.Add($searchTextBox)

    $caseSensitiveCheckBox = New-Object System.Windows.Forms.CheckBox
    $caseSensitiveCheckBox.Text = "Case Sensitive"
    $caseSensitiveCheckBox.Location = New-Object System.Drawing.Point(230, 2)
    $caseSensitiveCheckBox.AutoSize = $true
    $controlPanel.Controls.Add($caseSensitiveCheckBox)

    $findNextButton = New-Object System.Windows.Forms.Button
    $findNextButton.Text = "Find Next"
    $findNextButton.Location = New-Object System.Drawing.Point(350, 2)
    $findNextButton.AutoSize = $true
    $controlPanel.Controls.Add($findNextButton)

    # Row 2: Filter
    $filterLabel = New-Object System.Windows.Forms.Label
    $filterLabel.Text = "Filter:"
    $filterLabel.Location = New-Object System.Drawing.Point(10, 30)
    $filterLabel.AutoSize = $true
    $controlPanel.Controls.Add($filterLabel)

    $filterComboBox = New-Object System.Windows.Forms.ComboBox
    $filterComboBox.Location = New-Object System.Drawing.Point(70, 28)
    $filterComboBox.Width = 150
    $filterComboBox.Items.AddRange(@("All", "Errors", "Warnings", "Success", "Custom"))
    $filterComboBox.SelectedIndex = 0
    $controlPanel.Controls.Add($filterComboBox)

    $filterTextBox = New-Object System.Windows.Forms.TextBox
    $filterTextBox.Location = New-Object System.Drawing.Point(230, 28)
    $filterTextBox.Width = 150
    $filterTextBox.Visible = $false
    $controlPanel.Controls.Add($filterTextBox)

    # Row 3: Other Controls
    $pauseResumeButton = New-Object System.Windows.Forms.Button
    $pauseResumeButton.Text = "Pause"
    $pauseResumeButton.Location = New-Object System.Drawing.Point(10, 58)
    $pauseResumeButton.AutoSize = $true
    $controlPanel.Controls.Add($pauseResumeButton)

    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "Export Log"
    $exportButton.Location = New-Object System.Drawing.Point(100, 58)
    $exportButton.AutoSize = $true
    $controlPanel.Controls.Add($exportButton)

    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Text = "Clear Log"
    $clearButton.Location = New-Object System.Drawing.Point(190, 58)
    $clearButton.AutoSize = $true
    $controlPanel.Controls.Add($clearButton)

    $chooseFileButton = New-Object System.Windows.Forms.Button
    $chooseFileButton.Text = "Choose Log File"
    $chooseFileButton.AutoSize = $true
    $chooseFileButton.Location = New-Object System.Drawing.Point(280, 58)
    $controlPanel.Controls.Add($chooseFileButton)

    $manageHighlightsButton = New-Object System.Windows.Forms.Button
    $manageHighlightsButton.Text = "Manage Highlights"
    $manageHighlightsButton.AutoSize = $true
    $manageHighlightsButton.Location = New-Object System.Drawing.Point(380, 58)
    $controlPanel.Controls.Add($manageHighlightsButton)

    $openEventViewerButton = New-Object System.Windows.Forms.Button
    $openEventViewerButton.Text = "Open Event Viewer"
    $openEventViewerButton.AutoSize = $true
    $openEventViewerButton.Location = New-Object System.Drawing.Point(500, 58)
    $controlPanel.Controls.Add($openEventViewerButton)

    $restartBESClientButton = New-Object System.Windows.Forms.Button
    $restartBESClientButton.Text = "Restart BESClient"
    $restartBESClientButton.AutoSize = $true
    $restartBESClientButton.Location = New-Object System.Drawing.Point(620, 58)
    $restartBESClientButton.BackColor = [System.Drawing.Color]::LightGreen
    $controlPanel.Controls.Add($restartBESClientButton)

    # Row 4: Refresh Interval and Stats
    $refreshLabel = New-Object System.Windows.Forms.Label
    $refreshLabel.Text = "Refresh Interval (ms):"
    $refreshLabel.Location = New-Object System.Drawing.Point(10, 85)
    $refreshLabel.AutoSize = $true
    $controlPanel.Controls.Add($refreshLabel)

    $refreshNumeric = New-Object System.Windows.Forms.NumericUpDown
    $refreshNumeric.Location = New-Object System.Drawing.Point(140, 83)
    $refreshNumeric.Minimum = 100
    $refreshNumeric.Maximum = 5000
    $refreshNumeric.Value = if ($savedSettings -and $savedSettings.PSObject.Properties.Name -contains "RefreshInterval") { $savedSettings.RefreshInterval } else { $defaultRefreshInterval }
    $controlPanel.Controls.Add($refreshNumeric)

    $statsLabel = New-Object System.Windows.Forms.Label
    $statsLabel.Text = "Errors: 0, Warnings: 0, Success: 0"
    $statsLabel.Location = New-Object System.Drawing.Point(320, 85)
    $statsLabel.AutoSize = $true
    $controlPanel.Controls.Add($statsLabel)

    # -- Log Content Panel --
    $richTextBox = New-Object System.Windows.Forms.RichTextBox
    $richTextBox.Multiline = $true
    $richTextBox.ReadOnly = $true
    $richTextBox.ScrollBars = 'Vertical'
    $richTextBox.Dock = 'Fill'
    $richTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $richTextBox.BackColor = [System.Drawing.Color]::White
    $logViewerPanel.Controls.Add($richTextBox)

    # ----- Tab Page 2: Installed Apps -----
    $installedAppsTab = New-Object System.Windows.Forms.TabPage
    $installedAppsTab.Text = "Installed Apps"
    $tabControl.TabPages.Add($installedAppsTab)

    $installedAppsPanel = New-Object System.Windows.Forms.Panel
    $installedAppsPanel.Dock = 'Fill'
    $installedAppsTab.Controls.Add($installedAppsPanel)

    $appsSearchPanel = New-Object System.Windows.Forms.Panel
    $appsSearchPanel.Height = 30
    $appsSearchPanel.Dock = 'Top'
    $installedAppsPanel.Controls.Add($appsSearchPanel)

    $appsSearchLabel = New-Object System.Windows.Forms.Label
    $appsSearchLabel.Text = "Search:"
    $appsSearchLabel.Location = New-Object System.Drawing.Point(10, 5)
    $appsSearchLabel.AutoSize = $true
    $appsSearchPanel.Controls.Add($appsSearchLabel)

    $appsSearchTextBox = New-Object System.Windows.Forms.TextBox
    $appsSearchTextBox.Location = New-Object System.Drawing.Point(70, 2)
    $appsSearchTextBox.Width = 200
    $appsSearchPanel.Controls.Add($appsSearchTextBox)

    $appsSearchButton = New-Object System.Windows.Forms.Button
    $appsSearchButton.Text = "Search"
    $appsSearchButton.Location = New-Object System.Drawing.Point(280, 0)
    $appsSearchButton.AutoSize = $true
    $appsSearchPanel.Controls.Add($appsSearchButton)

    $appsRefreshButton = New-Object System.Windows.Forms.Button
    $appsRefreshButton.Text = "Refresh"
    $appsRefreshButton.AutoSize = $true
    $appsRefreshButton.Location = New-Object System.Drawing.Point(360, 0)
    $appsSearchPanel.Controls.Add($appsRefreshButton)

    $appsTable = New-Object System.Windows.Forms.TableLayoutPanel
    $appsTable.Dock = 'Fill'
    $appsTable.RowCount = 2
    $appsTable.ColumnCount = 1
    $appsTable.RowStyles.Clear()
    $appsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 25)))
    $appsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $installedAppsPanel.Controls.Add($appsTable)

    $appsHeaderPanel = New-Object System.Windows.Forms.Panel
    $appsHeaderPanel.Dock = 'Fill'
    $appsHeaderPanel.BackColor = [System.Drawing.Color]::LightGray
    $appsTable.Controls.Add($appsHeaderPanel, 0, 0)

    $appsNameHeader = New-Object System.Windows.Forms.Label
    $appsNameHeader.Text = "Name"
    $appsNameHeader.Width = 400
    $appsNameHeader.Dock = 'Left'
    $appsNameHeader.TextAlign = 'MiddleLeft'
    $appsHeaderPanel.Controls.Add($appsNameHeader)

    $appsVersionHeader = New-Object System.Windows.Forms.Label
    $appsVersionHeader.Text = "Version"
    $appsVersionHeader.Width = 100
    $appsVersionHeader.Dock = 'Left'
    $appsVersionHeader.TextAlign = 'MiddleLeft'
    $appsHeaderPanel.Controls.Add($appsVersionHeader)

    $appsPublisherHeader = New-Object System.Windows.Forms.Label
    $appsPublisherHeader.Text = "Publisher"
    $appsPublisherHeader.Width = 200
    $appsPublisherHeader.Dock = 'Fill'
    $appsPublisherHeader.TextAlign = 'MiddleLeft'
    $appsHeaderPanel.Controls.Add($appsPublisherHeader)

    $appsListView = New-Object System.Windows.Forms.ListView
    $appsListView.View = [System.Windows.Forms.View]::Details
    $appsListView.FullRowSelect = $true
    $appsListView.GridLines = $true
    $appsListView.Dock = 'Fill'
    $appsListView.Columns.Add("Name", 400)
    $appsListView.Columns.Add("Version", 100)
    $appsListView.Columns.Add("Publisher", 200)
    $appsTable.Controls.Add($appsListView, 0, 1)

    # ----- Tab Page 3: Running Processes -----
    $processesTab = New-Object System.Windows.Forms.TabPage
    $processesTab.Text = "Running Processes"
    $tabControl.TabPages.Add($processesTab)

    $processesPanel = New-Object System.Windows.Forms.Panel
    $processesPanel.Dock = 'Fill'
    $processesTab.Controls.Add($processesPanel)

    $procControlPanel = New-Object System.Windows.Forms.Panel
    $procControlPanel.Height = 40
    $procControlPanel.Dock = 'Top'
    $processesPanel.Controls.Add($procControlPanel)

    $refreshProcessesButton = New-Object System.Windows.Forms.Button
    $refreshProcessesButton.Text = "Refresh Processes"
    $refreshProcessesButton.AutoSize = $true
    $refreshProcessesButton.Location = New-Object System.Drawing.Point(10, 5)
    $procControlPanel.Controls.Add($refreshProcessesButton)

    $killProcessButton = New-Object System.Windows.Forms.Button
    $killProcessButton.Text = "Kill Process"
    $killProcessButton.AutoSize = $true
    $killProcessButton.Location = New-Object System.Drawing.Point(150, 5)
    $procControlPanel.Controls.Add($killProcessButton)

    $procTable = New-Object System.Windows.Forms.TableLayoutPanel
    $procTable.Dock = 'Fill'
    $procTable.RowCount = 2
    $procTable.ColumnCount = 1
    $procTable.RowStyles.Clear()
    $procTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    $procTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $processesPanel.Controls.Add($procTable)

    $procHeaderPanel = New-Object System.Windows.Forms.Panel
    $procHeaderPanel.Dock = 'Fill'
    $procHeaderPanel.BackColor = [System.Drawing.Color]::LightGray
    $procTable.Controls.Add($procHeaderPanel, 0, 0)

    $procNameHeader = New-Object System.Windows.Forms.Label
    $procNameHeader.Text = "Process Name"
    $procNameHeader.Width = 300
    $procNameHeader.Dock = 'Left'
    $procNameHeader.TextAlign = 'MiddleLeft'
    $procHeaderPanel.Controls.Add($procNameHeader)

    $procIDHeader = New-Object System.Windows.Forms.Label
    $procIDHeader.Text = "ID"
    $procIDHeader.Width = 80
    $procIDHeader.Dock = 'Left'
    $procIDHeader.TextAlign = 'MiddleLeft'
    $procHeaderPanel.Controls.Add($procIDHeader)

    $procMemoryHeader = New-Object System.Windows.Forms.Label
    $procMemoryHeader.Text = "Memory (MB)"
    $procMemoryHeader.Width = 100
    $procMemoryHeader.Dock = 'Left'
    $procMemoryHeader.TextAlign = 'MiddleLeft'
    $procHeaderPanel.Controls.Add($procMemoryHeader)

    $procCPUHeader = New-Object System.Windows.Forms.Label
    $procCPUHeader.Text = "CPU (s)"
    $procCPUHeader.Width = 80
    $procCPUHeader.Dock = 'Fill'
    $procCPUHeader.TextAlign = 'MiddleLeft'
    $procHeaderPanel.Controls.Add($procCPUHeader)

    $procListView = New-Object System.Windows.Forms.ListView
    $procListView.View = [System.Windows.Forms.View]::Details
    $procListView.FullRowSelect = $true
    $procListView.GridLines = $true
    $procListView.Dock = 'Fill'
    $procListView.Columns.Add("Process Name", 300)
    $procListView.Columns.Add("ID", 80)
    $procListView.Columns.Add("Memory (MB)", 100)
    $procListView.Columns.Add("CPU (s)", 80)
    $procTable.Controls.Add($procListView, 0, 1)

    $procContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $killMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem "Kill Process"
    $procContextMenu.Items.Add($killMenuItem)
    $procListView.ContextMenuStrip = $procContextMenu

    $procListView.Add_MouseDown({
        param($sender, $e)
        try {
            if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) {
                $hitTestInfo = $sender.HitTest($e.X, $e.Y)
                if ($hitTestInfo.Item -ne $null) {
                    $sender.SelectedItems.Clear()
                    $null = $hitTestInfo.Item.Selected = $true # Suppress output
                }
            }
        } catch {
            Write-DebugLog -Message "Error in procListView MouseDown event: $_" -Level "Error"
        }
    })

    ###############################################################
    # Section 3: Helper Functions
    ###############################################################

    function Scale-Font {
        param (
            [System.Windows.Forms.Control]$control,
            [double]$scaleFactor
        )
        try {
            if ($control -and $control.Font -and -not $control.IsDisposed) {
                $newSize = ([double]$control.Font.Size) * $scaleFactor
                $control.Font = New-Object System.Drawing.Font($control.Font.FontFamily, $newSize, $control.Font.Style)
            }
            foreach ($child in $control.Controls) {
                Scale-Font -control $child -scaleFactor $scaleFactor
            }
        } catch {
            Write-DebugLog -Message "Error scaling font for control: $_" -Level "Error"
        }
    }

    function Populate-ListView {
        param (
            [System.Windows.Forms.ListView]$listView,
            [array]$items,
            [string[]]$properties
        )
        try {
            if (-not $listView -or $listView.IsDisposed) { return }
            $listView.Items.Clear()
            foreach ($item in $items) {
                if (-not $item) { continue }
                $listViewItem = New-Object System.Windows.Forms.ListViewItem($item.($properties[0]))
                for ($i = 1; $i -lt $properties.Count; $i++) {
                    $value = if ($item.($properties[$i])) { $item.($properties[$i]) } else { "" }
                    $null = $listViewItem.SubItems.Add($value) # Suppress output
                }
                $null = $listView.Items.Add($listViewItem) # Suppress output
            }
        } catch {
            Write-DebugLog -Message "Error populating ListView: $_" -Level "Error"
        }
    }

    function Read-LogFileContent {
        param (
            [string]$filePath,
            [long]$startPos,
            [int]$maxLines = 1000
        )
        $retryCount = 3
        $retryDelay = 500 # ms
        for ($i = 0; $i -lt $retryCount; $i++) {
            try {
                $fs = [System.IO.File]::Open($filePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $fs.Seek($startPos, [System.IO.SeekOrigin]::Begin) | Out-Null
                $sr = New-Object System.IO.StreamReader($fs)
                $lines = New-Object System.Collections.Generic.List[string]
                while (-not $sr.EndOfStream -and $lines.Count -lt $maxLines) {
                    $line = $sr.ReadLine()
                    if ($line) { $lines.Add($line) }
                }
                $newPos = $fs.Position
                $sr.Close()
                $fs.Close()
                return @{ Content = $lines; NewPosition = $newPos }
            } catch {
                Write-DebugLog -Message "Attempt $($i+1) failed reading log file '$filePath': $_" -Level "Error"
                if ($i -lt $retryCount - 1) {
                    Start-Sleep -Milliseconds $retryDelay
                }
            }
        }
        Write-DebugLog -Message "Failed to read log file '$filePath'Thursday, May 22, 2025 after $retryCount attempts" -Level "Error"
        return $null
    }

    function Append-LogLines {
        param ([string[]]$lines)
        try {
            if (-not $richTextBox -or $richTextBox.IsDisposed) { return }
            $filter = if ($filterComboBox -and -not $filterComboBox.IsDisposed) { $filterComboBox.SelectedItem } else { "All" }
            $customFilter = if ($filterTextBox -and -not $filterTextBox.IsDisposed) { $filterTextBox.Text } else { "" }
            foreach ($line in $lines) {
                if ($line.Trim() -eq "") { continue }
                $shouldDisplay = $true
                if ($filter -eq "Errors" -and -not ($line -match '(?i)error')) { $shouldDisplay = $false }
                elseif ($filter -eq "Warnings" -and -not ($line -match '(?i)warning')) { $shouldDisplay = $false }
                elseif ($filter -eq "Success" -and -not ($line -match '(?i)success')) { $shouldDisplay = $false }
                elseif ($filter -eq "Custom" -and -not [string]::IsNullOrEmpty($customFilter) -and -not ($line -match "(?i)$customFilter")) { $shouldDisplay = $false }
                if (-not $shouldDisplay) { continue }
                $start = $richTextBox.TextLength
                $richTextBox.AppendText($line + "`n")
                $richTextBox.Select($start, $line.Length)
                if ($line -match '(?i)error') {
                    $richTextBox.SelectionBackColor = [System.Drawing.Color]::LightSalmon
                    $global:AppConfig.ErrorCount++
                }
                elseif ($line -match '(?i)warning') {
                    $richTextBox.SelectionBackColor = [System.Drawing.Color]::Yellow
                    $global:AppConfig.WarningCount++
                }
                elseif ($line -match '(?i)success') {
                    $richTextBox.SelectionBackColor = [System.Drawing.Color]::LightGreen
                    $global:AppConfig.SuccessCount++
                }
                else {
                    $customMatched = $false
                    foreach ($rule in $global:AppConfig.CustomHighlightRules) {
                        if ($rule -and $line -match "(?i)$($rule.Keyword)") {
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
            }
            Update-Stats
        } catch {
            Write-DebugLog -Message "Error appending log lines: $_" -Level "Error"
        }
    }

    function Update-Stats {
        try {
            if (-not $statsLabel -or $statsLabel.IsDisposed) { return }
            $text = "Errors: $($global:AppConfig.ErrorCount), Warnings: $($global:AppConfig.WarningCount), Success: $($global:AppConfig.SuccessCount)"
            if ($statsLabel.InvokeRequired) {
                $statsLabel.Invoke([System.Action]{ $statsLabel.Text = $text })
            } else {
                $statsLabel.Text = $text
            }
        } catch {
            Write-DebugLog -Message "Error updating stats: $_" -Level "Error"
        }
    }

    function Rehighlight-ExistingText {
        try {
            if (-not $richTextBox -or $richTextBox.IsDisposed) { return }
            $currentText = $richTextBox.Text
            $richTextBox.Clear()
            $global:AppConfig.ErrorCount = 0
            $global:AppConfig.WarningCount = 0
            $global:AppConfig.SuccessCount = 0
            $lines = $currentText -split "`r?`n"
            Append-LogLines -lines $lines
        } catch {
            Write-DebugLog -Message "Error rehighlighting text: $_" -Level "Error"
        }
    }

    function Load-LogFile {
        param ([string]$filePath)
        try {
            $fileInfo = Get-Item $filePath -ErrorAction Stop
            if ($fileInfo.Length -gt 100MB) {
                $result = [System.Windows.Forms.MessageBox]::Show("The log file is large ($([math]::Round($fileInfo.Length / 1MB, 2)) MB). Loading may be slow. Continue?", "Large File Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                if ($result -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            }
            if ($filePath.ToLower().EndsWith(".evtx")) {
                $global:AppConfig.IsEvtx = $true
                try {
                    $events = Get-WinEvent -Path $filePath -MaxEvents 1000 -ErrorAction Stop
                    $output = @()
                    foreach ($event in $events) {
                        if (-not $event) { continue }
                        $time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                        $id = $event.Id
                        $level = $event.LevelDisplayName
                        $message = $event.Message
                        $output += "$time  [$level] (ID: $id) - $message"
                    }
                    if ($richTextBox.InvokeRequired) {
                        $richTextBox.Invoke([System.Action]{ $richTextBox.Text = $output -join "`n" })
                    } else {
                        $richTextBox.Text = $output -join "`n"
                    }
                } catch {
                    Write-DebugLog -Message "Error reading EVTX file: $_" -Level "Error"
                }
            } else {
                $global:AppConfig.IsEvtx = $false
                $result = Read-LogFileContent -filePath $filePath -startPos 0
                if ($result) {
                    $global:AppConfig.LastPos = $result.NewPosition
                    if ($richTextBox.InvokeRequired) {
                        $richTextBox.Invoke([System.Action]{
                            $richTextBox.Clear()
                            $global:AppConfig.ErrorCount = 0
                            $global:AppConfig.WarningCount = 0
                            $global:AppConfig.SuccessCount = 0
                            Append-LogLines -lines $result.Content
                            $richTextBox.SelectionStart = $richTextBox.TextLength
                            $richTextBox.ScrollToCaret()
                        })
                    } else {
                        $richTextBox.Clear()
                        $global:AppConfig.ErrorCount = 0
                        $global:AppConfig.WarningCount = 0
                        $global:AppConfig.SuccessCount = 0
                        Append-LogLines -lines $result.Content
                        $richTextBox.SelectionStart = $richTextBox.TextLength
                        $richTextBox.ScrollToCaret()
                    }
                }
            }
            Update-LogFileLabel
        } catch {
            Write-DebugLog -Message "Error loading log file '$filePath': $_" -Level "Error"
            [System.Windows.Forms.MessageBox]::Show("Failed to load log file: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    function Update-Log {
        try {
            if ($global:AppConfig.IsEvtx) {
                try {
                    $events = Get-WinEvent -Path $global:AppConfig.LogFilePath -MaxEvents 1000 -ErrorAction Stop
                    $output = @()
                    foreach ($event in $events) {
                        if (-not $event) { continue }
                        $time = $event.TimeCreated.ToString("yyyy-MM-dd HH:mm:ss")
                        $id = $event.Id
                        $level = $event.LevelDisplayName
                        $message = $event.Message
                        $output += "$time  [$level] (ID: $id) - $message"
                    }
                    if ($richTextBox.InvokeRequired) {
                        $richTextBox.Invoke([System.Action]{ $richTextBox.Text = $output -join "`n"; $richTextBox.ScrollToCaret() })
                    } else {
                        $richTextBox.Text = $output -join "`n"
                        $richTextBox.ScrollToCaret()
                    }
                    Update-LogFileLabel
                } catch {
                    Write-DebugLog -Message "Error refreshing EVTX file: $_" -Level "Error"
                }
                return
            }
            $fs = [System.IO.File]::Open($global:AppConfig.LogFilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $currentLength = $fs.Length
            if ($currentLength -lt $global:AppConfig.LastPos) {
                $global:AppConfig.LastPos = 0
                if ($richTextBox.InvokeRequired) {
                    $richTextBox.Invoke([System.Action]{ $richTextBox.Clear() })
                } else {
                    $richTextBox.Clear()
                }
            }
            if ($currentLength -gt $global:AppConfig.LastPos) {
                $result = Read-LogFileContent -filePath $global:AppConfig.LogFilePath -startPos $global:AppConfig.LastPos
                if ($result -and $result.Content.Count -gt 0) {
                    $global:AppConfig.LastPos = $result.NewPosition
                    if ($richTextBox.InvokeRequired) {
                        $richTextBox.Invoke([System.Action]{
                            Append-LogLines -lines $result.Content
                            $richTextBox.ScrollToCaret()
                        })
                    } else {
                        Append-LogLines -lines $result.Content
                        $richTextBox.ScrollToCaret()
                    }
                    Update-LogFileLabel
                }
            }
            $fs.Close()
        } catch {
            Write-DebugLog -Message "Error updating log: $_" -Level "Error"
        }
    }

    function Update-BESClientStatus {
        try {
            $svc = Get-Service -Name "BESClient" -ErrorAction SilentlyContinue
            $status = if ($svc) { $svc.Status } else { "Not Installed" }
            $besInfoText = "BESClient Version: $clientVersion | BESClient Service Status: $status"
            if ($besInfoLabel.InvokeRequired) {
                $besInfoLabel.Invoke([System.Action]{ $besInfoLabel.Text = $besInfoText })
            } else {
                $besInfoLabel.Text = $besInfoText
            }
        } catch {
            Write-DebugLog -Message "Error updating BESClient status: $_" -Level "Error"
            $besInfoText = "BESClient Version: $clientVersion | BESClient Service Status: Error"
            if ($besInfoLabel.InvokeRequired) {
                $besInfoLabel.Invoke([System.Action]{ $besInfoLabel.Text = $besInfoText })
            } else {
                $besInfoLabel.Text = $besInfoText
            }
        }
    }

    function Load-InstalledApps {
        try {
            $apps = @()
            $regPaths = @(
                "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            )
            foreach ($path in $regPaths) {
                try {
                    if (Test-Path $path -ErrorAction Stop) {
                        $keys = Get-ChildItem $path -ErrorAction Stop
                        foreach ($key in $keys) {
                            try {
                                $app = Get-ItemProperty $key.PSPath -ErrorAction Stop
                                if ($app -and $app.PSObject.Properties.Name -contains "DisplayName") {
                                    $apps += [PSCustomObject]@{
                                        Name = $app.DisplayName
                                        Version = $app.DisplayVersion
                                        Publisher = $app.Publisher
                                    }
                                }
                            } catch {
                                Write-DebugLog -Message "Error processing registry key '$($key.PSPath)': $_" -Level "Error"
                            }
                        }
                    }
                } catch {
                    Write-DebugLog -Message "Error accessing registry path '$path': $_" -Level "Error"
                }
            }
            Populate-ListView -listView $appsListView -items ($apps | Sort-Object Name) -properties @("Name", "Version", "Publisher")
        } catch {
            Write-DebugLog -Message "Error loading installed apps: $_" -Level "Error"
        }
    }

    function Load-Processes {
        try {
            $procs = @()
            $processes = Get-Process -ErrorAction Stop | Sort-Object -Property ProcessName
            foreach ($proc in $processes) {
                if (-not $proc) { continue }
                $memMB = [math]::Round($proc.WorkingSet64 / 1MB, 2)
                $cpuSec = try { [math]::Round($proc.TotalProcessorTime.TotalSeconds, 2) } catch { "N/A" }
                $procs += [PSCustomObject]@{
                    ProcessName = $proc.ProcessName
                    Id = $proc.Id.ToString()
                    MemoryMB = $memMB.ToString()
                    CPUSec = $cpuSec.ToString()
                }
            }
            Populate-ListView -listView $procListView -items $procs -properties @("ProcessName", "Id", "MemoryMB", "CPUSec")
        } catch {
            Write-DebugLog -Message "Error loading processes: $_" -Level "Error"
        }
    }

    $criticalProcesses = @("svchost", "csrss", "winlogon", "smss")
    function Kill-Process {
        param ([int]$procId, [string]$procName)
        try {
            if ($criticalProcesses -contains $procName.ToLower()) {
                [System.Windows.Forms.MessageBox]::Show("Killing '$procName' is not allowed as it is a critical system process.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            $confirm = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to kill process '$procName' (ID: $procId)?", "Confirm Kill", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
                Stop-Process -Id $procId -Force -ErrorAction Stop
                [System.Windows.Forms.MessageBox]::Show("Process '$procName' (ID: $procId) has been killed.", "Kill Process", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                Load-Processes
            }
        } catch {
            Write-DebugLog -Message "Error killing process '$procName' (ID: $procId): $_" -Level "Error"
            [System.Windows.Forms.MessageBox]::Show("Failed to kill process '$procName': $_", "Kill Process Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }

    function Show-ManageHighlightsDialog {
        try {
            $dialog = New-Object System.Windows.Forms.Form
            $dialog.Text = "Manage Custom Highlight Rules"
            $dialog.Size = New-Object System.Drawing.Size(400, 350)
            $dialog.StartPosition = "CenterParent"

            $rulesListBox = New-Object System.Windows.Forms.ListBox
            $rulesListBox.Location = New-Object System.Drawing.Point(10, 10)
            $rulesListBox.Size = New-Object System.Drawing.Size(360, 100)
            foreach ($rule in $global:AppConfig.CustomHighlightRules) {
                if ($rule) {
                    $null = $rulesListBox.Items.Add("$($rule.Keyword) : $($rule.Color.Name)") # Suppress output
                }
            }
            $dialog.Controls.Add($rulesListBox)

            $keywordLabel = New-Object System.Windows.Forms.Label
            $keywordLabel.Text = "Keyword:"
            $keywordLabel.Location = New-Object System.Drawing.Point(10, 120)
            $keywordLabel.AutoSize = $true
            $dialog.Controls.Add($keywordLabel)

            $keywordTextBox = New-Object System.Windows.Forms.TextBox
            $keywordTextBox.Location = New-Object System.Drawing.Point(80, 117)
            $keywordTextBox.Width = 150
            $dialog.Controls.Add($keywordTextBox)

            $chooseColorButton = New-Object System.Windows.Forms.Button
            $chooseColorButton.Text = "Choose Color"
            $chooseColorButton.Location = New-Object System.Drawing.Point(240, 115)
            $chooseColorButton.AutoSize = $true
            $dialog.Controls.Add($chooseColorButton)

            $colorLabel = New-Object System.Windows.Forms.Label
            $colorLabel.Text = "No color selected"
            $colorLabel.Location = New-Object System.Drawing.Point(10, 150)
            $colorLabel.AutoSize = $true
            $dialog.Controls.Add($colorLabel)

            $previewLabel = New-Object System.Windows.Forms.Label
            $previewLabel.Text = "Sample Text"
            $previewLabel.Location = New-Object System.Drawing.Point(10, 170)
            $previewLabel.Size = New-Object System.Drawing.Size(200, 20)
            $previewLabel.BackColor = [System.Drawing.Color]::White
            $dialog.Controls.Add($previewLabel)

            $selectedColor = [ref]([System.Drawing.Color]::Empty)

            $chooseColorButton.Add_Click({
                try {
                    $colorDialog = New-Object System.Windows.Forms.ColorDialog
                    if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                        $selectedColor.Value = $colorDialog.Color
                        $colorLabel.Text = "Selected: " + $selectedColor.Value.Name
                        $previewLabel.BackColor = $selectedColor.Value
                        $previewLabel.ForeColor = if ($selectedColor.Value.GetBrightness() -lt 0.5) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
                    }
                } catch {
                    Write-DebugLog -Message "Error in chooseColorButton click: $_" -Level "Error"
                }
            })

            $addRuleButton = New-Object System.Windows.Forms.Button
            $addRuleButton.Text = "Add Rule"
            $addRuleButton.Location = New-Object System.Drawing.Point(10, 200)
            $addRuleButton.AutoSize = $true
            $dialog.Controls.Add($addRuleButton)

            $addRuleButton.Add_Click({
                try {
                    if (-not [string]::IsNullOrEmpty($keywordTextBox.Text) -and (-not $selectedColor.Value.IsEmpty)) {
                        $rule = [PSCustomObject]@{
                            Keyword = $keywordTextBox.Text
                            Color = $selectedColor.Value
                        }
                        $global:AppConfig.CustomHighlightRules += $rule
                        $null = $rulesListBox.Items.Add("$($rule.Keyword) : $($rule.Color.Name)") # Suppress output
                        $keywordTextBox.Clear()
                        $selectedColor.Value = [System.Drawing.Color]::Empty
                        $colorLabel.Text = "No color selected"
                        $previewLabel.BackColor = [System.Drawing.Color]::White
                        $previewLabel.ForeColor = [System.Drawing.Color]::Black
                        Rehighlight-ExistingText
                    } else {
                        [System.Windows.Forms.MessageBox]::Show("Please enter a keyword and select a color.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    }
                } catch {
                    Write-DebugLog -Message "Error in addRuleButton click: $_" -Level "Error"
                }
            })

            $removeRuleButton = New-Object System.Windows.Forms.Button
            $removeRuleButton.Text = "Remove Selected Rule"
            $removeRuleButton.Location = New-Object System.Drawing.Point(120, 200)
            $removeRuleButton.AutoSize = $true
            $dialog.Controls.Add($removeRuleButton)

            $removeRuleButton.Add_Click({
                try {
                    if ($rulesListBox.SelectedIndex -ge 0) {
                        $index = $rulesListBox.SelectedIndex
                        $global:AppConfig.CustomHighlightRules = $global:AppConfig.CustomHighlightRules | Where-Object { $global:AppConfig.CustomHighlightRules.IndexOf($_) -ne $index }
                        $null = $rulesListBox.Items.RemoveAt($index) # Suppress output
                        Rehighlight-ExistingText
                    }
                } catch {
                    Write-DebugLog -Message "Error in removeRuleButton click: $_" -Level "Error"
                }
            })

            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "OK"
            $okButton.Location = New-Object System.Drawing.Point(10, 230)
            $okButton.AutoSize = $true
            $dialog.Controls.Add($okButton)

            $okButton.Add_Click({
                try {
                    $dialog.Close()
                } catch {
                    Write-DebugLog -Message "Error in okButton click: $_" -Level "Error"
                }
            })

            $dialog.ShowDialog() | Out-Null
            $dialog.Dispose()
        } catch {
            Write-DebugLog -Message "Error in Show-ManageHighlightsDialog: $_" -Level "Error"
        }
    }

    ###############################################################
    # Section 4: Event Handlers
    ###############################################################

    # Fix: Validate timer initialization after form creation
    $timer = New-Object System.Windows.Forms.Timer
    if ($timer -is [System.Windows.Forms.Timer]) {
        $timer.Interval = $refreshNumeric.Value
        $timer.Add_Tick({
            try {
                $fileInfo = Get-Item $global:AppConfig.LogFilePath -ErrorAction SilentlyContinue
                if ($fileInfo -and $fileInfo.LastWriteTime -gt $global:AppConfig.LastModified) {
                    Update-Log
                    $global:AppConfig.LastModified = $fileInfo.LastWriteTime
                    $timer.Interval = [math]::Min($refreshNumeric.Value, 1000)
                } else {
                    $timer.Interval = [math]::Max($refreshNumeric.Value, 5000)
                }
                Update-BESClientStatus
            } catch {
                Write-DebugLog -Message "Error in timer tick: $_" -Level "Error"
            }
        })
        $timer.Start()
        Write-DebugLog -Message "Timer initialized and started successfully" -Level "Info"
    } else {
        Write-DebugLog -Message "Failed to initialize timer: Not a valid System.Windows.Forms.Timer object" -Level "Error"
    }

    $chooseFileButton.Add_Click({
        try {
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.InitialDirectory = Split-Path $global:AppConfig.LogFilePath
            $openFileDialog.Filter = "Log Files (*.log;*.txt;*.evtx)|*.log;*.txt;*.evtx|All Files (*.*)|*.*"
            if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                if (Validate-FilePath $openFileDialog.FileName) {
                    $global:AppConfig.LogFilePath = $openFileDialog.FileName
                    try {
                        $global:AppConfig.LastModified = (Get-Item $global:AppConfig.LogFilePath -ErrorAction Stop).LastWriteTime
                    } catch {
                        Write-DebugLog -Message "Error getting last modified time for new log file: $_" -Level "Error"
                        $global:AppConfig.LastModified = [DateTime]::Now
                    }
                    Update-LogFileLabel
                    if ($form.InvokeRequired) {
                        $form.Invoke([System.Action]{ $form.Text = "BigFix Log Viewer - $(Split-Path $global:AppConfig.LogFilePath -Leaf)" })
                    } else {
                        $form.Text = "BigFix Log Viewer - $(Split-Path $global:AppConfig.LogFilePath -Leaf)"
                    }
                    $global:AppConfig.LastPos = 0
                    if ($richTextBox.InvokeRequired) {
                        $richTextBox.Invoke([System.Action]{
                            $richTextBox.Clear()
                            $global:AppConfig.ErrorCount = 0
                            $global:AppConfig.WarningCount = 0
                            $global:AppConfig.SuccessCount = 0
                        })
                    } else {
                        $richTextBox.Clear()
                        $global:AppConfig.ErrorCount = 0
                        $global:AppConfig.WarningCount = 0
                        $global:AppConfig.SuccessCount = 0
                    }
                    Load-LogFile -filePath $global:AppConfig.LogFilePath
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Invalid or inaccessible file path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                }
            }
        } catch {
            Write-DebugLog -Message "Error in chooseFileButton click: $_" -Level "Error"
        }
    })

    $pauseResumeButton.Add_Click({
        try {
            if ($timer -and $timer.Enabled) {
                $timer.Stop()
                if ($pauseResumeButton.InvokeRequired) {
                    $pauseResumeButton.Invoke([System.Action]{ $pauseResumeButton.Text = "Resume" })
                } else {
                    $pauseResumeButton.Text = "Resume"
                }
            } else {
                if ($timer -is [System.Windows.Forms.Timer]) {
                    $timer.Start()
                    if ($pauseResumeButton.InvokeRequired) {
                        $pauseResumeButton.Invoke([System.Action]{ $pauseResumeButton.Text = "Pause" })
                    } else {
                        $pauseResumeButton.Text = "Pause"
                    }
                } else {
                    Write-DebugLog -Message "Cannot resume timer: Not a valid System.Windows.Forms.Timer object" -Level "Error"
                }
            }
        } catch {
            Write-DebugLog -Message "Error in pauseResumeButton click: $_" -Level "Error"
        }
    })

    $exportButton.Add_Click({
        try {
            $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                if ($richTextBox.SelectedText) {
                    $richTextBox.SelectedText | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
                } else {
                    $richTextBox.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
                }
                [System.Windows.Forms.MessageBox]::Show("Log exported successfully.", "Export", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
        } catch {
            Write-DebugLog -Message "Error in exportButton click: $_" -Level "Error"
            [System.Windows.Forms.MessageBox]::Show("Failed to export log: $_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $clearButton.Add_Click({
        try {
            if ($richTextBox.InvokeRequired) {
                $richTextBox.Invoke([System.Action]{
                    $richTextBox.Clear()
                    $global:AppConfig.ErrorCount = 0
                    $global:AppConfig.WarningCount = 0
                    $global:AppConfig.SuccessCount = 0
                    Update-Stats
                })
            } else {
                $richTextBox.Clear()
                $global:AppConfig.ErrorCount = 0
                $global:AppConfig.WarningCount = 0
                $global:AppConfig.SuccessCount = 0
                Update-Stats
            }
        } catch {
            Write-DebugLog -Message "Error in clearButton click: $_" -Level "Error"
        }
    })

    $refreshNumeric.Add_ValueChanged({
        try {
            if ($timer -is [System.Windows.Forms.Timer]) {
                $timer.Interval = $refreshNumeric.Value
            } else {
                Write-DebugLog -Message "Cannot set timer interval: Not a valid System.Windows.Forms.Timer object" -Level "Error"
            }
        } catch {
            Write-DebugLog -Message "Error in refreshNumeric ValueChanged: $_" -Level "Error"
        }
    })

    $findNextButton.Add_Click({
        try {
            $searchTerm = $searchTextBox.Text
            if ([string]::IsNullOrEmpty($searchTerm)) { return }
            $startPos = $richTextBox.SelectionStart + $richTextBox.SelectionLength
            $findOptions = if ($caseSensitiveCheckBox.Checked) { [System.Windows.Forms.RichTextBoxFinds]::None } else { [System.Windows.Forms.RichTextBoxFinds]::MatchCase -bxor [System.Windows.Forms.RichTextBoxFinds]::MatchCase }
            $index = $richTextBox.Find($searchTerm, $startPos, $findOptions)
            if ($index -eq -1) {
                [System.Windows.Forms.MessageBox]::Show("No further occurrences found.", "Search", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            } else {
                $richTextBox.Focus()
            }
        } catch {
            Write-DebugLog -Message "Error in findNextButton click: $_" -Level "Error"
        }
    })

    $filterComboBox.Add_SelectedIndexChanged({
        try {
            if ($filterTextBox.InvokeRequired) {
                $filterTextBox.Invoke([System.Action]{ $filterTextBox.Visible = ($filterComboBox.SelectedItem -eq "Custom") })
            } else {
                $filterTextBox.Visible = ($filterComboBox.SelectedItem -eq "Custom")
            }
            Rehighlight-ExistingText
        } catch {
            Write-DebugLog -Message "Error in filterComboBox SelectedIndexChanged: $_" -Level "Error"
        }
    })

    $openEventViewerButton.Add_Click({
        try {
            Start-Process "eventvwr.exe" -ErrorAction Stop
        } catch {
            Write-DebugLog -Message "Error in openEventViewerButton click: $_" -Level "Error"
            [System.Windows.Forms.MessageBox]::Show("Failed to open Event Viewer: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    # Fix: Added indicator for BESClient service restart
    $restartBESClientButton.Add_Click({
        try {
            # Update button text to show restarting status
            if ($restartBESClientButton.InvokeRequired) {
                $restartBESClientButton.Invoke([System.Action]{ $restartBESClientButton.Text = "Restarting..." })
            } else {
                $restartBESClientButton.Text = "Restarting..."
            }
            Write-DebugLog -Message "Initiating BESClient service restart" -Level "Info"
            Restart-Service -Name "BESClient" -Force -ErrorAction Stop
            # Restore button text on success
            if ($restartBESClientButton.InvokeRequired) {
                $restartBESClientButton.Invoke([System.Action]{ $restartBESClientButton.Text = "Restart BESClient" })
            } else {
                $restartBESClientButton.Text = "Restart BESClient"
            }
            [System.Windows.Forms.MessageBox]::Show("BESClient service restarted successfully.", "Service Restart", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            Write-DebugLog -Message "BESClient service restarted successfully" -Level "Info"
        } catch {
            # Restore button text on error
            if ($restartBESClientButton.InvokeRequired) {
                $restartBESClientButton.Invoke([System.Action]{ $restartBESClientButton.Text = "Restart BESClient" })
            } else {
                $restartBESClientButton.Text = "Restart BESClient"
            }
            Write-DebugLog -Message "Error in restartBESClientButton click: $_" -Level "Error"
            [System.Windows.Forms.MessageBox]::Show("Failed to restart BESClient service: $_", "Service Restart Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $manageHighlightsButton.Add_Click({
        try {
            Show-ManageHighlightsDialog
        } catch {
            Write-DebugLog -Message "Error in manageHighlightsButton click: $_" -Level "Error"
        }
    })

    $appsSearchButton.Add_Click({
        try {
            $searchText = $appsSearchTextBox.Text
            if (-not [string]::IsNullOrEmpty($searchText)) {
                foreach ($item in $appsListView.Items) {
                    if ($item -and -not $item.IsDisposed) {
                        $item.BackColor = [System.Drawing.Color]::White
                    }
                }
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
        } catch {
            Write-DebugLog -Message "Error in appsSearchButton click: $_" -Level "Error"
        }
    })

    $appsRefreshButton.Add_Click({
        try {
            Load-InstalledApps
        } catch {
            Write-DebugLog -Message "Error in appsRefreshButton click: $_" -Level "Error"
        }
    })

    $refreshProcessesButton.Add_Click({
        try {
            Load-Processes
        } catch {
            Write-DebugLog -Message "Error in refreshProcessesButton click: $_" -Level "Error"
        }
    })

    $killProcessButton.Add_Click({
        try {
            if ($procListView.SelectedItems.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select a process to kill.", "Kill Process", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            } else {
                $selectedItem = $procListView.SelectedItems[0]
                $procId = [int]$selectedItem.SubItems[1].Text
                $procName = $selectedItem.Text
                Kill-Process -procId $procId -procName $procName
            }
        } catch {
            Write-DebugLog -Message "Error in killProcessButton click: $_" -Level "Error"
        }
    })

    $killMenuItem.Add_Click({
        try {
            if ($procListView.SelectedItems.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No process is selected.", "Kill Process", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            } else {
                $selectedItem = $procListView.SelectedItems[0]
                $procId = [int]$selectedItem.SubItems[1].Text
                $procName = $selectedItem.Text
                Kill-Process -procId $procId -procName $procName
            }
        } catch {
            Write-DebugLog -Message "Error in killMenuItem click: $_" -Level "Error"
        }
    })

    ###############################################################
    # Section 5: Initialize and Run the Application
    ###############################################################

    try {
        Load-InstalledApps
        Load-Processes
        Load-LogFile -filePath $global:AppConfig.LogFilePath
        $tabControl.SelectedTab = $logViewerTab
    } catch {
        Write-DebugLog -Message "Initialization error: $_" -Level "Error"
        [System.Windows.Forms.MessageBox]::Show("Failed to initialize application: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }

    $form.Add_FormClosing({
        try {
            $newSettings = @{
                RefreshInterval = $refreshNumeric.Value
                WindowSize = @{ Width = $form.Width; Height = $form.Height }
                WindowLocation = @{ X = $form.Location.X; Y = $form.Location.Y }
                LastLogFile = $global:AppConfig.LogFilePath
                CustomHighlightRules = @()
            }
            foreach ($rule in $global:AppConfig.CustomHighlightRules) {
                if ($rule) {
                    $newSettings.CustomHighlightRules += @{
                        Keyword = $rule.Keyword
                        ColorName = $rule.Color.Name
                    }
                }
            }
            $newSettings | ConvertTo-Json -Depth 5 | Out-File -FilePath $global:AppConfig.ConfigFile -Encoding UTF8 -ErrorAction Stop
            Write-DebugLog -Message "Settings saved to $($global:AppConfig.ConfigFile)" -Level "Info"
        } catch {
            Write-DebugLog -Message "Error saving settings on form close: $_" -Level "Error"
        }
        try {
            if ($timer -is [System.Windows.Forms.Timer] -and $timer.Enabled) {
                $timer.Stop()
                $timer.Dispose()
            }
            if ($richTextBox -and -not $richTextBox.IsDisposed) { $richTextBox.Dispose() }
            if ($tabControl -and -not $tabControl.IsDisposed) { $tabControl.Dispose() }
            if ($form -and -not $form.IsDisposed) { $form.Dispose() }
            Write-DebugLog -Message "Resources disposed successfully" -Level "Info"
        } catch {
            Write-DebugLog -Message "Error disposing resources on form close: $_" -Level "Error"
        }
    })

    Scale-Font -control $form -scaleFactor 1.0

    try {
        [System.Windows.Forms.Application]::Run($form)
    } catch {
        Write-DebugLog -Message "Unhandled exception in application run: $_" -Level "Error"
        [System.Windows.Forms.MessageBox]::Show("An unexpected error occurred: $_`nSee $debugLogFile for details.", "Critical Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
} catch {
    Write-DebugLog -Message "Critical script error: $_" -Level "Error"
    Write-Host "Critical error: $_`nSee $debugLogFile for details."
    exit
}
