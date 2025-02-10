# Load required .NET assemblies.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

###############################################################
# Section 1: Setup – Configuration, Log File, System Info, and Global Variables
###############################################################

# Set configuration file path.
$configFile = Join-Path $PSScriptRoot "BigFixLogViewerSettings.json"

# Define default values.
$defaultLogDir = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
$defaultRefreshInterval = 1000
$defaultWindowSize = @{ Width = 1000; Height = 750 }
$defaultWindowLocation = @{ X = 100; Y = 100 }
$defaultCustomHighlightRules = @()

# Load saved settings if available.
$savedSettings = $null
if (Test-Path $configFile) {
    try {
        $savedSettings = Get-Content $configFile -Raw | ConvertFrom-Json
    }
    catch {
        Write-Host "Error reading config file: $_"
    }
}

# Determine the log file to open.
if ($savedSettings -and $savedSettings.LastLogFile -and (Test-Path $savedSettings.LastLogFile)) {
    $global:logFilePath = $savedSettings.LastLogFile
} else {
    if (-not (Test-Path $defaultLogDir)) {
        [System.Windows.Forms.MessageBox]::Show("Log directory not found:`n$defaultLogDir", "Error", 'OK', 'Error')
        exit
    }
    $latestLogFile = Get-ChildItem -Path $defaultLogDir -File |
                     Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latestLogFile) {
        [System.Windows.Forms.MessageBox]::Show("No log files found in:`n$defaultLogDir", "Error", 'OK', 'Error')
        exit
    }
    $global:logFilePath = $latestLogFile.FullName
}

# Load system information.
$computerName = $env:COMPUTERNAME
try {
    # Exclude both 127.0.0.1 and link-local addresses (169.254.x.x).
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
# Correct BESClient.exe path (note the space between "BES" and "Client").
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
                  "Relay Server: $relayServer`n" +
                  "BESClient Version: $clientVersion"

# Global variables.
$global:lastPos = 0
$global:errorCount   = 0
$global:warningCount = 0
$global:successCount = 0
$global:IsEvtx = $false  # Flag for EVTX file.

# Load custom highlight rules from saved settings.
if ($savedSettings -and $savedSettings.CustomHighlightRules) {
    $global:customHighlightRules = foreach ($rule in $savedSettings.CustomHighlightRules) {
        [PSCustomObject]@{
            Keyword = $rule.Keyword
            Color   = [System.Drawing.Color]::FromName($rule.ColorName)
        }
    }
} else {
    $global:customHighlightRules = $defaultCustomHighlightRules
}

###############################################################
# Section 2: Build the GUI
###############################################################

# Create main form.
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Log Viewer - $(Split-Path $global:logFilePath -Leaf)"
if ($savedSettings -and $savedSettings.WindowSize) {
    $form.Width = $savedSettings.WindowSize.Width
    $form.Height = $savedSettings.WindowSize.Height
} else {
    $form.Width = $defaultWindowSize.Width
    $form.Height = $defaultWindowSize.Height
}
if ($savedSettings -and $savedSettings.WindowLocation) {
    $form.Location = New-Object System.Drawing.Point($savedSettings.WindowLocation.X, $savedSettings.WindowLocation.Y)
} else {
    $form.Location = New-Object System.Drawing.Point($defaultWindowLocation.X, $defaultWindowLocation.Y)
}
$form.StartPosition = 'Manual'

# -- System Info Panel (Top) --
$systemInfoPanel = New-Object System.Windows.Forms.Panel
$systemInfoPanel.Dock = 'Top'
$systemInfoPanel.Height = 120
# Set background to black.
$systemInfoPanel.BackColor = [System.Drawing.Color]::Black
$systemInfoPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$form.Controls.Add($systemInfoPanel)

# Use a FlowLayoutPanel inside System Info for multiple lines.
$flowPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$flowPanel.Dock = 'Fill'
$flowPanel.FlowDirection = 'TopDown'
$systemInfoPanel.Controls.Add($flowPanel)

$systemInfoLabel = New-Object System.Windows.Forms.Label
$systemInfoLabel.AutoSize = $true
$systemInfoLabel.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
# Set text color to cyan.
$systemInfoLabel.ForeColor = [System.Drawing.Color]::Cyan
$systemInfoLabel.Text = $systemInfoText
$flowPanel.Controls.Add($systemInfoLabel)

# New label for BESClient service status.
$besClientStatusLabel = New-Object System.Windows.Forms.Label
$besClientStatusLabel.AutoSize = $true
$besClientStatusLabel.Font = New-Object System.Drawing.Font("Consolas", 10, [System.Drawing.FontStyle]::Bold)
$besClientStatusLabel.ForeColor = [System.Drawing.Color]::Cyan
$besClientStatusLabel.Text = "BESClient Service Status: Unknown"
$flowPanel.Controls.Add($besClientStatusLabel)

# -- Log File Panel (Below System Info) --
$logFilePanel = New-Object System.Windows.Forms.Panel
$logFilePanel.Dock = 'Top'
$logFilePanel.Height = 30
$form.Controls.Add($logFilePanel)

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
$logFilePanel.Controls.Add($logFileLabel)

# -- Control Panel (Search, Pause/Resume, Export, Clear, Choose Log, Manage Highlights, Open Event Viewer, Restart BESClient, Refresh Interval, Stats) --
$controlPanel = New-Object System.Windows.Forms.Panel
$controlPanel.Dock = 'Top'
$controlPanel.Height = 80
$controlPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($controlPanel)

$searchLabel = New-Object System.Windows.Forms.Label
$searchLabel.Text = "Search:"
$searchLabel.Location = New-Object System.Drawing.Point(10, 10)
$searchLabel.AutoSize = $true
$controlPanel.Controls.Add($searchLabel)

$searchTextBox = New-Object System.Windows.Forms.TextBox
$searchTextBox.Location = New-Object System.Drawing.Point(70, 7)
$searchTextBox.Width = 150
$controlPanel.Controls.Add($searchTextBox)

$findNextButton = New-Object System.Windows.Forms.Button
$findNextButton.Text = "Find Next"
$findNextButton.Location = New-Object System.Drawing.Point(230, 5)
$findNextButton.AutoSize = $true
$controlPanel.Controls.Add($findNextButton)

$pauseResumeButton = New-Object System.Windows.Forms.Button
$pauseResumeButton.Text = "Pause"
$pauseResumeButton.Location = New-Object System.Drawing.Point(320, 5)
$pauseResumeButton.AutoSize = $true
$controlPanel.Controls.Add($pauseResumeButton)

$exportButton = New-Object System.Windows.Forms.Button
$exportButton.Text = "Export Log"
$exportButton.Location = New-Object System.Drawing.Point(400, 5)
$exportButton.AutoSize = $true
$controlPanel.Controls.Add($exportButton)

$clearButton = New-Object System.Windows.Forms.Button
$clearButton.Text = "Clear Log"
$clearButton.Location = New-Object System.Drawing.Point(490, 5)
$clearButton.AutoSize = $true
$controlPanel.Controls.Add($clearButton)

$chooseFileButton = New-Object System.Windows.Forms.Button
$chooseFileButton.Text = "Choose Log File"
$chooseFileButton.AutoSize = $true
$chooseFileButton.Location = New-Object System.Drawing.Point(580, 5)
$controlPanel.Controls.Add($chooseFileButton)

$manageHighlightsButton = New-Object System.Windows.Forms.Button
$manageHighlightsButton.Text = "Manage Highlights"
$manageHighlightsButton.AutoSize = $true
$manageHighlightsButton.Location = New-Object System.Drawing.Point(680, 5)
$controlPanel.Controls.Add($manageHighlightsButton)

$openEventViewerButton = New-Object System.Windows.Forms.Button
$openEventViewerButton.Text = "Open Event Viewer"
$openEventViewerButton.AutoSize = $true
$openEventViewerButton.Location = New-Object System.Drawing.Point(800, 5)
$controlPanel.Controls.Add($openEventViewerButton)

$restartBESClientButton = New-Object System.Windows.Forms.Button
$restartBESClientButton.Text = "Restart BESClient"
$restartBESClientButton.AutoSize = $true
$restartBESClientButton.Location = New-Object System.Drawing.Point(920, 5)
$restartBESClientButton.BackColor = [System.Drawing.Color]::LightGreen
$controlPanel.Controls.Add($restartBESClientButton)

$refreshLabel = New-Object System.Windows.Forms.Label
$refreshLabel.Text = "Refresh Interval (ms):"
$refreshLabel.Location = New-Object System.Drawing.Point(10, 40)
$refreshLabel.AutoSize = $true
$controlPanel.Controls.Add($refreshLabel)

$refreshNumeric = New-Object System.Windows.Forms.NumericUpDown
$refreshNumeric.Location = New-Object System.Drawing.Point(140, 38)
$refreshNumeric.Minimum = 100
$refreshNumeric.Maximum = 5000
if ($savedSettings -and $savedSettings.RefreshInterval) {
    $refreshNumeric.Value = $savedSettings.RefreshInterval
} else {
    $refreshNumeric.Value = $defaultRefreshInterval
}
$controlPanel.Controls.Add($refreshNumeric)

$statsLabel = New-Object System.Windows.Forms.Label
$statsLabel.Text = "Errors: 0, Warnings: 0, Successes: 0"
$statsLabel.Location = New-Object System.Drawing.Point(320, 40)
$statsLabel.AutoSize = $true
$controlPanel.Controls.Add($statsLabel)

# -- Log Content Panel --
$richTextBox = New-Object System.Windows.Forms.RichTextBox
$richTextBox.Multiline = $true
$richTextBox.ReadOnly = $true
$richTextBox.ScrollBars = 'Vertical'
$richTextBox.Dock = 'Fill'
$richTextBox.Font = New-Object System.Drawing.Font("Consolas",10)
$richTextBox.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($richTextBox)

# Ensure proper ordering of panels.
$form.Controls.SetChildIndex($systemInfoPanel, 0)
$form.Controls.SetChildIndex($logFilePanel, 1)
$form.Controls.SetChildIndex($controlPanel, 2)

###############################################################
# Section 3: Log Tailing, Rehighlight Function, and Helper Functions
###############################################################

# Helper: Append a line with color highlighting.
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

# Helper: Update statistics label.
function Update-Stats {
    $statsLabel.Text = "Errors: $global:errorCount, Warnings: $global:warningCount, Successes: $global:successCount"
}

# Function: Rehighlight all existing text (called when custom rules change).
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

# Function: Load the initial log file content.
function Load-LogFile {
    param (
        [string]$filePath
    )
    if ($filePath.ToLower().EndsWith(".evtx")) {
        $global:IsEvtx = $true
        try {
            # For EVTX files, use Get-WinEvent (limit to the most recent 1000 events)
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
            Write-Host "Error reading EVTX file: $_"
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
            Write-Host "Error reading log content: $_"
        }
    }
    Update-LogFileLabel
}

# Function: Update (tail) the log file.
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
            Write-Host "Error refreshing EVTX file: $_"
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
        Write-Host "Error reading log file: $_"
    }
}

###############################################################
# Section 4: Set Up Timer and Control Event Handlers
###############################################################

# Function: Update BESClient service status.
function Update-BESClientStatus {
    try {
        $svc = Get-Service -Name "BESClient" -ErrorAction SilentlyContinue
        if ($svc) {
            $status = $svc.Status
        } else {
            $status = "Not Installed"
        }
        $besClientStatusLabel.Invoke([System.Action]{
            $besClientStatusLabel.Text = "BESClient Service Status: $status"
        })
    }
    catch {
        $besClientStatusLabel.Invoke([System.Action]{
            $besClientStatusLabel.Text = "BESClient Service Status: Error"
        })
    }
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = $refreshNumeric.Value
# On each tick, update the log file and the BESClient service status.
$timer.Add_Tick({
    Update-Log
    Update-BESClientStatus
})
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

$refreshNumeric.Add_ValueChanged({
    $timer.Interval = $refreshNumeric.Value
})

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

$openEventViewerButton.Add_Click({
    Start-Process "eventvwr.exe"
})

$restartBESClientButton.Add_Click({
    try {
        Restart-Service -Name "BESClient" -Force -ErrorAction Stop
        [System.Windows.Forms.MessageBox]::Show("BESClient service restarted successfully.","Service Restart",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to restart BESClient service: $_","Service Restart Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

# Function: Show Manage Highlights Dialog.
function Show-ManageHighlightsDialog {
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Manage Custom Highlight Rules"
    $dialog.Size = New-Object System.Drawing.Size(400,300)
    $dialog.StartPosition = "CenterParent"

    # ListBox for existing rules.
    $rulesListBox = New-Object System.Windows.Forms.ListBox
    $rulesListBox.Location = New-Object System.Drawing.Point(10,10)
    $rulesListBox.Size = New-Object System.Drawing.Size(360,100)
    foreach ($rule in $global:customHighlightRules) {
         $rulesListBox.Items.Add("$($rule.Keyword) : $($rule.Color.Name)")
    }
    $dialog.Controls.Add($rulesListBox)

    # Label for new rule keyword.
    $keywordLabel = New-Object System.Windows.Forms.Label
    $keywordLabel.Text = "Keyword:"
    $keywordLabel.Location = New-Object System.Drawing.Point(10, 120)
    $keywordLabel.AutoSize = $true
    $dialog.Controls.Add($keywordLabel)

    # TextBox for new rule keyword.
    $keywordTextBox = New-Object System.Windows.Forms.TextBox
    $keywordTextBox.Location = New-Object System.Drawing.Point(80, 117)
    $keywordTextBox.Width = 150
    $dialog.Controls.Add($keywordTextBox)

    # Button for choosing color.
    $chooseColorButton = New-Object System.Windows.Forms.Button
    $chooseColorButton.Text = "Choose Color"
    $chooseColorButton.Location = New-Object System.Drawing.Point(240, 115)
    $chooseColorButton.AutoSize = $true
    $dialog.Controls.Add($chooseColorButton)

    # Label to show chosen color.
    $colorLabel = New-Object System.Windows.Forms.Label
    $colorLabel.Text = "No color selected"
    $colorLabel.Location = New-Object System.Drawing.Point(10, 150)
    $colorLabel.AutoSize = $true
    $dialog.Controls.Add($colorLabel)

    # Variable to store the selected color as a [ref] variable.
    $selectedColor = [ref]([System.Drawing.Color]::Empty)

    $chooseColorButton.Add_Click({
      $colorDialog = New-Object System.Windows.Forms.ColorDialog
      if ($colorDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
         $selectedColor.Value = $colorDialog.Color
         $colorLabel.Text = "Selected: " + $selectedColor.Value.Name
      }
    })

    # Add Rule button.
    $addRuleButton = New-Object System.Windows.Forms.Button
    $addRuleButton.Text = "Add Rule"
    $addRuleButton.Location = New-Object System.Drawing.Point(10, 180)
    $addRuleButton.AutoSize = $true
    $dialog.Controls.Add($addRuleButton)

    $addRuleButton.Add_Click({
      if (-not [string]::IsNullOrEmpty($keywordTextBox.Text) -and (-not $selectedColor.Value.IsEmpty)) {
         $rule = [PSCustomObject]@{
             Keyword = $keywordTextBox.Text
             Color   = $selectedColor.Value
         }
         $global:customHighlightRules += $rule
         $rulesListBox.Items.Add("$($rule.Keyword) : $($rule.Color.Name)")
         $keywordTextBox.Clear()
         $selectedColor.Value = [System.Drawing.Color]::Empty
         $colorLabel.Text = "No color selected"
         Rehighlight-ExistingText
      } else {
         [System.Windows.Forms.MessageBox]::Show("Please enter a keyword and select a color.","Input Required",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
      }
    })

    # Remove Rule button.
    $removeRuleButton = New-Object System.Windows.Forms.Button
    $removeRuleButton.Text = "Remove Selected Rule"
    $removeRuleButton.Location = New-Object System.Drawing.Point(120, 180)
    $removeRuleButton.AutoSize = $true
    $dialog.Controls.Add($removeRuleButton)

    $removeRuleButton.Add_Click({
      if ($rulesListBox.SelectedIndex -ge 0) {
         $index = $rulesListBox.SelectedIndex
         $global:customHighlightRules = $global:customHighlightRules | Where-Object { $global:customHighlightRules.IndexOf($_) -ne $index }
         $rulesListBox.Items.RemoveAt($index)
         Rehighlight-ExistingText
      }
    })

    # OK button to close the dialog.
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(10, 220)
    $okButton.AutoSize = $true
    $dialog.Controls.Add($okButton)

    $okButton.Add_Click({
      $dialog.Close()
    })

    $dialog.ShowDialog() | Out-Null
}

$manageHighlightsButton.Add_Click({ Show-ManageHighlightsDialog })

###############################################################
# Section 5: Initialize and Run the Application
###############################################################

Load-LogFile -filePath $global:logFilePath

# Save user preferences when the form is closing.
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
    $newSettings | ConvertTo-Json -Depth 5 | Out-File -FilePath $configFile -Encoding UTF8
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
        # Cast the font size to double and multiply.
        $newSize = ([double]$control.Font.Size) * $scaleFactor
        $control.Font = New-Object System.Drawing.Font($control.Font.FontFamily, $newSize, $control.Font.Style)
    }
    foreach ($child in $control.Controls) {
        Scale-Font -control $child -scaleFactor $scaleFactor
    }
}
# Here, using a scaling factor of 1.0 so that buttons and labels retain their original sizes.
Scale-Font -control $systemInfoPanel -scaleFactor 1.0
Scale-Font -control $logFilePanel -scaleFactor 1.0
Scale-Font -control $controlPanel -scaleFactor 1.0

[System.Windows.Forms.Application]::Run($form)
$timer.Stop()
$timer.Dispose()
