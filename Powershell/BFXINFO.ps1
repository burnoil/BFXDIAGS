# Import necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create debug log file
$debugLog = "C:\Temp\BigFixDebug.txt"
if (-not (Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null }
"Starting BigFix Action Viewer at $(Get-Date)" | Out-File -FilePath $debugLog -Force

# Function to get last 100 BigFix actions from logs
function Get-LastBigFixActions {
    $logPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
    "Checking log path: $logPath" | Out-File -FilePath $debugLog -Append

    if (-not (Test-Path $logPath)) {
        "Error: Log path does not exist: $logPath" | Out-File -FilePath $debugLog -Append
        return @([PSCustomObject]@{ Timestamp = "N/A"; Description = "Log directory not found: $logPath" })
    }

    try {
        $logFiles = Get-ChildItem -Path $logPath -Filter "*.log" -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 10
        "Found $($logFiles.Count) log files" | Out-File -FilePath $debugLog -Append
    } catch {
        "Error accessing log files: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
        return @([PSCustomObject]@{ Timestamp = "N/A"; Description = "Error accessing log files: $($_.Exception.Message)" })
    }

    $actions = @()
    foreach ($file in $logFiles) {
        try {
            $content = Get-Content $file.FullName -ErrorAction Stop
            # Broaden filter to catch more action-related lines
            $actionLines = $content | Where-Object { $_ -match "Action|Command|Started|Completed|Failed|Downloading|Executing|Reported" }
            "Found $($actionLines.Count) action lines in $($file.Name)" | Out-File -FilePath $debugLog -Append
            $actions += $actionLines
        } catch {
            "Error reading $($file.Name): $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
            $actions += "Error reading $($file.Name): $($_.Exception.Message)"
        }
    }

    # Get the last 100 actions, reverse for newest first
    $actions = $actions | Select-Object -Last 100
    $actions = $actions | ForEach-Object { 
        $parts = $_ -split " ", 3
        [PSCustomObject]@{ 
            Timestamp = if ($parts.Length -ge 2) { $parts[0..1] -join " " } else { "N/A" }
            Description = if ($parts.Length -ge 3) { $parts[2] } else { $_ }
        }
    }

    if ($actions.Count -eq 0) {
        "No action-related log entries found" | Out-File -FilePath $debugLog -Append
        $actions = @([PSCustomObject]@{ Timestamp = "N/A"; Description = "No action-related log entries found." })
    }
    "Returning $($actions.Count) actions" | Out-File -FilePath $debugLog -Append
    return $actions
}

# Function to get BESClient last check-in time
function Get-BESClientLastCheckIn {
    $regPaths = @(
        "HKLM:\SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client",
        "HKLM:\SOFTWARE\BigFix\EnterpriseClient\Settings\Client"
    )
    $logPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"

    # Try registry paths
    foreach ($regPath in $regPaths) {
        "Checking registry: $regPath" | Out-File -FilePath $debugLog -Append
        try {
            $lastReportTime = Get-ItemProperty -Path $regPath -Name "LastReportTime" -ErrorAction Stop
            $unixTime = $lastReportTime.LastReportTime / 1000
            $checkInTime = (Get-Date "1970-01-01").AddSeconds($unixTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
            "Found check-in time in registry: $checkInTime" | Out-File -FilePath $debugLog -Append
            return $checkInTime
        } catch {
            "Registry error at $regPath: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
        }
    }

    # Fallback to log parsing
    "Falling back to log parsing for check-in time" | Out-File -FilePath $debugLog -Append
    try {
        $latestLog = Get-ChildItem -Path $logPath -Filter "*.log" -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latestLog) {
            "Parsing log: $($latestLog.FullName)" | Out-File -FilePath $debugLog -Append
            $content = Get-Content $latestLog.FullName -ErrorAction Stop
            $checkInLine = $content | Where-Object { $_ -match "Report posted successfully|Gather completed" } | Select-Object -Last 1
            if ($checkInLine) {
                $timestamp = ($checkInLine -split " ")[0..1] -join " "
                "Found check-in time in log: $timestamp" | Out-File -FilePath $debugLog -Append
                return $timestamp
            } else {
                "No check-in lines found in log" | Out-File -FilePath $debugLog -Append
                return "No check-in data found in logs."
            }
        } else {
            "No log files found for check-in" | Out-File -FilePath $debugLog -Append
            return "No log files found."
        }
    } catch {
        "Log parsing error: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
        return "Error parsing logs: $($_.Exception.Message)"
    }
}

# Create the GUI form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Viewer"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"

# Label for last check-in
$labelCheckIn = New-Object System.Windows.Forms.Label
$checkInTime = Get-BESClientLastCheckIn
$labelCheckIn.Text = "Last BESClient Check-In: $checkInTime"
$labelCheckIn.Location = New-Object System.Drawing.Point(10, 10)
$labelCheckIn.AutoSize = $true
$form.Controls.Add($labelCheckIn)

# DataGridView for actions
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(10, 40)
$dataGrid.Size = New-Object System.Drawing.Size(760, 500)
$dataGrid.ReadOnly = $true
$dataGrid.AutoSizeColumnsMode = "Fill"
$dataGrid.DataSource = $null  # Clear before assigning
$actions = Get-LastBigFixActions
$dataGrid.DataSource = $actions
"Assigned $($actions.Count) actions to DataGridView" | Out-File -FilePath $debugLog -Append
$form.Controls.Add($dataGrid)

# Show the form
$form.ShowDialog() | Out-Null
"Form closed" | Out-File -FilePath $debugLog -Append
