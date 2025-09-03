# Ensure debug log creation at the start
$debugLog = "C:\Temp\BigFixDebug.txt"
try {
    if (-not (Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null }
    "Starting BigFix Action Viewer at $(Get-Date)" | Out-File -FilePath $debugLog -Force
    Write-Host "Debug log created at $debugLog"
} catch {
    Write-Host "Failed to create debug log: $($_.Exception.Message)"
}

# Import necessary assemblies for GUI
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    "Loaded assemblies successfully" | Out-File -FilePath $debugLog -Append
    Write-Host "Loaded Windows Forms and Drawing assemblies"
} catch {
    "Error loading assemblies: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
    Write-Host "Error loading assemblies: $($_.Exception.Message)"
    exit
}

# Function to get last 100 BigFix actions from logs
function Get-LastBigFixActions {
    $logPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
    "Checking log path: $logPath" | Out-File -FilePath $debugLog -Append
    Write-Host "Checking log path: $logPath"

    if (-not (Test-Path $logPath)) {
        "Error: Log path does not exist: $logPath" | Out-File -FilePath $debugLog -Append
        Write-Host "Error: Log path does not exist"
        return @([PSCustomObject]@{ Timestamp = "N/A"; Description = "Log directory not found: $logPath" })
    }

    try {
        $logFiles = Get-ChildItem -Path $logPath -Filter "*.log" -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 10
        "Found $($logFiles.Count) log files" | Out-File -FilePath $debugLog -Append
        Write-Host "Found $($logFiles.Count) log files"
    } catch {
        "Error accessing log files: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
        Write-Host "Error accessing log files: $($_.Exception.Message)"
        return @([PSCustomObject]@{ Timestamp = "N/A"; Description = "Error accessing log files: $($_.Exception.Message)" })
    }

    $actions = @()
    foreach ($file in $logFiles) {
        try {
            $content = Get-Content $file.FullName -ErrorAction Stop
            # Simplified filter to catch any non-empty line for testing
            $actionLines = $content | Where-Object { $_ -match ".*" } | Select-Object -First 100
            "Found $($actionLines.Count) lines in $($file.Name)" | Out-File -FilePath $debugLog -Append
            Write-Host "Found $($actionLines.Count) lines in $($file.Name)"
            $actions += $actionLines
        } catch {
            "Error reading $($file.Name): $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
            Write-Host "Error reading $($file.Name): $($_.Exception.Message)"
            $actions += "Error reading $($file.Name): $($_.Exception.Message)"
        }
    }

    # Get the last 100 actions
    $actions = $actions | Select-Object -Last 100
    $actions = $actions | ForEach-Object { 
        $parts = $_ -split " ", 3
        [PSCustomObject]@{ 
            Timestamp = if ($parts.Length -ge 2) { $parts[0..1] -join " " } else { "N/A" }
            Description = if ($parts.Length -ge 3) { $parts[2] } else { $_ }
        }
    }

    if ($actions.Count -eq 0) {
        "No log entries found" | Out-File -FilePath $debugLog -Append
        Write-Host "No log entries found"
        $actions = @([PSCustomObject]@{ Timestamp = "N/A"; Description = "No log entries found." })
    }
    "Returning $($actions.Count) actions" | Out-File -FilePath $debugLog -Append
    Write-Host "Returning $($actions.Count) actions"
    return $actions
}

# Function to get BESClient last check-in time
function Get-BESClientLastCheckIn {
    $regPaths = @(
        "HKLM:\SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client",
        "HKLM:\SOFTWARE\BigFix\EnterpriseClient\Settings\Client"
    )
    $logPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"

    foreach ($regPath in $regPaths) {
        "Checking registry: $regPath" | Out-File -FilePath $debugLog -Append
        Write-Host "Checking registry: $regPath"
        try {
            $lastReportTime = Get-ItemProperty -Path $regPath -Name "LastReportTime" -ErrorAction Stop
            $unixTime = $lastReportTime.LastReportTime / 1000
            $checkInTime = (Get-Date "1970-01-01").AddSeconds($unixTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
            "Found check-in time in registry: $checkInTime" | Out-File -FilePath $debugLog -Append
            Write-Host "Found check-in time: $checkInTime"
            return $checkInTime
        } catch {
            "Registry error at $regPath: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
            Write-Host "Registry error at $regPath: $($_.Exception.Message)"
        }
    }

    # Fallback to log parsing
    "Falling back to log parsing for check-in time" | Out-File -FilePath $debugLog -Append
    Write-Host "Falling back to log parsing"
    try {
        $latestLog = Get-ChildItem -Path $logPath -Filter "*.log" -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($latestLog) {
            "Parsing log: $($latestLog.FullName)" | Out-File -FilePath $debugLog -Append
            Write-Host "Parsing log: $($latestLog.FullName)"
            $content = Get-Content $latestLog.FullName -ErrorAction Stop
            $checkInLine = $content | Where-Object { $_ -match "Report posted successfully|Gather completed" } | Select-Object -Last 1
            if ($checkInLine) {
                $timestamp = ($checkInLine -split " ")[0..1] -join " "
                "Found check-in time in log: $timestamp" | Out-File -FilePath $debugLog -Append
                Write-Host "Found check-in time: $timestamp"
                return $timestamp
            } else {
                "No check-in lines found in log" | Out-File -FilePath $debugLog -Append
                Write-Host "No check-in lines found"
                return "No check-in data found in logs."
            }
        } else {
            "No log files found for check-in" | Out-File -FilePath $debugLog -Append
            Write-Host "No log files found"
            return "No log files found."
        }
    } catch {
        "Log parsing error: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
        Write-Host "Log parsing error: $($_.Exception.Message)"
        return "Error parsing logs: $($_.Exception.Message)"
    }
}

# Create the GUI form
try {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "BigFix Action Viewer"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    "Created GUI form" | Out-File -FilePath $debugLog -Append
    Write-Host "Created GUI form"
} catch {
    "Error creating form: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
    Write-Host "Error creating form: $($_.Exception.Message)"
    exit
}

# Label for last check-in
try {
    $labelCheckIn = New-Object System.Windows.Forms.Label
    $checkInTime = Get-BESClientLastCheckIn
    $labelCheckIn.Text = "Last BESClient Check-In: $checkInTime"
    $labelCheckIn.Location = New-Object System.Drawing.Point(10, 10)
    $labelCheckIn.AutoSize = $true
    $form.Controls.Add($labelCheckIn)
    "Added check-in label: $checkInTime" | Out-File -FilePath $debugLog -Append
    Write-Host "Added check-in label: $checkInTime"
} catch {
    "Error adding check-in label: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
    Write-Host "Error adding check-in label: $($_.Exception.Message)"
}

# DataGridView for actions
try {
    $dataGrid = New-Object System.Windows.Forms.DataGridView
    $dataGrid.Location = New-Object System.Drawing.Point(10, 40)
    $dataGrid.Size = New-Object System.Drawing.Size(760, 500)
    $dataGrid.ReadOnly = $true
    $dataGrid.AutoSizeColumnsMode = "Fill"
    $dataGrid.DataSource = $null
    $actions = Get-LastBigFixActions
    $dataGrid.DataSource = $actions
    $form.Controls.Add($dataGrid)
    "Added DataGridView with $($actions.Count) actions" | Out-File -FilePath $debugLog -Append
    Write-Host "Added DataGridView with $($actions.Count) actions"
} catch {
    "Error adding DataGridView: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
    Write-Host "Error adding DataGridView: $($_.Exception.Message)"
}

# Show the form
try {
    $form.ShowDialog() | Out-Null
    "Form closed" | Out-File -FilePath $debugLog -Append
    Write-Host "Form closed"
} catch {
    "Error showing form: $($_.Exception.Message)" | Out-File -FilePath $debugLog -Append
    Write-Host "Error showing form: $($_.Exception.Message)"
}
