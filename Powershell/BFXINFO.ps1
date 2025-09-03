# Import necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to get last 100 BigFix actions from logs
function Get-LastBigFixActions {
    $logPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
    try {
        $logFiles = Get-ChildItem -Path $logPath -Filter "*.log" -ErrorAction Stop | Sort-Object LastWriteTime -Descending | Select-Object -First 10
    } catch {
        return @([PSCustomObject]@{ Timestamp = "N/A"; Description = "Error accessing log files: $($_.Exception.Message)" })
    }

    $actions = @()
    foreach ($file in $logFiles) {
        try {
            $content = Get-Content $file.FullName -ErrorAction Stop
            $actionLines = $content | Where-Object { $_ -match "Action|Command|Started|Completed|Failed" }
            $actions += $actionLines
        } catch {
            $actions += "Error reading $($file.Name): $($_.Exception.Message)"
        }
    }

    # Get the last 100 actions, reverse for newest first
    $actions = $actions | Select-Object -Last 100
    $actions = $actions | ForEach-Object { [PSCustomObject]@{ Timestamp = ($_ -split " ")[0..1] -join " "; Description = $_ } }

    if ($actions.Count -eq 0) {
        $actions = @([PSCustomObject]@{ Timestamp = "N/A"; Description = "No action-related log entries found." })
    }
    return $actions
}

# Function to get BESClient last check-in time
function Get-BESClientLastCheckIn {
    # Try primary registry path (32-bit app on 64-bit OS)
    $regPath1 = "HKLM:\SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client"
    $regPath2 = "HKLM:\SOFTWARE\BigFix\EnterpriseClient\Settings\Client"  # Fallback for non-WOW64 or custom setups
    $logPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"

    # Try registry first
    try {
        $lastReportTime = Get-ItemProperty -Path $regPath1 -Name "LastReportTime" -ErrorAction Stop
        $unixTime = $lastReportTime.LastReportTime / 1000
        return (Get-Date "1970-01-01").AddSeconds($unixTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
    } catch {
        try {
            # Try alternative registry path
            $lastReportTime = Get-ItemProperty -Path $regPath2 -Name "LastReportTime" -ErrorAction Stop
            $unixTime = $lastReportTime.LastReportTime / 1000
            return (Get-Date "1970-01-01").AddSeconds($unixTime).ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss")
        } catch {
            # Fallback to parsing logs
            try {
                $latestLog = Get-ChildItem -Path $logPath -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
                if ($latestLog) {
                    $content = Get-Content $latestLog.FullName
                    $checkInLine = $content | Where-Object { $_ -match "Report posted successfully" } | Select-Object -Last 1
                    if ($checkInLine) {
                        $timestamp = ($checkInLine -split " ")[0..1] -join " "
                        return $timestamp
                    }
                }
                return "Unable to retrieve last check-in time (no registry key or log data found)."
            } catch {
                return "Error accessing logs for check-in time: $($_.Exception.Message)"
            }
        }
    }
}

# Create the GUI form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Action Viewer"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"

# Label for last check-in
$labelCheckIn = New-Object System.Windows.Forms.Label
$labelCheckIn.Text = "Last BESClient Check-In: " + (Get-BESClientLastCheckIn)
$labelCheckIn.Location = New-Object System.Drawing.Point(10, 10)
$labelCheckIn.AutoSize = $true
$form.Controls.Add($labelCheckIn)

# DataGridView for actions
$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(10, 40)
$dataGrid.Size = New-Object System.Drawing.Size(760, 500)
$dataGrid.ReadOnly = $true
$dataGrid.AutoSizeColumnsMode = "Fill"
$dataGrid.DataSource = Get-LastBigFixActions
$form.Controls.Add($dataGrid)

# Show the form
$form.ShowDialog() | Out-Null
