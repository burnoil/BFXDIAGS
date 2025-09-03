# Import necessary assemblies for GUI
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to get last 100 BigFix actions from logs
function Get-LastBigFixActions {
    $logPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"
    $logFiles = Get-ChildItem -Path $logPath -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 10  # Limit to recent files to avoid overload

    $actions = @()
    foreach ($file in $logFiles) {
        $content = Get-Content $file.FullName
        $actionLines = $content | Where-Object { $_ -match "Action|Command|Started|Completed|Failed" }  # Filter for action-related lines
        $actions += $actionLines
    }

    # Get the last 100 (or fewer if not available), reverse to show newest first
    $actions = $actions | Select-Object -Last 100
    $actions = $actions | ForEach-Object { [PSCustomObject]@{ Timestamp = ($_ -split " ")[0..1] -join " "; Description = $_ } }

    return $actions
}

# Function to get BESClient last check-in time
function Get-BESClientLastCheckIn {
    $regPath = "HKLM:\SOFTWARE\WOW6432Node\BigFix\EnterpriseClient\Settings\Client"
    $lastReportTime = Get-ItemProperty -Path $regPath -Name "LastReportTime" -ErrorAction SilentlyContinue

    if ($lastReportTime) {
        $unixTime = $lastReportTime.LastReportTime / 1000  # Convert from milliseconds
        $checkInTime = (Get-Date "1970-01-01").AddSeconds($unixTime).ToLocalTime()
        return $checkInTime.ToString("yyyy-MM-dd HH:mm:ss")
    } else {
        return "Unable to retrieve last check-in time (registry key not found)."
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