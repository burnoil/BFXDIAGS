# Load required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define the BigFix log path (adjust if needed)
$bigFixLogPath = "C:\Program Files (x86)\BigFix Enterprise\BES Client\__BESData\__Global\Logs"

# Get machine info
$machineName = [System.Net.Dns]::GetHostName()
$ipAddress = [System.Net.Dns]::GetHostEntry($machineName).AddressList | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1 | ForEach-Object { $_.IPAddressToString }

# Global keywords (default; can be customized)
$globalKeywords = @{
    Success = @{Patterns = @('SUCCESS', 'SUCCESSFULLY'); Color = [System.Drawing.Color]::LightGreen}
    Warning = @{Patterns = @('WARNING'); Color = [System.Drawing.Color]::LightYellow}
    Fail = @{Patterns = @('FAIL', 'FAILURE', 'FAILED'); Color = [System.Drawing.Color]::LightCoral}
}

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "BigFix Log Viewer"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"

# Create TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$form.Controls.Add($tabControl)

# Menu strip with integrated search
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem("File")
$menuOpen = New-Object System.Windows.Forms.ToolStripMenuItem("Open Log...")
$menuExport = New-Object System.Windows.Forms.ToolStripMenuItem("Export View...")
$menuExit = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
$menuPause = New-Object System.Windows.Forms.ToolStripMenuItem("Pause Tailing")
$menuPause.CheckOnClick = $true
$menuFile.DropDownItems.Add($menuOpen)
$menuFile.DropDownItems.Add($menuExport)
$menuFile.DropDownItems.Add($menuPause)
$menuFile.DropDownItems.Add($menuExit)

$menuTools = New-Object System.Windows.Forms.ToolStripMenuItem("Tools")
$menuRestartBES = New-Object System.Windows.Forms.ToolStripMenuItem("Restart BESClient")
$menuTools.DropDownItems.Add($menuRestartBES)

$menuView = New-Object System.Windows.Forms.ToolStripMenuItem("View")
$menuWrap = New-Object System.Windows.Forms.ToolStripMenuItem("Word Wrap")
$menuWrap.CheckOnClick = $true
$menuWrap.Checked = $true
$menuZoomIn = New-Object System.Windows.Forms.ToolStripMenuItem("Zoom In")
$menuZoomOut = New-Object System.Windows.Forms.ToolStripMenuItem("Zoom Out")
$menuStats = New-Object System.Windows.Forms.ToolStripMenuItem("Show Stats")
$menuView.DropDownItems.Add($menuWrap)
$menuView.DropDownItems.Add($menuZoomIn)
$menuView.DropDownItems.Add($menuZoomOut)
$menuView.DropDownItems.Add($menuStats)

$menuFilter = New-Object System.Windows.Forms.ToolStripMenuItem("Filter")
$menuShowAll = New-Object System.Windows.Forms.ToolStripMenuItem("Show All")
$menuShowErrors = New-Object System.Windows.Forms.ToolStripMenuItem("Show Errors Only")
$menuShowWarnings = New-Object System.Windows.Forms.ToolStripMenuItem("Show Warnings Only")
$menuShowSuccess = New-Object System.Windows.Forms.ToolStripMenuItem("Show Success Only")
$menuCustomizeKeywords = New-Object System.Windows.Forms.ToolStripMenuItem("Customize Keywords...")
$menuCustomizeColors = New-Object System.Windows.Forms.ToolStripMenuItem("Customize Colors...")
$menuFilter.DropDownItems.Add($menuShowAll)
$menuFilter.DropDownItems.Add($menuShowErrors)
$menuFilter.DropDownItems.Add($menuShowWarnings)
$menuFilter.DropDownItems.Add($menuShowSuccess)
$menuFilter.DropDownItems.Add($menuCustomizeKeywords)
$menuFilter.DropDownItems.Add($menuCustomizeColors)

$menuSearch = New-Object System.Windows.Forms.ToolStripMenuItem("Search")
$searchBox = New-Object System.Windows.Forms.ToolStripTextBox
$searchBox.Width = 200
$searchButton = New-Object System.Windows.Forms.ToolStripButton -Property @{Text = "Search"}
$menuSearch.DropDownItems.Add($searchBox)
$menuSearch.DropDownItems.Add($searchButton)

$menuStrip.Items.Add($menuFile)
$menuStrip.Items.Add($menuTools)
$menuStrip.Items.Add($menuView)
$menuStrip.Items.Add($menuFilter)
$menuStrip.Items.Add($menuSearch)
$form.MainMenuStrip = $menuStrip
$form.Controls.Add($menuStrip)

# Tab context menu for closing
$tabContextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$closeTabMenu = New-Object System.Windows.Forms.ToolStripMenuItem("Close Tab")
$tabContextMenu.Items.Add($closeTabMenu)
$tabControl.ContextMenuStrip = $tabContextMenu

# Timer for auto-tailing (checks every 1 sec)
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$timer.Add_Tick({
    if ($menuPause.Checked) { return }  # Paused
    foreach ($tab in $tabControl.TabPages) {
        $rtb = $tab.Controls[0].Controls[0]  # RichTextBox in panel
        $ht = $tab.Tag
        $fp = $ht.FilePath
        $currentLength = (Get-Item $fp).Length
        if ($currentLength -gt $ht.LastLength) {
            try {
                $fs = New-Object System.IO.FileStream($fp, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $fs.Seek($ht.LastLength, [System.IO.SeekOrigin]::Begin)
                $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
                $newPart = $sr.ReadToEnd().Trim()  # Trim all whitespace/newlines
                $sr.Close()
                $fs.Close()

                if ($newPart) {
                    $lines = $newPart -split "`n" | Where-Object { $_.Trim() }
                    $ht.FullContent += $newPart + "`n"  # Append to full content
                    Add-ColoredText -rtb $rtb -text $newPart -filter $ht.Filter
                    $rtb.SelectionStart = $rtb.Text.Length
                    $rtb.ScrollToCaret()
                    $ht.LastLength = $currentLength

                    # Update status
                    $statusLabel = $tab.Controls[0].Controls[1]
                    $fileSizeKB = [math]::Round($currentLength / 1024, 2)
                    $charCount = $rtb.Text.Length
                    $statusLabel.Text = "Machine: $machineName | IP: $ipAddress | Path: $fp | Size: $fileSizeKB KB | Chars: $charCount | Filter: $($ht.Filter)"
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error tailing file $fp : $($_)", "Tailing Error")
            }
        }
    }
})
$timer.Start()

# Directory watcher for log rotation
$dirWatcher = New-Object System.IO.FileSystemWatcher
$dirWatcher.Path = $bigFixLogPath
$dirWatcher.Filter = "*.log"
$dirWatcher.IncludeSubdirectories = $false
$dirWatcher.EnableRaisingEvents = $true
$dirWatcher.Add_Created({
    $newFile = $_.FullPath
    $msg = [System.Windows.Forms.MessageBox]::Show("New log detected: $(Split-Path $newFile -Leaf). Open in new tab?", "New Log", "YesNo")
    if ($msg -eq "Yes") {
        Add-LogTab -filePath $newFile
    }
})

# Function to append text with row highlighting and optional filter
function Add-ColoredText {
    param (
        $rtb,
        [string]$text,
        [string]$filter = "All"
    )

    $lines = $text -split "`n" | Where-Object { $_.Trim() }  # Skip empty or whitespace-only lines
    foreach ($line in $lines) {
        $matchType = Get-MatchType -line $line
        if ($filter -ne "All" -and $matchType -ne $filter) { continue }

        $rtb.SelectionStart = $rtb.Text.Length
        $rtb.SelectionLength = 0
        $rtb.SelectionColor = [System.Drawing.Color]::Black
        $rtb.SelectionBackColor = [System.Drawing.Color]::White

        $startPos = $rtb.Text.Length
        $rtb.AppendText($line + "`n")

        $backColor = $null
        if ($matchType -eq "Fail") { $backColor = $globalKeywords.Fail.Color }
        elseif ($matchType -eq "Warning") { $backColor = $globalKeywords.Warning.Color }
        elseif ($matchType -eq "Success") { $backColor = $globalKeywords.Success.Color }

        if ($backColor) {
            $rtb.SelectionStart = $startPos
            $rtb.SelectionLength = $line.Length + 1
            $rtb.SelectionBackColor = $backColor
        }

        $rtb.SelectionStart = $rtb.Text.Length
        $rtb.SelectionLength = 0
    }
}

# Helper to get match type and count for stats
function Get-MatchType {
    param ([string]$line)
    foreach ($key in $globalKeywords.Fail.Patterns) { if ($line -imatch "\b$key\b") { return "Fail" } }
    foreach ($key in $globalKeywords.Warning.Patterns) { if ($line -imatch "\b$key\b") { return "Warning" } }
    foreach ($key in $globalKeywords.Success.Patterns) { if ($line -imatch "\b$key\b") { return "Success" } }
    return $null
}

# Function to get log stats
function Get-LogStats {
    param ([string]$text)
    $successCount = 0
    $warningCount = 0
    $failCount = 0
    $lines = $text -split "`n" | Where-Object { $_.Trim() }
    foreach ($line in $lines) {
        $matchType = Get-MatchType -line $line
        if ($matchType -eq "Success") { $successCount++ }
        elseif ($matchType -eq "Warning") { $warningCount++ }
        elseif ($matchType -eq "Fail") { $failCount++ }
    }
    return "Success: $successCount, Warning: $warningCount, Fail: $failCount"
}

# Function to refresh all tabs with new colors
function Refresh-Tabs {
    foreach ($tab in $tabControl.TabPages) {
        $rtb = $tab.Controls[0].Controls[0]
        $ht = $tab.Tag
        $rtb.Text = ""
        Add-ColoredText -rtb $rtb -text $ht.FullContent -filter $ht.Filter
        $rtb.SelectionStart = $rtb.Text.Length
        $rtb.ScrollToCaret()
        $statusLabel = $tab.Controls[0].Controls[1]
        $fileSizeKB = [math]::Round((Get-Item $ht.FilePath).Length / 1024, 2)
        $statusLabel.Text = "Machine: $machineName | IP: $ipAddress | Path: $($ht.FilePath) | Size: $fileSizeKB KB | Chars: $($rtb.Text.Length) | Filter: $($ht.Filter)"
    }
}

# Function to restart BESClient service with progress
function Restart-BESClient {
    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Restarting BESClient"
    $progressForm.Size = New-Object System.Drawing.Size(300, 150)
    $progressForm.StartPosition = "CenterScreen"

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(10, 10)
    $progressBar.Size = New-Object System.Drawing.Size(260, 20)
    $progressBar.Maximum = 100
    $progressForm.Controls.Add($progressBar)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(10, 40)
    $statusLabel.Size = New-Object System.Drawing.Size(260, 20)
    $progressForm.Controls.Add($statusLabel)

    $progressForm.Show()

    try {
        $statusLabel.Text = "Stopping BESClient service..."
        Stop-Service -Name "BESClient" -Force
        $progressBar.Value = 50

        $statusLabel.Text = "Starting BESClient service..."
        Start-Service -Name "BESClient"
        $progressBar.Value = 100

        [System.Windows.Forms.MessageBox]::Show("BESClient restarted successfully!", "Success")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error restarting BESClient: $($_)", "Error")
    } finally {
        $progressForm.Close()
    }
}

# Function to add a new tab for a log file
function Add-LogTab {
    param ([string]$filePath)

    if (-not (Test-Path $filePath)) {
        [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Error")
        return
    }

    $tabPage = New-Object System.Windows.Forms.TabPage
    $tabPage.Text = (Get-Item $filePath).Name

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = "Fill"

    $statusLabel = New-Object System.Windows.Forms.TextBox
    $statusLabel.Dock = "Bottom"
    $statusLabel.Height = 40
    $statusLabel.Multiline = $true
    $statusLabel.ReadOnly = $true
    $statusLabel.BorderStyle = "None"
    $statusLabel.BackColor = [System.Drawing.Color]::LightGray
    $fileSizeKB = [math]::Round((Get-Item $filePath).Length / 1024, 2)
    $statusLabel.Text = "Machine: $machineName | IP: $ipAddress | Path: $filePath | Size: $fileSizeKB KB | Chars: 0 | Filter: All"

    $richTextBox = New-Object System.Windows.Forms.RichTextBox
    $richTextBox.Dock = "Fill"
    $richTextBox.ReadOnly = $true
    $richTextBox.Multiline = $true
    $richTextBox.ScrollBars = "Vertical"
    $richTextBox.Font = New-Object System.Drawing.Font("Consolas", 10)
    $richTextBox.MaxLength = [int]::MaxValue
    $richTextBox.WordWrap = $true

    $panel.Controls.Add($richTextBox)
    $panel.Controls.Add($statusLabel)
    $tabPage.Controls.Add($panel)

    try {
        $initialContent = Get-Content -Path $filePath -Raw -Encoding UTF8
        Add-ColoredText -rtb $richTextBox -text $initialContent
        $statusLabel.Text = "Machine: $machineName | IP: $ipAddress | Path: $filePath | Size: $fileSizeKB KB | Chars: $($richTextBox.Text.Length) | Filter: All"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading file: $($_)", "Error")
        return
    }

    $richTextBox.SelectionStart = $richTextBox.Text.Length
    $richTextBox.ScrollToCaret()

    $lastLength = (Get-Item $filePath).Length
    $tabPage.Tag = @{FilePath = $filePath; LastLength = $lastLength; Filter = "All"; FullContent = $initialContent}

    $tabControl.TabPages.Add($tabPage)
}

# Apply filter to current tab
function Apply-Filter {
    param ([string]$filterType)

    $tab = $tabControl.SelectedTab
    if ($tab -eq $null) { return }
    $rtb = $tab.Controls[0].Controls[0]
    $ht = $tab.Tag
    $ht.Filter = $filterType
    $rtb.Text = ""
    Add-ColoredText -rtb $rtb -text $ht.FullContent -filter $filterType
    $statusLabel = $tab.Controls[0].Controls[1]
    $fileSizeKB = [math]::Round((Get-Item $ht.FilePath).Length / 1024, 2)
    $statusLabel.Text = "Machine: $machineName | IP: $ipAddress | Path: $($ht.FilePath) | Size: $fileSizeKB KB | Chars: $($rtb.Text.Length) | Filter: $filterType"
    $rtb.SelectionStart = $rtb.Text.Length
    $rtb.ScrollToCaret()
}

# Events
$menuOpen.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        Add-LogTab -filePath $openFileDialog.FileName
    }
})

$menuExport.Add_Click({
    $tab = $tabControl.SelectedTab
    if ($tab -eq $null) { return }
    $rtb = $tab.Controls[0].Controls[0]
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Log Files (*.log)|*.log|All Files (*.*)|*.*"
    $saveFileDialog.FileName = "Exported_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        try {
            $rtb.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show("Exported to $($saveFileDialog.FileName)", "Success")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting file: $($_)", "Error")
        }
    }
})

$menuRestartBES.Add_Click({ Restart-BESClient })

$menuExit.Add_Click({ $form.Close() })

$menuWrap.Add_Click({
    $rtb = $tabControl.SelectedTab.Controls[0].Controls[0]
    $rtb.WordWrap = $menuWrap.Checked
})

$menuZoomIn.Add_Click({
    $rtb = $tabControl.SelectedTab.Controls[0].Controls[0]
    $newSize = $rtb.Font.Size + 2
    $rtb.Font = New-Object System.Drawing.Font("Consolas", $newSize)
})

$menuZoomOut.Add_Click({
    $rtb = $tabControl.SelectedTab.Controls[0].Controls[0]
    $newSize = [math]::Max($rtb.Font.Size - 2, 8)
    $rtb.Font = New-Object System.Drawing.Font("Consolas", $newSize)
})

$menuStats.Add_Click({
    $tab = $tabControl.SelectedTab
    if ($tab -eq $null) { return }
    $rtb = $tab.Controls[0].Controls[0]
    $stats = Get-LogStats -text $rtb.Text
    [System.Windows.Forms.MessageBox]::Show("Log Stats: $stats", "Log Statistics")
})

$searchButton.Add_Click({
    $rtb = $tabControl.SelectedTab.Controls[0].Controls[0]
    $searchText = $searchBox.Text
    if ($searchText) {
        $start = $rtb.SelectionStart + $rtb.SelectionLength
        $pos = $rtb.Find($searchText, $start, [System.Windows.Forms.RichTextBoxFinds]::None)
        if ($pos -ge 0) {
            $rtb.SelectionStart = $pos
            $rtb.SelectionLength = $searchText.Length
            $rtb.SelectionBackColor = [System.Drawing.Color]::LightBlue
            $rtb.Focus()
        } else {
            [System.Windows.Forms.MessageBox]::Show("No more matches found.")
        }
    }
})

$menuShowAll.Add_Click({ Apply-Filter -filterType "All" })
$menuShowErrors.Add_Click({ Apply-Filter -filterType "Fail" })
$menuShowWarnings.Add_Click({ Apply-Filter -filterType "Warning" })
$menuShowSuccess.Add_Click({ Apply-Filter -filterType "Success" })

$menuCustomizeKeywords.Add_Click({
    $keywordForm = New-Object System.Windows.Forms.Form
    $keywordForm.Text = "Customize Keywords"
    $keywordForm.Size = New-Object System.Drawing.Size(400, 200)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Add Keyword (e.g., ERROR) to Category (Success/Warning/Fail):"
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $keywordForm.Controls.Add($label)

    $txtKeyword = New-Object System.Windows.Forms.TextBox
    $txtKeyword.Location = New-Object System.Drawing.Point(10, 40)
    $keywordForm.Controls.Add($txtKeyword)

    $cmbCategory = New-Object System.Windows.Forms.ComboBox
    $cmbCategory.Items.AddRange(@("Success", "Warning", "Fail"))
    $cmbCategory.Location = New-Object System.Drawing.Point(150, 40)
    $keywordForm.Controls.Add($cmbCategory)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.Text = "Add"
    $btnAdd.Location = New-Object System.Drawing.Point(10, 80)
    $btnAdd.Add_Click({
        if ($txtKeyword.Text -and $cmbCategory.SelectedItem) {
            $globalKeywords[$cmbCategory.SelectedItem].Patterns += $txtKeyword.Text
            [System.Windows.Forms.MessageBox]::Show("Keyword added! Reload tabs to apply.")
        }
        $keywordForm.Close()
    })
    $keywordForm.Controls.Add($btnAdd)

    $keywordForm.ShowDialog()
})

$menuCustomizeColors.Add_Click({
    $colorForm = New-Object System.Windows.Forms.Form
    $colorForm.Text = "Customize Colors"
    $colorForm.Size = New-Object System.Drawing.Size(400, 250)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select colors for each category:"
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $colorForm.Controls.Add($label)

    $btnSuccess = New-Object System.Windows.Forms.Button
    $btnSuccess.Text = "Success Color"
    $btnSuccess.Location = New-Object System.Drawing.Point(10, 40)
    $colorForm.Controls.Add($btnSuccess)

    $btnWarning = New-Object System.Windows.Forms.Button
    $btnWarning.Text = "Warning Color"
    $btnWarning.Location = New-Object System.Drawing.Point(10, 80)
    $colorForm.Controls.Add($btnWarning)

    $btnFail = New-Object System.Windows.Forms.Button
    $btnFail.Text = "Fail Color"
    $btnFail.Location = New-Object System.Drawing.Point(10, 120)
    $colorForm.Controls.Add($btnFail)

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = "Apply"
    $btnApply.Location = New-Object System.Drawing.Point(10, 160)
    $colorForm.Controls.Add($btnApply)

    $colorDialog = New-Object System.Windows.Forms.ColorDialog

    $btnSuccess.Add_Click({
        if ($colorDialog.ShowDialog() -eq "OK") {
            $globalKeywords.Success.Color = $colorDialog.Color
        }
    })

    $btnWarning.Add_Click({
        if ($colorDialog.ShowDialog() -eq "OK") {
            $globalKeywords.Warning.Color = $colorDialog.Color
        }
    })

    $btnFail.Add_Click({
        if ($colorDialog.ShowDialog() -eq "OK") {
            $globalKeywords.Fail.Color = $colorDialog.Color
        }
    })

    $btnApply.Add_Click({
        Refresh-Tabs
        $colorForm.Close()
    })

    $colorForm.ShowDialog()
})

$closeTabMenu.Add_Click({
    if ($tabControl.SelectedTab) {
        $tabControl.TabPages.Remove($tabControl.SelectedTab)
    }
})

# Load the most recent BigFix log (1 file)
if (Test-Path $bigFixLogPath) {
    $logFile = Get-ChildItem -Path $bigFixLogPath -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($logFile) {
        Add-LogTab -filePath $logFile.FullName
    } else {
        [System.Windows.Forms.MessageBox]::Show("No log files found in BigFix path. You can open logs manually.", "Info")
    }
} else {
    [System.Windows.Forms.MessageBox]::Show("BigFix log path not found. You can open logs manually.", "Info")
}

# Form closing cleanup
$form.Add_Closing({
    $timer.Stop()
    $dirWatcher.Dispose()
})

# Show the form
$form.ShowDialog() | Out-Null
