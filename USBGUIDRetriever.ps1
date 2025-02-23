<#
.SYNOPSIS
Enterprise USB GUID Retriever - Enhanced Edition

.DESCRIPTION
PowerShell GUI with automatic GUID loading and multi-drive support
#>

#region Initialization
#Requires -RunAsAdministrator
#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$ExecutionPolicy = Get-ExecutionPolicy
if ($ExecutionPolicy -notin "RemoteSigned", "Unrestricted") {
    Write-Error "Insufficient execution policy: $ExecutionPolicy"
    exit 1
}

$LogPath = "$env:ProgramData\USBGUIDRetriever\logs"
$LogFile = "$LogPath\activity.log"
if (-not (Test-Path $LogPath)) { New-Item -Path $LogPath -ItemType Directory -Force | Out-Null }
#endregion

#region UI Configuration (Unchanged)
class Theme {
    static [hashtable] $Current = @{
        Primary    = [System.Drawing.Color]::FromArgb(63, 81, 181)
        Secondary  = [System.Drawing.Color]::FromArgb(76, 175, 80)
        Tertiary   = [System.Drawing.Color]::FromArgb(158, 158, 158)
        Background = [System.Drawing.Color]::FromArgb(245, 245, 245)
        Text       = [System.Drawing.Color]::FromArgb(33, 37, 41)
        Error      = [System.Drawing.Color]::DarkRed
        ButtonText = [System.Drawing.Color]::White
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "USB GUID Manager Pro"
$form.Size = New-Object System.Drawing.Size(420, 170)
$form.MinimumSize = New-Object System.Drawing.Size(400, 170)
$form.MaximumSize = New-Object System.Drawing.Size([System.Int32]::MaxValue, 170)
$form.StartPosition = "CenterScreen"
$form.BackColor = [Theme]::Current.Background
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.AutoScaleMode = "Dpi"
[System.Windows.Forms.Application]::EnableVisualStyles()

#region UI Controls (Unchanged)
$mainPanel = New-Object System.Windows.Forms.TableLayoutPanel
$mainPanel.Dock = "Fill"
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(5)
$mainPanel.ColumnCount = 2
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "AutoSize")) | Out-Null
$mainPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle "Percent", 100)) | Out-Null
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "AutoSize")) | Out-Null
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "AutoSize")) | Out-Null
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "AutoSize")) | Out-Null
$mainPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle "AutoSize")) | Out-Null

$lblDrive = New-Object System.Windows.Forms.Label
$lblDrive.Text = "Select USB Drive:"
$lblDrive.AutoSize = $true
$lblDrive.Margin = New-Object System.Windows.Forms.Padding(0, 0, 2, 5)

$cmbDrives = New-Object System.Windows.Forms.ComboBox
$cmbDrives.DropDownStyle = "DropDownList"
$cmbDrives.Anchor = "Left,Right"
$cmbDrives.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 5)

$btnPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$btnPanel.AutoSize = $true
$btnPanel.FlowDirection = "LeftToRight"
$btnPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 5)

$btnGet = New-Object System.Windows.Forms.Button
$btnGet.Text = "&Get GUID"
$btnGet.BackColor = [Theme]::Current.Primary
$btnGet.ForeColor = [Theme]::Current.ButtonText
$btnGet.AutoSize = $true
$btnGet.Padding = New-Object System.Windows.Forms.Padding(3)

$btnCopy = New-Object System.Windows.Forms.Button
$btnCopy.Text = "&Copy"
$btnCopy.BackColor = [Theme]::Current.Secondary
$btnCopy.ForeColor = [Theme]::Current.ButtonText
$btnCopy.AutoSize = $true
$btnCopy.Padding = New-Object System.Windows.Forms.Padding(3)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "&Refresh"
$btnRefresh.BackColor = [Theme]::Current.Tertiary
$btnRefresh.ForeColor = [Theme]::Current.ButtonText
$btnRefresh.AutoSize = $true
$btnRefresh.Padding = New-Object System.Windows.Forms.Padding(3)

$btnPanel.Controls.AddRange(@($btnGet, $btnCopy, $btnRefresh)) | Out-Null

$lblGuid = New-Object System.Windows.Forms.Label
$lblGuid.Text = "GUID:"
$lblGuid.AutoSize = $true
$lblGuid.Margin = New-Object System.Windows.Forms.Padding(0, 5, 2, 0)

$txtGuid = New-Object System.Windows.Forms.TextBox
$txtGuid.ReadOnly = $true
$txtGuid.BackColor = [System.Drawing.Color]::White
$txtGuid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$txtGuid.Anchor = "Left,Right"
$txtGuid.Height = 20
$txtGuid.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtGuid.Margin = New-Object System.Windows.Forms.Padding(0)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Style = "Marquee"
$progressBar.Visible = $false
$progressBar.Height = 8
$progressBar.Anchor = "Left,Right"
$progressBar.Margin = New-Object System.Windows.Forms.Padding(0, 5, 0, 0)

$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.Dock = "Bottom"
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusBar.Items.Add($statusLabel) | Out-Null

$mainPanel.Controls.Add($lblDrive, 0, 0)
$mainPanel.Controls.Add($cmbDrives, 1, 0)
$mainPanel.Controls.Add($btnPanel, 0, 1)
$mainPanel.SetColumnSpan($btnPanel, 2)
$mainPanel.Controls.Add($lblGuid, 0, 2)
$mainPanel.Controls.Add($txtGuid, 1, 2)
$mainPanel.Controls.Add($progressBar, 0, 3)
$mainPanel.SetColumnSpan($progressBar, 2)

$form.Controls.AddRange(@($mainPanel, $statusBar))
#endregion

#region Core Logic (Enhanced)
class USBDrive {
    [string] $DisplayName
    [string] $DeviceID
    [string] $Guid
    [datetime] $LastUpdated
}

$driveCache = [System.Collections.Generic.List[USBDrive]]::new()

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error")]
        [string]$Level = "Info"
    )
    $logEntry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $LogFile -Value $logEntry
}

function Update-DriveList {
    try {
        $btnRefresh.Enabled = $false
        $progressBar.Visible = $true
        $statusLabel.Text = "Refreshing drives..."
        $statusLabel.ForeColor = [Theme]::Current.Text
        
        $volumes = Get-CimInstance -ClassName Win32_Volume | 
        Where-Object { $_.DriveType -eq 2 -and $_.DriveLetter } |
        Select-Object DriveLetter, Label, DeviceID

        $driveCache.Clear()
        $cmbDrives.Items.Clear()

        foreach ($vol in $volumes) {
            $driveLetter = $vol.DriveLetter.TrimEnd('\', ':')
            $drive = [USBDrive]@{
                DisplayName = if ($vol.Label) { 
                    "{0} ({1})" -f $vol.Label, $driveLetter 
                }
                else { 
                    "Unlabeled ({0})" -f $driveLetter 
                }
                DeviceID    = $vol.DeviceID
                Guid        = if ($vol.DeviceID -match '{([A-F0-9-]+)}') { $matches[1] }
                LastUpdated = [datetime]::Now
            }
            $driveCache.Add($drive)
        }

        if ($driveCache.Count -gt 0) {
            $driveCache | ForEach-Object { $cmbDrives.Items.Add($_.DisplayName) }
            $cmbDrives.SelectedIndex = 0
            $statusLabel.Text = "Ready - {0} drive(s) found" -f $driveCache.Count
            return $true
        }
        else {
            $statusLabel.Text = "No USB drives detected"
            return $false
        }
    }
    catch {
        Write-Log "Drive refresh failed: $_" -Level Error
        $statusLabel.Text = "Error: $($_.Exception.Message)"
        $statusLabel.ForeColor = [Theme]::Current.Error
        return $false
    }
    finally {
        $btnRefresh.Enabled = $true
        $progressBar.Visible = $false
    }
}

function Get-Guid {
    try {
        if (-not $cmbDrives.SelectedItem) { throw "No drive selected" }
        $selected = $driveCache[$cmbDrives.SelectedIndex]
        
        if (-not $selected.Guid) {
            throw "No GUID found for selected drive"
        }
        
        $txtGuid.Text = $selected.Guid
        $statusLabel.Text = "GUID retrieved successfully"
    }
    catch {
        Write-Log "GUID retrieval failed: $_" -Level Error
        $statusLabel.Text = $_.Exception.Message
        $statusLabel.ForeColor = [Theme]::Current.Error
    }
}

function Copy-Guid {
    try {
        if ([string]::IsNullOrEmpty($txtGuid.Text)) { return }
        
        [System.Windows.Forms.Clipboard]::SetText($txtGuid.Text)
        $statusLabel.Text = "GUID copied to clipboard"
    }
    catch {
        Write-Log "Copy failed: $_" -Level Error
        $statusLabel.Text = $_.Exception.Message
        $statusLabel.ForeColor = [Theme]::Current.Error
    }
}
#endregion

#region Event Handlers (Enhanced)
$form.Add_Load({
        if (Update-DriveList) {
            Get-Guid  # Auto-load first drive's GUID on startup
        }
    })

$btnRefresh.Add_Click({
        if (Update-DriveList) {
            Get-Guid  # Auto-load first drive's GUID after refresh
        }
    })

$btnGet.Add_Click({ Get-Guid })
$btnCopy.Add_Click({ Copy-Guid })

$form.Add_FormClosing({
        Write-Log "Application closed"
    })
#endregion

[void]$form.ShowDialog()
