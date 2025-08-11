<#
.SYNOPSIS
    Enterprise USB GUID/UUID Retriver - Fixed Version
    
.DESCRIPTION
    PowerShell WPF application for retrieving USB drive GUIDs and UUIDs with
    simplified detection that works reliably with all USB drives.

.NOTES
    Version:        2.0
    Requirements:   Windows PowerShell 5.1+, Administrator privileges
    Changes:        - Removed the redundant "Get Identifiers" button to improve and simplify the user interface.
                      Data is now retrieved automatically on selection change or refresh.
                   - Implemented a definitive hardware-level detection method.
#>

#region 1_Initialization_and_Requirements
#Requires -RunAsAdministrator
#Requires -Version 5.1

using namespace System.Windows
using namespace System.Windows.Controls
using namespace System.Windows.Media
using namespace System.Collections.Generic

# Application Constants
$SCRIPT_NAME = "USBGUIDUUIDRetriver"
$LOG_PATH = "$env:ProgramData\$SCRIPT_NAME\logs"
$LOG_FILE = "$LOG_PATH\activity.log"
$CACHE_EXPIRY_MINUTES = 5
$DEBOUNCE_INTERVAL_MS = 500

# Validate execution policy
$executionPolicy = Get-ExecutionPolicy
if ($executionPolicy -notin @("RemoteSigned", "Unrestricted", "Bypass")) {
    Write-Error "Insufficient execution policy: $executionPolicy. Required: RemoteSigned or higher."
    exit 1
}

# Initialize logging directory
if (-not (Test-Path -Path $LOG_PATH -PathType Container)) {
    try {
        New-Item -Path $LOG_PATH -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Warning "Failed to create log directory: $_"
    }
}

# Load required assemblies
try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
    Add-Type -AssemblyName System.Drawing, System.Management
}
catch {
    Write-Error "Failed to load required assemblies: $_"
    exit 1
}
#endregion

#region 2_UI_Definition_XAML
$xamlDefinition = @'
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="USB GUID/UUID Retriver Pro" 
    Height="250" 
    Width="550"
    MinHeight="250" 
    MinWidth="500"
    MaxHeight="250"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResizeWithGrip"
    Background="#FFF5F5F5"
    FontFamily="Segoe UI Variable Display"
    FontSize="11">
    
    <!-- Window Resources for Windows 11 Fluent Design -->
    <Window.Resources>
        <!-- Button Style with Fluent Design -->
        <Style x:Key="FluentButton" TargetType="Button">
            <Setter Property="Background" Value="#FF3F51B5"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="4,2"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}" 
                                CornerRadius="6"
                                BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"
                                            Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.9"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.8"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Small Button Style for Copy buttons -->
        <Style x:Key="SmallCopyButton" TargetType="Button" BasedOn="{StaticResource FluentButton}">
            <Setter Property="Padding" Value="8,3"/>
            <Setter Property="Margin" Value="4,0,0,0"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="Background" Value="#FF4CAF50"/>
            <Setter Property="MinWidth" Value="60"/>
        </Style>

        <!-- ComboBox Style -->
        <Style x:Key="FluentComboBox" TargetType="ComboBox">
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="BorderBrush" Value="#DDDDDD"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
        </Style>

        <!-- TextBox Style -->
        <Style x:Key="FluentTextBox" TargetType="TextBox">
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="BorderBrush" Value="#DDDDDD"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="FontFamily" Value="Cascadia Mono, Consolas"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Main Content Panel -->
        <Border Grid.Row="0" Margin="12" Background="White" CornerRadius="8">
            <Border.Effect>
                <DropShadowEffect ShadowDepth="0" BlurRadius="8" Opacity="0.1"/>
            </Border.Effect>
            
            <Grid Margin="16">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Drive Selection -->
                <Grid Grid.Row="0" Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="80"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    
                    <Label Grid.Column="0" 
                           Content="USB Drive:" 
                           VerticalAlignment="Center"/>
                    
                    <ComboBox Grid.Column="1" 
                              Name="cmbDrives" 
                              Style="{StaticResource FluentComboBox}"
                              ToolTip="Select a USB drive to retrieve its identifiers"/>
                </Grid>

                <!-- Button Panel -->
                    <StackPanel Grid.Row="1" 
                            Orientation="Horizontal" 
                            Margin="0,0,0,12">

                    <!-- Refresh: A secondary, light gray utility button -->
                    <Button Name="btnRefresh" 
                            Content="Refresh" 
                            Style="{StaticResource FluentButton}"
                            Background="#FFEEEEEE"
                            Foreground="#FF212121"
                            ToolTip="Refresh drive list"/>

                    <!-- Copy All: The primary action button, using the accent color -->
                    <Button Name="btnCopyAll" 
                            Content="Copy All Info" 
                            Style="{StaticResource FluentButton}"
                            Background="#FF3F51B5"
                            ToolTip="Copy all drive information to clipboard"/>
                </StackPanel>

                <!-- GUID Display -->
                <Grid Grid.Row="2" Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="80"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <Label Grid.Column="0" 
                           Content="GUID:" 
                           VerticalAlignment="Center"/>
                    
                    <TextBox Grid.Column="1" 
                             Name="txtGuid" 
                             Style="{StaticResource FluentTextBox}"
                             ToolTip="Volume GUID (Globally Unique Identifier)"/>
                    
                    <!-- Copy GUID: A specific action, styled green for success -->
                    <Button Grid.Column="2"
                            Name="btnCopyGuid"
                            Content="Copy"
                            Style="{StaticResource SmallCopyButton}"
                            ToolTip="Copy GUID to clipboard"/>
                </Grid>

                <!-- UUID Display -->
                <Grid Grid.Row="3" Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="80"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    
                    <Label Grid.Column="0" 
                           Content="UUID:" 
                           VerticalAlignment="Center"/>
                    
                    <TextBox Grid.Column="1" 
                             Name="txtUuid" 
                             Style="{StaticResource FluentTextBox}"
                             ToolTip="Disk UUID/Serial Number"/>
                    
                    <!-- Copy UUID: A specific action, styled green for success -->
                    <Button Grid.Column="2"
                            Name="btnCopyUuid"
                            Content="Copy"
                            Style="{StaticResource SmallCopyButton}"
                            ToolTip="Copy UUID to clipboard"/>
                </Grid>

                <!-- Progress Bar -->
                <ProgressBar Grid.Row="4" 
                             Name="progressBar" 
                             Height="4" 
                             IsIndeterminate="True"
                             Visibility="Collapsed"
                             Foreground="#FF3F51B5"
                             Background="#FFE0E0E0"
                             BorderThickness="0"/>
            </Grid>
        </Border>

        <!-- Status Bar -->
        <StatusBar Grid.Row="1" Background="#FFF0F0F0" Height="24">
            <StatusBarItem>
                <TextBlock Name="statusLabel" 
                           Text="Ready" 
                           FontSize="10"
                           Foreground="#FF616161"/>
            </StatusBarItem>
            <Separator Margin="10,0"/>
            <StatusBarItem>
                <TextBlock Name="statusDriveInfo" 
                           Text="" 
                           FontSize="10"
                           Foreground="#FF616161"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
'@
#endregion

#region 3_Data_Models_and_Types
class USBDrive {
    [ValidateNotNullOrEmpty()]
    [string]$DisplayName
    
    [ValidateNotNullOrEmpty()]
    [string]$DeviceID
    
    [string]$Guid
    
    [string]$UUID
    
    [string]$Model
    
    [string]$DriveType
    
    [double]$SizeGB
    
    [datetime]$LastUpdated
    
    USBDrive([string]$displayName, [string]$deviceId, [string]$guid, [string]$uuid, [string]$model, [string]$driveType, [double]$sizeGB) {
        $this.DisplayName = $displayName
        $this.DeviceID = $deviceId
        if ([string]::IsNullOrWhiteSpace($guid)) { $this.Guid = "N/A" } else { $this.Guid = $guid }
        if ([string]::IsNullOrWhiteSpace($uuid)) { $this.UUID = "N/A" } else { $this.UUID = $uuid }
        if ([string]::IsNullOrWhiteSpace($model)) { $this.Model = "Unknown" } else { $this.Model = $model }
        $this.DriveType = $driveType
        $this.SizeGB = $sizeGB
        $this.LastUpdated = [datetime]::Now
    }
}

# Global state management
$script:driveCache = [List[USBDrive]]::new()
$script:eventWatcher = $null
$script:debounceTimer = $null
#endregion

#region 4_Logging_and_Utility_Functions
function Write-ActivityLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Message,
        [Parameter(Mandatory = $false)][ValidateSet("Info", "Warning", "Error", "Debug")][string]$Level = "Info"
    )
    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        if (Test-Path -Path $LOG_PATH -PathType Container) {
            Add-Content -Path $LOG_FILE -Value $logEntry -ErrorAction SilentlyContinue
        }
    }
    catch {
        # Silent fail }
    }
}
#endregion

#region 5_USB_Detection_Logic
function Update-DriveList {
    <#
    .SYNOPSIS
        Updates the list of available USB drives using a robust, hardware-level detection method.
    #>
    [CmdletBinding()]
    param()
    
    try {
        $script:btnRefresh.IsEnabled = $false
        $script:progressBar.Visibility = "Visible"
        $script:statusLabel.Text = "Scanning for USB devices..."
        Write-ActivityLog -Message "Starting hardware-level USB device scan"
        
        $script:driveCache.Clear()
        $script:cmbDrives.Items.Clear()
        
        $usbDiskDeviceIDs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $systemDrive = $env:SystemDrive

        # Primary Method: Trace WMI associations from USB hubs down to disk drives.
        Write-ActivityLog -Message "Stage 1: Performing deep WMI association scan from USB hubs..."
        $usbHubs = Get-CimInstance -ClassName Win32_USBHub -ErrorAction SilentlyContinue
        if ($usbHubs) {
            foreach ($hub in $usbHubs) {
                try {
                    $assocQuery = "ASSOCIATORS OF {Win32_USBHub.DeviceID='$($hub.DeviceID)'} WHERE AssocClass=Win32_USBControllerDevice"
                    $connectedPnpEntities = Get-CimInstance -Query $assocQuery -ErrorAction SilentlyContinue
                    
                    foreach ($pnpEntity in $connectedPnpEntities) {
                        $pnpAssocQuery = "ASSOCIATORS OF {Win32_PnPEntity.DeviceID='$($pnpEntity.DeviceID)'}"
                        $pnpAssociated = Get-CimInstance -Query $pnpAssocQuery -ErrorAction SilentlyContinue
                        $diskDrives = $pnpAssociated | Where-Object { $_.CIMClassName -eq 'Win32_DiskDrive' }
                        
                        foreach ($disk in $diskDrives) {
                            $usbDiskDeviceIDs.Add($disk.DeviceID) | Out-Null
                        }
                    }
                }
                catch {
                    Write-ActivityLog -Message "Error traversing hub $($hub.DeviceID): $_" -Level Warning
                }
            }
        }
        
        # Fallback Method: Use Get-PhysicalDisk BusType for modern systems
        Write-ActivityLog -Message "Stage 2: Checking PhysicalDisk BusType as a fallback..."
        $physicalDisks = Get-PhysicalDisk -ErrorAction SilentlyContinue | Where-Object { $_.BusType -eq 'USB' }
        if ($physicalDisks) {
            foreach ($pdisk in $physicalDisks) {
                $wmiDisk = Get-CimInstance -ClassName Win32_DiskDrive -Filter "Index = $($pdisk.DeviceID)" -ErrorAction SilentlyContinue
                if ($wmiDisk) {
                    $usbDiskDeviceIDs.Add($wmiDisk.DeviceID) | Out-Null
                }
            }
        }
        
        Write-ActivityLog -Message "Found $($usbDiskDeviceIDs.Count) unique USB disk(s). Processing..."
        
        foreach ($deviceID in $usbDiskDeviceIDs) {
            try {
                $disk = Get-CimInstance -ClassName Win32_DiskDrive -Filter "DeviceID = '$($deviceID.Replace('\', '\\'))'" -ErrorAction Stop

                $partitions = Get-Partition -DiskNumber $disk.Index -ErrorAction SilentlyContinue
                if (-not $partitions) { continue }

                foreach ($partition in $partitions) {
                    if (-not $partition.DriveLetter) { continue }
                    $driveLetter = "$($partition.DriveLetter):"
                    if ($driveLetter -eq $systemDrive) { continue }

                    $volume = Get-CimInstance -ClassName Win32_Volume -Filter "DriveLetter = '$driveLetter'" -ErrorAction SilentlyContinue
                    if (-not $volume) { continue }
                    
                    $guid = if ($volume.DeviceID -match '\{([A-Fa-f0-9\-]+)\}') { $matches[1].ToUpper() } else { "N/A" }
                    $uuid = if ($disk.SerialNumber) { $disk.SerialNumber.Trim() } else { "N/A" }
                    $model = if ([string]::IsNullOrWhiteSpace($disk.Model)) { "USB Storage Device" } else { $disk.Model }
                    
                    $logicalDisk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID = '$driveLetter'" -ErrorAction SilentlyContinue
                    $label = if ($logicalDisk) { $logicalDisk.VolumeName } else { "No Label" }
                    $sizeGB = if ($logicalDisk) { [math]::Round($logicalDisk.Size / 1GB, 2) } else { [math]::Round($disk.Size / 1GB, 2) }
                    $displayName = if ([string]::IsNullOrWhiteSpace($label) -or $label -eq "No Label") { "$model ($driveLetter)" } else { "$label ($driveLetter)" }

                    $driveObject = [USBDrive]::new($displayName, $volume.DeviceID, $guid, $uuid, $model, "USB Device", $sizeGB)
                    $script:driveCache.Add($driveObject)
                    $script:cmbDrives.Items.Add($displayName) | Out-Null
                    Write-ActivityLog -Message "Added '$displayName' [GUID: $guid, UUID: $uuid]"
                }
            }
            catch {
                Write-ActivityLog -Message "Failed to process disk with DeviceID '$deviceID': $_" -Level Error
            }
        }

        if ($script:driveCache.Count -gt 0) {
            if ($script:cmbDrives.Items.Count -gt 0) { $script:cmbDrives.SelectedIndex = 0 }
            $script:statusLabel.Text = "Ready - $($script:driveCache.Count) USB device(s) found"
        }
        else {
            $script:statusLabel.Text = "No USB devices detected"
        }
        return $true
    }
    catch {
        $errorMessage = "USB detection failed: $_"
        Write-ActivityLog -Message $errorMessage -Level Error
        $script:statusLabel.Text = "Error detecting USB devices"
        return $false
    }
    finally {
        $script:btnRefresh.IsEnabled = $true
        $script:progressBar.Visibility = "Collapsed"
    }
}
#endregion

#region 6_GUID_UUID_Retrieval_Logic
function Get-SelectedDriveIdentifiers {
    [CmdletBinding()]
    param()
    
    try {
        if ($script:cmbDrives.SelectedIndex -lt 0) {
            $script:txtGuid.Text = ""; $script:txtUuid.Text = ""; $script:statusDriveInfo.Text = ""
            $script:statusLabel.Text = "Select a drive from the list."
            return
        }
        
        $selectedDrive = $script:driveCache[$script:cmbDrives.SelectedIndex]
        if ($null -eq $selectedDrive) { throw "Selected drive not found in cache" }
        
        $cacheAge = [datetime]::Now - $selectedDrive.LastUpdated
        if ($cacheAge.TotalMinutes -gt $CACHE_EXPIRY_MINUTES) {
            Write-ActivityLog -Message "Cache expired for $($selectedDrive.DisplayName), refreshing..."
            $selectedDisplayName = $selectedDrive.DisplayName
            if (Update-DriveList) {
                $script:cmbDrives.SelectedItem = $selectedDisplayName
                if ($script:cmbDrives.SelectedIndex -ge 0) {
                    $selectedDrive = $script:driveCache[$script:cmbDrives.SelectedIndex]
                }
                else {
                    $script:statusLabel.Text = "Drive list updated. Please re-select."
                    $script:txtGuid.Text = ""; $script:txtUuid.Text = ""; $script:statusDriveInfo.Text = ""
                    return
                }
            }
        }
        
        $script:txtGuid.Text = $selectedDrive.Guid
        $script:txtUuid.Text = $selectedDrive.UUID
        $script:statusDriveInfo.Text = "$($selectedDrive.Model) | $($selectedDrive.SizeGB) GB | $($selectedDrive.DriveType)"
        $script:statusLabel.Text = "Identifiers retrieved"
    }
    catch {
        $errorMessage = "Failed to retrieve identifiers: $_"
        Write-ActivityLog -Message $errorMessage -Level Error
        $script:statusLabel.Text = $_.Exception.Message
    }
}

function Copy-Identifier {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateSet("GUID", "UUID", "All")][string]$Type
    )
    
    try {
        $textToCopy = ""
        switch ($Type) {
            "GUID" {
                $textToCopy = $script:txtGuid.Text
                if ([string]::IsNullOrWhiteSpace($textToCopy) -or $textToCopy -eq "N/A") { $script:statusLabel.Text = "No GUID to copy"; return }
            }
            "UUID" {
                $textToCopy = $script:txtUuid.Text
                if ([string]::IsNullOrWhiteSpace($textToCopy) -or $textToCopy -eq "N/A") { $script:statusLabel.Text = "No UUID to copy"; return }
            }
            "All" {
                if ($script:cmbDrives.SelectedIndex -ge 0) {
                    $selectedDrive = $script:driveCache[$script:cmbDrives.SelectedIndex]
                    $textToCopy = @"
Drive Information
=================
Drive: $($selectedDrive.DisplayName)
Model: $($selectedDrive.Model)
Size: $($selectedDrive.SizeGB) GB
Type: $($selectedDrive.DriveType)
GUID: $($selectedDrive.Guid)
UUID: $($selectedDrive.UUID)
Device ID: $($selectedDrive.DeviceID)
"@
                }
                else { $script:statusLabel.Text = "No drive selected"; return }
            }
        }
        [System.Windows.Forms.Clipboard]::SetText($textToCopy)
        $script:statusLabel.Text = "$Type copied to clipboard"
        Write-ActivityLog -Message "Copied $Type to clipboard"
    }
    catch {
        $errorMessage = "Copy failed: $_"
        Write-ActivityLog -Message $errorMessage -Level Error
        $script:statusLabel.Text = "Failed to copy"
    }
}
#endregion

#region 7_UI_Control_Functions
function Initialize-UserInterface {
    [CmdletBinding()]
    param()
    
    try {
        $xamlReader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xamlDefinition))
        $script:window = [Windows.Markup.XamlReader]::Load($xamlReader)
        
        $controlsToFind = @(
            "cmbDrives", "btnRefresh", "btnCopyAll", "btnCopyGuid", "btnCopyUuid",
            "txtGuid", "txtUuid", "progressBar", "statusLabel", "statusDriveInfo"
        )
        foreach ($controlName in $controlsToFind) {
            $control = $script:window.FindName($controlName)
            if ($null -eq $control) { throw "Failed to find required control: $controlName" }
            Set-Variable -Name $controlName -Value $control -Scope "script"
        }
        
        Write-ActivityLog -Message "User interface initialized successfully"
        return $script:window
    }
    catch {
        Write-ActivityLog -Message "Failed to initialize UI: $_" -Level Error
        throw
    }
}
#endregion

#region 8_Event_Handlers
function Register-EventHandlers {
    [CmdletBinding()]
    param()
    
    try {
        $script:btnRefresh.Add_Click({ if (Update-DriveList) { Get-SelectedDriveIdentifiers } })
        $script:btnCopyAll.Add_Click({ Copy-Identifier -Type "All" })
        $script:btnCopyGuid.Add_Click({ Copy-Identifier -Type "GUID" })
        $script:btnCopyUuid.Add_Click({ Copy-Identifier -Type "UUID" })
        $script:cmbDrives.Add_SelectionChanged({ Get-SelectedDriveIdentifiers })
        
        $script:window.Add_Loaded({
                try {
                    $script:debounceTimer = New-Object System.Windows.Threading.DispatcherTimer
                    $script:debounceTimer.Interval = [TimeSpan]::FromMilliseconds($DEBOUNCE_INTERVAL_MS)
                    $script:debounceTimer.Add_Tick({
                            $script:debounceTimer.Stop()
                            $selectedItem = $script:cmbDrives.SelectedItem
                            Update-DriveList | Out-Null
                            if ($script:cmbDrives.Items.Contains($selectedItem)) {
                                $script:cmbDrives.SelectedItem = $selectedItem
                            }
                            if ($script:cmbDrives.SelectedIndex -lt 0 -and $script:cmbDrives.Items.Count -gt 0) {
                                $script:cmbDrives.SelectedIndex = 0
                            }
                            Get-SelectedDriveIdentifiers
                        })
                
                    Register-DriveWatcher | Out-Null
                    if (Update-DriveList) { Get-SelectedDriveIdentifiers }
                }
                catch { Write-ActivityLog -Message "Window load error: $_" -Level Error }
            })
        
        $script:window.Add_Closing({ try { Remove-DriveWatcher } catch {} })
        Write-ActivityLog -Message "Event handlers registered successfully"
    }
    catch {
        Write-ActivityLog -Message "Failed to register event handlers: $_" -Level Error
        throw
    }
}

function Register-DriveWatcher {
    [CmdletBinding()]
    param()
    
    try {
        $wqlQuery = "SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2 OR EventType = 3"
        $script:eventWatcher = Register-CimIndicationEvent -Query $wqlQuery -Action {
            if ($null -ne $script:debounceTimer) {
                $script:debounceTimer.Stop()
                $script:debounceTimer.Start()
            }
        } -ErrorAction Stop
        Write-ActivityLog -Message "Drive watcher registered"
    }
    catch {
        Write-ActivityLog -Message "Failed to register drive watcher: $_" -Level Error
    }
}

function Remove-DriveWatcher {
    [CmdletBinding()]
    param()
    
    try {
        if ($null -ne $script:eventWatcher) {
            Unregister-Event -SourceIdentifier $script:eventWatcher.SourceIdentifier -ErrorAction SilentlyContinue
            Remove-Job -Job $script:eventWatcher -ErrorAction SilentlyContinue
            $script:eventWatcher = $null
        }
        if ($null -ne $script:debounceTimer) {
            $script:debounceTimer.Stop()
            $script:debounceTimer = $null
        }
        Write-ActivityLog -Message "Drive watcher removed"
    }
    catch {
        Write-ActivityLog -Message "Error removing drive watcher: $_" -Level Warning
    }
}
#endregion

#region 9_Application_Lifecycle
function Start-Application {
    [CmdletBinding()]
    param()
    
    try {
        Write-ActivityLog -Message ("=" * 50)
        Write-ActivityLog -Message "Application starting - Version 6.0"
        Write-ActivityLog -Message ("=" * 50)
        
        [void](Initialize-UserInterface)
        [void](Register-EventHandlers)
        
        [void]$script:window.ShowDialog()
        
        Write-ActivityLog -Message "Application exited"
    }
    catch {
        $errorMessage = "Application failed to start: $_"
        Write-ActivityLog -Message $errorMessage -Level Error
        [System.Windows.MessageBox]::Show(
            "Failed to start USB GUID/UUID Retriver:`n`n$_", "Application Error",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        exit 1
    }
    finally {
        Remove-DriveWatcher
        Write-ActivityLog -Message "Application shutdown complete"
    }
}
#endregion

#region 10_Main_Entry_Point
Start-Application
#endregion