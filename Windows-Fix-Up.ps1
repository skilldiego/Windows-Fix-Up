# --- SCRIPT OVERVIEW ---
# This script automates a sequence of common Windows repair and maintenance tasks. It is designed to be a comprehensive
# "fix-it" tool for resolving system instability, update failures, application problems, and file system corruption.
# -------------------------------------------------
# How to Run This Script:
# 1.  Open PowerShell as an Administrator: Right-click your Start Menu and select "Terminal (Admin)" or search for "Windows PowerShell" and run as administrator.
# 2.  Enable Script Execution (if needed): If this is your first time running a PowerShell script, execute:
#     Set-ExecutionPolicy Bypass -Force
# 3.  Save the Script: Save the entire contents of this file as "windows-fix-up.ps1" (e.g., using Notepad).
# 4.  Run the Script: Right-click the saved "windows-fix-up.ps1" file and select "Run with PowerShell".
# 5.  Revert Execution Policy (Optional): For security, you can revert the execution policy after the script completes by running:
#     Set-ExecutionPolicy Default -Force
# -------------------------------------------------
# Parameters for the script
param(
    [switch]$Unattended, # Runs the script without any user prompts. It will not ask for confirmation to start.
    [switch]$AutoReboot, # Automatically configures the script to restart the computer upon completion.
    [switch]$ResetWMI    # Forces a rebuild of the WMI repository without attempting to salvage it first.
)

# Self-elevate the script if required
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $ArgumentList = @("-File", "`"$($MyInvocation.MyCommand.Path)`"")
        # Re-add any passed parameters
        foreach ($Parameter in $PSBoundParameters.Keys) {
            $ArgumentList += "-$Parameter"
        }

        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $ArgumentList
        exit
    } 
}

# --- Start Logging ---
# Create a log file in the same directory as the script
$LogFile = Join-Path -Path $PSScriptRoot -ChildPath "Windows-Fix-Up_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
Start-Transcript -Path $LogFile | Out-Null

$ProgressPreference = 'SilentlyContinue'
$LineBreakCharacter = '-'
$LineBreak = $null
1..$($Host.UI.RawUI.BufferSize.Width) | ForEach-Object {
    $LineBreak += $LineBreakCharacter
}

function Get-TimeStamp {
    return (Get-Date -Format '[MM/dd/yyyy|HH:mm:ss]')
}

function Write-HostTimestamp {
    param (
        [string]$Message,
        [consolecolor]$ForegroundColor = $(try { ((Get-Host).ui.rawui.ForegroundColor) } catch { 'White' })
    )

    # Get the current timestamp and combine it with the user's message.
    # The output is then sent to the console using Write-Host with the specified color.
    Write-Host "$(Get-TimeStamp) $Message" -ForegroundColor $ForegroundColor
}

function Invoke-Task {
    param(
        [string]$Description,
        [scriptblock]$ScriptBlock
    )

    Write-HostTimestamp $Description
    & $ScriptBlock
    Write-Host $LineBreak
}

# Start of the Fix Up Script
Write-HostTimestamp "Running Windows Fix Up on $($env:ComputerName)..." -Foreground Yellow
Write-Host $LineBreak

if ($Unattended) {
    Write-HostTimestamp 'Running in Unattended mode. User prompts will be skipped.' -ForegroundColor Cyan
    if ($AutoReboot) {
        $Script:AutoRestart = 'y'
    }
}
else {
    # Prompt user for confirmation
    Invoke-Task -Description 'Can take SEVERAL hours to complete and maybe required to be ran twice to completely fix issues.' -ScriptBlock {
        Write-Host 'NOTE: This script will reset some Windows components to defaults and run Disk Cleanup clearing old files.' -ForegroundColor Yellow
        Write-Host 'Please ensure you have backed up any important files...'
        Write-Host
        if ($AutoReboot) { $Script:AutoRestart = 'y' } else { $Script:AutoRestart = Read-Host -Prompt 'Do you want to automatically restart when the script is finished? (Y/N)' }
    }
}

# Change directory to System32 in case path is not set correctly
try {
    $System32Path = "$env:windir\System32"
    Set-Location -Path $System32Path -ErrorAction Stop -ErrorVariable NoSystem32
}
catch {
    $System32Path = 'C:\Windows\System32'
    Set-Location -Path $System32Path -ErrorAction SilentlyContinue -ErrorVariable NoSystem32
}

# If we can't set location to System32, we have some huge problems
if ($NoSystem32) {
    Write-HostTimestamp 'STOPPING SCRIPT: Unable to change directory to System32. This maybe an issue with the script or a major issue with Windows system that this script cannot fix.' -ForegroundColor Red
    Start-Sleep -Seconds 10; exit 1
}

# Find drive letter of Windows
$WindowsDriveLetter = $System32Path.Substring(0, 2)
if (-not (Test-Path $WindowsDriveLetter)) {
    Write-HostTimestamp "STOPPING SCRIPT: Unable to find $WindowsDriveLetter" -ForegroundColor Red
    Start-Sleep -Seconds 10; exit 1
}

# Verify and Salvage WMI Repository
Invoke-Task -Description 'Checking and repairing the WMI repository...' -ScriptBlock {
    try {
        $wmiService = Get-Service -Name winmgmt -ErrorAction Stop
        if ($wmiService.StartType -ne 'Automatic') {
            Write-HostTimestamp "Setting WMI service (winmgmt) to Automatic startup."
            Set-Service -Name winmgmt -StartupType Automatic
        }
        if ($wmiService.Status -ne 'Running') {
            Write-HostTimestamp "Starting WMI service (winmgmt)."
            Start-Service -Name winmgmt
            Start-Sleep -Seconds 5 # Give it a moment to start up
        }
    }
    catch {
        Write-HostTimestamp "Could not find or manage the WMI service (winmgmt). WMI checks might fail." -ForegroundColor Red
    }

    function Start-BuildWMI {
        Write-Host "Rebuilding WMI repository."
        $mofcompPath = "$System32Path\wbem\mofcomp.exe"
        if (Test-Path $mofcompPath) {
            winmgmt.exe /resetrepository
            Write-HostTimestamp "WMI repository is rebuilding. This can take some time."
            Get-ChildItem "$System32Path\wbem\*.mof" -File | Where-Object { $_.Name -notmatch 'uninstall|remove' } | ForEach-Object {
                Write-Host "Processing $($_.FullName)"
                Start-Process -FilePath $mofcompPath -ArgumentList $_.FullName -Wait -WindowStyle Hidden
            }
            Get-ChildItem -Path "$System32Path\wbem\en-us\*.mfl" -File | Where-Object { $_.Name -notmatch 'uninstall|remove' } | ForEach-Object {
                Write-Host "Processing $($_.FullName)"
                Start-Process -FilePath $mofcompPath -ArgumentList $_.FullName -Wait -WindowStyle Hidden
            }
            Write-HostTimestamp "WMI repository has been rebuilt."
        }
        else {
            Write-Host "Unable to start WMI repository reset. Missing mofcomp.exe." -ForegroundColor Red
        }
    }
    
    if ($ResetWMI) {
        Write-HostTimestamp "-ResetWMI switch detected. Forcing WMI repository rebuild." -ForegroundColor Cyan
        Start-BuildWMI
    }
    else {
        $WinMgmtOutput = winmgmt.exe /verifyrepository
        if ($WinMgmtOutput -eq 'WMI repository is consistent') {
            Write-Host 'WMI repository appears to be healthy.'
        }
        else {
            Write-Host 'WMI repository may have issues. Trying to salvage it.'
            $null = winmgmt.exe /salvagerepository # Run salvage twice for good measure
            $WinMgmtOutput = winmgmt.exe /salvagerepository
            if ($WinMgmtOutput -eq 'WMI repository is consistent') {
                Write-Host 'WMI repository has been salvaged.'
            }
            else {
                Write-Host 'WMI repository salvage was unsuccessful.'
                Start-BuildWMI
            }
        }
    }
}

# Windows to run the System File Checker utility - Part 1
Invoke-Task -Description 'Scanning and repairing system files with SFC...' -ScriptBlock {
    sfc.exe /scannow
}

# Clean up WinSxS
Invoke-Task -Description 'Cleaning up the component store (WinSxS) with DISM...' -ScriptBlock {
    DISM.exe /Online /Cleanup-Image /StartComponentCleanup /ResetBase
}

Invoke-Task -Description 'Checking system image health with DISM (/CheckHealth)...' -ScriptBlock {
    DISM.exe /Online /Cleanup-Image /CheckHealth
}

Invoke-Task -Description 'Scanning for component store corruption with DISM (/ScanHealth)...' -ScriptBlock {
    DISM.exe /Online /Cleanup-Image /ScanHealth
}

Invoke-Task -Description 'Repairing the system image with DISM (/RestoreHealth)...' -ScriptBlock {
    # NOTE: While the previous steps seem redundant, there have been certain fixes deployed by Microsoft that require /ScanHealth and /ScanHealth to run first before fixes are applied by /RestoreHealth
    DISM.exe /Online /Cleanup-Image /RestoreHealth
}

# Windows to run the System File Checker utility - Part 2
Invoke-Task -Description 'Running SFC again to fix any remaining issues...' -ScriptBlock {
    sfc.exe /scannow
}

# Run Disk Cleanup
Invoke-Task -Description 'Configuring and running Disk Cleanup for all categories...' -ScriptBlock {
    # Set registry keys to select all items for Disk Cleanup
    $RegPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches'
    Get-ChildItem -Path $RegPath | ForEach-Object {
        if ($_.PSChildName -ne 'DownloadsFolder') {
            Set-ItemProperty -Path $_.PSPath -Name 'StateFlags0001' -Value 2 -ErrorAction SilentlyContinue
        }
    }
    # Run Disk Cleanup with the configured settings
    Start-Process -FilePath 'cleanmgr.exe' -ArgumentList '/sagerun:1' -WindowStyle Hidden
    # Due to cleanmgr commonly getting stuck, the following has been added as a workaround
    # Check to see if cleanmgr is doing anything
    do {
        $CleanmgrTime = (Get-Process -Name cleanmgr -ErrorAction SilentlyContinue).TotalProcessorTime
        $DismHostTime = (Get-Process -Name dismhost -ErrorAction SilentlyContinue).TotalProcessorTime
        if ($CleanmgrTime -or $DismHostTime) {
            Start-Sleep -Seconds 30
        }
    } until (($CleanmgrTime -eq (Get-Process -Name cleanmgr -ErrorAction SilentlyContinue).TotalProcessorTime) -and ($DismHostTime -eq (Get-Process -Name dismhost -ErrorAction SilentlyContinue).TotalProcessorTime))
    Stop-Process -Name cleanmgr -Force
}

# Install Windows Update Module
Invoke-Task -Description "Installing the 'PSWindowsUpdate' PowerShell module..." -ScriptBlock {
    if (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue) {
        Write-HostTimestamp "NuGet package provider is installed. Updating it..."
    } else {
        Write-HostTimestamp "NuGet package provider needs to be installed..."
        Install-PackageProvider -Name NuGet -Force -ForceBootstrap
    }
    if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
        Write-HostTimestamp "PSWindowsUpdate module is already installed. Checking for updates..."
        Update-Module -Name PSWindowsUpdate -Force -Confirm:$false
    }
    else {
        Install-Module -Name PSWindowsUpdate -Force -Confirm:$false
    }
    Import-Module -Name PSWindowsUpdate -Force
}

# Reset Windows Update Services
Invoke-Task -Description 'Resetting Windows Update components...' -ScriptBlock {
    if (Get-Command Reset-WUComponents -ErrorAction SilentlyContinue) {
        Reset-WUComponents
    }
    else {
        Write-HostTimestamp "Command 'Reset-WUComponents' not found. Skipping Windows Update component reset." -ForegroundColor Yellow
    }
}

# Reset Windows/Microsoft Store
if (Get-Command wsreset.exe -ErrorAction SilentlyContinue) {
    Invoke-Task -Description 'The Microsoft Store cache will now be cleared and the application will be restarted to resolve any download or launch-related issues.' -ScriptBlock {
        Start-Process -FilePath 'wsreset.exe' -ArgumentList '-i' -Wait -WindowStyle Hidden
    }

    # Invoke Windows Store Updates
    Invoke-Task -Description 'Invoking Windows Store app updates...' -ScriptBlock {
        Get-CimInstance -Namespace 'Root\cimv2\mdm\dmmap' -ClassName 'MDM_EnterpriseModernAppManagement_AppManagement01' | Invoke-CimMethod -MethodName UpdateScanMethod | Out-Null
    }
}
else {
    Write-HostTimestamp 'wsreset.exe not found. Skipping Microsoft Store cache reset.' -ForegroundColor Yellow
}

# Reinstall AppX packages
if (Get-Command Get-AppxPackage -ErrorAction SilentlyContinue) {
    Invoke-Task -Description 'Re-registering all Windows AppX packages for all users...' -ScriptBlock {
        Write-HostTimestamp 'This may take some time...' -ForegroundColor Yellow
        Get-AppxPackage -AllUsers | ForEach-Object {
            try {
                Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppxManifest.xml" -ErrorAction Stop
            }
            catch {
                # Keep the red text down
            }
        }
    }
}
else {
    Write-HostTimestamp 'Get-AppxPackage not found. Skipping AppX package re-registration.' -ForegroundColor Yellow
}

# Install Windows Updates
Invoke-Task -Description 'Checking for and installing Windows Updates...' -ScriptBlock {
    if (Get-Command Get-WindowsUpdate -ErrorAction SilentlyContinue) {
        Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot -MicrosoftUpdate
    }
    else {
        Write-HostTimestamp "Command 'Get-WindowsUpdate' not found. Skipping Windows Update installation." -ForegroundColor Yellow
    }
}

# Install latest version of apps
if ((Get-Command winget.exe -ErrorAction SilentlyContinue) -and (-not ([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem))) {
    Invoke-Task -Description 'Upgrading all applications with Winget...' -ScriptBlock {
        winget.exe upgrade --silent --all --accept-package-agreements --accept-source-agreements --force
        Write-HostTimestamp 'Re-running again in case we ran into problems in the last go around.'
        winget.exe upgrade --silent --all --accept-package-agreements --accept-source-agreements --force
    }
}
elseif ([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem) {
    Write-HostTimestamp 'Script is running as SYSTEM. Skipping Winget upgrades as it requires a user context.' -ForegroundColor Yellow
}
else {
    Write-HostTimestamp 'Winget is not installed or not in PATH. Skipping Winget upgrades.' -ForegroundColor Yellow
}

# Resetting network stack
Invoke-Task -Description 'Resetting network adapters (Winsock, TCP/IP, DNS cache, IP release/renew)...' -ScriptBlock {
    netsh.exe winsock reset
    netsh.exe int ip reset
    ipconfig.exe /release
    ipconfig.exe /renew
    ipconfig.exe /flushdns
}

# Run check disk on need startup
Invoke-Task -Description 'Scheduling a disk check (CHKDSK) for the next restart...' -ScriptBlock {
    'y' | CHKDSK.exe $WindowsDriveLetter /F /V /R /offlinescanandfix
}

# Optimize Windows disk
Invoke-Task -Description "Optimizing drive $WindowsDriveLetter..." -ScriptBlock {
    $DriveLetterNoColon = $WindowsDriveLetter.Trim(':')
    $Disk = $null
    $retries = 5
    while ($retries -gt 0 -and -not $Disk) {
        try {
            $Disk = Get-Partition -DriveLetter $DriveLetterNoColon -ErrorAction Stop | Get-Disk -ErrorAction Stop | Get-PhysicalDisk -ErrorAction Stop
        }
        catch {
            # With the rebuild of the WMI, this may require some waiting
            $null = 'rescan' | diskpart.exe
            Write-HostTimestamp "Could not get physical disk information. Retrying in 60 seconds... ($retries retries remaining)" -ForegroundColor Yellow
            Start-Sleep -Seconds 60
            $retries--
        }
    }

    if ($Disk) {
        if ($Disk.MediaType -eq 'SSD') {
            Write-HostTimestamp 'Drive is an SSD. Performing trim...'
            Optimize-Volume -DriveLetter $DriveLetterNoColon -ReTrim -Verbose -ErrorAction Stop
        }
        elseif ($Disk.MediaType -eq 'HDD') {
            Write-HostTimestamp 'Drive is an HDD. Performing defragmentation...'
            Optimize-Volume -DriveLetter $DriveLetterNoColon -Defrag -Verbose -ErrorAction Stop
        }
        else {
            Write-HostTimestamp 'Drive is unspecified. Skipping disk optimization.' -ForegroundColor Yellow
        }
    } else {
        Write-HostTimestamp "Could not get physical disk information. Skipping disk optimization." -ForegroundColor Yellow
    }
}

# Clear Windows Search Index
Invoke-Task -Description 'Clearing Windows Search Index...' -ScriptBlock {
    try {
        Write-HostTimestamp 'Stopping and disabling the Windows Search service to rebuild the index...'
        Set-Service -Name WSearch -StartupType Disabled -ErrorAction Stop
        Stop-Service -Name WSearch -Force -ErrorAction Stop

        $SearchDbPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows"
        $DbFile = Join-Path -Path $SearchDbPath -ChildPath 'Windows.db'
        $GatherDbFile = Join-Path -Path $SearchDbPath -ChildPath 'Windows-gather.db'

        Write-HostTimestamp 'Deleting Windows Search database files...'
        Remove-Item -Path $DbFile -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $GatherDbFile -Force -ErrorAction SilentlyContinue

        Write-HostTimestamp 'Setting Windows Search service to Automatic and starting it...'
        Set-Service -Name WSearch -StartupType Automatic -ErrorAction Stop
        # Attempt to start the service multiple times if it fails initially
        $maxRetries = 5
        $retryCount = 0
        while ((Get-Service -Name WSearch).Status -ne 'Running' -and $retryCount -lt $maxRetries) {
            $retryCount++
            Write-HostTimestamp "Attempting to start Windows Search service (Attempt $retryCount/$maxRetries)..."
            Start-Service -Name WSearch -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            Start-Sleep -Seconds 10
        }
        if ((Get-Service -Name WSearch).Status -ne 'Running') {
            throw "Failed to start Windows Search service after $maxRetries attempts."
        }
        Write-HostTimestamp 'Windows Search index will be rebuilt in the background.'
    }
    catch {
        Write-HostTimestamp "Could not reset the Windows Search Index. The service may not be installed or is in a bad state. Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Done, restart when necessary
Write-HostTimestamp 'Windows Fix Up completed!' -Foreground Green
Write-Host 'A restart is required to complete the disk check (CHKDSK).'
if ($Script:AutoRestart -in @('y', 'yes')) {
    (60..1) | ForEach-Object {
        if ($_ -lt 10) {
            Write-HostTimestamp "Restart in $_ $(if ($_ -eq 1){'second'}else{'seconds'})" -ForegroundColor Yellow
        }
        else {
            if ($_ % 10 -eq 0) {
                Write-HostTimestamp "Restart in $_ seconds"
            }
        }
        Start-Sleep 1
    }
    shutdown.exe -r -t 5 -c 'Restarting to finish fix up...'
}
else {
    Write-HostTimestamp 'Restart not initiated. Please remember to restart your computer manually to complete the repairs.' -ForegroundColor Yellow
    if (-not $Unattended) {
        Read-Host -Prompt 'Close window or press enter to exit.'
    }
}

# Stop logging
Stop-Transcript
# --- End Logging ---