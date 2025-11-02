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
# Self-elevate the script if required
param(
    [switch]$Unattended,
    [switch]$AutoReboot
)

if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
     $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
     Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
     Exit
    } 
}

# --- Start Logging ---
# Create a log file in the same directory as the script
$LogFile = Join-Path -Path $PSScriptRoot -ChildPath "Windows-Fix-Up_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
Start-Transcript -Path $LogFile

$ProgressPreference = 'SilentlyContinue'
$LineBreakCharacter = "-"
1..$($Host.UI.RawUI.BufferSize.Width) | ForEach-Object {
    $LineBreak += $LineBreakCharacter
}

function Get-TimeStamp {
    return (Get-Date -Format "[MM/dd/yyyy|HH:mm:ss]")
}

function Write-HostTimestamp {
    param (
        [string]$Message,
        [consolecolor]$ForegroundColor = $(try {((Get-Host).ui.rawui.ForegroundColor)} catch {"White"})
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
Write-HostTimestamp "Running Windows Fix Up on $($env:ComputerName)..."  -Foreground Yellow

if ($Unattended) {
    Write-HostTimestamp "Running in Unattended mode. User prompts will be skipped." -ForegroundColor Cyan
    if ($AutoReboot) {
        $Script:autoRestart = 'y'
    }
} else {
    # Prompt user for confirmation
    Invoke-Task -Description "This script can take several hours to complete and maybe required to be ran twice in order to fix common issues." -ScriptBlock {
        Write-Host "It will delete files using Microsoft's Disk Cleanup (excluding the Downloads folder)."
        Write-Host "It will install a third-party module ('PSWindowsUpdate') to manage Windows Updates through PowerShell."
        Write-Host "It will reset some Windows settings (e.g., network settings) to their defaults."
        Write-Host "Restart is required for all steps of the Fix Up to complete."
        Write-Host ""
        if ($AutoReboot) { $Script:autoRestart = 'y' } else { $Script:autoRestart = Read-Host -Prompt "Do you want to automatically restart when the script is finished? (Y/N)" }
    }
}

# Windows to run the System File Checker utility - Part 1
Invoke-Task -Description "Scanning and repairing system files with SFC..." -ScriptBlock {
    sfc /scannow
}

# Reset winmgmt Repository
Invoke-Task -Description "Resetting and rebuilding the WMI repository..." -ScriptBlock {
    winmgmt.exe /resetrepository
}

# Clean up WinSxS
Invoke-Task -Description "Cleaning up the component store (WinSxS) with DISM..." -ScriptBlock {
    DISM /Online /Cleanup-Image /StartComponentCleanup /ResetBase
}

Invoke-Task -Description "Checking system image health with DISM (/CheckHealth)..." -ScriptBlock {
    DISM /Online /Cleanup-Image /CheckHealth
}

Invoke-Task -Description "Scanning for component store corruption with DISM (/ScanHealth)..." -ScriptBlock {
    DISM /Online /Cleanup-Image /ScanHealth
}

Invoke-Task -Description "Repairing the system image with DISM (/RestoreHealth)..." -ScriptBlock {
    # NOTE: While the previous steps seem redundant, there have been certain fixes deployed by Microsoft that require /ScanHealth and /ScanHealth to run first before fixes are applied by /RestoreHealth
    DISM /Online /Cleanup-Image /RestoreHealth
}

# Windows to run the System File Checker utility - Part 2
Invoke-Task -Description "Running SFC again to fix any remaining issues..." -ScriptBlock {
    sfc /scannow
}

# Run check disk on need startup
Invoke-Task -Description "Scheduling a disk check (CHKDSK) for the next restart..." -ScriptBlock {
    'y' | CHKDSK C: /F /V /R /offlinescanandfix
}

# Run Disk Cleanup
Invoke-Task -Description "Configuring and running Disk Cleanup for all categories..." -ScriptBlock {
    # Set registry keys to select all items for Disk Cleanup
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
    Get-ChildItem -Path $RegPath | ForEach-Object {
        if ($_.PSChildName -ne 'DownloadsFolder') {
            Set-ItemProperty -Path $_.PSPath -Name "StateFlags0001" -Value 2 -ErrorAction SilentlyContinue
        }
    }
    # Run Disk Cleanup with the configured settings
    Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -WindowStyle Hidden
    # Due to cleanmgr commonly getting stuck, the following has been added as a workaround
    # Check for a Window Title, if it exists, Disk Cleanup is still running
    do {
        Start-Sleep -Seconds 5
    } Until ([string]::IsNullOrWhiteSpace((Get-Process -Name cleanmgr -ErrorAction SilentlyContinue).MainWindowTitle))
    Write-Host "Disk Cleanup is almost finished..."
    # Second checks as dismhost maybe cleaning things as well
    do {
        $cleanmgrTime = (Get-Process -Name cleanmgr -ErrorAction SilentlyContinue).TotalProcessorTime
        $dismHostTime = (Get-Process -Name dismhost -ErrorAction SilentlyContinue).TotalProcessorTime
        if ($cleanmgrTime -or $dismHostTime){
            Start-Sleep -Seconds 30
        }
    } Until (($cleanmgrTime -eq (Get-Process -Name cleanmgr -ErrorAction SilentlyContinue).TotalProcessorTime) -and ($dismHostTime -eq (Get-Process -Name dismhost -ErrorAction SilentlyContinue).TotalProcessorTime))
    Stop-Process -Name cleanmgr -Force
}


# Install Windows Update Module
Invoke-Task -Description "Installing the 'PSWindowsUpdate' PowerShell module..." -ScriptBlock {
    Install-PackageProvider -Name NuGet -Force -ForceBootstrap
    Install-Module PSWindowsUpdate -Force -Confirm:$false
}

# Reset Windows Update Services
Invoke-Task -Description "Resetting Windows Update components..." -ScriptBlock {
    Reset-WUComponents
}

# Reset Windows/Microsoft Store
Invoke-Task -Description "The Microsoft Store cache will now be cleared and the application will be restarted to resolve any download or launch-related issues." -ScriptBlock {
    Start-Process -FilePath "wsreset.exe" -ArgumentList "-i" -Wait
}

# Invoke Windows Store Updates
Invoke-Task -Description "Invoking Windows Store app updates..." -ScriptBlock {
    Get-CimInstance -Namespace "Root\cimv2\mdm\dmmap" -ClassName "MDM_EnterpriseModernAppManagement_AppManagement01" | Invoke-CimMethod -MethodName UpdateScanMethod | Out-Null
}

# Reinstall AppX packages
Invoke-Task -Description "Re-registering all Windows AppX packages for all users..." -ScriptBlock {
    Write-HostTimestamp "This may take some time..." -ForegroundColor Yellow
    Get-AppxPackage -AllUsers | ForEach-Object {
        try {
            Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppxManifest.xml" -ErrorAction Stop
        } catch {
            # Keep the red text down
        }
    }
}

# Install Windows Updates
Invoke-Task -Description "Checking for and installing Windows Updates..." -ScriptBlock {
    Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot -MicrosoftUpdate
}

# Install latest version of apps
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-HostTimestamp "Winget is not installed or not in PATH. Skipping Winget upgrades." -ForegroundColor Yellow
} else {
    Invoke-Task -Description "Upgrading all applications with Winget..." -ScriptBlock {
        winget upgrade --silent --all --accept-package-agreements --accept-source-agreements --force
        Write-HostTimestamp "Re-running again in case we ran into problems in the last go around."
        winget upgrade --silent --all --accept-package-agreements --accept-source-agreements --force
    }
}

# Resetting network stack
Invoke-Task -Description "Resetting network adapters (Winsock, TCP/IP, DNS cache, IP release/renew)..." -ScriptBlock {
    netsh winsock reset
    netsh int ip reset
    ipconfig /release
    ipconfig /renew
    ipconfig /flushdns
}

# Done, restart when necessary
Write-HostTimestamp "Windows Fix Up completed!" -Foreground Green
Write-Host "A restart is required to complete the disk check (CHKDSK)."
if ($autoRestart -in @('y', 'yes')) {
    (60..1) | ForEach-Object {
        if ($_ -lt 10){
            Write-HostTimestamp "Restart in $_ $(if ($_ -eq 1){"second"}else{"seconds"})" -ForegroundColor Yellow
        } else {
            if ($_ % 10 -eq 0) {
            Write-HostTimestamp "Restart in $_ seconds"
            }
        }
        Start-Sleep 1
    }
    shutdown.exe -r -t 5 -c "Restarting to finish fix up..."
} else {
    Write-HostTimestamp "Restart not initiated. Please remember to restart your computer manually to complete the repairs." -ForegroundColor Yellow
    if (-not $Unattended) {
        Read-Host -Prompt "Close window or press enter to exit."
    }
}

# Stop logging
Stop-Transcript
# --- End Logging ---