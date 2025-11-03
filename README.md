# Windows Fix-Up Script

> [!NOTE]
> This script was born out of the necessity for one IT professional to find a lazy-yet-effective way to fix the recurring computer problems of friends and family. Think of it as the "turn it off and on again" of Windows repair, but on steroids. While it's generally safe and has saved countless hours of phone support, please exercise caution when running it in a production or enterprise environment.

This PowerShell script automates a sequence of common Windows repair and maintenance tasks. It is designed to be a comprehensive "fix-it" tool for resolving system instability, update failures, application problems, and file system corruption.

## ⚠️ Important Warnings

*   **Run as Administrator:** This script requires administrative privileges to perform its tasks. It includes a self-elevation feature to prompt for administrator rights if not already running with them.
*   **Creates a Log File:** The script automatically creates a detailed log file (e.g., `Windows-Fix-Up_2023-10-27_14-30-00.log`) in the same folder where the script is located.
*   **Time-Consuming:** The entire process can take **several hours** to complete, depending on the state of your system.
*   **File Deletion:** The script will run Windows Disk Cleanup and automatically select **all** categories for cleaning, **except for the `Downloads` folder**. Ensure any important files are backed up from temporary locations.
*   **Third-Party Module:** The script will automatically install the `PSWindowsUpdate` module from the PowerShell Gallery to manage Windows Updates.
*   **System Resets:** Several system components will be reset to their default configurations, including:
    *   Windows Management Instrumentation (WMI) repository.
    *   Windows Update components.
    *   Microsoft Store cache.
    *   Network stack (Winsock, TCP/IP).
*   **Restart Required:** A system restart is required to complete the disk check (CHKDSK). The script will ask you at the beginning if you want to restart automatically when it's done.

## How to Run This Script

1.  **Save the Script:** Save the script file as `Windows-Fix-Up.ps1` on your computer.
2.  **Run with PowerShell:** Right-click the `Windows-Fix-Up.ps1` file and select **"Run with PowerShell"**.
3.  **Administrator Prompt:** If you are not already in an administrative session, a User Account Control (UAC) window will appear. Click **Yes** to allow the script to run with administrative privileges.
4.  **Execution Policy (If Needed):** If you encounter an error about script execution being disabled, you may need to change the execution policy. Open a new PowerShell window **as an Administrator** and run the following command:
    ```powershell
    Set-ExecutionPolicy Bypass -Force
    ```
    Then, try running the script again. For security, you can revert this change after the script is finished by running `Set-ExecutionPolicy Default -Force`.
5.  **Confirmation:** The script will display a summary of its actions and ask for your confirmation. Type `Y` and press `Enter` to begin.
    *   When run interactively, it will ask if you want to **automatically restart** when the script is complete. Type `Y` or `N` and press Enter. The script will then proceed with its tasks.

## Command-Line Parameters

The script supports the following optional parameters for automation:

*   `-Unattended`: Runs the script without any user prompts. It will not ask for confirmation to start.
*   `-AutoReboot`: When used with `-Unattended`, this will automatically configure the script to restart the computer upon completion. If used without `-Unattended`, it pre-answers 'Y' to the automatic restart question.
*   `-ResetWMI`: Forces a rebuild of the WMI repository without attempting to salvage it first. This can be useful if you suspect deep-rooted WMI corruption.

Example of an unattended run with automatic reboot:
```powershell
.\Windows-Fix-Up.ps1 -Unattended -AutoReboot -ResetWMI
```

## What the Script Does

The script performs the following actions in sequence to repair and optimize your Windows installation.

1.  **WMI Repository Verification and Repair**
    *   Checks the health of the Windows Management Instrumentation (WMI) repository. If it is found to be inconsistent, the script first attempts to salvage it. If salvaging is unsuccessful, it proceeds to rebuild the repository to resolve issues with system management tools and services.

2.  **System File Checker (SFC)**
    *   Runs `sfc /scannow` to scan for and repair corrupted or missing Windows system files.

3.  **DISM Component Store Cleanup & Repair**
    *   **Cleanup (`/StartComponentCleanup /ResetBase`):** Cleans up and compresses the component store (WinSxS folder) to save disk space.
    *   **Health Check (`/CheckHealth` & `/ScanHealth`):** Scans the Windows component store for corruption.
    *   **Restore Health (`/RestoreHealth`):** Performs repair operations automatically using Windows Update to fix any detected corruption.

4.  **System File Checker (SFC) - Second Pass**
    *   Runs `sfc /scannow` again to address any issues that may have been uncovered or made fixable by the DISM repairs.

5.  **Disk Cleanup**
    *   Automates the Windows Disk Cleanup utility (`cleanmgr.exe`) to remove temporary files, system logs, old update files, and other unnecessary data. **The Downloads folder is explicitly excluded.**
    *   The script includes a monitor to prevent the Disk Cleanup process from getting stuck indefinitely.

6.  **Windows Update Module Installation**
    *   Installs or updates the `PSWindowsUpdate` PowerShell module, which allows for advanced management of Windows Updates via the command line. It also ensures the required `NuGet` package provider is present.

7.  **Windows Update Reset**
    *   Resets the Windows Update components to their default state, which can fix issues with updates failing to download or install.

8.  **Microsoft Store Reset & Update**
    *   Clears the Microsoft Store cache (`wsreset.exe`) to resolve problems with apps not downloading or launching.
    *   Triggers a scan for pending Microsoft Store app updates.

9.  **Re-register Windows Apps**
    *   Attempts to re-register all built-in and installed Microsoft Store (AppX) packages for all users. This can fix issues with modern apps that fail to start or function correctly.

10. **Install Windows Updates**
    *   Uses the `PSWindowsUpdate` module to check for, download, and install all available updates from Microsoft Update.

11. **Upgrade Applications with Winget**
    *   If the Windows Package Manager (`winget`) is available and the script is not running as the SYSTEM account, it will attempt to upgrade all installed applications silently. It runs twice to handle dependencies or failed initial attempts.

12. **Network Stack Reset**
    *   Resets the network configuration to resolve common connectivity issues:
        *   Resets the Winsock Catalog (`netsh winsock reset`).
        *   Resets the TCP/IP stack (`netsh int ip reset`).
        *   Releases and renews the IP address configuration (`ipconfig /release` & `ipconfig /renew`).
        *   Flushes the DNS resolver cache (`ipconfig /flushdns`).

13. **Disk Check (CHKDSK)**
    *   Schedules a comprehensive disk check (`chkdsk /f /r`) to run on the C: drive during the next system restart. This finds and repairs file system errors and scans for bad sectors.

14. **Disk Optimization**
    *   Checks the media type of the system drive.
    *   If it's an SSD, it performs a re-trim operation (`Optimize-Volume -ReTrim`).
    *   If it's an HDD, it performs a defragmentation (`Optimize-Volume -Defrag`).

15. **Windows Search Index Reset**
    *   Stops the Windows Search service, deletes the index database files to clear out corruption, and then restarts the service to allow it to rebuild the index in the background.

16. **Final Restart**
    *   If you agreed to the automatic restart at the beginning or used the `-AutoReboot` parameter, the script will initiate a 60-second countdown before rebooting. Otherwise, it will remind you to restart manually.
