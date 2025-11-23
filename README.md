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
*   `-DisableHibernation`: Disables hibernation and fast startup by running `powercfg.exe /hibernate off`. See [Why Disable Hibernation and Fast Startup?](#why-disable-hibernation-and-fast-startup) for more details.
*   `-DisableBrandBloat`: Disables startup services from common computer manufacturers (e.g., HP, Dell, ASUS, Lenovo, Acer) to reduce background processes.

Example of an unattended run with automatic reboot:
```powershell
.\Windows-Fix-Up.ps1 -Unattended -AutoReboot -ResetWMI -DisableBrandBloat
```

## What the Script Does

The script performs the following actions in sequence to repair and optimize your Windows installation.

1.  **Disable Manufacturer Bloatware (Optional)**
    *   If the `-DisableBrandBloat` parameter is used, the script will find, stop, and disable common services from manufacturers like HP, Dell, ASUS, Lenovo, and Acer. This helps reduce unnecessary background processes.

2.  **WMI Repository Verification and Repair**
    *   Checks the health of the Windows Management Instrumentation (WMI) repository. If it is found to be inconsistent, the script first attempts to salvage it. If salvaging is unsuccessful, it proceeds to rebuild the repository to resolve issues with system management tools and services.

3.  **System File Checker (SFC)**
    *   Runs `sfc /scannow` to scan for and repair corrupted or missing Windows system files.

4.  **DISM Component Store Cleanup & Repair**
    *   **Cleanup (`/StartComponentCleanup /ResetBase`):** Cleans up and compresses the component store (WinSxS folder) to save disk space.
    *   **Health Check (`/CheckHealth` & `/ScanHealth`):** Scans the Windows component store for corruption.
    *   **Restore Health (`/RestoreHealth`):** Performs repair operations automatically using Windows Update to fix any detected corruption.

5.  **System File Checker (SFC) - Second Pass**
    *   Runs `sfc /scannow` again to address any issues that may have been uncovered or made fixable by the DISM repairs.

6.  **Disk Cleanup**
    *   Automates the Windows Disk Cleanup utility (`cleanmgr.exe`) to remove temporary files, system logs, old update files, and other unnecessary data. **The Downloads folder is explicitly excluded.**
    *   The script includes a monitor to prevent the Disk Cleanup process from getting stuck indefinitely.

7.  **Windows Update Module Installation**
    *   Installs or updates the `PSWindowsUpdate` PowerShell module, which allows for advanced management of Windows Updates via the command line. It also ensures the required `NuGet` package provider is present.

8.  **Print Spooler Reset**
    *   Stops the Print Spooler service, clears out any stuck print jobs from the `C:\Windows\System32\spool\PRINTERS` directory, and then restarts the service. This can resolve issues where printers are offline or jobs won't print.

9.  **Windows Update Reset**
    *   Uses the `Reset-WUComponents` command from the `PSWindowsUpdate` module to stop Windows Update services, rename the `SoftwareDistribution` and `catroot2` folders, and re-register necessary DLLs. This resolves many common update failures.

10. **Microsoft Store Reset & Update**
    *   Clears the Microsoft Store cache (`wsreset.exe`) to resolve problems with apps not downloading or launching.
    *   Triggers a scan for pending Microsoft Store app updates.

11. **Re-register Windows Apps**
    *   Attempts to re-register all built-in and installed Microsoft Store (AppX) packages for all users. This can fix issues with modern apps that fail to start or function correctly.

12. **Install Windows Updates**
    *   Uses the `PSWindowsUpdate` module to check for, download, and install all available updates from Microsoft Update.

13. **Upgrade Applications with Winget**
    *   If the Windows Package Manager (`winget`) is available and the script is not running as the SYSTEM account, it will attempt to upgrade all installed applications silently. It runs twice to handle dependencies or failed initial attempts.

14. **Network Stack Reset**
    *   Resets the network configuration to resolve common connectivity issues:
        *   Resets the Winsock Catalog (`netsh winsock reset`).
        *   Resets the TCP/IP stack (`netsh int ip reset`).
        *   Releases and renews the IP address configuration (`ipconfig /release` & `ipconfig /renew`).
        *   Flushes the DNS resolver cache (`ipconfig /flushdns`).

15. **Disable Hibernation (Optional)**
    *   If the `-DisableHibernation` parameter is used, this step will turn off hibernation, delete the `hiberfil.sys` file, and disable Windows Fast Startup.

16. **Disk Check (CHKDSK)**
    *   Schedules a comprehensive disk check (`chkdsk /f /r`) to run on the C: drive during the next system restart. This finds and repairs file system errors and scans for bad sectors.

17. **Disk Optimization**
    *   Checks the media type of the system drive.
    *   If it's an SSD, it performs a re-trim operation (`Optimize-Volume -ReTrim`).
    *   If it's an HDD, it performs a defragmentation (`Optimize-Volume -Defrag`).

18. **Windows Search Index Reset**
    *   Stops and temporarily disables the Windows Search service, deletes the index database files (`Windows.db`) to clear out corruption, and then re-enables and restarts the service to allow it to rebuild the index from scratch in the background.

19. **Final Restart**
    *   If you agreed to the automatic restart at the beginning or used the `-AutoReboot` parameter, the script will initiate a 60-second countdown before rebooting. Otherwise, it will remind you to restart manually.

### Why Disable Hibernation and Fast Startup?

> [!NOTE]
> This action is only performed if you use the `-DisableHibernation` switch. It is not enabled by default because some users, particularly on laptops, rely on hibernation to save their session and conserve battery. Disabling it can also sometimes interfere with power management features on certain hardware.

Windows Fast Startup doesn't fully shut down the system. Instead, it hibernates the core operating system to speed up the next boot. While fast, this can cause issues with drivers, software updates, and dual-booting environments because the system never gets a completely fresh start. Disabling it provides several benefits:

*   **Ensures a True "Fresh Start":** Forces a full shutdown, which can resolve persistent driver and software glitches that survive a normal reboot.
*   **Improves Update Reliability:** A full shutdown allows system files to be properly replaced during updates, preventing common failures.
*   **Fixes Driver State Issues:** Because drivers are fully re-initialized on a cold boot, this can resolve odd hardware behavior. Note that on the first restart after disabling hibernation, display settings (like resolution or multi-monitor arrangement) may temporarily change before correcting themselves.
*   **Frees Up Disk Space:** Deletes the `hiberfil.sys` file, reclaiming several gigabytes of space on your system drive.
*   **Aids Dual-Booting:** Prevents file system corruption issues when accessing the Windows partition from another operating system (like Linux).

### Does This Script Work?

Yes, though this won't fix every possible error, it addresses the most common Windows issues to improve overall performance. You may need to run the script twice to ensure all fixes apply correctly. If issues persist, wait 24 hours before running it again. Windows requires time to complete specific background tasks and maintenance cycles before certain repairs take effect.

### Will This Remove Viruses?

This process may repair the built-in Windows Defender service. However, if you suspect an active malware infection, do not rely on this fix alone. Verify that Windows Defender is currently running, then scan your system with a reputable second-opinion scanner such as the [ESET Online Scanner](https://download.eset.com/com/eset/tools/online_scanner/latest/esetonlinescanner.exe) or [Malwarebytes Free Scanner](https://www.malwarebytes.com/solutions/virus-scanner) to ensure nothing was missed.