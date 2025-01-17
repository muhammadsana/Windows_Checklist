# Windows_Checklist

**Overview**

The Windows Checklist GUI is a tool designed to provide vital system information about a Windows computer. This application displays details such as Windows updates, BitLocker status, local administrators, computer name, and other necessary information. Additionally, it includes buttons to run updates, enable features, and refresh the GUI.

**Administrator privileges (for certain features like running updates)**

Execute the script to launch the GUI.
The main window will display system information such as Windows updates, BitLocker status, local administrators, and computer name and more.

**Interacting with the GUI:**
Use the buttons provided to:
- Run Windows Updates: This will initiate the Windows update process.
- Turn On Features: Enable necessary features on the system.
- Refresh GUI: Refreshes the information displayed in the GUI to reflect the current system status.

## Features

- **Hostname**: Displays the computer's hostname.
- **Active Directory Status**: Checks if the computer is bound to Active Directory.
- **Local Administrators**: Lists local administrators.
- **Windows Updates**: Shows available Windows updates and checks if any updates require a restart.
- **BitLocker Status**: Displays the BitLocker status of the C: drive.
- **SCCM Client**: Verifies if the SCCM client is installed.
- **Dell Command Update**: Runs Dell Command Update to check and apply updates.

## Usage

To run the script, follow these steps:

1. Save the script to a `.ps1` file.
2. Open PowerShell with administrative privileges.
3. Navigate to the directory containing the script.
4. Execute the script by typing `.\YourScriptName.ps1`.
