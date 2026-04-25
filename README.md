<img width="842" height="86" alt="Logo" src="https://github.com/user-attachments/assets/61359630-82bb-460d-adcb-8aba8f61f4ca" />
<img width="1200" height="747" alt="Searching function" src="https://github.com/user-attachments/assets/d34e61a4-3d64-43d8-9ef8-cd421fbb65ba" />


# Intune Help Desk Dashboard

A fast, interactive CLI dashboard for IT Help Desk teams to query Microsoft Intune devices, view compliance health, and check storage capacity directly from the terminal. Built with PowerShell and Spectre Console.

## Features
* **Interactive UI:** Searchable dropdowns and modern terminal styling.
* **Live Graph Data:** Pulls live device diagnostics (Compliance, Sync Time, Storage, Recent Apps) via the Microsoft Graph API.
* **Smart Prerequisites Check:** Automatically detects and prompts users to install missing required modules.

## Prerequisites
The script will automatically check for these, but they must be installed:
* `PwshSpectreConsole`
* `Microsoft.Graph.Authentication`
* `Microsoft.Graph.DeviceManagement`

## Usage
1. Clone the repository or download the script.
2. Open PowerShell as Administrator (required for initial module installation).
3. Run the script:
   ```powershell
   .\Help_Desk_App.ps1
