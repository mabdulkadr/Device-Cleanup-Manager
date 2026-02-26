# DevicesInactiveCleanupTool GUI

![Scope](https://img.shields.io/static/v1?label=scope&message=ActiveDirectory&color=blue)
![Mode](https://img.shields.io/static/v1?label=mode&message=CLI%2BGUI&color=lightgrey)
![Scripts](https://img.shields.io/static/v1?label=scripts&message=2&color=green)
![Pattern](https://img.shields.io/static/v1?label=pattern&message=StandaloneScripts&color=blue)
![Tech](https://img.shields.io/static/v1?label=tech&message=AD&color=blue)

---

## 📖 Overview
This folder contains **2 PowerShell script(s)**. The documentation below is generated from actual script content and includes technical behavior, dependencies, integration points, and exit-code patterns.


## ✨ Features
- Folder scope: `ActiveDirectory`
- Execution mode: `CLI+GUI`
- Scripts detected: **2**
- Path: `ActiveDirectory-Scripts\DevicesInactiveCleanupTool GUI\README.md`


## ⚙️ Requirements
- Windows PowerShell 5.1 or newer.
- Permissions aligned with script operations (file system, services, tasks, registry, API).
- Required modules and APIs are listed per script in Technical Details.

## 📂 Script Inventory
| File | Type | Synopsis |
|---|---|---|
| `DevicesInactiveCleanupTool.ps1` | Automation | Automation script for DevicesInactiveCleanupTool. |
| `test.ps1` | Automation | Automation script for test. |


## 🔍 Technical Details
### `DevicesInactiveCleanupTool.ps1`
- **Functional Type:** Automation
- **Purpose:** Automation script for DevicesInactiveCleanupTool.
- **Technical Description:** This script automates tasks related to DevicesInactiveCleanupTool. Review prerequisites, permissions, and execution context before production deployment. Exit codes: - Exit 0: Completed successfully - Exit 1: Failed or requires further action
- **Expected Run Context (Run As):** System or User (according to assignment settings and script requirements).
- **Path:** `ActiveDirectory-Scripts\DevicesInactiveCleanupTool GUI\DevicesInactiveCleanupTool.ps1`
- **Observed Exit Codes:** `0`, `1`
- **Technical Dependencies:**
  - RSAT ActiveDirectory module and AD read/write permissions based on target operations.

#### Internal Functions
- `Populate-OUComboBox`
- `Search-InactiveComputers`
- `Show-WPFConfirmation`
- `Show-WPFMessage`

#### Key Cmdlets/Commands
- `Add-Type`
- `Add-WindowsCapability`
- `Disable-ADAccount`
- `Export-Csv`
- `Get-ADComputer`
- `Get-ADDomain`
- `Get-ADObject`
- `Get-Command`
- `Get-Date`
- `Get-Module`
- `Import-Module`
- `Install-WindowsFeature`
- `New-Object`
- `New-TimeSpan`
- `Populate-OUComboBox`
- *(+10 additional commands found in script)*

### `test.ps1`
- **Functional Type:** Automation
- **Purpose:** Automation script for test.
- **Technical Description:** This script automates tasks related to test. Review prerequisites, permissions, and execution context before production deployment. Exit codes: - Exit 0: Completed successfully - Exit 1: Failed or requires further action
- **Expected Run Context (Run As):** System or User (according to assignment settings and script requirements).
- **Path:** `ActiveDirectory-Scripts\DevicesInactiveCleanupTool GUI\test.ps1`
- **Observed Exit Codes:** `0`, `1`
- **Technical Dependencies:**
  - RSAT ActiveDirectory module and AD read/write permissions based on target operations.

#### Key Cmdlets/Commands
- `Get-ADComputer`
- `Import-Module`
- `Write-Error`
- `Write-Output`


## 🚀 Usage
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\DevicesInactiveCleanupTool.ps1
.\test.ps1
```


## 🛡️ Operational Notes
- ✅ Validate scripts in a pilot environment before production rollout.
- 🔎 Review execution logs (if present) and verify exit codes match expected behavior.
- ⚠️ For Intune use cases, validate assignment context and **Run this script using logged-on credentials** configuration.


## 📦 Additional Files
- `DevicesInactiveCleanupTool.exe`
- `Screenshot.png`
- `trash (1).ico`
- `trash (1).png`
- `trash.ico`


## 🧷 Compatibility and Revision
- Documentation last updated: **2026-02-15**
- This README is standardized and generated from local script analysis to keep documentation aligned with implementation.

---

## 📜 License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).

---

## ⚠️ Disclaimer

This script is provided **as-is** without warranty.
The author is **not responsible** for unintended modifications or data loss.
Always test thoroughly before deploying in production.

