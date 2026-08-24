<div align="center">

# 🧹 Device Cleanup Manager

**Active Directory computer cleanup, safely**

Discover inactive and disabled computer objects, then disable, enable, or delete them in bulk — with confirmations and protected-OU rules.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue.svg)
![Platform](https://img.shields.io/badge/Windows-10%2F11-blue.svg)
![UI](https://img.shields.io/badge/UI-WPF%20GUI-blue.svg)
![AD](https://img.shields.io/badge/Directory-Active%20Directory-0078D4.svg)
![Version](https://img.shields.io/badge/version-2.0-green.svg)

[Features](#-core-features) • [Usage](#-usage) • [Requirements](#️-requirements) • [Troubleshooting](#-troubleshooting)

</div>

---

# 📖 Overview

**Device Cleanup Manager** is a modern WPF-based PowerShell tool designed to centrally manage **Active Directory computer objects** for cleanup and control.

It provides a safe, structured interface to:

* Discover **inactive devices** using `LastLogonDate` with a configurable **Days Inactive** threshold
* Discover **disabled computers** (`Enabled = False`)
* Scope searches to a selected **OU (LDAP)** using a searchable browser
* Import device names from **CSV** for bulk operations
* Execute controlled actions (**Disable / Enable / Delete**) with confirmations and protection rules
* Export results to CSV and track actions in a real-time **Message Center** log

---

## 🖼️ Screenshots

![Device Cleanup Manager main window — search modes, results DataGrid, and Message Center](Screenshot.png)

*Main window: search mode selection, OU scope, the results DataGrid with per-row selection, and the real-time Message Center.*

---

# ✨ Core Features

### 🔹 Search Modes
* **Inactive Devices (by days)** — uses `LastLogonDate` against a configurable cutoff date; optionally include devices with **no logon timestamp**
* **Disabled Computers** — server-side filter: `Enabled -eq $false`

### 🔹 OU Scope Selector
* Fast OU discovery via LDAP with a searchable picker window
* Scope options: **Entire Domain** or a **Specific OU** (Browse…)

### 🔹 Results Dashboard
DataGrid with per-row checkboxes and **Select All**, showing: Computer Name, Enabled, Last Logon Date, Inactive Days, OU Path, Distinguished Name, and Source (Search / CSV)

### 🔹 Bulk Actions
Actions available for **selected devices**:

| Action | Protected By |
|---|---|
| **Disable Selected** | Confirmation dialog + protected DN rules |
| **Enable Selected** | Confirmation dialog + protected DN rules |
| **Delete Selected** | Confirmation dialog + protected DN rules |

Protected DN rules skip any object under protected OUs/DNs before the action runs.

### 🔹 CSV Import (Bulk List Mode)
Import a names-only device list into the grid. Supported columns: `ComputerName`, `Name`.

```csv
ComputerName
PC-001
PC-002
LAB-010
```

Imported items are labeled `Source = CSV`; if the DN is missing, the tool resolves the device in AD at action time.

### 🔹 Export & Message Center
* **Export results** to CSV — ComputerName, Enabled, LastLogonDate, InactiveDays, OUPath, DistinguishedName, Source
* **Message Center** — real-time console log (INFO / SUCCESS / WARNING / ERROR) covering searches, scope selection, import/export tracking, action results, and LDAP/RSAT errors

---

# 🚀 Usage

### Launch

**Option 1 — PowerShell script:**

```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\DeviceCleanupManager.ps1
```

**Option 2 — Packaged EXE (self-signed with PSWrap):**

```text
DeviceCleanupManager.exe
```

The `.exe` was compiled and self-signed using [PSWrap](https://github.com/mabdulkadr/PSWrap) — no PowerShell console required.

### Typical Workflow

1. Launch the tool
2. Choose the **Mode**: Inactive Devices (set Days) or Disabled Computers
3. Select the **OU Scope**: Browse → search and pick an OU
4. Click **Run Search**
5. Review results and check devices
6. Execute actions: Disable / Enable / Delete Selected
7. Export results if needed

---

# ⚙️ Requirements

| Requirement | Details |
|-------------|---------|
| **OS** | Windows 10 / 11 |
| **Domain** | Connectivity to a domain controller (recommended) |
| **Elevation** | Run as **Admin** when your role/permissions require it |
| **Module** | `ActiveDirectory` (RSAT) |

Install RSAT (Windows 10/11):

```powershell
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

### Data & Logs

```text
C:\ProgramData\DeviceCleanupManager\
└── Logs\
```

---

# 🧠 Inactivity Logic (Important)

The tool uses `Get-ADComputer ... -Properties LastLogonDate`. `LastLogonDate` is derived from **lastLogonTimestamp replication**:

* ✅ Suitable for cleanup reporting and identifying stale objects
* ⚠ Not a real-time "last logon across all DCs" metric — recently active devices may lag by up to the replication tolerance

---

# 🔍 Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|--------------|-----|
| OU list doesn't load | Not connected to domain or no LDAP access | Ensure domain connectivity; check firewall |
| "RSAT module not found" | ActiveDirectory module not installed | Install RSAT via `Add-WindowsCapability` (see Requirements) |
| Search returns no results | Days threshold too low or wrong OU scope | Increase Days Inactive or select "Entire Domain" |
| Delete/Disable fails | Protected DN rule or insufficient permissions | Check protected OUs; verify delegated AD rights |
| CSV import shows empty | Wrong column name in CSV | Use `ComputerName` or `Name` as the column header |

---

# 🛡 Operational Notes

* **Built-in safeguards** — RSAT validation before AD operations, confirmation prompts for every destructive action, protected DN rule evaluation, and safe grid refresh after actions
* **Configure protected OUs first** — review the protected DN rules before your first Delete run; they are the last line of defense
* **Test in staging** — run a Disable pass on a pilot OU before any Delete; disabled devices can be re-enabled, deleted ones need AD Recycle Bin (or a restore)
* **Least privilege** — delegate only the rights needed (e.g., delete computer objects in target OUs), not full Domain Admin
* **Keep an audit trail** — export results before and after each action; the Message Center log is your session evidence

---

## 👤 Author

**Mohammad Abdulkader Omar**  
GitHub: [@mabdulkadr](https://github.com/mabdulkadr)  
Website: [momar.tech](https://momar.tech)  

---

## 📜 License

This project is licensed under the [MIT License](LICENSE).

---

## ⚠ Disclaimer

This skill and every script it generates are provided as-is with no warranty
of any kind. Test generated tools in a staging environment before deploying to
production. The authors assume no liability for any damage or data loss
resulting from their use.

---

<div align="center">

⭐ **If this tool saves you time, star the repo — it helps others find it.**

[Report an Issue](../../issues) · [momar.tech](https://momar.tech)

[![Buy Me A Coffee](https://img.shields.io/badge/Buy%20Me%20A%20Coffee-FFDD00?style=for-the-badge&logo=buy-me-a-coffee&logoColor=black)](https://www.buymeacoffee.com/mabdulkadrx)

Built with [**PowerShell Enterprise Admin**](https://github.com/mabdulkadr/powershell-enterprise-admin-skill)

</div>
