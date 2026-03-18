# CLAUDE.md - SoftwareScannerGUI

## Overview
WPF GUI for auditing all installed software on a Windows system. Debloat development aid — scan everything, then use the output to write targeted removal scripts. v1.0.0.

## Tech Stack
- PowerShell 5.1, WPF GUI
- Dark blue theme (#1a1a2e)

## Key Details
- ~1,194 lines, single-file
- Scans: AppX packages, Win32 registry uninstall paths, running services, scheduled tasks, startup entries
- Real-time verbose output panel
- DataGrid with publisher, version, install date, path columns
- Export to CSV / clipboard
- Search and filter

## Build/Run
```powershell
# Run as Administrator for full visibility
.\SoftwareScannerGUI.ps1
```

## Version
1.0.0
