# SoftwareScannerGUI

WPF GUI tool for auditing all installed software on a Windows system. Designed as a **debloat development aid** -- scan everything installed, then use the output to write targeted debloat scripts.

## Features

- Scans multiple sources: AppX packages, Win32 registry uninstall paths, running services, scheduled tasks, startup entries
- Real-time verbose output panel
- Dark theme UI
- DataGrid with publisher, version, install date, path columns
- Export to CSV / clipboard
- Search and filter

## Usage

```powershell
# Run as Administrator for full visibility
.\SoftwareScannerGUI.ps1
```

## Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges (recommended)
