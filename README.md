<p align="center"><img src="icon.svg" width="128" height="128" alt="SoftwareScannerGUI"></p>

# SoftwareScannerGUI

![Version](https://img.shields.io/badge/version-v1.1.0-blue) ![License](https://img.shields.io/badge/license-MIT-green) ![Platform](https://img.shields.io/badge/platform-PowerShell-lightgrey)

WPF GUI tool for auditing all installed software on a Windows system. Designed as a **debloat development aid** -- scan everything installed, then use the output to write targeted debloat/removal scripts.


## Features

- **Multi-source scanning**: AppX packages, Win32 registry (x64/x86/user), provisioned packages, WinGet, services, scheduled tasks, startup entries
- **Real-time verbose output** panel with timestamped log
- **Dark theme UI** (#1a1a2e deep blue)
- **Install size column**: Reads `EstimatedSize` from registry (KB), measures AppX package folder sizes. Total disk usage shown in status bar
- **Category filter**: ComboBox to filter view by source type (All, Registry, AppX, Services, etc.)
- **Startup impact indicators**: High/Low impact rating for startup items based on known heavy-hitters. Startup type classification (Run Key, Startup Folder, Service - Auto, Scheduled Task)
- **Generate Removal Script**: Select rows across any DataGrid, generates a complete .ps1 script with appropriate removal method per source type, error handling, and results summary
- **Snapshot system**: Save scan results to JSON. Compare two snapshots to see added/removed items (before/after debloat comparison)
- **Export**: Current tab CSV, all data CSV, debloat script, selected items, copy removal commands to clipboard
- **Bloatware detection**: Pattern-based flagging of known bloatware

## Usage

```powershell
# Run as Administrator for full visibility
.\SoftwareScannerGUI.ps1
```

## Requirements

- Windows 10/11
- PowerShell 5.1+
- Administrator privileges (recommended)
