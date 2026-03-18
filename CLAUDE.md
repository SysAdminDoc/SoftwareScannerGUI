# CLAUDE.md - SoftwareScannerGUI

## Overview
WPF GUI for auditing all installed software on a Windows system. Debloat development aid -- scan everything, then use the output to write targeted removal scripts. v1.1.0.

## Tech Stack
- PowerShell 5.1, WPF GUI
- Dark blue theme (#1a1a2e)

## Key Details
- Single-file PowerShell script
- Scans: AppX packages, Win32 registry uninstall paths, running services, scheduled tasks, startup entries, WinGet, provisioned packages
- Real-time verbose output panel
- DataGrid with publisher, version, install size, arch, bloat flag columns
- Export to CSV / clipboard
- Search and filter by category (ComboBox filter hides/shows tabs)
- Install size column: Registry EstimatedSize (KB) + AppX folder measurement
- Total disk usage in status bar
- Startup impact indicators: High/Low rating based on known heavy-hitters, startup type classification (Run Key, Startup Folder, Service - Auto, Scheduled Task)
- Generate Removal Script: select rows across any grid, generates .ps1 with per-source removal logic (Remove-AppxPackage, uninstall strings, Stop-Service, Disable-ScheduledTask, Remove-ItemProperty) with try/catch and results summary
- Snapshot system: Save scan results to JSON, compare two snapshots to see added/removed items
- Category filter: ComboBox to show All Sources or filter to single source type (collapses other tabs)

## Build/Run
```powershell
# Run as Administrator for full visibility
.\SoftwareScannerGUI.ps1
```

## Version History
- 1.1.0: Install size column, generate removal script, snapshot save/compare, startup impact indicators, category filter ComboBox
- 1.0.0: Initial release -- scan, export CSV, debloat script, copy removal commands
