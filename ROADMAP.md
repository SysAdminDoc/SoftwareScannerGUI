# SoftwareScannerGUI Roadmap

Roadmap for SoftwareScannerGUI - the WPF software-audit tool that inventories installed AppX, Win32, services, scheduled tasks, startup entries, and generates targeted debloat scripts.

## Planned Features

### Scanning depth
- MSIX package audit (shared packages vs main packages vs modification packages)
- Winget-source discrepancy view (show packages installed but no longer in any winget source = likely orphaned)
- Chocolatey and Scoop-managed package inventory as additional source columns
- Browser extension inventory (Chrome, Edge, Firefox, Brave) - many of these are just as bloaty
- Office add-ins (COM, VSTO, Excel XLL) - notoriously hard to remove, worth surfacing
- Windows Store provisioned packages *with* per-user install state (not just the global provisioned list)

### Evidence & telemetry
- Digital signature validation per binary with unsigned-binary warning
- Binary count + total install size breakdown per vendor
- Autoruns-style inventory of shell extensions, context menu handlers, explorer namespace extensions
- Scheduled task triggers fully parsed (calendar, logon, event, boot) instead of "Scheduled Task"
- Service dependency graph so removal scripts sequence in correct order

### Removal-script generator
- Multi-engine output: PowerShell, batch, winget-cli script, Intune remediation script
- Detection-script pair (Intune-compatible) alongside every generated removal script
- Rollback script generator from the same UI selection
- Safety ring: auto-insert System Restore checkpoint, event log entries, and exit codes
- Dry-run mode in generated scripts (`-WhatIf`) by default, with `-Force` to actually remove

### UI
- Saved-view profiles (e.g. "Dell new-deploy audit", "Enterprise golden image diff")
- Diff mode against a baseline JSON (added/removed/changed) as a first-class tab
- Keyword flag list editor with regex + per-source scoping
- Column picker per tab, export-only columns, sort persistence
- Export Markdown report beside CSV/JSON for direct posting in tickets

### Deployment scenarios
- `-Silent -Snapshot out.json` CLI mode (no GUI) for RMM inventory runs
- Baseline-snapshot diff report pair with HTML output for change review
- Compare two remote machines over WinRM / PS Remoting

## Competitive Research

- **Revo Uninstaller / Geek Uninstaller** - strong on leftover detection after uninstall. SoftwareScannerGUI currently audits *before*; add a "uninstall + leftover sweep" pairing with Scour for the actual cleanup.
- **Autoruns (Sysinternals)** - canonical for startup inventory. Borrow its source taxonomy (Logon, Explorer, Scheduled Tasks, Services, Drivers, Winsock) and expose as tabs.
- **Uninstall Tool / Bulk Crap Uninstaller** - BCU's pre-made removal rules for unwanted apps (Dell, HP, Lenovo bloat) are community-maintained; ship a similar YAML rule-pack format this tool can consume.
- **PatchMyPC Home Updater** - runs silent scans suitable for Intune/SCCM deployment; align CLI flags and JSON output so SoftwareScannerGUI can drop in.

## Nice-to-Haves

- "Vendor explorer" graph - cluster rows by publisher, show sizes and last-used dates
- PSGallery publication as `SoftwareScanner` module with a `Get-InstalledSoftware` cmdlet
- Built-in upload to a self-hostable dashboard (optional) for fleet-wide bloat reporting
- Scheduled recurring scans written to `%ProgramData%\SoftwareScannerGUI\Snapshots\` with auto-rotation
- One-click "generate debloat issue" - opens a new GitHub issue with the bloat summary attached
- Integration with MavenWinUtil and SystemUpdatePro so removal scripts auto-register as toggles

## Open-Source Research (Round 2)

### Related OSS Projects
- https://github.com/auberginehill/get-installed-programs — Combines registry (HKLM x64/x86, HKCU) + AppX enumeration via Get-AppxPackage, elevated-vs-not differences documented
- https://github.com/BulentTuncbilek/PowerShell-InventoryTool — WMI + Remote Registry over AD or manual lists, parallel scanning, Excel export
- https://github.com/MicrosoftDocs/PowerShell-Docs (Working-with-Software-Installations.md) — Canonical MS guidance on Win32_Product vs registry perf trade-off
- https://github.com/smbrine/WinSW-Inventory — Windows service inventory across fleets
- https://github.com/PSBicep/Bicep-Inventory — Schema reference for structured software-inventory JSON
- https://github.com/PowerShell/Microsoft.PowerShell.ConsoleGuiTools — Out-ConsoleGridView as an alternative to WPF GridView
- https://github.com/topics/software-inventory — Topic hub

### Features to Borrow
- Parallel per-computer scan via `Start-ThreadJob` or `ForEach-Object -Parallel` (PS7) for fleet mode (PowerShell-InventoryTool)
- Excel export with formatted headers + frozen top row + conditional formatting on install-size column (PowerShell-InventoryTool)
- HKCU-as-well enumeration — current-user-only installs are invisible to HKLM-only scans; requires iterating loaded NTUSER.DAT hives (get-installed-programs)
- Explicit avoidance of `Win32_Product` — documented 24s vs <1s and triggers MSI reconsistency checks (MS Docs)
- Loaded-hive scan for offline profiles on the disk (`reg load` each `NTUSER.DAT` under `C:\Users\*`) so scanner catches installs of logged-off users
- Remote-machine mode via WinRM/PSRemoting + credential prompt (PowerShell-InventoryTool)
- Out-ConsoleGridView alt view for headless / SSH / PS7 scenarios (ConsoleGuiTools)

### Patterns & Architectures Worth Studying
- Collector -> normalizer -> sink pipeline: each source (AppX, HKLM, HKCU, WinGet, services, tasks, startup) emits a common `InstalledItem` PSCustomObject with `Source, Name, Version, Publisher, InstallDate, InstallSize, Uninstall, Scope` (several repos hint, none canonical)
- Snapshot diff format as JSON-Patch for machine-readable "what changed" rather than full dump compare
- Removal-script emitter uses per-source removal templates (`AppX -> Remove-AppxPackage`, `HKLM64 -> msiexec /x`, `WinGet -> winget uninstall`) — already have, document as "adapter" extension point for plugins
- Schema-validated JSON manifest for snapshots so downstream consumers (Intune, SCCM, MavenWinUtil) can ingest safely
