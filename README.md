# README - PSC WSUS Reporting
Menu-driven PowerShell tooling to report on and maintain Windows Server Update Services (WSUS).
Includes a lightweight launcher module, an interactive dashboard script, and both manual and automated (MDT/WDS) installers.

## Components

psc_wsusreporting.ps1 — Main interactive dashboard (reports + cleanup).

psc_wsusreporting.psm1 — Launcher module exposing the psc_wsusreporting command (starts the tool in a new, maximized PowerShell window).

manual_Install-PSC_wsus-report-tool.ps1 — Local/manual installer using the repository’s .\Data folder.

custom_Install-PSC_wsus-report-tool.ps1 — Automated installer for MDT/WDS that pulls from \\<MDT_FileSrv>\DeploymentShare$.

Launchers / Shortcuts

psc_wsusreporting.cmd in %SystemRoot%\System32 (CLI convenience)

launch_psc_wsusreporting.bat (used by desktop shortcut)

Desktop shortcut EF_WSUS-Report-Tool.lnk (created by installers)

## What it does
Live WSUS Dashboard

From psc_wsusreporting.ps1, you get an at-a-glance status using the WSUS Admin API:

Environment: hostname, domain, IP, timestamp

Endpoint health: up-to-date, needing updates, with client errors, total clients

Update health: installed, needed by computers, with client errors

Server stats: approved / declined / pending approval counts

Synchronization: last sync time + result, auto/manual, in-progress state and %

Content download: progress/bytes (if any)

Connection: HTTP/HTTPS, port, WSUS server version

Cleanup: last cleanup export timestamp and suggested next run

Actions (menu)

Refresh dashboard

Generate Endpoint Status Report (HTML)

Generate Update Status Report (HTML)

Generate Last Synchronization Report (HTML)

Run WSUS Cleanup (declines superseded, etc.)

Test-Run WSUS Cleanup (safe preview, no decline)

All operations are timestamp-logged.

## Requirements

Run on the WSUS server with Administrator privileges

Windows PowerShell 5.1 (or PowerShell 7.x on Windows)

WSUS Administration components (Microsoft.UpdateServices.Administration assembly)

Local write access for logs/exports (default under C:\_it\psc_wsusreporting)

(Automated install) Network access to:

\\<MDT_FileSrv>\DeploymentShare$\Scripts\custom\psc_wsusreporting\Data

\\<MDT_FileSrv>\Logs$\Custom\Configuration

## Install
A) Manual install (local)
Run as Administrator from the repo root (or your package folder)
Unblock-File .\manual_Install-PSC_wsus-report-tool.ps1
powershell.exe -ExecutionPolicy Bypass -File .\manual_Install-PSC_wsus-report-tool.ps1


What it does:

Copies tool files from .\Data → C:\_it\psc_wsusreporting\

Installs module → C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting\

Places psc_wsusreporting.cmd in %SystemRoot%\System32

Imports the module

Creates desktop shortcut EF_WSUS-Report-Tool.lnk

B) Automated install (MDT/WDS)
Run as Administrator (deployed as a task-sequence step)
Unblock-File .\custom_Install-PSC_wsus-report-tool.ps1
powershell.exe -ExecutionPolicy Bypass -File .\custom_Install-PSC_wsus-report-tool.ps1


What it does:

Copies from deployment share → C:\_it\psc_wsusreporting\ (preserves structure)

Installs module + system launcher

Imports module

Creates desktop shortcut

Uploads the installer log to \\<MDT_FileSrv>\Logs$\Custom\Configuration

## Usage
Start the dashboard

If the module is installed:
psc_wsusreporting

Or run the script directly:
powershell.exe -ExecutionPolicy Bypass -File "C:\_it\psc_wsusreporting\psc_wsusreporting.ps1"

Typical flow

Launch the dashboard → review live status

Generate HTML reports as needed

Run a test cleanup (option 6) to preview impact

Run cleanup (option 5) when ready

## Configuration Highlights

The dashboard shows $VersionNumber in its header (e.g., 0.1.x).

Installers create/use:

C:\_it\psc_wsusreporting\

C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting\

%SystemRoot%\System32\psc_wsusreporting.cmd

If you maintain a deployment share:

Update custom_Install-PSC_wsus-report-tool.ps1 with your server name/IP:

\\<MDT_FileSrv>\DeploymentShare$\Scripts\custom\psc_wsusreporting\Data

\\<MDT_FileSrv>\Logs$\Custom\Configuration

## Logs & Outputs

Tool log:
C:\_it\psc_wsusreporting\Logfiles\psc_wsusreporting.log

Installer logs:
C:\_it\Configure_psc_wsusreporting_<COMPUTERNAME>_<YYYY-MM-DD_HH-mm-ss>.log
(Automated installer also uploads to \\<MDT_FileSrv>\Logs$\Custom\Configuration)

Reports/Exports:

HTML reports saved beside the script (or your chosen report path)

Cleanup export CSVs under:
..\script folder\Exports\SupersededUpdates*.csv

## Security Notes

Run from elevated PowerShell.

Limit access to report and log shares.

If you extend the tool to copy reports to a remote share, prefer:

Computer account permissions / Kerberos

Managed identities / GMSA

SecretManagement for credential retrieval

Review WSUS cleanup behavior in a lab before production use.

## Troubleshooting

“Cannot load WSUS Admin assembly”
Ensure WSUS server components are installed on the machine running the tool.

“Access denied” when creating files
Verify NTFS permissions on C:\_it\psc_wsusreporting\ and subfolders.

No sync progress shown
Sync may not be running; start a WSUS sync and refresh the dashboard.

Empty counts
First-time or stale WSUS deployments might need synchronization and/or client check-ins.


## References

Windows Update Agent (WUA) API
https://learn.microsoft.com/en-us/windows/win32/wua_sdk/windows-update-agent--wua--api-reference

WSUS Admin & Ops Guidance
https://learn.microsoft.com/en-us/windows-server/administration/windows-server-update-services/

PowerShell WSUS Admin Samples (community articles and examples are referenced in script headers)
