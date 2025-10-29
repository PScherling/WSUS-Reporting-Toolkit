# PSC WSUS Reporting Toolkit
<img width="1596" height="874" alt="image" src="https://github.com/user-attachments/assets/020e3988-7ffb-49e3-85b7-00776839921d" />

Menu-driven PowerShell tooling to **inspect**, **report**, and **maintain** Windows Server Update Services (WSUS).  
Includes automated and manual installers, a lightweight module launcher, and the interactive reporting dashboard.

### Release Notes
| Version| Release Date | Title | Description | Features | Bug Fixes | Known Issues |
|--------------|--------------|--------------|--------------|--------------|--------------|--------------|
| **0.0.1** |  | Pre-Release | Initial first attempt | Showing System Information |  |  |
| **0.0.2** |  | Pre-Release | WSUS Synchronisation Report function | WSUS Synchronisation Report function finished |  |  |
| **0.0.3** |  | Pre-Release | WSUS Endpoint Status Report function | WSUS Endpoint Status Report function finished |  |  |
| **0.0.4** |  | Pre-Release | WSUS Update Status Report function | WSUS Update Status Report function finished |  |  |
| **0.0.5** |  | Pre-Release | Bug fixing |  |  |  |
| **0.0.6** |  | Pre-Release |Show 'in sync' status; Bug fixing |  |  |  |
| **0.1.0** | 02/Jun/2025 | Initial release | Finalized first stable version |  |  | Sync progress in UI shows always 100% during synchronisation |
| **0.1.1** | 17/Jun/2025 | Hot Fix | Bug Fix for Sync Progress |  | Sync Progress Status |  |
| **0.1.2** | 12/Aug/2025 | Hot Fix | Fixing PowerShell Syntax Errors |  |  |  |
| **0.1.3** | 23/Sep/2025 | Minor Release | Adding WSUS Cleanup Functions | WSUS Cleanup | Endpoint HTML Report | Endpoint Report Issue in HTML Report -> Device to Group mapping not correct |
| **0.1.4** | 25/Sep/2025 | Hot Fix | Fixed Endpoint Report |  |  |  |

---

## ✨ Features

- **Interactive dashboard** (console UI) showing:
  - Endpoint health: up-to-date, needing updates, with client errors, total clients
  - Update health: installed, needed by computers, with client errors
  - Server stats: approved, declined, pending-approval counts
  - Synchronization: last sync result & time, auto/manual mode, progress %, content download status
  - Connection info: HTTP/HTTPS, port, WSUS server version
  - Cleanup info: last cleanup timestamp & suggested next run
- **One-key actions**:
  1. Refresh dashboard  
  2. Generate **Endpoint Status Report** (HTML)  
  3. Generate **Update Status Report** (HTML)  
  4. Generate **Last Synchronization Report** (HTML)  
  5. **Run WSUS Cleanup**  
  6. **Test-Run WSUS Cleanup** (no-decline preview)
- **Robust logging** with timestamps
- **Role separation**: main tool + module + installers (manual & MDT/WDS)

---

## 📦 Repository Contents

| Path / Script | Purpose |
|---|---|
| `Data\` | Payload used by the installers (module files, launcher, any assets). |
| `Logfiles\` | Local log directory used by the tool (created at runtime under `C:\_it\psc_wsusreporting\Logfiles`). |
| `Data\launch_psc_wsusreporting.bat` | For starting and auto-launching the module. |
| `Data\psc_wsusreporting.cmd` | For starting the main powershell script. |
| `Data\psc_wsusreporting.ps1` | Main interactive dashboard + reporting + cleanup actions (runs on WSUS server). |
| `Data\psc_wsusreporting.psm1` | Module launcher that opens the tool in a new, maximized PowerShell window. |
| `manual_Install-PSC_wsus-report-tool.ps1` | Local/manual installer (uses local `.\Data` folder as source). |
| `custom_Install-PSC_wsus-report-tool.ps1` | Automated installer for MDT/WDS (pulls from `\\<MDT_FileSrv>\DeploymentShare$`). |


> The dashboard displays the value of `$VersionNumber` embedded in `psc_wsusreporting.ps1`.

---

## 🧱 Requirements

- Run on the **WSUS server** (local WSUS Admin API access)
- **Administrator** PowerShell session
- **PowerShell 5.1+** (or PowerShell 7.x on Windows)
- **WSUS Administration API** available: `Microsoft.UpdateServices.Administration`
- Local write access for logs and reports
- (Automated install) Network access to your deployment and log shares

---

## 🔧 Installation

### Option A — Automated (MDT/WDS)
Use: `custom_Install-PSC_wsus-report-tool.ps1`

What it does:
- Copies payload from: `\\<MDT_FileSrv>\DeploymentShare$\Scripts\custom\psc_wsusreporting\Data` → `C:\_it\psc_wsusreporting`
- Creates module path: `C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting`
- Copies `psc_wsusreporting.psm1/.psd1` into module path
- Copies `psc_wsusreporting.cmd` into `C:\Windows\System32`
- Imports module, creates desktop shortcut **WSUS-Report-Tool.lnk** to `launch_psc_wsusreporting.bat`
- Logs locally then uploads to: `\\<MDT_FileSrv>\Logs$\Custom\Configuration`

Run (as Admin) in your task sequence:
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\custom_Install-PSC_wsus-report-tool.ps1
```

### Option B — Manual (Local)  
Use: `manual_Install-PSC_wsus-report-tool.ps1`

What it does:
- Copies payload from local `.\Data` → `C:\_it\psc_wsusreporting`
- Installs module files and `psc_wsusreporting.cmd`
- Imports module, creates desktop shortcut
- Logs to `C:\_it\Configure_psc_wsusreporting_*.log`

Run (as Admin):
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\manual_Install-PSC_wsus-report-tool.ps1
```

---

## ⚙️ Configuration

Key locations used by the solution:

- **Main tool logs**: `C:\_it\psc_wsusreporting\Logfiles\psc_wsusreporting.log`
- **Installer logs** (pattern): `C:\_it\Configure_psc_wsusreporting_<COMPUTERNAME>_<YYYY-MM-DD_HH-mm-ss>.log`
- **Desktop shortcut**: `WSUS-Report-Tool.lnk` → `C:\_it\psc_wsusreporting\launch_psc_wsusreporting.bat`
- **Module path**: `C:\Program Files\WindowsPowerShell\Modules\psc_wsusreporting\`

> Update the installer variables for your deployment server (`$FileSrv`) and log share paths if needed.

---

## 🚀 Usage
### Start via desktop link (recommended)
<img width="123" height="71" alt="image" src="https://github.com/user-attachments/assets/0f0acf98-7ecc-48b6-b5f0-fd6df4b9e9e4" />

### Start via Module (recommended)
```powershell
Import-Module psc_wsusreporting
psc_wsusreporting
```

### Start via Script
```powershell
powershell.exe -ExecutionPolicy Bypass -File "C:\_it\psc_wsusreporting\psc_wsusreporting.ps1"
```

You’ll see the **Windows Server Update Services Reporting Tool** menu. Use the keys shown to generate HTML reports, run cleanup, or refresh.

---

## 📑 Reports & Exports

- **Endpoint Status Report** (HTML)
  <img width="1596" height="874" alt="image" src="https://github.com/user-attachments/assets/5dd3e318-0cb5-49fc-95d4-dd8a6799e13c" />
  <img width="1593" height="801" alt="image" src="https://github.com/user-attachments/assets/812d4be0-fd79-4576-b886-eacc10df3d8a" />

- **Update Status Report** (HTML)
  <img width="1596" height="874" alt="image" src="https://github.com/user-attachments/assets/5a2f6e1d-aead-4b46-98cf-5c9493763f85" />
  <img width="1593" height="801" alt="image" src="https://github.com/user-attachments/assets/6775e81f-5a9f-41a4-9550-8be5aedf3bf4" />

- **Last Synchronization Report** (HTML)
  <img width="1596" height="874" alt="image" src="https://github.com/user-attachments/assets/8d624d95-5f7e-4555-b19c-50ac54888257" />
  <img width="1593" height="801" alt="image" src="https://github.com/user-attachments/assets/c48195f3-d8c7-4a41-94b7-213da916dd21" />

- **Cleanup Exports**: `Exports\SupersededUpdates*.csv` (the tool reads latest export timestamp to suggest next cleanup)
  <img width="1593" height="801" alt="image" src="https://github.com/user-attachments/assets/3a774857-32a3-4a5c-a2fe-be32e7b30725" />


> Reports are saved alongside the script unless otherwise configured inside your report functions.

---

## 📝 Logging

The tool uses a simple `Write-Log` function to append timestamped entries:
- `C:\_it\psc_wsusreporting\Logfiles\psc_wsusreporting.log`

Installers write their own **timestamped** logs and (for automated installs) attempt to upload them to your deployment log share.

---

## 🔐 Security Notes

- Run only on trusted admin workstations/servers.
- Limit access to deployment/log shares.
- If you adapt the tool to publish reports to a share or web location, ensure proper ACLs and avoid exposing internal WSUS data externally.

---

## 🛠️ Troubleshooting

- **WSUS API not found**  
  Ensure the WSUS console / admin components are installed on the machine running the tool.

- **Access denied / reports not generated**  
  Run PowerShell as **Administrator**. Verify folder permissions and paths.

- **Sync progress shows “Collecting Data”**  
  That’s normal at the very beginning of a sync; the tool switches to percentages once totals are known.

- **Cleanup suggestions show “No Data”**  
  The suggestion relies on a prior `SupersededUpdates*.csv` export. Run a cleanup/export first.

---

## ❓ FAQ

**Q: Can I run this from a management server that isn’t the WSUS server?**  
A: The tool connects via the local WSUS Admin API. Run it **on the WSUS server** for best results.

**Q: Does the tool require PSWindowsUpdate?**  
A: No. It uses the WSUS Administration API (`Microsoft.UpdateServices.Administration`).

**Q: Where do I change the version number shown in the UI?**  
A: Update the `$VersionNumber` variable inside `psc_wsusreporting.ps1`.

---

## 🧭 Roadmap

- Optional publish of reports to a central, access-controlled share
- Optional HTML theme customization & branding
- Additional WSUS health checks (database, IIS bindings, SSL validation)

---

## 🤝 Contributing

Issues and PRs are welcome. Please avoid including real server names, IPs, or credentials in examples.

---

## 🔗 References

- Windows Update Agent (WUA) API:  
  https://learn.microsoft.com/en-us/windows/win32/wua_sdk/windows-update-agent--wua--api-reference
- WSUS product info & guidance:  
  https://learn.microsoft.com/en-us/windows-server/administration/windows-server-update-services/
- Legacy WSUS API refs (for background):  
  https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms744624(v=vs.85)  
  https://learn.microsoft.com/en-us/previous-versions/windows/desktop/ms748969(v=vs.85)  
  https://learn.microsoft.com/en-us/previous-versions/windows/desktop/mt748187(v=vs.85)

---

## 👤 Author

**Author:** Patrick Scherling  
**Contact:** @Patrick Scherling  

---

> ⚡ *“Automate. Standardize. Simplify.”*  
> Part of Patrick Scherling’s IT automation suite for modern Windows Server infrastructure management.
