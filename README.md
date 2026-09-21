# AKTIS Helper

A collection of Windows PowerShell automation scripts that connect **Aktis** (a proposal-tracking system built on SharePoint Online lists) with **Clockify** time tracking and Excel workplans. The scripts pull financial data out of proposal workplans, post proposal statistics back to SharePoint, sync proposals into Clockify as projects, and produce monthly billing spreadsheets.

> **Platform:** Windows only (uses WinForms, Excel COM automation, and the Windows Event Log).

---

## Scripts

| Script | Type | What it does |
|---|---|---|
| `Aktis Helper v9.ps1` | Interactive (GUI) | Drag-and-drop tool. Drop an Excel workplan onto the window and it extracts the proposal's financials, lets you review/enter time data, then posts the results to Aktis and the local *Analytics Weekly Snapshot* workbook. |
| `Aktis Helper Automated.ps1` | Background service | Headless version of Aktis Helper. Continuously polls the **Proposal Queues** list and automatically processes proposals when their ticket status changes to `Draft Complete` or `Complete`. |
| `Aktis Proposal Monitor (Clockify Integration).ps1` | Background service | Watches Aktis for new proposals, creates a workplan from the correct client template, and creates the matching project and tasks in Clockify. |
| `Generate Monthly Billing Spreadsheet V7.ps1` | Interactive (GUI) | Prompts for a date range and output folder, then combines Customer Profile data, proposal analytics, and Clockify reports into a monthly billing spreadsheet per MSP client. |

### Aktis Helper (v9 and Automated)

Both versions share the same core pipeline:

1. **Extract** – Opens the Excel workplan via COM and reads proposal name, proposed hours, hourly rate, MRR, professional-services revenue, products revenue, and NRR (`Get-FinancialDataFromWorkplan`).
2. **Enrich** – Looks up existing records in the **Proposal Statistics** and **Customer Profile** SharePoint lists and checks whether a proposal exceeds the client's Enterprise/Basic subscription.
3. **Time data** – Pulls Clockify time entries for the proposal at the project and task level and converts them to Aktis time format.
4. **Post** – Updates the **Proposal Statistics** / **Proposal Queues** lists and writes to the *Analytics Weekly Snapshot* workbook (DDE new proposal, DDE revision, and Proposal Coordinator sheets).
5. **Notify** – Shows a Windows balloon notification on success and writes to the `AKTIS` event log.

The **Automated** version replaces the drag-and-drop form with a polling loop: it caches the current queue to `C:\Aktis\Aktis Helper\CurrentProposalQueue.csv`, compares it with the live list every few seconds, and processes any proposal whose status has changed. It stops if an error is encountered.

### Proposal Monitor (Clockify Integration)

1. Resolves Clockify client IDs for each supported MSP client.
2. Reads proposals from the **Proposal Queues** list and skips those already recorded in the posted-proposals log.
3. Copies the correct client workplan template (or the previous version, for revisions) into `C:\Aktis\Temp`.
4. Creates the Clockify project and its tasks.
5. Logs actions to `C:\Aktis\Clockify Integration\Logs\`.

---

## Requirements

- Windows 10/11 or Windows Server
- Windows PowerShell 5.1
- [PnP.PowerShell](https://pnp.github.io/powershell/) module
- Microsoft Excel (desktop) — used through COM automation
- Access to the SharePoint Online site hosting the Aktis lists:
  - `Proposal Queues`
  - `Proposal Statistics`
  - `Customer Profile`
- A Clockify workspace and API key
- OneDrive/SharePoint sync of the OPS document library (for workplan templates and analytics files)
- For the background services: an Entra ID (Azure AD) app registration with certificate authentication to SharePoint

---

## Setup

### 1. Install PnP.PowerShell

```powershell
Install-Module PnP.PowerShell -Scope AllUsers
```

### 2. Create the event log source

The scripts log to a custom `AKTIS` event log. Run once from an elevated PowerShell prompt:

```powershell
New-EventLog -LogName AKTIS -Source "Aktis Helper"
```

### 3. Create working folders

```powershell
New-Item -ItemType Directory -Force -Path `
  "C:\Aktis\Aktis Helper", `
  "C:\Aktis\Temp", `
  "C:\Aktis\Clockify Integration\Logs", `
  "C:\Temp\Aktis\Aktis Helper\Staging"
```

### 4. Configure credentials

The following values are set near the top of the relevant functions and must match your environment:

| Setting | Used by | Where |
|---|---|---|
| SharePoint site URL / tenant | All scripts | `AuthenticateToSPO` |
| Entra ID Client ID + certificate thumbprint | Automated, Proposal Monitor | `AuthenticateToSPO` |
| Clockify workspace ID | Helper, Automated, Monitor, Billing | `RetrieveClockifyTimeData`, `StartClockifyIntegration`, etc. |
| Clockify API key | Helper, Automated, Monitor, Billing | same as above |

It is strongly recommended to load secrets from the environment or a secret store instead of the script files, for example:

```powershell
$Script:ClockifyAPIKey = $env:CLOCKIFY_API_KEY
# or, with Microsoft.PowerShell.SecretManagement:
$Script:ClockifyAPIKey = Get-Secret -Name ClockifyAPIKey -AsPlainText
```

For certificate authentication, import the certificate into the service account's certificate store on the machine running the script and reference it by thumbprint — do not keep the `.pfx` in the repository.

---

## Usage

**Interactive Aktis Helper**

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Aktis Helper v9.ps1"
```

Sign in to SharePoint when prompted, then drag an Excel workplan onto the window and follow the prompts.

**Automated Aktis Helper / Proposal Monitor**

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Aktis Helper Automated.ps1"
powershell.exe -ExecutionPolicy Bypass -File ".\Aktis Proposal Monitor (Clockify Integration).ps1"
```

These are designed to run continuously, e.g. as a Scheduled Task set to start at boot under a service account.

**Monthly billing**

```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Generate Monthly Billing Spreadsheet V7.ps1"
```

Choose the start date, end date, and output folder, then click **Submit**.

---

## Logging

| Location | Contents |
|---|---|
| Event Viewer → `AKTIS` log, source `Aktis Helper` | Information, warning, and error events from all scripts |
| `C:\Aktis\Clockify Integration\Logs\ActionLog--MM-yyyy.log` | Proposals created by the Clockify integration |
| `C:\Aktis\Clockify Integration\Logs\PostedProposalsLog--MM-yyyy.log` | KT numbers already synced to Clockify |
| `C:\Aktis\Aktis Helper\CurrentProposalQueue.csv` | Last known proposal queue snapshot (Automated) |

---

## Repository contents

```
AKTIS-Helper/
├── Aktis Helper v9.ps1
├── Aktis Helper Automated.ps1
├── Aktis Proposal Monitor (Clockify Integration).ps1
├── Generate Monthly Billing Spreadsheet V7.ps1
├── AKTISFLOWS.cer
└── AKTISFLOWS.pfx
```

---

## Security notes

- Never commit API keys, client secrets, or private-key files (`.pfx`) to source control. Add them to `.gitignore`:
  ```gitignore
  *.pfx
  *.key
  secrets.*
  ```
- If credentials have ever been committed, rotate them (regenerate the Clockify API key, replace the app registration certificate) and purge them from Git history with a tool such as [`git filter-repo`](https://github.com/newren/git-filter-repo) or BFG.

---

## Known limitations

- Workplan template paths and client names are hard-coded per MSP client; adding a client requires editing the scripts.
- Paths assume the default OneDrive sync location for the *Project Fuel* SharePoint libraries.
- Excel must be installed and the workplan must not be open/locked by another user while processing.
