# Windows Firewall Inspector

Tools to collect and analyze Windows Firewall logs across local and remote Windows machines using Azure Table Storage or local CSV files. **No Azure PowerShell modules or Azure CLI required** — all Azure communication uses pure REST API calls.

---

## Overview

The project consists of three programs that work together in a pipeline: **collect** firewall logs on Windows machines, **store** them centrally, and **visualize** them in a rich dashboard.

```
┌──────────────────────────────┐       ┌───────────────────────────┐
│   Windows Machine(s)         │       │   Your Workstation        │
│                              │       │                           │
│  ┌────────────────────────┐  │       │  ┌─────────────────────┐  │
│  │  FirewallLogAgent.ps1  │──┼──┐    │  │ FirewallLogViewer   │  │
│  │  (GUI or headless)     │  │  │    │  │   .ps1              │  │
│  └────────────────────────┘  │  │    │  └────────┬────────────┘  │
│                              │  │    │           │               │
│  ┌────────────────────────┐  │  │    │   Queries │ via REST      │
│  │ FirewallLogCollector   │  │  │    │           ▼               │
│  │   .ps1                 │──┼──┤    │  ┌─────────────────────┐  │
│  │  (CSV or scheduled     │  │  │    │  │ Azure Table Storage │  │
│  │   task)                │  │  ├───►│  │  (central store)    │  │
│  └────────────────────────┘  │  │    │  └─────────────────────┘  │
└──────────────────────────────┘  │    │                           │
                                  │    │  ─── OR ───               │
        (multiple VMs)            │    │                           │
                                  │    │  ┌─────────────────────┐  │
                                  └───►│  │   Local CSV files   │  │
                                       │  └─────────────────────┘  │
                                       └───────────────────────────┘
```

| Component | Path | Purpose |
|---|---|---|
| **Firewall Log Agent** | `FirewallAgent/FirewallLogAgent.ps1` | Interactive GUI or headless agent that enables firewall logging, resolves processes, and uploads entries to Azure Table Storage in real time. |
| **Firewall Log Collector** | `FirewallAgent/FirewallLogCollector.ps1` | Standalone script for local or remote VMs — writes firewall logs to date-stamped CSV files. Can be deployed as a scheduled task or run via `Invoke-AzVMRunCommand`. |
| **Collector Installer** | `FirewallAgent/FirewallLogCollectorInstaller.ps1` | Install/uninstall-only wrapper for Azure Run Command. Deploys a self-contained scheduled task collector without exposing the other collector actions. |
| **Firewall Log Viewer** | `FirewallViewer/FirewallLogViewer.ps1` | Dark-mode WPF dashboard — queries Azure Table Storage or loads CSV files, filters/sorts data, charts activity, looks up IP owners, and exports to CSV. |

---

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 / Server 2016+ (Session Host compatible) |
| **PowerShell** | 5.1+ (Windows PowerShell) |
| **Elevation** | Agent and Collector require **Run as Administrator** |
| **Azure Storage Account** | General-purpose v2 with Table Storage (only needed for the Azure upload path) |
| **Authentication** | Storage Account **Access Key**, or an Azure AD account with **Storage Table Data Contributor** RBAC role |

---

## Azure Storage Setup (One-Time)

If you plan to use Azure Table Storage as the central store:

1. In the **Azure Portal**, create a **Storage Account** (general-purpose v2).
2. Table Storage is included by default — no extra configuration needed.
3. Choose an authentication method:
   - **Access Key** — copy a key from *Storage Account → Access Keys*.
   - **Azure AD (RBAC)** — assign the **Storage Table Data Contributor** role to your user on the storage account.

> If you only need local CSV collection (e.g., via the Collector script), you can skip Azure setup entirely.

---

## Quick Start

### Step 1 — Collect Firewall Logs

You have **two options** for collection. Pick the one that fits your scenario:

#### Option A: Firewall Log Agent (interactive, uploads to Azure)

Best for: monitoring a single machine in real time with a GUI, or running headless on a server that uploads directly to Azure.

```powershell
# GUI mode — right-click PowerShell → Run as Administrator
.\FirewallAgent\FirewallLogAgent.ps1
```

The GUI walks you through:
1. Click **Connect to Azure** — authenticates via device code flow.
2. Select your **Subscription → Storage Account → Table** from cascading dropdowns.
3. Click **Start Monitoring** — the agent enables firewall logging, reads `pfirewall.log` every 60 seconds, resolves PIDs to process names, and uploads entries to Azure Table Storage.
4. Click **Stop Monitoring** or close the window — original firewall logging settings are restored automatically.

**Headless mode** (for servers / automation):

```powershell
# With Access Key
.\FirewallAgent\FirewallLogAgent.ps1 -NoGUI -StorageAccount "myaccount" -AccessKey "base64key=="

# With Azure AD device code flow
.\FirewallAgent\FirewallLogAgent.ps1 -NoGUI -StorageAccount "myaccount" -TenantId "contoso.onmicrosoft.com"
```

| Parameter | Default | Description |
|---|---|---|
| `-NoGUI` | `$false` | Run without the WPF interface |
| `-StorageAccount` | — | Storage account name (required for headless) |
| `-AccessKey` | — | Storage access key (omit to use Azure AD) |
| `-TableName` | `FirewallLogs` | Azure Table name |
| `-IntervalSeconds` | `60` | Seconds between collection cycles |
| `-TenantId` | `common` | Azure AD tenant for device code flow |

#### Option B: Firewall Log Collector (local CSV, remote-friendly)

Best for: deploying to VMs via `Invoke-AzVMRunCommand`, running as a scheduled task, or collecting logs to CSV without an Azure dependency.

```powershell
# One-shot collection — writes a CSV and exits
.\FirewallAgent\FirewallLogCollector.ps1 -Action Collect

# Continuous collection loop (Ctrl+C to stop)
.\FirewallAgent\FirewallLogCollector.ps1 -Action Enable

# Install as a scheduled task (runs as SYSTEM at startup)
.\FirewallAgent\FirewallLogCollector.ps1 -Action Install

# Remove scheduled task and disable logging
.\FirewallAgent\FirewallLogCollector.ps1 -Action Uninstall

# Disable logging without removing the task
.\FirewallAgent\FirewallLogCollector.ps1 -Action Disable
```

**Remote VM deployment** via `Invoke-AzVMRunCommand`:

```powershell
# One-shot collection on a remote VM
Invoke-AzVMRunCommand -ResourceGroupName 'rg' -VMName 'vm' `
    -CommandId 'RunPowerShellScript' `
    -ScriptPath '.\FirewallAgent\FirewallLogCollector.ps1' `
    -Parameter @{ Action = 'Collect' }

# Install the scheduled task on a remote VM
Invoke-AzVMRunCommand -ResourceGroupName 'rg' -VMName 'vm' `
    -CommandId 'RunPowerShellScript' `
    -ScriptPath '.\FirewallAgent\FirewallLogCollector.ps1' `
    -Parameter @{ Action = 'Install'; IntervalSeconds = '30' }

# Install the scheduled task using the install/uninstall-only wrapper
Invoke-AzVMRunCommand -ResourceGroupName 'rg' -VMName 'vm' `
    -CommandId 'RunPowerShellScript' `
    -ScriptPath '.\FirewallAgent\FirewallLogCollectorInstaller.ps1' `
    -Parameter @{ Action = 'Install'; IntervalSeconds = '30' }
```

| Parameter | Default | Description |
|---|---|---|
| `-Action` | `Enable` | `Enable`, `Disable`, `Install`, `Uninstall`, or `Collect` |
| `-OutputPath` | `C:\ProgramData\FirewallLogCollector\Logs` | Directory for CSV output |
| `-IntervalSeconds` | `60` | Seconds between collection cycles |

CSV files are named `FirewallLog_YYYY-MM-DD.csv` (one per day) and include columns for process name, firewall rule name, and application path resolved from the Security Event Log.

---

### Step 2 — View & Analyze Logs

```powershell
.\FirewallViewer\FirewallLogViewer.ps1
```

The Viewer supports **two data sources**:

| Source | How to connect |
|---|---|
| **Azure Table Storage** | Menu → *Actions → Connect to Azure*. Authenticate via device code or access key, then select your storage account and table. Use the date range and filter controls, then click **Query**. |
| **Local CSV file** | Menu → *File → Load CSV File*. Browse to a `FirewallLog_*.csv` produced by the Collector (or exported from a previous session). |

#### Viewer Features

- **Filtering** — date range, action (ALLOW/DROP), protocol (TCP/UDP/ICMP), direction (SEND/RECEIVE), IP address, port, and process name.
- **Sortable Data Grid** — click any column header to sort. Alternating row colors in dark mode.
- **IP Owner Lookup** — *Actions → Lookup IP Owners* resolves public IPs via ip-api.com (batch API, 100 IPs/request) with automatic caching. Private/reserved IPs are detected locally.
- **Microsoft Service Tag Lookup** — *Actions → Lookup MS Service Tags* downloads the official Azure IP Ranges JSON and matches destination IPs to Azure service names (e.g., `AzureActiveDirectory`, `Storage.WestUS2`).
- **Charts** — bar charts (action, protocol, direction, top destination ports) and pie charts (action, protocol, direction, top IP owners) drawn on WPF Canvas.
- **CSV Export** — *File → Export to CSV* saves all visible columns including IP owner and service tag info.

---

## How the Applications Work Together

### Scenario 1: Single Machine with Azure (Agent + Viewer)

1. Run `FirewallLogAgent.ps1` as Administrator on the machine you want to monitor.
2. Connect to Azure and start monitoring — logs flow into Azure Table Storage.
3. On any workstation, run `FirewallLogViewer.ps1`, connect to the same Azure storage account, and query/analyze the collected data.

### Scenario 2: Multiple Remote VMs with CSV (Collector + Viewer)

1. Deploy `FirewallLogCollector.ps1` to each VM using `Invoke-AzVMRunCommand` with `-Action Install`. This creates a scheduled task that runs at startup as SYSTEM.
2. CSV files accumulate at `C:\ProgramData\FirewallLogCollector\Logs\` on each VM.
3. Copy the CSV files to your workstation (manually, via Azure file share, etc.).
4. Open `FirewallLogViewer.ps1` and load each CSV via *File → Load CSV File*.

### Scenario 3: Multiple Machines with Azure (Agent on each + Viewer)

1. Run `FirewallLogAgent.ps1` in headless mode on each machine, all pointing to the same Azure Storage Account and table.
2. Each machine's logs are partitioned by `{ComputerName}_{Date}` in Azure Table Storage.
3. Use `FirewallLogViewer.ps1` to query all machines' data from the single table. Filter by computer name, IP, or process to focus on a specific host.

---

## Authentication

Both the Agent and Viewer support two authentication methods:

| Method | How it works |
|---|---|
| **Access Key** | Provide the storage account key directly. Requests are signed with `SharedKeyLite`. Keys are held in memory only — never written to disk. |
| **Azure AD (Device Code Flow)** | The app initiates a device code flow using the well-known Azure PowerShell public client ID (`1950a258-227b-4e31-a9cf-717495945fc2`). You authenticate in a browser, and an OAuth2 Bearer token is used for all subsequent requests. Tokens are refreshed automatically. The Agent GUI also acquires an ARM token to browse subscriptions and storage accounts. |

---

## Table Schema (Azure Table Storage)

Entries uploaded by the Agent follow this schema:

| Column | Type | Description |
|---|---|---|
| `PartitionKey` | String | `{ComputerName}_{Date}` |
| `RowKey` | String | `{Timestamp}_{GUID}` |
| `Date` | String | `yyyy-MM-dd` |
| `Time` | String | `HH:mm:ss` |
| `Action` | String | `ALLOW`, `DROP`, `INFO-EVENTS-LOST` |
| `Protocol` | String | `TCP`, `UDP`, `ICMP` |
| `SrcIP` / `DstIP` | String | Source / destination IP |
| `SrcPort` / `DstPort` | Int32 | Source / destination port |
| `Direction` | String | `SEND`, `RECEIVE` |
| `Size` | Int32 | Packet size |
| `ProcessId` | Int32 | Resolved PID (0 if unknown) |
| `ProcessName` | String | Executable name |
| `ComputerName` | String | Hostname |

The Collector's CSV files include additional columns: `RuleName` (Windows Firewall rule name resolved from the Security Event Log) and `EventAppName` (application path from the event).

---

## Project Structure

```
Windows Firewall Inspector/
├── README.md                              ← you are here
│
├── FirewallAgent/
│   ├── FirewallLogAgent.ps1               ← Agent entry point (GUI + headless)
│   ├── FirewallLogCollector.ps1           ← Collector entry point (CSV / scheduled task)
│   ├── README.md                          ← detailed Agent & Collector docs
│   └── Modules/
│       ├── AgentGUI.ps1                   ← WPF GUI for the Agent
│       ├── AzureAuth.ps1                  ← OAuth2 device code flow + access key signing
│       ├── AzureResourceManager.ps1       ← ARM REST: subscriptions, storage accounts, keys
│       ├── AzureTableStorage.ps1          ← Table Storage REST: create table, insert entities, query
│       └── FirewallLogManager.ps1         ← Enable/restore firewall logging, parse pfirewall.log
│
└── FirewallViewer/
    ├── FirewallLogViewer.ps1              ← Viewer entry point
    └── Modules/
        ├── AzureAuth.ps1                  ← OAuth2 device code flow + access key signing
        ├── AzureTableStorage.ps1          ← Table Storage REST: query entities
        ├── IPOwnerLookup.ps1              ← ip-api.com batch lookup with caching
        ├── ServiceTagLookup.ps1           ← Microsoft Azure IP Ranges / Service Tag matching
        └── ViewerGUI.ps1                  ← WPF dark-mode dashboard with charts
```

---

## Security Notes

- Access keys are kept in memory only; never written to disk.
- Device Code Flow uses a well-known public client ID — no app registration required.
- The Agent automatically restores original firewall logging settings on exit or stop.
- The Collector's scheduled task runs as SYSTEM — no user credentials stored.
- The firewall log file is read with `FileShare.ReadWrite` to avoid locking conflicts with the Windows Firewall service.
