# Firewall Log Agent & Viewer

Two self-contained PowerShell programs for collecting Windows Firewall logs and visualizing them via Azure Table Storage. **No Azure PowerShell modules or Azure CLI required** — all Azure communication uses pure REST API calls.

---

## Architecture

```
FirewallAgent/                        FirewallViewer/
├── FirewallLogAgent.ps1  (entry)     ├── FirewallLogViewer.ps1  (entry)
└── Modules/                          └── Modules/
    ├── AzureAuth.ps1                     ├── AzureAuth.ps1
    ├── AzureTableStorage.ps1             ├── AzureTableStorage.ps1
    ├── FirewallLogManager.ps1            ├── IPOwnerLookup.ps1
    └── AgentGUI.ps1                      └── ViewerGUI.ps1
```

## Prerequisites

| Requirement | Details |
|---|---|
| **OS** | Windows 10 / 11 (Session Host compatible) |
| **PowerShell** | 5.1+ (Windows PowerShell) |
| **Elevation** | Agent requires **Run as Administrator** to enable firewall logging |
| **Azure Storage Account** | With Table Storage enabled |
| **Authentication** | Storage Account **Access Key**, or Azure AD account with **Storage Table Data Contributor** RBAC role |

---

## Program 1 — Firewall Log Agent

### What it does

1. **Enables** Windows Firewall logging on all profiles (domain/private/public) at launch.
2. Reads the firewall log (`pfirewall.log`) every 60 seconds for new entries.
3. **Resolves PIDs** to executable names by correlating active TCP/UDP connections.
4. **Uploads** parsed entries to Azure Table Storage via REST API.
5. **Restores** original firewall logging settings on exit.

### Usage

#### GUI mode (default)
```powershell
# Right-click PowerShell → Run as Administrator
.\FirewallAgent\FirewallLogAgent.ps1
```

#### Headless / no-GUI mode
```powershell
# With Access Key
.\FirewallAgent\FirewallLogAgent.ps1 -NoGUI -StorageAccount "myaccount" -AccessKey "base64key=="

# With Azure AD (device code flow)
.\FirewallAgent\FirewallLogAgent.ps1 -NoGUI -StorageAccount "myaccount"
```

#### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-NoGUI` | `$false` | Run without the WPF interface |
| `-StorageAccount` | — | Storage account name (required for NoGUI) |
| `-AccessKey` | — | Storage access key (omit to use Azure AD) |
| `-TableName` | `FirewallLogs` | Azure Table name |
| `-IntervalSeconds` | `60` | Upload interval |
| `-TenantId` | `common` | Azure AD tenant for device code flow |

### Authentication Flow

1. **Access Key** — provide the storage account key; requests are signed with `SharedKeyLite`.
2. **Azure AD (RBAC fallback)** — initiates a Device Code Flow using the well-known Azure PowerShell client ID. The user authenticates in a browser, and an OAuth2 Bearer token is used for subsequent requests. Tokens are automatically refreshed.

### Table Schema

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

---

## Program 2 — Firewall Log Viewer

### What it does

1. **Connects** to Azure Table Storage (same auth options as the Agent).
2. **Queries** firewall log data with OData filters + client-side filtering.
3. **Displays** results in a sortable data grid (dark-mode WPF).
4. **Looks up IP owners** via RDAP / ARIN WHOIS REST APIs with caching.
5. **Visualizes** data with bar charts and pie charts drawn on WPF Canvas.
6. **Exports** filtered results to CSV.

### Usage

```powershell
.\FirewallViewer\FirewallLogViewer.ps1
```

### Features

| Feature | Details |
|---|---|
| **Dark Mode** | Deep dark theme with purple accent (`#7C3AED`) |
| **Filters** | Date range, Action, Protocol, Direction, IP, Port, Process name |
| **Data Grid** | Sortable columns, alternating row colors, IP owner columns |
| **Bar Charts** | By action, protocol, direction, top destination ports |
| **Pie Charts** | Action distribution, protocol, direction, top IP owners |
| **IP Lookup** | RDAP → ARIN fallback, private IP detection, result caching |
| **CSV Export** | All visible columns including IP owner info |

---

## Azure Storage Setup

1. Create a **Storage Account** in the Azure Portal.
2. Enable **Table Storage** (included by default in general-purpose v2 accounts).
3. For **Access Key** auth: copy a key from *Storage Account → Access Keys*.
4. For **RBAC** auth: assign the **Storage Table Data Contributor** role to your Azure AD user on the storage account.

---

## Security Notes

- Access keys are kept in memory only; never written to disk.
- Device Code Flow uses the well-known Azure PowerShell public client ID (`1950a258-227b-4e31-a9cf-717495945fc2`).
- Firewall logging changes are automatically reverted when the Agent exits or is stopped.
- The Agent reads the firewall log via `FileShare.ReadWrite` to avoid locking conflicts.
