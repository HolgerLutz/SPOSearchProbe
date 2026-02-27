# SPO Search Probe

A .NET 10 Windows Forms tool for monitoring **SharePoint Online search index freshness**. Authenticates via OAuth2 device code flow, executes scheduled KQL queries against the SharePoint REST API, tracks page indexing status, and produces detailed HTML session reports.

Published as a **single self-contained `.exe`** — no .NET runtime required on the target machine.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## Features

- **Two operating modes** — End-User (simple single-user GUI) and Admin (multi-user management, full config editor)
- **OAuth2 PKCE authentication** — Device code flow via browser; tokens encrypted with Windows DPAPI
- **Scheduled search probing** — Configurable interval (seconds or minutes); executes KQL queries against SPO REST API
- **Multi-user support** (Admin mode) — Run concurrent probes for multiple user accounts simultaneously
- **Page validation tracking** — Resolve a page URL to its WorkId and monitor when it appears in search results
- **Live execution timeline chart** — Real-time per-user stacked sub-charts with color-coded validation status
- **HTML session reports** — Standalone reports with per-user charts, statistics, and full log archives
- **Log collection** — One-click ZIP packaging of session logs, TSV data, and request/response traces
- **EndUser package creation** — Admin can create distribution ZIPs with optional pre-configured workspace URL
- **Request/response archiving** — Every API call is saved as a ZIP for diagnostic analysis
- **No external dependencies** — Zero NuGet packages; uses only .NET BCL types

---

## Quick Start

### End-User Mode

1. Receive the `SPOSearchProbe.exe` + `search-config.json` package from your administrator
2. Double-click `SPOSearchProbe.exe` to launch
3. Click **Login** and authenticate via browser
4. Click **Start** to begin scheduled search probing

### Admin Mode

```powershell
.\SPOSearchProbe.exe -admin
```

1. Configure the search parameters (Site URL, Query, Properties, etc.)
2. Add user accounts and authenticate them
3. Set the **Page URL** of the page to monitor
4. Click **Validate Page** to resolve the page's WorkId
5. Click **Start** to begin scheduled probing for all enabled users
6. Use **Collect Logs** to generate a ZIP archive with an HTML session report

---

## Command-Line Arguments

| Argument | Description |
|---|---|
| *(none)* | Launch in **End-User** mode (simple GUI) |
| `-admin` | Launch in **Admin** mode (full config editor, multi-user) |
| `-config <path>` | Use a custom `search-config.json` file |
| `-help` or `-?` | Show help dialog |

Accepts prefix styles: `-admin`, `--admin`, `/admin`

---

## Configuration

The `search-config.json` file controls all probe behavior. In **Admin mode**, all settings can be
configured directly in the GUI and are saved automatically — no manual file editing required.
The JSON reference below is for informational purposes or manual editing if preferred:

```json
{
    "siteUrl": "https://contoso.sharepoint.com/sites/MySite",
    "tenantId": "contoso.onmicrosoft.com",
    "queryText": "contentclass:STS_ListItem",
    "selectProperties": ["Title", "Path", "LastModifiedTime", "WorkId"],
    "rowLimit": 10,
    "sortList": "",
    "pageUrl": "",
    "intervalValue": 10,
    "intervalUnit": "seconds",
    "clientId": "9bc3ab49-b65d-410a-85ad-de819febfddc",
    "workspaceUrl": ""
}
```

| Field | Description |
|---|---|
| `siteUrl` | SharePoint Online site collection URL to search against |
| `tenantId` | Azure AD tenant name (e.g. `contoso.onmicrosoft.com`) or GUID |
| `queryText` | KQL (Keyword Query Language) search query |
| `selectProperties` | Managed properties to include in search results |
| `rowLimit` | Maximum result rows per query (1–500) |
| `sortList` | Optional sort order (e.g. `Created:descending`) |
| `pageUrl` | URL of a specific page to monitor for indexing |
| `intervalValue` / `intervalUnit` | Probe scheduling interval (seconds or minutes) |
| `clientId` | Azure AD app registration ID (default: PnP Management Shell) |
| `workspaceUrl` | Optional URL for log upload workspace |

---

## Admin Mode Workflow

### Button State Logic

1. **Start** is disabled until the page is validated
2. Click **Validate Page** to resolve the page URL → enables **Start**, disables **Validate Page**
3. Changing the **Page URL** text invalidates the current validation and re-enables **Validate Page**
4. **Reset** clears the validation and disables **Start** again
5. **Test Query** runs a single query without starting the scheduler

### Creating EndUser Packages

In Admin mode, click **📦 Create EndUser Package** to:

1. Optionally enter a workspace URL (for log upload destinations)
2. Choose a save location
3. The tool creates a ZIP containing the exe + a clean `search-config.json` (admin user accounts are automatically stripped)

---

## Building from Source

For detailed build instructions (including manual build steps), see the **[Build & Deployment Guide](build-guide.md)**.

### Quick Build

The `Build.ps1` script checks prerequisites (offers to install .NET 10 SDK via winget or automated download), then builds both win-x64 and win-arm64 with versioned distribution ZIPs in `dist/`.

```powershell
.\Build.ps1                  # Full build with prerequisite checks
.\Build.ps1 -SkipPrereqs    # Skip prerequisite checks
```

| Requirement | Version |
|---|---|
| **Windows** | 10 (1809+) or 11 |
| **.NET 10 SDK** | 10.0.x (Preview) — installed automatically by `Build.ps1` if missing |
| **PowerShell** | 5.1+ |

> **Note:** End users do **not** need the .NET SDK — the published binary is fully self-contained.

---

## Project Structure

```
SPOSearchProbe/
├── SPOSearchProbe/
│   ├── SPOSearchProbe.csproj       # .NET 10 WinForms project
│   ├── Program.cs                  # Entry point, CLI arg parsing, mode selection
│   ├── SearchConfig.cs             # JSON config model (search-config.json)
│   ├── TokenCache.cs               # DPAPI-encrypted token persistence
│   ├── OAuthHelper.cs              # OAuth2 PKCE device code flow
│   ├── SearchClient.cs             # SharePoint REST API search client
│   ├── MainForm.cs                 # End-User mode GUI
│   ├── AdminForm.cs                # Admin mode GUI (multi-user, config editor)
│   ├── LogCollector.cs             # Log packaging + HTML report generation
│   ├── LiveChartForm.cs            # Real-time execution timeline chart
│   └── search-config.json          # Configuration template
├── Build.ps1                       # Release build with prerequisite checks
├── LICENSE                         # MIT License
├── README.md                       # This file
├── build-guide.md                  # Detailed build & deployment guide
└── .github/
    └── copilot-instructions.md     # GitHub Copilot coding conventions
```

---

## Architecture

```
┌────────────────────────────────────────────────┐
│              SPOSearchProbe.exe                │
│  ┌──────────────┐  ┌───────────────────────┐   │
│  │  AdminForm   │  │  MainForm (End-User)  │   │
│  │  Multi-user  │  │  Single-user, simple  │   │
│  │  Config edit │  │  Config from JSON     │   │
│  └──────┬───────┘  └──────────┬────────────┘   │
│         │                     │                │
│  ┌──────┴─────────────────────┴─────────────┐  │
│  │           Core Services                  │  │
│  │  OAuthHelper    (PKCE + token refresh)   │  │
│  │  TokenCache     (DPAPI encryption)       │  │
│  │  SearchClient   (SPO REST API)           │  │
│  │  LogCollector   (ZIP + HTML reports)     │  │
│  │  LiveChartForm  (Real-time chart)        │  │
│  └──────────────────────────────────────────┘  │
│                                                │
│  search-config.json  (companion file)          │
└────────────────────────────────────────────────┘
```

### Data Flow

1. User authenticates → OAuth2 device code flow opens browser → tokens received
2. Tokens encrypted via DPAPI → stored as `.token-{email}.dat`
3. On each timer tick, token auto-refreshed if within 5 min of expiry
4. `SearchClient` calls SPO REST API `/_api/search/query` → parses results
5. Results logged to `.tsv` + `.log` file; each request/response archived as ZIP
6. Page validation: resolves URL to WorkId, then monitors if WorkId appears in results

---

## License

This project is licensed under the [MIT License](LICENSE).

**Disclaimer:** This tool stores OAuth tokens encrypted with Windows DPAPI. Token security is tied to the Windows user account. The authors are not liable for any data loss, security incidents, or service disruptions. Use at your own risk.
