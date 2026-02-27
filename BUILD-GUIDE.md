# SPO Search Probe Tool — Build & Deployment Guide

## Overview

**SPO Search Probe** is a .NET 10 Windows Forms application for monitoring SharePoint Online search
index freshness. It authenticates via OAuth2 PKCE (device code flow through browser), executes
scheduled search queries against the SharePoint REST API, and logs results with full
request/response archiving.

The application publishes as a **single self-contained `.exe`** — no .NET runtime installation is
required on the target machine. The only companion file is `search-config.json` which the
administrator edits before distribution.

---

## Prerequisites

| Requirement | Version | Purpose |
|---|---|---|
| **Windows** | 10 (1809+) or 11 | WinForms GUI + DPAPI token encryption |
| **.NET 10 SDK** | 10.0.x (Preview) | Build and publish the application |
| **PowerShell** | 5.1+ or 7.x | Run the `Build.ps1` build script |

> **Note:** The `Build.ps1` script automatically checks for prerequisites and offers to install
> the .NET 10 SDK if missing (via winget or the official `dotnet-install.ps1` script).

> **Note:** The end user running the compiled `.exe` does **not** need the .NET SDK or runtime
> installed — the published binary is fully self-contained.

---

## Step 1 — Install .NET 10 SDK

> **Recommended:** Just run `.\Build.ps1` — it checks for prerequisites and offers to install the
> .NET 10 SDK automatically. The manual options below are only needed if you prefer to install
> the SDK yourself.

### Option A: winget (Windows Package Manager)

```powershell
winget install Microsoft.DotNet.SDK.Preview
dotnet --version
```

### Option B: PowerShell (Automated — works on servers without winget)

```powershell
# Download and run the official dotnet-install script
Invoke-WebRequest -Uri "https://dot.net/v1/dotnet-install.ps1" -OutFile "$env:TEMP\dotnet-install.ps1"

# Install .NET 10 SDK (latest preview)
& "$env:TEMP\dotnet-install.ps1" -Channel 10.0 -InstallDir "$env:ProgramFiles\dotnet"

# Verify installation
dotnet --version
```

### Option C: Manual Download

1. Go to [https://dotnet.microsoft.com/download/dotnet/10.0](https://dotnet.microsoft.com/download/dotnet/10.0)
2. Under **SDK**, download the installer for your architecture:
   - **x64** → `dotnet-sdk-10.0.xxx-win-x64.exe`
   - **ARM64** → `dotnet-sdk-10.0.xxx-win-arm64.exe`
3. Run the installer and follow the prompts
4. Verify:
   ```powershell
   dotnet --version
   # Expected: 10.0.xxx
   ```

---

## Step 2 — Verify the Solution Structure

After extracting or cloning the source, you should have:

```
SPOSearchProbe\
├── Build.ps1                       # Build script with prerequisite checks (PowerShell)
├── Build.bat                       # Build script with prerequisite checks (Command Prompt)
├── LICENSE                         # MIT License
├── README.md                       # Project overview and usage guide
├── BUILD-GUIDE.md                  # This file
├── .github\
│   └── copilot-instructions.md     # Copilot coding conventions
└── SPOSearchProbe\
    ├── SPOSearchProbe.csproj       # Project file (.NET 10 WinForms)
    ├── Program.cs                  # Entry point (mode selection)
    ├── SearchConfig.cs             # Configuration model
    ├── TokenCache.cs               # DPAPI-encrypted token storage
    ├── OAuthHelper.cs              # OAuth2 PKCE authentication
    ├── SearchClient.cs             # SharePoint Search REST API client
    ├── MainForm.cs                 # End-user GUI
    ├── AdminForm.cs                # Admin GUI (multi-user, config editor)
    ├── LogCollector.cs             # Log packaging (ZIP + HTML report)
    ├── LiveChartForm.cs            # Real-time execution timeline chart
    └── search-config.json          # Template configuration file
```

---

## Step 3 — Build the Application

### Option A: Build Script *(Recommended)*

The `Build.ps1` script checks all prerequisites (offers to install .NET 10 SDK if missing),
then builds self-contained single-file EXEs for both win-x64 and win-arm64:

```powershell
# PowerShell — build both platforms with versioned ZIPs
.\Build.ps1

# Command Prompt alternative
.\Build.bat

# Skip prerequisite checks if SDK is already installed
.\Build.ps1 -SkipPrereqs
```

This produces distribution ZIPs in `dist\`:

```
dist\
├── SPOSearchProbe-v1.26.226.1-win-x64.zip
└── SPOSearchProbe-v1.26.226.1-win-arm64.zip
```

Each ZIP contains:
```
├── SPOSearchProbe.exe          # ~48 MB, fully self-contained
└── search-config.json          # Configuration template
```

### Option B: Manual Debug Build

```powershell
cd SPOSearchProbe
dotnet build
```

This compiles to `bin\Debug\net10.0-windows\` — useful for development and testing.
Requires .NET 10 runtime on the machine to run.

### What the Publish Does

The script runs the following `dotnet publish` command under the hood:

```powershell
dotnet publish SPOSearchProbe -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=none
```

| Flag | Purpose |
|---|---|
| `--self-contained true` | Bundles the entire .NET runtime — no install needed on target |
| `PublishSingleFile=true` | Merges all DLLs into one `.exe` |
| `IncludeNativeLibrariesForSelfExtract=true` | Includes native libs inside the single file |
| `EnableCompressionInSingleFile=true` | Compresses bundled assemblies to reduce size |
| `DebugType=none` | Excludes `.pdb` debug symbols from output |

---

## Step 4 — Configure Before Distribution

Edit `search-config.json` before sending to end users:

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
| `siteUrl` | SharePoint site collection URL to search against |
| `tenantId` | Entra ID tenant (e.g. `contoso.onmicrosoft.com` or a GUID) |
| `queryText` | KQL search query |
| `selectProperties` | Managed properties to return |
| `rowLimit` | Max results per query (1–500) |
| `sortList` | Sort order (e.g. `Created:descending`) |
| `pageUrl` | Optional page URL for validation monitoring |
| `intervalValue` / `intervalUnit` | Scheduling interval |
| `clientId` | Azure AD app registration (default: PnP Management Shell) |
| `workspaceUrl` | Optional URL for log upload workspace |

---

## Step 5 — Run the Application

1. Extract `SPOSearchProbe-win-x64.zip` to a folder
2. Edit `search-config.json` with your tenant and search parameters
3. Double-click `SPOSearchProbe.exe` → launches **End-User mode** (simple GUI)

### Launch Modes

| Mode | How to launch | Description |
|---|---|---|
| **End-User** (default) | `SPOSearchProbe.exe` | Simple GUI: login, run search, view results |
| **Admin** | `SPOSearchProbe.exe -admin` | Full GUI: multi-user management, page validation, config editor, package creation |

To launch Admin mode, either:
- Open a terminal and run: `.\SPOSearchProbe.exe -admin`
- Create a shortcut and add `-admin` to the Target field
- Accepts `-admin`, `--admin`, or `/admin`

> **Tip:** In Admin mode, the config file is auto-created if missing. In End-User mode,
> a pre-configured `search-config.json` must exist next to the `.exe`.

---

## Architecture Summary

```
┌─────────────────────────────────────────────┐
│              SPOSearchProbe.exe              │
│  ┌────────────┐  ┌────────────────────────┐ │
│  │ AdminForm  │  │  MainForm (End-User)   │ │
│  │ Multi-user │  │  Single-user, config   │ │
│  │ Config edit│  │  embedded from JSON    │ │
│  └─────┬──────┘  └──────────┬─────────────┘ │
│        │                    │               │
│  ┌─────┴────────────────────┴─────────────┐ │
│  │         Core Services                  │ │
│  │  OAuthHelper   (PKCE + refresh)        │ │
│  │  TokenCache    (DPAPI encryption)      │ │
│  │  SearchClient  (SPO REST API)          │ │
│  │  LogCollector  (ZIP + HTML reports)    │ │
│  └────────────────────────────────────────┘ │
│                                             │
│  search-config.json  (companion file)       │
└─────────────────────────────────────────────┘
```
