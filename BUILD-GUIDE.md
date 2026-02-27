# SPO Search Probe Tool — Build & Deployment Guide

## Overview

**SPO Search Probe** is a .NET 10 Windows Forms application for monitoring SharePoint Online search
index freshness. It authenticates via OAuth2 PKCE (device code flow through browser), executes
scheduled search queries against the SharePoint REST API, and logs results with full
request/response archiving.

The application publishes as a **single self-contained `.exe`** — no .NET runtime installation is
required on the target machine. The only companion file is `search-config.json` which the
administrator edits before distribution.

> **Note:** The end user running the compiled `.exe` does **not** need the .NET SDK or runtime
> installed — the published binary is fully self-contained.

---

## Option 1: Build Script *(Recommended)*

The included `Build.ps1` (PowerShell) and `Build.bat` (Command Prompt) scripts handle everything
automatically — they check for prerequisites, offer to install the .NET 10 SDK if missing
(via winget or the official `dotnet-install.ps1` script), and then build self-contained
single-file EXEs for both win-x64 and win-arm64.

```powershell
# PowerShell
.\Build.ps1

# Command Prompt
.\Build.bat

# Skip prerequisite checks if SDK is already installed
.\Build.ps1 -SkipPrereqs
```

Output:

```
dist\
├── SPOSearchProbe-v1.26.226.1-win-x64.zip      # Intel/AMD 64-bit
└── SPOSearchProbe-v1.26.226.1-win-arm64.zip     # ARM64
```

Each ZIP contains:
```
├── SPOSearchProbe.exe          # ~48 MB, fully self-contained
└── search-config.json          # Configuration template
```

That's it — proceed to [Configure Before Distribution](#configure-before-distribution) below.

---

## Option 2: Manual Build

If you prefer to install prerequisites and build manually, follow these steps.

### Step 1 — Install .NET 10 SDK

| Requirement | Version | Purpose |
|---|---|---|
| **Windows** | 10 (1809+) or 11 | WinForms GUI + DPAPI token encryption |
| **.NET 10 SDK** | 10.0.x (Preview) | Build and publish the application |
| **PowerShell** | 5.1+ or 7.x | ZIP packaging in the build script |

Install the .NET 10 SDK using one of these methods:

**winget (Windows Package Manager):**

```powershell
winget install Microsoft.DotNet.SDK.Preview
dotnet --version
```

**PowerShell (works on servers without winget):**

```powershell
Invoke-WebRequest -Uri "https://dot.net/v1/dotnet-install.ps1" -OutFile "$env:TEMP\dotnet-install.ps1"
& "$env:TEMP\dotnet-install.ps1" -Channel 10.0 -InstallDir "$env:ProgramFiles\dotnet"
dotnet --version
```

**Manual download:**

1. Go to [https://dotnet.microsoft.com/download/dotnet/10.0](https://dotnet.microsoft.com/download/dotnet/10.0)
2. Download the SDK installer for your architecture (x64 or ARM64)
3. Run the installer, then verify: `dotnet --version`

### Step 2 — Verify the Solution Structure

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

### Step 3 — Build

**Debug build** (requires .NET 10 runtime on the machine to run):

```powershell
cd SPOSearchProbe
dotnet build
```

Output: `bin\Debug\net10.0-windows\`

**Release build** (self-contained single-file EXE):

```powershell
dotnet publish SPOSearchProbe -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=none
```

Repeat with `-r win-arm64` for ARM64 builds.

| Flag | Purpose |
|---|---|
| `--self-contained true` | Bundles the entire .NET runtime — no install needed on target |
| `PublishSingleFile=true` | Merges all DLLs into one `.exe` |
| `IncludeNativeLibrariesForSelfExtract=true` | Includes native libs inside the single file |
| `EnableCompressionInSingleFile=true` | Compresses bundled assemblies to reduce size |
| `DebugType=none` | Excludes `.pdb` debug symbols from output |

---

## Configure Before Distribution

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

## Run the Application

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
