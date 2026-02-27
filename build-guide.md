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

The included `Build.ps1` script handles everything automatically — it checks for prerequisites,
offers to install the .NET 10 SDK if missing (via winget or the official `dotnet-install.ps1`
script), and then builds self-contained single-file EXEs for both win-x64 and win-arm64.

```powershell
.\Build.ps1

# Skip prerequisite checks if SDK is already installed
.\Build.ps1 -SkipPrereqs

# Build and launch the app in admin mode afterwards
.\Build.ps1 -Launch
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

That's it — see the **[README](README.md)** for configuration and usage instructions.

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
├── Build.ps1                       # Build script with prerequisite checks
├── LICENSE                         # MIT License
├── README.md                       # Project overview and usage guide
├── build-guide.md                  # This file
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

## Next Steps

After building, see the **[README](README.md)** for configuration, usage instructions, and Admin mode workflow.

---

## Architecture Summary

```
┌─────────────────────────────────────────────┐
│              SPOSearchProbe.exe             │
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
