# Build.ps1 - Check prerequisites, install if needed, and build SPO Search Probe
#
# Usage:
#   .\Build.ps1              # Full build (x64 + arm64) with distribution ZIPs
#   .\Build.ps1 -SkipPrereqs # Skip prerequisite checks
#
# Prerequisites checked:
#   - .NET 10 SDK (offers install via winget or dotnet-install.ps1 download)
#   - PowerShell 5.1+ (informational only — you're already running it)

param(
    [switch]$SkipPrereqs
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$projectDir = Join-Path $scriptDir "SPOSearchProbe"
$distDir = Join-Path $scriptDir "dist"
$runtimes = @("win-x64", "win-arm64")

# =====================================================================
# Prerequisites
# =====================================================================
function Check-Prerequisites {
    Write-Host ""
    Write-Host "=== Checking Prerequisites ===" -ForegroundColor Cyan

    # --- PowerShell version ---
    $psVer = $PSVersionTable.PSVersion
    if ($psVer.Major -ge 5) {
        Write-Host "  [OK] PowerShell $psVer" -ForegroundColor Green
    } else {
        Write-Host "  [WARN] PowerShell $psVer detected. 5.1+ recommended." -ForegroundColor Yellow
    }

    # --- .NET SDK ---
    $dotnetOk = $false
    try {
        $dotnetVer = & dotnet --version 2>$null
        if ($dotnetVer -and $dotnetVer -match "^10\.") {
            Write-Host "  [OK] .NET SDK $dotnetVer" -ForegroundColor Green
            $dotnetOk = $true
        } elseif ($dotnetVer) {
            Write-Host "  [WARN] .NET SDK $dotnetVer found, but .NET 10.x is required." -ForegroundColor Yellow
        }
    } catch { }

    if (-not $dotnetOk) {
        Write-Host "  [MISSING] .NET 10 SDK is required to build this project." -ForegroundColor Red
        Write-Host ""
        Write-Host "  Install options:" -ForegroundColor White
        Write-Host "    1) winget (if available)" -ForegroundColor Gray
        Write-Host "    2) Automatic download via dotnet-install.ps1 (works on servers)" -ForegroundColor Gray
        Write-Host "    3) Manual download from https://dotnet.microsoft.com/download/dotnet/10.0" -ForegroundColor Gray
        Write-Host ""

        $choice = Read-Host "  Install .NET 10 SDK now? (Y/N)"
        if ($choice -match "^[Yy]") {
            # Try winget first, fall back to dotnet-install.ps1
            $hasWinget = $null -ne (Get-Command winget -ErrorAction SilentlyContinue)
            if ($hasWinget) {
                Write-Host "  Installing .NET 10 SDK via winget..." -ForegroundColor Cyan
                winget install Microsoft.DotNet.SDK.Preview --accept-package-agreements --accept-source-agreements
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "  [OK] .NET 10 SDK installed via winget." -ForegroundColor Green
                    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
                    $dotnetOk = $true
                } else {
                    Write-Host "  [WARN] winget install failed, trying dotnet-install.ps1..." -ForegroundColor Yellow
                    $hasWinget = $false  # fall through to script install
                }
            }

            if (-not $hasWinget -and -not $dotnetOk) {
                Write-Host "  Downloading official dotnet-install.ps1 ..." -ForegroundColor Cyan
                $installScript = Join-Path $env:TEMP "dotnet-install.ps1"
                try {
                    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                    Invoke-WebRequest -Uri "https://dot.net/v1/dotnet-install.ps1" -OutFile $installScript -UseBasicParsing
                    Write-Host "  Running dotnet-install.ps1 -Channel 10.0 ..." -ForegroundColor Cyan
                    & $installScript -Channel 10.0 -InstallDir "$env:ProgramFiles\dotnet"
                    # Refresh PATH for this session
                    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("PATH", "User")
                    if (-not ($env:PATH -like "*$env:ProgramFiles\dotnet*")) {
                        $env:PATH = "$env:ProgramFiles\dotnet;$env:PATH"
                    }
                    $checkVer = & dotnet --version 2>$null
                    if ($checkVer -and $checkVer -match "^10\.") {
                        Write-Host "  [OK] .NET SDK $checkVer installed via dotnet-install.ps1" -ForegroundColor Green
                        $dotnetOk = $true
                    } else {
                        Write-Host "  [WARN] dotnet-install.ps1 completed but .NET 10 not detected (got: $checkVer)." -ForegroundColor Yellow
                        Write-Host "         You may need to restart your terminal or install manually." -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "  [FAIL] Download failed: $_" -ForegroundColor Red
                    Write-Host "         Please install .NET 10 SDK manually:" -ForegroundColor Red
                    Write-Host "         https://dotnet.microsoft.com/download/dotnet/10.0" -ForegroundColor Yellow
                }
            }
        } else {
            Write-Host "  Skipping .NET SDK install." -ForegroundColor Yellow
        }

        if (-not $dotnetOk) {
            throw ".NET 10 SDK is required. Please install it and try again."
        }
    }

    # --- Verify dotnet can find the project ---
    if (-not (Test-Path (Join-Path $projectDir "SPOSearchProbe.csproj"))) {
        throw "Project not found at $projectDir. Run this script from the repository root."
    }
    Write-Host "  [OK] Project found: SPOSearchProbe.csproj" -ForegroundColor Green
    Write-Host ""
}

if (-not $SkipPrereqs) {
    Check-Prerequisites
}

# =====================================================================
# Version
# =====================================================================
$now = Get-Date
$yy = $now.ToString("yy")
$m = $now.Month.ToString()
$dd = $now.ToString("dd")
$dateKey = $now.ToString("yyyyMMdd")

$counterFile = Join-Path $scriptDir ".build-counter"
$buildNum = 1
if (Test-Path $counterFile) {
    $lines = Get-Content $counterFile
    if ($lines.Count -ge 2 -and $lines[0] -eq $dateKey) {
        $buildNum = [int]$lines[1] + 1
    }
}
Set-Content $counterFile "$dateKey`n$buildNum"
$version = "1.$yy.$($m)$($dd).$buildNum"
Write-Host "Version: v$version" -ForegroundColor Magenta

# =====================================================================
# Stop running instances
# =====================================================================
$procs = Get-Process -Name "SPOSearchProbe" -ErrorAction SilentlyContinue
if ($procs) {
    Write-Host "Stopping $($procs.Count) running SPOSearchProbe instance(s)..." -ForegroundColor Yellow
    foreach ($p in $procs) {
        try { Stop-Process -Id $p.Id -Force; Write-Host "  Stopped PID $($p.Id)" -ForegroundColor Yellow }
        catch { Write-Host "  Could not stop PID $($p.Id): $_" -ForegroundColor Red }
    }
    Start-Sleep -Seconds 2
}

# =====================================================================
# Backup config + tokens
# =====================================================================
$testDir = Join-Path (Split-Path $scriptDir -Parent) "test"
if (-not (Test-Path $testDir)) { New-Item -Path $testDir -ItemType Directory -Force | Out-Null }
Write-Host "Backing up config and tokens to $testDir ..." -ForegroundColor Cyan
$configBacked = $false
foreach ($rt in $runtimes) {
    $pubDir = Join-Path (Join-Path (Join-Path $projectDir "bin") "publish") $rt
    if (-not (Test-Path $pubDir)) { continue }

    if (-not $configBacked) {
        $cfgFile = Join-Path $pubDir "search-config.json"
        if (Test-Path $cfgFile) {
            Copy-Item $cfgFile $testDir -Force
            Write-Host "  Backed up config: $cfgFile" -ForegroundColor DarkGray
            $configBacked = $true
        }
    }

    Get-ChildItem $pubDir -Filter "*.dat" -ErrorAction SilentlyContinue | ForEach-Object {
        Copy-Item $_.FullName $testDir -Force
        Write-Host "  Backed up token: $($_.Name)" -ForegroundColor DarkGray
    }
}

# =====================================================================
# Clean + Build
# =====================================================================
Write-Host "Cleaning build directories..." -ForegroundColor Cyan
foreach ($dir in @("bin", "obj")) {
    $path = Join-Path $projectDir $dir
    if (Test-Path $path) {
        Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed $dir/" -ForegroundColor DarkGray
    }
}

if (-not (Test-Path $distDir)) { New-Item -Path $distDir -ItemType Directory -Force | Out-Null }

foreach ($runtime in $runtimes) {
    Write-Host ""
    Write-Host "=== Publishing v$version for $runtime ===" -ForegroundColor Cyan

    $publishDir = Join-Path (Join-Path (Join-Path $projectDir "bin") "publish") $runtime
    dotnet publish $projectDir -c Release -r $runtime --self-contained true `
        -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true `
        -p:EnableCompressionInSingleFile=true -p:DebugType=none `
        -p:Version=$version -p:InformationalVersion="v$version" `
        -o $publishDir

    if ($LASTEXITCODE -ne 0) { throw "Publish failed for $runtime." }

    # Create distribution zip
    $zipName = "SPOSearchProbe-v$version-$runtime.zip"
    $zipPath = Join-Path $distDir $zipName
    if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

    $tempStaging = Join-Path $scriptDir "staging-$runtime"
    if (Test-Path $tempStaging) { Remove-Item $tempStaging -Recurse -Force }
    New-Item -Path $tempStaging -ItemType Directory -Force | Out-Null

    Copy-Item (Join-Path $publishDir "SPOSearchProbe.exe") $tempStaging
    Copy-Item (Join-Path $projectDir "search-config.json") $tempStaging

    Compress-Archive -Path (Join-Path $tempStaging "*") -DestinationPath $zipPath -Force
    Remove-Item $tempStaging -Recurse -Force

    Write-Host "Package created: $zipPath" -ForegroundColor Green
}

Write-Host ""
Write-Host "=== Build complete - v$version ===" -ForegroundColor Green

# =====================================================================
# Restore config + tokens
# =====================================================================
Write-Host "Restoring config and tokens from $testDir ..." -ForegroundColor Cyan
foreach ($runtime in $runtimes) {
    $pubDir = Join-Path (Join-Path (Join-Path $projectDir "bin") "publish") $runtime
    $cfgSrc = Join-Path $testDir "search-config.json"
    if (Test-Path $cfgSrc) {
        Copy-Item $cfgSrc $pubDir -Force
        Write-Host "  Restored search-config.json -> $runtime" -ForegroundColor DarkGray
    }
    Get-ChildItem $testDir -Filter "*.dat" -ErrorAction SilentlyContinue | ForEach-Object {
        Copy-Item $_.FullName $pubDir -Force
        Write-Host "  Restored $($_.Name) -> $runtime" -ForegroundColor DarkGray
    }
}

# =====================================================================
# Summary
# =====================================================================
Write-Host ""
Write-Host "Distribution packages:" -ForegroundColor Cyan
foreach ($runtime in $runtimes) {
    $zp = Join-Path $distDir "SPOSearchProbe-v$version-$runtime.zip"
    if (Test-Path $zp) {
        $sz = [math]::Round((Get-Item $zp).Length / 1MB, 1)
        Write-Host "  $zp ($sz MB)" -ForegroundColor White
    }
}
Write-Host "Contents: SPOSearchProbe.exe + search-config.json" -ForegroundColor White

# --- Launch the matching build in admin mode ---
$arch = if ($env:PROCESSOR_ARCHITECTURE -eq "ARM64") { "win-arm64" } else { "win-x64" }
$launchExe = Join-Path (Join-Path (Join-Path (Join-Path $projectDir "bin") "publish") $arch) "SPOSearchProbe.exe"
if (Test-Path $launchExe) {
    Write-Host ""
    Write-Host "Launching SPOSearchProbe ($arch) in admin mode..." -ForegroundColor Magenta
    Start-Process $launchExe -ArgumentList "-admin"
}
