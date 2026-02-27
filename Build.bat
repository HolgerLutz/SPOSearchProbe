@echo off
setlocal enabledelayedexpansion
title SPO Search Probe - Build
echo.
echo ============================================
echo   SPO Search Probe - Build Script
echo ============================================

set "SCRIPTDIR=%~dp0"
set "PROJECTDIR=%SCRIPTDIR%SPOSearchProbe"
set "DISTDIR=%SCRIPTDIR%dist"
set "COUNTERFILE=%SCRIPTDIR%.build-counter"

:: =====================================================================
:: Prerequisites
:: =====================================================================
echo.
echo === Checking Prerequisites ===

:: --- PowerShell ---
where powershell >nul 2>&1
if %ERRORLEVEL%==0 (
    echo   [OK] PowerShell found
) else (
    echo   [FAIL] PowerShell not found. Required for ZIP packaging.
    pause
    exit /b 1
)

:: --- .NET SDK ---
set "DOTNET_OK=0"
where dotnet >nul 2>&1
if %ERRORLEVEL%==0 (
    for /f "tokens=*" %%V in ('dotnet --version 2^>nul') do set "DOTNET_VER=%%V"
    echo !DOTNET_VER! | findstr /B "10." >nul 2>&1
    if !ERRORLEVEL!==0 (
        echo   [OK] .NET SDK !DOTNET_VER!
        set "DOTNET_OK=1"
    ) else (
        echo   [WARN] .NET SDK !DOTNET_VER! found, but .NET 10.x is required.
    )
)

if !DOTNET_OK!==0 (
    echo   [MISSING] .NET 10 SDK is required to build this project.
    echo.
    echo   Install options:
    echo     1^) winget ^(if available^)
    echo     2^) Automatic download via dotnet-install.ps1 ^(works on servers^)
    echo     3^) Manual download from https://dotnet.microsoft.com/download/dotnet/10.0
    echo.
    set /p "INSTALL_CHOICE=  Install .NET 10 SDK now? (Y/N): "
    if /I "!INSTALL_CHOICE!"=="Y" (
        :: Try winget first
        where winget >nul 2>&1
        if !ERRORLEVEL!==0 (
            echo   Installing .NET 10 SDK via winget...
            winget install Microsoft.DotNet.SDK.Preview --accept-package-agreements --accept-source-agreements
            if !ERRORLEVEL!==0 (
                echo   [OK] .NET 10 SDK installed via winget.
                echo   Refreshing PATH...
                for /f "tokens=2*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYSPATH=%%B"
                for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USRPATH=%%B"
                set "PATH=!SYSPATH!;!USRPATH!"
                set "DOTNET_OK=1"
            ) else (
                echo   [WARN] winget install failed, trying dotnet-install.ps1...
            )
        ) else (
            echo   winget not available, using dotnet-install.ps1...
        )

        if !DOTNET_OK!==0 (
            echo   Downloading official dotnet-install.ps1 ...
            powershell -NoProfile -ExecutionPolicy Bypass -Command ^
                "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://dot.net/v1/dotnet-install.ps1' -OutFile '%TEMP%\dotnet-install.ps1' -UseBasicParsing"
            if exist "%TEMP%\dotnet-install.ps1" (
                echo   Running dotnet-install.ps1 -Channel 10.0 ...
                powershell -NoProfile -ExecutionPolicy Bypass -File "%TEMP%\dotnet-install.ps1" -Channel 10.0 -InstallDir "%ProgramFiles%\dotnet"
                :: Refresh PATH from registry + ensure dotnet dir is included
                for /f "tokens=2*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYSPATH=%%B"
                for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USRPATH=%%B"
                set "PATH=!SYSPATH!;!USRPATH!"
                echo !PATH! | findstr /I /C:"\dotnet" >nul 2>&1
                if !ERRORLEVEL! neq 0 set "PATH=%ProgramFiles%\dotnet;!PATH!"
                :: Verify
                dotnet --version >nul 2>&1
                if !ERRORLEVEL!==0 (
                    for /f "tokens=*" %%V in ('dotnet --version 2^>nul') do set "DOTNET_VER=%%V"
                    echo !DOTNET_VER! | findstr /B "10." >nul 2>&1
                    if !ERRORLEVEL!==0 (
                        echo   [OK] .NET SDK !DOTNET_VER! installed via dotnet-install.ps1
                        set "DOTNET_OK=1"
                    ) else (
                        echo   [WARN] dotnet-install.ps1 completed but .NET 10 not detected ^(got: !DOTNET_VER!^).
                    )
                ) else (
                    echo   [WARN] dotnet command not found after install. Restart your terminal.
                )
            ) else (
                echo   [FAIL] Download of dotnet-install.ps1 failed.
                echo          Please install .NET 10 SDK manually:
                echo          https://dotnet.microsoft.com/download/dotnet/10.0
            )
        )
    ) else (
        echo   Skipping .NET SDK install.
    )

    if !DOTNET_OK!==0 (
        echo.
        echo   ERROR: .NET 10 SDK is required. Please install it and try again.
        pause
        exit /b 1
    )
)

:: --- Project file ---
if not exist "%PROJECTDIR%\SPOSearchProbe.csproj" (
    echo   [FAIL] Project not found at %PROJECTDIR%
    echo          Run this script from the repository root.
    pause
    exit /b 1
)
echo   [OK] Project found: SPOSearchProbe.csproj
echo.

:: =====================================================================
:: Compute version: v1.YY.MDD.Build
:: =====================================================================
set "DATETMP=%TEMP%\spobuild_date.tmp"
powershell -NoProfile -Command "$d=Get-Date; $d.ToString('yy'); $d.Month; $d.ToString('dd'); $d.ToString('yyyyMMdd')" > "%DATETMP%"
set /a DIDX=0
for /f %%A in (%DATETMP%) do (
    set /a DIDX+=1
    if !DIDX!==1 set "YY=%%A"
    if !DIDX!==2 set "M=%%A"
    if !DIDX!==3 set "DD=%%A"
    if !DIDX!==4 set "DATEKEY=%%A"
)
del "%DATETMP%" 2>nul

set "BUILDNUM=1"
if exist "%COUNTERFILE%" (
    set /p STORED_DATE=<"%COUNTERFILE%"
    if "!STORED_DATE!"=="!DATEKEY!" (
        for /f "skip=1" %%N in (%COUNTERFILE%) do set "PREV=%%N"
        set /a BUILDNUM=!PREV!+1
    )
)
>"%COUNTERFILE%" (
    echo !DATEKEY!
    echo !BUILDNUM!
)
set "VERSION=1.%YY%.%M%%DD%.%BUILDNUM%"
echo Version: v%VERSION%

:: =====================================================================
:: Stop running instances
:: =====================================================================
echo.
echo Checking for running SPOSearchProbe instances...
tasklist /FI "IMAGENAME eq SPOSearchProbe.exe" 2>nul | find /I "SPOSearchProbe.exe" >nul
if %ERRORLEVEL%==0 (
    echo Stopping running SPOSearchProbe instances...
    for /f "tokens=2" %%P in ('tasklist /FI "IMAGENAME eq SPOSearchProbe.exe" /NH 2^>nul ^| findstr /I "SPOSearchProbe"') do (
        echo   Stopping PID %%P
        taskkill /PID %%P /F >nul 2>&1
    )
    timeout /t 2 /nobreak >nul
) else (
    echo No running instances found.
)

:: =====================================================================
:: Backup config + tokens
:: =====================================================================
set "TESTDIR=%SCRIPTDIR%..\test"
if not exist "%TESTDIR%" mkdir "%TESTDIR%"
echo.
echo Backing up config and tokens to %TESTDIR% ...
set "CFG_BACKED=0"
for %%R in (win-x64 win-arm64) do (
    set "PUBSRC=%PROJECTDIR%\bin\publish\%%R"
    if exist "!PUBSRC!" (
        if !CFG_BACKED!==0 (
            if exist "!PUBSRC!\search-config.json" (
                copy /y "!PUBSRC!\search-config.json" "%TESTDIR%\" >nul
                echo   Backed up config: !PUBSRC!\search-config.json
                set "CFG_BACKED=1"
            )
        )
        for %%F in ("!PUBSRC!\*.dat") do (
            copy /y "%%F" "%TESTDIR%\" >nul
            echo   Backed up token: %%~nxF
        )
    )
)

:: =====================================================================
:: Clean + Build
:: =====================================================================
echo.
echo Cleaning build directories...
if exist "%PROJECTDIR%\bin" (rd /s /q "%PROJECTDIR%\bin" & echo   Removed bin/)
if exist "%PROJECTDIR%\obj" (rd /s /q "%PROJECTDIR%\obj" & echo   Removed obj/)

if not exist "%DISTDIR%" mkdir "%DISTDIR%"

for %%R in (win-x64 win-arm64) do (
    echo.
    echo === Publishing v!VERSION! for %%R ===

    dotnet publish "%PROJECTDIR%" -c Release -r %%R --self-contained true ^
        -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true ^
        -p:EnableCompressionInSingleFile=true -p:DebugType=none ^
        -p:Version=!VERSION! -p:InformationalVersion=v!VERSION! ^
        -o "%PROJECTDIR%\bin\publish\%%R"

    if !ERRORLEVEL! neq 0 (
        echo ERROR: Publish failed for %%R
        pause
        exit /b 1
    )

    :: Create distribution zip
    set "ZIPPATH=%DISTDIR%\SPOSearchProbe-v!VERSION!-%%R.zip"
    set "STAGING=%SCRIPTDIR%staging-%%R"

    if exist "!ZIPPATH!" del /f "!ZIPPATH!"
    if exist "!STAGING!" rd /s /q "!STAGING!"
    mkdir "!STAGING!"

    copy /y "%PROJECTDIR%\bin\publish\%%R\SPOSearchProbe.exe" "!STAGING!\" >nul
    copy /y "%PROJECTDIR%\search-config.json" "!STAGING!\" >nul

    powershell -NoProfile -Command "Compress-Archive -Path '!STAGING!\*' -DestinationPath '!ZIPPATH!' -Force"
    rd /s /q "!STAGING!"

    echo Package created: !ZIPPATH!
)

echo.
echo === Build complete - v%VERSION% ===

:: =====================================================================
:: Restore config + tokens
:: =====================================================================
echo.
echo Restoring config and tokens from %TESTDIR% ...
for %%R in (win-x64 win-arm64) do (
    set "PUBDST=%PROJECTDIR%\bin\publish\%%R"
    if exist "%TESTDIR%\search-config.json" (
        copy /y "%TESTDIR%\search-config.json" "!PUBDST!\" >nul
        echo   Restored search-config.json -^> %%R
    )
    for %%F in ("%TESTDIR%\*.dat") do (
        copy /y "%%F" "!PUBDST!\" >nul
        echo   Restored %%~nxF -^> %%R
    )
)

echo.
echo Distribution packages in: %DISTDIR%
echo Contents: SPOSearchProbe.exe + search-config.json
echo.

:: --- Launch arm64 build in admin mode ---
if exist "%PROJECTDIR%\bin\publish\win-arm64\SPOSearchProbe.exe" (
    echo Launching SPOSearchProbe arm64 in admin mode...
    start "" "%PROJECTDIR%\bin\publish\win-arm64\SPOSearchProbe.exe" -admin
)
pause
