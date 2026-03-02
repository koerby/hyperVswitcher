@echo off
setlocal enabledelayedexpansion

set "ROOT=%~dp0"
cd /d "%ROOT%"

set "VERSION=2.1.0"
set "NO_PAUSE=false"
set "NO_VERSION_PROMPT=false"
set "SKIP_GUEST_BUILD=false"
set "VERSION_ARG="
set "EXPECT_VERSION_VALUE=false"
set "VERSION_PROMPT=Bitte Version fuer den Guest Installer eingeben (Default 2.1.0): "

for %%A in (%*) do (
    set "ARG=%%~A"

    if /I "!EXPECT_VERSION_VALUE!"=="true" (
        set "VERSION_ARG=version=%%~A"
        set "EXPECT_VERSION_VALUE=false"
    )

    if /I "%%~A"=="no-pause" set "NO_PAUSE=true"
    if /I "%%~A"=="no-version-prompt" set "NO_VERSION_PROMPT=true"
    if /I "%%~A"=="skip-guest-build" set "SKIP_GUEST_BUILD=true"
    if /I "!ARG:~0,8!"=="version=" set "VERSION_ARG=%%~A"
    if /I "!ARG!"=="version" set "EXPECT_VERSION_VALUE=true"
)

if defined VERSION_ARG (
    for /f "tokens=1,* delims==" %%K in ("%VERSION_ARG%") do set "VERSION=%%L"
)

if not defined VERSION_ARG if /I "%NO_VERSION_PROMPT%"=="false" (
    set /p "VERSION=!VERSION_PROMPT!"
)

if not defined VERSION set "VERSION=2.1.0"

if /I "%SKIP_GUEST_BUILD%"=="false" (
    echo Aktualisiere Guest DIST-Build fuer Version %VERSION%...
    call "%ROOT%build-guest.bat" "version=%VERSION%" no-version-prompt no-pause no-installer
    if errorlevel 1 (
        echo Guest Build vor Installer-Erstellung fehlgeschlagen.
        if /I "%NO_PAUSE%"=="false" pause
        exit /b 1
    )
)

if not exist "%ROOT%dist\HyperTool.Guest\HyperTool.Guest.exe" (
    echo Guest DIST-Build nicht gefunden. Fuehre zuerst build-guest.bat aus.
    if /I "%NO_PAUSE%"=="false" pause
    exit /b 1
)

set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist "%ISCC%" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"

if not exist "%ISCC%" (
    echo Inno Setup wurde nicht gefunden.
    echo Bitte Inno Setup 6 installieren: https://jrsoftware.org/isinfo.php
    if /I "%NO_PAUSE%"=="false" pause
    exit /b 1
)

set "OUT_DIR=%ROOT%dist\installer-guest"
if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"
if exist "%OUT_DIR%\prerequisites" rmdir /s /q "%OUT_DIR%\prerequisites"

echo Erzeuge Guest Installer fuer Version %VERSION%...
"%ISCC%" /DMyAppVersion=%VERSION% /DMySourceDir="%ROOT%dist\HyperTool.Guest" /DMyOutputDir="%OUT_DIR%" "%ROOT%installer\HyperTool.Guest.iss"

if errorlevel 1 (
    echo Guest Installer-Erstellung fehlgeschlagen.
    if /I "%NO_PAUSE%"=="false" pause
    exit /b 1
)

echo.
echo SUCCESS: Guest Installer erstellt in %OUT_DIR%
dir /b "%OUT_DIR%"

if /I "%NO_PAUSE%"=="false" pause
exit /b 0
