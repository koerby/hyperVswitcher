@echo off
setlocal enabledelayedexpansion

set "ROOT=%~dp0"
cd /d "%ROOT%"

set "NO_PAUSE=false"
set "DEFAULT_VERSION=2.1.0"
set "BUILD_HOST=false"
set "BUILD_GUEST=false"
set "HOST_INSTALLER=false"
set "GUEST_INSTALLER=false"
set "HOST_VERSION=%DEFAULT_VERSION%"
set "GUEST_VERSION=%DEFAULT_VERSION%"
set "HOST_SELECTED=false"
set "GUEST_SELECTED=false"
set "HOST_INSTALLER_SELECTED=false"
set "GUEST_INSTALLER_SELECTED=false"
set "HOST_VERSION_ARG="
set "GUEST_VERSION_ARG="
set "GLOBAL_VERSION_ARG="

for %%A in (%*) do (
    set "ARG=%%~A"

    if /I "%%~A"=="no-pause" set "NO_PAUSE=true"

    if /I "%%~A"=="host" (
        set "BUILD_HOST=true"
        set "HOST_SELECTED=true"
    )
    if /I "%%~A"=="no-host" (
        set "BUILD_HOST=false"
        set "HOST_SELECTED=true"
    )

    if /I "%%~A"=="guest" (
        set "BUILD_GUEST=true"
        set "GUEST_SELECTED=true"
    )
    if /I "%%~A"=="no-guest" (
        set "BUILD_GUEST=false"
        set "GUEST_SELECTED=true"
    )

    if /I "%%~A"=="host-installer" (
        set "HOST_INSTALLER=true"
        set "HOST_INSTALLER_SELECTED=true"
    )
    if /I "%%~A"=="no-host-installer" (
        set "HOST_INSTALLER=false"
        set "HOST_INSTALLER_SELECTED=true"
    )

    if /I "%%~A"=="guest-installer" (
        set "GUEST_INSTALLER=true"
        set "GUEST_INSTALLER_SELECTED=true"
    )
    if /I "%%~A"=="no-guest-installer" (
        set "GUEST_INSTALLER=false"
        set "GUEST_INSTALLER_SELECTED=true"
    )

    if /I "!ARG:~0,8!"=="version=" set "GLOBAL_VERSION_ARG=%%~A"
    if /I "!ARG:~0,13!"=="host-version=" set "HOST_VERSION_ARG=%%~A"
    if /I "!ARG:~0,14!"=="guest-version=" set "GUEST_VERSION_ARG=%%~A"
)

if defined GLOBAL_VERSION_ARG (
    for /f "tokens=1,* delims==" %%K in ("%GLOBAL_VERSION_ARG%") do (
        set "HOST_VERSION=%%L"
        set "GUEST_VERSION=%%L"
    )
)

if defined HOST_VERSION_ARG (
    for /f "tokens=1,* delims==" %%K in ("%HOST_VERSION_ARG%") do set "HOST_VERSION=%%L"
)

if defined GUEST_VERSION_ARG (
    for /f "tokens=1,* delims==" %%K in ("%GUEST_VERSION_ARG%") do set "GUEST_VERSION=%%L"
)

echo ==========================================
echo HyperTool Build-All Script
echo ROOT: %ROOT%
echo ==========================================
echo.

if /I "%HOST_SELECTED%"=="false" (
    choice /C JN /N /M "Host (WinUI) erzeugen? [J/N]: "
    if errorlevel 2 (
        set "BUILD_HOST=false"
    ) else (
        set "BUILD_HOST=true"
    )
)

if /I "%BUILD_HOST%"=="true" (
    echo.
    if not defined HOST_VERSION_ARG if not defined GLOBAL_VERSION_ARG (
        set /p "HOST_VERSION=Version fuer Host/WinUI (Default %DEFAULT_VERSION%): "
        if not defined HOST_VERSION set "HOST_VERSION=%DEFAULT_VERSION%"
    )

    if /I "%HOST_INSTALLER_SELECTED%"=="false" (
        choice /C JN /N /M "Host-Installer auch erzeugen? [J/N]: "
        if errorlevel 2 (
            set "HOST_INSTALLER=false"
        ) else (
            set "HOST_INSTALLER=true"
        )
    )
)

if /I "%GUEST_SELECTED%"=="false" (
    choice /C JN /N /M "Guest erzeugen? [J/N]: "
    if errorlevel 2 (
        set "BUILD_GUEST=false"
    ) else (
        set "BUILD_GUEST=true"
    )
)

if /I "%BUILD_GUEST%"=="true" (
    echo.
    if not defined GUEST_VERSION_ARG if not defined GLOBAL_VERSION_ARG (
        set /p "GUEST_VERSION=Version fuer Guest (Default %DEFAULT_VERSION%): "
        if not defined GUEST_VERSION set "GUEST_VERSION=%DEFAULT_VERSION%"
    )

    if /I "%GUEST_INSTALLER_SELECTED%"=="false" (
        choice /C JN /N /M "Guest-Installer auch erzeugen? [J/N]: "
        if errorlevel 2 (
            set "GUEST_INSTALLER=false"
        ) else (
            set "GUEST_INSTALLER=true"
        )
    )
)

if /I "%BUILD_HOST%"=="false" if /I "%BUILD_GUEST%"=="false" (
    echo.
    echo Kein Build ausgewaehlt. Abbruch.
    if /I "%NO_PAUSE%"=="false" pause
    exit /b 1
)

echo.
echo Auswahl:
echo - BUILD_HOST=%BUILD_HOST% VERSION=%HOST_VERSION% INSTALLER=%HOST_INSTALLER%
echo - BUILD_GUEST=%BUILD_GUEST% VERSION=%GUEST_VERSION% INSTALLER=%GUEST_INSTALLER%
echo.

if /I "%BUILD_HOST%"=="true" (
    if /I "%HOST_INSTALLER%"=="true" (
        echo [1/2] Starte Host-Build inkl. Installer...
        call "%ROOT%build-winui.bat" "version=%HOST_VERSION%" installer no-pause
    ) else (
        echo [1/2] Starte Host-Build ohne Installer...
        call "%ROOT%build-winui.bat" "version=%HOST_VERSION%" no-installer no-pause
    )
    if errorlevel 1 goto :fail
)

if /I "%BUILD_GUEST%"=="true" (
    if /I "%GUEST_INSTALLER%"=="true" (
        echo [2/2] Starte Guest-Build inkl. Installer...
        call "%ROOT%build-guest.bat" "version=%GUEST_VERSION%" installer no-pause
    ) else (
        echo [2/2] Starte Guest-Build ohne Installer...
        call "%ROOT%build-guest.bat" "version=%GUEST_VERSION%" no-installer no-pause
    )
    if errorlevel 1 goto :fail
)

echo.
echo SUCCESS: Build-All abgeschlossen.
echo Host-Output:  %ROOT%dist\HyperTool.WinUI
echo Host-Installer: %ROOT%dist\installer-winui
echo Guest-Output: %ROOT%dist\HyperTool.Guest
echo Guest-Installer: %ROOT%dist\installer-guest
if /I "%NO_PAUSE%"=="false" pause
exit /b 0

:fail
echo.
echo FEHLER: Build-All fehlgeschlagen.
if /I "%NO_PAUSE%"=="false" pause
exit /b 1
