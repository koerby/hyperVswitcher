@echo off
setlocal enabledelayedexpansion

set "ROOT=%~dp0"
cd /d "%ROOT%"

set "CONFIG=Release"
set "RUNTIME=win-x64"
set "SELF_CONTAINED=true"
set "NO_PAUSE=false"
set "CLEAN_BUILD=false"
set "CREATE_INSTALLER=false"
set "INSTALLER_DECIDED=false"
set "VERSION=2.1.0"
set "VERSION_ARG="
set "EXPECT_VERSION_VALUE=false"
set "VERSION_PROMPT=Bitte Version fuer HyperTool.Guest eingeben (Default 2.1.0): "

for %%A in (%*) do (
    set "ARG=%%~A"

    if /I "!EXPECT_VERSION_VALUE!"=="true" (
        set "VERSION_ARG=version=%%~A"
        set "EXPECT_VERSION_VALUE=false"
    )

    if /I "%%~A"=="self-contained" set "SELF_CONTAINED=true"
    if /I "%%~A"=="framework-dependent" set "SELF_CONTAINED=false"
    if /I "%%~A"=="no-pause" set "NO_PAUSE=true"
    if /I "%%~A"=="clean" set "CLEAN_BUILD=true"
    if /I "%%~A"=="installer" (
        set "CREATE_INSTALLER=true"
        set "INSTALLER_DECIDED=true"
    )
    if /I "%%~A"=="no-installer" (
        set "CREATE_INSTALLER=false"
        set "INSTALLER_DECIDED=true"
    )
    if /I "!ARG:~0,8!"=="version=" set "VERSION_ARG=%%~A"
    if /I "!ARG!"=="version" set "EXPECT_VERSION_VALUE=true"
)

if defined VERSION_ARG (
    for /f "tokens=1,* delims==" %%K in ("%VERSION_ARG%") do set "VERSION=%%L"
)

if not defined VERSION_ARG if /I "%NO_PAUSE%"=="false" (
    set /p "VERSION=!VERSION_PROMPT!"
)

if not defined VERSION set "VERSION=2.1.0"

echo ==========================================
echo HyperTool.Guest Build Script
echo ROOT: %ROOT%
echo CONFIG: %CONFIG%
echo RUNTIME: %RUNTIME%
echo SELF_CONTAINED: %SELF_CONTAINED%
echo VERSION: %VERSION%
echo CLEAN_BUILD: %CLEAN_BUILD%
echo ==========================================
echo.

if /I "%CLEAN_BUILD%"=="true" (
    echo [0/4] Clean Build Mode aktiv...
    dotnet clean src\HyperTool.Guest\HyperTool.Guest.csproj -c %CONFIG%
    if errorlevel 1 goto :fail
    echo.
)

echo [1/4] Restore Guest project...
dotnet restore src\HyperTool.Guest\HyperTool.Guest.csproj
if errorlevel 1 goto :fail

echo [2/4] Build Guest project...
dotnet build src\HyperTool.Guest\HyperTool.Guest.csproj -c %CONFIG% --no-restore /p:Version=%VERSION% /p:FileVersion=%VERSION% /p:AssemblyVersion=%VERSION% /p:InformationalVersion=%VERSION%
if errorlevel 1 goto :fail

set "DIST_DIR=%ROOT%dist\HyperTool.Guest"
if exist "%DIST_DIR%" rmdir /s /q "%DIST_DIR%"
mkdir "%DIST_DIR%"

echo [3/4] Publish Guest project...
dotnet publish src\HyperTool.Guest\HyperTool.Guest.csproj -c %CONFIG% -r %RUNTIME% --self-contained %SELF_CONTAINED% -o "%DIST_DIR%" /p:Version=%VERSION% /p:FileVersion=%VERSION% /p:AssemblyVersion=%VERSION% /p:InformationalVersion=%VERSION%
if errorlevel 1 goto :fail

if /I "%INSTALLER_DECIDED%"=="false" (
    if /I "%NO_PAUSE%"=="false" (
        echo.
        choice /C JN /N /M "Guest-Installer auch erstellen? [J/N]: "
        if errorlevel 2 (
            set "CREATE_INSTALLER=false"
        ) else (
            set "CREATE_INSTALLER=true"
        )
    )
)

if /I "%CREATE_INSTALLER%"=="true" (
    echo [4/5] Erzeuge Guest Installer...
    call "%ROOT%build-installer-guest.bat" "version=%VERSION%" no-version-prompt no-pause skip-guest-build
    if errorlevel 1 goto :fail
    echo [5/5] Build Guest abgeschlossen.
) else (
    echo [4/4] Build Guest abgeschlossen.
)

echo.
echo SUCCESS: HyperTool.Guest Build und Publish abgeschlossen.
echo Ausgabe liegt in:
echo %DIST_DIR%
echo.
dir /b "%DIST_DIR%"

if /I "%NO_PAUSE%"=="false" pause
goto :success

:fail
echo.
echo FEHLER: Guest Build/Publish fehlgeschlagen.
if /I "%NO_PAUSE%"=="false" pause
exit /b 1

:success
exit /b 0
