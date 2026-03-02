; HyperTool Inno Setup script
; Build with:
; ISCC.exe /DMyAppVersion=1.2.0 /DMySourceDir="...\dist\HyperTool" /DMyOutputDir="...\dist\installer" installer\HyperTool.iss

#ifndef MyAppVersion
  #define MyAppVersion "1.2.0"
#endif

#ifndef MySourceDir
  #define MySourceDir "..\\dist\\HyperTool"
#endif

#ifndef MyOutputDir
  #define MyOutputDir "..\\dist\\installer"
#endif

[Setup]
AppId={{E3AF03D2-9A6A-4E17-9E42-1B95A4D0FA93}
AppName=HyperTool
AppVersion={#MyAppVersion}
AppPublisher=koerby
DefaultDirName={autopf}\HyperTool
DefaultGroupName=HyperTool
OutputDir={#MyOutputDir}
OutputBaseFilename=HyperTool-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
UninstallDisplayIcon={app}\HyperTool.exe

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
english.DesktopIconTask=Create desktop shortcut
english.AdditionalTasks=Additional tasks:
english.UninstallShortcut=Uninstall HyperTool
english.RunAfterInstall=Launch HyperTool
english.UsbipdInstallTask=Download and install usbipd-win (optional, requires internet connection)

german.DesktopIconTask=Desktop-Verknüpfung erstellen
german.AdditionalTasks=Zusätzliche Aufgaben:
german.UninstallShortcut=HyperTool deinstallieren
german.RunAfterInstall=HyperTool starten
german.UsbipdInstallTask=usbipd-win herunterladen und installieren (optional, Internetverbindung erforderlich)

[Tasks]
Name: "desktopicon"; Description: "{cm:DesktopIconTask}"; GroupDescription: "{cm:AdditionalTasks}"; Flags: unchecked
Name: "installusbipd"; Description: "{cm:UsbipdInstallTask}"; GroupDescription: "{cm:AdditionalTasks}"; Flags: unchecked; Check: not IsUsbipdInstalled

[Files]
Source: "{#MySourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion createallsubdirs

[Icons]
Name: "{group}\HyperTool"; Filename: "{app}\HyperTool.exe"
Name: "{group}\{cm:UninstallShortcut}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\HyperTool"; Filename: "{app}\HyperTool.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\HyperTool.exe"; Description: "{cm:RunAfterInstall}"; Flags: nowait postinstall skipifsilent

[Code]
function IsUsbipdInstalled(): Boolean;
var
  InstallPath: string;
begin
  Result :=
    RegQueryStringValue(HKLM64, 'SOFTWARE\\usbipd-win', 'APPLICATIONFOLDER', InstallPath) and
    FileExists(AddBackslash(InstallPath) + 'usbipd.exe');

  if not Result then
  begin
    Result :=
      RegQueryStringValue(HKLM32, 'SOFTWARE\\usbipd-win', 'APPLICATIONFOLDER', InstallPath) and
      FileExists(AddBackslash(InstallPath) + 'usbipd.exe');
  end;

  if not Result then
  begin
    Result := FileExists(ExpandConstant('{pf}\\usbipd-win\\usbipd.exe'));
  end;
end;

procedure TryInstallUsbipdFromInternet;
var
  ResultCode: Integer;
  PsExe: string;
  DownloadMsiCommand: string;
  DownloadExeCommand: string;
  MsiPath: string;
  ExePath: string;
begin
  if not WizardIsTaskSelected('installusbipd') then
    Exit;

  if IsUsbipdInstalled() then
    Exit;

  PsExe := ExpandConstant('{sys}\\WindowsPowerShell\\v1.0\\powershell.exe');
  if not FileExists(PsExe) then
    Exit;

  MsiPath := ExpandConstant('{tmp}\\usbipd-win-setup.msi');
  ExePath := ExpandConstant('{tmp}\\usbipd-win-setup.exe');

  if FileExists(MsiPath) then
    DeleteFile(MsiPath);
  if FileExists(ExePath) then
    DeleteFile(ExePath);

  DownloadMsiCommand :=
    '-NoProfile -ExecutionPolicy Bypass -Command "' +
    '$ProgressPreference=''SilentlyContinue''; ' +
    '$headers=@{ ''User-Agent''=''HyperTool-Installer'' }; ' +
    '$release=Invoke-RestMethod -Uri ''https://api.github.com/repos/dorssel/usbipd-win/releases/latest'' -Headers $headers; ' +
    '$asset=$null; foreach($a in $release.assets){ if($a.name -match ''\\.msi$''){ $asset=$a; break } }; ' +
    'if($asset){ Invoke-WebRequest -Uri $asset.browser_download_url -Headers $headers -OutFile ''' + MsiPath + ''' }"';

  Exec(PsExe, DownloadMsiCommand, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if FileExists(MsiPath) then
  begin
    Exec('msiexec.exe', '/i "' + MsiPath + '" /qn /norestart', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    Exit;
  end;

  DownloadExeCommand :=
    '-NoProfile -ExecutionPolicy Bypass -Command "' +
    '$ProgressPreference=''SilentlyContinue''; ' +
    '$headers=@{ ''User-Agent''=''HyperTool-Installer'' }; ' +
    '$release=Invoke-RestMethod -Uri ''https://api.github.com/repos/dorssel/usbipd-win/releases/latest'' -Headers $headers; ' +
    '$asset=$null; foreach($a in $release.assets){ if($a.name -match ''\\.exe$''){ $asset=$a; break } }; ' +
    'if($asset){ Invoke-WebRequest -Uri $asset.browser_download_url -Headers $headers -OutFile ''' + ExePath + ''' }"';

  Exec(PsExe, DownloadExeCommand, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if FileExists(ExePath) then
  begin
    Exec(ExePath, '/SP- /VERYSILENT /SUPPRESSMSGBOXES /NOCANCEL /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    TryInstallUsbipdFromInternet;
  end;
end;
