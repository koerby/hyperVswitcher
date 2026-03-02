; HyperTool.Guest Inno Setup script
; Build with:
; ISCC.exe /DMyAppVersion=2.1.0 /DMySourceDir="...\dist\HyperTool.Guest" /DMyOutputDir="...\dist\installer-guest" installer\HyperTool.Guest.iss

#ifndef MyAppVersion
  #define MyAppVersion "2.1.0"
#endif

#ifndef MySourceDir
  #define MySourceDir "..\\dist\\HyperTool.Guest"
#endif

#ifndef MyOutputDir
  #define MyOutputDir "..\\dist\\installer-guest"
#endif

[Setup]
AppId={{4B7BB8BE-2B17-4B63-8EA2-67B429B7AB33}
AppName=HyperTool Guest
AppVersion={#MyAppVersion}
AppPublisher=koerby
DefaultDirName={autopf}\HyperTool Guest
DefaultGroupName=HyperTool Guest
OutputDir={#MyOutputDir}
OutputBaseFilename=HyperTool-Guest-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
UninstallDisplayIcon={app}\HyperTool.Guest.exe

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
english.DesktopIconTask=Create desktop shortcut
english.AdditionalTasks=Additional tasks:
english.UninstallShortcut=Uninstall HyperTool Guest
english.RunAfterInstall=Launch HyperTool Guest
english.UsbipInstallTask=Download and install usbip-win2 (optional, requires internet connection)

german.DesktopIconTask=Desktop-Verknüpfung erstellen
german.AdditionalTasks=Zusätzliche Aufgaben:
german.UninstallShortcut=HyperTool Guest deinstallieren
german.RunAfterInstall=HyperTool Guest starten
german.UsbipInstallTask=usbip-win2 herunterladen und installieren (optional, Internetverbindung erforderlich)

[Tasks]
Name: "desktopicon"; Description: "{cm:DesktopIconTask}"; GroupDescription: "{cm:AdditionalTasks}"; Flags: unchecked
Name: "installusbip"; Description: "{cm:UsbipInstallTask}"; GroupDescription: "{cm:AdditionalTasks}"; Flags: unchecked; Check: not IsUsbipClientInstalled

[Files]
Source: "{#MySourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion createallsubdirs

[Icons]
Name: "{group}\HyperTool Guest"; Filename: "{app}\HyperTool.Guest.exe"
Name: "{group}\{cm:UninstallShortcut}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\HyperTool Guest"; Filename: "{app}\HyperTool.Guest.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\HyperTool.Guest.exe"; Description: "{cm:RunAfterInstall}"; Flags: nowait postinstall skipifsilent

[Code]
function IsUsbipClientInstalled: Boolean;
begin
  Result :=
    FileExists(ExpandConstant('{pf64}\USBip\usbip.exe')) or
    FileExists(ExpandConstant('{pf}\USBip\usbip.exe')) or
    FileExists(ExpandConstant('{localappdata}\Programs\USBip\usbip.exe'));
end;

procedure TryInstallUsbipWin2;
var
  ResultCode: Integer;
  SetupPath: string;
  PowerShellArgs: string;
begin
  if not WizardIsTaskSelected('installusbip') then
    exit;

  if IsUsbipClientInstalled() then
    exit;

  SetupPath := ExpandConstant('{tmp}\usbip-win2-setup.exe');
  if FileExists(SetupPath) then
  begin
    DeleteFile(SetupPath);
  end;

  PowerShellArgs :=
    '-NoProfile -ExecutionPolicy Bypass -Command ' +
    '"$ProgressPreference=''SilentlyContinue''; ' +
    '$headers=@{ ''User-Agent''=''HyperTool-Guest-Installer'' }; ' +
    '$release=Invoke-RestMethod -Uri ''https://api.github.com/repos/vadimgrn/usbip-win2/releases/latest'' -Headers $headers; ' +
    '$asset=$null; ' +
    'foreach($a in $release.assets){ if($a.name -match ''USBip-.*x64.*Release\.exe$''){ $asset=$a; break } }; ' +
    'if(-not $asset){ foreach($a in $release.assets){ if($a.name -match ''USBip-.*Release\.exe$''){ $asset=$a; break } } }; ' +
    'if($asset){ Invoke-WebRequest -Uri $asset.browser_download_url -Headers $headers -OutFile ''' + SetupPath + ''' }"';

  Exec('powershell.exe', PowerShellArgs, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if FileExists(SetupPath) then
  begin
    Exec(SetupPath, '/SP- /VERYSILENT /SUPPRESSMSGBOXES /NOCANCEL /NORESTART /COMPONENTS="main,client"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    TryInstallUsbipWin2;
  end;
end;
