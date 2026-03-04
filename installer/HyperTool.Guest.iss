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
AppPublisher=github.com/koerby
AppPublisherURL=https://github.com/koerby
AppSupportURL=https://github.com/koerby/HyperTool/issues
AppUpdatesURL=https://github.com/koerby/HyperTool/releases
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany=github.com/koerby
VersionInfoDescription=HyperTool Guest Setup
VersionInfoProductName=HyperTool Guest
VersionInfoProductVersion={#MyAppVersion}
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
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
english.DesktopIconTask=Create desktop shortcut
english.AdditionalTasks=Additional tasks:
english.UninstallShortcut=Uninstall HyperTool Guest
english.RunAfterInstall=Launch HyperTool Guest
english.UsbipInstallTask=Download and install usbip-win2 (optional, needed for USB-Share in HyperTool Guest)
english.WinFspInstallTask=Download and install WinFsp (optional, needed for Shared Folder network drives in HyperTool Guest)
english.WinFspInstallFailed=WinFsp could not be installed automatically. Please install it manually from https://github.com/winfsp/winfsp/releases and then restart HyperTool Guest.
english.UninstallDepsTitle=Optional dependency cleanup
english.UninstallDepsText=Select optional dependencies that should also be removed during HyperTool Guest uninstall.
english.UninstallWinFspCheckbox=Also uninstall WinFsp (if installed)
english.UninstallUsbipCheckbox=Also uninstall usbip-win2 (if installed)

german.DesktopIconTask=Desktop-Verknüpfung erstellen
german.AdditionalTasks=Zusätzliche Aufgaben:
german.UninstallShortcut=HyperTool Guest deinstallieren
german.RunAfterInstall=HyperTool Guest starten
german.UsbipInstallTask=usbip-win2 herunterladen und installieren (optional, benötigt für USB-Share in HyperTool Guest)
german.WinFspInstallTask=WinFsp herunterladen und installieren (optional, benötigt für Shared-Folder Netzlaufwerke in HyperTool Guest)
german.WinFspInstallFailed=WinFsp konnte nicht automatisch installiert werden. Bitte manuell von https://github.com/winfsp/winfsp/releases installieren und HyperTool Guest danach neu starten.
german.UninstallDepsTitle=Optionale Abhängigkeits-Bereinigung
german.UninstallDepsText=Wähle optionale Abhängigkeiten aus, die bei der HyperTool Guest-Deinstallation ebenfalls entfernt werden sollen.
german.UninstallWinFspCheckbox=WinFsp ebenfalls deinstallieren (falls installiert)
german.UninstallUsbipCheckbox=usbip-win2 ebenfalls deinstallieren (falls installiert)

[Tasks]
Name: "desktopicon"; Description: "{cm:DesktopIconTask}"; GroupDescription: "{cm:AdditionalTasks}"
Name: "installusbip"; Description: "{cm:UsbipInstallTask}"; GroupDescription: "{cm:AdditionalTasks}"; Check: not IsUsbipClientInstalled; Flags: checkedonce
Name: "installwinfsp"; Description: "{cm:WinFspInstallTask}"; GroupDescription: "{cm:AdditionalTasks}"; Check: not IsWinFspInstalled; Flags: checkedonce

[Files]
Source: "{#MySourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion createallsubdirs

[Registry]
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\6c4eb1be-40e8-4c8b-a4d6-5b0f67d7e40f"; ValueType: string; ValueName: "ElementName"; ValueData: "HyperTool Hyper-V Socket USB Tunnel"; Flags: uninsdeletekey

[Icons]
Name: "{group}\HyperTool Guest"; Filename: "{app}\HyperTool.Guest.exe"; IconFilename: "{app}\HyperTool.Guest.exe"; IconIndex: 0; AppUserModelID: "HyperTool.Guest"
Name: "{group}\{cm:UninstallShortcut}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\HyperTool Guest"; Filename: "{app}\HyperTool.Guest.exe"; IconFilename: "{app}\HyperTool.Guest.exe"; IconIndex: 0; AppUserModelID: "HyperTool.Guest"; Tasks: desktopicon

[Run]
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall add rule name=""HyperTool Guest USB Discovery (UDP-Out)"" dir=out action=allow protocol=UDP remoteport=32491 profile=private,domain program=""{app}\HyperTool.Guest.exe"""; Flags: runhidden waituntilterminated
Filename: "{app}\HyperTool.Guest.exe"; Description: "{cm:RunAfterInstall}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/IM HyperTool.Guest.exe /T"; Flags: runhidden waituntilterminated skipifdoesntexist; RunOnceId: "HyperToolGuest-Uninstall-StopApp"
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=""HyperTool Guest USB Discovery (UDP-Out)"""; Flags: runhidden waituntilterminated; RunOnceId: "HyperToolGuest-Uninstall-DeleteFirewall-UDP-Out"
Filename: "{cmd}"; Parameters: "/C reg delete ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\6c4eb1be-40e8-4c8b-a4d6-5b0f67d7e40f"" /f"; Flags: runhidden waituntilterminated; RunOnceId: "HyperToolGuest-Uninstall-DeleteSvc-UsbTunnel"
Filename: "{cmd}"; Parameters: "/C winget.exe uninstall --id WinFsp.WinFsp --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveWinFspOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-WinFsp-WingetId"
Filename: "{cmd}"; Parameters: "/C winget.exe uninstall --name ""WinFsp"" --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveWinFspOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-WinFsp-WingetName"
Filename: "{cmd}"; Parameters: "/C ""%LOCALAPPDATA%\Microsoft\WindowsApps\winget.exe"" uninstall --id WinFsp.WinFsp --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveWinFspOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-WinFsp-WingetId-LocalAppData"
Filename: "{cmd}"; Parameters: "/C ""%LOCALAPPDATA%\Microsoft\WindowsApps\winget.exe"" uninstall --name ""WinFsp"" --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveWinFspOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-WinFsp-WingetName-LocalAppData"
Filename: "{commonpf64}\WinFsp\uninstall.exe"; Parameters: "/quiet /norestart"; Flags: runhidden waituntilterminated skipifdoesntexist; Check: ShouldRemoveWinFspOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-WinFsp-Uninstall64"
Filename: "{commonpf}\WinFsp\uninstall.exe"; Parameters: "/quiet /norestart"; Flags: runhidden waituntilterminated skipifdoesntexist; Check: ShouldRemoveWinFspOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-WinFsp-Uninstall32"
Filename: "{cmd}"; Parameters: "/C winget.exe uninstall --id vadimgrn.usbip-win2 --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveUsbipOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-Usbip-WingetId"
Filename: "{cmd}"; Parameters: "/C winget.exe uninstall --name ""USBip"" --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveUsbipOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-Usbip-WingetName"
Filename: "{cmd}"; Parameters: "/C ""%LOCALAPPDATA%\Microsoft\WindowsApps\winget.exe"" uninstall --id vadimgrn.usbip-win2 --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveUsbipOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-Usbip-WingetId-LocalAppData"
Filename: "{cmd}"; Parameters: "/C ""%LOCALAPPDATA%\Microsoft\WindowsApps\winget.exe"" uninstall --name ""USBip"" --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveUsbipOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-Usbip-WingetName-LocalAppData"
Filename: "{commonpf64}\USBip\unins000.exe"; Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; Flags: runhidden waituntilterminated skipifdoesntexist; Check: ShouldRemoveUsbipOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-Usbip-Unins64"
Filename: "{commonpf}\USBip\unins000.exe"; Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; Flags: runhidden waituntilterminated skipifdoesntexist; Check: ShouldRemoveUsbipOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-Usbip-Unins32"
Filename: "{commonpf64}\usbip-win2\unins000.exe"; Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; Flags: runhidden waituntilterminated skipifdoesntexist; Check: ShouldRemoveUsbipOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-Usbip-Unins64-Alt"
Filename: "{commonpf}\usbip-win2\unins000.exe"; Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; Flags: runhidden waituntilterminated skipifdoesntexist; Check: ShouldRemoveUsbipOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-Usbip-Unins32-Alt"
Filename: "{localappdata}\Programs\USBip\unins000.exe"; Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; Flags: runhidden waituntilterminated skipifdoesntexist; Check: ShouldRemoveUsbipOnUninstall; RunOnceId: "HyperToolGuest-Uninstall-Optional-Usbip-Unins-LocalAppData"

[Code]
var
  RemoveWinFspOnUninstall: Boolean;
  RemoveUsbipOnUninstall: Boolean;
  UninstallOptionsPrompted: Boolean;
  OptionalDepsUninstallExecuted: Boolean;

function HasUninstallDisplayNameMatching(const NamePart: string): Boolean;
var
  SubKeys: TArrayOfString;
  Root: Integer;
  BaseKey: string;
  DisplayName: string;
  I: Integer;
begin
  Result := False;

  for Root := 0 to 2 do
  begin
    BaseKey := 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall';

    if (Root = 0) and RegGetSubkeyNames(HKLM64, BaseKey, SubKeys) then
    begin
      for I := 0 to GetArrayLength(SubKeys) - 1 do
      begin
        if RegQueryStringValue(HKLM64, BaseKey + '\\' + SubKeys[I], 'DisplayName', DisplayName)
           and (Pos(LowerCase(NamePart), LowerCase(DisplayName)) > 0) then
        begin
          Result := True;
          Exit;
        end;
      end;
    end;

    if (Root = 1) and RegGetSubkeyNames(HKLM32, BaseKey, SubKeys) then
    begin
      for I := 0 to GetArrayLength(SubKeys) - 1 do
      begin
        if RegQueryStringValue(HKLM32, BaseKey + '\\' + SubKeys[I], 'DisplayName', DisplayName)
           and (Pos(LowerCase(NamePart), LowerCase(DisplayName)) > 0) then
        begin
          Result := True;
          Exit;
        end;
      end;
    end;

    if (Root = 2) and RegGetSubkeyNames(HKCU, BaseKey, SubKeys) then
    begin
      for I := 0 to GetArrayLength(SubKeys) - 1 do
      begin
        if RegQueryStringValue(HKCU, BaseKey + '\\' + SubKeys[I], 'DisplayName', DisplayName)
           and (Pos(LowerCase(NamePart), LowerCase(DisplayName)) > 0) then
        begin
          Result := True;
          Exit;
        end;
      end;
    end;
  end;
end;

procedure TryUninstallByDisplayName(const NamePart: string);
var
  SubKeys: TArrayOfString;
  BaseKey: string;
  DisplayName: string;
  CommandLine: string;
  Root: Integer;
  I: Integer;
  ResultCode: Integer;
begin
  BaseKey := 'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall';

  for Root := 0 to 2 do
  begin
    if ((Root = 0) and RegGetSubkeyNames(HKLM64, BaseKey, SubKeys))
       or ((Root = 1) and RegGetSubkeyNames(HKLM32, BaseKey, SubKeys))
       or ((Root = 2) and RegGetSubkeyNames(HKCU, BaseKey, SubKeys)) then
    begin
      for I := 0 to GetArrayLength(SubKeys) - 1 do
      begin
        DisplayName := '';
        CommandLine := '';

        if Root = 0 then
        begin
          RegQueryStringValue(HKLM64, BaseKey + '\\' + SubKeys[I], 'DisplayName', DisplayName);
          if not RegQueryStringValue(HKLM64, BaseKey + '\\' + SubKeys[I], 'QuietUninstallString', CommandLine) then
            RegQueryStringValue(HKLM64, BaseKey + '\\' + SubKeys[I], 'UninstallString', CommandLine);
        end
        else if Root = 1 then
        begin
          RegQueryStringValue(HKLM32, BaseKey + '\\' + SubKeys[I], 'DisplayName', DisplayName);
          if not RegQueryStringValue(HKLM32, BaseKey + '\\' + SubKeys[I], 'QuietUninstallString', CommandLine) then
            RegQueryStringValue(HKLM32, BaseKey + '\\' + SubKeys[I], 'UninstallString', CommandLine);
        end
        else
        begin
          RegQueryStringValue(HKCU, BaseKey + '\\' + SubKeys[I], 'DisplayName', DisplayName);
          if not RegQueryStringValue(HKCU, BaseKey + '\\' + SubKeys[I], 'QuietUninstallString', CommandLine) then
            RegQueryStringValue(HKCU, BaseKey + '\\' + SubKeys[I], 'UninstallString', CommandLine);
        end;

        if (Pos(LowerCase(NamePart), LowerCase(DisplayName)) > 0) and (Trim(CommandLine) <> '') then
        begin
          Exec(ExpandConstant('{cmd}'), '/C ' + CommandLine, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
        end;
      end;
    end;
  end;
end;

procedure CloseRunningGuestApp;
var
  ResultCode: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/IM HyperTool.Guest.exe /F /T', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

function IsUsbipClientInstalled: Boolean;
begin
  Result :=
    FileExists(ExpandConstant('{commonpf64}\USBip\usbip.exe')) or
    FileExists(ExpandConstant('{commonpf}\USBip\usbip.exe')) or
    FileExists(ExpandConstant('{localappdata}\Programs\USBip\usbip.exe'));

  if not Result then
  begin
    Result := HasUninstallDisplayNameMatching('usbip-win2') or HasUninstallDisplayNameMatching('usbip');
  end;
end;

function IsWinFspInstalled: Boolean;
var
  InstallDir: string;
begin
  Result :=
    RegQueryStringValue(HKLM64, 'SOFTWARE\\WinFsp', 'InstallDir', InstallDir) and
    FileExists(AddBackslash(InstallDir) + 'bin\\winfsp-x64.dll');

  if not Result then
  begin
    Result :=
      RegQueryStringValue(HKLM32, 'SOFTWARE\\WinFsp', 'InstallDir', InstallDir) and
      FileExists(AddBackslash(InstallDir) + 'bin\\winfsp-x64.dll');
  end;

  if not Result then
  begin
    Result :=
      FileExists(ExpandConstant('{commonpf64}\WinFsp\bin\winfsp-x64.dll')) or
      FileExists(ExpandConstant('{commonpf}\WinFsp\bin\winfsp-x64.dll'));
  end;

  if not Result then
  begin
    Result := HasUninstallDisplayNameMatching('winfsp');
  end;
end;

function ShouldRemoveWinFspOnUninstall(): Boolean;
begin
  Result := RemoveWinFspOnUninstall;
end;

function ShouldRemoveUsbipOnUninstall(): Boolean;
begin
  Result := RemoveUsbipOnUninstall;
end;

procedure PromptUninstallOptions;
begin
  RemoveWinFspOnUninstall := False;
  RemoveUsbipOnUninstall := False;

  if IsWinFspInstalled() then
  begin
    if SuppressibleMsgBox(
      ExpandConstant('{cm:UninstallDepsText}') + #13#10#13#10 + ExpandConstant('{cm:UninstallWinFspCheckbox}'),
      mbConfirmation,
      MB_YESNO,
      IDNO) = IDYES then
    begin
      RemoveWinFspOnUninstall := True;
    end;
  end;

  if IsUsbipClientInstalled() then
  begin
    if SuppressibleMsgBox(
      ExpandConstant('{cm:UninstallDepsText}') + #13#10#13#10 + ExpandConstant('{cm:UninstallUsbipCheckbox}'),
      mbConfirmation,
      MB_YESNO,
      IDNO) = IDYES then
    begin
      RemoveUsbipOnUninstall := True;
    end;
  end;
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

procedure TryInstallWinFsp;
var
  ResultCode: Integer;
  InstallerPath: string;
  PowerShellArgs: string;
begin
  if not WizardIsTaskSelected('installwinfsp') then
    Exit;

  if IsWinFspInstalled() then
    Exit;

  InstallerPath := ExpandConstant('{tmp}\winfsp-setup.msi');
  if FileExists(InstallerPath) then
  begin
    DeleteFile(InstallerPath);
  end;

  PowerShellArgs :=
    '-NoProfile -ExecutionPolicy Bypass -Command ' +
    '"$ProgressPreference=''SilentlyContinue''; ' +
    '$headers=@{ ''User-Agent''=''HyperTool-Guest-Installer'' }; ' +
    '$release=Invoke-RestMethod -Uri ''https://api.github.com/repos/winfsp/winfsp/releases/latest'' -Headers $headers; ' +
    '$asset=$null; ' +
    'foreach($a in $release.assets){ if($a.name -match ''winfsp-.*x64.*\.msi$''){ $asset=$a; break } }; ' +
    'if(-not $asset){ foreach($a in $release.assets){ if($a.name -match ''winfsp-.*\.msi$''){ $asset=$a; break } } }; ' +
    'if($asset){ Invoke-WebRequest -Uri $asset.browser_download_url -Headers $headers -OutFile ''' + InstallerPath + ''' }"';

  Exec('powershell.exe', PowerShellArgs, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  if FileExists(InstallerPath) then
  begin
    Exec('msiexec.exe', '/i "' + InstallerPath + '" /qn /norestart', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;

  if not IsWinFspInstalled() then
  begin
    MsgBox(ExpandConstant('{cm:WinFspInstallFailed}'), mbInformation, MB_OK);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
  begin
    CloseRunningGuestApp;
  end;

  if CurStep = ssPostInstall then
  begin
    TryInstallUsbipWin2;
    TryInstallWinFsp;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if (CurUninstallStep = usAppMutexCheck) and (not UninstallOptionsPrompted) then
  begin
    CloseRunningGuestApp;
    UninstallOptionsPrompted := True;
    PromptUninstallOptions;
  end;

  if (CurUninstallStep = usUninstall) and (not OptionalDepsUninstallExecuted) then
  begin
    OptionalDepsUninstallExecuted := True;

    if ShouldRemoveWinFspOnUninstall() then
    begin
      TryUninstallByDisplayName('winfsp');
    end;

    if ShouldRemoveUsbipOnUninstall() then
    begin
      TryUninstallByDisplayName('usbip-win2');
      TryUninstallByDisplayName('usbip');
    end;
  end;
end;
