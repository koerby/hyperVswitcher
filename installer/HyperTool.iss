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
AppPublisher=github.com/koerby
AppPublisherURL=https://github.com/koerby
AppSupportURL=https://github.com/koerby/HyperTool/issues
AppUpdatesURL=https://github.com/koerby/HyperTool/releases
VersionInfoVersion={#MyAppVersion}
VersionInfoCompany=github.com/koerby
VersionInfoDescription=HyperTool Setup
VersionInfoProductName=HyperTool
VersionInfoProductVersion={#MyAppVersion}
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
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
english.DesktopIconTask=Create desktop shortcut
english.AdditionalTasks=Additional tasks:
english.UninstallShortcut=Uninstall HyperTool
english.RunAfterInstall=Launch HyperTool
english.UsbipdInstallTask=Install usbipd-win for Host USB sharing / passthrough to Guest VMs (optional, requires internet connection)
english.UsbipdInstallFailed=usbipd-win could not be installed automatically. Please install it manually from https://github.com/dorssel/usbipd-win/releases and then restart HyperTool.
english.UsbipdInstallStarted=usbipd-win installation was started in background. The setup will now finish; if usbipd-win is still missing later, use the install button in HyperTool.
english.HyperToolFileServiceTask=Register HyperTool Hyper-V socket file service (optional, recommended)
english.HyperToolFileServiceFailed=HyperTool Hyper-V socket file service could not be registered automatically. Start HyperTool as administrator and retry in Shared Folder menu.
english.UninstallDepsTitle=Optional dependency cleanup
english.UninstallDepsText=Select optional dependencies that should also be removed during HyperTool uninstall.
english.UninstallUsbipdCheckbox=Also uninstall usbipd-win (if installed)

german.DesktopIconTask=Desktop-Verknüpfung erstellen
german.AdditionalTasks=Zusätzliche Aufgaben:
german.UninstallShortcut=HyperTool deinstallieren
german.RunAfterInstall=HyperTool starten
german.UsbipdInstallTask=usbipd-win fuer Host-USB-Sharing / Passthrough zu Guest-VMs installieren (optional, Internetverbindung erforderlich)
german.UsbipdInstallFailed=usbipd-win konnte nicht automatisch installiert werden. Bitte manuell von https://github.com/dorssel/usbipd-win/releases installieren und HyperTool danach neu starten.
german.UsbipdInstallStarted=Die usbipd-win Installation wurde im Hintergrund gestartet. Das Setup wird jetzt beendet; falls usbipd-win spaeter noch fehlt, nutze den Installationsbutton in HyperTool.
german.HyperToolFileServiceTask=HyperTool Hyper-V Socket File Service registrieren (optional, empfohlen)
german.HyperToolFileServiceFailed=HyperTool Hyper-V Socket File Service konnte nicht automatisch registriert werden. HyperTool als Administrator starten und im Shared-Folder-Menü erneut ausführen.
german.UninstallDepsTitle=Optionale Abhängigkeits-Bereinigung
german.UninstallDepsText=Wähle optionale Abhängigkeiten aus, die bei der HyperTool-Deinstallation ebenfalls entfernt werden sollen.
german.UninstallUsbipdCheckbox=usbipd-win ebenfalls deinstallieren (falls installiert)

[Tasks]
Name: "desktopicon"; Description: "{cm:DesktopIconTask}"; GroupDescription: "{cm:AdditionalTasks}"
Name: "installusbipd"; Description: "{cm:UsbipdInstallTask}"; GroupDescription: "{cm:AdditionalTasks}"; Flags: checkedonce; Check: not IsUsbipdInstalled

[Files]
Source: "{#MySourceDir}\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion createallsubdirs; Excludes: "HyperTool.config.json,Scripts\HyperToolCredentialProvisioning.ps1,tools\*Credential*.ps1"
Source: "{#MySourceDir}\HyperTool.config.json"; DestDir: "{app}"; Flags: onlyifdoesntexist ignoreversion uninsneveruninstall

[Registry]
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\6c4eb1be-40e8-4c8b-a4d6-5b0f67d7e40f"; ValueType: string; ValueName: "ElementName"; ValueData: "HyperTool Hyper-V Socket USB Tunnel"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\67c53bca-3f3d-4628-98e4-e45be5d6d1ad"; ValueType: string; ValueName: "ElementName"; ValueData: "HyperTool Hyper-V Socket Diagnostics"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\e7db04df-0e32-4f30-a4dc-c6cbc31a8792"; ValueType: string; ValueName: "ElementName"; ValueData: "HyperTool Hyper-V Socket Shared Folder Catalog"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\54b2c423-6f79-47d8-a77d-8cab14e3f041"; ValueType: string; ValueName: "ElementName"; ValueData: "HyperTool Hyper-V Socket Host Identity"; Flags: uninsdeletekey
Root: HKLM64; Subkey: "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\91df7cec-c5ba-452a-b072-42e5f672d5f9"; ValueType: string; ValueName: "ElementName"; ValueData: "HyperTool Hyper-V Socket File Service"; Flags: uninsdeletekey

[Icons]
Name: "{group}\HyperTool"; Filename: "{app}\HyperTool.exe"; IconFilename: "{app}\Assets\HyperTool.ico"
Name: "{group}\{cm:UninstallShortcut}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\HyperTool"; Filename: "{app}\HyperTool.exe"; IconFilename: "{app}\Assets\HyperTool.ico"; Tasks: desktopicon

[Run]
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""$n='HyperTool USB Discovery (UDP-In)'; if(-not (Get-NetFirewallRule -DisplayName $n -ErrorAction SilentlyContinue)){{ New-NetFirewallRule -DisplayName $n -Direction Inbound -Action Allow -Protocol UDP -LocalPort 32491 -Profile Domain,Private -Program '{app}\HyperTool.exe' | Out-Null }}"""; Flags: runhidden waituntilterminated
Filename: "{sys}\WindowsPowerShell\v1.0\powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -Command ""$n='HyperTool USB Discovery (UDP-Out)'; if(-not (Get-NetFirewallRule -DisplayName $n -ErrorAction SilentlyContinue)){{ New-NetFirewallRule -DisplayName $n -Direction Outbound -Action Allow -Protocol UDP -RemotePort 32491 -Profile Domain,Private -Program '{app}\HyperTool.exe' | Out-Null }}"""; Flags: runhidden waituntilterminated
Filename: "{app}\HyperTool.exe"; Description: "{cm:RunAfterInstall}"; Flags: postinstall nowait skipifsilent

[UninstallRun]
Filename: "{sys}\taskkill.exe"; Parameters: "/IM HyperTool.exe /T"; Flags: runhidden waituntilterminated skipifdoesntexist; RunOnceId: "HyperTool-Uninstall-StopApp"
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=""HyperTool USB Discovery (UDP-In)"""; Flags: runhidden waituntilterminated; RunOnceId: "HyperTool-Uninstall-DeleteFirewall-UDP-In"
Filename: "{sys}\netsh.exe"; Parameters: "advfirewall firewall delete rule name=""HyperTool USB Discovery (UDP-Out)"""; Flags: runhidden waituntilterminated; RunOnceId: "HyperTool-Uninstall-DeleteFirewall-UDP-Out"
Filename: "{cmd}"; Parameters: "/C reg delete ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\6c4eb1be-40e8-4c8b-a4d6-5b0f67d7e40f"" /f"; Flags: runhidden waituntilterminated; RunOnceId: "HyperTool-Uninstall-DeleteSvc-UsbTunnel"
Filename: "{cmd}"; Parameters: "/C reg delete ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\67c53bca-3f3d-4628-98e4-e45be5d6d1ad"" /f"; Flags: runhidden waituntilterminated; RunOnceId: "HyperTool-Uninstall-DeleteSvc-Diagnostics"
Filename: "{cmd}"; Parameters: "/C reg delete ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\e7db04df-0e32-4f30-a4dc-c6cbc31a8792"" /f"; Flags: runhidden waituntilterminated; RunOnceId: "HyperTool-Uninstall-DeleteSvc-Catalog"
Filename: "{cmd}"; Parameters: "/C reg delete ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\54b2c423-6f79-47d8-a77d-8cab14e3f041"" /f"; Flags: runhidden waituntilterminated; RunOnceId: "HyperTool-Uninstall-DeleteSvc-HostIdentity"
Filename: "{cmd}"; Parameters: "/C reg delete ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Virtualization\GuestCommunicationServices\91df7cec-c5ba-452a-b072-42e5f672d5f9"" /f"; Flags: runhidden waituntilterminated; RunOnceId: "HyperTool-Uninstall-DeleteSvc-FileService"
Filename: "{cmd}"; Parameters: "/C winget.exe uninstall --id dorssel.usbipd-win --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveUsbipdOnUninstall; RunOnceId: "HyperTool-Uninstall-Optional-Usbipd-WingetId"
Filename: "{cmd}"; Parameters: "/C winget.exe uninstall --name ""usbipd-win"" --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveUsbipdOnUninstall; RunOnceId: "HyperTool-Uninstall-Optional-Usbipd-WingetName"
Filename: "{cmd}"; Parameters: "/C ""%LOCALAPPDATA%\Microsoft\WindowsApps\winget.exe"" uninstall --id dorssel.usbipd-win --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveUsbipdOnUninstall; RunOnceId: "HyperTool-Uninstall-Optional-Usbipd-WingetId-LocalAppData"
Filename: "{cmd}"; Parameters: "/C ""%LOCALAPPDATA%\Microsoft\WindowsApps\winget.exe"" uninstall --name ""usbipd-win"" --exact --silent --accept-source-agreements --accept-package-agreements"; Flags: runhidden waituntilterminated; Check: ShouldRemoveUsbipdOnUninstall; RunOnceId: "HyperTool-Uninstall-Optional-Usbipd-WingetName-LocalAppData"
Filename: "{commonpf64}\usbipd-win\unins000.exe"; Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; Flags: runhidden waituntilterminated skipifdoesntexist; Check: ShouldRemoveUsbipdOnUninstall; RunOnceId: "HyperTool-Uninstall-Optional-Usbipd-Unins64"
Filename: "{commonpf}\usbipd-win\unins000.exe"; Parameters: "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART"; Flags: runhidden waituntilterminated skipifdoesntexist; Check: ShouldRemoveUsbipdOnUninstall; RunOnceId: "HyperTool-Uninstall-Optional-Usbipd-Unins32"

[Code]
var
  RemoveUsbipdOnUninstall: Boolean;
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

procedure CloseRunningHostApp;
var
  ResultCode: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/IM HyperTool.exe /F /T', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

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
    Result := FileExists(ExpandConstant('{commonpf}\\usbipd-win\\usbipd.exe'));
  end;
end;

function ShouldRemoveUsbipdOnUninstall(): Boolean;
begin
  Result := RemoveUsbipdOnUninstall;
end;

procedure PromptUninstallOptions;
begin
  RemoveUsbipdOnUninstall := False;

  if IsUsbipdInstalled() then
  begin
    if SuppressibleMsgBox(
      ExpandConstant('{cm:UninstallDepsText}') + #13#10#13#10 + ExpandConstant('{cm:UninstallUsbipdCheckbox}'),
      mbConfirmation,
      MB_YESNO,
      IDNO) = IDYES then
    begin
      RemoveUsbipdOnUninstall := True;
    end;
  end;
end;

function IsHyperToolFileServiceMissing(): Boolean;
var
  Value: string;
begin
  Result :=
    not RegQueryStringValue(
      HKLM64,
      'SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Virtualization\\GuestCommunicationServices\\91df7cec-c5ba-452a-b072-42e5f672d5f9',
      'ElementName',
      Value);
end;

procedure TryInstallHyperToolFileService;
var
  ResultCode: Integer;
begin
  if not FileExists(ExpandConstant('{app}\HyperTool.exe')) then
  begin
    MsgBox(ExpandConstant('{cm:HyperToolFileServiceFailed}'), mbInformation, MB_OK);
    Exit;
  end;

  if not Exec(ExpandConstant('{app}\HyperTool.exe'), '--register-sharedfolder-socket', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) or (ResultCode <> 0) then
  begin
    MsgBox(ExpandConstant('{cm:HyperToolFileServiceFailed}'), mbInformation, MB_OK);
  end;
end;

procedure TryInstallUsbipdFromInternet;
var
  ResultCode: Integer;
  WingetArgs: string;
begin
  if not WizardIsTaskSelected('installusbipd') then
    Exit;

  if IsUsbipdInstalled() then
    Exit;

  WingetArgs :=
    'install --id dorssel.usbipd-win --exact --silent --accept-source-agreements --accept-package-agreements';

  if Exec(ExpandConstant('{cmd}'), '/C start "" /MIN winget.exe ' + WingetArgs, '', SW_HIDE, ewNoWait, ResultCode) then
  begin
    MsgBox(ExpandConstant('{cm:UsbipdInstallStarted}'), mbInformation, MB_OK);
    Exit;
  end;

  MsgBox(ExpandConstant('{cm:UsbipdInstallFailed}'), mbInformation, MB_OK);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
  begin
    CloseRunningHostApp;
  end;

  if CurStep = ssPostInstall then
  begin
    TryInstallUsbipdFromInternet;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if (CurUninstallStep = usAppMutexCheck) and (not UninstallOptionsPrompted) then
  begin
    CloseRunningHostApp;
    UninstallOptionsPrompted := True;
    PromptUninstallOptions;
  end;

  if (CurUninstallStep = usUninstall) and (not OptionalDepsUninstallExecuted) then
  begin
    OptionalDepsUninstallExecuted := True;

    if ShouldRemoveUsbipdOnUninstall() then
    begin
      TryUninstallByDisplayName('usbipd-win');
    end;
  end;
end;
