# HyperTool

HyperTool ist ein Windows-Tool zur Steuerung von Hyper-V VMs mit moderner WinUI-3 Oberfläche, Tray-Menü und klaren One-Click-Aktionen.

## Überblick

- VM-Aktionen: Start, Stop, Hard Off, Restart, Konsole öffnen
- VM-Backup: Exportieren und Importieren mit Fortschritt in Prozent
- Netzwerk: adaptergenaues Switch verbinden/trennen (Multi-NIC) per Direktbuttons
- Host Network Popup: alle gefundenen Host-Adapter inkl. IP/Subnetz/Gateway/DNS, Badges für Gateway und Default Switch
- Snapshots: Tree-Ansicht (Parent/Child) mit Markierung für neuesten und aktuellen Stand
- Tray-Integration: VM starten/stoppen, Konsole öffnen, Snapshot, Switch umstellen; Menübereiche optional ausblendbar
- Collapsible Notifications Log: eingeklappt/ausgeklappt, Copy/Clear
- VM-Adapter umbenennen inkl. Eingabevalidierung
- Konfiguration per JSON, Logging mit Serilog

## Tech Stack

- .NET 8, WinUI 3 (Windows App SDK)
- MVVM mit CommunityToolkit.Mvvm
- Serilog für Datei-Logging
- Hyper-V-Operationen über PowerShell

## Voraussetzungen

- Windows 10 oder 11
- Hyper-V aktiviert
- PowerShell mit funktionsfähigem Hyper-V Modul (Get-VM)
- Für Entwicklung: .NET SDK 8.x
- Für USB-Funktionen: `usbipd-win` (HyperTool versucht bei Bedarf eine automatische Installation über `winget`)

## Projektstruktur

- HyperTool.sln
- src/HyperTool.Core
- src/HyperTool.WinUI
- src/HyperTool.Guest
- HyperTool.config.json
- build-winui.bat
- build-installer-winui.bat
- build-guest.bat
- build-installer-guest.bat
- build-all.bat
- dist/HyperTool.WinUI (Publish-Ausgabe)

## Schnellstart (Entwicklung)

1. dotnet restore HyperTool.sln
2. dotnet build HyperTool.sln -c Debug
3. dotnet run --project src/HyperTool.WinUI/HyperTool.WinUI.csproj

## HyperTool.Guest (neu)

`HyperTool.Guest` ist jetzt eine Guest-Desktop-App (WinUI) mit Host-ähnlichem Verhalten (Start/Exit-Visuals, Light/Dark Theme, Minimize-to-Tray, Start mit Windows). Die bisherigen Agent-Befehle bleiben per Startargument verfügbar.

- Header/Sidebar-Layout ist an den Host angepasst (Guest nutzt USB, Einstellungen, Info).
- Tray-Rechtsklick öffnet ein USB-zentriertes Guest Control Center (Einblenden/Ausblenden, Connect/Disconnect, Beenden).
- Single-Instance aktiv: ein zweiter Start blendet die laufende Instanz wieder ein.

Build/Run:

- dotnet build src/HyperTool.Guest/HyperTool.Guest.csproj -c Release
- build-guest.bat
- build-guest.bat version=1.2.0 no-pause
- dotnet run --project src/HyperTool.Guest/HyperTool.Guest.csproj (UI-Modus)
- dotnet run --project src/HyperTool.Guest/HyperTool.Guest.csproj -- once
- dotnet run --project src/HyperTool.Guest/HyperTool.Guest.csproj -- status
- dotnet run --project src/HyperTool.Guest/HyperTool.Guest.csproj -- run
- dotnet run --project src/HyperTool.Guest/HyperTool.Guest.csproj -- install-autostart
- dotnet run --project src/HyperTool.Guest/HyperTool.Guest.csproj -- install-autostart task
- dotnet run --project src/HyperTool.Guest/HyperTool.Guest.csproj -- autostart-status
- dotnet run --project src/HyperTool.Guest/HyperTool.Guest.csproj -- remove-autostart
- dotnet run --project src/HyperTool.Guest/HyperTool.Guest.csproj -- handshake

Konfiguration:

- Standardpfad: `%ProgramData%\HyperTool\HyperTool.Guest.json`
- Beim ersten Start wird eine Beispielkonfiguration erzeugt.
- Optionaler Override: `--config <Pfad>`
- Phase 1 enthält zusätzlich:
	- Autostart via `Run-Registry` oder `Task Scheduler`
	- strukturiertes Logfile (NDJSON) unter `%ProgramData%\HyperTool\logs`
	- Host-Handshake-Datei für spätere Host↔Guest-Kopplung unter `%ProgramData%\HyperTool\HyperTool.Guest.handshake.json`

## Build & Publish

Standard:

- build-winui.bat

Varianten:

- build-all.bat (interaktiv: Host/Guest, Installer, Versionen)
- build-all.bat host guest host-version=1.2.0 guest-version=1.2.1 no-host-installer guest-installer no-pause
- build-winui.bat framework-dependent no-pause
- build-winui.bat self-contained version=1.2.0
- build-guest.bat self-contained version=1.2.0
- build-guest.bat framework-dependent version=1.2.0
- build-guest.bat self-contained version=1.2.0 installer
- build-guest.bat no-installer no-pause
- build-installer-winui.bat (fragt Version interaktiv ab)
- build-installer-winui.bat version=1.2.0
- build-installer-guest.bat (fragt Version interaktiv ab)
- build-installer-guest.bat version=1.2.0

WinUI-Migration-Ausgabe liegt unter dist/HyperTool.WinUI.
Guest-Ausgabe liegt unter dist/HyperTool.Guest.

Installer-Ausgabe liegt unter dist/installer-winui (benötigt Inno Setup 6 / ISCC).
Guest-Installer-Ausgabe liegt unter dist/installer-guest (benötigt Inno Setup 6 / ISCC).
Der Installer ist für den self-contained WinUI-Build ausgelegt und enthält keine separate .NET Desktop Runtime-Abfrage.

## WinUI 3 Stand

- Produktivzweig ist WinUI-only (`HyperTool.Core` + `HyperTool.WinUI`).

## Konfiguration

Aktive Datei (Priorität):

1. `HYPERTOOL_CONFIG_PATH` (falls gesetzt)
2. Neuere Datei von:
	- `%LOCALAPPDATA%/HyperTool/HyperTool.config.json`
	- `HyperTool.config.json` im Installationsordner
3. Falls keine Datei existiert: `%LOCALAPPDATA%/HyperTool/HyperTool.config.json`

Im Config-Tab zeigt HyperTool den tatsächlich verwendeten Pfad als "Aktive Config" an.

Wichtige Felder:

- defaultVmName: bevorzugte VM
- lastSelectedVmName: letzte aktive VM
- defaultSwitchName: bevorzugter Switch
- vmConnectComputerName: vmconnect Host (z. B. `localhost` oder ein Zertifikats-Hostname)
- hns: HNS-Verhalten
- ui: Tray/Autostart Optionen
- ui.enableTrayMenu: blendet VM/Switch/Aktualisieren im Tray-Menü ein/aus (Show/Hide/Exit bleiben immer sichtbar)
- ui.theme: `Dark` oder `Light`
- update: GitHub Updateprüfung
- ui.trayVmNames: optionale Liste der VM-Namen, die im Tray-Menü erscheinen sollen (leer = alle)
- ui.startMinimized: App startet minimiert (in Verbindung mit Tray ideal für Hintergrundbetrieb)

VMs werden zur Laufzeit automatisch aus Hyper-V geladen (Auto-Discovery).

## Theme (Dark/Light)

- Umschaltung im Config-Tab über `Theme` (`Dark` / `Light`)
- Wechsel wird live auf die komplette UI angewendet
- Speicherung in `%LOCALAPPDATA%/HyperTool/HyperTool.config.json` unter `ui.theme`
- `Bright` wird aus älteren Configs weiterhin akzeptiert und automatisch zu `Light` normalisiert

## Hilfe-Popup

- Der `❔ Hilfe`-Button oben rechts öffnet ein eigenes Hilfe-Fenster (nicht mehr Info-Tab-Navigation)
- Enthält Kurz-Erklärungen zu Start/Stop, Network, Snapshots, HNS, Tray
- Schnellaktionen im Popup: `Logs öffnen`, `Config öffnen`, `GitHub Repo`

## Network & Host Adapter

- Netzwerk-Aktionen sind adaptergenau (pro VM-NIC auswählbar)
- `Host Network` zeigt alle gefundenen Host-Adapter, nicht nur Uplink-Adapter
- Default Switch (ICS) wird gesondert erkannt und mit Badge markiert

## Export & Import Details

- Export zeigt Fortschritt (0-100%) und prüft vorher den verfügbaren Speicherplatz
- Import läuft als neue VM (`-Copy -GenerateNewId`) und fragt Zielordner ab
- Bei Namenskonflikten wird automatisch ein eindeutiger Name mit Suffix erzeugt

## Tray Verhalten

- Das Tray-Icon bleibt aktiv (wichtig für minimierten Start und Wiederöffnen)
- Option `Tasktray-Menü aktiv` steuert nur Zusatzpunkte:
	- aktiv: VM Aktionen, Switch umstellen, Aktualisieren sichtbar
	- inaktiv: nur Show/Hide/Exit sichtbar

## Easter Egg

- Klick auf das Logo oben rechts startet eine kurze Dreh-Animation
- Optionaler Sound über `src/HyperTool.WinUI/Assets/logo-spin.wav` (Wiedergabe mit 30% Lautstärke)

## UI-Verhalten (wichtig)

- VM-Auswahl erfolgt über Chips im Header
- Hauptaktionen arbeiten immer auf der aktuell ausgewählten VM
- Network-Tab zeigt VM-Status + aktuellen Switch
- Notifications Log:
	- standardmäßig eingeklappt
	- eingeklappt: nur letzte Meldung
	- ausgeklappt: vollständige, scrollbare Liste + Copy/Clear

## Logging & Troubleshooting

Logpfade:

- dist/HyperTool.WinUI/logs
- Fallback: %LOCALAPPDATA%/HyperTool/logs

Wenn die App nicht startet:

1. Aus dist/HyperTool.WinUI starten
2. Logdateien prüfen
3. HyperTool.exe manuell in PowerShell starten

## Rechtehinweis

- Die App benötigt nicht grundsätzlich Adminrechte.
- Einzelne Hyper-V/HNS Aktionen können erhöhte Rechte benötigen.
- USB-Aktionen können UAC auslösen (Bind/Unbind/Detach sowie ggf. Start des `usbipd`-Dienstes).

## USB (usbipd) Verhalten

- HyperTool nutzt `usbipd.exe` als CLI-Bridge für USB-Share/-Unshare.
- Wenn `usbipd-win` fehlt, versucht HyperTool einmalig eine automatische Installation via `winget`.
- Läuft der `usbipd`-Dienst nicht, versucht HyperTool ihn automatisch zu starten (ggf. mit UAC).
- Falls Installation/Dienststart fehlschlägt, fällt die UI auf einen klaren USB-Status zurück statt hart zu brechen (Tool + Control Center).

## Update- und Installer-Flow

- HyperTool prüft GitHub Releases anhand der Versionsnummer (SemVer inkl. Prerelease).
- Wenn im Release ein Installer-Asset (`.exe`/`.msi`) erkannt wird, ist im Info-Tab der Button `Update installieren` nutzbar.
- Der Installer wird nach `%TEMP%\HyperTool\updates` geladen und direkt gestartet.
- Der ausgelieferte WinUI-Installer installiert nur HyperTool (self-contained), ohne zusätzliche Runtime-Installation im Setup-Dialog.
- Für eigene Releases: erst `build-winui.bat ...`, dann `build-installer-winui.bat version=x.y.z`, anschließend Setup-Datei als Release-Asset auf GitHub anhängen.