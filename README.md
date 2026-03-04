# HyperTool

HyperTool ist ein WinUI-3 Toolset für Hyper-V-Host und Windows-Guest mit Fokus auf schnelle VM-/Netzwerkaktionen und USB/IP-Workflows.

## Aktueller Release-Stand

- Version: **v2.3.5**
- Shared Folder ist end-to-end integriert: Host-Katalog/Freigaben + Guest-Mounting über `hypertool-file`.
- Guest nutzt `WinFsp` als zusätzliche Runtime für Shared-Folder-Mounts inkl. Runtime-Status, Installationshinweis und Quellenangabe.
- Host und Guest enthalten konsistente Runtime-Statusanzeigen (USB/Shared Folder) mit Installations- und Neustart-Aktionen.
- `Tool neu starten` ist in Host/Guest (USB/Shared Folder/Config) vereinheitlicht und zeigt wie beim Theme-Wechsel kurz den Reload-Screen.
- Guest-Info führt externe Quellen (`usbip-win2`, `winfsp`) nebeneinander mit symmetrischem Kartenlayout.

## Projekte

- HyperTool (Host): Hyper-V Steuerung, Netzwerk, Snapshots, USB-Share und Shared-Folder-Katalog.
- HyperTool Guest: USB-Client sowie Shared-Folder-Mounting gegen Host-Freigaben.
- Gemeinsame Basis in HyperTool.Core.

## Funktionen

### Host (HyperTool.exe)

- VM-Aktionen: Start, Stop, Hard Off, Restart, Konsole.
- Netzwerk: adaptergenaues Switch-Handling (auch Multi-NIC).
- Host-Network-Details: klare Status-Chips für `Gateway` (grün) und `Default Switch` (orange), dark/light lesbar.
- Snapshots: Baumdarstellung mit Restore/Delete/Create.
- USB: Refresh, Share, Unshare über usbipd.
- Tray Control Center: usbipd-Dienststatus (grün/rot), kompakter USB-Bereich und Installationsbutton bei fehlendem usbipd-win.
- Tray + Control Center mit Schnellaktionen.
- In-App Updatecheck und Installer-Update.

### Guest (HyperTool.Guest.exe)

- USB-Geräte vom Host laden, Connect/Disconnect.
- USB-Host-Sektion mit sichtbaren Transportmodus-Chips (Hyper-V Socket / IP-Mode) und modeabhängiger Aktivierung des Host-IP-Felds.
- Tray Control Center: usbip-win2-Status (grün/rot), kompakter USB-Bereich, Installationsbutton bei fehlendem Client und direkte Modusanzeige (Hyper-V Socket/IP).
- Shared Folder: Host-Katalog laden, Laufwerkszuordnungen anwenden und Mount-Status überwachen (WinFsp-basiert).
- Start mit Windows, Start minimiert, Minimize-to-Tray.
- Guest Control Center im Tray mit USB-Aktionen.
- Wenn Tasktray-Menü deaktiviert ist: nur Ein-/Ausblenden und Beenden.
- Theme-Unterstützung (Dark/Light) und Single-Instance-Verhalten.
- Theme-Neustart erhält die aktuell gewählte Menüseite in der Guest-App.

## Externe Runtime-Repositories (wichtig)

HyperTool vendort diese Projekte nicht als Produktabhängigkeit in die App, sondern nutzt installierte Laufzeiten:

- Host USB Runtime: dorssel/usbipd-win
  - Repository: https://github.com/dorssel/usbipd-win
- Guest USB Runtime: vadimgrn/usbip-win2
  - Repository: https://github.com/vadimgrn/usbip-win2
- Guest Shared-Folder Runtime: winfsp/winfsp
  - Repository: https://github.com/winfsp/winfsp

Hinweise:

- Alle Runtimes werden über deren eigene Releases/Lizenzen bezogen.
- Die HyperTool-Installer bieten optionale Online-Installation dieser Abhängigkeiten.
- Wenn eine Runtime fehlt, werden USB-Funktionen in der UI deaktiviert und mit Hinweis dargestellt.
- Für Shared-Folder-Mounts im Guest wird zusätzlich WinFsp benötigt; fehlt WinFsp, bleibt der Shared-Folder-Runtime-Status auf „Nicht installiert“.

## Lizenz & Drittanbieter-Hinweise

- HyperTool selbst steht unter der MIT-Lizenz (siehe `LICENSE`).
- Externe Runtimes (`usbipd-win`, `usbip-win2`, `winfsp`) sind eigenständige Projekte mit eigenen Lizenzen.
- In Host-/Guest-Info und in der Hilfe sind die jeweiligen Quellen verlinkt; verbindlich sind immer die Lizenztexte der Original-Repositories.

## Voraussetzungen

- Windows 10/11
- Für Host: Hyper-V aktiviert
- Für Entwicklung: .NET SDK 8.x
- Für Installer-Build: Inno Setup 6 (ISCC)

## Repository-Struktur

- HyperTool.sln
- src/HyperTool.Core
- src/HyperTool.WinUI
- src/HyperTool.Guest
- installer/HyperTool.iss
- installer/HyperTool.Guest.iss
- build-host.bat
- build-installer-host.bat
- build-guest.bat
- build-installer-guest.bat
- build-all.bat

## Build

### Host

- build-host.bat
- build-installer-host.bat version=2.3.5

### Guest

- build-guest.bat
- build-installer-guest.bat version=2.3.5

### Komplett

- build-all.bat
- build-all.bat version=2.3.5 host guest host-installer guest-installer no-pause

Ausgaben:

- dist/HyperTool.WinUI
- dist/HyperTool.Guest
- dist/installer-winui
- dist/installer-guest

## Konfiguration

Host-Konfigurationsdatei:

- HyperTool.config.json

Guest-Konfigurationsdatei:

- %ProgramData%/HyperTool/HyperTool.Guest.json

Relevante UI-Schalter:

- ui.enableTrayMenu (Host Tray-Menü erweitern/reduzieren)
- ui.MinimizeToTray bzw. Tasktray-Menü aktiv (Guest Control Center Verhalten)
- ui.startMinimized
- ui.theme

### Shared-Folder Transport (Guest)

Der Guest nutzt ausschließlich den Transport `hypertool-file` (Hyper-V Socket File Service).

Für `hypertool-file` wird zusätzlich eine installierte WinFsp Runtime im Guest benötigt.

Konfigurationsblock in `%ProgramData%/HyperTool/HyperTool.Guest.json`:

```json
"fileService": {
  "enabled": true,
  "mappingMode": "hypertool-file",
  "preferHyperVSocket": true
}
```

Hinweis zur Zuständigkeit:

- Host-seitige Freigaben/Optionen werden weiterhin über HyperTool Host verwaltet.
- Guest-seitiges Mapping/Transport wird über HyperTool Guest gesteuert.
- HyperTool.Guest mountet die Laufwerksbuchstaben direkt per WinFsp (Explorer sichtbar, ohne klassische SMB-Logon-Abhängigkeit).

## Update-Flow

- Updates basieren auf GitHub Releases.
- Asset-Auswahl für Host/Guest ist auf gemeinsame Releases abgestimmt.
- Installer werden nach %TEMP%/HyperTool/updates heruntergeladen und gestartet.

## Logging

- Host: %LOCALAPPDATA%/HyperTool/logs (Fallback je nach Startkontext)
- Guest: %ProgramData%/HyperTool/logs

## Rechtehinweis

- Nicht alle Funktionen benötigen Adminrechte.
- Hyper-V- und USB-Operationen können erhöhte Rechte/UAC erfordern.