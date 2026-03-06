# HyperTool Release Notes

## v2.3.8

### Highlights

- Host-Netzprofil-Workflow ist jetzt direkt in der UI bedienbar (VM-View + Host-Network pro Adapter).
- Netzprofil-Änderungen nutzen UAC-Elevation, damit die Funktion auch ohne als Admin gestartete App nutzbar ist.
- Neuer optionaler NumLock-Wächter im Host, steuerbar per Checkbox, mit konfigurierbarem Hintergrund-Intervall.

### Neu

- Host VM-Ansicht:
	- Sichtbarer Host-Netzprofilstatus im Footer (`Öffentlich` / `Privat` / `Domäne`).
	- Direkter Aktionsbutton mit Gegenzustand (z. B. `Auf Privat umstellen` bei `Öffentlich`).
	- Bei `Domäne` ist die Aktion bewusst gesperrt.
- Host-Network Fenster:
	- Profil-Chips pro Adapter (`Privat`/`Öffentlich`/`Domäne`) ergänzt.
	- Direkte Profil-Umstellung pro Adapter per Aktionsbutton.
- Config:
	- Neue Option `ui.restoreNumLockAfterVmStart` (Checkbox in der Oberfläche).
	- Erweiterte Option `ui.numLockWatcherIntervalSeconds` (nur Config-Datei, Default `30`).

### Verbessert

- Netzprofil-Fehlerhandling:
	- Klare Meldungen für UAC-Abbruch, fehlende Rechte, Domain-Profil-Sperre und GPO-Blockierung.
	- UI-Zustände bleiben konsistent bei fehlgeschlagenen Umschaltungen.
- Import-Flow:
	- Zielordner-Handling für `copy/register/restore` präzisiert.
	- Import-Hinweise in UI und Konfiguration klarer strukturiert.
- Snapshot-Flow:
	- `Create` per Dialog (Name/Beschreibung).
	- `Restore/Delete` mit Bestätigungsdialogen.

### Behoben

- Host USB-UI: robustere Selection-Synchronisierung gegen Index/State-Race in WinRT-ListView-Brücken.
- Troll-Overlay Host/Guest: Shake/Wobble/Warp wiederhergestellt und Reset/Centering stabilisiert.
- Update-Sicherheit: Konfigurationsdateien werden bei Updates nicht mehr unbeabsichtigt überschrieben (Host/Guest Installer + Laufzeitpfade).

### Doku

- README auf `v2.3.8` aktualisiert.

## v2.3.7

### Highlights

- Host- und Guest-Oberflächen wurden für den täglichen VM-Workflow sichtbar entschlackt und stärker vereinheitlicht.
- VM-Chips, Header-Status und Kontextaktionen sind im Host präziser und schneller bedienbar.
- Ungespeicherte Konfigurationsänderungen reagieren in Host und Guest konsistent beim Wechseln/Neu-Laden.

### Neu

- Host VM-Kontextaktionen:
	- `Als Default-VM setzen` direkt über VM-Chip-/VM-Menü verfügbar.
	- `Schnellstart-Verknüpfung erstellen` direkt über VM-Chip-/VM-Menü verfügbar.
- Info-Menü (Host/Guest):
	- Neuer gelber `Buy Me a Coffee`-Button mit Kaffee-Icon und Direktlink: `https://buymeacoffee.com/koerby`.
- Guest USB:
	- Option zum automatischen USB-Disconnect beim Beenden ergänzt.
	- USB-Refresh beim Guest-Start verbessert, damit Device-Status früher konsistent ist.
- Header-Status (Host):
	- `Selected VM` zeigt den aktuellen State als farbigen Status-Chip (Running grün, Off rot), theme-sensitiv für Dark/Light.
- Host/Guest Easteregg:
	- Ein neues, verstecktes Easteregg wurde eingebaut - ohne Spoiler, nur so viel: Es lohnt sich, die UI aufmerksam zu erkunden.


### Verbessert

- Host VM-Chips:
	- Chip-Breiten verhalten sich stärker inhaltsbasiert (kurze VM-Namen wirken nicht mehr unnötig breit).
	- PC-Icon und Default-Stern wurden visuell vergrößert und entquetscht, ohne die Chip-Größe aufzublähen.
	- Default-Markierung als Badge/Overlay optisch präzisiert.
- Guest Header:
	- VM-Chips stabil unter dem Titel platziert, um Resize-Jitter und inkonsistente Rückwechsel zu vermeiden.
- Layout/UX (Host/Guest):
	- System-/Update-Bereiche und Abstände in mehreren Ansichten nachgeschliffen.
	- Checkbox-/Header-Abstände konsistenter für bessere Lesbarkeit.
	- Optionale Runtime-Aufgaben im Installer sichtbarer gemacht, damit notwendige Komponenten    	schneller erkennbar sind.


### Behoben

- Host:
	- Potenzieller UI-Freeze beim Verwerfen (`Nein`) ungespeicherter Änderungen adressiert.
	- Reload-Pfad nach `Nein` auf konsistentes Snapshot-Reload umgestellt.
- Guest:
	- `Nein`-/Reload-Verhalten beim Verwerfen ungespeicherter Änderungen robuster gemacht.
	- Menüwechsel-Prompt ergänzt, damit Änderungen nicht still verworfen werden.
	- USB/IP-Client-Erkennung robuster gemacht, damit "Nicht installiert"-/"Verfügbar"-Status nach Installation/Deinstallation konsistent ist.

### Doku

- README auf `v2.3.7` aktualisiert (Release-Stand, Feature-Hinweise, Build-Beispiele).

## v2.3.5

### Highlights

- Shared Folder wurde vollständig in Host und Guest integriert (Katalog, Mapping, Runtime-Status, Diagnose, UI-Gating).
- Host/Guest-Runtime-UX ist jetzt durchgängig konsistent: Status, Installationsbuttons, Neustart-Buttons und Reload-basierter Tool-Neustart.
- Guest-Info und Hilfe wurden inhaltlich erweitert (inkl. WinFsp als externe Quelle) und visuell präzisiert.

### Neu

- Shared Folder (Host/Guest):
	- Host verwaltet Freigaben zentral, Guest mappt über `hypertool-file`.
	- Guest nutzt WinFsp als Mount-Runtime; fehlende Runtime wird explizit angezeigt.
	- Host-/Guest-Feature-Gating für Shared Folder inkl. klarer Aktiv/Inaktiv-Status und Overlay-Hinweise.
- Info/Hilfe Guest:
	- Neue externe Quelle `winfsp/winfsp` im Info-Menü inkl. direktem Quellen-Link.
	- Externe Quellenkarten (`usbip-win2`, `winfsp`) nebeneinander im 50/50-Layout.
	- Kartenlayout mit identischer Mindesthöhe und abgestimmter vertikaler Ausrichtung.
- Config UX:
	- `Tool neu starten` zusätzlich in Host- und Guest-Config-Headern neben `Speichern` und `Neu laden`.

### Verbessert

- Neustartverhalten:
	- Host-`Tool neu starten` entspricht jetzt exakt dem Theme-Wechsel-Ablauf (kurzer Reload-Screen, danach Reopen).
	- Guest-`Tool neu starten` nutzt denselben Reload-Flow über den bestehenden Theme-Reopen-Mechanismus.
	- Einheitliches Icon/Label-Schema für alle `Tool neu starten` Buttons.
- Runtime-Status/UI:
	- Installations-/Neustart-Buttons werden bei erfüllten Runtime-Abhängigkeiten automatisch ausgeblendet.
	- Guest-Status priorisiert fehlenden USB-Client klar als „Nicht installiert“.
	- Guest-Fensterhöhe moderat reduziert für die 4-Menü-Struktur (USB, Share, Einstellungen, Info), ohne Layout-Bruch.
- Installer/Uninstaller:
	- Optionale Runtime-Deinstallation robuster durch Registry-basierte Erkennung und Fallback-Uninstall-Aufrufe.
	- Host/Guest-Installertexte und Defaults konsolidiert; Startverhalten nach Installation präzisiert.

### Doku / Lizenz / Cleanup

- README auf `v2.3.5` aktualisiert (Features, Build-Beispiele, Runtime-/Lizenzhinweise).
- LICENSE um Third-Party-Runtime-Notice (`usbipd-win`, `usbip-win2`, `winfsp`) ergänzt und präzisiert.
- Hilfe-Texte in Host/Guest um Shared-Folder-/WinFsp-/Tool-Neustart-Kontext erweitert.

## v2.1.7

### Highlights

- Host und Guest nutzen jetzt konsistenten `Tool neu starten`-Flow mit kurzem Reload-Screen (analog Theme-Wechsel).
- Config-Bereiche in Host und Guest wurden um `Tool neu starten` ergänzt (neben `Speichern` / `Neu laden`).
- Guest-Info dokumentiert nun zusätzlich die externe Shared-Folder-Runtime `WinFsp` inkl. direktem Quellen-Link.

### Neu

- Guest Info:
	- Neue Karte `Externe Shared-Folder Runtime` mit Quelle `winfsp/winfsp`.
	- Neuer Button `WinFsp Quelle` im Info-Aktionsbereich.
- UI Konsistenz:
	- Einheitliches Icon/Label-Schema für alle `Tool neu starten` Buttons in Host und Guest.

### Verbessert

- Hilfe Host:
	- Config/Info-Beschreibung enthält jetzt den Reload-basierten `Tool neu starten` Ablauf.
- Hilfe Guest:
	- Shared-Folder-Hinweis enthält explizit die WinFsp-Abhängigkeit.
	- Einstellungs-Hinweis ergänzt um `Tool neu starten` mit Reload-Screen.

### Doku / Lizenzhinweise

- README um WinFsp-Abhängigkeit und konsolidierte Drittanbieter-/Lizenzhinweise erweitert.
- Lizenzdatei um klare Hinweise zu externen Runtimes (`usbipd-win`, `usbip-win2`, `winfsp`) ergänzt.

## v2.1.6

### Highlights

- Host- und Guest-Tray-Control-Center zeigen den USB-Runtime-Status klar über Farbpunkt und Statuszeile (grün/rot).
- Host-Network und Guest-USB wurden auf konsistente, moderne Status-Chips umgestellt.
- Guest-Info-Diagnose ist kompakter, mit sauber rechts platziertem Test-Button ohne Einfluss auf Zeilenabstände.

### Neu

- Host Tray USB:
	- Runtime-Statusanzeige für usbipd-Dienst (aktiv/inaktiv/nicht installiert).
	- Installationsbutton bei fehlendem usbipd-win mit Download aus dem offiziellen GitHub-Release.
- Host Network:
	- Adapter-Detailansicht mit Status-Chips für `Gateway` und `Default Switch`.
	- Badge-Farblogik für klare Semantik: `Gateway` grün, `Default Switch` orange.
- Guest Tray USB:
	- Runtime-Statusanzeige für usbip-win2 Client.
	- Installationsbutton bei fehlendem usbip-win2 mit Download aus dem offiziellen GitHub-Release.
	- Modusanzeige im USB-Bereich (Hyper-V Socket / IP-Fallback) als klickbare Status-Chips.
	- Modusabhängige Aktivierung/Anzeige des Host-IP-Eingabefelds.

### Verbessert

- USB-Bereich in Host und Guest kompakter aufgebaut (geringere Abstände, klarere Informationsdichte).
- USB-Aktionszustände orientieren sich stärker am Runtime-Status.
- Guest aktualisiert die Transportmodus-Anzeige nach Socket-Umschaltung unmittelbarer im UI.
- Guest-Themewechsel erhält die aktuell gewählte Menüseite nach dem Neustart der Oberfläche.
- Info-/Diagnosebereich im Guest auf reduzierte Textdichte und konsistente Abstände optimiert.

### Doku / Hilfe

- Host-Hilfe um aktuelle Hinweise zu Host-Network-Status-Chips ergänzt.
- Guest-Hilfe um aktuelle Hinweise zu Transport-Status-Chips und Live-Diagnose ergänzt.
- README auf v2.1.6 mit finalem Feature-Stand (Host/Guest UI, Build-Aufruf, Runtime-Hinweise) aktualisiert.

## v2.1.4

### Highlights

- Hyper-V Socket Diagnosepfad zwischen Guest und Host erweitert und robuster gemacht.
- Info-Bereiche in Host/Guest kompakter ausgerichtet; relevante Diagnoseanzeigen gezielt vereinfacht.
- Guest-Option „Beim Start auf Updates prüfen“ vollständig an Config und Startup-Verhalten angebunden.

### Neu

- Guest Info:
	- Neuer Diagnose-Button „Hyper-V Socket testen“ inkl. Ergebnisrückmeldung.
	- Diagnose-Button im Info-Bereich rechts ausgerichtet.
- Host Diagnose:
	- Host-seitiger Listener für Guest-Diagnose-Ack aus Hyper-V Socket Testpfad.
	- Zusätzliche Telemetrie-/Logfelder für Transportpfad (`hyperv` / `ip-fallback`).

### Verbessert

- Transport/Logging:
	- Transportpfad wird in Erfolg- und Fehlerfällen explizit protokolliert.
	- Hyper-V-first Verhalten mit klarerem IP-Fallback-Verhalten stabilisiert.
- UI/UX:
	- Info-Kopfzeilen in Host und Guest kompakter (Info + Version in einer Zeile).
	- Host-Info zeigt keinen überflüssigen Text „Fallback auf IP aktiv“ mehr.
	- Disconnect-Refresh im Guest bewusst verzögert (3 Sekunden) für stabilere Gerätezustände.
- Installer/Abhängigkeiten:
	- Host-App versucht fehlende usbipd-Runtime nicht mehr in-app nachzuinstallieren.
	- Optionaler Installer-Flow für USB-Runtimes klarer abgegrenzt.

### Behoben

- Guest Settings:
	- Checkbox „Beim Start auf Updates prüfen“ war deaktiviert und nicht gespeichert; jetzt persistiert und wirksam.
	- Startup-Updatecheck läuft nur noch, wenn die Option aktiv ist.
- Control Center:
	- Visuelles „Zappeln“ beim Öffnen/Positionieren reduziert.
- Diagnosepfad:
	- Mehrere Stabilitätsprobleme im Hyper-V Socket Testablauf und bei Fallback-Übergängen adressiert.

## v2.1.1

### Highlights

- Host und Guest nutzen jetzt ein konsistentes externes USB-Runtime-Modell mit optionaler Online-Installation im Setup.
- USB-Bereiche in UI und Control Center reagieren robuster auf fehlende Abhängigkeiten und zeigen klare Hinweise.
- Guest Control Center wurde für den USB-Bereich und das kompakte Tray-Verhalten sichtbar überarbeitet.

### Neu

- Host Installer (`HyperTool.iss`):
	- Optionale Aufgabe zur Installation von usbipd-win aus dem offiziellen Release-Feed.
	- Kein erzwungener Installationsabbruch, wenn die optionale Runtime nicht installiert wird.
- Guest Installer (`HyperTool.Guest.iss`):
	- Optionale Aufgabe zur Installation von usbip-win2 aus dem offiziellen Release-Feed.
	- Silent-Install mit expliziter Komponentenwahl ohne GUI-Komponente.
- In-App Update:
	- Verbesserte Asset-Auswahl für kombinierte Host/Guest-Releases über Installer-Hints.

### Verbessert

- Host USB:
	- Laufzeitprüfung und Deaktivierung der USB-Aktionen bei fehlendem usbipd.
	- Klarer Laufzeit-Hinweis im USB-Bereich.
	- Info-Bereich mit separatem Hinweis auf externe Quelle/Lizenzkontext.
- Guest USB:
	- Stabilere Statusdarstellung für verbundene Geräte.
	- Refresh/Connect/Disconnect im Guest Control Center sauber nebeneinander.
	- Dynamische Control-Center-Höhe abhängig vom Tray-Menü-Modus.
	- Bei deaktiviertem Tasktray-Menü nur Ein-/Ausblenden und Beenden.
- Guest Notifications:
	- Copy/Clear-Handling analog zum Host ergänzt.

### Behoben

- Guest Attach-Flow nutzt keine nicht unterstützte usbip-Option mehr.
- Mehrere Probleme bei USB-Status-Refresh direkt nach Disconnect reduziert.
- Control-Center-Button-Layout in Host/Guest konsistenter umgesetzt.

### Externe Abhängigkeiten

- Host USB Runtime: dorssel/usbipd-win
- Guest USB Runtime: vadimgrn/usbip-win2

Hinweis: HyperTool verweist auf diese externen Projekte und deren Lizenzen; Installationen erfolgen optional über die jeweiligen offiziellen Releases.

## v2.0.0

### Highlights

- Major Release: HyperTool ist jetzt vollständig auf WinUI 3 umgestellt.
- Modernisierte Oberfläche mit konsistentem Dark/Light Theme und verbessertem Window-/Tray-Verhalten.
- Build- und Release-Prozess auf WinUI-only vereinheitlicht (App + Installer über BAT-Skripte).
- Installer-Flow für self-contained WinUI vereinfacht (keine separate Runtime-Abfrage im Setup).

### Neu

- WinUI-3 App als neue Hauptanwendung (`HyperTool.Core` + `HyperTool.WinUI`).
- Überarbeiteter Theme-Flow mit sauberem Übergang und robustem Rebuild der Hauptansicht.
- Inno-Setup-Installer für self-contained WinUI-Auslieferung ohne zusätzliche Runtime-Installation.
- WinUI-Build-/Installer-Pipeline:
	- `build-host.bat`
	- `build-installer-host.bat`

### Verbessert

- Export-Statusanzeige überarbeitet:
	- Fortschritt wird nur angezeigt, wenn Hyper-V verlässliche Werte liefert.
	- Irreführende Sprünge/Flicker in Prozentanzeige und Progressbar reduziert.
	- Monotones Fortschrittsverhalten (kein Zurückspringen im Balken).
- Fehlerrobustheit bei Hyper-V Aktionen verbessert (klarere Meldungen für Berechtigung/PowerShell-Fehler).
- Repository auf WinUI-only bereinigt (Legacy-WPF-Struktur entfernt).

### Behoben

- Fataler Fehler beim VM-Export in der Speicherplatzprüfung (PowerShell-RegEx/UNC-Parsing) behoben.
- Export-Fehlerpfad stabilisiert: Ausnahmen führen nicht mehr zu hartem App-Abbruch.
- Mehrere Probleme im Theme-Wechsel-/Rebuild-Flow behoben.

### Kompatibilität

- Windows 10/11
- Hyper-V aktiviert
- Keine separate .NET Desktop Runtime-Installation über den HyperTool-Installer erforderlich (self-contained Build).

## v1.3.4

### Highlights

- App-/Tray-Icon überarbeitet: saubere Transparenz außen herum, ohne den alten HyperTool-Iconstil zu verlieren.
- Fensteroptik modernisiert: echte abgerundete App-Ecken und weicheres Gesamtbild.
- „Modern clean“-Feintuning für Buttons, VM-Chips und Hauptpanels umgesetzt.
- Dark- und Light-Theme weiterhin vollständig unterstützt und konsistent gehalten.

### Neu

- Fenster-Rundung technisch ergänzt:
	- Abgerundete Fensterecken werden jetzt aktiv auf Window-Ebene angewendet.
	- Rundung wird bei Größenänderung/Fensterstatus sauber aktualisiert.
	- Bei maximiertem Fenster wird korrekt auf rechteckige Darstellung zurückgeschaltet.
- Theme-Interaktion erweitert:
	- Neue Theme-Brushes für Button `Hover` und `Pressed` (Dark/Light).

### Verbessert

- Icons:
	- Altes `HyperTool.ico` bleibt Basis.
	- Äußerer Hintergrund/Verlauf wurde transparent gemacht für sauberere Darstellung in `.exe` und Tasktray.
	- Tray nutzt dediziertes Icon-Fallback, dadurch konsistenteres Erscheinungsbild bei kleinen Größen.
- UI/Design:
	- Rahmenfarben in Dark/Light subtil entschärft (weniger harte Kanten).
	- `ActionButton` mit modernerem Verhalten (Hover/Pressed/Disabled, besseres Padding, Hand-Cursor).
	- VM-Auswahl-Chips leicht modernisiert (Padding/Interaktionsfeedback), ohne Funktionsänderung.
	- Hauptpanel-Abstände und Card-Padding dezent erhöht für luftigere, ruhigere Oberfläche.

### Behoben

- Fensterkanten wirkten trotz vorheriger Anpassung teils noch eckig; jetzt echte Rundung der App-Form.
- Uneinheitliche Button-/Chip-Anmutung wurde vereinheitlicht, ohne bestehende Abläufe zu verändern.

### Kompatibilität

- Windows 10/11
- Hyper-V aktiviert
- .NET 8

## v1.3.3

### Highlights

- Netzwerkverwaltung auf Multi-NIC erweitert: adaptergenaues Verbinden/Trennen statt pauschal pro VM.
- Host-Network Popup deutlich ausgebaut: alle gefundenen Host-Adapter inkl. Netzwerkdetails und Badge-Infos.
- Snapshot-Bereich auf Baumansicht mit Status-Badges umgestellt (neuester/aktueller Stand).
- Export/Import-Workflow robuster: Prozent-Fortschritt, Speicherplatzprüfung, sicherer Import als neue VM.
- Tray-Verhalten feiner steuerbar: Show/Hide/Exit immer verfügbar, Fachmenüs optional ausblendbar.

### Neu

- Network-Tab:
	- Auswahl von VM-Netzwerkadaptern (pro Adapter eigener Switch-Connect/Disconnect).
	- `Host Network`-Button mit Detailfenster für Host-Adapter.
- Host-Network Fenster:
	- Anzeige von Adapter, Beschreibung, IP, Subnetz, Gateway und DNS.
	- Gateway-Badge für Adapter mit Gateway.
	- Default-Switch-Badge für Hyper-V Default Switch (ICS).
- Snapshots:
	- Tree-View (Parent/Child) statt flacher Liste.
	- Kennzeichnung `Neueste` und `Jetzt`.
- Config/VM:
	- Tray-Adapter pro VM konfigurierbar.
	- Umbenennen von VM-Adaptern inkl. Validierung (leer/ungültige Zeichen/Duplikate/identischer Name).
- Tray:
	- Neue Option `Tasktray-Menü aktiv`.
	- Bei deaktivierter Option bleiben `Show`, `Hide`, `Exit` sichtbar; `VM Aktionen`, `Switch umstellen`, `Aktualisieren` werden ausgeblendet.
- Easter Egg:
	- Klick auf Logo startet Rotation und spielt optionalen custom WAV-Sound.

### Verbessert

- Dropdown-Usability: Aufklappen über Klick auf die gesamte ComboBox-Fläche.
- Host-Adapter-Erkennung robuster inklusive Default-Switch-Fallbacks.
- Sortierung im Host-Network Fenster verbessert (Gateway-relevante Adapter zuerst).
- Dunkelmodus-/Kontextmenü-Darstellung in VM-Chips überarbeitet.
- Host-Network Fenstergröße angepasst, um unnötige Scrollbars zu reduzieren.

### Export/Import

- Export:
	- Fortschrittsanzeige in Prozent.
	- Vorabprüfung auf ausreichend freien Speicher im Zielpfad.
- Import:
	- Immer als neue VM (`-Copy -GenerateNewId`).
	- Zielpfad wird abgefragt.
	- Namenskonflikte werden automatisch mit Suffix aufgelöst.

### Behoben

- Default Switch (ICS) wurde im Host-Network Popup teilweise ohne Details angezeigt.
- Tray-Switching bei Multi-NIC ist jetzt auf konfigurierbaren Adapter begrenzbar.
- Mehrere UI-Konsistenzprobleme im Netzwerk-/Tray-/Darkmode-Bereich.

### Kompatibilität

- Windows 10/11
- Hyper-V aktiviert
- .NET 8

## v1.3.0

### Highlights

- Dark/Light Theme vollständig integriert und live umschaltbar.
- VM-Backup-Workflow erweitert: Export und Import direkt in der App.
- Snapshot-/Checkpoint-Handling deutlich robuster gemacht (inkl. Sonderzeichen-Fixes).
- Config- und Info-Bereiche visuell bereinigt und klarer strukturiert.
- Notification/Log-Bereich überarbeitet: dynamische Größe und direkter Zugriff auf Logdatei.

### Neu

- Theme-Umschaltung (`Dark` / `Light`) im Config-Bereich mit Live-Anwendung.
- VM Export/Import in der UI integriert.
- Config-Tab: Export arbeitet auf der aktuell ausgewählten VM; Default-VM wird separat gesetzt.
- Notification-Bereich: neuer Button `Logdatei öffnen`.

### Verbessert

- Config-UX überarbeitet (klarere Trennung von VM-Auswahl, Default-VM und Export-Flow).
- Network-Tab aufgeräumt (doppelte/irritierende Statuszeilen entfernt).
- Info-Tab „Links“-Bereich cleaner dargestellt.
- Notification-Bereich verhält sich beim Ein-/Ausklappen kontrollierter.
- Snapshot-Bezeichnungen konsistenter (`Restore` statt `Apply` in der UI).

### Behoben

- Snapshot-Create-Button blieb in bestimmten Zuständen fälschlich deaktiviert.
- Snapshot-Sektion konnte durch globales Busy-State blockiert werden; Checkpoint-Laden entkoppelt.
- Checkpoint-Erstellung robuster bei Production-Checkpoint-Problemen (Fallback-Handling).
- Checkpoint Restore/Delete mit Sonderzeichen im Namen funktioniert zuverlässig.
- Mehrere kleinere UI-Layout-Probleme (u. a. horizontale Scroll-Irritationen im Config-Bereich).

### Kompatibilität

- Windows 10/11
- Hyper-V aktiviert
- .NET 8

---

## v1.2.0

### Highlights

- Stabilitätsupdate für Startverhalten und Bindings
- Verbesserte Tray-Funktionen für VM-Alltag
- Update-/Installer-Flow für einfachere Aktualisierung aus der App
- Überarbeitete Dokumentation und Build-Prozess

### Neu

- Tasktray: neuer Menüpunkt "Konsole öffnen" pro VM
- Tasktray: "VM starten" öffnet direkt danach automatisch die VM-Konsole
- Tasktray: Menü schließt sich nach Aktionen zuverlässig
- Tasktray: Menüpunkt heißt jetzt "Aktualisieren" und lädt die Konfiguration neu
- Config: `ui.trayVmNames` erlaubt die Auswahl, welche VMs im Tray angezeigt werden
- Config/UI: neue Option `ui.startMinimized` inkl. Checkbox in der App
- Update: Default-Repo auf `koerby/HyperTool` umgestellt
- Update: semantischer Versionsvergleich (inkl. `v`-Prefix und Prerelease)
- Update: Installer-Asset-Erkennung in GitHub Releases (`.exe`/`.msi`)
- Info-Tab: neuer Button "Update installieren" (Download + Start des Installers)
- Build: neuer Installer-Workflow über `build-installer-host.bat` und Inno Setup Script (`installer/HyperTool.iss`)

### Verbessert

- Release-Prozess erweitert: `build-host.bat installer version=x.y.z` erstellt App + Setup
- README um Installer/Update-Prozess und neue Config-Felder ergänzt

### Behoben

- Tray-Usability verbessert: Menü bleibt nicht mehr offen nach VM-Aktion

### Kompatibilität

- Windows 10/11
- Hyper-V aktiviert
- .NET 8

### Hinweis zum Update

Für zukünftige Releases empfiehlt sich ein GitHub-Release mit angehängtem Installer-Asset (`HyperTool-Setup-<version>.exe`), damit die In-App-Funktion "Update installieren" automatisch genutzt werden kann.
