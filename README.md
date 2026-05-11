# LB-Automatisierung

Automatisierungssystem für das Lichtblick-Kinoprogramm auf Windows.  
Überwacht einen Eingangsordner auf neue JPG-Bilder, erstellt daraus DCPs via DCP-o-matic, lädt sie per FTP auf den Doremi DCP2000 hoch und triggert den Ingest via HTTP. Steuerung und Statusmeldungen über Telegram.

---

## Funktionsweise

```
JPG-Bild erkannt
      ↓
Naming-Dialog (Telegram) → Name bestätigen
      ↓
Phase 1: DCP erstellen (dcpomatic2_create + dcpomatic2_cli)
      ↓
Phase 2: FTP-Upload → Doremi /gui/{dcp_name}/
      ↓
Phase 3: Ingest starten (HTTP → Doremi Web-Interface)
      ↓
Phase 4: Monitoring (HTTP Monitor-Seite pollen) → Erfolg/Fehler per Telegram
```

---

## Dateien & Ordner

| Datei | Bedeutung |
|---|---|
| `main.py` | Hauptprogramm – Queue, Naming-Dialog, Job-Pipeline, Telegram-Befehle |
| `watchdog.py` | Überwacht ob main.py läuft, meldet Absturz per Telegram |
| `updater.py` | Selbst-Update-Logik (läuft via Windows-Taskplaner) |
| `modules/doremi_web.py` | HTTP-API zum Doremi (Login, Ingest-Trigger, Monitor-Polling) |
| `modules/job_manager.py` | Job-Persistenz, Phasen-Tracking, Batch-Fertig-Meldungen |
| `modules/telegram_bot.py` | Nachrichten senden/empfangen, Dialog-State, Listener-Thread |
| `modules/analyzer.py` | OCR (Tesseract) zur Typ-Erkennung aus Bildern |
| `modules/queue_manager.py` | Warteschlange für neue Bilder |
| `modules/watcher.py` | Ordner-Scan für neue JPG-Dateien |
| `version.txt` | Aktuelle Versionsnummer |
| `update_manifest.json` | Dateiliste + Version für das Auto-Update-System |

Auf dem Server (nicht im Repo): `C:\dcp_automatisierung\config.yaml`

---

## Voraussetzungen

- Windows 10/11
- Python 3.11+
- [DCP-o-matic](https://dcpomatic.com/) installiert
- [NSSM](https://nssm.cc/) für Service-Verwaltung
- Tesseract OCR
- Doremi DCP2000 im Netzwerk erreichbar

---

## Installation

```powershell
# 1. Repository klonen
git clone https://github.com/Sh-Kod/LB-automatisierung.git C:\dcp_automatisierung

# 2. Abhängigkeiten installieren
pip install -r requirements.txt

# 3. Config anlegen (Vorlage anpassen)
# C:\dcp_automatisierung\config.yaml

# 4. Hauptservice registrieren
nssm install dcp_automatisierung "C:\Python\bin\python.exe" "C:\dcp_automatisierung\main.py"
nssm start dcp_automatisierung

# 5. Watchdog-Service registrieren
nssm install dcp_watchdog "C:\Python\bin\python.exe" "C:\dcp_automatisierung\watchdog.py"
nssm set dcp_watchdog AppRestartDelay 5000
nssm start dcp_watchdog
```

---

## config.yaml – Struktur

```yaml
telegram:
  token: "BOT_TOKEN"
  chat_id: "CHAT_ID"

doremi:
  ip: "172.20.23.11"
  ftp_user: "ingest"
  ftp_pass: "ingest"
  web_user: "admin"
  web_pass: "1234"

zeitplan:
  intervall_minuten: 60

update:
  auto_update_intervall_stunden: 24

logging:
  log_datei: "C:\\dcp_automatisierung\\logs\\dcp_automatisierung.log"
  log_level: "INFO"
```

---

## Telegram-Befehle

| Befehl | Funktion |
|---|---|
| `/hilfe` | Alle Befehle anzeigen |
| `/status` | Aktuellen Status abfragen |
| `/jobs` | Fehlerhafte Jobs anzeigen |
| `/retry_alle` | Alle fehlerhaften Jobs neu starten |
| `/retry <id>` | Einzelnen Job neu starten |
| `/check` | Ordner-Prüfung manuell starten |
| `/pause` | Scan pausieren/fortsetzen |
| `/neustart` | Service neu starten |
| `/update` | Update suchen und installieren |
| `/intervall <n>` | Scan-Intervall in Minuten setzen |

---

## Watchdog

`watchdog.py` läuft als eigener NSSM-Service und überwacht `main.py` unabhängig:

- `main.py` schreibt alle **30 Sekunden** einen Timestamp in `heartbeat.txt` (1 Zeile, wird überschrieben)
- Watchdog prüft alle **60 Sekunden** ob der Timestamp frisch ist
- Ist der Heartbeat älter als **2 Minuten** → Telegram: *„DCP-Automatisierung nicht erreichbar!"*
- Beim sauberen Beenden von `main.py` → Telegram: *„DCP-Automatisierung wurde beendet."*
- Beim Neustart → Telegram: *„DCP-Automatisierung vX.XX gestartet."*

---

## Auto-Update

Updates werden über Telegram ausgelöst:

1. `/update` senden
2. `updater.py` wird via Windows-Taskplaner gestartet
3. Lädt neue Dateien von GitHub herunter (laut `update_manifest.json`)
4. Startet den Service neu
5. Bestätigung per Telegram

---

## Versionsverlauf

| Version | Änderung |
|---|---|
| v2.49 | Watchdog-Heartbeat – Telegram-Benachrichtigung bei Absturz/Beenden |
| v2.48 | Anti-Spam Aggregator – gebündelte Ingest-Nachrichten |
| v2.47 | HTTP-Monitoring über Doremi Monitor-Seite, JPG-Dateiname als Namensvorschlag |
