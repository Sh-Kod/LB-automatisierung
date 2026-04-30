# CLAUDE.md – LB-Automatisierung

## Projektbeschreibung

Automatisierungssystem für Lichtblick-Kinoprogramm (LB) auf Windows.
Überwacht Eingangsordner auf neue JPG-Bilder, erstellt DCPs via DCP-o-matic,
lädt sie per FTP auf den Doremi DCP2000 hoch und triggert den Ingest via HTTP.
Steuerung und Statusmeldungen über Telegram.

---

## Wichtige Dateien und Ordner

| Pfad | Bedeutung |
|---|---|
| `main.py` | Hauptprogramm – Queue, Naming-Dialog, Job-Pipeline, Befehle |
| `modules/doremi_web.py` | HTTP-API zum Doremi (Login, Ingest-Trigger, Monitor-Polling) |
| `modules/job_manager.py` | Job-Persistenz, Phasen-Tracking, Batch-Fertig-Meldungen |
| `modules/telegram_bot.py` | Nachrichten senden/empfangen, Dialog-State, Listener-Thread |
| `modules/analyzer.py` | OCR (Tesseract) zur Typ-Erkennung aus Bildern |
| `modules/queue_manager.py` | Warteschlange für neue Bilder |
| `modules/watcher.py` | Ordner-Scan für neue JPG-Dateien |
| `version.txt` | Aktuelle Versionsnummer (z.B. `2.48`) |
| `update_manifest.json` | Dateiliste + Version für das Auto-Update-System |
| `updater.py` | Selbst-Update-Logik (wird via Taskplaner ausgeführt) |

Auf dem Windows-Server: `C:\dcp_automatisierung\config.yaml` (nicht im Repo)

---

## Build- und Test-Befehle

Kein Build-System. Direkt starten:
```
python main.py
```
Als NSSM-Service auf dem Windows-Server registriert.

Updates werden über `/update` per Telegram ausgelöst → `updater.py` läuft via Windows-Taskplaner.

---

## Architektur-Entscheidungen

### Job-Pipeline (4 Phasen)
1. **DCP erstellen** – dcpomatic2_create + dcpomatic2_cli
2. **FTP-Upload** – `ftplib` → Doremi `/gui/{dcp_name}/`
3. **Ingest starten** – HTTP `SCHEDULE_INGEST_TASK` via `doremi_web.login()` + `doremi_web.starte_ingest()`
4. **Monitoring** – HTTP Monitor-Seite pollen (`doremi_web.pruefe_ingest_status()`); Fallback: FTP-Ordner-Check

### KLV TCP (Port 11730) – AUFGEGEBEN
`IngestAddJob` gibt immer error_code=1. HTTP-Web-Interface ist der einzige funktionierende Ingest-Weg.

### Ingest-Monitor-Seite
`GET http://172.20.23.11/web/sys_control/index.php?page=ingest_manager/ingest_monitor.php`
HTML-Tabelle mit: `value="{job_id}"`, Icon (`success_16.png` / `error-icon-16.png`), `<label>{dcp_name}</label>`

### Anti-Spam Aggregator (`_agg_worker` / `_sammle`)
Parallele Jobs sammeln gleichartige Meldungen 3 Sekunden, dann eine gebündelte Telegram-Nachricht.
Während des Naming-Dialogs werden Ingest-Meldungen zurückgehalten.

### Naming-Dialog
1. JPG-Dateiname (bereinigt) als primärer Vorschlag `[1]`
2. OCR-Vorschlag (Tesseract + Typ-Erkennung) als sekundärer Vorschlag `[2]`
3. Eigener Name `[3]`, Überspringen `[4]`

### FTP-Pfade auf Doremi
- FTP `/gui/{dcp_name}/` = Dateisystem `/data/incoming/gui/{dcp_name}/`
- Doremi **löscht** den Ordner nach erfolgreichem Ingest (= Erfolgsindikator)

---

## Doremi-Konfiguration

| Eigenschaft | Wert |
|---|---|
| IP | `172.20.23.11` |
| FTP-User | `ingest` / `ingest` |
| Web-User | `admin` / `1234` |
| Firmware | 21.5b, Software 2.8.52-0 |

Config-Schlüssel in `config.yaml`: `doremi.ip`, `doremi.ftp_user`, `doremi.ftp_pass`, `doremi.web_user`, `doremi.web_pass`

---

## Code-Style und Vermeidungsregeln

- **Nie KLV** (`doremi_api`) für Ingest-Trigger verwenden – gibt immer error_code=1
- **Nie direkt** `telegram_bot.sende_nachricht()` aus parallelen Job-Threads → immer `_sammle()` für Ingest-Meldungen
- **Nie Code schreiben ohne Bestätigung** – User sagt explizit "einverstanden" / "mach das" bevor Änderungen gemacht werden
- **Kein Feature-Creep** – nur das implementieren was besprochen wurde
- Alle Phasen-Fehler gehen durch `_phase_ausfuehren()` → Error-Nachrichten automatisch
- Version in `version.txt` UND `update_manifest.json` synchron halten

---

## Typischer Workflow – Neues Feature

1. Feature besprechen, Plan vorstellen, auf Bestätigung warten
2. Relevante Dateien lesen (vorher + nachher prüfen)
3. Änderungen machen, direkt verifizieren
4. `version.txt` und `update_manifest.json` bump
5. Commit + Push auf Branch `claude/check-github-access-JYt09`
6. User führt `/update` auf Telegram aus → Auto-Update

## Aktueller Branch

`claude/check-github-access-JYt09` (immer hier entwickeln, nie auf `main` pushen)
