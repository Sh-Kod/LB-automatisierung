# CLAUDE.md – LB-Automatisierung

## WHAT

Python-Dienst (NSSM-Service, Windows) der automatisch DCPs für das
Lichtblick-Kino erstellt und einspielt:

- **Watcher** überwacht Eingangsordner auf neue JPG-Bilder
- **Naming-Dialog** per Telegram: Dateiname [1] → OCR/Tesseract [2] → Eigener [3] → Skip [4]
- **Job-Pipeline**: DCP erstellen (DCP-o-matic) → FTP-Upload → Ingest-Trigger → Monitoring
- **Watchdog** (eigener NSSM-Service): prüft alle 60 Sek. ob main.py noch läuft
  (via `heartbeat.txt`), meldet Ausfall + Wiederherstellung per Telegram
- **Health-Monitoring**: Ingest-Status per HTTP-Polling; Fallback: FTP-Ordner-Check
- **Auto-Update**: `/update` per Telegram → `updater.py` via Windows-Taskplaner

Stack: Python, ftplib, Tesseract-OCR, requests, python-telegram-bot, NSSM

Session-Start: `CONTEXT.md` lesen (Stand + nächste Schritte)

---

## WHY – Architekturentscheidungen

**KLV TCP (Port 11730) → AUFGEGEBEN**
`IngestAddJob` gibt immer error_code=1. Nie wieder versuchen.
Einziger funktionierender Weg: HTTP Web-Interface.

**Anti-Spam Aggregator statt direktem sende_nachricht()**
Parallele Job-Threads → nie direkt `telegram_bot.sende_nachricht()`.
Immer `_sammle()` → bündelt gleichartige Meldungen für 3 Sek.

**Ingest-Erfolg = FTP-Ordner verschwindet**
Doremi löscht `/gui/{dcp_name}/` nach erfolgreichem Ingest.
Das ist der primäre Erfolgsindikator.

---

## HOW

```bash
python main.py          # direkt starten
# Deploy: /update via Telegram → updater.py (Taskplaner)
```

Workflow: besprechen → Bestätigung abwarten → lesen → ändern →
`version.txt` + `update_manifest.json` bump → commit + push auf Feature-Branch →
**PR nach `main` erstellen** → User mergt → `/update` via Telegram

**WICHTIG:** Der Server überwacht nur `main`. Nach jedem Push immer sofort
einen PR nach `main` erstellen (`gh pr create --base main`).
Ohne PR → kein Update auf dem Server.

**Nie:** KLV, direktes `sende_nachricht()` aus Threads,
Code ohne explizite Bestätigung ("einverstanden" / "mach das"), Feature-Creep

---

## Sprache

Alle Antworten auf **Deutsch**. Technische Begriffe und Code-Bezeichner bleiben auf Englisch.
