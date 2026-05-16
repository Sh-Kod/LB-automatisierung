# CONTEXT.md

## Erledigt

- **v2.53**: Updater stoppt/startet jetzt auch den Watchdog-Service
  - `updater.py`: `stoppe_alle()` / `starte_alle()` – beide Services (`dcp_automatisierung` + `dcp_watchdog`) werden gemeinsam verwaltet
  - Reihenfolge: Watchdog ZUERST stoppen (kein Falschalarm während main.py stoppt), main.py ZUERST starten (heartbeat sofort frisch bevor Watchdog prüft)
  - Behebt das Problem dass der Watchdog-Service mit altem Code im RAM weiterlief, obwohl `watchdog.py` aktualisiert wurde → neue Watchdog-Versionen werden jetzt automatisch aktiv
  - Behebt Falschalarm "Heartbeat-Datei fehlt" während `/update`, der vorher gesendet wurde weil der Watchdog während des Service-Stopps von main.py weiterlief

- **v2.52**: Watchdog-Falschmeldungen beim Start behoben
  - `watchdog.py`: 5-Minuten Grace-Period (`STARTVERZOEGERUNG = 300`) vor dem ersten Heartbeat-Check
  - Verhindert Falschalarm wenn Watchdog-Service neustartet und main.py noch nicht läuft



- **v2.51**: Fehlende "gestartet"-Meldung nach PC-Neustart behoben
  - `modules/telegram_bot.py`: `sende_nachricht()` gibt jetzt `True`/`False` zurück (rückwärtskompatibel)
  - `main.py`: `_sende_start_meldung()` – Retry-Thread sendet Start-Meldung alle 30s bis max. 10 Versuche (~5 Min), falls Netzwerk beim Boot noch nicht bereit
  - `watchdog.py`: Sendet "DCP-Automatisierung wieder erreichbar." wenn Heartbeat nach Ausfall zurückkommt (doppelter Boden)

- **v2.47**: HTTP-Monitoring über Doremi Monitor-Seite implementiert
  - `doremi_web.pruefe_ingest_status()` hinzugefügt: parst HTML-Tabelle, erkennt `success_16.png` / `error-icon-16.png` / `pending`
  - `_ingest_starten()`: Job-ID wird jetzt via HTTP Monitor-Seite ermittelt (4 Versuche à 2s), nicht mehr via KLV
  - `_monitoring_ueberwachen()`: HTTP Monitor-Polling alle 10s (primär), FTP-Ordner-Check als Fallback, PHPSESSID-Erneuerung bei Fehler
  - FTP-Fallback (job_id=0) bleibt erhalten
- **v2.47**: JPG-Dateiname als Namensvorschlag im Naming-Dialog
  - Dateiname (bereinigt) als [1], OCR-Vorschlag als [2], Eigener Name als [3], Überspringen als [4]
  - Funktioniert korrekt für alle 4 Fälle: beide/nur JPG/nur OCR/keiner
- **v2.48**: Anti-Spam Aggregator für parallele Jobs
  - `_sammle(key, item, formatter)` + `_agg_worker()` Thread
  - 6 Nachrichtentypen: `upload_retry`, `ingest_start`, `ingest_pending`, `ingest_fertig`, `ingest_warten`, `ingest_timeout`
  - 3-Sekunden-Sammelzeit → eine gebündelte Nachricht statt N Einzelnachrichten
  - Während Naming-Dialog: Puffer hält Ingest-Meldungen zurück
- **v2.49**: Watchdog-Heartbeat
  - `main.py` schreibt alle 30s einen Timestamp in `heartbeat.txt` (1 Zeile, wird überschrieben)
  - `atexit` + Signal-Handler: Bei sauberem Beenden → Telegram „DCP-Automatisierung wurde beendet."
  - Neues `watchdog.py` als eigener NSSM-Service: prüft alle 60s den Heartbeat, meldet bei Ausfall (>2 Min) per Telegram
  - Auf beiden Standorten (Düsseldorf + Recklinghausen) eingerichtet und getestet
- **Praxis-Test**: 7 DCPs parallel ingestet in 1 Minute — vollständig erfolgreich
- **README.md** erstellt mit Projektübersicht, Installation, Befehlen, Watchdog-Doku
- **CLAUDE.md** und **CONTEXT.md** erstellt und gepflegt

---

## Geänderte Dateien (v2.53)

- `updater.py` – `WATCHDOG_SERVICE` Konstante, generische `_stoppe()`/`_starte()`, neue Helfer `stoppe_alle()`/`starte_alle()`
- `version.txt` – `2.53`
- `update_manifest.json` – `version: 2.53`

---

## Geänderte Dateien (v2.52)

- `watchdog.py` – `STARTVERZOEGERUNG = 300` + initialer Sleep vor der Hauptschleife
- `version.txt` – `2.52`
- `update_manifest.json` – `version: 2.52`

---

## Geänderte Dateien (v2.51)

- `main.py` – `_sende_start_meldung()` hinzugefügt, Startup-Meldung als Retry-Thread
- `watchdog.py` – Recovery-Meldung "wieder erreichbar" ergänzt
- `modules/telegram_bot.py` – `sende_nachricht()` gibt `True`/`False` zurück
- `version.txt` – `2.51`
- `update_manifest.json` – `version: 2.51`

---

## Geänderte Dateien (v2.49)

- `main.py` – Heartbeat-Worker, atexit/Signal-Handler, `import atexit/signal`, `HEARTBEAT_PFAD`
- `watchdog.py` – neu erstellt
- `version.txt` – `2.49`
- `update_manifest.json` – `version: 2.49`, `watchdog.py` eingetragen
- `README.md` – neu erstellt
- `CLAUDE.md` – Session-Start-Pflicht (CONTEXT.md + README.md lesen) ergänzt
- `CONTEXT.md` – auf v2.49 aktualisiert

---

## Offene Probleme

- `_sammle("ingest_start", ...)` wird kurz VOR `doremi_web.login()` aufgerufen → bei Login-Fehler erscheint „Starte Ingest" trotzdem in Telegram, aber danach kommt die Fehlermeldung. Akzeptabel, da selten.
- Die „Fertig: {name}"-Meldung bei Einzel-Job kommt weiterhin aus `job_manager._flush_buffer()` (nicht aggregiert). Kein Problem, da Einzel-Jobs kein Spam verursachen.

---

## Nächster Schritt

- v2.53 deployen (`/update` via Telegram)
- **Wichtig (einmalig):** Da v2.53 selbst der Fix für den nicht-neugestarteten Watchdog ist,
  muss auf BEIDEN Standorten der Watchdog-Service einmalig manuell neu gestartet werden,
  damit v2.52/v2.53 Code aktiv wird:
  `nssm restart dcp_watchdog`
  Ab v2.53 erledigt der Updater das selbst.
