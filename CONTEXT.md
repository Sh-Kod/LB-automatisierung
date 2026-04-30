# CONTEXT.md

## Erledigt

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
- **Praxis-Test**: 7 DCPs parallel ingestet in 1 Minute — vollständig erfolgreich
- **CLAUDE.md und CONTEXT.md** erstellt

---

## Geänderte Dateien

- `main.py` – v2.47 + v2.48: Aggregator, JPG-Vorschlag, HTTP-Monitoring, Job-ID via Monitor-Seite
- `modules/doremi_web.py` – `pruefe_ingest_status()` hinzugefügt, `import re` ergänzt
- `version.txt` – `2.48`
- `update_manifest.json` – `version: 2.48`
- `CLAUDE.md` – neu erstellt
- `CONTEXT.md` – neu erstellt

---

## Offene Probleme

- Bei `/update` um 16:05 zeigte das System noch v2.45 (Timing – v2.47 war noch nicht gepusht). Jetzt v2.48 auf GitHub. User muss `/update` erneut ausführen.
- `_sammle("ingest_start", ...)` wird kurz VOR `doremi_web.login()` aufgerufen → bei Login-Fehler erscheint "Starte Ingest" trotzdem in Telegram, aber danach kommt die Fehlermeldung. Akzeptabel, da selten.
- Die "Fertig: {name}"-Meldung bei Einzel-Job kommt weiterhin aus `job_manager._flush_buffer()` (nicht aggregiert). Kein Problem, da Einzel-Jobs kein Spam verursachen.

---

## Nächster Schritt

- `/update` auf Telegram ausführen → v2.48 installieren
- Test: mehrere Bilder einlegen (verschiedene Namen), Naming-Dialog ausprobieren (JPG-Name als Vorschlag `[1]`)
- Test: paralleler Ingest mit mehreren DCPs → prüfen ob Aggregator korrekt 4 gebündelte Nachrichten sendet
- Falls gewünscht: Aggregator auch für Upload-Phase-Meldungen ("✓ Hochgeladen") erweitern
