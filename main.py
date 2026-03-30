import json
import os
import py_compile
import shutil
import subprocess
import tempfile
import threading
import time
import urllib.request
from datetime import datetime

import schedule
import yaml

from modules import analyzer, job_manager, queue_manager, telegram_bot, watcher

CONFIG_PFAD = "C:\\dcp_automatisierung\\config.yaml"
VERSION_PFAD = "C:\\dcp_automatisierung\\version.txt"
STAGING_PFAD = "C:\\dcp_automatisierung\\staging"
PENDING_UPDATE_PFAD = "C:\\dcp_automatisierung\\pending_update.json"
UPDATE_RESULT_PFAD = "C:\\dcp_automatisierung\\update_result.json"
UPDATER_PFAD = "C:\\dcp_automatisierung\\updater.py"

# ──────────────────────────────────────────────
# Hilfsfunktionen
# ──────────────────────────────────────────────

def lade_config():
    with open(CONFIG_PFAD, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def speichere_config(config):
    tmp = CONFIG_PFAD + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False)
    os.replace(tmp, CONFIG_PFAD)

def lese_version():
    try:
        return open(VERSION_PFAD, encoding="utf-8").read().strip()
    except Exception:
        return "unbekannt"

# ──────────────────────────────────────────────
# Ordner-Scan → Queue
# ──────────────────────────────────────────────

def _scan_ordner(ordner, dauer_sek):
    bilder = watcher.suche_neue_bilder(ordner)
    neu = 0
    for p in bilder:
        if queue_manager.hinzufuegen(p, dauer_sek):
            neu += 1
    return neu

def starte_verarbeitung():
    """Scannt Eingangsordner und legt neue Bilder in die Benennungs-Queue."""
    cfg = lade_config().get("ordner", {})
    gesamt = 0
    gesamt += _scan_ordner(cfg.get("eingang_7sec", ""), 7)
    gesamt += _scan_ordner(cfg.get("eingang_10sec", ""), 10)
    gesamt += _scan_ordner(cfg.get("eingang_15sec", ""), 15)
    if gesamt > 0:
        telegram_bot.sende_nachricht(f"{gesamt} neues Bild(er) in Queue aufgenommen.")

# ──────────────────────────────────────────────
# Queue-Worker: sequentieller Benennungs-Dialog
# ──────────────────────────────────────────────

def queue_worker():
    """Laeuft dauerhaft in eigenem Thread.
    Verarbeitet Queue-Items SEQUENTIELL (genau 1 Dialog gleichzeitig).
    Hintergrund-Jobs laufen parallel."""
    time.sleep(5)
    while True:
        item = queue_manager.naechstes_pending()
        if not item:
            time.sleep(10)
            continue

        if not telegram_bot.starte_dialog():
            queue_manager.zuruecksetzen(item["id"])
            time.sleep(5)
            continue

        try:
            bildpfad = item["bildpfad"]
            dauer_sek = item["dauer_sek"]
            pending_rest = queue_manager.pending_anzahl()

            caption = f"Neues Bild ({dauer_sek}s Werbung)"
            if pending_rest > 0:
                caption += f"  |  Noch {pending_rest} in Queue"
            try:
                telegram_bot.sende_bild(bildpfad, caption=caption)
            except Exception:
                telegram_bot.sende_nachricht(f"Bild: {os.path.basename(bildpfad)}")

            ocr_text = ""
            try:
                ocr_text = analyzer.lese_text_aus_bild(bildpfad)
            except Exception:
                pass

            prompt = (
                f"Erkannter Text:\n{ocr_text[:400]}\n\n"
                f"Bitte DCP-Namen eingeben.\n"
                f"(oder /skip zum Ueberspringen, Timeout: 60 Min)"
            )
            telegram_bot.sende_nachricht(prompt)

            antwort = telegram_bot.warte_auf_dialog_antwort(timeout=3600)

            if antwort is None or antwort.strip().lower() == "/skip":
                queue_manager.abschliessen(item["id"])
                telegram_bot.sende_nachricht(
                    f"Uebersprungen: {os.path.basename(bildpfad)}"
                )
            else:
                final_name = antwort.strip()
                queue_manager.abschliessen(item["id"])
                job_id = job_manager.erstelle_job(bildpfad, final_name)
                telegram_bot.sende_nachricht(
                    f"Name: {final_name}\nVerarbeitung laeuft im Hintergrund..."
                )
                threading.Thread(
                    target=verarbeite_job,
                    args=(job_id, 1),
                    daemon=True
                ).start()

        except Exception as e:
            telegram_bot.sende_nachricht(f"Fehler in Queue-Worker: {e}")
            queue_manager.zuruecksetzen(item["id"])
        finally:
            telegram_bot.beende_dialog()
            time.sleep(1)

# ──────────────────────────────────────────────
# Job-Pipeline: DCP → Upload → Ingest → Monitoring
# ──────────────────────────────────────────────

def _phase_ausfuehren(job_id, phase, fn):
    """Fuehrt eine Phase aus. Gibt True bei Erfolg, False bei Fehler zurueck."""
    try:
        job_manager.aktualisiere_phase(job_id, phase, "running")
        fn(job_id)
        job_manager.aktualisiere_phase(job_id, phase, "done")
        return True
    except Exception as e:
        job_manager.markiere_fehler(job_id, phase, str(e), retryable=True)
        return False

def verarbeite_job(job_id, ab_phase=1):
    """Fuehrt Job-Pipeline ab der angegebenen Phase aus."""
    if ab_phase <= 1:
        if not _phase_ausfuehren(job_id, 1, _dcp_erstellen):
            return
    if ab_phase <= 2:
        if not _phase_ausfuehren(job_id, 2, _upload_durchfuehren):
            return
    if ab_phase <= 3:
        if not _phase_ausfuehren(job_id, 3, _ingest_starten):
            return
    if ab_phase <= 4:
        if not _phase_ausfuehren(job_id, 4, _monitoring_ueberwachen):
            return
    job_manager.markiere_fertig(job_id)

def _dcp_erstellen(job_id):
    raise NotImplementedError("DCP-Erstellung: noch nicht implementiert")

def _upload_durchfuehren(job_id):
    raise NotImplementedError("FTP-Upload: noch nicht implementiert")

def _ingest_starten(job_id):
    raise NotImplementedError("Doremi-Ingest: noch nicht implementiert")

def _monitoring_ueberwachen(job_id):
    raise NotImplementedError("Monitoring: noch nicht implementiert")

# ──────────────────────────────────────────────
# In-App Update System (v2.4)
# ──────────────────────────────────────────────

def pruefe_update_ergebnis():
    """Beim Start: Ergebnis des letzten In-App-Updates melden."""
    if not os.path.exists(UPDATE_RESULT_PFAD):
        return
    try:
        with open(UPDATE_RESULT_PFAD, "r", encoding="utf-8") as f:
            result = json.load(f)
        os.remove(UPDATE_RESULT_PFAD)
        if result.get("erfolg"):
            telegram_bot.sende_nachricht(
                f"Update auf v{result.get('neue_version', '?')} erfolgreich installiert!"
            )
        else:
            telegram_bot.sende_nachricht(
                f"Update fehlgeschlagen: {result.get('fehler', 'Unbekannter Fehler')}\n"
                f"Rollback wurde durchgefuehrt."
            )
    except Exception:
        pass

def pruefe_update():
    """Prueft auf neue Version, laedt Dateien ins Staging, startet updater.py."""
    try:
        cfg = lade_config()
        update_cfg = cfg.get("update", {})
        version_url = update_cfg.get("github_version_url", "")
        manifest_url = update_cfg.get("github_manifest_url", "")
        base_url = update_cfg.get("github_base_url", "")

        if not version_url or not manifest_url or not base_url:
            telegram_bot.sende_nachricht("Update-URLs nicht konfiguriert.")
            return

        github_version = urllib.request.urlopen(version_url, timeout=10).read().decode().strip()
        local_version = lese_version()

        if github_version == local_version:
            telegram_bot.sende_nachricht(f"Bereits aktuell: v{local_version}")
            return

        telegram_bot.sende_nachricht(
            f"Update verfuegbar: v{local_version} -> v{github_version}\n"
            f"Lade Dateien herunter..."
        )

        # Manifest laden
        manifest_data = json.loads(
            urllib.request.urlopen(manifest_url, timeout=10).read().decode()
        )
        dateien = manifest_data.get("files", [])

        # Staging-Verzeichnis vorbereiten
        if os.path.exists(STAGING_PFAD):
            shutil.rmtree(STAGING_PFAD)
        os.makedirs(os.path.join(STAGING_PFAD, "modules"), exist_ok=True)

        # Alle Dateien herunterladen und Syntax pruefen
        for entry in dateien:
            src = entry["src"]
            url = base_url + src
            local_staged = os.path.join(STAGING_PFAD, src.replace("/", os.sep))
            urllib.request.urlretrieve(url, local_staged)
            if local_staged.endswith(".py"):
                py_compile.compile(local_staged, doraise=True)

        # pending_update.json schreiben
        pending = {
            "neue_version": github_version,
            "staging_pfad": STAGING_PFAD,
            "dateien": dateien,
            "timestamp": datetime.now().isoformat()
        }
        tmp = PENDING_UPDATE_PFAD + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(pending, f, ensure_ascii=False, indent=2)
        os.replace(tmp, PENDING_UPDATE_PFAD)

        telegram_bot.sende_nachricht(
            f"Alle Dateien geladen und validiert.\n"
            f"Update wird jetzt installiert (Service-Neustart)..."
        )

        # updater.py als separaten Prozess starten (ausserhalb des Dienstes)
        import sys
        subprocess.Popen(
            [sys.executable, UPDATER_PFAD],
            creationflags=subprocess.CREATE_NEW_CONSOLE
        )
        # updater.py stoppt den Service (und damit diesen Prozess)

    except Exception as e:
        telegram_bot.sende_nachricht(f"Update fehlgeschlagen: {e}")
        try:
            if os.path.exists(STAGING_PFAD):
                shutil.rmtree(STAGING_PFAD)
        except Exception:
            pass

# ──────────────────────────────────────────────
# Intervall-Einstellungen
# ──────────────────────────────────────────────

def aendere_intervall(minuten_str):
    try:
        minuten = int(minuten_str.strip())
        if not (1 <= minuten <= 1440):
            telegram_bot.sende_nachricht("Intervall muss zwischen 1 und 1440 Minuten liegen.")
            return
        cfg = lade_config()
        cfg.setdefault("zeitplan", {})["intervall_minuten"] = minuten
        speichere_config(cfg)
        schedule.clear("scan")
        schedule.every(minuten).minutes.do(starte_verarbeitung).tag("scan")
        telegram_bot.sende_nachricht(f"Check-Intervall: {minuten} Minuten (gespeichert).")
    except ValueError:
        telegram_bot.sende_nachricht("Ungueltige Eingabe. Beispiel: /intervall 60")

def aendere_update_intervall(stunden_str):
    try:
        stunden = int(stunden_str.strip())
        if not (0 <= stunden <= 168):
            telegram_bot.sende_nachricht("Intervall muss zwischen 0 und 168 Stunden liegen (0 = deaktiviert).")
            return
        cfg = lade_config()
        cfg.setdefault("update", {})["auto_update_intervall_stunden"] = stunden
        speichere_config(cfg)
        schedule.clear("update")
        if stunden > 0:
            schedule.every(stunden).hours.do(
                lambda: threading.Thread(target=pruefe_update, daemon=True).start()
            ).tag("update")
            telegram_bot.sende_nachricht(f"Auto-Update-Intervall: {stunden} Stunden (gespeichert).")
        else:
            telegram_bot.sende_nachricht("Auto-Update deaktiviert.")
    except ValueError:
        telegram_bot.sende_nachricht("Ungueltige Eingabe. Beispiel: /update_intervall 24")

# ──────────────────────────────────────────────
# Fehler-Menue
# ──────────────────────────────────────────────

def zeige_fehler():
    fehler = job_manager.hole_fehler()
    if not fehler:
        telegram_bot.sende_nachricht("Keine Fehler vorhanden.")
        return
    phasen = {1: "DCP", 2: "Upload", 3: "Ingest", 4: "Monitoring"}
    msg = f"{len(fehler)} Fehler-Job(s):\n"
    for j in fehler:
        phase = phasen.get(j.get("error_phase"), "?")
        name = (j.get("final_name") or os.path.basename(j.get("bildpfad", "")))[:35]
        fehlertext = (j.get("fehler_text") or "")[:100]
        msg += f"\nID: {j['id']}\nName: {name}\nPhase: {phase}\nFehler: {fehlertext}\n"
    msg += "\n/retry_alle  oder  /retry <ID>"
    telegram_bot.sende_nachricht(msg)

# ──────────────────────────────────────────────
# Befehlshandler
# ──────────────────────────────────────────────

def bearbeite_befehl(text):
    cmd = text.strip()
    low = cmd.lower()

    if low == "/version":
        telegram_bot.sende_nachricht(f"DCP-Automatisierung v{lese_version()}")

    elif low == "/update":
        telegram_bot.sende_nachricht("Suche nach Updates...")
        threading.Thread(target=pruefe_update, daemon=True).start()

    elif low == "/status":
        fehler = job_manager.hole_fehler()
        aktive = job_manager.hole_aktive()
        pending = queue_manager.pending_anzahl()
        cfg = lade_config()
        intervall = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
        update_intervall = cfg.get("update", {}).get("auto_update_intervall_stunden", 24)
        update_str = f"alle {update_intervall}h" if update_intervall > 0 else "deaktiviert"
        telegram_bot.sende_nachricht(
            f"System: Aktiv  |  v{lese_version()}\n"
            f"Queue: {pending} wartend\n"
            f"Jobs: {len(aktive)} laufend, {len(fehler)} Fehler\n"
            f"Check-Intervall: {intervall} Min\n"
            f"Auto-Update: {update_str}"
        )

    elif low == "/check":
        telegram_bot.sende_nachricht("Starte manuelle Ordner-Pruefung...")
        threading.Thread(target=starte_verarbeitung, daemon=True).start()

    elif low.startswith("/intervall"):
        teile = cmd.split(maxsplit=1)
        if len(teile) < 2:
            cfg = lade_config()
            aktuell = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
            telegram_bot.sende_nachricht(
                f"Aktuelles Intervall: {aktuell} Minuten\nAendern: /intervall <Minuten>"
            )
        else:
            aendere_intervall(teile[1])

    elif low.startswith("/update_intervall"):
        teile = cmd.split(maxsplit=1)
        if len(teile) < 2:
            cfg = lade_config()
            aktuell = cfg.get("update", {}).get("auto_update_intervall_stunden", 24)
            telegram_bot.sende_nachricht(
                f"Auto-Update-Intervall: {aktuell} Stunden (0 = deaktiviert)\n"
                f"Aendern: /update_intervall <Stunden>"
            )
        else:
            aendere_update_intervall(teile[1])

    elif low == "/fehler":
        zeige_fehler()

    elif low == "/retry_alle":
        jobs = job_manager.alle_retry()
        if not jobs:
            telegram_bot.sende_nachricht("Keine Fehler-Jobs zum Wiederholen.")
        else:
            telegram_bot.sende_nachricht(f"{len(jobs)} Job(s) werden neu gestartet...")
            for j in jobs:
                threading.Thread(
                    target=verarbeite_job, args=(j["id"], j["current_phase"]), daemon=True
                ).start()

    elif low.startswith("/retry "):
        job_id = cmd.split(maxsplit=1)[1].strip()
        j = job_manager.retry_job(job_id)
        if j:
            telegram_bot.sende_nachricht(
                f"Job {job_id} neu gestartet ab Phase {j['current_phase']}..."
            )
            threading.Thread(
                target=verarbeite_job, args=(j["id"], j["current_phase"]), daemon=True
            ).start()
        else:
            telegram_bot.sende_nachricht(f"Job '{job_id}' nicht gefunden oder nicht wiederholbar.")

    elif low == "/hilfe":
        telegram_bot.sende_nachricht(
            "Befehle:\n"
            "/version               Aktuelle Version\n"
            "/update                Update von GitHub\n"
            "/update_intervall <h>  Auto-Update-Intervall (0=aus)\n"
            "/check                 Ordner jetzt pruefen\n"
            "/status                Systemstatus\n"
            "/intervall <n>         Check-Intervall (Minuten)\n"
            "/fehler                Fehler-Jobs anzeigen\n"
            "/retry <ID>            Job neu starten\n"
            "/retry_alle            Alle Fehler-Jobs neu\n"
            "/hilfe                 Diese Hilfe"
        )

    else:
        telegram_bot.sende_nachricht(f"Unbekannt: {text}\n/hilfe fuer alle Befehle.")

# ──────────────────────────────────────────────
# Einstiegspunkt
# ──────────────────────────────────────────────

if __name__ == "__main__":
    # Absturz-Recovery: naming → pending
    queue_manager.naming_zuruecksetzen()

    cfg = lade_config()
    intervall = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
    update_intervall = cfg.get("update", {}).get("auto_update_intervall_stunden", 24)

    # Status-Callback fuer job_manager setzen
    job_manager.setze_status_callback(telegram_bot.sende_nachricht)

    # Update-Ergebnis vom letzten Neustart melden
    pruefe_update_ergebnis()

    telegram_bot.sende_nachricht(f"DCP-Automatisierung v{lese_version()} gestartet.")

    # Telegram-Listener
    threading.Thread(
        target=telegram_bot.starte_listener,
        args=(bearbeite_befehl,),
        daemon=True
    ).start()

    # Queue-Worker
    threading.Thread(target=queue_worker, daemon=True).start()

    # Scheduler
    schedule.every(intervall).minutes.do(starte_verarbeitung).tag("scan")
    schedule.every(5).minutes.do(job_manager.sende_bundle_wenn_noetig).tag("bundle")
    if update_intervall > 0:
        schedule.every(update_intervall).hours.do(
            lambda: threading.Thread(target=pruefe_update, daemon=True).start()
        ).tag("update")

    starte_verarbeitung()

    while True:
        schedule.run_pending()
        time.sleep(10)
