import json
import os
import py_compile
import re
import shutil
import subprocess
import threading
import time
import urllib.request
from datetime import datetime

import schedule
import yaml

from modules import analyzer, job_manager, queue_manager, telegram_bot, watcher

CONFIG_PFAD       = "C:\\dcp_automatisierung\\config.yaml"
VERSION_PFAD      = "C:\\dcp_automatisierung\\version.txt"
RULES_PFAD        = "C:\\dcp_automatisierung\\rules\\naming_rules.yaml"
STAGING_PFAD      = "C:\\dcp_automatisierung\\staging"
PENDING_UPDATE_PFAD = "C:\\dcp_automatisierung\\pending_update.json"
UPDATE_RESULT_PFAD  = "C:\\dcp_automatisierung\\update_result.json"
UPDATER_PFAD      = "C:\\dcp_automatisierung\\updater.py"

_scan_pausiert      = False
_scan_pausiert_lock = threading.Lock()

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

def _trenn():
    return "─" * 30

# ──────────────────────────────────────────────
# Naming-Vorschlag aus Rules
# ──────────────────────────────────────────────

def schlage_namen_vor(ocr_text):
    """Versucht aus OCR-Text und naming_rules.yaml einen DCP-Namen vorzuschlagen."""
    try:
        with open(RULES_PFAD, "r", encoding="utf-8") as f:
            rules_data = yaml.safe_load(f)
        regeln = (rules_data or {}).get("regeln", [])
        for regel in regeln:
            muster  = regel.get("muster", "")
            vorlage = regel.get("vorlage", "")
            if muster and vorlage and re.search(muster, ocr_text, re.IGNORECASE):
                return vorlage.replace("{JAHR}", str(datetime.now().year))
    except Exception:
        pass
    return None

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
    with _scan_pausiert_lock:
        if _scan_pausiert:
            return
    cfg = lade_config().get("ordner", {})
    gesamt  = _scan_ordner(cfg.get("eingang_7sec",  ""), 7)
    gesamt += _scan_ordner(cfg.get("eingang_10sec", ""), 10)
    gesamt += _scan_ordner(cfg.get("eingang_15sec", ""), 15)
    if gesamt > 0:
        telegram_bot.sende_nachricht(f"{gesamt} neues Bild(er) in Queue aufgenommen.")

# ──────────────────────────────────────────────
# Queue-Worker: sequentieller Benennungs-Dialog
# ──────────────────────────────────────────────

def _dialog_zeitüberschreitung(item_id):
    queue_manager.zuruecksetzen(item_id)
    telegram_bot.sende_nachricht("Timeout – Bild zurück in Queue.")

def queue_worker():
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
            bildpfad    = item["bildpfad"]
            dauer_sek   = item["dauer_sek"]
            pending_rest = queue_manager.pending_anzahl()

            # Bild senden
            caption = f"Neues Bild ({dauer_sek}s)"
            if pending_rest > 0:
                caption += f"  |  Queue: {pending_rest} wartend"
            try:
                telegram_bot.sende_bild(bildpfad, caption=caption)
            except Exception:
                telegram_bot.sende_nachricht(f"Bild: {os.path.basename(bildpfad)}")

            # OCR
            ocr_text = ""
            try:
                ocr_text = analyzer.lese_text_aus_bild(bildpfad)
            except Exception:
                pass

            # Vorschlag aus Naming-Rules
            vorschlag = schlage_namen_vor(ocr_text)

            # Dialog aufbauen
            t = _trenn()
            msg = f"{t}\n"
            if ocr_text:
                msg += f"OCR-Text:\n{ocr_text[:300].strip()}\n\n"
            if vorschlag:
                msg += f"Vorschlag:\n{vorschlag}\n\n"
                msg += "[1]  Vorschlag übernehmen\n"
                msg += "[2]  Eigenen Namen eingeben\n"
                msg += "[3]  Überspringen\n"
            else:
                msg += "Kein Vorschlag (keine Regel gefunden)\n\n"
                msg += "[1]  Namen eingeben\n"
                msg += "[2]  Überspringen\n"
            msg += f"{t}\n(Timeout: 60 Min)"
            telegram_bot.sende_nachricht(msg)

            # Schritt 1: Auswahl abwarten
            antwort = telegram_bot.warte_auf_dialog_antwort(timeout=3600)
            if antwort is None:
                _dialog_zeitüberschreitung(item["id"])
                continue

            a = antwort.strip()
            final_name = None
            skip = False

            if vorschlag:
                if a == "1":
                    final_name = vorschlag
                elif a == "2":
                    telegram_bot.sende_nachricht("Bitte DCP-Namen eingeben:")
                    a2 = telegram_bot.warte_auf_dialog_antwort(timeout=3600)
                    if a2 is None:
                        _dialog_zeitüberschreitung(item["id"])
                        continue
                    final_name = a2.strip()
                elif a in ("3", "/skip"):
                    skip = True
                else:
                    # Direkte Namenseingabe akzeptieren
                    final_name = a
            else:
                if a == "1":
                    telegram_bot.sende_nachricht("Bitte DCP-Namen eingeben:")
                    a2 = telegram_bot.warte_auf_dialog_antwort(timeout=3600)
                    if a2 is None:
                        _dialog_zeitüberschreitung(item["id"])
                        continue
                    final_name = a2.strip()
                elif a in ("2", "/skip"):
                    skip = True
                else:
                    final_name = a

            if skip or (final_name or "").lower() == "/skip":
                queue_manager.abschliessen(item["id"])
                telegram_bot.sende_nachricht(
                    f"Übersprungen: {os.path.basename(bildpfad)}"
                )
            elif final_name:
                queue_manager.abschliessen(item["id"])
                job_id = job_manager.erstelle_job(bildpfad, final_name)
                telegram_bot.sende_nachricht(
                    f"Name gesetzt: {final_name}\nVerarbeitung läuft..."
                )
                threading.Thread(
                    target=verarbeite_job, args=(job_id, 1), daemon=True
                ).start()

        except Exception as e:
            telegram_bot.sende_nachricht(f"Fehler im Dialog: {e}")
            queue_manager.zuruecksetzen(item["id"])
        finally:
            telegram_bot.beende_dialog()
            time.sleep(1)

# ──────────────────────────────────────────────
# Job-Pipeline
# ──────────────────────────────────────────────

def _phase_ausfuehren(job_id, phase, fn):
    try:
        job_manager.aktualisiere_phase(job_id, phase, "running")
        fn(job_id)
        job_manager.aktualisiere_phase(job_id, phase, "done")
        return True
    except Exception as e:
        job_manager.markiere_fehler(job_id, phase, str(e), retryable=True)
        return False

def verarbeite_job(job_id, ab_phase=1):
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
# In-App Update System
# ──────────────────────────────────────────────

def pruefe_update_ergebnis():
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
                f"Rollback wurde durchgeführt."
            )
    except Exception:
        pass

def pruefe_update():
    try:
        cfg        = lade_config()
        update_cfg = cfg.get("update", {})
        version_url  = update_cfg.get("github_version_url", "")
        manifest_url = update_cfg.get("github_manifest_url", "")
        base_url     = update_cfg.get("github_base_url", "")

        if not version_url or not manifest_url or not base_url:
            telegram_bot.sende_nachricht("Update-URLs nicht konfiguriert.")
            return

        github_version = urllib.request.urlopen(version_url, timeout=10).read().decode().strip()
        local_version  = lese_version()

        if github_version == local_version:
            telegram_bot.sende_nachricht(f"Bereits aktuell: v{local_version}")
            return

        telegram_bot.sende_nachricht(
            f"Update verfügbar: v{local_version} → v{github_version}\n"
            f"Lade Dateien herunter..."
        )

        manifest_data = json.loads(
            urllib.request.urlopen(manifest_url, timeout=10).read().decode()
        )
        dateien = manifest_data.get("files", [])

        if os.path.exists(STAGING_PFAD):
            shutil.rmtree(STAGING_PFAD)
        os.makedirs(os.path.join(STAGING_PFAD, "modules"), exist_ok=True)

        for entry in dateien:
            src = entry["src"]
            url = base_url + src
            local_staged = os.path.join(STAGING_PFAD, src.replace("/", os.sep))
            urllib.request.urlretrieve(url, local_staged)
            if local_staged.endswith(".py"):
                py_compile.compile(local_staged, doraise=True)

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
            f"Dateien geladen und validiert.\n"
            f"Update wird jetzt installiert (Service-Neustart)..."
        )

        import sys
        subprocess.Popen(
            [sys.executable, UPDATER_PFAD],
            creationflags=subprocess.CREATE_NEW_CONSOLE
        )

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
        telegram_bot.sende_nachricht(f"Scan-Intervall: {minuten} Minuten gespeichert.")
    except ValueError:
        telegram_bot.sende_nachricht("Ungültige Eingabe. Beispiel: /intervall 60")

def aendere_update_intervall(stunden_str):
    try:
        stunden = int(stunden_str.strip())
        if not (0 <= stunden <= 168):
            telegram_bot.sende_nachricht("Intervall muss zwischen 0 und 168 Stunden liegen.")
            return
        cfg = lade_config()
        cfg.setdefault("update", {})["auto_update_intervall_stunden"] = stunden
        speichere_config(cfg)
        schedule.clear("update")
        if stunden > 0:
            schedule.every(stunden).hours.do(
                lambda: threading.Thread(target=pruefe_update, daemon=True).start()
            ).tag("update")
            telegram_bot.sende_nachricht(f"Auto-Update: alle {stunden} Stunden gespeichert.")
        else:
            telegram_bot.sende_nachricht("Auto-Update deaktiviert.")
    except ValueError:
        telegram_bot.sende_nachricht("Ungültige Eingabe. Beispiel: /update_intervall 24")

# ──────────────────────────────────────────────
# Wartung
# ──────────────────────────────────────────────

def zeige_jobs():
    fehler = job_manager.hole_fehler()
    aktive = job_manager.hole_aktive()
    phasen = {1: "DCP", 2: "Upload", 3: "Ingest", 4: "Monitoring"}
    t   = _trenn()
    msg = f"{t}\nJobs & Fehler\n{t}\n"
    msg += f"Laufend: {len(aktive)}  |  Fehler: {len(fehler)}\n"
    if not fehler:
        msg += "\nKeine Fehler vorhanden."
    else:
        msg += "\nFehler-Jobs:\n"
        for j in fehler:
            phase     = phasen.get(j.get("error_phase"), "?")
            name      = (j.get("final_name") or os.path.basename(j.get("bildpfad", "")))[:28]
            fehlertext = (j.get("fehler_text") or "")[:80]
            msg += f"\n[{j['id']}] {name}\n"
            msg += f"  Phase:  {phase}\n"
            msg += f"  Fehler: {fehlertext}\n"
            msg += f"  → /retry {j['id']}\n"
        msg += f"\n{t}\n/retry_alle  – Alle Fehler neu starten"
    telegram_bot.sende_nachricht(msg)

def toggle_pause():
    global _scan_pausiert
    with _scan_pausiert_lock:
        _scan_pausiert = not _scan_pausiert
        return _scan_pausiert

def neustart_service():
    telegram_bot.sende_nachricht("Service wird neu gestartet...")
    time.sleep(1)
    subprocess.Popen(
        ["C:\\nssm\\nssm.exe", "restart", "dcp_automatisierung"],
        creationflags=subprocess.CREATE_NEW_CONSOLE
    )

# ──────────────────────────────────────────────
# Befehlshandler
# ──────────────────────────────────────────────

def bearbeite_befehl(text):
    cmd = text.strip()
    low = cmd.lower()
    t   = _trenn()

    if low == "/version":
        telegram_bot.sende_nachricht(f"DCP-Automatisierung v{lese_version()}")

    elif low == "/update":
        telegram_bot.sende_nachricht("Suche nach Updates...")
        threading.Thread(target=pruefe_update, daemon=True).start()

    elif low == "/status":
        fehler  = job_manager.hole_fehler()
        aktive  = job_manager.hole_aktive()
        pending = queue_manager.pending_anzahl()
        cfg     = lade_config()
        intervall       = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
        update_intervall = cfg.get("update", {}).get("auto_update_intervall_stunden", 24)
        update_str  = f"alle {update_intervall}h" if update_intervall > 0 else "aus"
        with _scan_pausiert_lock:
            pausiert = _scan_pausiert
        scan_str = "PAUSIERT" if pausiert else "Aktiv"
        telegram_bot.sende_nachricht(
            f"{t}\n"
            f"System: {scan_str}  |  v{lese_version()}\n"
            f"{t}\n"
            f"Queue:      {pending} wartend\n"
            f"Jobs:       {len(aktive)} laufend  |  {len(fehler)} Fehler\n"
            f"Scan:       alle {intervall} Min\n"
            f"Auto-Upd:   {update_str}\n"
            f"{t}"
        )

    elif low == "/check":
        with _scan_pausiert_lock:
            if _scan_pausiert:
                telegram_bot.sende_nachricht("Scan ist pausiert.\n/pause zum Fortsetzen.")
                return
        telegram_bot.sende_nachricht("Starte Ordner-Prüfung...")
        threading.Thread(target=starte_verarbeitung, daemon=True).start()

    elif low.startswith("/intervall"):
        teile = cmd.split(maxsplit=1)
        if len(teile) < 2:
            cfg    = lade_config()
            aktuell = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
            telegram_bot.sende_nachricht(
                f"Scan-Intervall: {aktuell} Minuten\nÄndern: /intervall <Minuten>"
            )
        else:
            aendere_intervall(teile[1])

    elif low.startswith("/update_intervall"):
        teile = cmd.split(maxsplit=1)
        if len(teile) < 2:
            cfg    = lade_config()
            aktuell = cfg.get("update", {}).get("auto_update_intervall_stunden", 24)
            telegram_bot.sende_nachricht(
                f"Auto-Update: {aktuell}h (0 = deaktiviert)\nÄndern: /update_intervall <Stunden>"
            )
        else:
            aendere_update_intervall(teile[1])

    elif low in ("/jobs", "/fehler"):
        zeige_jobs()

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
                f"Job {job_id} neu gestartet (Phase {j['current_phase']})..."
            )
            threading.Thread(
                target=verarbeite_job, args=(j["id"], j["current_phase"]), daemon=True
            ).start()
        else:
            telegram_bot.sende_nachricht(f"Job '{job_id}' nicht gefunden oder nicht wiederholbar.")

    elif low == "/pause":
        pausiert = toggle_pause()
        if pausiert:
            telegram_bot.sende_nachricht(
                "Scan pausiert.\nNeue Bilder werden nicht erkannt.\n/pause zum Fortsetzen."
            )
        else:
            telegram_bot.sende_nachricht("Scan fortgesetzt.")

    elif low == "/neustart":
        threading.Thread(target=neustart_service, daemon=True).start()

    elif low == "/hilfe":
        telegram_bot.sende_nachricht(
            f"{t}\n"
            f"DCP-Automatisierung\n"
            f"{t}\n"
            f"/status              System-Status\n"
            f"/check               Ordner jetzt prüfen\n"
            f"/version             Version anzeigen\n"
            f"\n"
            f"── Update ────────────────────\n"
            f"/update              Update suchen\n"
            f"/update_intervall <h>  Auto-Update\n"
            f"\n"
            f"── Einstellungen ─────────────\n"
            f"/intervall <n>       Scan alle n Min\n"
            f"\n"
            f"── Wartung ───────────────────\n"
            f"/jobs                Fehler & Retry\n"
            f"/pause               Scan pausieren\n"
            f"/neustart            Service neu starten\n"
            f"{t}"
        )

    else:
        telegram_bot.sende_nachricht(f"Unbekannt: {text}\n/hilfe für alle Befehle.")

# ──────────────────────────────────────────────
# Einstiegspunkt
# ──────────────────────────────────────────────

if __name__ == "__main__":
    queue_manager.naming_zuruecksetzen()

    cfg = lade_config()
    intervall        = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
    update_intervall = cfg.get("update", {}).get("auto_update_intervall_stunden", 24)

    job_manager.setze_status_callback(telegram_bot.sende_nachricht)

    pruefe_update_ergebnis()

    telegram_bot.sende_nachricht(f"DCP-Automatisierung v{lese_version()} gestartet.")

    threading.Thread(
        target=telegram_bot.starte_listener,
        args=(bearbeite_befehl,),
        daemon=True
    ).start()

    threading.Thread(target=queue_worker, daemon=True).start()

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
