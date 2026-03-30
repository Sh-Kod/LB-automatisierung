#!/usr/bin/env python3
# updater.py - DCP Automatisierung In-App Updater
# Laeuft als separater Prozess ausserhalb des Windows-Dienstes.
# Stdlib only - keine externen Abhaengigkeiten.

import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime

BASE_PFAD = "C:\\dcp_automatisierung"
PENDING_PFAD = os.path.join(BASE_PFAD, "pending_update.json")
RESULT_PFAD = os.path.join(BASE_PFAD, "update_result.json")
BACKUP_PFAD = os.path.join(BASE_PFAD, "backup")
NSSM = "C:\\nssm\\nssm.exe"
SERVICE = "dcp_automatisierung"
LOG_PFAD = os.path.join(BASE_PFAD, "logs", "updater.log")


def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    try:
        os.makedirs(os.path.dirname(LOG_PFAD), exist_ok=True)
        with open(LOG_PFAD, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def schreibe_result(erfolg, neue_version="", fehler=""):
    result = {
        "erfolg": erfolg,
        "neue_version": neue_version,
        "fehler": fehler,
        "timestamp": datetime.now().isoformat()
    }
    try:
        tmp = RESULT_PFAD + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        os.replace(tmp, RESULT_PFAD)
    except Exception as e:
        log(f"Konnte update_result.json nicht schreiben: {e}")


def stoppe_service():
    log("Stoppe Service...")
    try:
        subprocess.run([NSSM, "stop", SERVICE], capture_output=True, timeout=30)
    except Exception as e:
        log(f"Fehler beim Stoppen: {e}")
    for _ in range(20):
        time.sleep(2)
        try:
            r = subprocess.run([NSSM, "status", SERVICE],
                               capture_output=True, text=True, timeout=10)
            out = r.stdout + r.stderr
            if "SERVICE_STOPPED" in out:
                log("Service gestoppt.")
                return True
        except Exception:
            pass
    log("Service stop timeout - fahre trotzdem fort.")
    return True


def starte_service():
    log("Starte Service...")
    try:
        subprocess.run([NSSM, "start", SERVICE], capture_output=True, timeout=30)
        time.sleep(3)
        log("Service gestartet.")
    except Exception as e:
        log(f"Fehler beim Starten: {e}")


def erstelle_backup(dateien, backup_dir):
    log(f"Erstelle Backup in: {backup_dir}")
    os.makedirs(backup_dir, exist_ok=True)
    gesichert = 0
    for entry in dateien:
        dest = entry["dest"]
        if os.path.exists(dest):
            rel = os.path.relpath(dest, BASE_PFAD)
            backup_dest = os.path.join(backup_dir, rel)
            os.makedirs(os.path.dirname(backup_dest), exist_ok=True)
            shutil.copy2(dest, backup_dest)
            gesichert += 1
    log(f"Backup abgeschlossen: {gesichert} Dateien gesichert.")


def rollback(dateien, backup_dir):
    log(f"Fuehre Rollback durch aus: {backup_dir}")
    wiederhergestellt = 0
    for entry in dateien:
        dest = entry["dest"]
        rel = os.path.relpath(dest, BASE_PFAD)
        backup_src = os.path.join(backup_dir, rel)
        if os.path.exists(backup_src):
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.copy2(backup_src, dest)
            wiederhergestellt += 1
    log(f"Rollback abgeschlossen: {wiederhergestellt} Dateien wiederhergestellt.")


def wende_update_an(dateien, staging_pfad):
    log("Wende Update an...")
    for entry in dateien:
        src_rel = entry["src"]
        dest = entry["dest"]
        staged = os.path.join(staging_pfad, src_rel.replace("/", os.sep))
        if not os.path.exists(staged):
            raise FileNotFoundError(f"Staged file nicht gefunden: {staged}")
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        tmp = dest + ".new"
        shutil.copy2(staged, tmp)
        os.replace(tmp, dest)
        log(f"  -> {dest}")
    log("Update angewendet.")


def main():
    log("=" * 50)
    log("DCP Automatisierung Updater gestartet")
    log("=" * 50)

    if not os.path.exists(PENDING_PFAD):
        log(f"Kein pending_update.json gefunden: {PENDING_PFAD}")
        sys.exit(1)

    try:
        with open(PENDING_PFAD, "r", encoding="utf-8") as f:
            pending = json.load(f)
    except Exception as e:
        log(f"Konnte pending_update.json nicht lesen: {e}")
        sys.exit(1)

    neue_version = pending.get("neue_version", "?")
    staging_pfad = pending.get("staging_pfad", "")
    dateien = pending.get("dateien", [])

    log(f"Update auf Version: {neue_version}")
    log(f"Anzahl Dateien: {len(dateien)}")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = os.path.join(BACKUP_PFAD, f"v{neue_version}_{ts}")

    # Service stoppen (stoppt auch den laufenden main.py-Prozess)
    stoppe_service()
    time.sleep(2)

    # Backup erstellen
    try:
        erstelle_backup(dateien, backup_dir)
    except Exception as e:
        log(f"Backup fehlgeschlagen: {e}")
        schreibe_result(False, neue_version, f"Backup fehlgeschlagen: {e}")
        starte_service()
        sys.exit(1)

    # Update anwenden
    try:
        wende_update_an(dateien, staging_pfad)
    except Exception as e:
        log(f"Update fehlgeschlagen: {e}")
        log("Starte Rollback...")
        try:
            rollback(dateien, backup_dir)
            log("Rollback erfolgreich.")
        except Exception as re:
            log(f"Rollback fehlgeschlagen: {re}")
        schreibe_result(False, neue_version, str(e))
        starte_service()
        sys.exit(1)

    # version.txt aktualisieren
    try:
        version_pfad = os.path.join(BASE_PFAD, "version.txt")
        tmp = version_pfad + ".new"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(neue_version)
        os.replace(tmp, version_pfad)
        log(f"Version aktualisiert: {neue_version}")
    except Exception as e:
        log(f"Konnte version.txt nicht aktualisieren: {e}")

    # Staging-Verzeichnis bereinigen
    try:
        if os.path.exists(staging_pfad):
            shutil.rmtree(staging_pfad)
    except Exception:
        pass

    # pending_update.json entfernen
    try:
        os.remove(PENDING_PFAD)
    except Exception:
        pass

    schreibe_result(True, neue_version)
    log(f"Update auf v{neue_version} erfolgreich abgeschlossen!")

    # Service neu starten
    starte_service()
    log("Updater beendet.")


if __name__ == "__main__":
    main()
