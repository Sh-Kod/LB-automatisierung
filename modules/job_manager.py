import json
import os
import threading
import uuid
from datetime import datetime

JOBS_PFAD = "C:\\dcp_automatisierung\\jobs.json"
_lock = threading.Lock()
_status_buffer = []
_status_buffer_lock = threading.Lock()
_status_callback = None
_naming_check = None  # callable → True wenn Naming noch läuft

PHASEN = {1: "DCP", 2: "Upload", 3: "Ingest", 4: "Monitoring"}

def setze_status_callback(fn):
    global _status_callback
    _status_callback = fn

def setze_naming_check(fn):
    global _naming_check
    _naming_check = fn

def _lade():
    if not os.path.exists(JOBS_PFAD):
        return {"jobs": []}
    try:
        with open(JOBS_PFAD, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"jobs": []}

def _speichere(data):
    tmp = JOBS_PFAD + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, JOBS_PFAD)

def erstelle_job(bildpfad, final_name):
    """Legt neuen Job an. Gibt Job-ID zurueck."""
    with _lock:
        data = _lade()
        job = {
            "id": str(uuid.uuid4())[:8],
            "bildpfad": bildpfad,
            "final_name": final_name,
            "current_phase": 1,
            "last_success_phase": 0,
            "error_phase": None,
            "current_status": "running",
            "retryable": True,
            "fehler_text": None,
            "timestamp": datetime.now().isoformat()
        }
        data["jobs"].append(job)
        _speichere(data)
        return job["id"]

def aktualisiere_phase(job_id, phase, status="running"):
    """Aktualisiert Phase und Status eines Jobs."""
    with _lock:
        data = _lade()
        for job in data["jobs"]:
            if job["id"] == job_id:
                job["current_phase"] = phase
                job["current_status"] = status
                if status == "done":
                    job["last_success_phase"] = phase
                    job["error_phase"] = None
                    job["fehler_text"] = None
                _speichere(data)
                return

def markiere_fehler(job_id, phase, fehler_text, retryable=True):
    """Markiert einen Job als fehlerhaft. Loest Status-Meldung aus."""
    job_data = None
    aktive_nach = 0
    with _lock:
        data = _lade()
        for job in data["jobs"]:
            if job["id"] == job_id:
                job["current_phase"] = phase
                job["current_status"] = "error"
                job["error_phase"] = phase
                job["fehler_text"] = str(fehler_text)[:500]
                job["retryable"] = retryable
                job_data = dict(job)
                break
        _speichere(data)
        aktive_nach = sum(1 for j in data["jobs"] if j["current_status"] == "running")
    if job_data:
        _melde_status(job_data, aktive_nach)

def markiere_fertig(job_id):
    """Markiert Job als vollstaendig abgeschlossen. Loest Status-Meldung aus."""
    job_data = None
    aktive_nach = 0
    with _lock:
        data = _lade()
        for job in data["jobs"]:
            if job["id"] == job_id:
                job["current_status"] = "done"
                job["last_success_phase"] = 4
                job_data = dict(job)
                break
        _speichere(data)
        aktive_nach = sum(1 for j in data["jobs"] if j["current_status"] == "running")
    if job_data:
        _melde_status(job_data, aktive_nach)

def hole_job(job_id):
    """Gibt Job-Dict fuer job_id zurueck, oder None."""
    with _lock:
        data = _lade()
        for j in data["jobs"]:
            if j["id"] == job_id:
                return dict(j)
    return None

def hole_fehler():
    with _lock:
        data = _lade()
        return [j for j in data["jobs"] if j["current_status"] == "error"]

def hole_aktive():
    with _lock:
        data = _lade()
        return [j for j in data["jobs"] if j["current_status"] == "running"]

def retry_job(job_id):
    """Setzt einen Fehler-Job auf retry_pending ab der richtigen Phase. Gibt Job-Dict zurueck."""
    with _lock:
        data = _lade()
        for job in data["jobs"]:
            if job["id"] == job_id and job.get("retryable") and job["current_status"] == "error":
                if job.get("error_phase") == 1:
                    retry_phase = 1
                else:
                    retry_phase = max(1, (job.get("last_success_phase") or 0) + 1)
                job["current_phase"] = retry_phase
                job["current_status"] = "retry_pending"
                job["error_phase"] = None
                job["fehler_text"] = None
                _speichere(data)
                return dict(job)
        return None

def alle_retry():
    """Setzt alle Fehler-Jobs auf retry_pending. Gibt Liste der Jobs zurueck."""
    with _lock:
        data = _lade()
        result = []
        for job in data["jobs"]:
            if job["current_status"] == "error" and job.get("retryable"):
                retry_phase = 1 if job.get("error_phase") == 1 else max(1, (job.get("last_success_phase") or 0) + 1)
                job["current_phase"] = retry_phase
                job["current_status"] = "retry_pending"
                job["error_phase"] = None
                job["fehler_text"] = None
                result.append(dict(job))
        _speichere(data)
        return result

# --- Status-Buendelung ---
# Regel: Gibt es noch aktive Jobs → puffern.
#        Letzter Job fertig → Bundle sofort senden.
#        Alle 5 Minuten → Puffer senden falls gefuellt.
#        1 Job allein → Sofort senden.

def _melde_status(job, aktive_nach):
    with _status_buffer_lock:
        _status_buffer.append(job)
        if aktive_nach == 0:
            # Nur flushen wenn Naming fertig
            naming = _naming_check() if _naming_check else False
            if not naming:
                _flush_buffer()

def _flush_buffer():
    """Muss mit _status_buffer_lock aufgerufen werden."""
    if not _status_buffer or not _status_callback:
        return
    fertig = [j for j in _status_buffer if j["current_status"] == "done"]
    fehler = [j for j in _status_buffer if j["current_status"] == "error"]
    if not fertig and not fehler:
        _status_buffer.clear()
        return

    if len(_status_buffer) == 1 and fertig:
        # Einzelner Job: einfache Fertig-Meldung
        name = (fertig[0].get("final_name") or os.path.basename(fertig[0].get("bildpfad", "")))
        _status_callback(f"Fertig: {name}")
    else:
        # Batch-Abschluss
        t = "─" * 30
        msg = f"{t}\nAlle fertig!\n{t}\n"
        if fertig:
            msg += f"✓ {len(fertig)} DCP(s) erfolgreich\n"
        if fehler:
            msg += f"✗ {len(fehler)} Fehler:\n"
            for j in fehler[:5]:
                phase = PHASEN.get(j.get("error_phase"), "?")
                name  = (j.get("final_name") or os.path.basename(j.get("bildpfad", "")))[:30]
                msg  += f"  • {name} [{phase}]\n"
            msg += "/jobs für Details  |  /retry_alle"
        _status_callback(msg)
    _status_buffer.clear()

def _sende_einzeln(job):
    # Nicht mehr verwendet – Fehler werden von main._phase_ausfuehren gemeldet
    pass
