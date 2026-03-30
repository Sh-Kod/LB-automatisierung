import json
import os
import threading
import uuid
from datetime import datetime

QUEUE_PFAD = "C:\\dcp_automatisierung\\queue.json"
_lock = threading.Lock()

def _lade():
    if not os.path.exists(QUEUE_PFAD):
        return {"items": []}
    try:
        with open(QUEUE_PFAD, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"items": []}

def _speichere(data):
    tmp = QUEUE_PFAD + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, QUEUE_PFAD)

def hinzufuegen(bildpfad, dauer_sek):
    """Fuegt Bild zur Queue hinzu. Gibt ID zurueck, oder None wenn bereits vorhanden."""
    with _lock:
        data = _lade()
        for item in data["items"]:
            if item["bildpfad"] == bildpfad and item["status"] in ("pending", "naming"):
                return None
        item_id = str(uuid.uuid4())[:8]
        data["items"].append({
            "id": item_id,
            "bildpfad": bildpfad,
            "dauer_sek": dauer_sek,
            "timestamp": datetime.now().isoformat(),
            "status": "pending"
        })
        _speichere(data)
        return item_id

def naechstes_pending():
    """Holt naechstes pending-Item und markiert es als 'naming'. Thread-safe."""
    with _lock:
        data = _lade()
        for item in data["items"]:
            if item["status"] == "pending":
                item["status"] = "naming"
                _speichere(data)
                return dict(item)
        return None

def zuruecksetzen(item_id):
    """Setzt ein 'naming'-Item zurueck auf 'pending' (z.B. bei Timeout oder Fehler)."""
    with _lock:
        data = _lade()
        for item in data["items"]:
            if item["id"] == item_id and item["status"] == "naming":
                item["status"] = "pending"
                _speichere(data)
                return

def abschliessen(item_id):
    """Entfernt Item aus Queue nach Abschluss."""
    with _lock:
        data = _lade()
        data["items"] = [i for i in data["items"] if i["id"] != item_id]
        _speichere(data)

def pending_anzahl():
    """Gibt Anzahl wartender Items zurueck."""
    with _lock:
        data = _lade()
        return sum(1 for i in data["items"] if i["status"] == "pending")

def naming_zuruecksetzen():
    """Beim Systemstart: alle 'naming'-Items zurueck auf 'pending' (Absturz-Recovery)."""
    with _lock:
        data = _lade()
        geaendert = False
        for item in data["items"]:
            if item["status"] == "naming":
                item["status"] = "pending"
                geaendert = True
        if geaendert:
            _speichere(data)
