import requests
import yaml
import time
import threading

_last_update_id = 0
_update_id_lock = threading.Lock()

# --- Dialog-State: genau 1 aktiver Benennungsdialog ---
_dialog_aktiv = False
_dialog_aktiv_lock = threading.Lock()
_dialog_antwort = None
_dialog_event = threading.Event()

def lade_config():
    with open("C:\\dcp_automatisierung\\config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def sende_nachricht(text):
    try:
        config = lade_config()
        token = config["telegram"]["token"]
        chat_id = config["telegram"]["chat_id"]
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        requests.post(url, data={"chat_id": chat_id, "text": text}, timeout=10)
    except Exception as e:
        print(f"Telegram Fehler: {e}")

def sende_bild(bildpfad, caption=""):
    try:
        config = lade_config()
        token = config["telegram"]["token"]
        chat_id = config["telegram"]["chat_id"]
        url = f"https://api.telegram.org/bot{token}/sendPhoto"
        with open(bildpfad, "rb") as f:
            requests.post(url, data={"chat_id": chat_id, "caption": caption},
                          files={"photo": f}, timeout=30)
    except Exception as e:
        print(f"Telegram Bild Fehler: {e}")

def starte_dialog():
    """Belegt den Dialog-Slot. Gibt True zurueck wenn erfolgreich (kein anderer Dialog aktiv)."""
    global _dialog_aktiv
    with _dialog_aktiv_lock:
        if _dialog_aktiv:
            return False
        _dialog_aktiv = True
        _dialog_event.clear()
        return True

def beende_dialog():
    """Gibt den Dialog-Slot frei und entsperrt wartendes warte_auf_dialog_antwort()."""
    global _dialog_aktiv
    with _dialog_aktiv_lock:
        _dialog_aktiv = False
    _dialog_event.set()

def warte_auf_dialog_antwort(timeout=3600):
    """Wartet auf Nutzer-Eingabe waehrend eines aktiven Dialogs.
    Gibt den eingegebenen Text zurueck, oder None bei Timeout."""
    global _dialog_antwort
    _dialog_antwort = None
    _dialog_event.clear()
    _dialog_event.wait(timeout=timeout)
    with _dialog_aktiv_lock:
        return _dialog_antwort

def starte_listener(callback):
    """Lauft dauerhaft. Leitet Nachrichten an Dialog oder Befehlshandler weiter."""
    global _last_update_id, _dialog_antwort
    try:
        config = lade_config()
        token = config["telegram"]["token"]
        url = f"https://api.telegram.org/bot{token}/getUpdates"
        resp = requests.get(url, params={"timeout": 0}, timeout=10)
        updates = resp.json().get("result", [])
        with _update_id_lock:
            _last_update_id = max([u["update_id"] for u in updates], default=0) if updates else 0
    except Exception:
        pass

    while True:
        try:
            config = lade_config()
            token = config["telegram"]["token"]
            chat_id = str(config["telegram"]["chat_id"])
            url = f"https://api.telegram.org/bot{token}/getUpdates"
            with _update_id_lock:
                offset = _last_update_id + 1
            resp = requests.get(url, params={"offset": offset, "timeout": 10}, timeout=15)
            updates = resp.json().get("result", [])
            for update in updates:
                with _update_id_lock:
                    _last_update_id = update["update_id"]
                msg = update.get("message", {})
                if str(msg.get("chat", {}).get("id", "")) != chat_id:
                    continue
                text = msg.get("text", "").strip()
                if not text:
                    continue
                # Routing: Dialog aktiv → Antwort an Dialog; sonst → Befehlshandler
                with _dialog_aktiv_lock:
                    dialog = _dialog_aktiv
                if dialog:
                    _dialog_antwort = text
                    _dialog_event.set()
                elif callback:
                    threading.Thread(target=callback, args=(text,), daemon=True).start()
        except Exception:
            time.sleep(5)
