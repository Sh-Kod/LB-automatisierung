"""
Watchdog für DCP-Automatisierung.
Prüft alle 60 Sekunden ob main.py noch läuft (via heartbeat.txt).
Sendet Telegram-Benachrichtigung bei Ausfall und bei Wiederherstellung.

Einrichten als NSSM-Service auf dem Server:
  nssm install dcp_watchdog python C:\dcp_automatisierung\watchdog.py
  nssm start dcp_watchdog
"""

import time
from datetime import datetime

import requests
import yaml

HEARTBEAT_PFAD  = "C:\\dcp_automatisierung\\heartbeat.txt"
CONFIG_PFAD     = "C:\\dcp_automatisierung\\config.yaml"
PRUEF_INTERVALL = 60    # Sekunden zwischen Prüfungen
TIMEOUT_SEKUNDEN = 120  # Heartbeat gilt als abgelaufen nach 2 Minuten


def _lade_config():
    with open(CONFIG_PFAD, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def _sende_nachricht(text):
    try:
        config = _lade_config()
        token   = config["telegram"]["token"]
        chat_id = config["telegram"]["chat_id"]
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        requests.post(url, data={"chat_id": chat_id, "text": text}, timeout=10)
    except Exception as e:
        print(f"Telegram Fehler: {e}")


def _lese_heartbeat():
    """Gibt den Timestamp aus heartbeat.txt zurück, oder None bei Fehler."""
    try:
        with open(HEARTBEAT_PFAD, "r", encoding="utf-8") as f:
            return float(f.read().strip())
    except Exception:
        return None


def main():
    print("Watchdog gestartet.")
    ausgefallen = False

    while True:
        time.sleep(PRUEF_INTERVALL)

        ts = _lese_heartbeat()
        jetzt = time.time()

        if ts is None:
            alter = TIMEOUT_SEKUNDEN + 1  # Datei fehlt → gilt als abgelaufen
        else:
            alter = jetzt - ts

        if alter > TIMEOUT_SEKUNDEN and not ausgefallen:
            ausgefallen = True
            if ts is not None:
                letztes = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
                meldung = f"DCP-Automatisierung nicht erreichbar!\nLetztes Lebenszeichen: {letztes}"
            else:
                meldung = "DCP-Automatisierung nicht erreichbar!\nHeartbeat-Datei fehlt."
            print(meldung)
            _sende_nachricht(meldung)

        elif alter <= TIMEOUT_SEKUNDEN and ausgefallen:
            ausgefallen = False
            # Hinweis: main.py sendet selbst eine "gestartet"-Nachricht –
            # der Watchdog sendet hier nichts, um doppelte Meldungen zu vermeiden.
            print("Heartbeat wieder aktiv.")


if __name__ == "__main__":
    main()
