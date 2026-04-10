"""
Doremi DCP2000 – HTTP Web Interface API
Startet Ingest über das Web-Interface (SCHEDULE_INGEST_TASK)
und überwacht den Fortschritt über die Monitor-Seite.

Vorteil gegenüber KLV TCP: Zuverlässig, browser-kompatibel,
kein binäres Protokoll nötig.

Login:  POST http://{ip}/web/index.php
        Felder: username, password, screen
        Antwort: PHPSESSID Cookie

Ingest: POST http://{ip}/web/sys_control/ingest_manager/ajax.php
        Felder: request=SCHEDULE_INGEST_TASK,
                tasks[0][uri]=/data/incoming/gui/{dcp_name}/{pkl_datei}
                tasks[0][description], tasks[0][type]=PackagingList
                destination=, pingest=false

Monitor: GET http://{ip}/web/sys_control/index.php?page=ingest_manager/ingest_monitor.php
         HTML-Tabelle mit job_id (checkbox value), Status-Icon und DCP-Beschreibung.
         Icons: success_16.png = Erfolg, error-icon-16.png = Fehler

Benötigt: pip install requests
"""

import logging
import os
import re
import glob as glob_mod

log = logging.getLogger("dcp_automatisierung")


def login(ip, user="admin", password="1234"):
    """
    Meldet sich am Doremi-Webinterface an.
    Gibt PHPSESSID zurück oder wirft RuntimeError.
    """
    try:
        import requests
    except ImportError:
        raise RuntimeError("requests nicht installiert: pip install requests")

    url = f"http://{ip}/web/index.php"
    data = {
        "username": user,
        "password": password,
        "screen": "auto",
    }
    session = requests.Session()
    try:
        session.post(url, data=data, timeout=15, allow_redirects=True)
        phpsessid = session.cookies.get("PHPSESSID")
        if not phpsessid:
            raise RuntimeError("Kein PHPSESSID nach Login erhalten.")
        log.info(f"[Web] Login OK – PHPSESSID={phpsessid[:8]}...")
        return phpsessid
    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"HTTP-Login fehlgeschlagen: {e}")


def finde_pkl_datei(dcp_pfad_lokal):
    """
    Sucht die PKL-Datei (pkl_*.xml) im lokalen DCP-Ordner.
    Gibt den Dateinamen zurück (nur Dateiname, kein Pfad).
    Wirft RuntimeError wenn keine gefunden.
    """
    for muster in ("pkl_*.xml", "PKL_*.xml", "PKL*.xml"):
        treffer = glob_mod.glob(os.path.join(dcp_pfad_lokal, muster))
        if treffer:
            return os.path.basename(treffer[0])
    raise RuntimeError(
        f"Keine PKL-Datei (pkl_*.xml) in '{dcp_pfad_lokal}' gefunden."
    )


def starte_ingest(ip, phpsessid, dcp_name, pkl_datei):
    """
    Startet Ingest via HTTP SCHEDULE_INGEST_TASK.

    pkl_datei: Nur der Dateiname, z.B. 'pkl_36ac7cf5-e43b-40d6-b0ab-80372114f580.xml'
    URI auf Doremi: /data/incoming/gui/{dcp_name}/{pkl_datei}

    Gibt True zurück bei HTTP 200.
    Wirft RuntimeError bei Fehler.
    """
    try:
        import requests
    except ImportError:
        raise RuntimeError("requests nicht installiert: pip install requests")

    url = f"http://{ip}/web/sys_control/ingest_manager/ajax.php"
    uri = f"/data/incoming/gui/{dcp_name}/{pkl_datei}"

    data = {
        "request": "SCHEDULE_INGEST_TASK",
        "tasks[0][uri]": uri,
        "tasks[0][description]": dcp_name,
        "tasks[0][type]": "PackagingList",
        "destination": "",
        "pingest": "false",
    }
    cookies = {
        "PHPSESSID": phpsessid,
        "interfaceSize": "auto",
    }

    try:
        resp = requests.post(url, data=data, cookies=cookies, timeout=20)
        resp.raise_for_status()
        log.info(
            f"[Web] SCHEDULE_INGEST_TASK gesendet – URI: {uri} – HTTP {resp.status_code}"
        )
        log.debug(f"[Web] Antwort: {resp.text[:300]}")
        return True
    except Exception as e:
        raise RuntimeError(f"SCHEDULE_INGEST_TASK fehlgeschlagen: {e}")


def pruefe_ingest_status(ip, phpsessid, dcp_name):
    """
    Liest die Doremi Ingest-Monitor-Seite und sucht den Job für dcp_name.

    Gibt (job_id, status) zurück:
      - job_id: Integer (Doremi-Job-ID) oder None wenn nicht gefunden
      - status:  "success"   → success_16.png
                 "error"     → error-icon-16.png (oder ähnlich)
                 "pending"   → Job gefunden, aber noch kein Abschluss-Icon
                 "not_found" → dcp_name nicht in der Tabelle

    Bei mehreren Treffern mit gleichem Namen wird der Job mit der
    höchsten Job-ID (= aktuellster) zurückgegeben.
    """
    try:
        import requests
    except ImportError:
        raise RuntimeError("requests nicht installiert: pip install requests")

    url = (
        f"http://{ip}/web/sys_control/index.php"
        f"?page=ingest_manager/ingest_monitor.php"
    )
    cookies = {"PHPSESSID": phpsessid, "interfaceSize": "auto"}

    try:
        resp = requests.get(url, cookies=cookies, timeout=20)
        resp.raise_for_status()
        html = resp.text
    except Exception as e:
        raise RuntimeError(f"Monitor-Seite nicht erreichbar: {e}")

    # HTML-Tabelle parsen: jede <tr> auf Job-ID, Icon und Beschreibung prüfen
    rows = re.findall(r"<tr\b[^>]*>(.*?)</tr>", html, re.DOTALL | re.IGNORECASE)
    treffer = []
    for row_html in rows:
        # Job-ID aus checkbox value
        jid_m = re.search(
            r'<input[^>]+value=["\'](\d+)["\']',
            row_html,
            re.IGNORECASE,
        )
        if not jid_m:
            continue
        job_id = int(jid_m.group(1))

        # Status-Icon src
        icon_m = re.search(
            r'<img[^>]+src=["\']([^"\']+)["\']',
            row_html,
            re.IGNORECASE,
        )
        icon_src = icon_m.group(1) if icon_m else ""

        # DCP-Beschreibung aus <label>
        label_m = re.search(r"<label[^>]*>([^<]+)</label>", row_html, re.IGNORECASE)
        if not label_m:
            continue
        beschreibung = label_m.group(1).strip()

        if beschreibung == dcp_name:
            if "success_16" in icon_src:
                status = "success"
            elif "error" in icon_src.lower():
                status = "error"
            else:
                status = "pending"
            treffer.append((job_id, status))

    if not treffer:
        log.debug(f"[Web] Monitor: '{dcp_name}' nicht gefunden")
        return None, "not_found"

    # Höchste Job-ID = aktuellster Job
    treffer.sort(key=lambda x: x[0], reverse=True)
    job_id, status = treffer[0]
    log.info(f"[Web] Monitor: '{dcp_name}' → job_id={job_id}, status={status}")
    return job_id, status
