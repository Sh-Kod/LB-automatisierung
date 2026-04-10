"""
Doremi DCP2000 – Native TCP API
Port: 11730
Protokoll: KLV (Key-Length-Value) binär

Verwendete Befehle:
  IngestAddJob      (0x070F00) – Ingest starten, gibt job_id zurück
  IngestGetJobStatus (0x071D00) – Status eines Ingest-Jobs abfragen

Ingest-Status-Codes:
  0 = pending
  1 = paused
  2 = running
  3 = scheduled
  4 = success   ← fertig + erfolgreich
  5 = aborted
  6 = unused
  7 = failed    ← Fehler
"""

import socket
import struct
import logging

log = logging.getLogger("dcp_automatisierung")

DOREMI_PORT = 11730
SOCKET_TIMEOUT = 30  # Sekunden

# KLV-Header (13 Bytes) – identisch bei allen DCP2000-Nachrichten
KLV_HEADER = bytes([
    0x06, 0x0E, 0x2B, 0x34, 0x02, 0x05, 0x01, 0x0A,
    0x0E, 0x10, 0x01, 0x01, 0x01
])

# Befehlsschlüssel (Request)
CMD_INGEST_ADD_JOB    = bytes([0x07, 0x0F, 0x00])
CMD_INGEST_CANCEL_JOB = bytes([0x07, 0x11, 0x00])   # IngestCancelJob
CMD_INGEST_GET_STATUS = bytes([0x07, 0x1D, 0x00])
CMD_INGEST_GET_LIST   = bytes([0x07, 0x19, 0x00])   # IngestGetJobList (Diagnose)
CMD_WHO_AM_I          = bytes([0x0E, 0x0B, 0x00])
CMD_GET_API_VERSION   = bytes([0x05, 0x05, 0x00])   # GetAPIProtocolVersion

# Antwortschlüssel (Response)
RESP_INGEST_ADD_JOB    = bytes([0x07, 0x10, 0x00])
RESP_INGEST_CANCEL_JOB = bytes([0x07, 0x12, 0x00])
RESP_INGEST_GET_STATUS = bytes([0x07, 0x1E, 0x00])
RESP_WHO_AM_I          = bytes([0x0E, 0x0C, 0x00])

STATUS_NAMEN = {
    0: "pending",
    1: "paused",
    2: "running",
    3: "scheduled",
    4: "success",
    5: "aborted",
    6: "unused",
    7: "failed",
}

_request_counter = 0


def _naechste_request_id():
    global _request_counter
    _request_counter = (_request_counter + 1) % 60000
    return _request_counter


def _ber_encode(laenge):
    """Länge im BER-Format kodieren."""
    if laenge < 0x80:
        return bytes([laenge])
    elif laenge < 0x100:
        return bytes([0x81, laenge])
    elif laenge < 0x10000:
        return bytes([0x82, (laenge >> 8) & 0xFF, laenge & 0xFF])
    else:
        return bytes([
            0x83,
            (laenge >> 16) & 0xFF,
            (laenge >> 8)  & 0xFF,
            laenge         & 0xFF
        ])


def _baue_nachricht(cmd_key, payload):
    """KLV-Nachricht aufbauen: Header + Schlüssel + BER-Länge + Request-ID + Payload."""
    req_id = _naechste_request_id()
    req_id_bytes = struct.pack(">I", req_id)      # 4 Byte big-endian
    voller_payload = req_id_bytes + payload
    laenge_bytes = _ber_encode(len(voller_payload))
    nachricht = KLV_HEADER + cmd_key + laenge_bytes + voller_payload
    return nachricht, req_id


def _lese_genau(sock, n):
    """Exakt n Bytes vom Socket lesen."""
    data = b""
    while len(data) < n:
        chunk = sock.recv(n - len(data))
        if not chunk:
            raise ConnectionError("Verbindung vom Doremi-Server geschlossen")
        data += chunk
    return data


def _lese_antwort(sock):
    """KLV-Antwortnachricht lesen und parsen. Gibt (cmd_key, payload) zurück."""
    # Header lesen (13 Bytes)
    header = _lese_genau(sock, 13)
    if header != KLV_HEADER:
        raise ValueError(f"Ungültiger KLV-Header: {header.hex()}")

    # Befehlsschlüssel lesen (3 Bytes)
    cmd_key = _lese_genau(sock, 3)

    # BER-Länge lesen
    erstes_byte = _lese_genau(sock, 1)[0]
    if erstes_byte < 0x80:
        laenge = erstes_byte
    elif erstes_byte == 0x81:
        laenge = _lese_genau(sock, 1)[0]
    elif erstes_byte == 0x82:
        b = _lese_genau(sock, 2)
        laenge = (b[0] << 8) | b[1]
    elif erstes_byte == 0x83:
        b = _lese_genau(sock, 3)
        laenge = (b[0] << 16) | (b[1] << 8) | b[2]
    else:
        raise ValueError(f"Unbekannte BER-Kodierung: 0x{erstes_byte:02X}")

    # Payload inkl. 4-Byte Request-ID lesen
    payload_mit_id = _lese_genau(sock, laenge)
    # Erste 4 Bytes = Request-ID, der Rest ist der eigentliche Payload
    payload = payload_mit_id[4:]

    return cmd_key, payload


def _verbinde(ip):
    """TCP-Verbindung zu Doremi herstellen."""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(SOCKET_TIMEOUT)
    sock.connect((ip, DOREMI_PORT))
    return sock


# ──────────────────────────────────────────────
# Öffentliche API
# ──────────────────────────────────────────────

def who_am_i(ip):
    """
    Diagnoseanfrage WhoAmI – prüft ob TCP/KLV-Protokoll grundsätzlich funktioniert.
    Gibt die Antwort als Text zurück (Gerätename o.ä.).
    Wirft Exception bei Verbindungs- oder Protokollfehler.
    """
    nachricht, _ = _baue_nachricht(CMD_WHO_AM_I, b"")
    with _verbinde(ip) as sock:
        sock.sendall(nachricht)
        cmd_key, resp_payload = _lese_antwort(sock)
    log.info(
        f"[Doremi API] WhoAmI Response – key={cmd_key.hex()}, "
        f"payload={resp_payload.hex()}"
    )
    try:
        return resp_payload.decode("utf-8", errors="replace").strip("\x00").strip()
    except Exception:
        return resp_payload.hex()


def get_api_version(ip):
    """
    GetAPIProtocolVersion – gibt API-Versionsinformation zurück.
    Gibt (cmd_key_hex, payload_hex) zurück.
    """
    nachricht, _ = _baue_nachricht(CMD_GET_API_VERSION, b"")
    try:
        with _verbinde(ip) as sock:
            sock.sendall(nachricht)
            cmd_key, resp_payload = _lese_antwort(sock)
        log.info(
            f"[Doremi API] GetAPIVersion Response – key={cmd_key.hex()}, "
            f"payload={resp_payload.hex()}"
        )
        return cmd_key.hex(), resp_payload.hex()
    except Exception as e:
        log.warning(f"[Doremi API] GetAPIVersion Fehler: {e}")
        return None, str(e)


def get_ingest_list(ip):
    """
    IngestGetJobList – listet aktive/wartende Ingest-Jobs auf dem Doremi.
    Gibt (cmd_key_hex, payload_hex) zurück – nützlich zur Diagnose.
    """
    nachricht, _ = _baue_nachricht(CMD_INGEST_GET_LIST, b"")
    try:
        with _verbinde(ip) as sock:
            sock.sendall(nachricht)
            cmd_key, resp_payload = _lese_antwort(sock)
        log.info(
            f"[Doremi API] IngestGetList Response – key={cmd_key.hex()}, "
            f"payload={resp_payload.hex()}"
        )
        return cmd_key.hex(), resp_payload.hex()
    except Exception as e:
        log.warning(f"[Doremi API] IngestGetList Fehler: {e}")
        return None, str(e)


def _generiere_pfad_varianten(assetmap_pfad):
    """Generiert Pfad-Varianten für IngestAddJob (primärer Pfad zuerst).

    Doremi-interne Pfade können je nach Firmware-Version und FTP-Konfiguration
    in verschiedenen Formaten erwartet werden. Getestete Varianten:
      - Vollpfad mit /incoming-Präfix: /incoming/gui/{name}/ASSETMAP.xml
      - Ohne /incoming-Präfix:         /gui/{name}/ASSETMAP.xml
      - Ohne .xml-Extension:           /incoming/gui/{name}/ASSETMAP
      - Nur Ordnerpfad:                /incoming/gui/{name}
      - Relativer Pfad:                {name}/ASSETMAP.xml
      - Nur DCP-Name:                  {name}
    """
    varianten = [assetmap_pfad]
    assetmap_upper = assetmap_pfad.upper()

    # Basis-Pfad ohne /incoming-Präfix (oder mit, falls nicht vorhanden)
    if assetmap_pfad.startswith("/incoming"):
        ohne_incoming = assetmap_pfad[len("/incoming"):]
    else:
        ohne_incoming = assetmap_pfad

    mit_incoming = (
        assetmap_pfad if assetmap_pfad.startswith("/incoming")
        else "/incoming" + assetmap_pfad
    )

    # --- ASSETMAP-Datei-Varianten ---
    if "ASSETMAP" in assetmap_upper:
        idx      = assetmap_upper.rfind("/ASSETMAP")
        # Pfad bis inkl. letztem "/" vor ASSETMAP
        ordner   = assetmap_pfad[:idx]          # z.B. /incoming/gui/{name}
        am_datei = assetmap_pfad[idx + 1:]      # z.B. ASSETMAP.xml

        # Ohne .xml-Extension
        am_ohne_ext = am_datei
        if am_datei.upper().endswith(".XML"):
            am_ohne_ext = am_datei[:-4]

        ordner_ohne = (ordner[len("/incoming"):] if ordner.startswith("/incoming")
                       else ordner)
        ordner_mit  = (ordner if ordner.startswith("/incoming")
                       else "/incoming" + ordner)

        # Vollpfad-Varianten (mit/ohne /incoming, mit/ohne .xml)
        varianten += [
            ohne_incoming,                              # /gui/{name}/ASSETMAP.xml (falls primary mit /incoming)
            mit_incoming,                               # /incoming/gui/{name}/ASSETMAP.xml (falls primary ohne)
            ordner_mit  + "/" + am_ohne_ext,            # /incoming/gui/{name}/ASSETMAP
            ordner_ohne + "/" + am_ohne_ext,            # /gui/{name}/ASSETMAP
        ]

        # Ordner-Varianten (ohne Dateiname)
        varianten += [
            ordner_mit  + "/",                          # /incoming/gui/{name}/
            ordner_ohne + "/",                          # /gui/{name}/
            ordner_mit,                                 # /incoming/gui/{name}
            ordner_ohne,                                # /gui/{name}
        ]

        # Relative Varianten (kein führender Slash)
        # DCP-Name extrahieren (letzter Pfadbestandteil des Ordners)
        dcp_name = ordner_ohne.rstrip("/").split("/")[-1]
        if dcp_name:
            varianten += [
                dcp_name + "/" + am_datei,              # {name}/ASSETMAP.xml
                dcp_name + "/" + am_ohne_ext,           # {name}/ASSETMAP
                dcp_name + "/",                         # {name}/
                dcp_name,                               # {name}
            ]

    # Deduplizieren (Reihenfolge beibehalten)
    seen = set()
    result = []
    for v in varianten:
        if v and v not in seen:
            seen.add(v)
            result.append(v)
    return result


def _ingest_add_job_einzel(ip, assetmap_pfad):
    """Einzelner IngestAddJob-Versuch für einen Pfad.
    Gibt job_id (int) zurück oder wirft RuntimeError.
    """
    payload_bytes = assetmap_pfad.encode("utf-8")
    log.info(
        f"[Doremi API] IngestAddJob – Pfad: {assetmap_pfad} "
        f"(hex: {payload_bytes.hex()})"
    )
    nachricht, _ = _baue_nachricht(CMD_INGEST_ADD_JOB, payload_bytes)

    with _verbinde(ip) as sock:
        sock.sendall(nachricht)
        cmd_key, resp_payload = _lese_antwort(sock)

    log.info(f"[Doremi API] IngestAddJob Response (hex): {resp_payload.hex()}")

    if cmd_key != RESP_INGEST_ADD_JOB:
        raise RuntimeError(
            f"Unerwarteter Response-Schlüssel: {cmd_key.hex()} "
            f"(erwartet: {RESP_INGEST_ADD_JOB.hex()})"
        )

    if len(resp_payload) < 8:
        raise RuntimeError(
            f"Antwort zu kurz: {len(resp_payload)} Bytes – "
            f"Rohinhalt: {resp_payload.hex()}"
        )

    job_id = struct.unpack(">q", resp_payload[:8])[0]

    response_code = 0
    if len(resp_payload) >= 12:
        response_code = struct.unpack(">I", resp_payload[8:12])[0]
    elif len(resp_payload) >= 9:
        response_code = resp_payload[8]

    if response_code != 0 or job_id == 0:
        raise RuntimeError(
            f"IngestAddJob Fehlercode={response_code}, job_id={job_id} "
            f"– Rohinhalt: {resp_payload.hex()}"
        )

    return job_id


def ingest_starten(ip, assetmap_pfad, content_uuid=None):
    """
    Startet Ingest auf dem Doremi via TCP API.

    assetmap_pfad: Pfad zur ASSETMAP-Datei auf dem Doremi-Filesystem.
    content_uuid:  Optionale Content-UUID aus der ASSETMAP.xml (urn:uuid:...).
                   Wird als weitere Variante versucht, falls Pfade scheitern.

    Probiert automatisch mehrere Pfad-Varianten + UUID-Formate.
    Gibt die Ingest-Job-ID zurück (int).
    Wirft RuntimeError wenn alle Varianten scheitern.
    """
    varianten = _generiere_pfad_varianten(assetmap_pfad)

    # UUID-Varianten: manche Doremi-Versionen erwarten die Content-UUID
    if content_uuid:
        uuid_clean = content_uuid.strip().replace("urn:uuid:", "")
        varianten += [
            f"urn:uuid:{uuid_clean}",   # vollständige URN
            uuid_clean,                  # nur UUID-String
        ]

    log.info(f"[Doremi API] Versuche {len(varianten)} Varianten: {varianten}")

    letzter_fehler = "Kein Versuch"
    for pfad in varianten:
        try:
            job_id = _ingest_add_job_einzel(ip, pfad)
            log.info(f"[Doremi API] IngestAddJob OK – Variante: '{pfad}', job_id={job_id}")
            return job_id
        except RuntimeError as e:
            log.warning(f"[Doremi API] Variante fehlgeschlagen: '{pfad}' → {e}")
            letzter_fehler = str(e)

    # Fallback: Doremi hat Content möglicherweise automatisch in die Ingest-Queue
    # gestellt (passiert nach FTP-Upload). Suche aktiven Job in job_ids 0..99.
    log.info(
        "[Doremi API] Alle IngestAddJob-Varianten fehlgeschlagen – "
        "suche vorhandenen Ingest-Job (job_ids 0..99)..."
    )
    gefunden = suche_aktive_ingest_job(ip, max_scan=99)
    if gefunden:
        job_id_vorhanden, sc, sn = gefunden
        if sc == 0:  # pending → versuche zu canceln und neu zu starten
            log.info(
                f"[Doremi API] Pending Job job_id={job_id_vorhanden} gefunden – "
                f"versuche zu canceln und IngestAddJob neu zu starten..."
            )
            try:
                import time as _time
                ingest_cancel(ip, job_id_vorhanden)
                _time.sleep(1)
                # Nach Cancel: erneuter IngestAddJob-Versuch (primäre Varianten)
                for pfad in varianten[:6]:
                    try:
                        job_id_neu = _ingest_add_job_einzel(ip, pfad)
                        log.info(
                            f"[Doremi API] IngestAddJob nach Cancel OK – "
                            f"Pfad: '{pfad}', job_id={job_id_neu}"
                        )
                        return job_id_neu
                    except RuntimeError as e:
                        log.debug(f"[Doremi API] Retry '{pfad}': {e}")
                log.warning(
                    "[Doremi API] IngestAddJob nach Cancel weiterhin fehlgeschlagen – "
                    "übernehme job_id vom Cancel (möglicherweise neu vergeben)"
                )
            except Exception as e:
                log.warning(f"[Doremi API] IngestCancelJob fehlgeschlagen: {e}")

        log.info(
            f"[Doremi API] Vorhandenen Ingest-Job übernommen: "
            f"job_id={job_id_vorhanden}, status={sn}({sc})"
        )
        return job_id_vorhanden

    raise RuntimeError(
        f"IngestAddJob fehlgeschlagen für alle {len(varianten)} Varianten "
        f"und kein vorhandener Ingest-Job gefunden. "
        f"Letzter Fehler: {letzter_fehler}"
    )


def suche_aktive_ingest_job(ip, max_scan=99):
    """Sucht aktiven Ingest-Job durch Scannen der job_ids 0..max_scan.

    Hintergrund: Doremi DCP2000 erstellt automatisch Ingest-Jobs wenn Content
    via FTP hochgeladen wird. IngestAddJob schlägt dann mit Fehler 1 fehl
    (Job existiert bereits). Diese Funktion findet den vorhandenen Job.

    Gibt (job_id, status_code, status_name) zurück wenn aktiver Job gefunden,
    sonst None.
    """
    aktive_status = {0, 2, 3}  # pending, running, scheduled
    for job_id_int in range(max_scan + 1):
        try:
            status_code, status_name, progress = ingest_status(ip, job_id_int)
            log.debug(
                f"[Doremi API] Scan job_id={job_id_int}: "
                f"status={status_name}({status_code}), progress={progress}%"
            )
            if status_code in aktive_status:
                log.info(
                    f"[Doremi API] Aktiver Ingest-Job gefunden: "
                    f"job_id={job_id_int}, status={status_name}({status_code})"
                )
                return job_id_int, status_code, status_name
        except (RuntimeError, ConnectionError, OSError, TimeoutError):
            pass  # Ungültige job_id oder Verbindungsfehler → überspringen
    return None


def suche_laufenden_ingest_job(ip, max_scan=20):
    """Sucht einen LAUFENDEN Ingest-Job (status=running/scheduled, NICHT pending).

    Wird im Monitoring verwendet um zu erkennen wenn der Benutzer manuell
    im Doremi-Webinterface auf 'Ingest' geklickt hat.
    job_id=0 mit status=pending wird bewusst ignoriert (Sentinel-Wert).

    Gibt die job_id (int) zurück wenn ein laufender Job gefunden wurde, sonst None.
    """
    for job_id_int in range(1, max_scan + 1):   # Start bei 1, 0 ist Sentinel
        try:
            status_code, status_name, _ = ingest_status(ip, job_id_int)
            if status_code in (2, 3):  # running, scheduled
                log.info(
                    f"[Doremi API] Laufender Ingest-Job gefunden: "
                    f"job_id={job_id_int}, status={status_name}({status_code})"
                )
                return job_id_int
        except (RuntimeError, ConnectionError, OSError, TimeoutError):
            pass
    return None


def ingest_cancel(ip, job_id):
    """
    Bricht einen Ingest-Job ab (IngestCancelJob, Befehl 0x071100).

    Nützlich wenn der Doremi automatisch einen pending-Job erstellt hat
    und IngestAddJob deshalb mit Fehler 1 scheitert.
    Gibt (cmd_key_hex, payload_hex) zurück.
    """
    payload = struct.pack(">q", job_id)
    nachricht, _ = _baue_nachricht(CMD_INGEST_CANCEL_JOB, payload)
    with _verbinde(ip) as sock:
        sock.sendall(nachricht)
        cmd_key, resp_payload = _lese_antwort(sock)
    log.info(
        f"[Doremi API] IngestCancelJob job_id={job_id} – "
        f"key={cmd_key.hex()}, payload={resp_payload.hex()}"
    )
    return cmd_key.hex(), resp_payload.hex()


def ingest_status(ip, job_id):
    """
    Status eines laufenden Ingest-Jobs abfragen.

    Gibt (status_code, status_name, progress_pct) zurück:
      status_code  : int (4 = success, 7 = failed, 5 = aborted, 2 = running)
      status_name  : str (z.B. "running", "success")
      progress_pct : int 0-100 (Prozessfortschritt)
    """
    payload = struct.pack(">q", job_id)   # int64 big-endian
    nachricht, req_id = _baue_nachricht(CMD_INGEST_GET_STATUS, payload)

    with _verbinde(ip) as sock:
        sock.sendall(nachricht)
        cmd_key, resp_payload = _lese_antwort(sock)

    if cmd_key != RESP_INGEST_GET_STATUS:
        raise RuntimeError(
            f"Unerwarteter Response-Schlüssel: {cmd_key.hex()}"
        )

    if len(resp_payload) < 28:
        raise RuntimeError(
            f"IngestGetJobStatus-Antwort zu kurz: {len(resp_payload)} Bytes"
        )

    # Byte-Offsets laut python-dcitools responses.py:
    # 0-4:   error_count
    # 4-8:   warning_count
    # 8-12:  event_count
    # 12-16: status
    # 16-20: download_progress
    # 20-24: process_progress
    # 24-28: actions
    # 28-end: title (text)
    status_code       = struct.unpack(">I", resp_payload[12:16])[0]
    download_progress = struct.unpack(">I", resp_payload[16:20])[0]
    process_progress  = struct.unpack(">I", resp_payload[20:24])[0]

    status_name  = STATUS_NAMEN.get(status_code, f"unbekannt({status_code})")
    progress_pct = max(download_progress, process_progress)

    log.debug(
        f"[Doremi API] IngestGetJobStatus job_id={job_id}: "
        f"status={status_name}({status_code}), progress={progress_pct}%"
    )

    return status_code, status_name, progress_pct
