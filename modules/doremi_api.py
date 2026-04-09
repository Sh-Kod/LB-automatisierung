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
CMD_INGEST_GET_STATUS = bytes([0x07, 0x1D, 0x00])

# Antwortschlüssel (Response)
RESP_INGEST_ADD_JOB    = bytes([0x07, 0x10, 0x00])
RESP_INGEST_GET_STATUS = bytes([0x07, 0x1E, 0x00])

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

def ingest_starten(ip, assetmap_pfad):
    """
    Startet Ingest auf dem Doremi via TCP API.

    assetmap_pfad: vollständiger Pfad zur ASSETMAP-Datei auf dem Doremi-Filesystem,
                   z.B. '/gui/MEIN_DCP/ASSETMAP.xml'
                   Konfigurierbar über doremi.content_path in config.yaml.

    Gibt die Ingest-Job-ID zurück (int).
    Wirft RuntimeError wenn der Ingest nicht gestartet werden konnte.
    """
    payload_bytes = assetmap_pfad.encode("utf-8")
    log.info(
        f"[Doremi API] IngestAddJob – Pfad: {assetmap_pfad} "
        f"(hex: {payload_bytes.hex()})"
    )

    nachricht, req_id = _baue_nachricht(CMD_INGEST_ADD_JOB, payload_bytes)

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
            f"IngestAddJob-Antwort zu kurz: {len(resp_payload)} Bytes – "
            f"Rohinhalt: {resp_payload.hex()}"
        )

    # Antwort: erste 8 Bytes = job_id (int64 big-endian)
    job_id = struct.unpack(">q", resp_payload[:8])[0]

    # Statusbyte prüfen (Byte 8 oder letztes Byte)
    response_code = 0
    if len(resp_payload) >= 12:
        response_code = struct.unpack(">I", resp_payload[8:12])[0]
    elif len(resp_payload) >= 9:
        response_code = resp_payload[8]

    if response_code != 0:
        raise RuntimeError(
            f"IngestAddJob Fehlercode: {response_code} "
            f"– Rohinhalt: {resp_payload.hex()}"
        )

    log.info(f"[Doremi API] IngestAddJob OK – job_id={job_id}")
    return job_id


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
