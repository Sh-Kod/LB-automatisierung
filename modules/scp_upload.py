"""
SCP-Upload zum Doremi DCP2000
Überträgt DCP-Ordner via SSH/SCP direkt nach /data/incoming/{dcp_name}/

Benötigt: pip install paramiko scp

Hintergrund: Der Doremi hat einen SSH-Server (Standard: doremi/doremi).
SCP-Upload nach /data/incoming/ kann einen anderen auto-ingest-Mechanismus
triggern als FTP-Upload nach /incoming/gui/ – und ermöglicht IngestAddJob
mit Pfad /data/incoming/{name}/ASSETMAP.xml.
"""

import logging
import os

log = logging.getLogger("dcp_automatisierung")


def upload_dcp(ip, dcp_pfad, dcp_name, ssh_user="doremi", ssh_pass="doremi",
               scp_ziel_pfad="/data/incoming"):
    """
    Lädt den lokalen DCP-Ordner via SCP auf den Doremi.

    dcp_pfad       : Lokaler Pfad zum DCP-Ordner (enthält ASSETMAP.xml, .mxf etc.)
    dcp_name       : Name des DCPs (= Ordnername, der auf dem Doremi erstellt wird)
    scp_ziel_pfad  : Basisverzeichnis auf dem Doremi (Standard: /data/incoming)

    Ergebnis auf Doremi: {scp_ziel_pfad}/{dcp_name}/   (mit allen DCP-Dateien darin)

    Wirft RuntimeError bei Verbindungs- oder Übertragungsfehler.
    Wirft RuntimeError wenn paramiko/scp nicht installiert sind.
    """
    try:
        import paramiko
        from scp import SCPClient
    except ImportError as e:
        raise RuntimeError(
            f"paramiko/scp nicht installiert: {e}. "
            f"Bitte 'pip install paramiko scp' auf dem Windows-Server ausführen."
        )

    log.info(f"[SCP] Verbinde zu {ssh_user}@{ip}:22 ...")
    ssh = paramiko.SSHClient()
    # AutoAddPolicy: Host-Key automatisch akzeptieren (kein known_hosts nötig)
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    try:
        ssh.connect(
            hostname=ip,
            port=22,
            username=ssh_user,
            password=ssh_pass,
            timeout=30,
            banner_timeout=30,
            auth_timeout=30,
            allow_agent=False,
            look_for_keys=False,
        )
        log.info(f"[SCP] SSH-Verbindung hergestellt. Starte Upload von '{dcp_name}' ...")

        def _fortschritt(dateiname, groesse, gesendet):
            """Fortschritts-Callback für SCPClient."""
            if groesse > 0:
                pct = int(gesendet / groesse * 100)
                if pct in (25, 50, 75, 100):
                    log.debug(
                        f"[SCP] {os.path.basename(str(dateiname))}: {pct}%  "
                        f"({gesendet}/{groesse} Bytes)"
                    )

        with SCPClient(ssh.get_transport(), progress=_fortschritt) as scp:
            # scp.put mit recursive=True:
            #   Quelle:  C:\dcp_ausgabe\{dcp_name}\  (lokaler Windows-Ordner)
            #   Ziel:    /data/incoming/              (remote-Basisverzeichnis)
            #   Ergebnis: /data/incoming/{dcp_name}/  (Ordnername wird beibehalten)
            scp.put(dcp_pfad, remote_path=scp_ziel_pfad, recursive=True)

        log.info(
            f"[SCP] Upload abgeschlossen: '{dcp_name}' → {scp_ziel_pfad}/{dcp_name}/"
        )

    except Exception as e:
        log.error(f"[SCP] Upload fehlgeschlagen: {e}")
        raise RuntimeError(f"SCP-Upload fehlgeschlagen: {e}")
    finally:
        try:
            ssh.close()
        except Exception:
            pass


def teste_verbindung(ip, ssh_user="doremi", ssh_pass="doremi"):
    """
    Testet SSH-Verbindung zum Doremi.
    Gibt (True, info_str) zurück bei Erfolg, (False, fehler_str) bei Fehler.
    """
    try:
        import paramiko
    except ImportError:
        return False, "paramiko nicht installiert (pip install paramiko scp)"

    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        ssh.connect(
            hostname=ip,
            port=22,
            username=ssh_user,
            password=ssh_pass,
            timeout=15,
            banner_timeout=15,
            allow_agent=False,
            look_for_keys=False,
        )
        # Einfache Diagnose: Verzeichnis auflisten
        stdin, stdout, stderr = ssh.exec_command("ls /data/incoming/ 2>/dev/null || echo '(leer)'")
        stdout.channel.recv_exit_status()
        ausgabe = stdout.read().decode("utf-8", errors="replace").strip()
        ssh.close()
        return True, f"SSH OK. /data/incoming Inhalt:\n{ausgabe[:300]}"
    except Exception as e:
        try:
            ssh.close()
        except Exception:
            pass
        return False, str(e)
