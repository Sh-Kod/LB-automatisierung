import base64
import json
import os
import py_compile
import re
import shutil
import subprocess
import threading
import time
import unicodedata
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
_naming_aktiv       = False
_naming_aktiv_lock  = threading.Lock()

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
# Naming: Typ-Erkennung + Vorschlag
# ──────────────────────────────────────────────

def _erkenne_typ_und_vorschlag(ocr_text):
    """Erkennt DCP-Typ aus OCR-Text (MEK, ZiK, TK, FK) und erstellt Vorschlag.
    Gibt (typ_string_oder_None, vorschlag_oder_None) zurück."""
    text  = ocr_text.upper()
    monat = datetime.now().strftime("%m")
    tag   = datetime.now().strftime("%d")

    if any(x in text for x in ["MEIN ERSTER KINOBESUCH", "KINOBESUCH", "MEK"]):
        typ, prefix = "MeK – Mein erster Kinobesuch", "LB_MeK"
    elif any(x in text for x in ["ZURÜCK IM KINO", "ZURUCK IM KINO", "ZURCK IM KINO", "ZIK"]):
        typ, prefix = "ZiK – Zurück im Kino", "LB_ZiK"
    elif "TRAUMKINO" in text:
        typ, prefix = "TK – Traumkino", "LB_TK"
    elif "FILMKLASSIKER" in text:
        typ, prefix = "FK – Filmklassiker", "LB_FK"
    else:
        # Fallback: naming_rules.yaml
        try:
            with open(RULES_PFAD, "r", encoding="utf-8") as f:
                rules_data = yaml.safe_load(f)
            for regel in (rules_data or {}).get("regeln", []):
                muster  = regel.get("muster", "")
                vorlage = regel.get("vorlage", "")
                if muster and vorlage and re.search(muster, ocr_text, re.IGNORECASE):
                    return None, vorlage.replace("{JAHR}", str(datetime.now().year))
        except Exception:
            pass
        return None, None

    return typ, f"{prefix}_[FILMNAME]_{monat}_{tag}"


def _bereinige_dcp_name(name):
    """Wandelt Umlaute um und entfernt unerlaubte Zeichen aus DCP-Namen."""
    for orig, repl in [("ä","ae"),("ö","oe"),("ü","ue"),("Ä","Ae"),("Ö","Oe"),("Ü","Ue"),("ß","ss")]:
        name = name.replace(orig, repl)
    name = re.sub(r"[^a-zA-Z0-9_\-]", "_", name)
    return re.sub(r"_+", "_", name).strip("_")

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
    global _naming_aktiv
    time.sleep(5)
    while True:
        # Pause-Check: Bei Pause keinen neuen Naming-Dialog starten
        with _scan_pausiert_lock:
            pausiert = _scan_pausiert
        if pausiert:
            with _naming_aktiv_lock:
                _naming_aktiv = False
            time.sleep(10)
            continue

        item = queue_manager.naechstes_pending()
        if not item:
            with _naming_aktiv_lock:
                _naming_aktiv = False
            time.sleep(10)
            continue

        with _naming_aktiv_lock:
            _naming_aktiv = True

        if not telegram_bot.starte_dialog():
            queue_manager.zuruecksetzen(item["id"])
            time.sleep(5)
            continue

        try:
            bildpfad     = item["bildpfad"]
            dauer_sek    = item["dauer_sek"]
            pending_rest = queue_manager.pending_anzahl()

            # Bild senden
            caption = f"Neues Bild ({dauer_sek}s)"
            if pending_rest > 0:
                caption += f"  |  Queue: {pending_rest} wartend"
            try:
                telegram_bot.sende_bild(bildpfad, caption=caption)
            except Exception:
                telegram_bot.sende_nachricht(f"Bild: {os.path.basename(bildpfad)}")

            # OCR + Typ-Erkennung
            ocr_text = ""
            try:
                ocr_text = analyzer.lese_text_aus_bild(bildpfad)
            except Exception:
                pass

            typ, vorschlag = _erkenne_typ_und_vorschlag(ocr_text)

            # Dialog aufbauen
            t = _trenn()
            msg = f"{t}\n"
            if ocr_text:
                msg += f"OCR-Text:\n{ocr_text[:300].strip()}\n\n"
            if typ:
                msg += f"Typ erkannt: {typ}\n\n"
            if vorschlag:
                msg += f"Vorschlag:\n{vorschlag}\n\n"
                msg += "[1]  Vorschlag übernehmen\n"
                msg += "[2]  Eigenen Namen eingeben\n"
                msg += "[3]  Überspringen\n"
            else:
                msg += "Kein Typ erkannt – bitte Namen eingeben\n\n"
                msg += "[1]  Namen eingeben\n"
                msg += "[2]  Überspringen\n"
            msg += f"{t}\n(Timeout: 60 Min)"
            telegram_bot.sende_nachricht(msg)

            # Auswahl abwarten
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
                    final_name = a  # direkte Namenseingabe
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
                final_name = _bereinige_dcp_name(final_name)
                queue_manager.abschliessen(item["id"])
                job_id = job_manager.erstelle_job(bildpfad, final_name)
                telegram_bot.sende_nachricht(
                    f"Name gesetzt: {final_name}\nVerarbeitung läuft im Hintergrund..."
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

_PHASE_MELDUNGEN = {
    1: "✓ DCP erstellt",
    2: "✓ Hochgeladen",
}

def _ist_solo():
    """True nur wenn: kein Naming läuft, Queue leer, max 1 Job aktiv."""
    with _naming_aktiv_lock:
        if _naming_aktiv:
            return False
    if len(job_manager.hole_aktive()) > 1:
        return False
    return queue_manager.pending_anzahl() == 0

def _phase_ausfuehren(job_id, phase, fn):
    job = job_manager.hole_job(job_id)
    name = (job.get("final_name") or "?") if job else "?"
    phasen_txt = {1: "DCP", 2: "Upload", 3: "Ingest", 4: "Monitor"}
    try:
        job_manager.aktualisiere_phase(job_id, phase, "running")
        fn(job_id)
        job_manager.aktualisiere_phase(job_id, phase, "done")
        if _ist_solo() and phase in _PHASE_MELDUNGEN:
            telegram_bot.sende_nachricht(f"{_PHASE_MELDUNGEN[phase]}\n{name}")
        return True
    except Exception as e:
        fehler_kurz = str(e)[:300]
        job_manager.markiere_fehler(job_id, phase, fehler_kurz, retryable=True)
        telegram_bot.sende_nachricht(
            f"Fehler [{phasen_txt.get(phase,'?')}]: {name}\n"
            f"{fehler_kurz}\n\n"
            f"→ /retry {job_id}"
        )
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

def _bereinige_pfad_fuer_dcp(bildpfad):
    """Kopiert Bild in temp-Ordner mit reinem ASCII-Dateinamen.
    Behebt den dcpomatic2 'codecvt to wstring: error [codecvt:2]' bei Sonderzeichen."""
    temp_dir = r"C:\dcp_automatisierung\temp"
    os.makedirs(temp_dir, exist_ok=True)
    ext = os.path.splitext(bildpfad)[1].lower()
    basename = os.path.splitext(os.path.basename(bildpfad))[0]
    # Umlaute + Sonderzeichen → ASCII
    ascii_name = unicodedata.normalize("NFKD", basename).encode("ascii", "ignore").decode("ascii")
    ascii_name = re.sub(r"[^a-zA-Z0-9_\-]", "_", ascii_name).strip("_")
    if not ascii_name:
        ascii_name = "bild"
    ziel = os.path.join(temp_dir, ascii_name + ext)
    shutil.copy2(bildpfad, ziel)
    return ziel


def _dcp_erstellen(job_id):
    job = job_manager.hole_job(job_id)
    if not job:
        raise RuntimeError(f"Job {job_id} nicht gefunden")

    bildpfad   = job["bildpfad"]
    final_name = job["final_name"]
    cfg        = lade_config()
    ausgabe    = cfg["ordner"]["dcp_ausgabe"]
    create_exe = cfg["dcpomatic"]["create_pfad"]
    cli_exe    = cfg["dcpomatic"]["cli_pfad"]

    # Dauer aus Quellordner ableiten
    bild_lower = bildpfad.replace("\\", "/").lower()
    if "10sec" in bild_lower:
        dauer = 10
    elif "15sec" in bild_lower:
        dauer = 15
    else:
        dauer = 7

    os.makedirs(ausgabe, exist_ok=True)
    tmp_dir = os.path.join(ausgabe, f"_tmp_{job_id}")
    os.makedirs(tmp_dir, exist_ok=True)

    # ASCII-Kopie erstellen – verhindert codecvt-Fehler bei Sonderzeichen im Pfad
    ascii_bildpfad = _bereinige_pfad_fuer_dcp(bildpfad)

    try:
        r1 = subprocess.run(
            [create_exe,
             "--name", final_name,
             "--no-use-isdcf-name",
             "--still-length", str(dauer),
             "--dcp-content-type", "ADV",
             "--twok",
             "--j2k-bandwidth", "100",
             "-o", tmp_dir,
             ascii_bildpfad],   # ASCII-Pfad statt Original
            capture_output=True, text=True, timeout=120
        )
        if r1.returncode != 0:
            raise RuntimeError(f"dcpomatic2_create: {(r1.stderr or r1.stdout)[:400]}")

        # Projekt suchen: In DCP-o-matic 2.x ist das Projekt ein VERZEICHNIS
        # mit .dcpomatic Endung (nicht eine Datei).
        projekt_pfad = None
        for root, dirs, files in os.walk(tmp_dir):
            for d in dirs:
                if d.endswith(".dcpomatic"):
                    projekt_pfad = os.path.join(root, d)
                    break
            if not projekt_pfad:
                for f in files:
                    if f.endswith(".dcpomatic"):
                        projekt_pfad = os.path.join(root, f)
                        break
            if projekt_pfad:
                break
        if not projekt_pfad and os.path.exists(os.path.join(tmp_dir, "metadata.xml")):
            projekt_pfad = tmp_dir
        if not projekt_pfad:
            inhalt = str(os.listdir(tmp_dir))
            raise RuntimeError(f"Kein .dcpomatic Projekt gefunden. tmp_dir: {inhalt}")

        # DCP rendern - dcpomatic2_cli hat kein -o Flag, nimmt nur <FILM>
        r2 = subprocess.run(
            [cli_exe, projekt_pfad],
            capture_output=True, text=True, timeout=7200
        )
        if r2.returncode != 0:
            raise RuntimeError(f"dcpomatic2_cli: {(r2.stderr or r2.stdout)[:400]}")

        # Gerenderter DCP-Ordner suchen (enthaelt .mxf Dateien)
        dcp_subdir = None
        for root, _dirs, files in os.walk(tmp_dir):
            if any(f.lower().endswith(".mxf") for f in files):
                dcp_subdir = root
                break

        if not dcp_subdir:
            raise RuntimeError("Gerenderter DCP-Ordner nicht gefunden")

        ziel = os.path.join(ausgabe, final_name)
        if os.path.exists(ziel):
            shutil.rmtree(ziel)
        shutil.move(dcp_subdir, ziel)

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        # ASCII-Kopie aufräumen
        try:
            os.remove(ascii_bildpfad)
        except Exception:
            pass

    # Originalbild ins Archiv verschieben
    try:
        archiv = cfg["ordner"]["archiv"]
        os.makedirs(archiv, exist_ok=True)
        shutil.move(bildpfad, os.path.join(archiv, os.path.basename(bildpfad)))
    except Exception:
        pass


def _upload_durchfuehren(job_id):
    import ftplib
    job = job_manager.hole_job(job_id)
    if not job:
        raise RuntimeError(f"Job {job_id} nicht gefunden")

    cfg       = lade_config()
    dcp_name  = job["final_name"]
    dcp_pfad  = os.path.join(cfg["ordner"]["dcp_ausgabe"], dcp_name)
    if not os.path.exists(dcp_pfad):
        raise FileNotFoundError(f"DCP-Ordner nicht gefunden: {dcp_pfad}")

    ip     = cfg["doremi"]["ip"]
    user   = cfg["doremi"]["ftp_user"]
    passwd = cfg["doremi"]["ftp_pass"]

    def _upload_dir(ftp, lok_pfad):
        for eintrag in sorted(os.listdir(lok_pfad)):
            lok = os.path.join(lok_pfad, eintrag)
            if os.path.isfile(lok):
                with open(lok, "rb") as f:
                    ftp.storbinary(f"STOR {eintrag}", f)
            elif os.path.isdir(lok):
                try:
                    ftp.mkd(eintrag)
                except ftplib.error_perm:
                    pass
                ftp.cwd(eintrag)
                _upload_dir(ftp, lok)
                ftp.cwd("..")

    with ftplib.FTP(timeout=60) as ftp:
        ftp.connect(ip, 21)
        ftp.login(user, passwd)
        ftp.set_pasv(True)
        # Doremi scannt nur /gui – DCP muss dort abgelegt werden
        ftp.cwd("/gui")
        try:
            ftp.mkd(dcp_name)
        except ftplib.error_perm:
            pass
        ftp.cwd(dcp_name)
        _upload_dir(ftp, dcp_pfad)

    # DCP-Ordner in Upload-Archiv verschieben
    dcp_archiv = cfg["ordner"]["dcp_archiv"]
    os.makedirs(dcp_archiv, exist_ok=True)
    ziel = os.path.join(dcp_archiv, dcp_name)
    if os.path.exists(ziel):
        shutil.rmtree(ziel)
    shutil.move(dcp_pfad, ziel)


def _warte_bis_ftp_bereit(cfg, dcp_name, timeout_min=20):
    """Wartet per FTP-Verify bis DCP auf Doremi /gui vollständig geschrieben ist.
    Gibt True zurück wenn ASSETMAP gefunden, False bei Timeout."""
    import ftplib
    ip     = cfg["doremi"]["ip"]
    user   = cfg["doremi"]["ftp_user"]
    passwd = cfg["doremi"]["ftp_pass"]
    deadline = time.time() + timeout_min * 60
    while time.time() < deadline:
        try:
            with ftplib.FTP(timeout=30) as ftp:
                ftp.connect(ip, 21)
                ftp.login(user, passwd)
                ftp.set_pasv(True)
                ftp.cwd(f"/gui/{dcp_name}")
                dateien = ftp.nlst()
                if any("ASSETMAP" in d.upper() for d in dateien):
                    return True
        except Exception:
            pass
        time.sleep(10)
    return False


def _ftp_ordner_loeschen(cfg, dcp_name):
    """Löscht DCP-Ordner rekursiv von Doremi /gui nach erfolgreichem Ingest."""
    import ftplib

    def rmtree(ftp, pfad):
        eintraege = []
        try:
            ftp.retrlines(f"LIST {pfad}", eintraege.append)
        except Exception:
            return
        for zeile in eintraege:
            teile = zeile.split(None, 8)
            if len(teile) < 9:
                continue
            name = teile[8]
            voll = f"{pfad}/{name}"
            if zeile.startswith("d"):
                rmtree(ftp, voll)
                try:
                    ftp.rmd(voll)
                except Exception:
                    pass
            else:
                try:
                    ftp.delete(voll)
                except Exception:
                    pass
        try:
            ftp.rmd(pfad)
        except Exception:
            pass

    try:
        ip     = cfg["doremi"]["ip"]
        user   = cfg["doremi"]["ftp_user"]
        passwd = cfg["doremi"]["ftp_pass"]
        with ftplib.FTP(timeout=60) as ftp:
            ftp.connect(ip, 21)
            ftp.login(user, passwd)
            ftp.set_pasv(True)
            rmtree(ftp, f"/gui/{dcp_name}")
    except Exception:
        pass  # Best-Effort – kein harter Fehler


def _erstelle_selenium_driver():
    """Erstellt einen Chrome-Headless-Treiber via webdriver_manager."""
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager

    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--log-level=3")
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=opts
    )


def _doremi_login(driver, ip, user, passwd):
    """Setzt HTTP-Basic-Auth-Header via Chrome DevTools Protocol und öffnet die Doremi-Seite."""
    auth_header = "Basic " + base64.b64encode(f"{user}:{passwd}".encode()).decode()
    driver.execute_cdp_cmd("Network.enable", {})
    driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {
        "headers": {"Authorization": auth_header}
    })
    driver.get(f"http://{ip}/web/")
    time.sleep(2)


def _ingest_starten(job_id):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC

    job = job_manager.hole_job(job_id)
    if not job:
        raise RuntimeError(f"Job {job_id} nicht gefunden")

    cfg      = lade_config()
    ip       = cfg["doremi"]["ip"]
    user     = cfg["doremi"]["web_user"]
    passwd   = cfg["doremi"]["web_pass"]
    dcp_name = job["final_name"]

    # FTP-Verify: warten bis DCP wirklich auf Doremi angekommen ist
    bereit = _warte_bis_ftp_bereit(cfg, dcp_name, timeout_min=20)
    if not bereit:
        raise RuntimeError(
            f"DCP '{dcp_name}' nach 20 Min nicht auf Doremi /gui verfügbar. "
            f"FTP-Upload möglicherweise unvollständig."
        )

    ingest_url = (
        f"http://{ip}/web/sys_control/index.php"
        "?page=ingest_manager/ingest_scan.php"
    )

    driver = _erstelle_selenium_driver()
    try:
        _doremi_login(driver, ip, user, passwd)

        gefunden = False
        # 3 Versuche – Doremi kann nach FTP noch kurz brauchen zum Indexieren
        for versuch in range(3):
            driver.get(ingest_url)
            wait = WebDriverWait(driver, 30)

            try:
                dropdown_el = wait.until(
                    EC.presence_of_element_located((By.TAG_NAME, "select"))
                )
                sel = Select(dropdown_el)
                sel.select_by_visible_text("Local Storage")
                time.sleep(8)
                driver.get(ingest_url)  # Seite neu laden damit Liste aktuell ist
                time.sleep(3)
            except Exception:
                time.sleep(5)

            rows = driver.find_elements(By.TAG_NAME, "tr")
            for row in rows:
                if dcp_name.lower() in row.text.lower():
                    btns = row.find_elements(By.TAG_NAME, "button")
                    if not btns:
                        btns = row.find_elements(
                            By.CSS_SELECTOR, "input[type='submit'], input[type='button']"
                        )
                    if btns:
                        btns[0].click()
                        gefunden = True
                        break

            if gefunden:
                break
            if versuch < 2:
                time.sleep(20)

        if not gefunden:
            sichtbar = driver.find_element(By.TAG_NAME, "body").text[:400]
            raise RuntimeError(
                f"DCP '{dcp_name}' nach 3 Versuchen nicht in Ingest-Liste.\n"
                f"Doremi zeigt: {sichtbar}"
            )
        time.sleep(2)
    finally:
        driver.quit()


def _monitoring_ueberwachen(job_id):
    from selenium.webdriver.common.by import By

    job = job_manager.hole_job(job_id)
    if not job:
        raise RuntimeError(f"Job {job_id} nicht gefunden")

    cfg      = lade_config()
    ip       = cfg["doremi"]["ip"]
    user     = cfg["doremi"]["web_user"]
    passwd   = cfg["doremi"]["web_pass"]
    dcp_name = job["final_name"]

    ingest_fertig = False
    driver = _erstelle_selenium_driver()
    try:
        _doremi_login(driver, ip, user, passwd)

        driver.get(
            f"http://{ip}/web/sys_control/index.php"
            "?page=ingest_manager/ingest_monitor.php"
        )

        # Polling – max. 15 Minuten (180 × 5s)
        for _ in range(180):
            time.sleep(5)
            try:
                driver.refresh()
                page_text = driver.find_element(By.TAG_NAME, "body").text.lower()

                # Nicht mehr in der Liste → Ingest abgeschlossen
                if dcp_name.lower() not in page_text:
                    ingest_fertig = True
                    break

                rows = driver.find_elements(By.TAG_NAME, "tr")
                for row in rows:
                    if dcp_name.lower() in row.text.lower():
                        row_text = row.text.lower()
                        row_html = row.get_attribute("innerHTML").lower()
                        if "100%" in row_text or "success" in row_html or "complete" in row_html:
                            ingest_fertig = True
                            break
                        if "error" in row_html or "failed" in row_html:
                            raise RuntimeError(
                                f"Doremi Ingest fehlgeschlagen für '{dcp_name}'"
                            )
                        break
                if ingest_fertig:
                    break
            except RuntimeError:
                raise
            except Exception:
                pass

        # Timeout ohne Fehler – DCP möglicherweise noch am Ingest
    finally:
        driver.quit()

    # FTP aufräumen nach erfolgreichem Ingest
    if ingest_fertig:
        _ftp_ordner_loeschen(cfg, dcp_name)

# ──────────────────────────────────────────────
# Status-Updates + Abschluss
# ──────────────────────────────────────────────

def sende_status_update():
    """Wird alle 15 Min aufgerufen. Sendet kompakten Status wenn Naming fertig
    aber noch Jobs laufen. Sendet Abschlussmeldung wenn alles fertig."""
    with _naming_aktiv_lock:
        naming_laeuft = _naming_aktiv
    if naming_laeuft:
        return  # Während Naming keine Statusmeldungen

    aktive = job_manager.hole_aktive()
    fehler = job_manager.hole_fehler()

    if not aktive and not fehler:
        return  # Nichts zu melden

    t = _trenn()
    phasen_txt = {1: "DCP läuft", 2: "Upload läuft", 3: "Ingest läuft", 4: "Monitor läuft"}

    if aktive:
        nach_phase = {}
        for j in aktive:
            p = phasen_txt.get(j.get("current_phase"), "läuft")
            nach_phase[p] = nach_phase.get(p, 0) + 1
        msg = f"{t}\nStatus\n{t}\n"
        for pname, anzahl in nach_phase.items():
            msg += f"{pname}: {anzahl}\n"
        if fehler:
            msg += f"Fehler: {len(fehler)}  → /jobs\n"
        telegram_bot.sende_nachricht(msg)
    else:
        # Alle fertig (nur Fehler übrig)
        msg = f"{t}\nAbgeschlossen\n{t}\n"
        msg += f"Fehler: {len(fehler)}\n"
        for j in fehler[:5]:
            name = (j.get("final_name") or "?")[:30]
            msg += f"• {name}\n"
        msg += f"\n/jobs für Details  |  /retry_alle"
        telegram_bot.sende_nachricht(msg)


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
        from datetime import timedelta
        # Task Scheduler: updater.py ausserhalb des Service-Prozessbaums starten.
        # Direkte subprocess.Popen wuerde durch NSSM AppKillProcessTree getoetet.
        start_time = (datetime.now() + timedelta(minutes=2)).strftime("%H:%M")
        r1 = subprocess.run([
            "schtasks", "/create", "/f",
            "/tn", "DCP_Automatisierung_Updater",
            "/tr", f'"{sys.executable}" "{UPDATER_PFAD}"',
            "/sc", "once", "/st", start_time, "/ru", "SYSTEM"
        ], capture_output=True, text=True)
        if r1.returncode != 0:
            raise RuntimeError(f"Task erstellen fehlgeschlagen: {r1.stderr.strip()}")
        r2 = subprocess.run([
            "schtasks", "/run", "/tn", "DCP_Automatisierung_Updater"
        ], capture_output=True, text=True)
        if r2.returncode != 0:
            raise RuntimeError(f"Task starten fehlgeschlagen: {r2.stderr.strip()}")

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

    elif low in ("/neustart", "/restart"):
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
            f"/neustart /restart   Service neu starten\n"
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
    job_manager.setze_naming_check(lambda: _naming_aktiv)

    pruefe_update_ergebnis()

    telegram_bot.sende_nachricht(f"DCP-Automatisierung v{lese_version()} gestartet.")

    threading.Thread(
        target=telegram_bot.starte_listener,
        args=(bearbeite_befehl,),
        daemon=True
    ).start()

    threading.Thread(target=queue_worker, daemon=True).start()

    schedule.every(intervall).minutes.do(starte_verarbeitung).tag("scan")
    schedule.every(15).minutes.do(sende_status_update).tag("status")
    if update_intervall > 0:
        schedule.every(update_intervall).hours.do(
            lambda: threading.Thread(target=pruefe_update, daemon=True).start()
        ).tag("update")

    starte_verarbeitung()

    while True:
        schedule.run_pending()
        time.sleep(10)
