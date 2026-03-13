# ============================================================
#   DCP AUTOMATISIERUNG - INSTALLER v1.0
#   Ausfuehren mit:
#   powershell -ExecutionPolicy Bypass -File install.ps1
# ============================================================

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "   DCP AUTOMATISIERUNG - INSTALLER v1.0" -ForegroundColor Cyan
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""

# Admin-Check
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[FEHLER] Bitte als Administrator ausfuehren!" -ForegroundColor Red
    Write-Host "Rechtsklick auf PowerShell dann 'Als Administrator ausfuehren'" -ForegroundColor Yellow
    Read-Host "Druecke Enter zum Beenden"
    exit 1
}

# ─── EINGABEN ───────────────────────────────────────────────
Write-Host "Auf welchem Laufwerk sollen die Ordner erstellt werden?" -ForegroundColor Yellow
Write-Host "Beispiel: E fuer E:\ oder C fuer C:\" -ForegroundColor Gray
$LAUFWERK = Read-Host "Laufwerk eingeben (Standard E)"
if ($LAUFWERK -eq "") { $LAUFWERK = "E" }
$LAUFWERK = $LAUFWERK.Replace(":", "")

Write-Host ""
Write-Host "Welche Doremi IP-Adresse?" -ForegroundColor Yellow
Write-Host "Kino 3 = 172.20.23.11" -ForegroundColor Gray
$DOREMI_IP = Read-Host "Doremi IP eingeben (Standard 172.20.23.11)"
if ($DOREMI_IP -eq "") { $DOREMI_IP = "172.20.23.11" }

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host "   Einstellungen:" -ForegroundColor Cyan
Write-Host "     Laufwerk  : ${LAUFWERK}:\" -ForegroundColor White
Write-Host "     Doremi IP : $DOREMI_IP" -ForegroundColor White
Write-Host "     Programm  : C:\dcp_automatisierung\" -ForegroundColor White
Write-Host "  ============================================================" -ForegroundColor Cyan
Write-Host ""
$OK = Read-Host "Installation starten? J/N"
if ($OK -eq "N" -or $OK -eq "n") { exit 0 }
Write-Host ""

# ─── SCHRITT 1: PYTHON ──────────────────────────────────────
Write-Host "[1/7] Pruefe Python..." -ForegroundColor Green
try {
    $v = python --version 2>&1
    Write-Host "      $v - OK" -ForegroundColor Gray
} catch {
    Write-Host "      Python nicht gefunden - wird installiert..." -ForegroundColor Yellow
    winget install --id Python.Python.3.11 --silent --accept-source-agreements --accept-package-agreements
    $env:PATH += ";C:\Users\$env:USERNAME\AppData\Local\Programs\Python\Python311"
    $env:PATH += ";C:\Users\$env:USERNAME\AppData\Local\Programs\Python\Python311\Scripts"
    Write-Host "      Python installiert!" -ForegroundColor Gray
}

# ─── SCHRITT 2: TESSERACT ───────────────────────────────────
Write-Host "[2/7] Pruefe Tesseract OCR..." -ForegroundColor Green
if (Test-Path "C:\Program Files\Tesseract-OCR\tesseract.exe") {
    Write-Host "      Tesseract bereits installiert - OK" -ForegroundColor Gray
} else {
    Write-Host "      Tesseract wird installiert..." -ForegroundColor Yellow
    winget install --id UB-Mannheim.TesseractOCR --silent --accept-source-agreements --accept-package-agreements --source winget
    Write-Host "      Tesseract installiert!" -ForegroundColor Gray
}

# ─── SCHRITT 3: CHROME ──────────────────────────────────────
Write-Host "[3/7] Pruefe Google Chrome..." -ForegroundColor Green
if ((Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") -or
    (Test-Path "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")) {
    Write-Host "      Chrome bereits installiert - OK" -ForegroundColor Gray
} else {
    Write-Host "      Chrome wird installiert..." -ForegroundColor Yellow
    winget install --id Google.Chrome --silent --accept-source-agreements --accept-package-agreements
    Write-Host "      Chrome installiert!" -ForegroundColor Gray
}

# ─── SCHRITT 4: NSSM ────────────────────────────────────────
Write-Host "[4/7] Pruefe NSSM..." -ForegroundColor Green
if (Test-Path "C:\nssm\nssm.exe") {
    Write-Host "      NSSM bereits installiert - OK" -ForegroundColor Gray
} else {
    Write-Host "      NSSM wird installiert..." -ForegroundColor Yellow
    curl.exe -L --retry 3 --retry-delay 2 -o "$env:TEMP\nssm.zip" "https://nssm.cc/release/nssm-2.24.zip"
    Expand-Archive -Path "$env:TEMP\nssm.zip" -DestinationPath "$env:TEMP\nssm_tmp" -Force
    New-Item -ItemType Directory -Path "C:\nssm" -Force | Out-Null
    Copy-Item "$env:TEMP\nssm_tmp\nssm-2.24\win64\nssm.exe" "C:\nssm\nssm.exe" -Force
    $machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
    [Environment]::SetEnvironmentVariable("PATH", "$machinePath;C:\nssm", "Machine")
    $env:PATH += ";C:\nssm"
    Write-Host "      NSSM installiert!" -ForegroundColor Gray
}

# ─── SCHRITT 5: ORDNER ──────────────────────────────────────
Write-Host "[5/7] Erstelle Ordner..." -ForegroundColor Green
$ordner = @(
    "C:\dcp_automatisierung"
    "C:\dcp_automatisierung\modules"
    "C:\dcp_automatisierung\rules"
    "C:\dcp_automatisierung\logs"
    "C:\dcp_automatisierung\temp"
    "${LAUFWERK}:\K.O.D Atomations\Neue LB"
    "${LAUFWERK}:\K.O.D Atomations\Neue LB 10sec"
    "${LAUFWERK}:\K.O.D Atomations\Neue LB 15sec"
    "${LAUFWERK}:\K.O.D Atomations\DCP"
    "${LAUFWERK}:\K.O.D Atomations\AUF TMS"
    "${LAUFWERK}:\K.O.D Atomations\Fehler"
    "${LAUFWERK}:\K.O.D Atomations\DCP Upload erledigt"
)
foreach ($o in $ordner) { New-Item -ItemType Directory -Path $o -Force | Out-Null }
Write-Host "      Alle Ordner erstellt - OK" -ForegroundColor Gray

# ─── SCHRITT 6: PYTHON DATEIEN ──────────────────────────────
Write-Host "[6/7] Erstelle Konfiguration und Scripts..." -ForegroundColor Green

# config.yaml
@"
ordner:
  eingang_7sec: "${LAUFWERK}:\\K.O.D Atomations\\Neue LB"
  eingang_10sec: "${LAUFWERK}:\\K.O.D Atomations\\Neue LB 10sec"
  eingang_15sec: "${LAUFWERK}:\\K.O.D Atomations\\Neue LB 15sec"
  dcp_ausgabe: "${LAUFWERK}:\\K.O.D Atomations\\DCP"
  archiv: "${LAUFWERK}:\\K.O.D Atomations\\AUF TMS"
  fehler: "${LAUFWERK}:\\K.O.D Atomations\\Fehler"
  dcp_archiv: "${LAUFWERK}:\\K.O.D Atomations\\DCP Upload erledigt"
dcpomatic:
  cli_pfad: "C:\\Program Files\\DCP-o-matic 2\\bin\\dcpomatic2_cli.exe"
zeitplan:
  intervall_minuten: 60
telegram:
  token: "8655165819:AAFqrEPOO8OGCR3jHBOoFe9vflceVYTfpAc"
  chat_id: "479976191"
logging:
  log_datei: "C:\\dcp_automatisierung\\logs\\dcp_system.log"
  log_level: "INFO"
gemini:
  api_key: "AIzaSyAcxSYME3T3hQNy7vK3wMdENQZuS4RXGzc"
tesseract:
  pfad: "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
doremi:
  ip: "$DOREMI_IP"
  ftp_user: "ingest"
  ftp_pass: "ingest"
  web_user: "admin"
  web_pass: "1234"
"@ | Set-Content -Path "C:\dcp_automatisierung\config.yaml" -Encoding UTF8

# naming_rules.yaml
"regeln: []" | Set-Content -Path "C:\dcp_automatisierung\rules\naming_rules.yaml" -Encoding UTF8

# __init__.py
"" | Set-Content -Path "C:\dcp_automatisierung\modules\__init__.py" -Encoding UTF8

# ── modules/logger.py ────────────────────────────────────────
@'
import logging
import yaml
from pathlib import Path

def erstelle_logger():
    with open("C:\\dcp_automatisierung\\config.yaml", "r") as f:
        config = yaml.safe_load(f)
    log_datei = config["logging"]["log_datei"]
    log_level = config["logging"].get("log_level", "INFO")
    Path(log_datei).parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=getattr(logging, log_level),
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_datei, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("dcp_automatisierung")

logger = erstelle_logger()
'@ | Set-Content -Path "C:\dcp_automatisierung\modules\logger.py" -Encoding UTF8

# ── modules/watcher.py ───────────────────────────────────────
@'
import os
from pathlib import Path

ERLAUBTE_ENDUNGEN = [".jpg", ".jpeg", ".png"]

def suche_neue_bilder(ordner):
    bilder = []
    if not os.path.exists(ordner):
        return bilder
    for datei in os.listdir(ordner):
        if Path(datei).suffix.lower() in ERLAUBTE_ENDUNGEN:
            bilder.append(os.path.join(ordner, datei))
    return bilder
'@ | Set-Content -Path "C:\dcp_automatisierung\modules\watcher.py" -Encoding UTF8

# ── modules/analyzer.py ──────────────────────────────────────
@'
import yaml
from PIL import Image
import pytesseract

def lade_config():
    with open("C:\\dcp_automatisierung\\config.yaml", "r") as f:
        return yaml.safe_load(f)

def lese_text_aus_bild(bildpfad):
    try:
        config = lade_config()
        pytesseract.pytesseract.tesseract_cmd = config["tesseract"]["pfad"]
        img = Image.open(bildpfad)
        text = pytesseract.image_to_string(img, lang="deu+eng")
        return text.strip()
    except Exception as e:
        print(f"OCR Fehler: {e}")
        return ""
'@ | Set-Content -Path "C:\dcp_automatisierung\modules\analyzer.py" -Encoding UTF8

# ── modules/telegram_bot.py ──────────────────────────────────
@'
import requests
import yaml
import time
import threading

def lade_config():
    with open("C:\\dcp_automatisierung\\config.yaml", "r") as f:
        return yaml.safe_load(f)

def sende_nachricht(text):
    try:
        config = lade_config()
        token = config["telegram"]["token"]
        chat_id = config["telegram"]["chat_id"]
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        requests.post(url, data={"chat_id": chat_id, "text": text}, timeout=10)
        print("Telegram Nachricht gesendet!")
    except Exception as e:
        print(f"Telegram Fehler: {e}")

def warte_auf_antwort(timeout=300):
    try:
        config = lade_config()
        token = config["telegram"]["token"]
        chat_id = str(config["telegram"]["chat_id"])
        url = f"https://api.telegram.org/bot{token}/getUpdates"
        response = requests.get(url, params={"timeout": 0}, timeout=10)
        updates = response.json().get("result", [])
        last_id = max([u["update_id"] for u in updates], default=0) if updates else 0
        start = time.time()
        while time.time() - start < timeout:
            response = requests.get(url, params={"offset": last_id + 1, "timeout": 5}, timeout=10)
            updates = response.json().get("result", [])
            for update in updates:
                last_id = update["update_id"]
                msg = update.get("message", {})
                if str(msg.get("chat", {}).get("id", "")) == chat_id:
                    return msg.get("text", "").strip()
        return None
    except Exception as e:
        print(f"Telegram Warten Fehler: {e}")
        return None

_last_update_id = 0

def starte_listener(callback):
    global _last_update_id
    try:
        config = lade_config()
        token = config["telegram"]["token"]
        url = f"https://api.telegram.org/bot{token}/getUpdates"
        response = requests.get(url, params={"timeout": 0}, timeout=10)
        updates = response.json().get("result", [])
        _last_update_id = max([u["update_id"] for u in updates], default=0) if updates else 0
    except:
        pass
    while True:
        try:
            config = lade_config()
            token = config["telegram"]["token"]
            chat_id = str(config["telegram"]["chat_id"])
            url = f"https://api.telegram.org/bot{token}/getUpdates"
            response = requests.get(url, params={"offset": _last_update_id + 1, "timeout": 10}, timeout=15)
            updates = response.json().get("result", [])
            for update in updates:
                _last_update_id = update["update_id"]
                msg = update.get("message", {})
                if str(msg.get("chat", {}).get("id", "")) == chat_id:
                    text = msg.get("text", "").strip()
                    if text and callback:
                        threading.Thread(target=callback, args=(text,), daemon=True).start()
        except:
            time.sleep(5)
'@ | Set-Content -Path "C:\dcp_automatisierung\modules\telegram_bot.py" -Encoding UTF8

# ── modules/naming.py ────────────────────────────────────────
@'
import os
import re
import yaml
from pathlib import Path
from datetime import datetime
from modules.telegram_bot import sende_nachricht, warte_auf_antwort

REGELN_PFAD = "C:\\dcp_automatisierung\\rules\\naming_rules.yaml"

def bereinige_name(name):
    ersetzungen = {
        "ae": ["ae"], "oe": ["oe"], "ue": ["ue"],
        "\u00e4": "ae", "\u00f6": "oe", "\u00fc": "ue",
        "\u00c4": "Ae", "\u00d6": "Oe", "\u00dc": "Ue",
        "\u00df": "ss", "+": "_", " ": "_", "-": "_"
    }
    for alt, neu in [("\u00e4","ae"),("\u00f6","oe"),("\u00fc","ue"),("\u00c4","Ae"),
                     ("\u00d6","Oe"),("\u00dc","Ue"),("\u00df","ss"),("+","_"),(" ","_"),("-","_")]:
        name = name.replace(alt, neu)
    name = re.sub(r"[^a-zA-Z0-9_]", "", name)
    name = re.sub(r"_+", "_", name)
    return name.strip("_")

def extrahiere_daten(text):
    muster = r"(\d{1,2})\.(\d{1,2})\.(\d{2,4})"
    gefunden = []
    for t in re.findall(muster, text or ""):
        try:
            tag, monat, jahr = int(t[0]), int(t[1]), int(t[2])
            if len(t[2]) == 2:
                jahr = 2000 + jahr
            if 1 <= tag <= 31 and 1 <= monat <= 12:
                gefunden.append((tag, monat, jahr))
        except:
            pass
    return sorted(set(gefunden))

def datum_zu_string(tag, monat, jahr):
    return f"{tag:02d}_{monat:02d}_{str(jahr)[-2:]}"

def lade_regeln():
    if os.path.exists(REGELN_PFAD):
        with open(REGELN_PFAD, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
            return data.get("regeln", [])
    return []

def speichere_regel(muster, dcp_name):
    regeln = lade_regeln()
    regeln.append({"muster": muster, "name": dcp_name})
    with open(REGELN_PFAD, "w", encoding="utf-8") as f:
        yaml.dump({"regeln": regeln}, f, allow_unicode=True)

def pruefe_gespeicherte_regeln(dateiname):
    for regel in lade_regeln():
        if regel.get("muster", "").lower() in dateiname.lower():
            return regel.get("name")
    return None

def erkenne_typ(dateiname, ocr_text):
    dl = dateiname.lower()
    ou = (ocr_text or "").upper()
    if "MEIN ERSTER KINOBESUCH" in ou:
        return "mek_uebersicht" if "vorschau" in dl else "mek_einzelfilm"
    if "TRAUMKINO" in ou or "traumkino" in dl:
        return "traumkino"
    if "FILMKLASSIKER" in ou:
        return "filmklassiker"
    if "ZUR" in ou and "CK IM KINO" in ou or "_zik_" in dl:
        daten = extrahiere_daten(ocr_text)
        return "zik_uebersicht" if "_zik_" in dl or len(daten) > 3 else "zik_einzelfilm"
    if "_lb_" in dl:
        return "regulaer"
    return "unbekannt"

def erstelle_vorschlag(typ, dateiname, ocr_text):
    stem = Path(dateiname).stem
    daten = extrahiere_daten(ocr_text or "")
    if typ == "mek_uebersicht":
        monate = re.findall(r"(Januar|Februar|M.rz|April|Mai|Juni|Juli|August|September|Oktober|November|Dezember)", ocr_text or "", re.IGNORECASE)
        if len(monate) >= 2:
            return f"LB_MeK_uebersicht_{bereinige_name(monate[0])}_{bereinige_name(monate[-1])}_{datetime.now().strftime('%y')}"
        return "LB_MeK_uebersicht"
    elif typ == "mek_einzelfilm":
        teile = stem.split("_"); film = bereinige_name("_".join(teile[2:]) if len(teile) > 2 else teile[-1])
        return f"LB_MeK_{film}_{datum_zu_string(*daten[-1])}" if daten else f"LB_MeK_{film}"
    elif typ == "traumkino":
        if len(daten) >= 2:
            t1,m1,_ = daten[0]; t2,m2,_ = daten[-1]
            return f"LB_Traumkino_{t1:02d}_{m1:02d}_bis_{t2:02d}_{m2:02d}"
        elif daten:
            t,m,_ = daten[0]; return f"LB_Traumkino_{t:02d}_{m:02d}"
        return "LB_Traumkino"
    elif typ == "filmklassiker":
        for z in (ocr_text or "").split("\n"):
            if z.strip() and "FILMKLASSIKER" not in z.upper():
                f = bereinige_name(z.strip())
                if f: return f"LB_FK_{f}"
        return "LB_FK"
    elif typ == "zik_einzelfilm":
        teile = stem.split("_"); film = bereinige_name("_".join(teile[2:]) if len(teile) > 2 else teile[-1])
        return f"LB_ZiK_{film}_{datum_zu_string(*daten[-1])}" if daten else f"LB_ZiK_{film}"
    elif typ == "zik_uebersicht":
        if len(daten) >= 2:
            t1,m1,j1 = daten[0]; t2,m2,j2 = daten[-1]
            return f"LB_ZiK_{t1:02d}_{m1:02d}_bis_{t2:02d}_{m2:02d}_{str(j2)[-2:]}"
        return "LB_ZiK_Uebersicht"
    elif typ == "regulaer":
        teile = stem.split("_lb_") if "_lb_" in stem.lower() else [stem]
        film = bereinige_name(teile[-1])
        if len(daten) == 1:
            return f"LB_{film}_{datum_zu_string(*daten[0])}"
        elif len(daten) > 1:
            t1,m1,j1 = daten[0]; t2,m2,j2 = daten[-1]
            return f"LB_{film}_{t1:02d}_{m1:02d}_bis_{t2:02d}_{m2:02d}_{str(j2)[-2:]}"
        return f"LB_{film}"
    return None

def bestimme_dcp_name_komplett(bildpfad, ocr_text):
    dateiname = Path(bildpfad).name
    stem = Path(bildpfad).stem

    gespeichert = pruefe_gespeicherte_regeln(dateiname)
    if gespeichert:
        print(f"Gespeicherte Regel: {gespeichert}")
        return gespeichert

    typ = erkenne_typ(dateiname, ocr_text)
    print(f"Erkannter Typ: {typ}")
    vorschlag = erstelle_vorschlag(typ, dateiname, ocr_text)

    daten = extrahiere_daten(ocr_text or "")
    daten_text = ""
    if daten:
        daten_text = "\nDaten: " + ", ".join([f"{t:02d}.{m:02d}.{j}" for t,m,j in daten])

    def name_aus_datei():
        teile = stem.split("_lb_") if "_lb_" in stem.lower() else [stem]
        return f"LB_{bereinige_name(teile[-1])}"

    if vorschlag:
        sende_nachricht(
            f"\U0001f5bc Neues Bild: {dateiname}\n"
            f"\U0001f50d Typ: {typ}{daten_text}\n\n"
            f"\U0001f4a1 Vorschlag: {vorschlag}\n\n"
            f"1 \u2192 Vorschlag \u00fcbernehmen\n"
            f"2 \u2192 Eigenen Namen eingeben\n"
            f"3 \u2192 Kein Datum\n"
            f"4 \u2192 \u00dcberspringen"
        )
        antwort = warte_auf_antwort(timeout=300)
        if antwort in ["1", "/1"]:
            return vorschlag
        elif antwort in ["2", "/2"]:
            sende_nachricht("\u270f Namen eingeben:")
            name = warte_auf_antwort(timeout=300)
            if name:
                n = bereinige_name(name)
                speichere_regel(stem[:20], n)
                return n
        elif antwort in ["3", "/3"]:
            return name_aus_datei()
        elif antwort in ["4", "/4"]:
            return None
        else:
            return vorschlag
    else:
        sende_nachricht(
            f"\U0001f5bc Neues Bild: {dateiname}\n"
            f"\u2753 Typ unbekannt{daten_text}\n\n"
            f"1 \u2192 Namen eingeben\n"
            f"2 \u2192 Kein Datum (aus Dateiname)\n"
            f"3 \u2192 \u00dcberspringen"
        )
        antwort = warte_auf_antwort(timeout=300)
        if antwort in ["1", "/1"]:
            sende_nachricht("\u270f Namen eingeben:")
            name = warte_auf_antwort(timeout=300)
            if name:
                n = bereinige_name(name)
                speichere_regel(stem[:20], n)
                return n
        elif antwort in ["2", "/2"]:
            return name_aus_datei()
        elif antwort in ["3", "/3"]:
            return None
    return vorschlag
'@ | Set-Content -Path "C:\dcp_automatisierung\modules\naming.py" -Encoding UTF8

# ── modules/dcpomatic.py ─────────────────────────────────────
@'
import os
import subprocess
import shutil
import yaml

def lade_config():
    with open("C:\\dcp_automatisierung\\config.yaml", "r") as f:
        return yaml.safe_load(f)

def erstelle_dcp(bildpfad, dcp_name, eingangsordner):
    from modules.telegram_bot import sende_nachricht
    config = lade_config()
    cli_pfad = config["dcpomatic"]["cli_pfad"]
    create_exe = cli_pfad.replace("dcpomatic2_cli.exe", "dcpomatic2_create.exe")
    dcp_ausgabe = config["ordner"]["dcp_ausgabe"]
    temp_ordner = "C:\\dcp_automatisierung\\temp"
    os.makedirs(temp_ordner, exist_ok=True)
    os.makedirs(dcp_ausgabe, exist_ok=True)
    laenge = 15 if "15sec" in eingangsordner else 10 if "10sec" in eingangsordner else 7
    projekt_ordner = os.path.join(temp_ordner, dcp_name)
    os.makedirs(projekt_ordner, exist_ok=True)
    try:
        sende_nachricht(f"Erstelle DCP:\n{dcp_name}\nLaenge: {laenge} Sekunden")
        subprocess.run([create_exe, "--name", dcp_name, "--content", bildpfad,
            "--still-length", str(laenge), "--dcp-content-type", "ADV",
            "--output", dcp_ausgabe, projekt_ordner],
            capture_output=True, text=True, timeout=120)
        subprocess.run([cli_pfad, projekt_ordner],
            capture_output=True, text=True, timeout=1800)
        if os.path.exists(os.path.join(dcp_ausgabe, dcp_name)):
            sende_nachricht(f"DCP erstellt!\n{dcp_name}")
            return True
        sende_nachricht(f"DCP Fehler!\n{dcp_name}")
        return False
    except Exception as e:
        print(f"DCP Fehler: {e}")
        sende_nachricht(f"DCP Fehler!\n{dcp_name}\n{str(e)[:200]}")
        return False
    finally:
        shutil.rmtree(projekt_ordner, ignore_errors=True)

def verschiebe_in_archiv(bildpfad):
    config = lade_config()
    archiv = config["ordner"]["archiv"]
    os.makedirs(archiv, exist_ok=True)
    ziel = os.path.join(archiv, os.path.basename(bildpfad))
    shutil.move(bildpfad, ziel)
    return ziel
'@ | Set-Content -Path "C:\dcp_automatisierung\modules\dcpomatic.py" -Encoding UTF8

# ── modules/doremi.py ────────────────────────────────────────
@'
import os
import time
import ftplib
import shutil
import yaml
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

LOG_ORDNER = "C:\\dcp_automatisierung\\logs"

def lade_config():
    with open("C:\\dcp_automatisierung\\config.yaml", "r") as f:
        return yaml.safe_load(f)

def hole_cfg():
    config = lade_config()
    d = config.get("doremi", {})
    ip = d.get("ip", "172.20.23.11")
    return {
        "ip": ip,
        "ftp_user": d.get("ftp_user", "ingest"),
        "ftp_pass": d.get("ftp_pass", "ingest"),
        "ftp_ordner": "/gui",
        "web_user": d.get("web_user", "admin"),
        "web_pass": d.get("web_pass", "1234"),
        "web_url": f"http://{ip}/web",
        "scan_url": f"http://{ip}/web/sys_control/index.php?page=ingest_manager/ingest_scan.php",
        "monitor_url": f"http://{ip}/web/sys_control/index.php?page=ingest_manager/ingest_monitor.php",
    }

def erstelle_driver():
    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--ignore-certificate-errors")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

def login(driver, cfg):
    print("Login auf Doremi...")
    driver.get(f"{cfg['web_url']}/")
    time.sleep(3)
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    ok = False
    for iframe in iframes:
        try:
            driver.switch_to.frame(iframe)
            felder = driver.find_elements(By.NAME, "username")
            if felder:
                felder[0].clear(); felder[0].send_keys(cfg["web_user"])
                driver.find_element(By.NAME, "password").send_keys(cfg["web_pass"])
                driver.find_element(By.XPATH, "//input[@value='Login']|//button[contains(text(),'Login')]").click()
                ok = True; break
            driver.switch_to.default_content()
        except:
            driver.switch_to.default_content()
    if not ok:
        driver.switch_to.default_content()
        wait = WebDriverWait(driver, 20)
        wait.until(lambda d: d.find_element(By.NAME, "username")).send_keys(cfg["web_user"])
        driver.find_element(By.NAME, "password").send_keys(cfg["web_pass"])
        driver.find_element(By.XPATH, "//input[@value='Login']|//button[contains(text(),'Login')]").click()
    time.sleep(3)
    driver.switch_to.default_content()
    print("Login erfolgreich!")

def finde_select(driver):
    s = driver.find_elements(By.TAG_NAME, "select")
    if s: return s[0], None
    for i, iframe in enumerate(driver.find_elements(By.TAG_NAME, "iframe")):
        try:
            driver.switch_to.frame(iframe)
            s = driver.find_elements(By.TAG_NAME, "select")
            if s: return s[0], i
            driver.switch_to.default_content()
        except:
            driver.switch_to.default_content()
    return None, None

def ftp_loesche_dcp(dcp_name):
    cfg = hole_cfg()
    try:
        ftp = ftplib.FTP(); ftp.connect(cfg["ip"], 21, timeout=30)
        ftp.login(cfg["ftp_user"], cfg["ftp_pass"]); ftp.set_pasv(True)
        ftp.cwd(cfg["ftp_ordner"])
        try:
            ftp.cwd(dcp_name)
            for d in ftp.nlst():
                try: ftp.delete(d)
                except: pass
            ftp.cwd(".."); ftp.rmd(dcp_name)
        except: pass
        ftp.quit()
        print(f"DCP vom FTP geloescht: {dcp_name}")
        return True
    except Exception as e:
        print(f"FTP Loeschen Fehler: {e}")
        return False

def ftp_upload(dcp_ordner_pfad, dcp_name):
    from modules.telegram_bot import sende_nachricht
    cfg = hole_cfg()
    print(f"FTP Upload: {dcp_name}")
    try:
        ftp = ftplib.FTP(); ftp.connect(cfg["ip"], 21, timeout=30)
        ftp.login(cfg["ftp_user"], cfg["ftp_pass"]); ftp.set_pasv(True)
        ftp.cwd(cfg["ftp_ordner"])
        try: ftp.mkd(dcp_name)
        except ftplib.error_perm: pass
        ftp.cwd(dcp_name)
        dateien = [d for d in os.listdir(dcp_ordner_pfad) if os.path.isfile(os.path.join(dcp_ordner_pfad, d))]
        gesamt = len(dateien)
        sende_nachricht(f"FTP Upload gestartet!\nDCP: {dcp_name}\n{gesamt} Dateien...")
        for i, datei in enumerate(dateien, 1):
            pfad = os.path.join(dcp_ordner_pfad, datei)
            mb = os.path.getsize(pfad) / (1024*1024)
            print(f"[{i}/{gesamt}] {datei} ({mb:.1f} MB)...")
            with open(pfad, "rb") as f:
                ftp.storbinary(f"STOR {datei}", f, blocksize=8192)
        ftp.quit()
        sende_nachricht(f"FTP Upload fertig!\nDCP: {dcp_name}\nStarte Ingest...")
        return True
    except Exception as e:
        sende_nachricht(f"FTP Fehler!\nDCP: {dcp_name}\n{str(e)[:300]}")
        return False

def starte_ingest(dcp_name):
    from modules.telegram_bot import sende_nachricht
    cfg = hole_cfg()
    driver = None
    try:
        driver = erstelle_driver()
        os.makedirs(LOG_ORDNER, exist_ok=True)
        login(driver, cfg)
        driver.get(cfg["scan_url"]); time.sleep(5)
        sel, _ = finde_select(driver)
        if not sel:
            sende_nachricht(f"Ingest Fehler!\nDCP: {dcp_name}\nDropdown nicht gefunden!")
            return False
        Select(sel).select_by_visible_text("Local Storage"); time.sleep(4)
        if dcp_name not in driver.page_source:
            sende_nachricht(f"DCP nicht gefunden!\n{dcp_name}"); return False
        driver.execute_script("""
            var cbs=document.querySelectorAll('input[type=checkbox]');
            for(var i=0;i<cbs.length;i++){if(cbs[i].id!=='pingestCheck'&&!cbs[i].checked)cbs[i].click();}
        """)
        time.sleep(2)
        btn = driver.find_element(By.ID, "ingestBtn")
        driver.execute_script("arguments[0].click();", btn); time.sleep(2)
        try:
            a = driver.switch_to.alert; a.dismiss()
        except: pass
        driver.switch_to.default_content()
        sende_nachricht(f"Ingest gestartet!\nDCP: {dcp_name}\nUeberwache Fortschritt...")
        return True
    except Exception as e:
        if driver: driver.save_screenshot(f"{LOG_ORDNER}\\doremi_fehler.png")
        sende_nachricht(f"Ingest Fehler!\nDCP: {dcp_name}\n{str(e)[:300]}")
        return False
    finally:
        if driver: driver.quit()

def ueberwache_ingest(dcp_name):
    cfg = hole_cfg()
    driver = None
    try:
        driver = erstelle_driver()
        login(driver, cfg)
        while True:
            print(f"Pruefe Monitor... ({time.strftime('%H:%M:%S')})")
            driver.get(cfg["monitor_url"]); time.sleep(5)
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            for iframe_v in [None] + iframes:
                try:
                    if iframe_v: driver.switch_to.frame(iframe_v)
                    else: driver.switch_to.default_content()
                    if dcp_name in driver.page_source:
                        for zeile in driver.find_elements(By.XPATH, f"//tr[contains(.,'{dcp_name}')]"):
                            zt = zeile.text
                            print(f"Status: {zt[:150]}")
                            try:
                                zeile.find_element(By.XPATH, ".//*[contains(@class,'success') or contains(@class,'check') or contains(@src,'success') or contains(@class,'ok')]")
                                driver.switch_to.default_content(); return "erfolg"
                            except: pass
                            if "100%" in zt:
                                driver.switch_to.default_content(); return "erfolg"
                            try:
                                zeile.find_element(By.XPATH, ".//*[contains(@class,'error') or contains(@class,'fail') or contains(@src,'error')]")
                                driver.switch_to.default_content(); return "fehler"
                            except: pass
                            print("Ingest laeuft noch...")
                        driver.switch_to.default_content(); break
                    driver.switch_to.default_content()
                except: driver.switch_to.default_content()
            print("Naechster Check in 30 Sekunden...")
            time.sleep(30)
    except Exception as e:
        print(f"Monitor Fehler: {e}")
        if driver: driver.save_screenshot(f"{LOG_ORDNER}\\doremi_monitor_fehler.png")
        return "fehler"
    finally:
        if driver: driver.quit()

def verschiebe_dcp_ins_archiv(dcp_ordner_pfad, dcp_name):
    config = lade_config()
    archiv = config["ordner"]["dcp_archiv"]
    os.makedirs(archiv, exist_ok=True)
    ziel = os.path.join(archiv, dcp_name)
    if os.path.exists(dcp_ordner_pfad):
        shutil.move(dcp_ordner_pfad, ziel)
        print(f"DCP archiviert: {ziel}")
    return ziel

def lade_hoch_und_ingest(dcp_name, dcp_ordner_pfad):
    from modules.telegram_bot import sende_nachricht
    if not ftp_upload(dcp_ordner_pfad, dcp_name):
        sende_nachricht(f"Upload fehlgeschlagen!\nDCP: {dcp_name}"); return False
    time.sleep(5)
    if not starte_ingest(dcp_name):
        sende_nachricht(f"Ingest nicht gestartet!\nDCP: {dcp_name}"); return False
    ergebnis = ueberwache_ingest(dcp_name)
    if ergebnis == "erfolg":
        verschiebe_dcp_ins_archiv(dcp_ordner_pfad, dcp_name)
        ftp_loesche_dcp(dcp_name)
        sende_nachricht(f"Alles erledigt!\nDCP: {dcp_name}\nHochgeladen\nIngest erfolgreich\nArchiviert\nFTP bereinigt")
        return True
    else:
        sende_nachricht(f"Ingest fehlgeschlagen!\nDCP: {dcp_name}\nBitte manuell pruefen!")
        return False
'@ | Set-Content -Path "C:\dcp_automatisierung\modules\doremi.py" -Encoding UTF8

# ── main.py ──────────────────────────────────────────────────
@'
import os
import time
import threading
import schedule
import yaml
from pathlib import Path
from modules.watcher import suche_neue_bilder
from modules.analyzer import lese_text_aus_bild
from modules.naming import bestimme_dcp_name_komplett
from modules.dcpomatic import erstelle_dcp, verschiebe_in_archiv
from modules.doremi import lade_hoch_und_ingest
from modules.telegram_bot import sende_nachricht, starte_listener

LAEUFT = True
VERARBEITUNG_AKTIV = False

def lade_config():
    with open("C:\\dcp_automatisierung\\config.yaml", "r") as f:
        return yaml.safe_load(f)

def hole_alle_eingangsordner():
    c = lade_config()
    return [c["ordner"]["eingang_7sec"], c["ordner"]["eingang_10sec"], c["ordner"]["eingang_15sec"]]

def verarbeite_bild(bildpfad, eingangsordner):
    dateiname = Path(bildpfad).name
    print(f"\n{'='*50}\nVerarbeite: {dateiname}\n{'='*50}")
    try:
        ocr_text = lese_text_aus_bild(bildpfad)
        print(f"OCR: {ocr_text[:200] if ocr_text else 'Kein Text'}")
        dcp_name = bestimme_dcp_name_komplett(bildpfad, ocr_text)
        if not dcp_name:
            print(f"Uebersprungen: {dateiname}"); return False
        print(f"DCP-Name: {dcp_name}")
        if not erstelle_dcp(bildpfad, dcp_name, eingangsordner):
            print(f"DCP-Erstellung fehlgeschlagen!"); return False
        verschiebe_in_archiv(bildpfad)
        config = lade_config()
        dcp_pfad = os.path.join(config["ordner"]["dcp_ausgabe"], dcp_name)
        if os.path.exists(dcp_pfad):
            lade_hoch_und_ingest(dcp_name, dcp_pfad)
        else:
            sende_nachricht(f"DCP Ordner nicht gefunden!\n{dcp_name}")
        return True
    except Exception as e:
        print(f"Fehler: {e}")
        sende_nachricht(f"Fehler!\n{dateiname}\n{str(e)[:300]}")
        return False

def verarbeite_fertige_dcps():
    config = lade_config()
    dcp_ausgabe = config["ordner"]["dcp_ausgabe"]
    if not os.path.exists(dcp_ausgabe): return 0, 0
    dcps = [d for d in os.listdir(dcp_ausgabe) if os.path.isdir(os.path.join(dcp_ausgabe, d))]
    if not dcps:
        print("Keine fertigen DCPs"); return 0, 0
    print(f"{len(dcps)} fertige DCP(s) gefunden!")
    sende_nachricht(f"{len(dcps)} fertige DCP(s)!\nStarte Upload...")
    ok, fehler = 0, 0
    for dcp_name in dcps:
        if lade_hoch_und_ingest(dcp_name, os.path.join(dcp_ausgabe, dcp_name)):
            ok += 1
        else:
            fehler += 1
    return ok, fehler

def pruefe_alle_ordner():
    global VERARBEITUNG_AKTIV
    if VERARBEITUNG_AKTIV:
        print("Verarbeitung laeuft..."); return
    VERARBEITUNG_AKTIV = True
    print(f"\nCheck... ({time.strftime('%H:%M:%S')})")
    ok, fehler = 0, 0
    try:
        alle_bilder = []
        for ordner in hole_alle_eingangsordner():
            if os.path.exists(ordner):
                for bild in suche_neue_bilder(ordner):
                    alle_bilder.append((bild, ordner))
        if alle_bilder:
            sende_nachricht(f"{len(alle_bilder)} neue Bilder!\nStarte Verarbeitung...")
            for bildpfad, ordner in alle_bilder:
                if verarbeite_bild(bildpfad, ordner): ok += 1
                else: fehler += 1
        else:
            print("Keine neuen Bilder")
        d_ok, d_fehler = verarbeite_fertige_dcps()
        ok += d_ok; fehler += d_fehler
        if ok > 0 or fehler > 0:
            msg = f"Check abgeschlossen!\nErfolgreich: {ok}\n"
            if fehler > 0: msg += f"Fehlgeschlagen: {fehler}\n"
            sende_nachricht(msg)
        else:
            print("Nichts zu tun")
    except Exception as e:
        print(f"Check Fehler: {e}")
        sende_nachricht(f"Check Fehler:\n{str(e)[:300]}")
    finally:
        VERARBEITUNG_AKTIV = False

def telegram_befehl(text):
    text = text.strip()
    if text in ["1", "/check"]:
        sende_nachricht("Starte Check...")
        threading.Thread(target=pruefe_alle_ordner, daemon=True).start()
    elif text in ["2", "/status"]:
        config = lade_config()
        msg = f"System laeuft!\nDoremi: {config.get('doremi',{}).get('ip','?')}\n\n"
        for o in hole_alle_eingangsordner():
            n = len(suche_neue_bilder(o)) if os.path.exists(o) else "?"
            msg += f"{Path(o).name}: {n} Bilder\n"
        dcps = [d for d in os.listdir(config["ordner"]["dcp_ausgabe"]) if os.path.isdir(os.path.join(config["ordner"]["dcp_ausgabe"],d))] if os.path.exists(config["ordner"]["dcp_ausgabe"]) else []
        msg += f"DCP-Ordner: {len(dcps)} DCP(s)\n\n1 /check\n2 /status\n3 /stop\n4 /restart\n5 /hilfe"
        sende_nachricht(msg)
    elif text in ["3", "/stop"]:
        sende_nachricht("Wird beendet..."); time.sleep(1); os._exit(0)
    elif text in ["4", "/restart"]:
        sende_nachricht("Neustart..."); time.sleep(1)
        import sys; os.execv(sys.executable, [sys.executable] + sys.argv)
    elif text in ["5", "/hilfe", "/help"]:
        sende_nachricht(
            "Befehle:\n/check Sofortiger Check\n/status Systemstatus\n"
            "/stop Beenden\n/restart Neustart\n/hilfe Hilfe\n\n"
            "Bei Bildern:\n1 Vorschlag\n2 Eigener Name\n3 Kein Datum\n4 Ueberspringen"
        )

def scheduler_thread():
    global LAEUFT
    config = lade_config()
    intervall = config.get("zeitplan", {}).get("intervall_minuten", 60)
    schedule.every(intervall).minutes.do(pruefe_alle_ordner)
    while LAEUFT:
        schedule.run_pending(); time.sleep(30)

def main():
    global LAEUFT
    print("="*60 + "\nDCP AUTOMATISIERUNG startet...\n" + "="*60)
    config = lade_config()
    for ordner in config["ordner"].values():
        os.makedirs(ordner, exist_ok=True)
    threading.Thread(target=starte_listener, args=(telegram_befehl,), daemon=True).start()
    threading.Thread(target=scheduler_thread, daemon=True).start()
    ip = config.get("doremi", {}).get("ip", "?")
    sende_nachricht(f"DCP Automatisierung gestartet!\nDoremi: {ip}\n\n/check /status /stop /restart /hilfe")
    print("Erster Check...")
    pruefe_alle_ordner()
    print("\nSystem laeuft!\n")
    try:
        while True: time.sleep(1)
    except KeyboardInterrupt:
        LAEUFT = False
        sende_nachricht("DCP Automatisierung beendet.")

if __name__ == "__main__":
    main()
'@ | Set-Content -Path "C:\dcp_automatisierung\main.py" -Encoding UTF8

Write-Host "      Alle Scripts erstellt - OK" -ForegroundColor Gray

# ─── SCHRITT 7: VENV + PAKETE ───────────────────────────────
Write-Host "[7/7] Installiere Python-Pakete (bitte warten)..." -ForegroundColor Green
Set-Location "C:\dcp_automatisierung"
python -m venv venv
& "C:\dcp_automatisierung\venv\Scripts\pip.exe" install --upgrade pip -q
& "C:\dcp_automatisierung\venv\Scripts\pip.exe" install requests pillow pytesseract pyyaml schedule selenium webdriver-manager watchdog -q
Write-Host "      Alle Pakete installiert - OK" -ForegroundColor Gray

# ─── NSSM DIENST ────────────────────────────────────────────
Write-Host "Richte Windows-Dienst ein..." -ForegroundColor Green
& "C:\nssm\nssm.exe" stop dcp_automatisierung 2>$null | Out-Null
& "C:\nssm\nssm.exe" remove dcp_automatisierung confirm 2>$null | Out-Null


& "C:\nssm\nssm.exe" install dcp_automatisierung "C:\dcp_automatisierung\venv\Scripts\python.exe" "C:\dcp_automatisierung\main.py"
& "C:\nssm\nssm.exe" set dcp_automatisierung AppDirectory "C:\dcp_automatisierung"
& "C:\nssm\nssm.exe" set dcp_automatisierung DisplayName "DCP Automatisierung"
& "C:\nssm\nssm.exe" set dcp_automatisierung Description "Automatische DCP Erstellung und Ingest"
& "C:\nssm\nssm.exe" set dcp_automatisierung Start SERVICE_AUTO_START
& "C:\nssm\nssm.exe" set dcp_automatisierung AppStdout "C:\dcp_automatisierung\logs\service.log"
& "C:\nssm\nssm.exe" set dcp_automatisierung AppStderr "C:\dcp_automatisierung\logs\service_error.log"
& "C:\nssm\nssm.exe" set dcp_automatisierung AppRestartDelay 5000
& "C:\nssm\nssm.exe" start dcp_automatisierung

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "   INSTALLATION ABGESCHLOSSEN!" -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "   Laufwerk  : ${LAUFWERK}:\" -ForegroundColor White
Write-Host "   Doremi IP : $DOREMI_IP" -ForegroundColor White
Write-Host "   Dienst    : Laeuft automatisch!" -ForegroundColor White
Write-Host ""
Write-Host "   Eingangsordner:" -ForegroundColor Yellow
Write-Host "     ${LAUFWERK}:\K.O.D Atomations\Neue LB         (7 Sek)" -ForegroundColor Gray
Write-Host "     ${LAUFWERK}:\K.O.D Atomations\Neue LB 10sec   (10 Sek)" -ForegroundColor Gray
Write-Host "     ${LAUFWERK}:\K.O.D Atomations\Neue LB 15sec   (15 Sek)" -ForegroundColor Gray
Write-Host ""
Write-Host "   Telegram-Befehle:" -ForegroundColor Yellow
Write-Host "     /check /status /stop /restart /hilfe" -ForegroundColor Gray
Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Read-Host "Druecke Enter zum Beenden"
