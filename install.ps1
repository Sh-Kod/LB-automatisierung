# ============================================================
#   DCP AUTOMATISIERUNG - INSTALLER v2.3
#   Ausfuehren mit:
#   powershell -ExecutionPolicy Bypass -File install.ps1
# ============================================================

param(
    [string]$LAUFWERK_PARAM = "",
    [string]$DOREMI_IP_PARAM = "",
    [string]$TELEGRAM_TOKEN_PARAM = "",
    [string]$TELEGRAM_CHAT_ID_PARAM = ""
)

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

$IS_UPDATE = ($LAUFWERK_PARAM -ne "")

if (-not $IS_UPDATE) {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "   DCP AUTOMATISIERUNG - INSTALLER v2.3" -ForegroundColor Cyan
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""
}

if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[FEHLER] Bitte als Administrator ausfuehren!" -ForegroundColor Red
    if (-not $IS_UPDATE) { Read-Host "Druecke Enter zum Beenden" }
    exit 1
}

$utf8NoBom = New-Object System.Text.UTF8Encoding $false

if ($LAUFWERK_PARAM -ne "") {
    $LAUFWERK = $LAUFWERK_PARAM.Replace(":", "")
} else {
    Write-Host "Auf welchem Laufwerk sollen die Ordner erstellt werden?" -ForegroundColor Yellow
    Write-Host "Beispiel: E fuer E:\\ oder D fuer D:\\" -ForegroundColor Gray
    $LAUFWERK = Read-Host "Laufwerk eingeben (Standard E)"
    if ($LAUFWERK -eq "") { $LAUFWERK = "E" }
    $LAUFWERK = $LAUFWERK.Replace(":", "")
}

if ($DOREMI_IP_PARAM -ne "") {
    $DOREMI_IP = $DOREMI_IP_PARAM
} else {
    Write-Host ""
    Write-Host "Welche Doremi IP-Adresse?" -ForegroundColor Yellow
    Write-Host "Kino 3 = 172.20.23.11" -ForegroundColor Gray
    $DOREMI_IP = Read-Host "Doremi IP eingeben (Standard 172.20.23.11)"
    if ($DOREMI_IP -eq "") { $DOREMI_IP = "172.20.23.11" }
}

if ($TELEGRAM_TOKEN_PARAM -ne "") {
    $TELEGRAM_TOKEN = $TELEGRAM_TOKEN_PARAM
    $TELEGRAM_CHAT_ID = $TELEGRAM_CHAT_ID_PARAM
} elseif (-not $IS_UPDATE) {
    Write-Host ""
    Write-Host "Telegram Bot Token?" -ForegroundColor Yellow
    Write-Host "Beispiel: 1234567890:AAFxxxxxx..." -ForegroundColor Gray
    $TELEGRAM_TOKEN = Read-Host "Bot Token eingeben"
    if ($TELEGRAM_TOKEN -eq "") { $TELEGRAM_TOKEN = "BITTE_TOKEN_EINTRAGEN" }
    Write-Host ""
    Write-Host "Telegram Chat ID?" -ForegroundColor Yellow
    Write-Host "Beispiel: 479976191" -ForegroundColor Gray
    $TELEGRAM_CHAT_ID = Read-Host "Chat ID eingeben"
    if ($TELEGRAM_CHAT_ID -eq "") { $TELEGRAM_CHAT_ID = "0" }
} else {
    $TELEGRAM_TOKEN = "AUS_BACKUP"
    $TELEGRAM_CHAT_ID = "0"
}

if (-not $IS_UPDATE) {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "   Laufwerk  : ${LAUFWERK}:\\" -ForegroundColor White
    Write-Host "   Doremi IP : $DOREMI_IP" -ForegroundColor White
    Write-Host "   Telegram  : $($TELEGRAM_TOKEN.Substring(0, [Math]::Min(12,$TELEGRAM_TOKEN.Length)))..." -ForegroundColor White
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host ""
    $OK = Read-Host "Installation starten? J/N"
    if ($OK -eq "N" -or $OK -eq "n") { exit 0 }
    Write-Host ""
}

Write-Host "Bereinige alte Installation..." -ForegroundColor Yellow

$configBackup = ""
$rulesBackup = ""
if ($IS_UPDATE) {
    if (Test-Path "C:\\dcp_automatisierung\\config.yaml") {
        $configBackup = Get-Content "C:\\dcp_automatisierung\\config.yaml" -Raw
    }
    if (Test-Path "C:\\dcp_automatisierung\\rules\\naming_rules.yaml") {
        $rulesBackup = Get-Content "C:\\dcp_automatisierung\\rules\\naming_rules.yaml" -Raw
    }
}

$svcAlt = Get-Service -Name "dcp_automatisierung" -ErrorAction SilentlyContinue
if ($svcAlt) {
    & "C:\\nssm\\nssm.exe" stop dcp_automatisierung | Out-Null
    Start-Sleep -Seconds 3
    & "C:\\nssm\\nssm.exe" remove dcp_automatisierung confirm | Out-Null
    Start-Sleep -Seconds 2
}
if (Test-Path "C:\\dcp_automatisierung") {
    Remove-Item -Path "C:\\dcp_automatisierung" -Recurse -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
}
Write-Host "Bereinigung abgeschlossen!" -ForegroundColor Gray
Write-Host ""

Write-Host "[1/8] Pruefe Python..." -ForegroundColor Green
try {
    $v = python --version 2>&1
    Write-Host "      $v - OK" -ForegroundColor Gray
} catch {
    Write-Host "      Python nicht gefunden - wird installiert..." -ForegroundColor Yellow
    winget install --id Python.Python.3.11 --silent --accept-source-agreements --accept-package-agreements
    $env:PATH += ";C:\\Users\\$env:USERNAME\\AppData\\Local\\Programs\\Python\\Python311"
    $env:PATH += ";C:\\Users\\$env:USERNAME\\AppData\\Local\\Programs\\Python\\Python311\\Scripts"
    Write-Host "      Python installiert!" -ForegroundColor Gray
}

Write-Host "[2/8] Pruefe Tesseract OCR..." -ForegroundColor Green
if (Test-Path "C:\\Program Files\\Tesseract-OCR\\tesseract.exe") {
    Write-Host "      Tesseract bereits installiert - OK" -ForegroundColor Gray
} else {
    Write-Host "      Tesseract wird installiert..." -ForegroundColor Yellow
    winget install --id UB-Mannheim.TesseractOCR --silent --accept-source-agreements --accept-package-agreements --source winget
    Write-Host "      Tesseract installiert!" -ForegroundColor Gray
}

Write-Host "[3/8] Pruefe DCP-o-matic 2..." -ForegroundColor Green
if (Test-Path "C:\\Program Files\\DCP-o-matic 2\\bin\\dcpomatic2_cli.exe") {
    Write-Host "      DCP-o-matic bereits installiert - OK" -ForegroundColor Gray
} else {
    Write-Host "      DCP-o-matic wird heruntergeladen (ca. 200MB)..." -ForegroundColor Yellow
    $dcpInstaller = "$env:TEMP\\dcpomatic_setup.exe"
    curl.exe -L --retry 3 --retry-delay 5 -o "$dcpInstaller" "https://dcpomatic.com/dl.php?id=windows-2.16.78"
    Start-Process -FilePath $dcpInstaller -ArgumentList "/S" -Wait
    Remove-Item $dcpInstaller -Force -ErrorAction SilentlyContinue
    Write-Host "      DCP-o-matic installiert!" -ForegroundColor Gray
}

Write-Host "[4/8] Pruefe Google Chrome..." -ForegroundColor Green
if ((Test-Path "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe") -or (Test-Path "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe")) {
    Write-Host "      Chrome bereits installiert - OK" -ForegroundColor Gray
} else {
    Write-Host "      Chrome wird installiert..." -ForegroundColor Yellow
    winget install --id Google.Chrome --silent --accept-source-agreements --accept-package-agreements
    Write-Host "      Chrome installiert!" -ForegroundColor Gray
}

Write-Host "[5/8] Pruefe NSSM..." -ForegroundColor Green
if (Test-Path "C:\\nssm\\nssm.exe") {
    Write-Host "      NSSM bereits installiert - OK" -ForegroundColor Gray
} else {
    Write-Host "      NSSM wird installiert..." -ForegroundColor Yellow
    curl.exe -L --retry 3 --retry-delay 2 -o "$env:TEMP\\nssm.zip" "https://nssm.cc/release/nssm-2.24.zip"
    Expand-Archive -Path "$env:TEMP\\nssm.zip" -DestinationPath "$env:TEMP\\nssm_tmp" -Force
    New-Item -ItemType Directory -Path "C:\\nssm" -Force | Out-Null
    Copy-Item "$env:TEMP\\nssm_tmp\\nssm-2.24\\win64\\nssm.exe" "C:\\nssm\\nssm.exe" -Force
    $machinePath = [Environment]::GetEnvironmentVariable("PATH", "Machine")
    [Environment]::SetEnvironmentVariable("PATH", "$machinePath;C:\\nssm", "Machine")
    $env:PATH += ";C:\\nssm"
    Write-Host "      NSSM installiert!" -ForegroundColor Gray
}

Write-Host "[6/8] Erstelle Ordner..." -ForegroundColor Green
$ordner = @(
    "C:\\dcp_automatisierung",
    "C:\\dcp_automatisierung\\modules",
    "C:\\dcp_automatisierung\\rules",
    "C:\\dcp_automatisierung\\logs",
    "C:\\dcp_automatisierung\\temp",
    "${LAUFWERK}:\\K.O.D Atomations\\Neue LB",
    "${LAUFWERK}:\\K.O.D Atomations\\Neue LB 10sec",
    "${LAUFWERK}:\\K.O.D Atomations\\Neue LB 15sec",
    "${LAUFWERK}:\\K.O.D Atomations\\DCP",
    "${LAUFWERK}:\\K.O.D Atomations\\AUF TMS",
    "${LAUFWERK}:\\K.O.D Atomations\\Fehler",
    "${LAUFWERK}:\\K.O.D Atomations\\DCP Upload erledigt"
)
foreach ($o in $ordner) { New-Item -ItemType Directory -Path $o -Force | Out-Null }
Write-Host "      Alle Ordner erstellt - OK" -ForegroundColor Gray

Write-Host "[7/8] Erstelle Scripts und Konfiguration..." -ForegroundColor Green

"2.3" | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\version.txt", $_, $utf8NoBom) }

if ($IS_UPDATE -and $configBackup -ne "") {
    [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\config.yaml", $configBackup, $utf8NoBom)
    $cfg = Get-Content "C:\\dcp_automatisierung\\config.yaml" -Raw
    if ($cfg -notmatch "github_version_url") {
        $addSection = "`nupdate:`n  github_version_url: `"https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/version.txt`"`n  github_update_url: `"https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/update.ps1`"`n"
        [System.IO.File]::AppendAllText("C:\\dcp_automatisierung\\config.yaml", $addSection, $utf8NoBom)
    }
    Write-Host "      Config aus Backup wiederhergestellt!" -ForegroundColor Gray
} else {
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
  create_pfad: "C:\\Program Files\\DCP-o-matic 2\\bin\\dcpomatic2_create.exe"
zeitplan:
  intervall_minuten: 60
telegram:
  token: "$TELEGRAM_TOKEN"
  chat_id: "$TELEGRAM_CHAT_ID"
logging:
  log_datei: "C:\\dcp_automatisierung\\logs\\dcp_system.log"
  log_level: "INFO"
tesseract:
  pfad: "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
doremi:
  ip: "$DOREMI_IP"
  ftp_user: "ingest"
  ftp_pass: "ingest"
  web_user: "admin"
  web_pass: "1234"
update:
  github_version_url: "https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/version.txt"
  github_update_url: "https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/update.ps1"
"@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\config.yaml", $_, $utf8NoBom) }
}

if ($rulesBackup -ne "") {
    [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\rules\\naming_rules.yaml", $rulesBackup, $utf8NoBom)
    Write-Host "      Naming-Regeln wiederhergestellt!" -ForegroundColor Gray
} else {
    "regeln: []" | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\rules\\naming_rules.yaml", $_, $utf8NoBom) }
}

"" | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\modules\\__init__.py", $_, $utf8NoBom) }

@'
import logging
import yaml
from pathlib import Path

def erstelle_logger():
    with open("C:\\dcp_automatisierung\\config.yaml", "r", encoding="utf-8") as f:
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
'@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\modules\\logger.py", $_, $utf8NoBom) }

@'
import os
from pathlib import Path

ERLAUBTE_ENDUNGEN = [".jpg", ".jpeg", ".png"]

def suche_neue_bilder(ordner):
    bilder = []
    if not os.path.exists(ordner):
        return bilder
    for datei in os.listdir(ordner):
        if datei.startswith("."):
            continue
        if Path(datei).suffix.lower() in ERLAUBTE_ENDUNGEN:
            bilder.append(os.path.join(ordner, datei))
    return sorted(bilder)
'@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\modules\\watcher.py", $_, $utf8NoBom) }

@'
import yaml
from PIL import Image
import pytesseract

def lade_config():
    with open("C:\\dcp_automatisierung\\config.yaml", "r", encoding="utf-8") as f:
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
'@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\modules\\analyzer.py", $_, $utf8NoBom) }

@'
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
'@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\modules\\telegram_bot.py", $_, $utf8NoBom) }

@'
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
    with open(QUEUE_PFAD, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

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
'@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\modules\\queue_manager.py", $_, $utf8NoBom) }

@'
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

PHASEN = {1: "DCP", 2: "Upload", 3: "Ingest", 4: "Monitoring"}

def setze_status_callback(fn):
    global _status_callback
    _status_callback = fn

def _lade():
    if not os.path.exists(JOBS_PFAD):
        return {"jobs": []}
    try:
        with open(JOBS_PFAD, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"jobs": []}

def _speichere(data):
    with open(JOBS_PFAD, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

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
                # DCP-Fehler: komplett neu; sonst ab letzter erfolgreicher Phase + 1
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
            _flush_buffer()

def _flush_buffer():
    """Muss mit _status_buffer_lock aufgerufen werden."""
    if not _status_buffer:
        return
    if len(_status_buffer) == 1:
        _sende_einzeln(_status_buffer[0])
    else:
        fertig = [j for j in _status_buffer if j["current_status"] == "done"]
        fehler = [j for j in _status_buffer if j["current_status"] == "error"]
        teile = []
        if fertig:
            teile.append(f"{len(fertig)} fertig")
        if fehler:
            teile.append(f"{len(fehler)} Fehler")
        msg = "Batch-Status: " + ", ".join(teile)
        if fehler:
            msg += "\n\nFehler-Details:"
            for j in fehler[:5]:
                phase = PHASEN.get(j.get("error_phase"), "?")
                name = (j.get("final_name") or os.path.basename(j.get("bildpfad", "")))[:30]
                msg += f"\n• {name} [{phase}]: {(j.get('fehler_text') or '')[:80]}"
        if _status_callback:
            _status_callback(msg)
    _status_buffer.clear()

def _sende_einzeln(job):
    if not _status_callback:
        return
    name = (job.get("final_name") or os.path.basename(job.get("bildpfad", "")))
    if job["current_status"] == "done":
        msg = f"Fertig: {name}"
    elif job["current_status"] == "error":
        phase = PHASEN.get(job.get("error_phase"), "?")
        msg = f"Fehler [{phase}]: {name}\n{job.get('fehler_text', '')[:200]}"
    else:
        return
    _status_callback(msg)

def sende_bundle_wenn_noetig():
    """Wird alle 5 Minuten vom Scheduler aufgerufen."""
    with _status_buffer_lock:
        _flush_buffer()
'@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\modules\\job_manager.py", $_, $utf8NoBom) }

@'
import os
import subprocess
import tempfile
import threading
import time
import urllib.request

import schedule
import yaml

from modules import analyzer, job_manager, queue_manager, telegram_bot, watcher

CONFIG_PFAD = "C:\\dcp_automatisierung\\config.yaml"
VERSION_PFAD = "C:\\dcp_automatisierung\\version.txt"

# ──────────────────────────────────────────────
# Hilfsfunktionen
# ──────────────────────────────────────────────

def lade_config():
    with open(CONFIG_PFAD, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def speichere_config(config):
    with open(CONFIG_PFAD, "w", encoding="utf-8") as f:
        yaml.dump(config, f, allow_unicode=True, default_flow_style=False)

def lese_version():
    try:
        return open(VERSION_PFAD, encoding="utf-8").read().strip()
    except Exception:
        return "unbekannt"

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
    """Scannt Eingangsordner und legt neue Bilder in die Benennungs-Queue."""
    cfg = lade_config().get("ordner", {})
    gesamt = 0
    gesamt += _scan_ordner(cfg.get("eingang_7sec", ""), 7)
    gesamt += _scan_ordner(cfg.get("eingang_10sec", ""), 10)
    gesamt += _scan_ordner(cfg.get("eingang_15sec", ""), 15)
    if gesamt > 0:
        telegram_bot.sende_nachricht(f"{gesamt} neues Bild(er) in Queue aufgenommen.")

# ──────────────────────────────────────────────
# Queue-Worker: sequentieller Benennungs-Dialog
# ──────────────────────────────────────────────

def queue_worker():
    """Laeuft dauerhaft in eigenem Thread.
    Verarbeitet Queue-Items SEQUENTIELL (genau 1 Dialog gleichzeitig).
    Hintergrund-Jobs laufen parallel."""
    time.sleep(5)  # kurz warten bis System vollstaendig gestartet
    while True:
        item = queue_manager.naechstes_pending()
        if not item:
            time.sleep(10)
            continue

        if not telegram_bot.starte_dialog():
            # Sollte nicht vorkommen (Worker laeuft single-threaded)
            queue_manager.zuruecksetzen(item["id"])
            time.sleep(5)
            continue

        try:
            bildpfad = item["bildpfad"]
            dauer_sek = item["dauer_sek"]
            pending_rest = queue_manager.pending_anzahl()

            # Bild senden
            caption = f"Neues Bild ({dauer_sek}s Werbung)"
            if pending_rest > 0:
                caption += f"  |  Noch {pending_rest} in Queue"
            try:
                telegram_bot.sende_bild(bildpfad, caption=caption)
            except Exception:
                telegram_bot.sende_nachricht(f"Bild: {os.path.basename(bildpfad)}")

            # OCR
            ocr_text = ""
            try:
                ocr_text = analyzer.lese_text_aus_bild(bildpfad)
            except Exception:
                pass

            prompt = (
                f"Erkannter Text:\n{ocr_text[:400]}\n\n"
                f"Bitte DCP-Namen eingeben.\n"
                f"(oder /skip zum Ueberspringen, Timeout: 60 Min)"
            )
            telegram_bot.sende_nachricht(prompt)

            # Auf Nutzer-Antwort warten (60 Minuten Timeout)
            antwort = telegram_bot.warte_auf_dialog_antwort(timeout=3600)

            if antwort is None or antwort.strip().lower() == "/skip":
                queue_manager.abschliessen(item["id"])
                telegram_bot.sende_nachricht(
                    f"Uebersprungen: {os.path.basename(bildpfad)}"
                )
            else:
                final_name = antwort.strip()
                queue_manager.abschliessen(item["id"])
                job_id = job_manager.erstelle_job(bildpfad, final_name)
                telegram_bot.sende_nachricht(
                    f"Name: {final_name}\nVerarbeitung laeuft im Hintergrund..."
                )
                threading.Thread(
                    target=verarbeite_job,
                    args=(job_id, 1),
                    daemon=True
                ).start()

        except Exception as e:
            telegram_bot.sende_nachricht(f"Fehler in Queue-Worker: {e}")
            queue_manager.zuruecksetzen(item["id"])
        finally:
            telegram_bot.beende_dialog()
            time.sleep(1)

# ──────────────────────────────────────────────
# Job-Pipeline: DCP → Upload → Ingest → Monitoring
# (Phasen koennen einzeln wiederholt werden)
# ──────────────────────────────────────────────

def _phase_ausfuehren(job_id, phase, fn):
    """Fuehrt eine Phase aus. Gibt True bei Erfolg, False bei Fehler zurueck."""
    try:
        job_manager.aktualisiere_phase(job_id, phase, "running")
        fn(job_id)
        job_manager.aktualisiere_phase(job_id, phase, "done")
        return True
    except Exception as e:
        job_manager.markiere_fehler(job_id, phase, str(e), retryable=True)
        return False

def verarbeite_job(job_id, ab_phase=1):
    """Fuehrt Job-Pipeline ab der angegebenen Phase aus.
    Laeuft im Hintergrund-Thread - mehrere Jobs parallel erlaubt."""
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

# --- Platzhalter-Implementierungen (Phase 2 ersetzt diese) ---

def _dcp_erstellen(job_id):
    # TODO: DCP-Erstellung mit dcpomatic2_cli
    raise NotImplementedError("DCP-Erstellung: noch nicht implementiert")

def _upload_durchfuehren(job_id):
    # TODO: FTP-Upload zur Doremi
    raise NotImplementedError("FTP-Upload: noch nicht implementiert")

def _ingest_starten(job_id):
    # TODO: Doremi-Ingest ueber Web-Interface
    raise NotImplementedError("Doremi-Ingest: noch nicht implementiert")

def _monitoring_ueberwachen(job_id):
    # TODO: Ingest-Monitoring ueberwachen
    raise NotImplementedError("Monitoring: noch nicht implementiert")

# ──────────────────────────────────────────────
# Update
# ──────────────────────────────────────────────

def pruefe_update():
    try:
        cfg = lade_config()
        url = cfg.get("update", {}).get("github_version_url", "")
        if not url:
            telegram_bot.sende_nachricht("Keine Update-URL konfiguriert.")
            return
        github_version = urllib.request.urlopen(url, timeout=10).read().decode().strip()
        local_version = lese_version()
        if github_version != local_version:
            telegram_bot.sende_nachricht(
                f"Update verfuegbar!\n"
                f"Installiert: v{local_version}  →  Neu: v{github_version}\n"
                f"Starte Update..."
            )
            update_url = cfg.get("update", {}).get("github_update_url", "")
            tmp = os.path.join(tempfile.gettempdir(), "dcp_update.ps1")
            urllib.request.urlretrieve(update_url, tmp)
            subprocess.Popen(
                ["powershell", "-ExecutionPolicy", "Bypass", "-File", tmp],
                creationflags=subprocess.CREATE_NEW_CONSOLE
            )
        else:
            telegram_bot.sende_nachricht(f"Bereits aktuell: v{local_version}")
    except Exception as e:
        telegram_bot.sende_nachricht(f"Update fehlgeschlagen: {e}")

# ──────────────────────────────────────────────
# Intervall aendern (persistent + Scheduler-Neustart)
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
        telegram_bot.sende_nachricht(f"Check-Intervall: {minuten} Minuten (gespeichert).")
    except ValueError:
        telegram_bot.sende_nachricht("Ungueltige Eingabe. Beispiel: /intervall 60")

# ──────────────────────────────────────────────
# Fehler-Menue
# ──────────────────────────────────────────────

def zeige_fehler():
    fehler = job_manager.hole_fehler()
    if not fehler:
        telegram_bot.sende_nachricht("Keine Fehler vorhanden.")
        return
    phasen = {1: "DCP", 2: "Upload", 3: "Ingest", 4: "Monitoring"}
    msg = f"{len(fehler)} Fehler-Job(s):\n"
    for j in fehler:
        phase = phasen.get(j.get("error_phase"), "?")
        name = (j.get("final_name") or os.path.basename(j.get("bildpfad", "")))[:35]
        fehlertext = (j.get("fehler_text") or "")[:100]
        msg += f"\nID: {j['id']}\nName: {name}\nPhase: {phase}\nFehler: {fehlertext}\n"
    msg += "\n/retry_alle  oder  /retry <ID>"
    telegram_bot.sende_nachricht(msg)

# ──────────────────────────────────────────────
# Befehlshandler
# ──────────────────────────────────────────────

def bearbeite_befehl(text):
    cmd = text.strip()
    low = cmd.lower()

    if low == "/version":
        telegram_bot.sende_nachricht(f"DCP-Automatisierung v{lese_version()}")

    elif low == "/update":
        telegram_bot.sende_nachricht("Suche nach Updates...")
        threading.Thread(target=pruefe_update, daemon=True).start()

    elif low == "/status":
        fehler = job_manager.hole_fehler()
        aktive = job_manager.hole_aktive()
        pending = queue_manager.pending_anzahl()
        cfg = lade_config()
        intervall = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
        telegram_bot.sende_nachricht(
            f"System: Aktiv  |  v{lese_version()}\n"
            f"Queue: {pending} wartend\n"
            f"Jobs: {len(aktive)} laufend, {len(fehler)} Fehler\n"
            f"Check-Intervall: {intervall} Min"
        )

    elif low == "/check":
        telegram_bot.sende_nachricht("Starte manuelle Ordner-Pruefung...")
        threading.Thread(target=starte_verarbeitung, daemon=True).start()

    elif low.startswith("/intervall"):
        teile = cmd.split(maxsplit=1)
        if len(teile) < 2:
            cfg = lade_config()
            aktuell = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
            telegram_bot.sende_nachricht(
                f"Aktuelles Intervall: {aktuell} Minuten\nAendern: /intervall <Minuten>"
            )
        else:
            aendere_intervall(teile[1])

    elif low == "/fehler":
        zeige_fehler()

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
                f"Job {job_id} neu gestartet ab Phase {j['current_phase']}..."
            )
            threading.Thread(
                target=verarbeite_job, args=(j["id"], j["current_phase"]), daemon=True
            ).start()
        else:
            telegram_bot.sende_nachricht(f"Job '{job_id}' nicht gefunden oder nicht wiederholbar.")

    elif low == "/hilfe":
        telegram_bot.sende_nachricht(
            "Befehle:\n"
            "/version          Aktuelle Version\n"
            "/update           Update von GitHub\n"
            "/check            Ordner jetzt pruefen\n"
            "/status           Systemstatus\n"
            "/intervall <n>    Check-Intervall aendern\n"
            "/fehler           Fehler-Jobs anzeigen\n"
            "/retry <ID>       Job neu starten\n"
            "/retry_alle       Alle Fehler-Jobs neu\n"
            "/hilfe            Diese Hilfe"
        )

    else:
        telegram_bot.sende_nachricht(f"Unbekannt: {text}\n/hilfe fuer alle Befehle.")

# ──────────────────────────────────────────────
# Einstiegspunkt
# ──────────────────────────────────────────────

if __name__ == "__main__":
    # Absturz-Recovery: naming → pending
    queue_manager.naming_zuruecksetzen()

    cfg = lade_config()
    intervall = cfg.get("zeitplan", {}).get("intervall_minuten", 60)

    # Status-Callback fuer job_manager setzen
    job_manager.setze_status_callback(telegram_bot.sende_nachricht)

    telegram_bot.sende_nachricht(f"DCP-Automatisierung v{lese_version()} gestartet.")

    # Telegram-Listener (leitet Nachrichten an Dialog oder Befehlshandler)
    threading.Thread(
        target=telegram_bot.starte_listener,
        args=(bearbeite_befehl,),
        daemon=True
    ).start()

    # Queue-Worker (sequentieller Benennungs-Dialog)
    threading.Thread(target=queue_worker, daemon=True).start()

    # Scheduler
    schedule.every(intervall).minutes.do(starte_verarbeitung).tag("scan")
    schedule.every(5).minutes.do(job_manager.sende_bundle_wenn_noetig).tag("bundle")

    starte_verarbeitung()

    while True:
        schedule.run_pending()
        time.sleep(10)
'@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\main.py", $_, $utf8NoBom) }

Write-Host "      Python-Module erstellt - OK" -ForegroundColor Gray

$utf8NoBomFix = New-Object System.Text.UTF8Encoding $false
foreach ($file in Get-ChildItem "C:\\dcp_automatisierung" -Recurse -Include *.py,*.yaml) {
    $txt = [System.IO.File]::ReadAllText($file.FullName)
    [System.IO.File]::WriteAllText($file.FullName, $txt, $utf8NoBomFix)
}

Write-Host "[8/8] Installiere Python-Pakete..." -ForegroundColor Green
Set-Location "C:\\dcp_automatisierung"
$pythonExe = (Get-Command python).Source
& "$pythonExe" -m pip install --upgrade pip -q
& "$pythonExe" -m pip install requests pillow pytesseract pyyaml schedule selenium webdriver-manager -q
Write-Host "      Alle Pakete installiert - OK" -ForegroundColor Gray

Write-Host "Richte Windows-Dienst ein..." -ForegroundColor Green
$dienst = Get-Service -Name "dcp_automatisierung" -ErrorAction SilentlyContinue
if ($dienst) {
    & "C:\\nssm\\nssm.exe" stop dcp_automatisierung | Out-Null
    Start-Sleep -Seconds 2
    & "C:\\nssm\\nssm.exe" remove dcp_automatisierung confirm | Out-Null
    Start-Sleep -Seconds 2
}
& "C:\nssm\nssm.exe" install dcp_automatisierung "$pythonExe" "C:\dcp_automatisierung\main.py" | Out-Null
& "C:\nssm\nssm.exe" set dcp_automatisierung AppDirectory "C:\dcp_automatisierung" | Out-Null
& "C:\nssm\nssm.exe" set dcp_automatisierung DisplayName "DCP-Automatisierung" | Out-Null
& "C:\nssm\nssm.exe" set dcp_automatisierung Start SERVICE_AUTO_START | Out-Null
& "C:\nssm\nssm.exe" set dcp_automatisierung AppStdout "C:\dcp_automatisierung\logs\service.log" | Out-Null
& "C:\nssm\nssm.exe" set dcp_automatisierung AppStderr "C:\dcp_automatisierung\logs\service_error.log" | Out-Null
& "C:\nssm\nssm.exe" set dcp_automatisierung AppRestartDelay 5000 | Out-Null
& "C:\nssm\nssm.exe" start dcp_automatisierung | Out-Null
Start-Sleep -Seconds 4
$svc = Get-Service -Name "dcp_automatisierung" -ErrorAction SilentlyContinue
$status = if ($svc) { $svc.Status } else { "Nicht gefunden" }

if ($IS_UPDATE) {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host "   AUTO-UPDATE ABGESCHLOSSEN! v2.3" -ForegroundColor Green
    Write-Host "   Dienst: $status" -ForegroundColor White
    Write-Host "  ============================================================" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host "   INSTALLATION ABGESCHLOSSEN! v2.3" -ForegroundColor Green
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "   Laufwerk  : ${LAUFWERK}:\\" -ForegroundColor White
    Write-Host "   Doremi IP : $DOREMI_IP" -ForegroundColor White
    Write-Host "   Python    : $pythonExe" -ForegroundColor White
    Write-Host "   Dienst    : $status" -ForegroundColor White
    Write-Host ""
    Write-Host "   Telegram-Befehle:" -ForegroundColor Yellow
    Write-Host "     /version /update /check /status /intervall /fehler /retry /retry_alle /hilfe" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host ""
    Read-Host "Druecke Enter zum Beenden"
}
