# ============================================================
#   DCP AUTOMATISIERUNG - INSTALLER v2.12
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
    Write-Host "   DCP AUTOMATISIERUNG - INSTALLER v2.12" -ForegroundColor Cyan
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
    "C:\\dcp_automatisierung\\staging",
    "C:\\dcp_automatisierung\\backup",
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

"2.11" | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\version.txt", $_, $utf8NoBom) }

if ($IS_UPDATE -and $configBackup -ne "") {
    [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\config.yaml", $configBackup, $utf8NoBom)
    $cfg = Get-Content "C:\\dcp_automatisierung\\config.yaml" -Raw
    if ($cfg -notmatch "github_version_url") {
        $addSection = "`nupdate:`n  github_version_url: `"https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/version.txt`"`n  github_manifest_url: `"https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/update_manifest.json`"`n  github_base_url: `"https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/`"`n  auto_update_intervall_stunden: 24`n"
        [System.IO.File]::AppendAllText("C:\\dcp_automatisierung\\config.yaml", $addSection, $utf8NoBom)
    } elseif ($cfg -notmatch "github_manifest_url") {
        $addFields = "`n  github_manifest_url: `"https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/update_manifest.json`"`n  github_base_url: `"https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/`"`n  auto_update_intervall_stunden: 24`n"
        [System.IO.File]::AppendAllText("C:\\dcp_automatisierung\\config.yaml", $addFields, $utf8NoBom)
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
  github_manifest_url: "https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/update_manifest.json"
  github_base_url: "https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/"
  auto_update_intervall_stunden: 24
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
                # /Befehle immer zum Befehlshandler - auch waehrend Namens-Dialog
                with _dialog_aktiv_lock:
                    dialog = _dialog_aktiv
                if text.startswith("/") and callback:
                    threading.Thread(target=callback, args=(text,), daemon=True).start()
                elif dialog:
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
    tmp = JOBS_PFAD + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, JOBS_PFAD)

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

def hole_job(job_id):
    """Gibt Job-Dict fuer job_id zurueck, oder None."""
    with _lock:
        data = _lade()
        for j in data["jobs"]:
            if j["id"] == job_id:
                return dict(j)
    return None

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
#!/usr/bin/env python3
# updater.py - DCP Automatisierung In-App Updater
# Laeuft als separater Prozess ausserhalb des Windows-Dienstes.
# Stdlib only - keine externen Abhaengigkeiten.

import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime

BASE_PFAD = "C:\\dcp_automatisierung"
PENDING_PFAD = os.path.join(BASE_PFAD, "pending_update.json")
RESULT_PFAD = os.path.join(BASE_PFAD, "update_result.json")
BACKUP_PFAD = os.path.join(BASE_PFAD, "backup")
NSSM = "C:\\nssm\\nssm.exe"
SERVICE = "dcp_automatisierung"
LOG_PFAD = os.path.join(BASE_PFAD, "logs", "updater.log")


def log(msg):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    try:
        os.makedirs(os.path.dirname(LOG_PFAD), exist_ok=True)
        with open(LOG_PFAD, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def schreibe_result(erfolg, neue_version="", fehler=""):
    result = {
        "erfolg": erfolg,
        "neue_version": neue_version,
        "fehler": fehler,
        "timestamp": datetime.now().isoformat()
    }
    try:
        tmp = RESULT_PFAD + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        os.replace(tmp, RESULT_PFAD)
    except Exception as e:
        log(f"Konnte update_result.json nicht schreiben: {e}")


def stoppe_service():
    log("Stoppe Service...")
    try:
        proc = subprocess.Popen(
            [NSSM, "stop", SERVICE],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        proc.wait(timeout=30)
    except Exception as e:
        log(f"nssm stop Fehler: {e}")
    for _ in range(20):
        time.sleep(2)
        try:
            r = subprocess.run(
                ["sc", "query", SERVICE],
                capture_output=True, text=True, timeout=10
            )
            if "STOPPED" in r.stdout:
                log("Service gestoppt.")
                return True
        except Exception:
            pass
    log("Service stop timeout - fahre trotzdem fort.")
    return True


def starte_service():
    log("Starte Service...")
    try:
        proc = subprocess.Popen(
            [NSSM, "start", SERVICE],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        proc.wait(timeout=30)
        time.sleep(3)
        log("Service gestartet.")
    except Exception as e:
        log(f"nssm start Fehler: {e}")


def erstelle_backup(dateien, backup_dir):
    log(f"Erstelle Backup in: {backup_dir}")
    os.makedirs(backup_dir, exist_ok=True)
    gesichert = 0
    for entry in dateien:
        dest = entry["dest"]
        if os.path.exists(dest):
            rel = os.path.relpath(dest, BASE_PFAD)
            backup_dest = os.path.join(backup_dir, rel)
            os.makedirs(os.path.dirname(backup_dest), exist_ok=True)
            shutil.copy2(dest, backup_dest)
            gesichert += 1
    log(f"Backup abgeschlossen: {gesichert} Dateien gesichert.")


def rollback(dateien, backup_dir):
    log(f"Fuehre Rollback durch aus: {backup_dir}")
    wiederhergestellt = 0
    for entry in dateien:
        dest = entry["dest"]
        rel = os.path.relpath(dest, BASE_PFAD)
        backup_src = os.path.join(backup_dir, rel)
        if os.path.exists(backup_src):
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            shutil.copy2(backup_src, dest)
            wiederhergestellt += 1
    log(f"Rollback abgeschlossen: {wiederhergestellt} Dateien wiederhergestellt.")


def wende_update_an(dateien, staging_pfad):
    log("Wende Update an...")
    for entry in dateien:
        src_rel = entry["src"]
        dest = entry["dest"]
        staged = os.path.join(staging_pfad, src_rel.replace("/", os.sep))
        if not os.path.exists(staged):
            raise FileNotFoundError(f"Staged file nicht gefunden: {staged}")
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        tmp = dest + ".new"
        shutil.copy2(staged, tmp)
        os.replace(tmp, dest)
        log(f"  -> {dest}")
    log("Update angewendet.")


def main():
    log("=" * 50)
    log("DCP Automatisierung Updater gestartet")
    log("=" * 50)

    if not os.path.exists(PENDING_PFAD):
        log(f"Kein pending_update.json gefunden: {PENDING_PFAD}")
        sys.exit(1)

    try:
        with open(PENDING_PFAD, "r", encoding="utf-8") as f:
            pending = json.load(f)
    except Exception as e:
        log(f"Konnte pending_update.json nicht lesen: {e}")
        sys.exit(1)

    neue_version = pending.get("neue_version", "?")
    staging_pfad = pending.get("staging_pfad", "")
    dateien = pending.get("dateien", [])

    log(f"Update auf Version: {neue_version}")
    log(f"Anzahl Dateien: {len(dateien)}")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = os.path.join(BACKUP_PFAD, f"v{neue_version}_{ts}")

    stoppe_service()
    time.sleep(2)

    try:
        erstelle_backup(dateien, backup_dir)
    except Exception as e:
        log(f"Backup fehlgeschlagen: {e}")
        schreibe_result(False, neue_version, f"Backup fehlgeschlagen: {e}")
        starte_service()
        sys.exit(1)

    try:
        wende_update_an(dateien, staging_pfad)
    except Exception as e:
        log(f"Update fehlgeschlagen: {e}")
        log("Starte Rollback...")
        try:
            rollback(dateien, backup_dir)
            log("Rollback erfolgreich.")
        except Exception as re:
            log(f"Rollback fehlgeschlagen: {re}")
        schreibe_result(False, neue_version, str(e))
        starte_service()
        sys.exit(1)

    try:
        version_pfad = os.path.join(BASE_PFAD, "version.txt")
        tmp = version_pfad + ".new"
        with open(tmp, "w", encoding="utf-8") as f:
            f.write(neue_version)
        os.replace(tmp, version_pfad)
        log(f"Version aktualisiert: {neue_version}")
    except Exception as e:
        log(f"Konnte version.txt nicht aktualisieren: {e}")

    try:
        if os.path.exists(staging_pfad):
            shutil.rmtree(staging_pfad)
    except Exception:
        pass

    try:
        os.remove(PENDING_PFAD)
    except Exception:
        pass

    schreibe_result(True, neue_version)
    log(f"Update auf v{neue_version} erfolgreich abgeschlossen!")

    starte_service()

    # Task Scheduler Aufgabe aufraeumen
    try:
        subprocess.Popen(
            ["schtasks", "/delete", "/f", "/tn", "DCP_Automatisierung_Updater"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
    except Exception:
        pass

    log("Updater beendet.")


if __name__ == "__main__":
    main()
'@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\updater.py", $_, $utf8NoBom) }

@'
import json
import os
import py_compile
import re
import shutil
import subprocess
import threading
import time
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
    return "\u2500" * 30

def schlage_namen_vor(ocr_text):
    try:
        with open(RULES_PFAD, "r", encoding="utf-8") as f:
            rules_data = yaml.safe_load(f)
        regeln = (rules_data or {}).get("regeln", [])
        for regel in regeln:
            muster  = regel.get("muster", "")
            vorlage = regel.get("vorlage", "")
            if muster and vorlage and re.search(muster, ocr_text, re.IGNORECASE):
                return vorlage.replace("{JAHR}", str(datetime.now().year))
    except Exception:
        pass
    return None

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

def _dialog_zeitüberschreitung(item_id):
    queue_manager.zuruecksetzen(item_id)
    telegram_bot.sende_nachricht("Timeout – Bild zurück in Queue.")

def queue_worker():
    time.sleep(5)
    while True:
        item = queue_manager.naechstes_pending()
        if not item:
            time.sleep(10)
            continue

        if not telegram_bot.starte_dialog():
            queue_manager.zuruecksetzen(item["id"])
            time.sleep(5)
            continue

        try:
            bildpfad    = item["bildpfad"]
            dauer_sek   = item["dauer_sek"]
            pending_rest = queue_manager.pending_anzahl()

            caption = f"Neues Bild ({dauer_sek}s)"
            if pending_rest > 0:
                caption += f"  |  Queue: {pending_rest} wartend"
            try:
                telegram_bot.sende_bild(bildpfad, caption=caption)
            except Exception:
                telegram_bot.sende_nachricht(f"Bild: {os.path.basename(bildpfad)}")

            ocr_text = ""
            try:
                ocr_text = analyzer.lese_text_aus_bild(bildpfad)
            except Exception:
                pass

            vorschlag = schlage_namen_vor(ocr_text)

            t = _trenn()
            msg = f"{t}\n"
            if ocr_text:
                msg += f"OCR-Text:\n{ocr_text[:300].strip()}\n\n"
            if vorschlag:
                msg += f"Vorschlag:\n{vorschlag}\n\n"
                msg += "[1]  Vorschlag übernehmen\n"
                msg += "[2]  Eigenen Namen eingeben\n"
                msg += "[3]  Überspringen\n"
            else:
                msg += "Kein Vorschlag (keine Regel gefunden)\n\n"
                msg += "[1]  Namen eingeben\n"
                msg += "[2]  Überspringen\n"
            msg += f"{t}\n(Timeout: 60 Min)"
            telegram_bot.sende_nachricht(msg)

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
                    final_name = a
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
                queue_manager.abschliessen(item["id"])
                job_id = job_manager.erstelle_job(bildpfad, final_name)
                telegram_bot.sende_nachricht(
                    f"Name gesetzt: {final_name}\nVerarbeitung läuft..."
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

_PHASE_NAMEN = {
    1: ("DCP wird erstellt...", "DCP erstellt"),
    2: ("Upload zum Doremi...", "Upload abgeschlossen"),
    3: ("Ingest wird gestartet...", "Ingest gestartet"),
    4: ("Ingest wird überwacht...", "Ingest abgeschlossen"),
}

def _phase_ausfuehren(job_id, phase, fn):
    job = job_manager.hole_job(job_id)
    name = (job.get("final_name") or "?")[:30] if job else "?"
    start_msg, done_msg = _PHASE_NAMEN.get(phase, (f"Phase {phase}...", f"Phase {phase} fertig"))
    telegram_bot.sende_nachricht(f"[{name}]\n{start_msg}")
    try:
        job_manager.aktualisiere_phase(job_id, phase, "running")
        fn(job_id)
        job_manager.aktualisiere_phase(job_id, phase, "done")
        telegram_bot.sende_nachricht(f"[{name}]\n{done_msg}")
        return True
    except Exception as e:
        job_manager.markiere_fehler(job_id, phase, str(e), retryable=True)
        return False

def verarbeite_job(job_id, ab_phase=1):
    if ab_phase <= 1:
        if not _phase_ausfuehren(job_id, 1, _dcp_erstellen): return
    if ab_phase <= 2:
        if not _phase_ausfuehren(job_id, 2, _upload_durchfuehren): return
    if ab_phase <= 3:
        if not _phase_ausfuehren(job_id, 3, _ingest_starten): return
    if ab_phase <= 4:
        if not _phase_ausfuehren(job_id, 4, _monitoring_ueberwachen): return
    job_manager.markiere_fertig(job_id)

def _dcp_erstellen(job_id):
    import ftplib
    job = job_manager.hole_job(job_id)
    if not job:
        raise RuntimeError(f"Job {job_id} nicht gefunden")

    bildpfad   = job["bildpfad"]
    final_name = job["final_name"]
    cfg        = lade_config()
    ausgabe    = cfg["ordner"]["dcp_ausgabe"]
    create_exe = cfg["dcpomatic"]["create_pfad"]
    cli_exe    = cfg["dcpomatic"]["cli_pfad"]

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

    try:
        r1 = subprocess.run(
            [create_exe,
             "--name", final_name,
             "--still-length", str(dauer),
             "--dcp-content-type", "ADV",
             "-o", tmp_dir,
             bildpfad],
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
        try:
            ftp.mkd(dcp_name)
        except ftplib.error_perm:
            pass
        ftp.cwd(dcp_name)
        _upload_dir(ftp, dcp_pfad)

    dcp_archiv = cfg["ordner"]["dcp_archiv"]
    os.makedirs(dcp_archiv, exist_ok=True)
    ziel = os.path.join(dcp_archiv, dcp_name)
    if os.path.exists(ziel):
        shutil.rmtree(ziel)
    shutil.move(dcp_pfad, ziel)


def _ingest_starten(job_id):
    import base64
    job = job_manager.hole_job(job_id)
    if not job:
        raise RuntimeError(f"Job {job_id} nicht gefunden")

    cfg      = lade_config()
    ip       = cfg["doremi"]["ip"]
    user     = cfg["doremi"]["web_user"]
    passwd   = cfg["doremi"]["web_pass"]
    dcp_name = job["final_name"]

    time.sleep(5)

    creds = base64.b64encode(f"{user}:{passwd}".encode()).decode()
    post_data = f"content={dcp_name}".encode()
    req = urllib.request.Request(
        f"http://{ip}/Ingest/",
        data=post_data, method="POST"
    )
    req.add_header("Authorization", f"Basic {creds}")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    try:
        with urllib.request.urlopen(req, timeout=30):
            pass
    except Exception:
        pass  # FTP-Upload loest Ingest meistens automatisch aus


def _monitoring_ueberwachen(job_id):
    import base64
    job = job_manager.hole_job(job_id)
    if not job:
        raise RuntimeError(f"Job {job_id} nicht gefunden")

    cfg      = lade_config()
    ip       = cfg["doremi"]["ip"]
    user     = cfg["doremi"]["web_user"]
    passwd   = cfg["doremi"]["web_pass"]
    dcp_name = job["final_name"]

    creds = base64.b64encode(f"{user}:{passwd}".encode()).decode()

    for _ in range(120):  # max 20 Minuten
        time.sleep(10)
        try:
            req = urllib.request.Request(f"http://{ip}/ContentInfo/?name={dcp_name}")
            req.add_header("Authorization", f"Basic {creds}")
            with urllib.request.urlopen(req, timeout=10) as resp:
                body = resp.read().decode("utf-8", errors="replace").lower()
                if dcp_name.lower() in body and any(
                    w in body for w in ("complete", "ready", "ok", "ingested")
                ):
                    return
        except Exception:
            pass
    # Timeout ist kein Fehler - DCP wurde hochgeladen und Ingest laeuft

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
                f"Update fehlgeschlagen: {result.get('fehler', '?')}\nRollback wurde durchgeführt."
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
            f"Update verfügbar: v{local_version} → v{github_version}\nLade Dateien herunter..."
        )
        manifest_data = json.loads(urllib.request.urlopen(manifest_url, timeout=10).read().decode())
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
            phase      = phasen.get(j.get("error_phase"), "?")
            name       = (j.get("final_name") or os.path.basename(j.get("bildpfad", "")))[:28]
            fehlertext = (j.get("fehler_text") or "")[:80]
            msg += f"\n[{j['id']}] {name}\n"
            msg += f"  Phase:  {phase}\n"
            msg += f"  Fehler: {fehlertext}\n"
            msg += f"  Aktion: /retry {j['id']}\n"
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
        intervall        = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
        update_intervall = cfg.get("update", {}).get("auto_update_intervall_stunden", 24)
        update_str = f"alle {update_intervall}h" if update_intervall > 0 else "aus"
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
            cfg     = lade_config()
            aktuell = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
            telegram_bot.sende_nachricht(
                f"Scan-Intervall: {aktuell} Minuten\nÄndern: /intervall <Minuten>"
            )
        else:
            aendere_intervall(teile[1])

    elif low.startswith("/update_intervall"):
        teile = cmd.split(maxsplit=1)
        if len(teile) < 2:
            cfg     = lade_config()
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

if __name__ == "__main__":
    queue_manager.naming_zuruecksetzen()

    cfg = lade_config()
    intervall        = cfg.get("zeitplan", {}).get("intervall_minuten", 60)
    update_intervall = cfg.get("update", {}).get("auto_update_intervall_stunden", 24)

    job_manager.setze_status_callback(telegram_bot.sende_nachricht)

    pruefe_update_ergebnis()

    telegram_bot.sende_nachricht(f"DCP-Automatisierung v{lese_version()} gestartet.")

    threading.Thread(
        target=telegram_bot.starte_listener,
        args=(bearbeite_befehl,),
        daemon=True
    ).start()

    threading.Thread(target=queue_worker, daemon=True).start()

    schedule.every(intervall).minutes.do(starte_verarbeitung).tag("scan")
    schedule.every(5).minutes.do(job_manager.sende_bundle_wenn_noetig).tag("bundle")
    if update_intervall > 0:
        schedule.every(update_intervall).hours.do(
            lambda: threading.Thread(target=pruefe_update, daemon=True).start()
        ).tag("update")

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

# Echtes Python finden (Windows Store Stub ignorieren)
$pythonExe = ""
try {
    $pythonExe = (where.exe python 2>$null) |
        Where-Object { $_ -notmatch "WindowsApps" -and $_ -match "python\.exe$" } |
        Select-Object -First 1
} catch {}
if (-not $pythonExe) {
    $pythonExe = (Get-Command python).Source
}
Write-Host "      Python: $pythonExe" -ForegroundColor Gray
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
    Write-Host "   AUTO-UPDATE ABGESCHLOSSEN! v2.12" -ForegroundColor Green
    Write-Host "   Dienst: $status" -ForegroundColor White
    Write-Host "  ============================================================" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host "   INSTALLATION ABGESCHLOSSEN! v2.12" -ForegroundColor Green
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "   Laufwerk  : ${LAUFWERK}:\\" -ForegroundColor White
    Write-Host "   Doremi IP : $DOREMI_IP" -ForegroundColor White
    Write-Host "   Python    : $pythonExe" -ForegroundColor White
    Write-Host "   Dienst    : $status" -ForegroundColor White
    Write-Host ""
    Write-Host "   Telegram-Befehle:" -ForegroundColor Yellow
    Write-Host "     /version /update /update_intervall /check /status /intervall /fehler /retry /retry_alle /hilfe" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host ""
    Read-Host "Druecke Enter zum Beenden"
}
