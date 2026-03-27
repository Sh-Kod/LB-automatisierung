# ============================================================
#   DCP AUTOMATISIERUNG - INSTALLER v2.1
#   Ausfuehren mit:
#   powershell -ExecutionPolicy Bypass -File install.ps1
# ============================================================

param(
    [string]$LAUFWERK_PARAM = "",
    [string]$DOREMI_IP_PARAM = ""
)

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

$IS_UPDATE = ($LAUFWERK_PARAM -ne "")

if (-not $IS_UPDATE) {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "   DCP AUTOMATISIERUNG - INSTALLER v2.1" -ForegroundColor Cyan
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

if (-not $IS_UPDATE) {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Cyan
    Write-Host "   Laufwerk  : ${LAUFWERK}:\\" -ForegroundColor White
    Write-Host "   Doremi IP : $DOREMI_IP" -ForegroundColor White
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

"2.1" | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\version.txt", $_, $utf8NoBom) }

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
  token: "8655165819:AAFqrEPOO8OGCR3jHBOoFe9vflceVYTfpAc"
  chat_id: "479976191"
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
from collections import deque

_last_update_id = 0
_update_id_lock = threading.Lock()
_naming_aktiv = False
_meldungs_queue = deque()
_queue_lock = threading.Lock()

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

def sende_hintergrund(text):
    if _naming_aktiv:
        with _queue_lock:
            _meldungs_queue.append(text)
        print(f"[QUEUED] {text[:80]}")
    else:
        sende_nachricht(text)

def setze_naming_aktiv(aktiv):
    global _naming_aktiv
    _naming_aktiv = aktiv
    if not aktiv:
        with _queue_lock:
            nachrichten = list(_meldungs_queue)
            _meldungs_queue.clear()
        for msg in nachrichten:
            sende_nachricht(msg)
            time.sleep(0.3)

def sende_bild(bildpfad, caption=""):
    try:
        config = lade_config()
        token = config["telegram"]["token"]
        chat_id = config["telegram"]["chat_id"]
        url = f"https://api.telegram.org/bot{token}/sendPhoto"
        with open(bildpfad, "rb") as f:
            requests.post(url, data={"chat_id": chat_id, "caption": caption}, files={"photo": f}, timeout=30)
    except Exception as e:
        print(f"Telegram Bild Fehler: {e}")

def warte_auf_antwort(timeout=300):
    global _last_update_id
    try:
        config = lade_config()
        token = config["telegram"]["token"]
        chat_id = str(config["telegram"]["chat_id"])
        url = f"https://api.telegram.org/bot{token}/getUpdates"
        start = time.time()
        while time.time() - start < timeout:
            try:
                with _update_id_lock:
                    offset = _last_update_id + 1
                resp = requests.get(url, params={"offset": offset, "timeout": 5}, timeout=10)
                updates = resp.json().get("result", [])
                for update in updates:
                    with _update_id_lock:
                        _last_update_id = update["update_id"]
                    msg = update.get("message", {})
                    if str(msg.get("chat", {}).get("id", "")) == chat_id:
                        text = msg.get("text", "").strip()
                        if text:
                            return text
            except Exception:
                time.sleep(2)
        return None
    except Exception as e:
        print(f"warte_auf_antwort Fehler: {e}")
        return None

def starte_listener(callback):
    global _last_update_id
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
            if _naming_aktiv:
                time.sleep(1)
                continue
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
                if str(msg.get("chat", {}).get("id", "")) == chat_id:
                    text = msg.get("text", "").strip()
                    if text and callback and not _naming_aktiv:
                        threading.Thread(target=callback, args=(text,), daemon=True).start()
        except Exception:
            time.sleep(5)
'@ | ForEach-Object { [System.IO.File]::WriteAllText("C:\\dcp_automatisierung\\modules\\telegram_bot.py", $_, $utf8NoBom) }

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
& "C:\\nssm\\nssm.exe" install dcp_automatisierung "$pythonExe" "C:\\dcp_automatisierung\\main.py" | Out-Null
& "C:\\nssm\\nssm.exe" set dcp_automatisierung AppDirectory "C:\\dcp_automatisierung" | Out-Null
& "C:\\nssm\\nssm.exe" set dcp_automatisierung DisplayName "DCP-Automatisierung" | Out-Null
& "C:\\nssm\\nssm.exe" set dcp_automatisierung Start SERVICE_AUTO_START | Out-Null
& "C:\\nssm\\nssm.exe" set dcp_automatisierung AppStdout "C:\\dcp_automatisierung\\logs\\service.log" | Out-Null
& "C:\\nssm\\nssm.exe" set dcp_automatisierung AppStderr "C:\\dcp_automatisierung\\logs\\service_error.log" | Out-Null
& "C:\\nssm\\nssm.exe" set dcp_automatisierung AppRestartDelay 5000 | Out-Null
& "C:\\nssm\\nssm.exe" start dcp_automatisierung | Out-Null
Start-Sleep -Seconds 4
$svc = Get-Service -Name "dcp_automatisierung" -ErrorAction SilentlyContinue
$status = if ($svc) { $svc.Status } else { "Nicht gefunden" }

if ($IS_UPDATE) {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host "   AUTO-UPDATE ABGESCHLOSSEN! v2.1" -ForegroundColor Green
    Write-Host "   Dienst: $status" -ForegroundColor White
    Write-Host "  ============================================================" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host "   INSTALLATION ABGESCHLOSSEN! v2.1" -ForegroundColor Green
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "   Laufwerk  : ${LAUFWERK}:\\" -ForegroundColor White
    Write-Host "   Doremi IP : $DOREMI_IP" -ForegroundColor White
    Write-Host "   Python    : $pythonExe" -ForegroundColor White
    Write-Host "   Dienst    : $status" -ForegroundColor White
    Write-Host ""
    Write-Host "   Telegram-Befehle:" -ForegroundColor Yellow
    Write-Host "     /check /status /queue /stop /restart /hilfe" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  ============================================================" -ForegroundColor Green
    Write-Host ""
    Read-Host "Druecke Enter zum Beenden"
}
