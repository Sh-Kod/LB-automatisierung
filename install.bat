@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

echo.
echo  ============================================================
echo   DCP AUTOMATISIERUNG - INSTALLER v1.0
echo  ============================================================
echo.

net session >nul 2>&1
if !errorLevel! neq 0 (
    echo [FEHLER] Bitte als Administrator ausfuehren!
    echo Rechtsklick auf install.bat dann "Als Administrator ausfuehren"
    pause
    exit /b 1
)

echo Auf welchem Laufwerk sollen die Ordner erstellt werden?
echo Beispiel: E fuer E:\ oder C fuer C:\
echo.
set /p LAUFWERK="Laufwerk eingeben (Standard E): "
if "!LAUFWERK!"=="" set LAUFWERK=E
set LAUFWERK=!LAUFWERK::=!
echo.

echo Welche Doremi IP-Adresse?
echo Kino 3 = 172.20.23.11
echo.
set /p DOREMI_IP="Doremi IP eingeben (Standard 172.20.23.11): "
if "!DOREMI_IP!"=="" set DOREMI_IP=172.20.23.11
echo.

echo ============================================================
echo  Einstellungen:
echo    Laufwerk  : !LAUFWERK!:\
echo    Doremi IP : !DOREMI_IP!
echo    Programm  : C:\dcp_automatisierung\
echo ============================================================
echo.
set /p OK="Installation starten? J/N: "
if /i "!OK!"=="N" exit /b 0
echo.

echo [1/7] Pruefe Python...
python --version >nul 2>&1
if !errorLevel! neq 0 (
    echo       Python nicht gefunden - wird installiert...
    winget install --id Python.Python.3.11 --silent --accept-source-agreements --accept-package-agreements
    set "PATH=!PATH!;C:\Users\!USERNAME!\AppData\Local\Programs\Python\Python311"
    set "PATH=!PATH!;C:\Users\!USERNAME!\AppData\Local\Programs\Python\Python311\Scripts"
    echo       Python installiert!
) else (
    echo       Python bereits installiert - OK
)

echo [2/7] Pruefe Tesseract OCR...
if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    echo       Tesseract bereits installiert - OK
) else (
    echo       Tesseract wird installiert...
    winget install --id UB-Mannheim.TesseractOCR --silent --accept-source-agreements --accept-package-agreements
    echo       Tesseract installiert!
)

echo [3/7] Pruefe Google Chrome...
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    echo       Chrome bereits installiert - OK
) else if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
    echo       Chrome bereits installiert - OK
) else (
    echo       Chrome wird installiert...
    winget install --id Google.Chrome --silent --accept-source-agreements --accept-package-agreements
    echo       Chrome installiert!
)

echo [4/7] Pruefe NSSM...
if exist "C:\nssm\nssm.exe" (
    echo       NSSM bereits installiert - OK
) else (
    echo       NSSM wird installiert...
    powershell -Command "Invoke-WebRequest -Uri 'https://nssm.cc/release/nssm-2.24.zip' -OutFile '$env:TEMP\nssm.zip' -UseBasicParsing" >nul 2>&1
    powershell -Command "Expand-Archive -Path '$env:TEMP\nssm.zip' -DestinationPath '$env:TEMP\nssm_tmp' -Force" >nul 2>&1
    mkdir "C:\nssm" >nul 2>&1
    copy /y "%TEMP%\nssm_tmp\nssm-2.24\win64\nssm.exe" "C:\nssm\nssm.exe" >nul 2>&1
    setx /M PATH "!PATH!;C:\nssm" >nul 2>&1
    echo       NSSM installiert!
)

echo [5/7] Erstelle Ordner...
mkdir "C:\dcp_automatisierung" >nul 2>&1
mkdir "C:\dcp_automatisierung\modules" >nul 2>&1
mkdir "C:\dcp_automatisierung\rules" >nul 2>&1
mkdir "C:\dcp_automatisierung\logs" >nul 2>&1
mkdir "C:\dcp_automatisierung\temp" >nul 2>&1
mkdir "!LAUFWERK!:\K.O.D Atomations\Neue LB" >nul 2>&1
mkdir "!LAUFWERK!:\K.O.D Atomations\Neue LB 10sec" >nul 2>&1
mkdir "!LAUFWERK!:\K.O.D Atomations\Neue LB 15sec" >nul 2>&1
mkdir "!LAUFWERK!:\K.O.D Atomations\DCP" >nul 2>&1
mkdir "!LAUFWERK!:\K.O.D Atomations\AUF TMS" >nul 2>&1
mkdir "!LAUFWERK!:\K.O.D Atomations\Fehler" >nul 2>&1
mkdir "!LAUFWERK!:\K.O.D Atomations\DCP Upload erledigt" >nul 2>&1
echo       Alle Ordner erstellt - OK

echo [6/7] Erstelle Konfiguration und Scripts...

(
echo ordner:
echo   eingang_7sec: "!LAUFWERK!:\\K.O.D Atomations\\Neue LB"
echo   eingang_10sec: "!LAUFWERK!:\\K.O.D Atomations\\Neue LB 10sec"
echo   eingang_15sec: "!LAUFWERK!:\\K.O.D Atomations\\Neue LB 15sec"
echo   dcp_ausgabe: "!LAUFWERK!:\\K.O.D Atomations\\DCP"
echo   archiv: "!LAUFWERK!:\\K.O.D Atomations\\AUF TMS"
echo   fehler: "!LAUFWERK!:\\K.O.D Atomations\\Fehler"
echo   dcp_archiv: "!LAUFWERK!:\\K.O.D Atomations\\DCP Upload erledigt"
echo dcpomatic:
echo   cli_pfad: "C:\\Program Files\\DCP-o-matic 2\\bin\\dcpomatic2_cli.exe"
echo zeitplan:
echo   intervall_minuten: 60
echo telegram:
echo   token: "8655165819:AAFqrEPOO8OGCR3jHBOoFe9vflceVYTfpAc"
echo   chat_id: "479976191"
echo logging:
echo   log_datei: "C:\\dcp_automatisierung\\logs\\dcp_system.log"
echo   log_level: "INFO"
echo gemini:
echo   api_key: "AIzaSyAcxSYME3T3hQNy7vK3wMdENQZuS4RXGzc"
echo tesseract:
echo   pfad: "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
echo doremi:
echo   ip: "!DOREMI_IP!"
echo   ftp_user: "ingest"
echo   ftp_pass: "ingest"
echo   web_user: "admin"
echo   web_pass: "1234"
) > "C:\dcp_automatisierung\config.yaml"

echo regeln: [] > "C:\dcp_automatisierung\rules\naming_rules.yaml"
echo. > "C:\dcp_automatisierung\modules\__init__.py"

python "C:\dcp_automatisierung\setup_scripts.py"
echo       Alle Scripts erstellt - OK

echo [7/7] Installiere Python-Pakete...
cd /d "C:\dcp_automatisierung"
python -m venv venv
call venv\Scripts\activate.bat
pip install --upgrade pip -q
pip install requests -q
pip install pillow -q
pip install pytesseract -q
pip install pyyaml -q
pip install schedule -q
pip install selenium -q
pip install webdriver-manager -q
pip install watchdog -q
echo       Alle Pakete installiert - OK

echo Richte Windows-Dienst ein...
C:\nssm\nssm.exe stop dcp_automatisierung >nul 2>&1
C:\nssm\nssm.exe remove dcp_automatisierung confirm >nul 2>&1
C:\nssm\nssm.exe install dcp_automatisierung "C:\dcp_automatisierung\venv\Scripts\python.exe" "C:\dcp_automatisierung\main.py"
C:\nssm\nssm.exe set dcp_automatisierung AppDirectory "C:\dcp_automatisierung"
C:\nssm\nssm.exe set dcp_automatisierung DisplayName "DCP Automatisierung"
C:\nssm\nssm.exe set dcp_automatisierung Description "Automatische DCP Erstellung und Ingest"
C:\nssm\nssm.exe set dcp_automatisierung Start SERVICE_AUTO_START
C:\nssm\nssm.exe set dcp_automatisierung AppStdout "C:\dcp_automatisierung\logs\service.log"
C:\nssm\nssm.exe set dcp_automatisierung AppStderr "C:\dcp_automatisierung\logs\service_error.log"
C:\nssm\nssm.exe set dcp_automatisierung AppRestartDelay 5000
C:\nssm\nssm.exe start dcp_automatisierung

echo.
echo  ============================================================
echo   INSTALLATION ABGESCHLOSSEN!
echo  ============================================================
echo.
echo   Laufwerk  : !LAUFWERK!:\
echo   Doremi IP : !DOREMI_IP!
echo   Dienst    : Laeuft automatisch beim Windows-Start
echo.
echo   Eingangsordner:
echo     !LAUFWERK!:\K.O.D Atomations\Neue LB         (7 Sek)
echo     !LAUFWERK!:\K.O.D Atomations\Neue LB 10sec   (10 Sek)
echo     !LAUFWERK!:\K.O.D Atomations\Neue LB 15sec   (15 Sek)
echo.
echo   Telegram-Befehle:
echo     /check   - Sofortiger Check
echo     /status  - Systemstatus
echo     /stop    - Beenden
echo     /restart - Neustart
echo.
echo  ============================================================
echo.
pause
