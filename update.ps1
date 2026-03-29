# ============================================================
#   DCP AUTOMATISIERUNG - AUTO-UPDATE
# ============================================================

$ErrorActionPreference = "Continue"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
chcp 65001 | Out-Null

Write-Host "DCP Automatisierung - Auto-Update wird ausgefuehrt..." -ForegroundColor Cyan

$configPfad = "C:\dcp_automatisierung\config.yaml"
$LAUFWERK = "E"
$DOREMI_IP = "172.20.23.11"

if (Test-Path $configPfad) {
    $configText = Get-Content $configPfad -Raw
    if ($configText -match 'eingang_7sec:\s*"([A-Z]):\\\\') { $LAUFWERK = $Matches[1] }
    if ($configText -match 'ip:\s*"([0-9.]+)"') { $DOREMI_IP = $Matches[1] }
    Write-Host "  Laufwerk: $LAUFWERK | Doremi: $DOREMI_IP" -ForegroundColor Gray
}

$rulesPfad = "C:\dcp_automatisierung\rules\naming_rules.yaml"
$rulesBackup = ""
if (Test-Path $rulesPfad) {
    $rulesBackup = Get-Content $rulesPfad -Raw
    Write-Host "  Naming-Regeln gesichert" -ForegroundColor Gray
}

$installPfad = "$env:TEMP\dcp_update_install.ps1"
$GITHUB_URL = "https://raw.githubusercontent.com/Sh-Kod/LB-automatisierung/main/install.ps1"

Write-Host "Lade neuen Installer von GitHub..." -ForegroundColor Yellow
$erfolgreich = $false
for ($v = 1; $v -le 3; $v++) {
    try {
        curl.exe -L --retry 2 -s -o "$installPfad" "$GITHUB_URL" 2>$null
        if ((Test-Path $installPfad) -and (Get-Item $installPfad).Length -gt 1000) {
            $erfolgreich = $true; break
        }
    } catch { }
    if ($v -lt 3) { Start-Sleep -Seconds (5 * $v) }
}

if (-not $erfolgreich) { Write-Host "Download fehlgeschlagen!" -ForegroundColor Red; exit 1 }

Write-Host "Starte Installation..." -ForegroundColor Green
Start-Process powershell -ArgumentList @("-ExecutionPolicy", "Bypass", "-File", $installPfad, "-LAUFWERK_PARAM", $LAUFWERK, "-DOREMI_IP_PARAM", $DOREMI_IP) -Wait

if ($rulesBackup -ne "" -and (Test-Path $rulesPfad)) {
    $aktuell = Get-Content $rulesPfad -Raw
    if ($aktuell.Trim() -eq "regeln: []") {
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($rulesPfad, $rulesBackup, $utf8NoBom)
        Write-Host "  Naming-Regeln wiederhergestellt!" -ForegroundColor Gray
    }
}

Remove-Item $installPfad -Force -ErrorAction SilentlyContinue
Write-Host "Auto-Update abgeschlossen!" -ForegroundColor Green