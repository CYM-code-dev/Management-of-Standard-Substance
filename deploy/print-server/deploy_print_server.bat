@echo off
cd /d "%~dp0"
title Niimbot Print Server Deploy

echo ============================================
echo   Niimbot label print server - one-click deploy
echo ============================================
echo.

rem ---- 1. admin check (firewall rule needs it) ----
net session >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Right-click this file and choose "Run as administrator"
    pause
    exit /b 1
)

rem ---- 2. Node.js check (auto download + silent install if missing) ----
where node >nul 2>&1
if errorlevel 1 (
    echo [INFO] Node.js not found. Auto-installing LTS v24.16.0 ...
    if not exist "%~dp0node-lts.msi" (
        curl -fL -o "%~dp0node-lts.msi" https://nodejs.org/dist/v24.16.0/node-v24.16.0-x64.msi
    )
    if not exist "%~dp0node-lts.msi" (
        echo [ERROR] Download failed. Needs internet, or copy node-lts.msi next to this bat.
        pause
        exit /b 1
    )
    echo [INFO] Installing silently, please wait 1-2 minutes...
    msiexec /i "%~dp0node-lts.msi" /qn /norestart
    set "PATH=%PATH%;C:\Program Files\nodejs"
    where node >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Install failed. Double-click node-lts.msi manually, then re-run.
        pause
        exit /b 1
    )
)
for /f "tokens=*" %%v in ('node --version') do echo [OK] Node.js %%v

rem ---- 3. dependencies (node_modules shipped in folder, offline OK) ----
if not exist "node_modules\@mmote\niimblue-node\cli.mjs" (
    echo [INFO] Dependencies missing, trying npm install - needs internet...
    call npm install --omit=dev --no-fund --no-audit
    if errorlevel 1 (
        echo [ERROR] npm install failed. Copy the FULL folder including node_modules.
        pause
        exit /b 1
    )
)
echo [OK] Dependencies ready

rem ---- 4. firewall: allow inbound TCP 5001 ----
netsh advfirewall firewall delete rule name="NiimbotPrint5001" >nul 2>&1
netsh advfirewall firewall add rule name="NiimbotPrint5001" dir=in action=allow protocol=TCP localport=5001 >nul
if errorlevel 1 (echo [WARN] Firewall rule failed, company server may not reach this PC) else (echo [OK] Firewall allows TCP 5001)

rem ---- 5. background service: scheduled task at boot, SYSTEM, no window ----
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v NiimbotPrintServer /f >nul 2>&1
schtasks /Delete /TN NiimbotPrint /F >nul 2>&1
schtasks /Create /TN NiimbotPrint /TR "\"%~dp0run_print_server.bat\"" /SC ONSTART /RU SYSTEM /F >nul
if errorlevel 1 (
    echo [WARN] Task creation failed. Fallback: auto-start on logon with visible window.
    reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v NiimbotPrintServer /t REG_SZ /d "\"%~dp0run_print_server.bat\"" /f >nul
) else (
    echo [OK] Background service installed: starts at boot, no window, no logon needed.
)

rem ---- 6. start service now (windowless) ----
schtasks /Run /TN NiimbotPrint >nul 2>&1
if errorlevel 1 start "" /min cmd /c "%~dp0run_print_server.bat"

echo.
echo [waiting for service to start...]
timeout /t 10 >nul

rem ---- 7. self test ----
curl -s -m 5 -o nul -w "local test http://localhost:5001  HTTP %%{http_code}\n" http://localhost:5001/connected

rem ---- 8. show current LAN IP (auto-detected, not hardcoded) ----
set "LOCAL_IP="
for /f "delims=" %%i in ('powershell -NoProfile -Command "(Get-NetRoute -DestinationPrefix 0.0.0.0/0 ^| Sort-Object RouteMetric ^| Select-Object -First 1 ^| Get-NetIPConfiguration ^| Select-Object -ExpandProperty IPv4Address).IPAddress" 2^>nul') do set "LOCAL_IP=%%i"
if not defined LOCAL_IP set "LOCAL_IP=localhost"

echo.
echo ============================================
echo   Deploy done. Fully automatic background service.
echo   Printer connects via Bluetooth or USB serial (COMx).
echo   Any old console windows can be closed now.
echo   This PC IP: %LOCAL_IP%
echo   Verify from company server:
echo     curl http://%LOCAL_IP%:5001/connected
echo ============================================
pause
