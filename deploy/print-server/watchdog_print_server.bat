@echo off
rem Watchdog: run every 5 min by scheduled task NiimbotPrintWatchdog (SYSTEM).
rem Service healthy -> exit. Dead or hung -> kill listener, restart task.
curl -sf -m 5 -o nul http://localhost:5001/connected
if not errorlevel 1 exit /b 0
for /f "tokens=5" %%p in ('netstat -ano ^| findstr /C:":5001 " ^| findstr /C:"LISTENING"') do taskkill /F /PID %%p >nul 2>&1
schtasks /Run /TN NiimbotPrint >nul 2>&1
exit /b 0
