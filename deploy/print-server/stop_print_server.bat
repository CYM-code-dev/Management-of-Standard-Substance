@echo off
cd /d "%~dp0"
title Stop Niimbot Print Server

net session >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Right-click this file and choose "Run as administrator"
    pause
    exit /b 1
)

schtasks /End /TN NiimbotPrint >nul 2>&1
for /f "tokens=5" %%p in ('netstat -ano ^| findstr /C:":5001 " ^| findstr /C:"LISTENING"') do taskkill /F /PID %%p >nul 2>&1
echo [OK] Print service stopped.
echo       It will start again at next boot (or run start_print_server.bat).
echo       To remove autostart completely:
echo       schtasks /Delete /TN NiimbotPrint /F
pause
