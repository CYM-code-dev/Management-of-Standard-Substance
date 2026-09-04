@echo off
cd /d "%~dp0"
title Start Niimbot Print Server

net session >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Right-click this file and choose "Run as administrator"
    pause
    exit /b 1
)

schtasks /Change /TN NiimbotPrintWatchdog /ENABLE >nul 2>&1
schtasks /Run /TN NiimbotPrint >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Task not found. Run deploy_print_server.bat first.
    pause
    exit /b 1
)
echo [OK] Print service started in background (no window).
pause
