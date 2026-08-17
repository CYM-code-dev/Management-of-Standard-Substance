@echo off
cd /d "%~dp0"
title Niimbot print server 5001

rem already running? (started by scheduled task or a previous window)
netstat -ano | findstr /C:":5001 " | findstr /C:"LISTENING" >nul
if not errorlevel 1 (
    echo Port 5001 is already LISTENING - print server is already running.
    echo Close this window. Do NOT start a second copy.
    ping -n 11 127.0.0.1 >nul
    exit /b
)

:loop
rem direct node call - works under SYSTEM scheduled task, no npm needed
node "%~dp0node_modules\@mmote\niimblue-node\cli.mjs" server -p 5001 --cors -h 0.0.0.0
echo.
echo [%date% %time%] server exited, restarting in 5s...
ping -n 6 127.0.0.1 >nul
goto loop
