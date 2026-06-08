@echo off
chcp 936 >nul 2>&1
title Update

set "PROJECT_DIR=%~dp0Management-of-Standard-Substance"

cd /d "%PROJECT_DIR%"

echo [1/3] git pull...
git pull
echo.

echo [2/3] pip install...
.venv\Scripts\pip install -r requirements.txt
echo.

echo [3/3] restart service...
nssm restart FlaskStdMgr 2>nul || echo service not registered, skip
echo.

echo Done!
pause
