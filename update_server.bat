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
set "NSSM_PATH="
where nssm >nul 2>&1 && set "NSSM_PATH=nssm"
if not defined NSSM_PATH if exist "%PROJECT_DIR%\nssm.exe" set "NSSM_PATH=%PROJECT_DIR%\nssm.exe"
if defined NSSM_PATH (
    %NSSM_PATH% restart FlaskStdMgr
    if errorlevel 1 echo [WARN] nssm restart failed
) else (
    echo [WARN] nssm not found, cannot restart service automatically
    echo        Please restart manually: nssm restart FlaskStdMgr
)
echo.

echo Done!
pause
