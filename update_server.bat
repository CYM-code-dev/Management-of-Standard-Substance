@echo off
chcp 936 >nul 2>&1
title Update

set "PROJECT_DIR=%~dp0."

cd /d "%PROJECT_DIR%"

echo [1/4] git fetch + reset...
git fetch origin
git reset --hard origin/main-1
echo.

echo [2/4] pip install + pywin32 DLL fix...
.venv\Scripts\pip install -r requirements.txt
REM copy pywin32 DLLs for NSSM service
for %%f in (.venv\Lib\site-packages\pywin32_system32\*.dll) do (
    copy /Y "%%f" ".venv\Scripts\" >nul 2>&1
)
REM create SYSTEM Desktop for Word COM
if not exist "C:\Windows\System32\config\systemprofile\Desktop" mkdir "C:\Windows\System32\config\systemprofile\Desktop"
if exist "C:\Windows\SysWOW64" if not exist "C:\Windows\SysWOW64\config\systemprofile\Desktop" mkdir "C:\Windows\SysWOW64\config\systemprofile\Desktop"
echo.

echo [3/4] verifying imports...
.venv\Scripts\python.exe -c "import pywintypes; import win32com.client; import reportlab; import pypdfium2; import docx; from PIL import Image; print('All OK')"
if %errorlevel% neq 0 (
    echo [WARN] some imports failed, check manually
)
echo.

echo [4/4] restart service...
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
