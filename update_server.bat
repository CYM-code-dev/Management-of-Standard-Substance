@echo off
chcp 936 >nul 2>&1
title Update

set "PROJECT_DIR=%~dp0."

cd /d "%PROJECT_DIR%"

REM stop print server BEFORE git reset - node locks .node files in deploy\print-server
schtasks /End /TN NiimbotPrint >nul 2>&1
ping -n 3 127.0.0.1 >nul
for /f "tokens=5" %%p in ('netstat -ano ^| findstr /C:":5001 " ^| findstr /C:"LISTENING"') do taskkill /F /T /PID %%p >nul 2>&1

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

echo [4/5] restart FlaskStdMgr service...
set "NSSM_PATH="
where nssm >nul 2>&1 && set "NSSM_PATH=nssm"
if not defined NSSM_PATH if exist "%PROJECT_DIR%\nssm.exe" set "NSSM_PATH=%PROJECT_DIR%\nssm.exe"
if defined NSSM_PATH (
    %NSSM_PATH% restart FlaskStdMgr
    if errorlevel 1 echo [WARN] nssm restart FlaskStdMgr failed
) else (
    echo [WARN] nssm not found, cannot restart FlaskStdMgr
    echo        Please restart manually: nssm restart FlaskStdMgr
)
echo.

echo [5/5] PortalAutoLogin service...
REM portal_config.ini is gitignored (holds plaintext wifi/portal password),
REM so it never arrives via git - must exist on server already or service crash-loops.
if not exist "%PROJECT_DIR%\portal_config.ini" (
    echo [SKIP] portal_config.ini missing - put it next to portal_auto_login.py, then re-run
) else if not defined NSSM_PATH (
    echo [SKIP] nssm not found, cannot install PortalAutoLogin
) else (
    %NSSM_PATH% status PortalAutoLogin >nul 2>&1
    if not errorlevel 1 (
        echo        Restarting PortalAutoLogin...
        %NSSM_PATH% restart PortalAutoLogin
        if errorlevel 1 echo [WARN] nssm restart PortalAutoLogin failed
    ) else (
        echo        Installing PortalAutoLogin...
        %NSSM_PATH% install PortalAutoLogin "%PROJECT_DIR%\.venv\Scripts\python.exe" "%PROJECT_DIR%\portal_auto_login.py"
        %NSSM_PATH% set PortalAutoLogin AppDirectory "%PROJECT_DIR%"
        %NSSM_PATH% set PortalAutoLogin AppStdout "%PROJECT_DIR%\portal_auto_login_nssm.log"
        %NSSM_PATH% set PortalAutoLogin AppStderr "%PROJECT_DIR%\portal_auto_login_nssm.log"
        %NSSM_PATH% set PortalAutoLogin AppExit Default Restart
        %NSSM_PATH% set PortalAutoLogin Start SERVICE_AUTO_START
        %NSSM_PATH% start PortalAutoLogin
    )
)
echo.

echo [6/6] print server (NiimbotPrint, deploy\print-server)...
schtasks /Run /TN NiimbotPrint >nul 2>&1
if errorlevel 1 (echo [WARN] schtasks /Run NiimbotPrint failed) else (echo        started)

echo.
echo Done!
pause
