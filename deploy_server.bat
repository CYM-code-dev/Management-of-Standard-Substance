@echo off
chcp 936 >nul 2>&1
title Deploy

echo ============================================
echo   deploy start
echo ============================================
echo.

set "REPO_URL=https://github.com/CYM-code-dev/Management-of-Standard-Substance.git"
set "PROJECT_DIR=%~dp0Management-of-Standard-Substance"
set "SERVICE_NAME=FlaskStdMgr"
set "PORT=5000"

:: --- Git ---
echo [check] Git...
where git >nul 2>&1
if errorlevel 1 (
    echo [install] Git via winget...
    winget install --id Git.Git -e --accept-source-agreements --accept-package-agreements 2>nul
    if errorlevel 1 (
        echo [install] Git via download...
        powershell -NoProfile -ExecutionPolicy Bypass -Command "irm 'https://github.com/git-for-windows/git/releases/download/v2.49.0.windows.1/Git-2.49.0-64-bit.exe' -OutFile $env:TEMP\git.exe; Start-Process $env:TEMP\git.exe -ArgumentList '/VERYSILENT /NORESTART' -Wait; del $env:TEMP\git.exe"
    )
    set "PATH=%PATH%;C:\Program Files\Git\cmd"
)
where git >nul 2>&1
if errorlevel 1 (
    echo [FAIL] Git install failed. Install manually: https://git-scm.com
    goto :end
)
git --version
echo.

:: --- Node.js ---
echo [check] Node.js...
where node >nul 2>&1
if errorlevel 1 (
    echo [install] Node.js via winget...
    winget install --id OpenJS.NodeJS.LTS -e --accept-source-agreements --accept-package-agreements 2>nul
    if errorlevel 1 (
        echo [install] Node.js via download...
        powershell -NoProfile -ExecutionPolicy Bypass -Command "irm 'https://nodejs.org/dist/v22.16.0/node-v22.16.0-x64.msi' -OutFile $env:TEMP\node.msi; Start-Process msiexec -ArgumentList '/i $env:TEMP\node.msi /quiet /norestart' -Wait; del $env:TEMP\node.msi"
    )
    set "PATH=%PATH%;C:\Program Files\nodejs"
)
where node >nul 2>&1
if errorlevel 1 (
    echo [SKIP] Node.js not available, printer service disabled
) else (
    node --version
)
echo.

:: --- Python ---
echo [check] Python...
where python >nul 2>&1
if errorlevel 1 (
    echo [install] Python via winget...
    winget install --id Python.Python.3.12 -e --accept-source-agreements --accept-package-agreements 2>nul
    if errorlevel 1 (
        echo [install] Python via download...
        powershell -NoProfile -ExecutionPolicy Bypass -Command "irm 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe' -OutFile $env:TEMP\py.exe; Start-Process $env:TEMP\py.exe -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1 Include_pip=1' -Wait; del $env:TEMP\py.exe"
    )
    set "PATH=%PATH%;C:\Program Files\Python312;C:\Program Files\Python312\Scripts"
)
where python >nul 2>&1
if errorlevel 1 (
    echo [FAIL] Python install failed. Install manually: https://python.org
    goto :end
)
python --version
echo.

:: =============================================
:: 1. clone or pull
:: =============================================
echo [1/5] Code...
if not exist "%PROJECT_DIR%\.git" (
    git clone %REPO_URL% "%PROJECT_DIR%"
    if errorlevel 1 (
        echo [FAIL] git clone failed. Check network.
        goto :end
    )
) else (
    cd /d "%PROJECT_DIR%"
    git pull
)
cd /d "%PROJECT_DIR%"
echo.

:: =============================================
:: 2. pip install
:: =============================================
echo [2/5] Python deps...
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
)
.venv\Scripts\pip install -r requirements.txt
echo.

:: =============================================
:: 3. npm install
:: =============================================
where node >nul 2>&1
if not errorlevel 1 (
    if not exist "node_modules" (
        echo [3/5] Node deps...
        call npm install
    ) else (
        echo [3/5] Node deps OK
    )
) else (
    echo [3/5] Skip npm
)
echo.

:: =============================================
:: 4. config.json
:: =============================================
if not exist "config.json" (
    echo [4/5] WARNING: config.json missing! Create it manually.
    echo.
    pause
) else (
    echo [4/5] config.json OK
)
echo.

:: =============================================
:: 5. NSSM service
:: =============================================
echo [5/5] Auto-start service...
set "NSSM_PATH="
where nssm >nul 2>&1
if not errorlevel 1 (
    set "NSSM_PATH=nssm"
) else (
    if exist "%PROJECT_DIR%\nssm.exe" (
        set "NSSM_PATH=%PROJECT_DIR%\nssm.exe"
    ) else (
        echo        Downloading NSSM...
        powershell -NoProfile -ExecutionPolicy Bypass -Command "irm 'https://nssm.cc/release/nssm-2.24.zip' -OutFile $env:TEMP\nssm.zip; Expand-Archive $env:TEMP\nssm.zip $env:TEMP\nssm -Force; copy $env:TEMP\nssm\nssm-2.24\win64\nssm.exe %PROJECT_DIR%\nssm.exe; del $env:TEMP\nssm.zip"
        if exist "%PROJECT_DIR%\nssm.exe" (
            set "NSSM_PATH=%PROJECT_DIR%\nssm.exe"
        )
    )
)

if not defined NSSM_PATH (
    echo [SKIP] NSSM not available
    goto :done
)

%nssm_path% status %SERVICE_NAME% >nul 2>&1
if not errorlevel 1 (
    echo        Restarting service...
    %NSSM_PATH% restart %SERVICE_NAME%
) else (
    echo        Installing service...
    %NSSM_PATH% install %SERVICE_NAME% "%PROJECT_DIR%\.venv\Scripts\python.exe" "%PROJECT_DIR%\login_html.py"
    %NSSM_PATH% set %SERVICE_NAME% AppDirectory "%PROJECT_DIR%"
    %NSSM_PATH% set %SERVICE_NAME% DisplayName "Flask StdMgr"
    %NSSM_PATH% set %SERVICE_NAME% Start SERVICE_AUTO_START
    %NSSM_PATH% start %SERVICE_NAME%
)

:done
echo.
echo ============================================
echo   Done! http://SERVER_IP:%PORT%/login
echo ============================================
echo.
echo Update later:
echo   cd /d "%PROJECT_DIR%"
echo   git pull
echo   nssm restart %SERVICE_NAME%
echo.

:end
pause
