@echo off
echo ============================================
echo   Full export dependencies setup for NSSM
echo ============================================
echo.

cd /d "%~dp0"

REM -- try to stop common NSSM service names --
echo [0/5] Stopping NSSM service ...
nssm stop StandardSubstance 2>nul
nssm stop flask_app 2>nul
nssm stop FlaskService 2>nul
echo   Waiting 3 seconds ...
timeout /t 3 /nobreak >nul
echo.

echo [1/5] Installing all dependencies ...
.venv\Scripts\pip install pywin32 reportlab pypdfium2 python-docx Pillow
if %errorlevel% neq 0 (
    echo [WARN] pip install had errors, retrying without force-reinstall ...
    .venv\Scripts\pip install pywin32 reportlab pypdfium2 python-docx Pillow --no-deps
)
echo.

echo [2/5] Copying pywin32 DLLs to venv Scripts ...
for %%f in (.venv\Lib\site-packages\pywin32_system32\*.dll) do (
    copy /Y "%%f" ".venv\Scripts\" >nul
    echo   copied %%~nxf
)
echo.

echo [3/5] Creating Desktop folders for SYSTEM profile ...
if not exist "C:\Windows\System32\config\systemprofile\Desktop" (
    mkdir "C:\Windows\System32\config\systemprofile\Desktop"
    echo   Created System32\Desktop
) else (
    echo   System32\Desktop OK
)
if exist "C:\Windows\SysWOW64" (
    if not exist "C:\Windows\SysWOW64\config\systemprofile\Desktop" (
        mkdir "C:\Windows\SysWOW64\config\systemprofile\Desktop"
        echo   Created SysWOW64\Desktop
    ) else (
        echo   SysWOW64\Desktop OK
    )
)
echo.

echo [4/5] Verifying all imports ...
.venv\Scripts\python.exe -c "import pywintypes; import win32com.client; import reportlab; import pypdfium2; import docx; from PIL import Image; print('All imports OK')"
if %errorlevel% neq 0 (
    echo.
    echo [FAIL] Some imports failed
    pause
    exit /b 1
)
echo.

echo [5/5] Checking Word COM ...
.venv\Scripts\python.exe -c "import win32com.client; w=win32com.client.Dispatch('Word.Application'); print('Word OK'); w.Quit()"
if %errorlevel% neq 0 (
    echo.
    echo [FAIL] Word COM not available, install Microsoft Word
    pause
    exit /b 1
)
echo.
echo ============================================
echo   All done! Starting NSSM service ...
echo ============================================
nssm start StandardSubstance 2>nul
nssm start flask_app 2>nul
nssm start FlaskService 2>nul
echo.
echo   If service did not start, run manually:
echo   nssm start YOUR_SERVICE_NAME
echo.
pause
