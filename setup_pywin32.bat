@echo off
echo ============================================
echo   pywin32 fix for venv
echo ============================================
echo.

cd /d "%~dp0"

echo [1/3] Installing pywin32 ...
.venv\Scripts\pip install pywin32 --force-reinstall
echo.

echo [2/3] Running post-install ...
.venv\Scripts\python.exe -m pywin32_postinstall -install
echo.

echo [3/3] Copying DLLs to venv Scripts ...
for %%f in (.venv\Lib\site-packages\pywin32_system32\*.dll) do (
    copy /Y "%%f" ".venv\Scripts\" >nul
    echo   copied %%~nxf
)
echo.

echo [4/4] Checking Word COM ...
.venv\Scripts\python.exe -c "import pywintypes; import win32com.client; w=win32com.client.Dispatch('Word.Application'); print('Word OK'); w.Quit()"
if %errorlevel% neq 0 (
    echo.
    echo [FAIL] Word COM not available.
    pause
    exit /b 1
)
echo.
echo ============================================
echo   All done, PDF export ready
echo ============================================
pause
