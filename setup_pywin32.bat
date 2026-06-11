@echo off
echo ============================================
echo   pywin32 + Word COM fix for NSSM service
echo ============================================
echo.

cd /d "%~dp0"

echo [1/4] Installing pywin32 ...
.venv\Scripts\pip install pywin32 --force-reinstall
echo.

echo [2/4] Copying DLLs to venv Scripts ...
for %%f in (.venv\Lib\site-packages\pywin32_system32\*.dll) do (
    copy /Y "%%f" ".venv\Scripts\" >nul
    echo   copied %%~nxf
)
echo.

echo [3/4] Creating Desktop folders for SYSTEM profile ...
if not exist "C:\Windows\System32\config\systemprofile\Desktop" (
    mkdir "C:\Windows\System32\config\systemprofile\Desktop"
    echo   Created System32\Desktop
) else (
    echo   System32\Desktop already exists
)
if exist "C:\Windows\SysWOW64" (
    if not exist "C:\Windows\SysWOW64\config\systemprofile\Desktop" (
        mkdir "C:\Windows\SysWOW64\config\systemprofile\Desktop"
        echo   Created SysWOW64\Desktop
    ) else (
        echo   SysWOW64\Desktop already exists
    )
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
echo   All done!
echo   Please restart the NSSM service now.
echo ============================================
pause
