@echo off
REM ============================================================
REM  Management-of-Standard-Substance remote update.
REM  Run from the dev/test machine over ssh.
REM  Default target: E:\server\Management-of-Standard-Substance\update_server.bat
REM  on the server (NSSM services FlaskStdMgr + PortalAutoLogin, port 5000).
REM  Pass another server-side .bat to update a different program:
REM    update_remote.bat E:\server\mup-web\update_server.bat
REM  Server ssh setup was done 2026-08 (see mup-web repo update_remote.bat
REM  header for the one-time sshd instructions).
REM ============================================================
setlocal
title stdmgr remote update

set "SRV=10.1.93.25"
set "SUSER=Administrator"
set "CMD=cmd /c E:\server\Management-of-Standard-Substance\update_server.bat"
if not "%~1"=="" set "CMD=cmd /c %~1"

REM stdin for the ssh session is redirected from NUL so the trailing
REM 'pause' inside the remote bat cannot hang the session
ssh %SUSER%@%SRV% "%CMD% < NUL"
if errorlevel 1 ( echo [X] remote update FAILED & pause & exit /b 1 )

echo === waiting for service to come back ===
for /l %%i in (1,1,60) do (
  curl -s -o nul -m 3 http://%SRV%:5000/login && goto ok
  ping -n 4 127.0.0.1 >nul
)
echo [X] service did not come back within ~5min & exit /b 1
:ok
echo OK: http://%SRV%:5000/login is back
pause
endlocal
