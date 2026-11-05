@echo off
setlocal
set "APP=%~dp0danhao.ps1"
if not exist "%APP%" (
    echo [!] danhao.ps1 not found next to this launcher.
    pause
    exit /b 1
)
start "" powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -WindowStyle Hidden -File "%APP%"
exit /b 0
