@echo off
setlocal enabledelayedexpansion
title [Cutting Eval Tool] Supreme Launcher v7.2
cd /d "%~dp0"

echo [1/4] Checking version...
set "TOOLS=%~dp0_tools"
set "REPO=shin9602/Easy-Cutting-Report"
set "VERSION_FILE=%~dp0version.txt"

if not exist "%VERSION_FILE%" echo v1.0.0 > "%VERSION_FILE%"
set /p CURRENT_VER=<"%VERSION_FILE%"
echo [*] Local: %CURRENT_VER%

:: GitHub Update Check (Strictly ASCII to avoid crashes)
powershell -NoProfile -ExecutionPolicy Bypass -Command "try{$r=Invoke-RestMethod https://api.github.com/repos/%REPO%/releases/latest -TimeoutSec 3;$r.tag_name}catch{}" > "%TEMP%\cever.txt" 2>nul
set /p LATEST_VER=<"%TEMP%\cever.txt"
del "%TEMP%\cever.txt" >nul 2>&1

if "!LATEST_VER!"=="" (
    echo [*] GitHub unreachable. Skipping update check.
) else if not "!CURRENT_VER!"=="!LATEST_VER!" (
    echo [!!!] NEW VERSION DETECTED: !LATEST_VER!
    echo [!!!] Please visit github.com/%REPO% to get the latest files.
    echo.
) else (
    echo [*] You are on the latest version.
)

echo [2/4] Python checking...
set "PY_EXE=python"
python --version >nul 2>&1
if %ERRORLEVEL% equ 0 goto SKIP_INSTALL

set "PY_EXE=%TOOLS%\python\python.exe"
if exist "%PY_EXE%" goto SKIP_INSTALL

echo [!] Python not found. AUTO-INSTALLING...
if not exist "%TOOLS%" mkdir "%TOOLS%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12';Invoke-WebRequest https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip -OutFile '%TOOLS%\python.zip' -UseBasicParsing"
mkdir "%TOOLS%\python" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive '%TOOLS%\python.zip' '%TOOLS%\python' -Force"
del "%TOOLS%\python.zip" >nul 2>&1
for /f "delims=" %%F in ('dir /b "%TOOLS%\python\python*._pth" 2^>nul') do powershell -NoProfile -Command "(Get-Content '%TOOLS%\python\%%F') -replace '#import site','import site' | Set-Content '%TOOLS%\python\%%F'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12';Invoke-WebRequest https://bootstrap.pypa.io/get-pip.py -OutFile '%TOOLS%\get-pip.py' -UseBasicParsing"
"%PY_EXE%" "%TOOLS%\get-pip.py" --no-warn-script-location >nul 2>&1

:SKIP_INSTALL
echo [3/4] Installing dependencies...
"%PY_EXE%" -m pip install flask flask-cors pillow openpyxl xlrd==1.2.0 pywin32 --quiet --no-warn-script-location

echo [4/4] Launching application...
start /b "" "%PY_EXE%" "Program_Files/app_server.py"

echo.
echo ===================================
echo  READY! Opening in browser...
echo ===================================
timeout /t 5
exit
