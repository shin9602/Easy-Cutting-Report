@echo off
setlocal enabledelayedexpansion
title [Cutting Eval Tool] Supreme Launcher v7.6
cd /d "%~dp0"

echo [1/4] Checking version...
set "TOOLS=%~dp0_tools"
set "REPO=shin9602/Easy-Cutting-Report"
set "VERSION_FILE=%~dp0version.txt"

if not exist "%VERSION_FILE%" echo v1.0.0 > "%VERSION_FILE%"
set /p CURRENT_VER=<"%VERSION_FILE%"
echo [*] Local: %CURRENT_VER%

:: [UPDATE LOGIC WITH Y/N]
powershell -NoProfile -ExecutionPolicy Bypass -Command "try{$r=Invoke-RestMethod https://api.github.com/repos/%REPO%/releases/latest -TimeoutSec 3;$r.tag_name}catch{}" > "%TEMP%\cever.txt" 2>nul
set /p LATEST_VER=<"%TEMP%\cever.txt"
del "%TEMP%\cever.txt" >nul 2>&1

if "!LATEST_VER!"=="" goto SKIP_UPDATE
if "!CURRENT_VER!"=="!LATEST_VER!" (
    echo [*] Latest version.
    goto SKIP_UPDATE
)

echo.
echo -------------------------------------------------
echo  NEW VERSION AVAILABLE: !LATEST_VER!
echo -------------------------------------------------
set /p DO_UPDATE=" >> Update to latest version now? [Y/N]: "
if /i not "!DO_UPDATE!"=="Y" (
    echo [*] Skipping update. Starting app...
    goto SKIP_UPDATE
)

echo [*] Downloading and applying update...
set "ZIPFILE=%~dp0_update.zip"
set "TMPDIR=%~dp0_update_temp"
set "ZIPURL=https://github.com/%REPO%/releases/download/!LATEST_VER!/CuttingEval-!LATEST_VER!.zip"

powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12';Invoke-WebRequest '!ZIPURL!' -OutFile '!ZIPFILE!' -UseBasicParsing"
if not exist "!ZIPFILE!" (
    echo [ERROR] Download failed.
    pause & goto SKIP_UPDATE
)

if exist "!TMPDIR!" rmdir /s /q "!TMPDIR!"
mkdir "!TMPDIR!"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive '!ZIPFILE!' '!TMPDIR!' -Force"
robocopy "!TMPDIR!" "%~dp0" /E /XD _data _tools _update_temp /XF version.txt launcher.log /NFL /NDL /NJH /NJS >nul 2>&1

echo !LATEST_VER!> "%VERSION_FILE%"
del "!ZIPFILE!" >nul 2>&1
rmdir /s /q "!TMPDIR!"

echo [SUCCESS] Updated to !LATEST_VER!. Restarting...
timeout /t 2 >nul
start "" "%~f0"
exit

:SKIP_UPDATE
:: [END UPDATE LOGIC]

echo [2/4] Python checking...
set "PY_EXE=python"
python --version >nul 2>&1
if %ERRORLEVEL% equ 0 goto STEP3
set "PY_EXE=%TOOLS%\python\python.exe"
if exist "%PY_EXE%" goto STEP3
echo [!] Python not found. Installing...
if not exist "%TOOLS%" mkdir "%TOOLS%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12';Invoke-WebRequest https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip -OutFile '%TOOLS%\python.zip' -UseBasicParsing"
mkdir "%TOOLS%\python" >nul 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive '%TOOLS%\python.zip' '%TOOLS%\python' -Force"
del "%TOOLS%\python.zip" >nul 2>&1
for /f "delims=" %%F in ('dir /b "%TOOLS%\python\python*._pth" 2^>nul') do powershell -NoProfile -Command "(Get-Content '%TOOLS%\python\%%F') -replace '#import site','import site' | Set-Content '%TOOLS%\python\%%F'"
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12';Invoke-WebRequest https://bootstrap.pypa.io/get-pip.py -OutFile '%TOOLS%\get-pip.py' -UseBasicParsing"
"%PY_EXE%" "%TOOLS%\get-pip.py" --no-warn-script-location >nul 2>&1

:STEP3
echo [3/4] Installing dependencies...
set "PATH=%TOOLS%\python;%TOOLS%\python\Scripts;%PATH%"
"%PY_EXE%" -m pip install flask flask-cors pillow openpyxl xlrd==1.2.0 pywin32 --quiet --no-warn-script-location

echo [4/4] Launching application...
start /b "" "%PY_EXE%" "Program_Files/app_server.py"

echo.
echo ===================================
echo  READY!
echo ===================================
timeout /t 5
exit
