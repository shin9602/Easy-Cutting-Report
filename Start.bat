@echo off
setlocal enabledelayedexpansion
title [Cutting Eval Tool] Unified Launcher + Auto Updater
cd /d "%~dp0"

echo =================================================
echo  Cutting Evaluation Tool - Launcher v5.5
echo =================================================
echo.

:: [0] Auto-Update System (GitHub)
:: ---------------------------------------------------
:: ⚠️ 주의: 레포지토리 주소를 실제 GitHub 주소로 변경하세요!
set "REPO=shin9602/Easy-Cutting-Report"
set "VERSION_FILE=%~dp0version.txt"

if not exist "%VERSION_FILE%" echo v0.0.0 > "%VERSION_FILE%"
set /p CURRENT_VER=<"%VERSION_FILE%"
echo [*] Local Version: %CURRENT_VER%

echo [*] Checking for updates...
powershell -NoProfile -NonInteractive -Command "try{$r=Invoke-RestMethod https://api.github.com/repos/%REPO%/releases/latest -TimeoutSec 5;$r.tag_name}catch{}" > "%TEMP%\cever.txt" 2>nul
set /p LATEST_VER=<"%TEMP%\cever.txt"
del "%TEMP%\cever.txt" >nul 2>&1

if "!LATEST_VER!"=="" (
    echo [SKIP] Could not reach GitHub or no releases found.
    goto APP_START
)

if "!CURRENT_VER!"=="!LATEST_VER!" (
    echo [OK] Already on the latest version.
    goto APP_START
)

echo.
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
echo  [NEW VERSION AVAILABLE] !LATEST_VER!
echo ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
set /p DO_UPDATE=" >> Update to latest version now? [Y/N]: "
if /i not "!DO_UPDATE!"=="Y" goto APP_START

:: Download & Update
set "ZIPURL=https://github.com/%REPO%/releases/download/!LATEST_VER!/CuttingEval-!LATEST_VER!.zip"
set "ZIPFILE=%~dp0_update.zip"
set "TMPDIR=%~dp0_update_temp"

echo [*] Downloading update...
powershell -NoProfile -NonInteractive -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12';Invoke-WebRequest '!ZIPURL!' -OutFile '!ZIPFILE!' -UseBasicParsing"
if not exist "!ZIPFILE!" (
    echo [ERROR] Download failed.
    pause
    goto APP_START
)

echo [*] Extracting and applying update...
if exist "!TMPDIR!" rmdir /s /q "!TMPDIR!"
mkdir "!TMPDIR!"
powershell -NoProfile -NonInteractive -Command "Expand-Archive '!ZIPFILE!' '!TMPDIR!' -Force"

:: Copy files (excluding data and logs)
robocopy "!TMPDIR!" "%~dp0" /E /XD _data _update_temp /XF version.txt launcher.log /NFL /NDL /NJH /NJS >nul 2>&1

echo !LATEST_VER!> "%VERSION_FILE%"
del "!ZIPFILE!" >nul 2>&1
rmdir /s /q "!TMPDIR!"

echo.
echo [SUCCESS] Updated to !LATEST_VER!. Restarting...
timeout /t 2 >nul
start "" "%~f0"
exit

:APP_START
:: ---------------------------------------------------

:: [1] Check Python
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python not found!
    echo Please install Python from https://www.python.org/
    pause
    exit
)

:: [2] Auto Install Dependencies
echo [*] Checking environment and installing packages...
python -m pip install flask flask-cors pillow openpyxl xlrd==1.2.0 pywin32 pyinstaller --quiet

:: [3] Build Standalone EXE
if not exist "CuttingEval_App.exe" (
    echo.
    echo [*] Creating standalone EXE... (Please wait)
    
    pyinstaller --onefile --noconsole ^
    --add-data "Program_Files/index.html;." ^
    --name "CuttingEval_App" ^
    "Program_Files/app_server.py" >nul 2>&1
    
    if exist "dist/CuttingEval_App.exe" (
        move /y "dist/CuttingEval_App.exe" "CuttingEval_App.exe" >nul
        rmdir /s /q "build" "dist" >nul
        del /q "CuttingEval_App.spec" >nul
        echo [SUCCESS] CuttingEval_App.exe ready!
    )
)

:: [4] Run
echo.
echo Launching Application...
if exist "CuttingEval_App.exe" (
    start "" "CuttingEval_App.exe"
) else (
    echo [WARN] EXE failed. Running script mode...
    start "" python "Program_Files/app_server.py"
)

echo.
echo =================================================
echo  System is active.
echo =================================================
timeout /t 3 >nul
exit
