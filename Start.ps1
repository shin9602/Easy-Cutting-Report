# 절삭평가 자동화 도구 런처 (PS 5.1 호환)
# LAUNCHER_ROOT: Start.bat에서 환경변수로 주입된 경로 (한글 경로 깨짐 방지)
$Root = if ($env:LAUNCHER_ROOT -and (Test-Path $env:LAUNCHER_ROOT)) {
    $env:LAUNCHER_ROOT.TrimEnd('\')
} elseif ($PSScriptRoot) {
    $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    Split-Path -Parent $MyInvocation.MyCommand.Path
} else {
    (Get-Location).Path
}
$AppDir = Join-Path $Root "절삭평가_App"
$ToolsDir = Join-Path $AppDir "_tools"
$PythonDir = Join-Path $ToolsDir "python"
$VersionFile = Join-Path $AppDir "version.txt"
$ServerScript = Join-Path $AppDir "Program_Files\app_server.py"
$LogFile = Join-Path $Root "launcher.log"
$Repo = "shin9602/Easy-Cutting-Report"

# 콘솔(launcher.log) 로그 기록 함수
function Log($msg) {
    try { Add-Content -Path $LogFile -Value ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $msg) -Encoding UTF8 } catch {}
}

[Net.ServicePointManager]::SecurityProtocol = 'Tls12'

# 한글 콘솔 출력을 위한 UTF-8 설정 (PS 5.1 호환)
try { chcp 65001 | Out-Null } catch {}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
try { $OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}

# === 시작 배너 ===
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "   절삭평가 자동화 도구" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

Log "[BOOT] Root=$Root | PS=$($PSVersionTable.PSVersion) | Script=$($MyInvocation.MyCommand.Path)"
Log "[1/4] Checking version..."

# 버전 파일 없으면 생성
if (-not (Test-Path $VersionFile)) {
    "v1.0.0" | Out-File -FilePath $VersionFile -Encoding ascii -NoNewline
}
$CurrentVer = ((Get-Content $VersionFile) -join "").Trim()
Log "[*] Local: $CurrentVer"

Write-Host " 현재 버전: $CurrentVer" -ForegroundColor White
Write-Host " 최신 버전 확인 중..." -ForegroundColor Gray

# 최신 버전 확인
try {
    $Release = Invoke-RestMethod "https://api.github.com/repos/$Repo/releases/latest" -TimeoutSec 3
    $LatestVer = if ($Release -and $Release.tag_name) { ([string]$Release.tag_name).Trim() } else { "" }
} catch {
    Log "[WARN] GitHub API failed: $_"
    Write-Host " [알림] 버전 확인 실패 - 오프라인으로 실행합니다." -ForegroundColor Yellow
    $LatestVer = ""
}

if ($LatestVer -ne "" -and $CurrentVer -ne $LatestVer) {
    Log "[*] New version available: $LatestVer (auto-updating)"
    Write-Host " 새 버전 발견: $CurrentVer → $LatestVer" -ForegroundColor Yellow
    Write-Host " 자동 업데이트를 시작합니다..." -ForegroundColor Yellow

    $ZipFile = Join-Path $Root "_update.zip"
    $TmpDir  = Join-Path $Root "_update_temp"
    $ZipUrl  = "https://github.com/$Repo/releases/download/$LatestVer/CuttingEval-$LatestVer.zip"

    $UpdateOk = $true
    Write-Host " 다운로드 중... ($LatestVer)" -ForegroundColor Gray
    try {
        Invoke-WebRequest $ZipUrl -OutFile $ZipFile -UseBasicParsing -TimeoutSec 60
    } catch {
        Log "[ERROR] Download failed: $_"
        Write-Host " [경고] 업데이트 실패 - 현재 버전으로 실행합니다." -ForegroundColor Red
        if (Test-Path $ZipFile) { Remove-Item $ZipFile -Force -ErrorAction SilentlyContinue }
        $UpdateOk = $false
    }

    if ($UpdateOk) {
        if (-not (Test-Path $ZipFile) -or (Get-Item $ZipFile).Length -lt 1024) {
            Log "[ERROR] Zip corrupted."
            Write-Host " [경고] 업데이트 실패 - 현재 버전으로 실행합니다." -ForegroundColor Red
            if (Test-Path $ZipFile) { Remove-Item $ZipFile -Force -ErrorAction SilentlyContinue }
        } else {
            if (Test-Path $TmpDir) { Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue }

            $ExtractOk = $true
            Write-Host " 설치 중..." -ForegroundColor Gray
            try {
                Expand-Archive $ZipFile $TmpDir -Force -ErrorAction Stop
            } catch {
                Log "[ERROR] Expand-Archive failed: $_"
                Write-Host " [경고] 업데이트 실패 - 현재 버전으로 실행합니다." -ForegroundColor Red
                if (Test-Path $ZipFile) { Remove-Item $ZipFile -Force -ErrorAction SilentlyContinue }
                if (Test-Path $TmpDir) { Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue }
                $ExtractOk = $false
            }

            if ($ExtractOk) {
                # ZIP 구조: 루트에 Start.bat/Start.ps1 + 절삭평가_App/ 폴더
                # $Root (Start.bat 위치) 전체를 업데이트하되, 보호 항목 제외
                robocopy $TmpDir $Root /E /XD _data _tools _update_temp _workspace .git .claude Harness_Plugin /XF version.txt launcher.log /NFL /NDL /NJH /NJS | Out-Null
                $rc = $LASTEXITCODE
                if ($rc -ge 8) {
                    Log "[ERROR] robocopy failed (exit=$rc). Skip version bump."
                    Write-Host " [경고] 업데이트 실패 - 현재 버전으로 실행합니다." -ForegroundColor Red
                    if (Test-Path $ZipFile) { Remove-Item $ZipFile -Force -ErrorAction SilentlyContinue }
                    if (Test-Path $TmpDir)  { Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue }
                } else {
                    $VersionWritten = $true
                    try {
                        $LatestVer | Out-File -FilePath $VersionFile -Encoding ascii -NoNewline -ErrorAction Stop
                        Log "[*] Version file updated to $LatestVer"
                    } catch {
                        Log "[ERROR] Failed to update version file: $_"
                        Write-Host " [경고] 업데이트 실패 - 현재 버전으로 실행합니다." -ForegroundColor Red
                        $VersionWritten = $false
                    }

                    Remove-Item $ZipFile -Force -ErrorAction SilentlyContinue
                    Remove-Item $TmpDir -Recurse -Force -ErrorAction SilentlyContinue

                    if ($VersionWritten) {
                        Log "[SUCCESS] Updated to $LatestVer (rc=$rc). Restarting..."
                        Write-Host " 업데이트 완료! 재시작합니다..." -ForegroundColor Green
                        Start-Sleep -Seconds 1
                        # 재실행: $Root는 이미 확정된 경로이므로 항상 사용 가능
                        $ScriptPath = Join-Path $Root "Start.ps1"
                        $BatPath    = Join-Path $Root "Start.bat"
                        if (Test-Path $ScriptPath) {
                            Log "[*] Restarting via: $ScriptPath"
                            Start-Process powershell -ArgumentList "-NoProfile","-ExecutionPolicy","Bypass","-File","`"$ScriptPath`""
                            exit
                        } elseif (Test-Path $BatPath) {
                            Log "[*] Restarting via bat: $BatPath"
                            Start-Process "cmd.exe" -ArgumentList "/c `"$BatPath`""
                            exit
                        } else {
                            Log "[ERROR] Cannot restart: neither Start.ps1 nor Start.bat found at $Root"
                            Write-Host " [경고] 재시작 파일을 찾을 수 없습니다. Start.bat 을 다시 실행해 주세요." -ForegroundColor Red
                        }
                    }
                }
            }
        }
    }
} else {
    Log "[*] Latest version."
    if ($LatestVer -ne "") {
        Write-Host " 최신 버전입니다. ($CurrentVer)" -ForegroundColor Green
    }
}

Write-Host ""
Write-Host " Python 확인 중..." -ForegroundColor Gray
Log "[2/4] Python checking..."

# Python 탐색 순서: 시스템 Python → 내장 Python
$PyExe  = $null
$PywExe = $null
$PyDir  = $null

$SysCmd = Get-Command python -ErrorAction SilentlyContinue
$SysPython = if ($SysCmd) { $SysCmd.Source } else { $null }
if ($SysPython) {
    $PyDir  = Split-Path $SysPython
    $PyExe  = $SysPython
    $PywCandidate = Join-Path $PyDir "pythonw.exe"
    $PywExe = if (Test-Path $PywCandidate) { $PywCandidate } else { $SysPython }
    Log "[*] System Python: $PyExe"
} elseif (Test-Path (Join-Path $PythonDir "python.exe")) {
    $PyDir  = $PythonDir
    $PyExe  = Join-Path $PythonDir "python.exe"
    $PywExe = Join-Path $PythonDir "pythonw.exe"
    if (-not (Test-Path $PywExe)) { $PywExe = $PyExe }
    Log "[*] Embedded Python: $PyExe"
} else {
    Log "[!] Python not found. Installing embedded Python 3.11..."
    Write-Host " Python 미설치 - 내장 Python 3.11 을 설치합니다 (수 분 소요)..." -ForegroundColor Yellow
    if (-not (Test-Path $ToolsDir)) { New-Item -ItemType Directory $ToolsDir | Out-Null }
    $PythonZip = Join-Path $ToolsDir "python.zip"
    [Net.ServicePointManager]::SecurityProtocol = 'Tls12'
    Invoke-WebRequest "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip" -OutFile $PythonZip -UseBasicParsing
    if (-not (Test-Path $PythonDir)) { New-Item -ItemType Directory $PythonDir | Out-Null }
    Expand-Archive $PythonZip $PythonDir -Force
    Remove-Item $PythonZip -Force

    # enable site-packages
    Get-ChildItem "$PythonDir\python*._pth" | ForEach-Object {
        (Get-Content $_.FullName) -replace '#import site','import site' | Set-Content $_.FullName
    }

    # install pip
    $GetPip = Join-Path $ToolsDir "get-pip.py"
    Invoke-WebRequest "https://bootstrap.pypa.io/get-pip.py" -OutFile $GetPip -UseBasicParsing
    & $PythonDir\python.exe $GetPip --no-warn-script-location | Out-Null
    Remove-Item $GetPip -Force

    $PyDir  = $PythonDir
    $PyExe  = Join-Path $PythonDir "python.exe"
    $PywExe = Join-Path $PythonDir "pythonw.exe"
    if (-not (Test-Path $PywExe)) { $PywExe = $PyExe }
}

Write-Host ""
Write-Host " 패키지 설치 중..." -ForegroundColor Gray
Log "[3/4] Installing dependencies..."
if ($PyDir) {
    $env:PATH = "$PyDir;$PyDir\Scripts;$env:PATH"
}
try { & $PyExe -m pip install flask flask-cors pillow openpyxl "xlrd==1.2.0" pywin32 --quiet --no-warn-script-location *> $null } catch {}

$PostInstall = Join-Path $PyDir "Scripts\pywin32_postinstall.py"
if (Test-Path $PostInstall) {
    try { & $PyExe $PostInstall -install *> $null } catch {}
}

Write-Host ""
Write-Host " 앱 실행 중..." -ForegroundColor Gray
Log "[4/4] Launching application..."
Start-Process $PywExe -ArgumentList "`"$ServerScript`"" -WindowStyle Hidden

Log "[READY] App is running."
Write-Host ""
Write-Host " 앱이 실행되었습니다. 브라우저에서 http://localhost:5000 으로 접속하세요." -ForegroundColor Green
Write-Host " 이 창은 닫아도 됩니다." -ForegroundColor Gray
Start-Sleep -Seconds 3
