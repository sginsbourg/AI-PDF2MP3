@echo off
CLS
pushd "%~dp0"
setlocal enabledelayedexpansion
SET VENV_DIR=venv
SET PYTHON_APP=pdf2mp3.py
SET RUNNER_BAT=run_audiobook.bat

echo [*] Starting Ironclad Repository Synchronization...

REM --- Robust Executable Discovery ---
where git >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [!] git is not in PATH. Searching standard locations...
    if exist "C:\Program Files\Git\cmd\git.exe" (
        set "PATH=%PATH%;C:\Program Files\Git\cmd"
        echo [+] Found git at C:\Program Files\Git\cmd
    ) else (
        echo [!] git.exe not found. Skipping sync.
        goto :skip_sync
    )
)

SET "PS_EXE=powershell"
where powershell >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    if exist "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" (
        SET "PS_EXE=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    ) else if exist "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" (
        SET "PS_EXE=C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    )
)

where ffmpeg >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [!] WARNING: ffmpeg not found in PATH. Audio merging will fail.
)

REM --- Git Repository Self-Healing ---
git rev-parse --is-inside-work-tree >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [!] Not a git repository. Initializing local workspace...
    git init
    git branch -M main
    git remote add origin https://github.com/sginsbourg/AI-PDF2MP3.git
    echo [+] Repository initialized and linked to GitHub.
)

REM Ensure GitLab remote exists
git remote | findstr /X "gitlab" >nul
if %ERRORLEVEL% NEQ 0 (
    echo [+] Adding GitLab mirror remote...
    git remote add gitlab https://gitlab.com/sginsbourg/AI-PDF2MP3.git
)

echo [+] Executing Dual-Cloud Sync (GitHub + GitLab)...
git add .
set "SYNC_MSG=Auto-sync from PDF2MP3 Pipeline: %DATE% %TIME%"
git commit -m "!SYNC_MSG!"
echo [+] Pushing to GitHub...
git pull origin main --rebase
git push origin main

echo [+] Pushing to GitLab mirror...
git pull gitlab main --rebase
git push gitlab main

echo [+] Synchronization Complete.

:skip_sync
if not exist %VENV_DIR% (
    echo [*] Creating Virtual Environment...
    python -m venv %VENV_DIR%
)

echo [*] Checking dependencies...
call %VENV_DIR%\Scripts\activate
python -m pip install --upgrade pip >nul
pip install -r requirements.txt >nul

echo.
set "CMD_ARG=%~1"
set "TARGET_PDF=%~1"

if /I "%CMD_ARG%"=="all" (
    echo [*] 'all' parameter detected. Synchronizing entire library...
    goto :batch_scan
)

if "%TARGET_PDF%"=="" (
    echo [*] No PDF specified. Defaulting to full library scan...
    goto :batch_scan
)

:single_file
set "FILE_BASE=%~n1"
call :check_and_process "%TARGET_PDF%" "!FILE_BASE!"
echo.
echo [*] Operation Completed.
timeout /t 5
goto :EOF

:batch_scan
echo [*] Scanning "pdf" folder for updates...
if not exist "pdf" mkdir "pdf"

for %%F in (pdf\*.pdf) do (
    set "FILE_PATH=%%F"
    set "FILE_BASE=%%~nF"
    call :check_and_process "%%F" "!FILE_BASE!"
)
echo.
echo [*] Bulk Operation Completed.
timeout /t 5
goto :EOF

:check_and_process
set "P_PATH=%~1"
set "P_BASE=%~2"
set "MP3_TARGET=mp3\!P_BASE!.mp3"
set "JSON_TARGET=json\!P_BASE!.json"

if not exist "mp3" mkdir "mp3"
if not exist "json" mkdir "json"

REM Skip only if BOTH MP3 and JSON exist AND MP3 is newer than PDF source
set "PS_CMD=%PS_EXE% -NoProfile -Command"
%PS_CMD% "$m=$env:MP3_TARGET; $j=$env:JSON_TARGET; $src=$env:P_PATH; if (-not (Test-Path $m) -or -not (Test-Path $j)) { exit 1 }; if ((Get-Item $m).LastWriteTime -lt (Get-Item $src).LastWriteTime) { exit 1 }; exit 0"

if %ERRORLEVEL% NEQ 0 goto :do_process
echo [.] Skipping "!P_BASE!" (Audiobook is up to date)
goto :EOF

:do_process
echo [+] Generating Audiobook ^& JSON: "!P_BASE!"
%VENV_DIR%\Scripts\python.exe %PYTHON_APP% "!P_PATH!"
goto :EOF
