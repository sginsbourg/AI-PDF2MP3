@echo off
SETLOCAL

:: %~dp0 refers to the directory where this batch file is located
set "SOURCE_DIR=%~dp0"

:: Target directories
set "DEST_DIR_SD=D:\pdf2mp3"
set "DEST_DIR_G=G:\My Drive\mygit\WPS-Media\pdf2mp3"

echo ========================================================
echo Mirroring repository to Local Targets
echo Source: %SOURCE_DIR%
echo Target 1 (SD): %DEST_DIR_SD%
echo Target 2 (G-Drive): %DEST_DIR_G%
echo Excluding: venv, .git
echo ========================================================
echo.

:: Execute robocopy for SD Card
:: /MIR  - Mirror a directory tree (equivalent to /E plus /PURGE)
:: /XD   - Exclude directories matching the given names or paths
:: /XJ   - Exclude junction points (prevents infinite loops)
:: /R:1  - 1 retry on failed copies
:: /W:1  - 1 second wait between retries
echo [+] Mirroring to SD Card...
robocopy "%SOURCE_DIR%." "%DEST_DIR_SD%" /MIR /XD venv .git /XJ /R:1 /W:1

:: Robocopy exit codes: 0-7 are success/info, >=8 are errors
if errorlevel 8 (
    echo.
    echo [ERROR] SD Mirroring failed with error level %ERRORLEVEL%.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [+] Mirroring to Google Drive...
robocopy "%SOURCE_DIR%." "%DEST_DIR_G%" /MIR /XD venv .git /XJ /R:1 /W:1

if errorlevel 8 (
    echo.
    echo [ERROR] G-Drive Mirroring failed with error level %ERRORLEVEL%.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo [SUCCESS] All Local Mirroring completed successfully.

echo ========================================================
echo Starting Cloud Repository Synchronization
echo ========================================================
echo.

where git >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [!] git is not in PATH. Searching standard locations...
    if exist "C:\Program Files\Git\cmd\git.exe" (
        set "PATH=%PATH%;C:\Program Files\Git\cmd"
        echo [+] Found git at C:\Program Files\Git\cmd
    ) else if exist "C:\Program Files\Git\bin\git.exe" (
        set "PATH=%PATH%;C:\Program Files\Git\bin"
        echo [+] Found git at C:\Program Files\Git\bin
    ) else (
        echo [!] git.exe still not found. Skipping cloud sync.
        goto :skip_sync
    )
)

echo [+] Executing Dual-Cloud Sync (GitHub + GitLab)...
git add .
set "SYNC_MSG=Auto-sync from SD mirror script: %DATE% %TIME%"
:: Using late expansion syntax or call to safely execute message
call git commit -m "%%SYNC_MSG%%"
git pull origin main --rebase
git push origin main

:: Check if gitlab remote exists before pushing
git remote get-url gitlab >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [+] Configuring GitLab remote...
    git remote add gitlab "https://sginsbourg%%40gmail.com:Xx11gd12@gitlab.com/sginsbourg/pdf2mp3.git"
)
echo [+] Syncing to GitLab mirror...
git pull gitlab main --rebase
git push gitlab main
echo [+] Cloud Synchronization Complete.

:skip_sync
timeout /t 10
exit /b 0

ENDLOCAL
