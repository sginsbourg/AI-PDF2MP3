@echo off
rem -- Batch file to launch the file-to-audio-converter.py script --

rem Change to the directory where this batch file is located
cd /d "%~dp0"

set PATH=%PATH%;C:\Users\sgins\Python312;C:\Users\sgins\Python312\Scripts;C:\ffmpeg\bin;C:\Program Files\Git\bin;C:\Windows\System32;

echo Checking for Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in your PATH. Please install Python.
    echo For Python 3.13+, ensure audioop-lts is installed.
    echo Download Python from https://www.python.org/downloads/
    pause
    exit /b
)

echo.
echo Checking for FFmpeg installation at C:\ffmpeg\bin...
if exist "C:\ffmpeg\bin\ffmpeg.exe" (
    echo FFmpeg found at C:\ffmpeg\bin\ffmpeg.exe
) else (
    echo FFmpeg not found at C:\ffmpeg\bin\ffmpeg.exe
    echo Please install FFmpeg from https://ffmpeg.org/download.html at C:\ffmpeg.
    echo See instructions in file-to-audio-converter.py for details.
    pause
    exit /b
)

rem Define the name of the virtual environment and the script
set VENV_NAME=audio_venv
set SCRIPT_NAME=file-to-audio-converter.py

echo.
echo Checking for virtual environment...
if not exist "%VENV_NAME%\Scripts\activate.bat" (
    echo Virtual environment not found. Creating a new one...
    python -m venv %VENV_NAME%
    if %errorlevel% neq 0 (
        echo Failed to create virtual environment.
        pause
        exit /b
    )
)

echo Activating virtual environment...
call "%VENV_NAME%\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo Failed to activate virtual environment.
    pause
    exit /b
)

echo.
echo Installing required Python libraries...
pip install python-docx PyPDF2 pyttsx3 pydub comtypes pypiwin32 pywin32 audioop-lts
if %errorlevel% neq 0 (
    echo Failed to install required libraries.
    echo Please make sure you have internet access and that pip is working correctly.
    pause
    exit /b
)

echo.
echo Launching the script: %SCRIPT_NAME%
echo.
python "%SCRIPT_NAME%"

echo.
echo Script finished. Press any key to exit.
pause
