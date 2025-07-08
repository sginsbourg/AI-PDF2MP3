@echo off
setlocal EnableDelayedExpansion

:: Set current directory to the batch file's location
cd /d "%~dp0"
echo Current directory set to: %CD%

:: Set the path to Python executable and script
set PYTHON_EXE=C:\Python313\python.exe
set PYTHON_SCRIPT=pdf_to_speech_ocr.py

:: Check if Python executable exists
if not exist "%PYTHON_EXE%" (
    echo Python not found at %PYTHON_EXE%. Please verify the installation path.
    pause
    exit /b 1
)

:: Check if the Python script exists
if not exist "%PYTHON_SCRIPT%" (
    echo Python script "%PYTHON_SCRIPT%" not found in %CD%.
    pause
    exit /b 1
)

:: Run the Python script
echo Running %PYTHON_SCRIPT%...
"%PYTHON_EXE%" "%PYTHON_SCRIPT%"
if %ERRORLEVEL% neq 0 (
    echo Error occurred while running the Python script.
    pause
    exit /b 1
)

:: Pause to view output
REM pause
endlocal