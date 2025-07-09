@echo off
setlocal EnableDelayedExpansion

:: Set log file for debugging
set LOG_FILE=C:\Users\sgins\OneDrive\Documents\GitHub\AI-PDF2MP3\run_log.log
:: Delete existing log file if it exists
if exist "%LOG_FILE%" (
    del "%LOG_FILE%"
    echo %DATE% %TIME% - Deleted existing %LOG_FILE% >> "%LOG_FILE%"
)
echo %DATE% %TIME% - Starting batch file >> "%LOG_FILE%"

:: Set paths
set PYTHON_BASE=C:\Python313\python.exe
set VENV_DIR=C:\Users\sgins\OneDrive\Documents\GitHub\AI-PDF2MP3\venv
set PYTHON_SCRIPT=C:\Users\sgins\OneDrive\Documents\GitHub\AI-PDF2MP3\pdf_to_speech_ocr.py
set POPPLER_PATH=C:\Program Files\poppler\bin
set TESSERACT_PATH=C:\Program Files\Tesseract-OCR
set PDF_FILE=C:\Users\sgins\OneDrive\Documents\GitHub\AI-PDF2MP3\grok_report_2.pdf
set MP3_FILE=C:\Users\sgins\OneDrive\Documents\GitHub\AI-PDF2MP3\grok_report_2_speech.mp3
set AUDIO_BOOKS_DIR=C:\Users\sgins\Downloads\AI-Generated-Audio-Books
set MP3_DEST=%AUDIO_BOOKS_DIR%\grok_report_2_speech.mp3

:: Set working directory
echo %DATE% %TIME% - Setting working directory to C:\Users\sgins\OneDrive\Documents\GitHub\AI-PDF2MP3 >> "%LOG_FILE%"
cd /d C:\Users\sgins\OneDrive\Documents\GitHub\AI-PDF2MP3
if %ERRORLEVEL% neq 0 (
    echo %DATE% %TIME% - Failed to set working directory >> "%LOG_FILE%"
    echo Failed to set working directory to C:\Users\sgins\OneDrive\Documents\GitHub\AI-PDF2MP3
    pause
    exit /b 1
)

:: Add Poppler and Tesseract to PATH
set PATH=%PATH%;%POPPLER_PATH%;%TESSERACT_PATH%
echo %DATE% %TIME% - Updated PATH with Poppler and Tesseract >> "%LOG_FILE%"

:: Check if base Python exists
if not exist "%PYTHON_BASE%" (
    echo %DATE% %TIME% - Python not found at %PYTHON_BASE% >> "%LOG_FILE%"
    echo Python not found at %PYTHON_BASE%. Please install Python 3.13 to C:\Python313.
    echo Run 'where python' to find installed Python versions or download from https://www.python.org/downloads/.
    pause
    exit /b 1
)

:: Check base Python version
"%PYTHON_BASE%" --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo %DATE% %TIME% - Python executable at %PYTHON_BASE% is invalid >> "%LOG_FILE%"
    echo Invalid Python executable at %PYTHON_BASE%. Ensure Python 3.13 is installed correctly.
    pause
    exit /b 1
)
echo %DATE% %TIME% - Base Python version check passed >> "%LOG_FILE%"

:: Create virtual environment if it doesn't exist
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo %DATE% %TIME% - Creating virtual environment at %VENV_DIR% >> "%LOG_FILE%"
    "%PYTHON_BASE%" -m venv "%VENV_DIR%"
    if %ERRORLEVEL% neq 0 (
        echo %DATE% %TIME% - Failed to create virtual environment >> "%LOG_FILE%"
        echo Failed to create virtual environment at %VENV_DIR%.
        pause
        exit /b 1
    )
    echo %DATE% %TIME% - Virtual environment created successfully >> "%LOG_FILE%"
)

:: Activate virtual environment
echo %DATE% %TIME% - Activating virtual environment >> "%LOG_FILE%"
call "%VENV_DIR%\Scripts\activate.bat"
if %ERRORLEVEL% neq 0 (
    echo %DATE% %TIME% - Failed to activate virtual environment >> "%LOG_FILE%"
    echo Failed to activate virtual environment at %VENV_DIR%.
    pause
    exit /b 1
)
echo %DATE% %TIME% - Virtual environment activated >> "%LOG_FILE%"

:: Upgrade pip in virtual environment
echo %DATE% %TIME% - Upgrading pip in virtual environment >> "%LOG_FILE%"
python -m pip install --upgrade pip >> "%LOG_FILE%" 2>&1
if %ERRORLEVEL% neq 0 (
    echo %DATE% %TIME% - Failed to upgrade pip >> "%LOG_FILE%"
    echo Failed to upgrade pip in virtual environment.
    pause
    exit /b 1
)

:: Install dependencies in virtual environment
echo %DATE% %TIME% - Installing dependencies in virtual environment >> "%LOG_FILE%"
set DEPENDENCIES=PyPDF2 pdf2image pytesseract pillow gTTS colorama tqdm pyttsx3
pip install %DEPENDENCIES% >> "%LOG_FILE%" 2>&1
if %ERRORLEVEL% neq 0 (
    echo %DATE% %TIME% - Failed to install dependencies: %DEPENDENCIES% >> "%LOG_FILE%"
    echo Failed to install dependencies. Check %LOG_FILE% for details.
    pause
    exit /b 1
)
echo %DATE% %TIME% - Dependencies installed successfully >> "%LOG_FILE%"

:: Check Python dependencies individually
echo %DATE% %TIME% - Checking Python dependencies >> "%LOG_FILE%"
set DEPENDENCY_ERROR=0
for %%M in (PyPDF2 pdf2image pytesseract PIL.Image pyttsx3) do (
    python -c "import %%M" 2>>"%LOG_FILE%.tmp"
    if !ERRORLEVEL! neq 0 (
        echo %DATE% %TIME% - Failed to import %%M >> "%LOG_FILE%"
        type "%LOG_FILE%.tmp" >> "%LOG_FILE%"
        echo Failed to import %%M
        set DEPENDENCY_ERROR=1
    ) else (
        echo %DATE% %TIME% - Successfully imported %%M >> "%LOG_FILE%"
    )
    del "%LOG_FILE%.tmp" 2>nul
)
:: Special handling for gTTS case sensitivity
python -c "import gTTS" 2>>"%LOG_FILE%.tmp"
if !ERRORLEVEL% neq 0 (
    python -c "import gtts" 2>>"%LOG_FILE%.tmp"
    if !ERRORLEVEL! neq 0 (
        echo %DATE% %TIME% - Failed to import gTTS or gtts >> "%LOG_FILE%"
        type "%LOG_FILE%.tmp" >> "%LOG_FILE%"
        echo Failed to import gTTS or gtts
        set DEPENDENCY_ERROR=1
    ) else (
        echo %DATE% %TIME% - Successfully imported gtts >> "%LOG_FILE%"
    )
) else (
    echo %DATE% %TIME% - Successfully imported gTTS >> "%LOG_FILE%"
)
del "%LOG_FILE%.tmp" 2>nul

if %DEPENDENCY_ERROR% equ 1 (
    echo %DATE% %TIME% - Missing Python dependencies. Run: pip install %DEPENDENCIES% in the virtual environment >> "%LOG_FILE%"
    echo Missing Python dependencies. Activate the virtual environment with '%VENV_DIR%\Scripts\activate.bat' and run: pip install %DEPENDENCIES%
    pause
    exit /b 1
)

:: Check if the Python script exists
if not exist "%PYTHON_SCRIPT%" (
    echo %DATE% %TIME% - Python script not found at %PYTHON_SCRIPT% >> "%LOG_FILE%"
    echo Python script "%PYTHON_SCRIPT%" not found.
    pause
    exit /b 1
)

:: Check if Poppler is accessible
if not exist "%POPPLER_PATH%\pdftoppm.exe" (
    echo %DATE% %TIME% - Poppler executable not found at %POPPLER_PATH%\pdftoppm.exe >> "%LOG_FILE%"
    echo Poppler executable not found at %POPPLER_PATH%\pdftoppm.exe. Ensure Poppler is installed correctly.
    pause
    exit /b 1
)
pdftoppm -v >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo %DATE% %TIME% - Poppler not found or not in PATH >> "%LOG_FILE%"
    echo Poppler not found or not in PATH. Ensure Poppler is installed at %POPPLER_PATH%.
    pause
    exit /b 1
)

:: Check if Tesseract is accessible
tesseract --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo %DATE% %TIME% - Tesseract not found or not in PATH >> "%LOG_FILE%"
    echo Tesseract not found or not in PATH. Ensure Tesseract is installed at %TESSERACT_PATH%.
    pause
    exit /b 1
)

:: Run the Python script with a timeout (300 seconds = 5 minutes)
echo %DATE% %TIME% - Running %PYTHON_SCRIPT% >> "%LOG_FILE%"
echo Running %PYTHON_SCRIPT%...
C:\Windows\System32\timeout.exe /t 300 /nobreak >nul & python "%PYTHON_SCRIPT%"
if %ERRORLEVEL% neq 0 (
    echo %DATE% %TIME% - Error occurred while running the Python script or timed out >> "%LOG_FILE%"
    echo Error occurred while running the Python script or timed out after 5 minutes. Check %LOG_FILE% for details.
    pause
    exit /b 1
)

:: Create the AI-Generated-Audio-Books directory if it doesn't exist
if not exist "%AUDIO_BOOKS_DIR%" (
    echo %DATE% %TIME% - Creating directory %AUDIO_BOOKS_DIR% >> "%LOG_FILE%"
    mkdir "%AUDIO_BOOKS_DIR%"
    if %ERRORLEVEL% neq 0 (
        echo %DATE% %TIME% - Failed to create directory %AUDIO_BOOKS_DIR% >> "%LOG_FILE%"
        echo Failed to create directory %AUDIO_BOOKS_DIR%
        pause
        exit /b 1
    )
    echo %DATE% %TIME% - Directory %AUDIO_BOOKS_DIR% created successfully >> "%LOG_FILE%"
)

:: Move the generated MP3 file to AI-Generated-Audio-Books directory
if exist "%MP3_FILE%" (
    echo %DATE% %TIME% - Moving MP3 file to %MP3_DEST% >> "%LOG_FILE%"
    echo Moving MP3 file to %MP3_DEST%
    move "%MP3_FILE%" "%MP3_DEST%" >nul
    if %ERRORLEVEL% neq 0 (
        echo %DATE% %TIME% - Failed to move MP3 file to %MP3_DEST% >> "%LOG_FILE%"
        echo Failed to move MP3 file to %MP3_DEST%
        pause
        exit /b 1
    )
    echo %DATE% %TIME% - Successfully moved MP3 file to %MP3_DEST% >> "%LOG_FILE%"
) else (
    echo %DATE% %TIME% - MP3 file not found: %MP3_FILE% >> "%LOG_FILE%"
    echo MP3 file not found: %MP3_FILE%
    pause
    exit /b 1
)

:: Launch the moved MP3 file
if exist "%MP3_DEST%" (
    echo %DATE% %TIME% - Launching MP3 file: %MP3_DEST% >> "%LOG_FILE%"
    echo Launching MP3 file: %MP3_DEST%
    start "" "%MP3_DEST%"
    if %ERRORLEVEL% neq 0 (
        echo %DATE% %TIME% - Failed to launch MP3 file: %MP3_DEST% >> "%LOG_FILE%"
        echo Failed to launch MP3 file: %MP3_DEST%
    ) else (
        echo %DATE% %TIME% - Successfully launched MP3 file >> "%LOG_FILE%"
    )
) else (
    echo %DATE% %TIME% - Moved MP3 file not found: %MP3_DEST% >> "%LOG_FILE%"
    echo Moved MP3 file not found: %MP3_DEST%
    pause
    exit /b 1
)

:: Launch the input PDF file
if exist "%PDF_FILE%" (
    echo %DATE% %TIME% - Launching PDF file: %PDF_FILE% >> "%LOG_FILE%"
    echo Launching PDF file: %PDF_FILE%
    start "" "%PDF_FILE%"
    if %ERRORLEVEL% neq 0 (
        echo %DATE% %TIME% - Failed to launch PDF file: %PDF_FILE% >> "%LOG_FILE%"
        echo Failed to launch PDF file: %PDF_FILE%
    ) else (
        echo %DATE% %TIME% - Successfully launched PDF file >> "%LOG_FILE%"
    )
) else (
    echo %DATE% %TIME% - PDF file not found: %PDF_FILE% >> "%LOG_FILE%"
    echo PDF file not found: %PDF_FILE%
    pause
    exit /b 1
)

:: Deactivate virtual environment
echo %DATE% %TIME% - Deactivating virtual environment >> "%LOG_FILE%"
call "%VENV_DIR%\Scripts\deactivate.bat" 2>nul
echo %DATE% %TIME% - Script completed successfully >> "%LOG_FILE%"

:: Pause to allow user to see output before exiting
pause
endlocal
