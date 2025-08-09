@echo off
rem --- Define Source and Destination Folders ---
set "source_folder=C:\Users\sgins\OneDrive\Documents\GitHub\AI-PDF2MP3"
set "primary_dest_folder=C:\Users\sgins\Downloads\AI-Generated-Audio-Books"
set "secondary_dest_folder=C:\Users\sgins\CrossDevice\Shay's S24+\storage\Audiobooks\AI-Generated-Audio-Books"

echo source_folder          = %source_folder%
echo primary_dest_folder    = %primary_dest_folder%
echo secondary_dest_folder  = %secondary_dest_folder%

rem --- Check if Destination Folders Exist and Create Them if Necessary ---
if not exist "%primary_dest_folder%" (
    echo Primary destination folder does not exist. Creating it...
    mkdir "%primary_dest_folder%"
)
if not exist "%secondary_dest_folder%" (
    echo Secondary destination folder does not exist. Creating it...
    mkdir "%secondary_dest_folder%"
)

rem --- Step 1: Move audio files from source to primary destination ---
echo.
echo Step 1: Moving audio files from source to primary destination...
if exist "%source_folder%\*.mp3" move /Y "%source_folder%\*.mp3" "%primary_dest_folder%\"
if exist "%source_folder%\*.m4a" move /Y "%source_folder%\*.m4a" "%primary_dest_folder%\"
if exist "%source_folder%\*.aac" move /Y "%source_folder%\*.aac" "%primary_dest_folder%\"
if exist "%source_folder%\*.wma" move /Y "%source_folder%\*.wma" "%primary_dest_folder%\"
if exist "%source_folder%\*.ogg" move /Y "%source_folder%\*.ogg" "%primary_dest_folder%\"
if exist "%source_folder%\*.wav" move /Y "%source_folder%\*.wav" "%primary_dest_folder%\"
if exist "%source_folder%\*.flac" move /Y "%source_folder%\*.flac" "%primary_dest_folder%\"

rem --- Step 2: Copy files from primary destination to secondary destination ---
echo.
echo Step 2: Copying audio files from primary destination to secondary destination...
if exist "%primary_dest_folder%\*.mp3" copy /Y "%primary_dest_folder%\*.mp3" "%secondary_dest_folder%\"
if exist "%primary_dest_folder%\*.m4a" copy /Y "%primary_dest_folder%\*.m4a" "%secondary_dest_folder%\"
if exist "%primary_dest_folder%\*.aac" copy /Y "%primary_dest_folder%\*.aac" "%secondary_dest_folder%\"
if exist "%primary_dest_folder%\*.wma" copy /Y "%primary_dest_folder%\*.wma" "%secondary_dest_folder%\"
if exist "%primary_dest_folder%\*.ogg" copy /Y "%primary_dest_folder%\*.ogg" "%secondary_dest_folder%\"
if exist "%primary_dest_folder%\*.wav" copy /Y "%primary_dest_folder%\*.wav" "%secondary_dest_folder%\"
if exist "%primary_dest_folder%\*.flac" copy /Y "%primary_dest_folder%\*.flac" "%secondary_dest_folder%\"

echo.
echo Operation completed.
pause