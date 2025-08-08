@echo off
rem This batch file syncs audio files from one folder to another using Robocopy.

rem Define the source and destination folders.
set "source_folder=C:\Users\sgins\Downloads\AI-Generated-Audio-Books"
set "destination_folder=C:\Users\sgins\CrossDevice\Shay's S24+\storage\Audiobooks\AI-Generated-Audio-Books"

rem Define the full path to the robocopy executable.
set "robocopy_exe=C:\Windows\System32\robocopy.exe"

rem Check if the source folder exists before running the sync.
if not exist "%source_folder%" (
    echo Source folder does not exist: %source_folder%
    pause
    exit /b
)

rem Robocopy command with a set of useful options:
rem The list of file extensions (*.mp3, *.m4a, etc.) restricts the sync to only these types.
rem /E   : Copies subdirectories, including empty ones.
rem /MIR : Mirrors a directory tree (equivalent to /E and /PURGE), but only for the specified files.
rem        This means it will copy new audio files, update changed audio files,
rem        and delete audio files from the destination that are no longer in the source.
rem /R:1 : Retries failed copies 1 time.
rem /W:1 : Waits 1 second between retries.
rem /NP  : No Progress - prevents Robocopy from showing the percentage.
rem /LOG:sync_log.txt : Writes the output to a log file instead of the console.
echo Syncing only audio files...
"%robocopy_exe%" "%source_folder%" "%destination_folder%" *.mp3 *.m4a *.aac *.flac *.wma *.ogg *.wav /MIR /R:1 /W:1 /NP /LOG:sync_log.txt

echo Sync complete. Check sync_log.txt for details.
pause
