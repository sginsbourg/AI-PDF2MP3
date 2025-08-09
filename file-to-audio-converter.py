import os
import sys
import pyttsx3
import PyPDF2
from docx import Document
from pydub import AudioSegment

# Install these libraries using pip before running the script:
# pip install python-docx
# pip install PyPDF2
# pip install pyttsx3
# pip install pydub

# To ensure the script works on Windows 11 64-bit, you need to install FFmpeg.
# FFmpeg is a crucial tool that pydub uses for converting audio formats.
# Here are the steps to install it on Windows:
# 1. Go to the FFmpeg download page: https://ffmpeg.org/download.html
# 2. Click the Windows icon and then select "git" or "release" build.
# 3. Download the .zip file for your architecture (e.g., "win64-gpl").
# 4. Extract the contents of the zip file to a location on your computer, for example, C:\ffmpeg.
# 5. Open the Windows Start Menu, search for "Environment Variables", and click "Edit the system environment variables".
# 6. In the "System Properties" window, click the "Environment Variables..." button.
# 7. Under "System variables", find and select the "Path" variable, then click "Edit...".
# 8. Click "New" and add the path to the `bin` folder inside your FFmpeg directory (e.g., C:\ffmpeg\bin).
# 9. Click "OK" on all windows to save the changes.
# 10. Restart your terminal or command prompt for the changes to take effect.
# 11. You can verify the installation by typing `ffmpeg -version` in your terminal.

def extract_text_from_file(file_path):
    """
    Extracts text from a TXT, DOCX, or PDF file.
    """
    file_extension = os.path.splitext(file_path)[1].lower()
    text = ""
    try:
        if file_extension == ".txt":
            with open(file_path, 'r', encoding='utf-8') as f:
                text = f.read()
        elif file_extension in [".docx", ".doc"]:
            document = Document(file_path)
            for paragraph in document.paragraphs:
                text += paragraph.text + "\n"
        elif file_extension == ".pdf":
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
        else:
            print(f"Error: Unsupported file type '{file_extension}'.")
            return None
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while reading the file: {e}")
        return None

    return text

def set_voice_to_uk_male(engine):
    """
    Attempts to set the pyttsx3 engine's voice to a UK male voice.
    """
    voices = engine.getProperty('voices')
    found_voice = False
    for voice in voices:
        # Check for UK locale and male gender.
        if "en-gb" in voice.id.lower() and "male" in voice.id.lower():
            engine.setProperty('voice', voice.id)
            print(f"Using UK male voice: {voice.name}")
            found_voice = True
            break
    
    if not found_voice:
        print("A UK male voice could not be found. Using the default system voice.")

def convert_text_to_audio(text, output_filename="output_audio.mp3"):
    """
    Converts a string of text into an audio file using pyttsx3 and pydub.
    """
    if not text:
        print("No text to convert to audio.")
        return

    try:
        engine = pyttsx3.init()
        set_voice_to_uk_male(engine)
        
        # Adjusting the rate and volume for better quality
        engine.setProperty('rate', 170) # You can adjust this value
        engine.setProperty('volume', 1.0) # Maximum volume
        
        # pyttsx3 saves to a temporary WAV file, which is then converted to MP3
        temp_wav_file = "temp_output.wav"
        engine.save_to_file(text, temp_wav_file)
        engine.runAndWait()
        
        # Convert the temporary WAV to MP3 using pydub
        audio = AudioSegment.from_wav(temp_wav_file)
        audio.export(output_filename, format="mp3")
        
        # Clean up the temporary WAV file
        os.remove(temp_wav_file)
        
        print(f"Successfully created audio file: {output_filename}")
    except Exception as e:
        print(f"An error occurred during audio conversion: {e}")

def main():
    """
    Main function to ask the user for a file and convert it to audio.
    """
    # Print available voices for user reference
    engine_temp = pyttsx3.init()
    voices = engine_temp.getProperty('voices')
    print("--- Available System Voices ---")
    for i, voice in enumerate(voices):
        print(f"{i}: Name: {voice.name}, ID: {voice.id}")
    print("-----------------------------\n")

    file_path = input("Please enter the full path to a TXT, DOCX, or PDF file: ")

    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        return

    print(f"Extracting text from: {os.path.basename(file_path)}...")
    extracted_text = extract_text_from_file(file_path)

    if extracted_text:
        print("Text extracted successfully. Converting to audio...")
        output_name = os.path.splitext(os.path.basename(file_path))[0] + "_pyttsx3.mp3"
        convert_text_to_audio(extracted_text, output_name)

if __name__ == "__main__":
    main()
