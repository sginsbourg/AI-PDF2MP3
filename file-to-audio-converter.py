import os
import sys
import pyttsx3
import PyPDF2
from docx import Document
try:
    from pydub import AudioSegment
except ImportError as e:
    print("Error: Failed to import pydub. This may be due to missing dependencies like 'audioop-lts' for Python 3.13+.")
    print("Ensure all required libraries are installed via pip, and FFmpeg is available.")
    print("Original error:", e)
    sys.exit(1)

# Set custom path for FFmpeg
ffmpeg_bin = r"C:\ffmpeg\bin"
ffmpeg_exe = os.path.join(ffmpeg_bin, "ffmpeg.exe")
ffprobe_exe = os.path.join(ffmpeg_bin, "ffprobe.exe")

if os.path.exists(ffmpeg_exe) and os.path.exists(ffprobe_exe):
    AudioSegment.converter = ffmpeg_exe
    AudioSegment.ffprobe = ffprobe_exe
    print(f"Using FFmpeg from {ffmpeg_bin}")
else:
    print(f"FFmpeg executables not found in {ffmpeg_bin}.")
    print("Please install FFmpeg at this location or adjust the path in the script.")
    sys.exit(1)

# Install these libraries using pip before running the script:
# pip install python-docx PyPDF2 pyttsx3 pydub audioop-lts
# audioop-lts is required for Python 3.13+ as a replacement for the deprecated audioop module.
# To ensure the script works on Windows 11 64-bit, you need to install FFmpeg.
# FFmpeg is a crucial tool that pydub uses for converting audio formats.
# Here are the steps to install it on Windows:
# 1. Go to the FFmpeg download page: https://ffmpeg.org/download.html
# 2. Click the Windows icon and select a build (e.g., from "git" or "release").
# 3. Download the .zip file for your architecture (e.g., "win64-gpl").
# 4. Extract the contents to a folder, e.g., C:\ffmpeg.
# 5. This script now uses the fixed path C:\ffmpeg\bin for FFmpeg executables.
# 6. Verify by ensuring ffmpeg.exe and ffprobe.exe are in C:\ffmpeg\bin.

def extract_text_from_file(file_path):
    """
    Extracts text from a TXT, DOCX, or PDF file as a list of (text, type) tuples.
    Type is 'heading' or 'paragraph'.
    """
    file_extension = os.path.splitext(file_path)[1].lower()
    text_segments = []
    try:
        if file_extension == ".txt":
            with open(file_path, 'r', encoding='utf-8') as f:
                text = f.read()
                # Split by double newlines
                segments = [s.strip() for s in text.split('\n\n') if s.strip()]
                for segment in segments:
                    # Heuristic: short segments (<50 chars) are headings
                    segment_type = "heading" if len(segment) < 50 else "paragraph"
                    text_segments.append((segment, segment_type))
        elif file_extension in [".docx", ".doc"]:
            document = Document(file_path)
            for paragraph in document.paragraphs:
                if paragraph.text.strip():
                    # Check if paragraph style indicates a heading
                    style_name = paragraph.style.name.lower()
                    segment_type = "heading" if "heading" in style_name else "paragraph"
                    text_segments.append((paragraph.text.strip(), segment_type))
        elif file_extension == ".pdf":
            with open(file_path, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        # Split page text by double newlines
                        page_segments = [s.strip() for s in page_text.split('\n\n') if s.strip()]
                        for segment in page_segments:
                            # Heuristic: short segments (<50 chars) are headings
                            segment_type = "heading" if len(segment) < 50 else "paragraph"
                            text_segments.append((segment, segment_type))
        else:
            print(f"Error: Unsupported file type '{file_extension}'.")
            return None
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
        return None
    except Exception as e:
        print(f"An unexpected error occurred while reading the file: {e}")
        return None

    return text_segments if text_segments else None

def save_text_to_file(text_segments, temp_file="temp.txt"):
    """
    Saves the extracted text to a single-line text file.
    """
    try:
        # Combine only the text parts of the segments
        unified_text = " ".join(segment[0] for segment in text_segments).replace("\n", " ")
        with open(temp_file, 'w', encoding='utf-8') as f:
            f.write(unified_text)
        print(f"Successfully saved text to {temp_file}")
        return temp_file
    except Exception as e:
        print(f"Error saving text to {temp_file}: {e}")
        return None

def create_lowercase_text_file(input_file, output_file="small_letter_temp.txt"):
    """
    Reads text from input_file, converts it to lowercase, and saves it to output_file.
    """
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            text = f.read().lower()  # Convert to lowercase
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(text)
        print(f"Successfully created lowercase text file: {output_file}")
        return output_file
    except Exception as e:
        print(f"Error creating lowercase text file {output_file}: {e}")
        return None

def read_text_from_file(temp_file):
    """
    Reads text from the temporary text file.
    """
    try:
        with open(temp_file, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"Error reading from {temp_file}: {e}")
        return None

def preprocess_text_for_tts(text):
    """
    Preprocesses text to fix pronunciation issues for the TTS engine.
    Replaces problematic words with phonetic equivalents for better pronunciation.
    """
    # Since text is already lowercase from small_letter_temp.txt, match lowercase
    text = text.replace("testing", "test-ing")
    text = text.replace("we", "wee")
    text = text.replace("data", "dah-ta")
    text = text.replace(" ai ", " eye ")  # Add spaces to avoid partial matches
    text = text.replace("read", "red")
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

def convert_text_to_audio(text_segments, output_filename="output_audio.mp3"):
    """
    Converts a list of (text, type) tuples into an audio file with different pauses.
    Uses 1.5-second pause after headings, 1-second pause after paragraphs.
    """
    if not text_segments:
        print("No text to convert to audio.")
        return

    try:
        engine = pyttsx3.init()
        set_voice_to_uk_male(engine)
        
        # Adjusting the rate and volume for better quality
        engine.setProperty('rate', 170)  # You can adjust this value
        engine.setProperty('volume', 1.0)  # Maximum volume
        
        temp_wav_files = []
        heading_pause_ms = 1500  # 1.5-second pause after headings
        paragraph_pause_ms = 1000  # 1-second pause after paragraphs
        heading_pause_audio = AudioSegment.silent(duration=heading_pause_ms)
        paragraph_pause_audio = AudioSegment.silent(duration=paragraph_pause_ms)

        # Convert each segment to a WAV file with preprocessing
        for i, (text, segment_type) in enumerate(text_segments):
            if text.strip():  # Skip empty segments
                # Preprocess text to fix pronunciation
                processed_text = preprocess_text_for_tts(text)
                temp_wav_file = f"temp_segment_{i}.wav"
                engine.save_to_file(processed_text, temp_wav_file)
                engine.runAndWait()
                temp_wav_files.append((temp_wav_file, segment_type))
        
        # Combine WAV files with appropriate pauses
        combined_audio = AudioSegment.empty()
        for i, (wav_file, segment_type) in enumerate(temp_wav_files):
            audio = AudioSegment.from_wav(wav_file)
            combined_audio += audio
            if i < len(temp_wav_files) - 1:  # Add pause after all but the last segment
                pause_audio = heading_pause_audio if segment_type == "heading" else paragraph_pause_audio
                combined_audio += pause_audio
        
        # Export combined audio to MP3
        combined_audio.export(output_filename, format="mp3")
        
        # Clean up temporary WAV files
        for wav_file, _ in temp_wav_files:
            try:
                os.remove(wav_file)
            except Exception as e:
                print(f"Error cleaning up {wav_file}: {e}")
        
        print(f"Successfully created audio file: {output_filename}")
    except Exception as e:
        print(f"An error occurred during audio conversion: {e}")
        print("Ensure FFmpeg is installed at C:\\ffmpeg\\bin and accessible.")
        sys.exit(1)

def main():
    """
    Main function to ask the user for a file, save text to temp.txt, create lowercase small_letter_temp.txt, and convert it to audio.
    """
    # Print available voices for user reference
    engine_temp = pyttsx3.init()
    voices = engine_temp.getProperty('voices')
    print("--- Available System Voices ---")
    for i, voice in enumerate(voices):
        print(f"{i}: Name: {voice.name}, ID: {voice.id}")
    print("-----------------------------\n")

    file_path = input("Please enter the full path to a TXT, DOCX, or PDF file: ").strip().strip('"').strip("'")

    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        return

    print(f"Extracting text from: {os.path.basename(file_path)}...")
    text_segments = extract_text_from_file(file_path)

    if text_segments:
        temp_file = "temp.txt"
        print(f"Saving extracted text to {temp_file}...")
        saved_file = save_text_to_file(text_segments, temp_file)
        if saved_file:
            print(f"Creating lowercase text file from {temp_file}...")
            lowercase_file = create_lowercase_text_file(temp_file, "small_letter_temp.txt")
            if lowercase_file:
                print("Lowercase text file created successfully. Converting to audio...")
                text_to_convert = read_text_from_file(lowercase_file)
                if text_to_convert:
                    # Split lowercase text into segments, using original segment types
                    lowercase_segments = []
                    lowercase_text_parts = text_to_convert.split('\n\n')
                    lowercase_text_parts = [p.strip() for p in lowercase_text_parts if p.strip()]
                    # Match lowercase parts to original segments, preserving types
                    for i, text in enumerate(lowercase_text_parts):
                        segment_type = text_segments[i][1] if i < len(text_segments) else "paragraph"
                        lowercase_segments.append((text, segment_type))
                    output_name = os.path.splitext(os.path.basename(file_path))[0] + "_pyttsx3.mp3"
                    convert_text_to_audio(lowercase_segments, output_name)
                    # Clean up temporary files
                    try:
                        os.remove(temp_file)
                        print(f"Cleaned up temporary file: {temp_file}")
                        os.remove(lowercase_file)
                        print(f"Cleaned up temporary file: {lowercase_file}")
                    except Exception as e:
                        print(f"Error cleaning up temporary files: {e}")
                else:
                    print("Failed to read text from lowercase file. Audio conversion aborted.")
            else:
                print("Failed to create lowercase text file. Audio conversion aborted.")
            # Clean up temp.txt if lowercase file creation failed
            if os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                    print(f"Cleaned up temporary file: {temp_file}")
                except Exception as e:
                    print(f"Error cleaning up {temp_file}: {e}")
        else:
            print("Failed to save text to temporary file. Audio conversion aborted.")

if __name__ == "__main__":
    main()