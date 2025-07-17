import PyPDF2
import pdf2image
import pytesseract
import os
import tkinter as tk
from tkinter import filedialog
import pyttsx3
import pygame
import time
import wave
import struct
try:
    from colorama import init, Fore, Style
    from tqdm import tqdm
    colorama_available = True
except ImportError:
    colorama_available = False
    print("Colorama or tqdm not installed. Run: pip install colorama tqdm for enhanced progress indicators.")

# Initialize colorama for Windows compatibility
if colorama_available:
    init(autoreset=True)

# Set paths for Tesseract, Poppler, and output directory
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\Program Files\poppler\bin"
AUDIO_BOOKS_DIR = r"C:\Users\sgins\Downloads\AI-Generated-Audio-Books"

# Configure Tesseract path
if os.path.exists(TESSERACT_PATH):
    if colorama_available:
        print(f"{Fore.GREEN}✓ Tesseract found at {TESSERACT_PATH}{Style.RESET_ALL}")
    else:
        print(f"Tesseract found at {TESSERACT_PATH}")
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
else:
    if colorama_available:
        print(f"{Fore.RED}✗ Tesseract not found at {TESSERACT_PATH}. Ensure it is installed.{Style.RESET_ALL}")
    else:
        print(f"Tesseract not found at {TESSERACT_PATH}. Ensure it is installed.")

def select_pdf_file():
    """Select a PDF file using a dialog."""
    if colorama_available:
        print(f"{Fore.YELLOW}🚀 Starting PDF file selection...{Style.RESET_ALL}")
    else:
        print("Starting PDF file selection...")
    
    try:
        if colorama_available:
            print(f"{Fore.CYAN}📂 Opening file dialog...{Style.RESET_ALL}")
        else:
            print("Opening file dialog...")
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(
            title="Select a PDF file",
            filetypes=[("PDF files", "*.pdf")]
        )
        root.destroy()
        if file_path:
            if colorama_available:
                print(f"{Fore.GREEN}✓ Selected PDF: {file_path}{Style.RESET_ALL}")
            else:
                print(f"Selected PDF: {file_path}")
            return file_path
        else:
            if colorama_available:
                print(f"{Fore.RED}✗ No file selected via dialog.{Style.RESET_ALL}")
            else:
                print("No file selected via dialog.")
            return None
    except Exception as e:
        if colorama_available:
            print(f"{Fore.RED}✗ Error with file dialog: {e}{Style.RESET_ALL}")
        else:
            print(f"Error with file dialog: {e}")
        return None

def extract_text_with_ocr(pdf_path, max_pages=999):
    """Extract text from PDF using PyPDF2, fall back to OCR if needed."""
    if colorama_available:
        print(f"{Fore.YELLOW}📄 Starting text extraction from {pdf_path}...{Style.RESET_ALL}")
    else:
        print(f"Starting text extraction from {pdf_path}...")
    
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            num_pages = min(len(reader.pages), max_pages)
            text = ""
            if colorama_available:
                print(f"{Fore.CYAN}🔍 Extracting text from {num_pages} pages...{Style.RESET_ALL}")
            else:
                print(f"Extracting text from {num_pages} pages...")
            
            page_range = tqdm(range(num_pages), desc="Processing pages", leave=True) if colorama_available else range(num_pages)
            for page_num in page_range:
                page = reader.pages[page_num]
                page_text = page.extract_text() or ""
                text += page_text
            
            if len(text.strip()) > 100:
                if colorama_available:
                    print(f"{Fore.GREEN}✓ Text extracted successfully without OCR (length: {len(text)} chars){Style.RESET_ALL}")
                else:
                    print(f"Text extracted successfully without OCR (length: {len(text)} chars)")
                return text
            
            if colorama_available:
                print(f"{Fore.YELLOW}⚠ Minimal text detected, attempting OCR...{Style.RESET_ALL}")
            else:
                print("Minimal text detected, attempting OCR...")
            try:
                images = convert_from_path(
                    pdf_path,
                    first_page=1,
                    last_page=num_pages,
                    poppler_path=POPPLER_PATH if os.path.exists(POPPLER_PATH) else None
                )
                ocr_text = ""
                if colorama_available:
                    print(f"{Fore.CYAN}🖼 Processing {len(images)} pages with OCR...{Style.RESET_ALL}")
                else:
                    print(f"Processing {len(images)} pages with OCR...")
                
                for i, image in enumerate(tqdm(images, desc="OCR Progress", leave=True) if colorama_available else images):
                    if colorama_available:
                        print(f"{Fore.CYAN}🔎 OCR processing page {i+1}...{Style.RESET_ALL}")
                    else:
                        print(f"Processing page {i+1} with OCR...")
                    ocr_text += pytesseract.image_to_string(image, lang='eng')
                if ocr_text.strip():
                    if colorama_available:
                        print(f"{Fore.GREEN}✓ OCR completed successfully (length: {len(ocr_text)} chars){Style.RESET_ALL}")
                    else:
                        print(f"OCR completed successfully (length: {len(ocr_text)} chars)")
                    return ocr_text.strip()
                else:
                    if colorama_available:
                        print(f"{Fore.RED}✗ OCR extracted no text.{Style.RESET_ALL}")
                    else:
                        print("OCR extracted no text.")
                    return None
            except Exception as e:
                if colorama_available:
                    print(f"{Fore.RED}✗ OCR failed: {e}{Style.RESET_ALL}")
                else:
                    print(f"OCR failed: {e}")
                return None
    
    except Exception as e:
        if colorama_available:
            print(f"{Fore.RED}✗ Error processing PDF: {e}{Style.RESET_ALL}")
        else:
            print(f"Error processing PDF: {e}")
        return None

def count_words(text):
    """Count non-empty words in text."""
    words = [word for word in text.split() if word.strip()]
    return len(words)

def add_silence_to_wav(input_wav, output_wav, silence_duration=2):
    """Add silence to the start of a WAV file."""
    try:
        with wave.open(input_wav, 'rb') as wav_in:
            channels = wav_in.getnchannels()
            sample_width = wav_in.getsampwidth()
            framerate = wav_in.getframerate()
            n_frames = wav_in.getnframes()
            audio_data = wav_in.readframes(n_frames)
            
            silence_frames = int(framerate * silence_duration)
            silence_data = struct.pack('<' + str(channels * silence_frames) + 'h', *([0] * channels * silence_frames))
            
            with wave.open(output_wav, 'wb') as wav_out:
                wav_out.setnchannels(channels)
                wav_out.setsampwidth(sample_width)
                wav_out.setframerate(framerate)
                wav_out.writeframes(silence_data)
                wav_out.writeframes(audio_data)
                
        return True
    except Exception as e:
        if colorama_available:
            print(f"{Fore.RED}✗ Error adding silence to WAV: {e}{Style.RESET_ALL}")
        else:
            print(f"Error adding silence to WAV: {e}")
        return False

def text_to_speech(text, temp_file, output_file):
    """Convert text to WAV using pyttsx3 with a male voice."""
    if colorama_available:
        print(f"{Fore.YELLOW}🎙 Starting text-to-speech conversion...{Style.RESET_ALL}")
    else:
        print("Starting text-to-speech conversion...")
    try:
        if os.path.exists(temp_file):
            if colorama_available:
                print(f"{Fore.YELLOW}🗑 Deleting existing temp WAV file: {temp_file}{Style.RESET_ALL}")
            else:
                print(f"Deleting existing temp WAV file: {temp_file}")
            os.remove(temp_file)
        
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        selected_voice = None
        for voice in voices:
            if "male" in voice.name.lower() or getattr(voice, 'gender', '').lower() == "male":
                selected_voice = voice
                engine.setProperty('voice', voice.id)
                break
        else:
            if colorama_available:
                print(f"{Fore.YELLOW}⚠ No male voice found, using default voice.{Style.RESET_ALL}")
            else:
                print("No male voice found, using default voice.")
            selected_voice = voices[0] if voices else None
            if selected_voice:
                engine.setProperty('voice', selected_voice.id)
        
        # Display voice name and gender
        if selected_voice:
            voice_name = selected_voice.name
            voice_gender = getattr(selected_voice, 'gender', 'Unknown')
            if colorama_available:
                print(f"{Fore.CYAN}🔊 Selected voice: {voice_name} (Gender: {voice_gender}){Style.RESET_ALL}")
            else:
                print(f"Selected voice: {voice_name} (Gender: {voice_gender})")
        else:
            if colorama_available:
                print(f"{Fore.RED}✗ No voices available.{Style.RESET_ALL}")
            else:
                print("No voices available.")
            return False
        
        engine.setProperty('rate', 150)
        engine.setProperty('volume', 0.9)
        
        engine.save_to_file(text, temp_file)
        engine.runAndWait()
        
        if os.path.exists(temp_file):
            if add_silence_to_wav(temp_file, output_file):
                if colorama_available:
                    print(f"{Fore.GREEN}✓ WAV file with silence saved as: {output_file}{Style.RESET_ALL}")
                else:
                    print(f"WAV file with silence saved as: {output_file}")
                os.remove(temp_file)
                return True
            else:
                if colorama_available:
                    print(f"{Fore.RED}✗ Failed to add silence to WAV file: {output_file}{Style.RESET_ALL}")
                else:
                    print(f"Failed to add silence to WAV file: {output_file}")
                return False
        else:
            if colorama_available:
                print(f"{Fore.RED}✗ Failed to create temporary WAV file: {temp_file}{Style.RESET_ALL}")
            else:
                print(f"Failed to create temporary WAV file: {temp_file}")
            return False
    except Exception as e:
        if colorama_available:
            print(f"{Fore.RED}✗ Error generating WAV: {e}{Style.RESET_ALL}")
        else:
            print(f"Error generating WAV: {e}")
        return False

def play_audio(file_path):
    """Play WAV file using pygame."""
    try:
        pygame.mixer.init()
        pygame.mixer.music.load(file_path)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(10)
    except Exception as e:
        if colorama_available:
            print(f"{Fore.RED}✗ Error playing audio: {e}{Style.RESET_ALL}")
        else:
            print(f"Error playing audio: {e}")

def main():
    """Main function to run the script."""
    if colorama_available:
        print(f"{Fore.MAGENTA}🌟 Starting PDF to WAV conversion process...{Style.RESET_ALL}")
    else:
        print("Starting PDF to WAV conversion process...")
    
    pdf_path = select_pdf_file()
    if not pdf_path or not os.path.exists(pdf_path):
        if colorama_available:
            print(f"{Fore.RED}✗ Invalid or no file selected. Exiting.{Style.RESET_ALL}")
        else:
            print("Invalid or no file selected. Exiting.")
        return
    
    text = extract_text_with_ocr(pdf_path)
    if text:
        word_count = count_words(text)
        if colorama_available:
            print(f"{Fore.GREEN}Number of words in the text: {word_count}{Style.RESET_ALL}")
        else:
            print(f"Number of words in the text: {word_count}")
        
        temp_file = os.path.splitext(pdf_path)[0] + "_temp.wav"
        output_base = os.path.splitext(os.path.basename(pdf_path))[0] + "_output.wav"
        output_file = os.path.join(AUDIO_BOOKS_DIR, output_base)
        
        if text_to_speech(text, temp_file, output_file):
            time.sleep(2)
            play_audio(output_file)
            if colorama_available:
                print(f"{Fore.GREEN}✅ Conversion and playback completed successfully!{Style.RESET_ALL}")
            else:
                print("Conversion and playback completed successfully!")
        else:
            if colorama_available:
                print(f"{Fore.RED}✗ Conversion failed during WAV generation. Exiting.{Style.RESET_ALL}")
            else:
                print("Conversion failed during WAV generation. Exiting.")
    else:
        if colorama_available:
            print(f"{Fore.RED}✗ No text extracted from PDF. Exiting.{Style.RESET_ALL}")
        else:
            print("No text extracted from PDF. Exiting.")

if __name__ == "__main__":
    main()