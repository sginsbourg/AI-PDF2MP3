import PyPDF2
from pdf2image import convert_from_path
import pytesseract
import pyttsx3
import os
from tqdm import tqdm
import sys
import threading
import time
try:
    import tkinter as tk
    from tkinter import filedialog
    tkinter_available = True
except ImportError:
    tkinter_available = False

def select_pdf_file():
    """Select a PDF file using a dialog (if tk available) or console input."""
    if tkinter_available:
        try:
            root = tk.Tk()
            root.withdraw()  # Hide the main window
            file_path = filedialog.askopenfilename(
                title="Select a PDF file",
                filetypes=[("PDF files", "*.pdf")]
            )
            root.destroy()
            return file_path
        except Exception as e:
            print(f"Error with file dialog: {e}")
            print("Falling back to manual input.")
    
    # Fallback to console input
    file_path = input("Enter the path to the PDF file: ")
    return file_path.strip()

def extract_text_with_ocr(pdf_path, max_pages=999):
    """Extract text from PDF using PyPDF2, fall back to OCR if needed."""
    try:
        # First, try direct text extraction with PyPDF2
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            num_pages = min(len(reader.pages), max_pages)
            text = ""
            print("Extracting text directly...")
            for page_num in tqdm(range(num_pages), desc="Processing pages", unit="page"):
                page = reader.pages[page_num]
                page_text = page.extract_text() or ""
                text += page_text
            
            # If text is substantial, return it
            if len(text.strip()) > 100:  # Adjusted threshold for robustness
                print("Text extracted successfully without OCR.")
                return text
            
            # If little/no text, use OCR
            print("Minimal text detected, attempting OCR...")
            images = convert_from_path(pdf_path, first_page=1, last_page=num_pages)
            ocr_text = ""
            for i, image in enumerate(tqdm(images, desc="OCR processing", unit="page")):
                ocr_text += pytesseract.image_to_string(image, lang='eng')
            return ocr_text.strip() if ocr_text else None
    
    except Exception as e:
        print(f"Error processing PDF: {e}")
        return None

def spinner(message="Processing"):
    """Display a spinning animation in the console."""
    spinner_chars = ['|', '/', '-', '\\']
    idx = 0
    while not getattr(spinner, "done", False):
        sys.stdout.write(f'\r{message} {spinner_chars[idx % len(spinner_chars)]}')
        sys.stdout.flush()
        idx += 1
        time.sleep(0.1)
    sys.stdout.write('\r' + ' ' * (len(message) + 2) + '\r')
    sys.stdout.flush()

def text_to_speech(text, output_file="output.mp3"):
    """Convert text to MP3 using a bright and clear UK male voice with pyttsx3."""
    try:
        # Initialize pyttsx3 engine
        engine = pyttsx3.init()
        
        # Find UK male voice (e.g., Microsoft George)
        uk_male_voice = None
        for voice in engine.getProperty('voices'):
            if "UK" in voice.name and "male" in voice.name.lower():
                uk_male_voice = voice.id
                break
            elif "George" in voice.name:  # Fallback for Microsoft George
                uk_male_voice = voice.id
                break
        
        if uk_male_voice:
            engine.setProperty('voice', uk_male_voice)
        else:
            print("Warning: UK male voice not found, using default voice.")
        
        # Set properties for bright and clear tone
        engine.setProperty('rate', 180)  # Faster rate for brightness
        engine.setProperty('pitch', 1.2)  # Slightly higher pitch (if supported)
        engine.setProperty('volume', 0.9)  # High volume for clarity
        
        # Start spinner
        spinner.done = False
        spinner_thread = threading.Thread(target=spinner, args=("Generating MP3",))
        spinner_thread.start()
        
        # Save to MP3
        engine.save_to_file(text, output_file)
        engine.runAndWait()
        
        # Stop spinner
        spinner.done = True
        spinner_thread.join()
        
        print(f"MP3 file saved as {output_file}")
        
        # Autoplay the MP3
        try:
            os.startfile(output_file)
            print("Playing MP3...")
        except Exception as e:
            print(f"Error playing MP3: {e}")
            sys.exit(1)
        
    except Exception as e:
        # Stop spinner on error
        spinner.done = True
        spinner_thread.join()
        print(f"Error generating MP3: {e}")
        sys.exit(1)

def main():
    """Main function to run the script for a single PDF."""
    pdf_path = select_pdf_file()
    if not pdf_path or not os.path.exists(pdf_path):
        print("Invalid or no file selected. Exiting.")
        sys.exit(1)
    
    # Create dedicated folder for MP3 in user's Downloads
    output_dir = os.path.join(os.path.expanduser("~"), "Downloads", "AI-Generated-Audio-Books")
    os.makedirs(output_dir, exist_ok=True)
    
    text = extract_text_with_ocr(pdf_path)
    if text:
        # Set MP3 path to user's Downloads/AI-Generated-Audio-Books
        output_file = os.path.join(output_dir, os.path.basename(os.path.splitext(pdf_path)[0] + "_speech.mp3"))
        text_to_speech(text, output_file)
        
        # Launch the original PDF
        try:
            os.startfile(pdf_path)
            print("Opening PDF and exiting...")
            sys.exit(0)
        except Exception as e:
            print(f"Error opening PDF: {e}")
            sys.exit(1)
    else:
        print("No text extracted from PDF. Exiting.")
        sys.exit(1)

if __name__ == "__main__":
    main()