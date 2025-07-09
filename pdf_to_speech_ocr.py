import PyPDF2
try:
    from gTTS import gTTS
except ImportError:
    from gtts import gTTS
from pdf2image import convert_from_path
import pytesseract
import os
try:
    import tkinter as tk
    from tkinter import filedialog
    tkinter_available = True
except ImportError:
    tkinter_available = False
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

# Set paths for Tesseract and Poppler
TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\Program Files\poppler\bin"

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
    """Select a PDF file using a dialog (if tk available) or console input."""
    if colorama_available:
        print(f"{Fore.YELLOW}🚀 Starting PDF file selection...{Style.RESET_ALL}")
    else:
        print("Starting PDF file selection...")

    if tkinter_available:
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
                print("Falling back to manual input.")
        except Exception as e:
            if colorama_available:
                print(f"{Fore.RED}✗ Error with file dialog: {e}{Style.RESET_ALL}")
            else:
                print(f"Error with file dialog: {e}")
            print("Falling back to manual input.")
    
    file_path = input("Enter the path to the PDF file: ")
    if file_path and os.path.exists(file_path):
        if colorama_available:
            print(f"{Fore.GREEN}✓ Valid PDF path entered: {file_path}{Style.RESET_ALL}")
        else:
            print(f"Valid PDF path entered: {file_path}")
        return file_path
    else:
        if colorama_available:
            print(f"{Fore.RED}✗ Invalid or non-existent file path: {file_path}{Style.RESET_ALL}")
        else:
            print(f"Invalid or non-existent file path: {file_path}")
        return None

def extract_text_with_ocr(pdf_path, max_pages=999):
    """Extract text from PDF using PyPDF2, fall back to OCR if needed."""
    if colorama_available:
        print(f"{Fore.YELLOW}📄 Starting text extraction from {pdf_path}...{Style.RESET_ALL}")
    else:
        print(f"Starting text extraction from {pdf_path}...")
    
    try:
        # First, try direct text extraction with PyPDF2
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            num_pages = min(len(reader.pages), max_pages)
            text = ""
            if colorama_available:
                print(f"{Fore.CYAN}🔍 Extracting text from {num_pages} pages...{Style.RESET_ALL}")
            else:
                print(f"Extracting text from {num_pages} pages...")
            
            # Use tqdm for progress bar if available
            page_range = tqdm(range(num_pages), desc="Processing pages", leave=True) if colorama_available else range(num_pages)
            for page_num in page_range:
                page = reader.pages[page_num]
                page_text = page.extract_text() or ""
                text += page_text
            
            # If text is substantial, return it
            if len(text.strip()) > 100:
                if colorama_available:
                    print(f"{Fore.GREEN}✓ Text extracted successfully without OCR (length: {len(text)} chars){Style.RESET_ALL}")
                else:
                    print(f"Text extracted successfully without OCR (length: {len(text)} chars)")
                return text
            
            # If little/no text, use OCR
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
                
                # Use tqdm for OCR progress bar
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

def text_to_speech(text, output_file="output.mp3", lang='en', tld='co.uk'):
    """Convert text to UK English MP3 using pyttsx3 with a male voice."""
    if colorama_available:
        print(f"{Fore.YELLOW}🎙 Starting text-to-speech conversion...{Style.RESET_ALL}")
    else:
        print("Starting text-to-speech conversion...")
    try:
        if os.path.exists(output_file):
            if colorama_available:
                print(f"{Fore.YELLOW}🗑 Deleting existing MP3 file: {output_file}{Style.RESET_ALL}")
            else:
                print(f"Deleting existing MP3 file: {output_file}")
            os.remove(output_file)
        
        import pyttsx3
        if colorama_available:
            print(f"{Fore.CYAN}🔊 Generating male MP3 for {output_file}...{Style.RESET_ALL}")
        else:
            print(f"Generating male MP3 for {output_file}...")
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        # Select a male voice (e.g., Microsoft David)
        for voice in voices:
            if "David" in voice.name or "male" in voice.name.lower():
                engine.setProperty('voice', voice.id)
                break
        engine.setProperty('rate', 150)  # Adjust speed for male-like tone
        engine.save_to_file(text, output_file)
        engine.runAndWait()
        
        if colorama_available:
            print(f"{Fore.GREEN}✓ MP3 file saved as {output_file}{Style.RESET_ALL}")
        else:
            print(f"MP3 file saved as {output_file}")
        return True
    except Exception as e:
        if colorama_available:
            print(f"{Fore.RED}✗ Error generating MP3: {e}{Style.RESET_ALL}")
        else:
            print(f"Error generating MP3: {e}")
        return False

def main():
    """Main function to run the script."""
    if colorama_available:
        print(f"{Fore.MAGENTA}🌟 Starting PDF to MP3 conversion process...{Style.RESET_ALL}")
    else:
        print("Starting PDF to MP3 conversion process...")
    
    pdf_path = select_pdf_file()
    if not pdf_path or not os.path.exists(pdf_path):
        if colorama_available:
            print(f"{Fore.RED}✗ Invalid or no file selected. Exiting.{Style.RESET_ALL}")
        else:
            print("Invalid or no file selected. Exiting.")
        return
    
    text = extract_text_with_ocr(pdf_path)
    if text:
        output_file = os.path.splitext(pdf_path)[0] + "_speech.mp3"
        if text_to_speech(text, output_file):
            if colorama_available:
                print(f"{Fore.GREEN}✅ Conversion process completed successfully!{Style.RESET_ALL}")
            else:
                print("Conversion process completed successfully!")
        else:
            if colorama_available:
                print(f"{Fore.RED}✗ Conversion failed during MP3 generation. Exiting.{Style.RESET_ALL}")
            else:
                print("Conversion failed during MP3 generation. Exiting.")
    else:
        if colorama_available:
            print(f"{Fore.RED}✗ No text extracted from PDF. Exiting.{Style.RESET_ALL}")
        else:
            print("No text extracted from PDF. Exiting.")

if __name__ == "__main__":
    main()