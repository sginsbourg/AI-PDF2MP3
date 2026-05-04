import os
import re
import sys
import json
import shutil
import argparse
import tempfile
import subprocess
import tempfile
import subprocess
import tkinter as tk
from tkinter import filedialog
import fitz  # PyMuPDF
import asyncio
import edge_tts

# --- 7-Step Procedural Structure ---

def step1_get_pdf_path(path_arg=None):
    """
    1. get pdf from parameter in command line of open gui to help user select a pdf file.
    """
    if path_arg and os.path.exists(path_arg):
        return path_arg
    
    # Open GUI if no path provided or path is invalid
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    pdf_path = filedialog.askopenfilename(
        title="Select PDF for Audiobook Conversion",
        filetypes=[("PDF files", "*.pdf")],
    )
    root.destroy()
    return pdf_path

def step2_validate_and_read_pdf(pdf_path):
    """
    2. validatate that the pdf exists and read it.
    If no text is found, attempt OCR.
    """
    if not pdf_path or not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")
    
    print(f"[*] Validating and reading: {pdf_path}")
    doc = fitz.open(pdf_path)
    text_content = []
    
    for page in doc:
        text = page.get_text().strip()
        if text:
            text_content.append(text)
    
    if not text_content:
        print("[!] No text layer found. Checking for images to perform OCR...")
        has_images = any(len(page.get_images()) > 0 for page in doc)
        
        if has_images:
            print("[*] Scanned PDF detected. Initiating OCR (Tesseract required)...")
            try:
                import pytesseract
                from PIL import Image
                import io
                
                import shutil

                # Robust Tesseract path discovery
                tess_paths = [
                    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                    os.path.join(os.environ.get("USERPROFILE", ""), r"AppData\Local\Tesseract-OCR\tesseract.exe"),
                    os.path.join(os.environ.get("USERPROFILE", ""), r"AppData\Local\Programs\Tesseract-OCR\tesseract.exe")
                ]
                
                def find_tesseract():
                    tess_in_path = shutil.which("tesseract")
                    if tess_in_path:
                        return tess_in_path
                    for p in tess_paths:
                        if os.path.exists(p):
                            return p
                    return None

                tesseract_cmd_path = find_tesseract()

                if not tesseract_cmd_path:
                    print("[*] Tesseract OCR not found. Attempting to install automatically via winget...")
                    try:
                        subprocess.run(
                            ["winget", "install", "UB-Mannheim.TesseractOCR", "--accept-package-agreements", "--accept-source-agreements", "--silent"],
                            check=True
                        )
                        print("[+] Tesseract OCR installed successfully. Re-checking paths...")
                        tesseract_cmd_path = find_tesseract()
                    except Exception as e:
                        print(f"[!] Auto-installation failed: {e}")

                if tesseract_cmd_path:
                    pytesseract.pytesseract.tesseract_cmd = tesseract_cmd_path
                else:
                    raise RuntimeError("Tesseract OCR engine not found! Please download and install it from https://github.com/UB-Mannheim/tesseract/wiki (or run 'winget install UB-Mannheim.TesseractOCR').")
                
                for i, page in enumerate(doc):
                    print(f"  [OCR] Processing page {i+1}/{len(doc)}...")
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # Higher DPI for better OCR
                    img_data = pix.tobytes("png")
                    img = Image.open(io.BytesIO(img_data))
                    page_text = pytesseract.image_to_string(img)
                    if page_text.strip():
                        text_content.append(page_text)
                
                if not text_content:
                    raise ValueError("OCR yielded no text content.")
                
                print("[+] OCR successful.")
            except ImportError:
                raise ImportError("pytesseract or Pillow not installed. Please run 'pip install pytesseract Pillow'.")
            except Exception as e:
                if "tesseract is not installed" in str(e).lower() or "not found" in str(e).lower():
                    raise RuntimeError("Tesseract OCR engine not found! Please download and install it from https://github.com/UB-Mannheim/tesseract/wiki (or run 'winget install UB-Mannheim.TesseractOCR').")
                raise e
        else:
            raise ValueError("The PDF contains no extractable text and no images for OCR.")
    
    return text_content

def step3_generate_structured_json(text_pages, pdf_path):
    """
    3. generate a Json for extracting the entire text from the pdf in a structured manner.
    """
    print("[*] Generating structured JSON from text content...")
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    
    elements = []
    max_seg_len = 1500 # pyttsx3 stability
    
    def sanitize_text(text):
        # Replace non-standard punctuation that might trip up TTS
        text = text.replace('–', '-').replace('—', '-')
        text = text.replace('“', '"').replace('”', '"')
        text = text.replace('‘', "'").replace('’', "'")
        text = text.replace('…', '...')
        # Remove multiple dots
        text = re.sub(r'\.{4,}', '...', text)
        # Remove "Page X" or "PAGE X"
        text = re.sub(r'(?i)page\s+\d+', '', text)
        return text

    def clean_boilerplate_candidate(line):
        # Strip page numbers and common noise to find truly repeating headers
        l = line.strip()
        l = re.sub(r'(?i)page\s+\d+(\s+of\s+\d+)?', '', l)
        l = re.sub(r'^\d+$', '', l) # Just a naked number
        return l.strip()

    # Identify potential repeating boilerplate across pages via Global Line Frequency
    content_pages = []
    if len(text_pages) > 1:
        from collections import Counter
        all_lines_cleaned = []
        for p in text_pages:
            # We unique-ify lines per page to count "page frequency" not "total frequency"
            page_lines = {clean_boilerplate_candidate(l) for l in p.split('\n') if l.strip()}
            all_lines_cleaned.extend(list(page_lines))
        
        line_page_counts = Counter(all_lines_cleaned)
        
        # Threshold: if a cleaned line appears on > 20% of pages and there's more than 2 pages
        threshold = max(2, len(text_pages) * 0.2)
        systemic_boilerplate = {l for l, count in line_page_counts.items() if l and count >= threshold}
        
        if systemic_boilerplate:
            print(f"[*] Detected {len(systemic_boilerplate)} repeating boilerplate patterns. Scrubbing...")

        for p in text_pages:
            filtered_lines = []
            for line in p.split('\n'):
                if clean_boilerplate_candidate(line) in systemic_boilerplate:
                    continue # Skip this line
                filtered_lines.append(line)
            content_pages.append("\n".join(filtered_lines))
    else:
        content_pages = text_pages

    for i, page_text in enumerate(content_pages):
        # Clean text
        cleaned_page = page_text.replace("\n", " ")
        cleaned_page = re.sub(r"\s+", " ", cleaned_page)
        cleaned_page = sanitize_text(cleaned_page)
        
        # Split into segments by paragraphs (rough heuristic)
        segments = re.split(r"(?<=[.!?])\s+(?=[A-Z])", cleaned_page)
        
        for seg in segments:
            seg = seg.strip()
            if len(seg) < 10:
                continue
            
            # If segment is too long, split it further by sentences or just length
            while len(seg) > max_seg_len:
                # Find the last period within the limit
                split_idx = seg.rfind('. ', 0, max_seg_len)
                if split_idx == -1:
                    split_idx = max_seg_len
                
                chunk = seg[:split_idx+1].strip()
                elements.append({"type": "paragraph", "text": chunk, "page": i + 1})
                seg = seg[split_idx+1:].strip()

            if not seg:
                continue

            # Granular Heading Detection
            words = seg.split()
            seg_type = "paragraph"
            
            # 1. Chapter Titles: Short, often all-caps or starts with "Chapter" or "Part"
            if len(words) < 8 and (seg.isupper() or any(keyword in seg.lower() for keyword in ["chapter", "part", "preface", "conclusion"])):
                seg_type = "chapter_title"
            # 2. Section Headlines: Numbered or all-caps but slightly longer
            elif len(words) < 12 and (seg.isupper() or re.match(r"^\d+(\.\d+)*\s?[A-Z]", seg)):
                seg_type = "section_headline"
            # 3. Sub-headers: Short phrases
            elif len(words) < 15 and seg[0].isupper():
                seg_type = "sub_heading"
            
            elements.append({
                "type": seg_type,
                "text": seg,
                "page": i + 1
            })
    
    json_data = {
        "title": base_name,
        "source": pdf_path,
        "total_segments": len(elements),
        "elements": elements
    }
    return json_data

def step4_ai_record_audiobook(json_data, temp_dir):
    """
    4. ai "record" the audiobook using Edge-TTS (High-Quality Neural Voices)
    """
    print("[*] Recording segments using Edge-TTS Neural Voice Engine...")
    
    audio_segments = []
    total = len(json_data["elements"])
    voice_id = "en-US-ChristopherNeural" # High-quality male voice
    
    async def generate_segment(text, filename):
        communicate = edge_tts.Communicate(text, voice_id)
        await communicate.save(filename)

    for i, element in enumerate(json_data["elements"]):
        chunk_file = os.path.abspath(os.path.join(temp_dir, f"segment_{i:04d}.mp3"))
        text_to_speak = element["text"]
        
        # Basic sanitation already done in step3, but minor fixes here
        text_to_speak = text_to_speak.replace(" vs. ", " versus ")
        text_to_speak = re.sub(r'([a-z])([A-Z])', r'\1 \2', text_to_speak)
        
        pid = os.getpid()
        snippet = text_to_speak[:50] + "..." if len(text_to_speak) > 50 else text_to_speak
        print(f"  [PID:{pid}][Recording {i+1}/{total}] {element['type']} ({len(text_to_speak)} chars): \"{snippet}\"")
        
        try:
            asyncio.run(generate_segment(text_to_speak, chunk_file))
            audio_segments.append(chunk_file)
        except Exception as e:
            print(f"  [!] Error during segment {i+1}: {e}")
            
    return audio_segments

def step5_merge_with_pauses(audio_segments, json_data, output_path):
    """
    5. edit a merged audiobook with inserting the pauses.
    """
    print("[*] Merging segments and inserting industry-standard pauses...")
    
    if not audio_segments:
        return False

    with tempfile.TemporaryDirectory() as merge_temp:
        # Create silent files (Parameters must match TTS engine output: 24000Hz, Mono)
        tail_sh = os.path.join(merge_temp, "tail.mp3")     # 0.5s
        pause_sh = os.path.join(merge_temp, "short.mp3")   # 1.2s
        pause_md = os.path.join(merge_temp, "medium.mp3")  # 1.8s
        pause_lg = os.path.join(merge_temp, "long.mp3")    # 3.0s
        
        def create_silence(path, duration):
            # Match Edge-TTS output parameters 24000Hz Mono
            cmd = ["ffmpeg", "-y", "-f", "lavfi", "-i", f"anullsrc=r=24000:cl=mono", "-t", str(duration), "-acodec", "libmp3lame", "-ar", "24000", "-ac", "1", "-q:a", "2", path]
            subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)

        create_silence(tail_sh, 0.5)
        create_silence(pause_sh, 1.2)
        create_silence(pause_md, 1.8)
        create_silence(pause_lg, 3.0)
        
        # Build concat list
        concat_file = os.path.join(merge_temp, "list.txt")
        with open(concat_file, "w", encoding="utf-8") as f:
            # 1. Leading Tail
            f.write(f"file '{tail_sh.replace('\\', '/')}'\n")

            for i, segment in enumerate(audio_segments):
                f.write(f"file '{segment.replace('\\', '/')}'\n")
                
                # Decide pause type based on current and next element
                if i < len(audio_segments) - 1:
                    current_type = json_data["elements"][i]["type"]
                    next_type = json_data["elements"][i+1]["type"]
                    
                    if current_type == "chapter_title":
                        # Chapter Titles: 3.0 seconds mental reset
                        f.write(f"file '{pause_lg.replace('\\', '/')}'\n")
                    elif current_type == "section_headline":
                        # Section Headlines: 2.0 seconds
                        # We use 1.8s + an extra beat if needed, or just 2.0s (md+)
                        f.write(f"file '{pause_md.replace('\\', '/')}'\n")
                    elif current_type == "sub_heading":
                        # Sub-headers: 1.2 seconds
                        f.write(f"file '{pause_sh.replace('\\', '/')}'\n")
                    else:
                        # After a paragraph
                        if next_type in ["chapter_title", "section_headline"]:
                            # Getting ready for a major shift: increase breather
                            f.write(f"file '{pause_md.replace('\\', '/')}'\n")
                        else:
                            # Standard internal breath
                            f.write(f"file '{pause_sh.replace('\\', '/')}'\n")
            
            # 2. Trailing Tail
            f.write(f"file '{tail_sh.replace('\\', '/')}'\n")
        
        # Merge
        cmd_merge = [
            "ffmpeg", "-y", "-f", "concat", "-safe", "0",
            "-i", concat_file, "-c:a", "libmp3lame", "-q:a", "2",
            output_path
        ]
        subprocess.run(cmd_merge, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)
        
    return True

def step6_save_json(json_data, pdf_path):
    """
    6. save the json in a text file with the original name of the pdf.
    """
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    target_dir = "json" if os.path.isdir("json") else "."
    os.makedirs(target_dir, exist_ok=True)
    json_path = os.path.join(target_dir, f"{base_name}.json")
    
    print(f"[*] Saving structured JSON to: {json_path}")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, indent=2, ensure_ascii=False)
    return json_path

def step7_save_mp3(temp_mp3, pdf_path):
    """
    7. save the mp3 in a audio file with the original name of the pdf.
    """
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    target_dir = "mp3" if os.path.isdir("mp3") else "."
    os.makedirs(target_dir, exist_ok=True)
    final_mp3_path = os.path.join(target_dir, f"{base_name}.mp3")
    
    print(f"[*] Saving final MP3 to: {final_mp3_path}")
    shutil.copy(temp_mp3, final_mp3_path)
    return final_mp3_path

def process_pdf(pdf_path):
    """
    Executes the 7-step procedural pipeline for a single PDF.
    """
    try:
        # 2. Validate & Read
        text_content = step2_validate_and_read_pdf(pdf_path)

        # 3. Structured JSON
        json_data = step3_generate_structured_json(text_content, pdf_path)
        
        # 3b. Save JSON immediately as "Intermediary"
        step6_save_json(json_data, pdf_path)

        with tempfile.TemporaryDirectory() as temp_audio_dir:
            # 4. Record
            audio_segments = step4_ai_record_audiobook(json_data, temp_audio_dir)

            # 5. Merge
            temp_mp3 = os.path.join(temp_audio_dir, "output.mp3")
            step5_merge_with_pauses(audio_segments, json_data, temp_mp3)

            # 6. (Final JSON save - redundant but ensures it's correct if schema changed)
            step6_save_json(json_data, pdf_path)

            # 7. Save MP3
            step7_save_mp3(temp_mp3, pdf_path)

        print(f"\n[SUCCESS] Audiobook Generated: {os.path.basename(pdf_path)}")
        return True

    except Exception as e:
        print(f"\n[ERROR] Failed to process {pdf_path}: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Procedural PDF2MP3 Converter")
    parser.add_argument("path", nargs="?", default=None)
    parser.add_argument("--pdf-path", dest="pdf_path_alt")
    args = parser.parse_args()
    
    path_arg = args.path or args.pdf_path_alt
    
    # 1. Get Path (Detect 'all' or standard path)
    if path_arg and path_arg.lower() == "all":
        pdf_dir = "pdf"
        if not os.path.isdir(pdf_dir):
            print(f"[ERROR] 'pdf' directory not found for batch processing.")
            sys.exit(1)
        
        pdf_files = [os.path.join(pdf_dir, f) for f in os.listdir(pdf_dir) if f.lower().endswith(".pdf")]
        if not pdf_files:
            print("[INFO] No PDF files found in 'pdf' directory.")
            return

        print(f"[*] Batch processing {len(pdf_files)} files...")
        for pdf_path in pdf_files:
            process_pdf(pdf_path)
    else:
        # Single file mode (may trigger GUI if path_arg is None)
        pdf_path = step1_get_pdf_path(path_arg)
        if not pdf_path:
            return
        process_pdf(pdf_path)

if __name__ == "__main__":
    main()
