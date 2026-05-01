from PyPDF2 import PdfReader
import sys

def diag_pdf(path):
    print(f"Opening {path}...")
    try:
        reader = PdfReader(path)
        print(f"Pages: {len(reader.pages)}")
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            print(f"Page {i+1}: length={len(text) if text else 0}")
            if text:
                print(f"Snippet: {text[:100]}...")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    diag_pdf(sys.argv[1])
