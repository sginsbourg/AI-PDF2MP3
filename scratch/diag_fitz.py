import fitz
import sys

def diag_fitz(path):
    print(f"Opening {path} with PyMuPDF...")
    try:
        doc = fitz.open(path)
        print(f"Pages: {len(doc)}")
        for i, page in enumerate(doc):
            text = page.get_text()
            print(f"Page {i+1}: length={len(text)}")
            if text:
                print(f"Snippet: {text[:100]}...")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    diag_fitz(sys.argv[1])
