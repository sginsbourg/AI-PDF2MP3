import fitz
import sys

def diag_deep(path):
    print(f"Deep Diag: {path}")
    doc = fitz.open(path)
    print(f"Encrypted: {doc.is_encrypted}")
    for i, page in enumerate(doc):
        print(f"\n--- Page {i+1} ---")
        text = page.get_text().strip()
        print(f"Text Length: {len(text)}")
        images = page.get_images()
        print(f"Images count: {len(images)}")
        if len(text) == 0 and len(images) > 0:
            print("Status: Likely a SCANNED PDF (images present, no text).")
        elif len(text) == 0 and len(images) == 0:
            print("Status: Truly empty or non-standard vector paths.")

if __name__ == "__main__":
    diag_deep(sys.argv[1])
