import fitz
import re
from collections import Counter
import sys

def clean_boilerplate_candidate(line):
    l = line.strip()
    l = re.sub(r'(?i)page\s+\d+(\s+of\s+\d+)?', '', l)
    l = re.sub(r'^\d+$', '', l)
    return l.strip()

def diag(pdf_path):
    doc = fitz.open(pdf_path)
    all_lines = []
    print(f"Total Pages: {len(doc)}")
    for i in range(min(15, len(doc))):
        page = doc[i]
        text = page.get_text()
        lines = text.split('\n')
        print(f"--- Page {i+1} ---")
        for j, l in enumerate(lines[:5]):
            print(f"L{j}: {repr(l)} -> Cleaned: {repr(clean_boilerplate_candidate(l))}")

if __name__ == "__main__":
    diag(sys.argv[1])
