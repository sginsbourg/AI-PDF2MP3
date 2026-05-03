import os
import urllib.request
import subprocess
import sys

def main():
    installer_path = "tesseract_installer.exe"
    url = "https://github.com/UB-Mannheim/tesseract/releases/download/v5.3.3.20231005/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
    print("Downloading Tesseract installer...")
    try:
        urllib.request.urlretrieve(url, installer_path)
    except Exception as e:
        print("Failed to download:", e)
        sys.exit(1)
        
    print("Running installer...")
    try:
        subprocess.run([installer_path, "/SILENT"], check=True)
        print("Installation complete.")
    except Exception as e:
        print("Failed to install:", e)
        sys.exit(1)
        
if __name__ == "__main__":
    main()
