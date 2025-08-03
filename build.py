import PyInstaller.__main__
import os
import sys

def build_exe():
    """Build executable using PyInstaller"""
    
    # PyInstaller arguments
    args = [
        'pdf_renamer.py',
        '--onefile',
        '--windowed',
        '--name=PDF_Renamer_Scan',
        '--icon=icon.ico',  # Add icon if available
        '--add-data=config.txt;.',
        '--hidden-import=PIL._tkinter_finder',
        '--hidden-import=pytesseract',
        '--hidden-import=cv2',
        '--hidden-import=fitz',
        '--collect-all=pytesseract',
        '--collect-all=cv2',
        '--distpath=dist',
        '--workpath=build',
        '--specpath=.',
        '--clean'
    ]
    
    # Remove icon argument if icon file doesn't exist
    if not os.path.exists('icon.ico'):
        args = [arg for arg in args if not arg.startswith('--icon')]
    
    print("Building executable...")
    print("Arguments:", ' '.join(args))
    
    try:
        PyInstaller.__main__.run(args)
        print("\nBuild completed successfully!")
        print("Executable location: dist/PDF_Renamer_Scan.exe")
    except Exception as e:
        print(f"Build failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    build_exe()
