import os
import sys
import shutil
import subprocess

IS_WINDOWS = sys.platform.startswith('win')
if IS_WINDOWS:
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("win32com or pythoncom not available. Running in LibreOffice-only mode.")
        IS_WINDOWS = False

def convert_with_libreoffice(input_path, output_dir):
    """Converts a document to PDF using LibreOffice headless command line."""
    soffice_path = shutil.which('soffice') or shutil.which('libreoffice')
    
    if not soffice_path and sys.platform.startswith('win'):
        possible_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ]
        for p in possible_paths:
            if os.path.exists(p):
                soffice_path = p
                break
                
    if not soffice_path:
        raise Exception(
            "LibreOffice (soffice) CLI not found. Make sure LibreOffice is installed "
            "and added to system variables/PATH (or MS Office is installed on Windows)."
        )
        
    cmd = [
        soffice_path,
        '--headless',
        '--convert-to', 'pdf',
        input_path,
        '--outdir', output_dir
    ]
    
    print(f"Executing LibreOffice conversion: {' '.join(cmd)}")
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if result.returncode != 0:
        raise Exception(f"LibreOffice conversion failed: {result.stderr or result.stdout}")
