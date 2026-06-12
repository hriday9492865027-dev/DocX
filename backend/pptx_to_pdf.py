import os
import shutil
from libreoffice_helper import IS_WINDOWS, convert_with_libreoffice

if IS_WINDOWS:
    import pythoncom
    import win32com.client

def pptx_to_pdf(pptx_path, pdf_path):
    if IS_WINDOWS:
        try:
            pythoncom.CoInitialize()
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            pres = powerpoint.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
            pres.SaveAs(os.path.abspath(pdf_path), FileFormat=32) # 32 is ppSaveAsPDF
            pres.Close()
            powerpoint.Quit()
            return
        except Exception as e:
            print(f"Windows COM automation failed: {e}. Falling back to LibreOffice CLI...")
        finally:
            pythoncom.CoUninitialize()
            
    # Linux / Fallback: Run LibreOffice
    output_dir = os.path.dirname(pdf_path)
    convert_with_libreoffice(pptx_path, output_dir)
    
    filename_no_ext = os.path.basename(pptx_path).rsplit('.', 1)[0]
    expected_out = os.path.join(output_dir, f"{filename_no_ext}.pdf")
    if os.path.exists(expected_out):
        shutil.move(expected_out, pdf_path)
    else:
        raise Exception("LibreOffice CLI did not produce PDF output.")
