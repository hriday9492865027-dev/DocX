import os
import shutil
from libreoffice_helper import IS_WINDOWS, convert_with_libreoffice

if IS_WINDOWS:
    import pythoncom
    import win32com.client

def docx_to_pdf(docx_path, pdf_path):
    if IS_WINDOWS:
        try:
            pythoncom.CoInitialize()
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(docx_path))
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17) # 17 is wdFormatPDF
            doc.Close()
            word.Quit()
            return
        except Exception as e:
            print(f"Windows COM automation failed: {e}. Falling back to LibreOffice CLI...")
        finally:
            pythoncom.CoUninitialize()
            
    # Linux / Fallback: Run LibreOffice
    output_dir = os.path.dirname(pdf_path)
    convert_with_libreoffice(docx_path, output_dir)
    
    filename_no_ext = os.path.basename(docx_path).rsplit('.', 1)[0]
    expected_out = os.path.join(output_dir, f"{filename_no_ext}.pdf")
    if os.path.exists(expected_out):
        shutil.move(expected_out, pdf_path)
    else:
        raise Exception("LibreOffice CLI did not produce PDF output.")
