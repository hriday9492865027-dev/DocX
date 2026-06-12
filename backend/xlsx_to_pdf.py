import os
import shutil
from libreoffice_helper import IS_WINDOWS, convert_with_libreoffice

if IS_WINDOWS:
    import pythoncom
    import win32com.client

def xlsx_to_pdf(xlsx_path, pdf_path):
    if IS_WINDOWS:
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path)) # 0 is xlTypePDF
            wb.Close(SaveChanges=False)
            excel.Quit()
            return
        except Exception as e:
            print(f"Windows COM automation failed: {e}. Falling back to LibreOffice CLI...")
        finally:
            pythoncom.CoUninitialize()
            
    # Linux / Fallback: Run LibreOffice
    output_dir = os.path.dirname(pdf_path)
    convert_with_libreoffice(xlsx_path, output_dir)
    
    filename_no_ext = os.path.basename(xlsx_path).rsplit('.', 1)[0]
    expected_out = os.path.join(output_dir, f"{filename_no_ext}.pdf")
    if os.path.exists(expected_out):
        shutil.move(expected_out, pdf_path)
    else:
        raise Exception("LibreOffice CLI did not produce PDF output.")
