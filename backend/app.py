# AetherPDF Backend - Python Document Transmutation Server
import os
import uuid
import shutil
import pythoncom
import win32com.client
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

# Import libraries for local sandboxed conversions
from pdf2docx import Converter
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches
import io
import pdfplumber
import openpyxl

app = Flask(__name__)
CORS(app)

TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp')
os.makedirs(TEMP_DIR, exist_ok=True)

# ==========================================
# --- CONVERSION ALGORITHMS ---
# ==========================================

# 1. PDF to DOCX
def pdf_to_docx(pdf_path, docx_path):
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()

# 2. PDF to PPTX
def pdf_to_pptx(pdf_path, pptx_path):
    doc = fitz.open(pdf_path)
    prs = Presentation()
    blank_layout = prs.slide_layouts[6] # blank layout
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=150)
        img_data = pix.tobytes("png")
        
        # Match dimensions (72 points = 1 inch)
        if page_num == 0:
            prs.slide_width = Inches(page.rect.width / 72.0)
            prs.slide_height = Inches(page.rect.height / 72.0)
            
        slide = prs.slides.add_slide(blank_layout)
        img_stream = io.BytesIO(img_data)
        slide.shapes.add_picture(img_stream, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
        
    prs.save(pptx_path)

# 3. PDF to XLSX
def pdf_to_xlsx(pdf_path, xlsx_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            sheet = wb.create_sheet(title=f"Page {i+1}")
            tables = page.extract_tables()
            row_idx = 1
            
            for table in tables:
                for row in table:
                    for col_idx, cell in enumerate(row, start=1):
                        sheet.cell(row=row_idx, column=col_idx, value=cell)
                    row_idx += 1
                row_idx += 2 # spacing between tables
                
            if not tables:
                # Text extraction fallback
                text = page.extract_text()
                if text:
                    for line in text.split('\n'):
                        parts = [p.strip() for p in line.split('   ') if p.strip()]
                        if not parts:
                            parts = line.split('\t')
                        for col_idx, val in enumerate(parts, start=1):
                            sheet.cell(row=row_idx, column=col_idx, value=val)
                        row_idx += 1
                        
    if not wb.sheetnames:
        wb.create_sheet(title="Sheet 1")
    wb.save(xlsx_path)

# 4. DOCX to PDF (Using Word COM automation)
def docx_to_pdf(docx_path, pdf_path):
    pythoncom.CoInitialize()
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17) # 17 is wdFormatPDF
        doc.Close()
        word.Quit()
    except Exception as e:
        raise Exception(f"Microsoft Word export failed: {str(e)}. Make sure MS Word is installed.")
    finally:
        pythoncom.CoUninitialize()

# 5. PPTX to PDF (Using PowerPoint COM automation)
def pptx_to_pdf(pptx_path, pdf_path):
    pythoncom.CoInitialize()
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        pres = powerpoint.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
        pres.SaveAs(os.path.abspath(pdf_path), FileFormat=32) # 32 is ppSaveAsPDF
        pres.Close()
        powerpoint.Quit()
    except Exception as e:
        raise Exception(f"Microsoft PowerPoint export failed: {str(e)}. Make sure MS PowerPoint is installed.")
    finally:
        pythoncom.CoUninitialize()

# 6. XLSX to PDF (Using Excel COM automation)
def xlsx_to_pdf(xlsx_path, pdf_path):
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path)) # 0 is xlTypePDF
        wb.Close(SaveChanges=False)
        excel.Quit()
    except Exception as e:
        raise Exception(f"Microsoft Excel export failed: {str(e)}. Make sure MS Excel is installed.")
    finally:
        pythoncom.CoUninitialize()

# ==========================================
# --- ROUTE HANDLERS ---
# ==========================================

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files or 'mode' not in request.form:
        return jsonify({"error": "Missing file or mode parameters"}), 400
        
    uploaded_file = request.files['file']
    mode = request.form['mode']
    
    if uploaded_file.filename == '':
        return jsonify({"error": "No file selected"}), 400
        
    # Generate unique filenames to prevent conflicts
    task_id = str(uuid.uuid4())
    filename = uploaded_file.filename
    ext = filename.split('.')[-1].lower()
    
    src_path = os.path.join(TEMP_DIR, f"{task_id}.{ext}")
    dest_ext = 'pdf' if 'to-pdf' in mode else ('docx' if 'to-doc' in mode else ('pptx' if 'to-ppt' in mode else 'xlsx'))
    dest_path = os.path.join(TEMP_DIR, f"{task_id}_out.{dest_ext}")
    
    # Save input file
    uploaded_file.save(src_path)
    
    try:
        if mode == 'pdf-to-doc':
            pdf_to_docx(src_path, dest_path)
        elif mode == 'pdf-to-ppt':
            pdf_to_pptx(src_path, dest_path)
        elif mode == 'pdf-to-sheets':
            pdf_to_xlsx(src_path, dest_path)
        elif mode == 'doc-to-pdf':
            docx_to_pdf(src_path, dest_path)
        elif mode == 'ppt-to-pdf':
            pptx_to_pdf(src_path, dest_path)
        elif mode == 'sheets-to-pdf':
            xlsx_to_pdf(src_path, dest_path)
        else:
            return jsonify({"error": f"Invalid conversion mode: {mode}"}), 400
            
        if not os.path.exists(dest_path):
            raise Exception("Output file was not generated by conversion module.")
            
        # Return converted file
        return send_file(
            dest_path,
            as_attachment=True,
            download_name=f"{filename.rsplit('.', 1)[0]}_converted.{dest_ext}"
        )
        
    except Exception as e:
        print(f"Error during conversion: {str(e)}")
        return jsonify({"error": str(e)}), 500
        
    finally:
        # Cleanup temp files
        try:
            if os.path.exists(src_path):
                os.remove(src_path)
            if os.path.exists(dest_path):
                os.remove(dest_path)
        except Exception as cleanup_err:
            print(f"Failed to delete temp files: {cleanup_err}")

if __name__ == '__main__':
    print("AetherPDF Python backend listening on http://localhost:5000...")
    app.run(port=5000, debug=False)
