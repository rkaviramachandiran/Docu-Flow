import os
import comtypes.client
import logging
import threading
import queue
import time
import uuid
import shutil
from PIL import Image
from docx import Document
from pdf2docx import Converter as PDFConverter

logger = logging.getLogger(__name__)

# Queue for conversion tasks
# Tuple format: (task_id, input_path, output_path, ext, target_format)
_conversion_queue = queue.Queue()

# Dictionary to store results
_results = {}

def _safe_remove(path):
    if os.path.exists(path):
        for _ in range(5):
            try:
                os.remove(path)
                return
            except:
                time.sleep(0.2)

def _conversion_worker():
    logger.info("Starting background conversion worker...")
    comtypes.CoInitialize()
    
    word = None
    excel = None
    
    try:
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = False
        word.DisplayAlerts = 0 # wdAlertsNone
        excel = comtypes.client.CreateObject('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        logger.info("Office applications pre-launched successfully.")
    except Exception as e:
        logger.warning(f"Could not pre-launch Office apps: {e}")

    while True:
        task = _conversion_queue.get()
        if task is None: break
            
        task_id, input_path, output_path, ext, target_format = task
        
        try:
            abs_out = os.path.abspath(output_path)
            
            # Special Case: Image Merging
            if ext == "images_merge":
                input_paths = input_path # It's a list in this case
                if target_format == 'pdf':
                    processed_imgs = []
                    for p in input_paths:
                        img = Image.open(p)
                        if img.mode != 'RGB': img = img.convert('RGB')
                        processed_imgs.append(img)
                    
                    if processed_imgs:
                        processed_imgs[0].save(abs_out, save_all=True, append_images=processed_imgs[1:])
                        for img in processed_imgs: img.close()
                
                elif target_format == 'docx':
                    doc_obj = Document()
                    for p in input_paths:
                        doc_obj.add_picture(os.path.abspath(p))
                    doc_obj.save(abs_out)
                
                _results[task_id] = {"status": "success"}
                return # Skip normal logic

            abs_in = os.path.abspath(input_path)

            if target_format == 'pdf':
                if ext in ['.docx', '.doc']:
                    if word is None: word = comtypes.client.CreateObject('Word.Application'); word.Visible = False; word.DisplayAlerts = 0
                    doc = word.Documents.Open(abs_in, ReadOnly=True)
                    try:
                        doc.SaveAs(abs_out, FileFormat=17) # wdFormatPDF
                    finally:
                        doc.Close(0) # wdDoNotSaveChanges
                elif ext in ['.xlsx', '.xls']:
                    if excel is None: excel = comtypes.client.CreateObject('Excel.Application'); excel.Visible = False; excel.DisplayAlerts = False
                    wb = excel.Workbooks.Open(abs_in, ReadOnly=True)
                    try:
                        wb.ExportAsFixedFormat(0, abs_out) # xlTypePDF
                    finally:
                        wb.Close(False)
                elif ext in ['.png', '.jpg', '.jpeg']:
                    img = Image.open(abs_in)
                    try:
                        if img.mode != 'RGB': img = img.convert('RGB')
                        img.save(abs_out, "PDF", resolution=100.0)
                    finally:
                        img.close()
                elif ext == '.txt':
                    temp_docx = input_path + ".docx"
                    doc_obj = Document()
                    with open(abs_in, 'r', encoding='utf-8', errors='ignore') as f:
                        for line in f: doc_obj.add_paragraph(line.strip())
                    doc_obj.save(os.path.abspath(temp_docx))
                    
                    if word is None: word = comtypes.client.CreateObject('Word.Application'); word.Visible = False; word.DisplayAlerts = 0
                    wdoc = word.Documents.Open(os.path.abspath(temp_docx), ReadOnly=True)
                    try:
                        wdoc.SaveAs(abs_out, FileFormat=17)
                    finally:
                        wdoc.Close(0)
                    _safe_remove(temp_docx)
                elif ext == '.pdf':
                    shutil.copy2(abs_in, abs_out)
                else:
                    raise ValueError(f"No PDF conversion path for {ext}")

            elif target_format == 'docx':
                if ext in ['.docx', '.doc']:
                    if ext == '.docx':
                        shutil.copy2(abs_in, abs_out)
                    else: # .doc to .docx
                        if word is None: word = comtypes.client.CreateObject('Word.Application'); word.Visible = False; word.DisplayAlerts = 0
                        doc = word.Documents.Open(abs_in, ReadOnly=True)
                        try:
                            doc.SaveAs(abs_out, FileFormat=16) # wdFormatXMLDocument
                        finally:
                            doc.Close(0)
                elif ext == '.pdf':
                    cv = PDFConverter(abs_in)
                    try:
                        cv.convert(abs_out)
                    finally:
                        cv.close()
                elif ext == '.txt':
                    doc_obj = Document()
                    with open(abs_in, 'r', encoding='utf-8', errors='ignore') as f:
                        for line in f: doc_obj.add_paragraph(line.strip())
                    doc_obj.save(abs_out)
                elif ext in ['.png', '.jpg', '.jpeg']:
                    doc_obj = Document()
                    doc_obj.add_picture(abs_in)
                    doc_obj.save(abs_out)
                else:
                    raise ValueError(f"No Word conversion path for {ext}")

            _results[task_id] = {"status": "success"}
        except Exception as e:
            logger.error(f"Conversion error for {task_id}: {e}")
            _results[task_id] = {"status": "error", "message": str(e)}
        finally:
            _conversion_queue.task_done()

    if word:
        try: word.Quit()
        except: pass
    if excel:
        try: excel.Quit()
        except: pass
    comtypes.CoUninitialize()

_worker_thread = threading.Thread(target=_conversion_worker, daemon=True)
_worker_thread.start()

def convert_to_pdf(input_path: str, output_path: str, ext: str, target_format: str = "pdf"):
    task_id = str(uuid.uuid4())
    _results[task_id] = {"status": "pending"}
    _conversion_queue.put((task_id, input_path, output_path, ext, target_format))
    while True:
        res = _results.get(task_id)
        if res["status"] == "success":
            del _results[task_id]
            return
        elif res["status"] == "error":
            err_msg = res["message"]
            del _results[task_id]
            raise Exception(err_msg)
        time.sleep(0.1)
