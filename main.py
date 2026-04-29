import os
import shutil
import tempfile
import zipfile
import time
from typing import List
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks, Form
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from converter import convert_to_pdf

app = FastAPI(title="DocuFlow API")

os.makedirs("static", exist_ok=True)
app.mount("/static", StaticFiles(directory="static"), name="static")

TEMP_DIR = tempfile.gettempdir()
APP_TEMP_DIR = os.path.join(TEMP_DIR, "doc_converter_files")
os.makedirs(APP_TEMP_DIR, exist_ok=True)

def cleanup_file(file_path: str):
    if os.path.exists(file_path):
        # On Windows, COM apps might hold a lock for a split second after close.
        # We try up to 3 times with a small delay.
        for _ in range(3):
            try:
                os.remove(file_path)
                break
            except Exception as e:
                time.sleep(0.2)
        else:
            print(f"Warning: Failed to clean up file after retries {file_path}")

@app.get("/")
async def read_index():
    return FileResponse(os.path.join(os.getcwd(), "static", "index.html"))

@app.get("/word")
async def read_word():
    return FileResponse(os.path.join(os.getcwd(), "static", "word.html"))

@app.post("/convert")
async def convert_file(background_tasks: BackgroundTasks, files: List[UploadFile] = File(...), output_format: str = Form("pdf")):
    if not files:
        raise HTTPException(status_code=400, detail="No files provided")
    
    import uuid
    batch_id = str(uuid.uuid4())
    
    # Check if all files are images for merging
    image_exts = ['.png', '.jpg', '.jpeg']
    all_images = True
    valid_files = []
    for file in files:
        if not file.filename: continue
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in image_exts: all_images = False
        valid_files.append(file)

    converted_results = []
    
    # Case: Multiple images and target is single output
    if all_images and len(valid_files) > 1:
        file_id = str(uuid.uuid4())
        actual_output_ext = f".{output_format}"
        output_path = os.path.join(APP_TEMP_DIR, f"{file_id}{actual_output_ext}")
        
        input_paths = []
        for file in valid_files:
            fid = str(uuid.uuid4())
            ext = os.path.splitext(file.filename)[1].lower()
            in_path = os.path.join(APP_TEMP_DIR, f"{fid}{ext}")
            with open(in_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            input_paths.append(in_path)
            
        try:
            # Special call for merging
            convert_to_pdf(input_paths, output_path, "images_merge", output_format)
            converted_results.append({
                "file_id": file_id,
                "name": f"Merged_Images{actual_output_ext}",
                "ext": actual_output_ext
            })
        except Exception as e:
            for p in input_paths: cleanup_file(p)
            raise HTTPException(status_code=500, detail=f"Merge failed. Error: {e}")
        finally:
            for p in input_paths: cleanup_file(p)
            
    else:
        # Normal individual conversion
        for file in valid_files:
            ext = os.path.splitext(file.filename)[1].lower()
            allowed_exts = ['.docx', '.doc', '.xlsx', '.xls', '.png', '.jpg', '.jpeg', '.txt', '.pdf']
            if ext not in allowed_exts: continue
                
            file_id = str(uuid.uuid4())
            input_path = os.path.join(APP_TEMP_DIR, f"{file_id}{ext}")
            actual_output_ext = f".{output_format}"
            output_path = os.path.join(APP_TEMP_DIR, f"{file_id}{actual_output_ext}")
            
            with open(input_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            
            try:
                convert_to_pdf(input_path, output_path, ext, output_format)
                converted_results.append({
                    "file_id": file_id, 
                    "name": f"{os.path.splitext(file.filename)[0]}{actual_output_ext}",
                    "ext": actual_output_ext
                })
            except Exception as e:
                cleanup_file(input_path)
                raise HTTPException(status_code=500, detail=f"Conversion failed. Error: {e}")
            finally:
                cleanup_file(input_path)
    
    if not converted_results:
        raise HTTPException(status_code=400, detail="No valid files were converted")

    return JSONResponse(content={
        "files": converted_results,
        "message": "Conversion successful"
    })

@app.get("/download/{file_id}")
async def download_file(file_id: str, background_tasks: BackgroundTasks, name: str = "document"):
    # Check for both .pdf and .docx
    for ext in ['.pdf', '.docx']:
        file_path = os.path.join(APP_TEMP_DIR, f"{file_id}{ext}")
        if os.path.exists(file_path):
            background_tasks.add_task(cleanup_file, file_path)
            media_type = "application/pdf" if ext == ".pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            return FileResponse(
                path=file_path,
                filename=name,
                media_type=media_type
            )
    
    raise HTTPException(status_code=404, detail="File not found or expired")

@app.get("/")
async def root():
    return FileResponse("static/index.html")

@app.get("/word")
async def word_mode():
    return FileResponse("static/word.html")
