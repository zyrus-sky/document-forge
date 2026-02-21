import os
import csv
import re
import uuid
import shutil
import zipfile
import threading
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
from pydantic import BaseModel
from typing import Dict, List, Optional
from docx import Document
import win32com.client

app = FastAPI(title="Document Forge API")

# Enable CORS for SvelteKit Dev Server
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = Path("temp_sessions")
TEMP_DIR.mkdir(exist_ok=True)
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)

WD_FORMAT_PDF = 17

class MappingItem(BaseModel):
    type: str # 'csv_column', 'custom_text', or 'combined'
    value: str
    prefix: Optional[str] = ""
    suffix: Optional[str] = ""

class GenerateRequest(BaseModel):
    session_id: str
    mapping: Dict[str, MappingItem]
    generate_docx: bool = True
    generate_pdf: bool = True

def cleanup_session(session_id: str):
    """Clean up temp files for a session after generation"""
    session_dir = TEMP_DIR / session_id
    if session_dir.exists():
        shutil.rmtree(session_dir, ignore_errors=True)

@app.post("/api/upload")
async def upload_files(csv_file: UploadFile = File(...), template_file: UploadFile = File(...)):
    session_id = str(uuid.uuid4())
    session_dir = TEMP_DIR / session_id
    session_dir.mkdir(parents=True, exist_ok=True)
    
    csv_path = session_dir / "input.csv"
    template_path = session_dir / "template.docx"
    
    with open(csv_path, "wb") as f:
        f.write(await csv_file.read())
        
    with open(template_path, "wb") as f:
        f.write(await template_file.read())
        
    return {"session_id": session_id}

@app.get("/api/metadata")
async def get_metadata(session_id: str):
    session_dir = TEMP_DIR / session_id
    csv_path = session_dir / "input.csv"
    template_path = session_dir / "template.docx"
    
    if not csv_path.exists() or not template_path.exists():
        raise HTTPException(status_code=404, detail="Session not found or files missing")
        
    try:
        # 1. Parse CSV Headers
        headers = []
        rows_count = 0
        with open(csv_path, "r", encoding="utf-8") as f:
            reader = csv.reader(f)
            headers_raw = next(reader, [])
            # Apply normalization
            headers = [re.sub(r"\s+", "", h.strip()) for h in headers_raw]
            headers.append("P_ ADDRESS")
            rows_count = sum(1 for row in reader if any(v.strip() for v in row))
            
        # 2. Parse DOCX Placeholders
        doc = Document(template_path)
        placeholders = set()
        
        def extract_from_text(text):
            found = re.findall(r'#\w+(?:\s+\w+)?', text)
            for f in found:
                 placeholders.add(f)
        
        for paragraph in doc.paragraphs:
            extract_from_text(paragraph.text)
            
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        extract_from_text(paragraph.text)
                        
        for section in doc.sections:
            for header_footer in (section.header, section.footer):
                if header_footer is not None:
                    for paragraph in header_footer.paragraphs:
                        extract_from_text(paragraph.text)

        return {
            "csv_headers": sorted(list(set(headers))),
            "docx_placeholders": sorted(list(placeholders)),
            "total_rows": rows_count
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/generate")
async def generate_documents(request: GenerateRequest, background_tasks: BackgroundTasks):
    session_id = request.session_id
    session_dir = TEMP_DIR / session_id
    csv_path = session_dir / "input.csv"
    template_path = session_dir / "template.docx"
    
    if not csv_path.exists() or not template_path.exists():
        raise HTTPException(status_code=404, detail="Session not found")
        
    out_dir = session_dir / "output_files"
    out_dir.mkdir(exist_ok=True)
    
    try:
        # Read CSV logic
        rows = []
        with open(csv_path, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if all(v.strip() == "" for v in row.values()):
                    continue
                record = {}
                for header, value in row.items():
                    key = re.sub(r"\s+", "", header.strip())
                    record[key] = value.strip() if value else ""
                rows.append(record)
                
        # Prep Word COM if PDF needed
        word_app = None
        if request.generate_pdf:
            try:
                word_app = win32com.client.Dispatch("Word.Application")
                word_app.Visible = False
                word_app.DisplayAlerts = False
            except Exception as e:
                raise HTTPException(status_code=500, detail=f"Word PDF Engine failed: {str(e)}")
        
        for i, record in enumerate(rows, start=1):
            replacements = {}
            for docx_tag, map_item in request.mapping.items():
                if map_item.type == 'custom_text':
                    replacements[docx_tag] = map_item.value
                elif map_item.type == 'csv_column':
                    lookup_col = "P_ADDRESS" if map_item.value == "P_ ADDRESS"  else map_item.value
                    replacements[docx_tag] = record.get(lookup_col, "")
                elif map_item.type == 'combined':
                    lookup_col = "P_ADDRESS" if map_item.value == "P_ ADDRESS"  else map_item.value
                    val = record.get(lookup_col, "")
                    if val:  # only apply prefix/suffix if the value exists
                        replacements[docx_tag] = f"{map_item.prefix}{val}{map_item.suffix}"
                    else:
                        replacements[docx_tag] = ""
                    
            # Sort longest key first
            replacements = dict(sorted(replacements.items(), key=lambda kv: len(kv[0]), reverse=True))

            name = record.get("NAME", f"row_{i}").strip().replace(" ", "_")
            if not name: name = f"row_{i}"
            filename_base = f"{i:04d}_{name}"
            
            # Fill document
            doc = _fill_template(template_path, replacements)
            out_docx = out_dir / f"{filename_base}.docx"
            doc.save(out_docx)
            
            # Convert to PDF
            if word_app and request.generate_pdf:
                out_pdf = out_dir / f"{filename_base}.pdf"
                pdf_doc = word_app.Documents.Open(str(out_docx.absolute()))
                pdf_doc.SaveAs(str(out_pdf.absolute()), FileFormat=WD_FORMAT_PDF)
                pdf_doc.Close(False)
                
                if not request.generate_docx:
                    os.remove(out_docx)

        if word_app:
            word_app.Quit()
            
        # Create ZIP file
        zip_path = session_dir / "DocumentForge_Output.zip"
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(out_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, out_dir)
                    zipf.write(file_path, arcname)
                    
        # Schedule cleanup after download
        background_tasks.add_task(cleanup_session, session_id)
        
        return FileResponse(
            path=zip_path, 
            filename="DocumentForge_Output.zip",
            media_type="application/zip"
        )
        
    except Exception as e:
        if word_app:
            word_app.Quit()
        raise HTTPException(status_code=500, detail=str(e))

# Helpers
def _replace_in_paragraph(paragraph, replacements):
    full_text = paragraph.text
    if "#" not in full_text: return

    for run in paragraph.runs:
        for placeholder, value in replacements.items():
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, value)

    remaining = paragraph.text
    for placeholder, value in replacements.items():
        if placeholder in remaining:
            new_text = remaining
            for ph, val in replacements.items():
                new_text = new_text.replace(ph, val)
            if paragraph.runs:
                paragraph.runs[0].text = new_text
                for run in paragraph.runs[1:]:
                    run.text = ""
            break

def _replace_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                _replace_in_paragraph(paragraph, replacements)

def _fill_template(template_path, replacements):
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        _replace_in_paragraph(paragraph, replacements)
    for table in doc.tables:
        _replace_in_table(table, replacements)
    for section in doc.sections:
        for header_footer in (section.header, section.footer):
            if header_footer is not None:
                for paragraph in header_footer.paragraphs:
                    _replace_in_paragraph(paragraph, replacements)
    return doc

if __name__ == "__main__":
    print("🚀 Starting Document Forge Fast API server on port 8000...")
    uvicorn.run("backend:app", host="127.0.0.1", port=8000, reload=True)
