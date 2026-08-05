import os
import uuid
import shutil
import time
from typing import List
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import uvicorn
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, as_completed

from extractor import extract_text, parse_resume
from excel_generator import generate_excel

def process_single_file(filename, filepath):
    """
    Worker function for parsing a single resume.
    Needs to be at the module level so it can be pickled by multiprocessing.
    """
    try:
        with open(filepath, "rb") as f:
            contents = f.read()
            
        raw_text = extract_text(contents, filename)
        parsed_data = parse_resume(raw_text)
        parsed_data["Source File"] = filename
        role = parsed_data.get('Role', 'Unknown')
        
        return {
            "success": True,
            "filename": filename,
            "data": parsed_data,
            "role": role
        }
    except Exception as e:
        return {
            "success": False,
            "filename": filename,
            "error": str(e)
        }


app = FastAPI(title="Resume Extractor Dashboard")
app.mount("/static", StaticFiles(directory="static"), name="static")

# Global store for jobs
jobs = {}
TEMP_DIR = "temp_jobs"
os.makedirs(TEMP_DIR, exist_ok=True)

@app.get("/", response_class=HTMLResponse)
async def read_root():
    with open("static/index.html", "r") as f:
        return f.read()

def process_batch(job_id: str, file_paths: list):
    try:
        parsed_resumes = []
        total = len(file_paths)
        jobs[job_id]["total"] = total
        jobs[job_id]["active"] = 1
        
        # Max 4 workers to respect Render's free/starter tier CPU limits,
        # or use all available CPUs if less than 4.
        max_workers = min(multiprocessing.cpu_count(), len(file_paths), 4)
        if max_workers < 1:
            max_workers = 1
            
        jobs[job_id]["logs"].append(f"Starting batch processing with {max_workers} worker(s)... [ACTIVE]")
        
        with ProcessPoolExecutor(max_workers=max_workers) as executor:
            # Submit all jobs
            future_to_file = {
                executor.submit(process_single_file, filename, filepath): filename
                for filename, filepath in file_paths
            }
            
            # Process as they complete
            completed_count = 0
            for future in as_completed(future_to_file):
                completed_count += 1
                filename = future_to_file[future]
                jobs[job_id]["current_file"] = f"{completed_count}/{total} files processed"
                
                try:
                    result = future.result()
                    if result["success"]:
                        parsed_resumes.append(result["data"])
                        jobs[job_id]["processed"] += 1
                        jobs[job_id]["logs"].append(
                            f"[{completed_count}/{total}] {filename} categorized as {result['role']} [SUCCESS]"
                        )
                    else:
                        jobs[job_id]["failed"] += 1
                        jobs[job_id]["logs"].append(
                            f"[{completed_count}/{total}] Error parsing {filename}: {result.get('error')} [FAILED]"
                        )
                except Exception as exc:
                    jobs[job_id]["failed"] += 1
                    jobs[job_id]["logs"].append(
                        f"[{completed_count}/{total}] Critical error parsing {filename}: {exc} [FAILED]"
                    )
            
        jobs[job_id]["active"] = 0
        jobs[job_id]["current_file"] = "Generating Excel..."
        jobs[job_id]["logs"].append("All files processed. Generating Excel workbook... [ACTIVE]")
        
        excel_bytes = generate_excel(parsed_resumes)
        
        excel_path = os.path.join(TEMP_DIR, f"{job_id}.xlsx")
        with open(excel_path, "wb") as f:
            f.write(excel_bytes)
            
        jobs[job_id]["excel_path"] = excel_path
        jobs[job_id]["status"] = "completed"
        jobs[job_id]["logs"].append(f"Completed! {jobs[job_id]['processed']} resumes processed. [SUCCESS]")
        
    except Exception as e:
        jobs[job_id]["status"] = "failed"
        jobs[job_id]["error"] = str(e)
        jobs[job_id]["logs"].append(f"Critical Error in batch process: {str(e)} [FAILED]")
    finally:
        # Cleanup uploaded files and their parent job_id folder
        job_dir = os.path.join(TEMP_DIR, job_id)
        for _, filepath in file_paths:
            if os.path.exists(filepath):
                os.remove(filepath)
        if os.path.exists(job_dir) and not os.listdir(job_dir):
            os.rmdir(job_dir)

@app.post("/upload")
async def upload_resumes(background_tasks: BackgroundTasks, files: List[UploadFile] = File(...)):
    job_id = str(uuid.uuid4())
    job_dir = os.path.join(TEMP_DIR, job_id)
    os.makedirs(job_dir, exist_ok=True)
    
    file_paths = []
    for file in files:
        unique_filename = f"{uuid.uuid4().hex[:8]}_{file.filename}"
        filepath = os.path.join(job_dir, unique_filename)
        with open(filepath, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        file_paths.append((file.filename, filepath))
        
    jobs[job_id] = {
        "status": "processing",
        "total": len(file_paths),
        "processed": 0,
        "active": 0,
        "failed": 0,
        "current_file": "Initializing...",
        "excel_path": None,
        "error": None,
        "logs": ["Initializing batch job... [ACTIVE]"]
    }
    
    background_tasks.add_task(process_batch, job_id, file_paths)
    return {"job_id": job_id}

@app.get("/status/{job_id}")
async def get_status(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    return jobs[job_id]

@app.get("/debug")
async def debug_jobs():
    return jobs

@app.get("/download/{job_id}")
async def download_excel(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    
    job = jobs[job_id]
    if job["status"] != "completed" or not job["excel_path"]:
        raise HTTPException(status_code=400, detail="Job is not completed yet")
        
    return FileResponse(
        path=job["excel_path"],
        filename=f"extracted_resumes_{job_id[:8]}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
