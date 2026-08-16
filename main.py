import os
import uuid
import shutil
import time
import math
from typing import List, Optional
from contextlib import asynccontextmanager

from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException, Depends, Query, Response
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import uvicorn
import multiprocessing
from concurrent.futures import ProcessPoolExecutor, as_completed
from sqlalchemy.orm import Session
from sqlalchemy import or_, func, desc, asc

from extractor import extract_text, parse_resume
from excel_generator import generate_excel
from database import init_db, get_db, SessionLocal, Resume


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Initialize database tables on startup
    init_db()
    yield


app = FastAPI(title="Resume Extractor & Candidate Database", lifespan=lifespan)
app.mount("/static", StaticFiles(directory="static"), name="static")

# Global store for active batch jobs
jobs = {}
TEMP_DIR = "temp_jobs"
os.makedirs(TEMP_DIR, exist_ok=True)


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
        role = parsed_data.get("Role", "Unknown")

        return {
            "success": True,
            "filename": filename,
            "data": parsed_data,
            "role": role,
        }
    except Exception as e:
        return {
            "success": False,
            "filename": filename,
            "error": str(e),
        }


def save_resumes_to_db(job_id: str, parsed_resumes: list):
    """Saves a batch of parsed resumes to the PostgreSQL / SQLite database."""
    db = SessionLocal()
    try:
        for r in parsed_resumes:
            resume_record = Resume(
                job_id=job_id,
                name=r.get("Name", ""),
                contact_no=r.get("Contact No", ""),
                email=r.get("Email", ""),
                role=r.get("Role", "Uncategorized"),
                skills=r.get("Skills", ""),
                education=r.get("Education", ""),
                projects=r.get("Projects", ""),
                experience=r.get("Experience & Internships", ""),
                area_of_interest=r.get("Area of Interest / Objective", ""),
                awards=r.get("Awards & Achievements", ""),
                extra_curriculars=r.get("Extra Curriculars & Leadership", ""),
                registration_no=r.get("Registration No", ""),
                linkedin=r.get("LinkedIn", ""),
                github=r.get("GitHub", ""),
                source_file=r.get("Source File", ""),
            )
            db.add(resume_record)
        db.commit()
    except Exception as e:
        db.rollback()
        print(f"Error saving resumes to DB: {e}")
    finally:
        db.close()


def process_batch(job_id: str, file_paths: list):
    try:
        parsed_resumes = []
        total = len(file_paths)
        jobs[job_id]["total"] = total
        jobs[job_id]["active"] = 1

        # Max 4 workers to respect cloud instance CPU limits
        max_workers = min(multiprocessing.cpu_count(), len(file_paths), 4)
        if max_workers < 1:
            max_workers = 1

        jobs[job_id]["logs"].append(f"Starting batch processing with {max_workers} worker(s)... [ACTIVE]")

        with ProcessPoolExecutor(max_workers=max_workers) as executor:
            future_to_file = {
                executor.submit(process_single_file, filename, filepath): filename
                for filename, filepath in file_paths
            }

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

        # Save to Database
        jobs[job_id]["current_file"] = "Saving to Database..."
        jobs[job_id]["logs"].append("Saving parsed resumes to Database... [ACTIVE]")
        save_resumes_to_db(job_id, parsed_resumes)
        jobs[job_id]["logs"].append("Saved to Database. [SUCCESS]")

        # Generate Excel workbook
        jobs[job_id]["current_file"] = "Generating Excel..."
        jobs[job_id]["logs"].append("Generating Excel workbook... [ACTIVE]")

        excel_bytes = generate_excel(parsed_resumes)
        excel_path = os.path.join(TEMP_DIR, f"{job_id}.xlsx")
        with open(excel_path, "wb") as f:
            f.write(excel_bytes)

        jobs[job_id]["excel_path"] = excel_path
        jobs[job_id]["status"] = "completed"
        jobs[job_id]["logs"].append(f"Completed! {jobs[job_id]['processed']} resumes processed & stored. [SUCCESS]")

    except Exception as e:
        jobs[job_id]["status"] = "failed"
        jobs[job_id]["error"] = str(e)
        jobs[job_id]["logs"].append(f"Critical Error in batch process: {str(e)} [FAILED]")
    finally:
        # Cleanup uploaded files and their parent job_id folder
        job_dir = os.path.join(TEMP_DIR, job_id)
        for _, filepath in file_paths:
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception:
                    pass
        if os.path.exists(job_dir) and not os.listdir(job_dir):
            try:
                os.rmdir(job_dir)
            except Exception:
                pass


@app.get("/", response_class=HTMLResponse)
async def read_root():
    with open("static/index.html", "r") as f:
        return f.read()


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
        "logs": ["Initializing batch job... [ACTIVE]"],
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
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ═══════════════════════════════════════════════════════════════════════════════
# CANDIDATE EXPLORER REST API ENDPOINTS
# ═══════════════════════════════════════════════════════════════════════════════

@app.get("/api/candidates")
def list_candidates(
    search: Optional[str] = Query(None, description="Search keyword across all fields"),
    role: Optional[str] = Query(None, description="Filter by candidate role"),
    skill: Optional[str] = Query(None, description="Filter by specific skill"),
    sort_by: str = Query("created_at", description="Sort field: created_at, name, role"),
    sort_order: str = Query("desc", description="Sort direction: asc, desc"),
    page: int = Query(1, ge=1, description="Page number"),
    limit: int = Query(15, ge=1, le=100, description="Items per page"),
    db: Session = Depends(get_db),
):
    """Search and filter candidates from the database with pagination."""
    query = db.query(Resume)

    # Filter by role
    if role and role.strip() and role.lower() != "all":
        query = query.filter(Resume.role == role.strip())

    # Filter by specific skill
    if skill and skill.strip():
        query = query.filter(Resume.skills.ilike(f"%{skill.strip()}%"))

    # Global search filter across multiple fields
    if search and search.strip():
        term = f"%{search.strip()}%"
        query = query.filter(
            or_(
                Resume.name.ilike(term),
                Resume.email.ilike(term),
                Resume.contact_no.ilike(term),
                Resume.role.ilike(term),
                Resume.skills.ilike(term),
                Resume.projects.ilike(term),
                Resume.experience.ilike(term),
                Resume.education.ilike(term),
                Resume.registration_no.ilike(term),
                Resume.source_file.ilike(term),
            )
        )

    total = query.count()

    # Sorting
    sort_column = getattr(Resume, sort_by, Resume.created_at)
    if sort_order.lower() == "asc":
        query = query.order_by(asc(sort_column))
    else:
        query = query.order_by(desc(sort_column))

    # Pagination
    offset = (page - 1) * limit
    candidates = query.offset(offset).limit(limit).all()

    total_pages = math.ceil(total / limit) if total > 0 else 1

    return {
        "candidates": [c.to_dict() for c in candidates],
        "total": total,
        "page": page,
        "pages": total_pages,
        "limit": limit,
    }


@app.get("/api/candidates/{candidate_id}")
def get_candidate(candidate_id: int, db: Session = Depends(get_db)):
    """Retrieve full details of a specific candidate."""
    candidate = db.query(Resume).filter(Resume.id == candidate_id).first()
    if not candidate:
        raise HTTPException(status_code=404, detail="Candidate not found")
    return candidate.to_dict()


@app.delete("/api/candidates/{candidate_id}")
def delete_candidate(candidate_id: int, db: Session = Depends(get_db)):
    """Delete a single candidate from the database."""
    candidate = db.query(Resume).filter(Resume.id == candidate_id).first()
    if not candidate:
        raise HTTPException(status_code=404, detail="Candidate not found")
    db.delete(candidate)
    db.commit()
    return {"message": "Candidate deleted successfully", "id": candidate_id}


class BulkDeleteRequest(BaseModel):
    ids: List[int]


@app.post("/api/candidates/bulk-delete")
def bulk_delete_candidates(req: BulkDeleteRequest, db: Session = Depends(get_db)):
    """Bulk delete multiple candidates from the database."""
    if not req.ids:
        return {"deleted_count": 0}
    deleted_count = db.query(Resume).filter(Resume.id.in_(req.ids)).delete(synchronize_session=False)
    db.commit()
    return {"deleted_count": deleted_count}


class ExportRequest(BaseModel):
    ids: Optional[List[int]] = None
    search: Optional[str] = None
    role: Optional[str] = None
    skill: Optional[str] = None


@app.post("/api/candidates/export")
def export_candidates_excel(req: ExportRequest, db: Session = Depends(get_db)):
    """Exports matching or selected candidate records to a formatted Excel workbook."""
    query = db.query(Resume)

    if req.ids and len(req.ids) > 0:
        query = query.filter(Resume.id.in_(req.ids))
    else:
        if req.role and req.role.strip() and req.role.lower() != "all":
            query = query.filter(Resume.role == req.role.strip())
        if req.skill and req.skill.strip():
            query = query.filter(Resume.skills.ilike(f"%{req.skill.strip()}%"))
        if req.search and req.search.strip():
            term = f"%{req.search.strip()}%"
            query = query.filter(
                or_(
                    Resume.name.ilike(term),
                    Resume.email.ilike(term),
                    Resume.contact_no.ilike(term),
                    Resume.role.ilike(term),
                    Resume.skills.ilike(term),
                    Resume.projects.ilike(term),
                    Resume.experience.ilike(term),
                    Resume.education.ilike(term),
                )
            )

    candidates = query.order_by(desc(Resume.created_at)).all()
    if not candidates:
        raise HTTPException(status_code=400, detail="No candidate records found to export.")

    parsed_list = [c.to_dict() for c in candidates]
    excel_bytes = generate_excel(parsed_list)

    filename = f"exported_candidates_{int(time.time())}.xlsx"
    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.get("/api/stats")
def get_stats(db: Session = Depends(get_db)):
    """Returns database summary metrics and role distribution."""
    total_candidates = db.query(func.count(Resume.id)).scalar() or 0

    # Role distribution
    roles_query = (
        db.query(Resume.role, func.count(Resume.id))
        .group_by(Resume.role)
        .order_by(func.count(Resume.id).desc())
        .all()
    )
    roles_distribution = {role: count for role, count in roles_query if role}

    # Unique roles list
    all_roles = [r[0] for r in roles_query if r[0]]

    return {
        "total_candidates": total_candidates,
        "roles_distribution": roles_distribution,
        "all_roles": all_roles,
    }


if __name__ == "__main__":
    init_db()
    uvicorn.run(app, host="0.0.0.0", port=8000)
