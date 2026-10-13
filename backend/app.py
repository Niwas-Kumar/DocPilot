# -*- coding: utf-8 -*-
"""DocPilot API + static frontend server.

Run:  python -m uvicorn backend.app:app --port 8930
or:   run.bat
"""
import os
import re

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from . import fonts as _fonts
from .categories import CATEGORIES, get_category
from .jobs import manager, Job, WORK

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FRONTEND_DIST = os.path.join(ROOT, "frontend", "dist")

app = FastAPI(title="DocPilot", version="2.0")


@app.on_event("startup")
def _install_bundled_fonts():
    """Put Baloo 2 / Rozha One on this machine before the first job runs.

    A font already installed here - by IT machine-wide or by the user - is
    left untouched. A failure is logged and ignored: the pipelines fall back
    to Nirmala UI and report it, which is worse-looking but never fatal.
    """
    try:
        print(_fonts.summary(), flush=True)
    except Exception as exc:                       # never block startup
        print(f"Bundled fonts: skipped ({type(exc).__name__}: {exc}).", flush=True)


@app.get("/api/categories")
def categories():
    return CATEGORIES


def _save_upload(job_dir, upload, data):
    name = os.path.basename(upload.filename or "input.docx")
    if not name.lower().endswith(".docx"):
        raise HTTPException(400, "Only .docx files are supported (export Google "
                                 "Docs as .docx or paste its link)")
    os.makedirs(job_dir, exist_ok=True)
    safe = re.sub(r'[<>:"/\\|?*]', "_", name)
    path = os.path.join(job_dir, safe)
    with open(path, "wb") as f:
        f.write(data)
    return name, path


@app.post("/api/jobs")
async def create_job(category_id: str = Form(...),
                     link: str = Form(None),
                     link2: str = Form(None),
                     start_page: str = Form(None),
                     end_page: str = Form(None),
                     file: UploadFile = File(None),
                     file2: UploadFile = File(None),
                     cover_date: str = Form(None),
                     cover_index: str = Form(None),
                     cover_shorts: str = Form(None),
                     cover_qr: str = Form(None),
                     cover_index2: str = Form(None)):
    cat = get_category(category_id)
    if cat is None:
        raise HTTPException(400, "Unknown category")
    if not cat["ready"]:
        raise HTTPException(400, "This category's template is not configured yet. "
                                 "Share a sample input & output pair to enable it.")
    if file is None and not link:
        raise HTTPException(400, "Provide a file upload or a Google Drive link")

    needs_two = cat.get("files", 1) == 2
    if needs_two and file2 is None and not link2:
        slots = cat.get("slots") or ["English paper", "Hindi paper"]
        raise HTTPException(400, f"This category needs two files: the {slots[0]} "
                                 f"and the {slots[1]}.")

    def as_page(value, label):
        if value in (None, "", "null"):
            return None
        try:
            n = int(str(value).strip())
        except ValueError:
            raise HTTPException(400, f"{label} must be a whole number")
        if n < 1:
            raise HTTPException(400, f"{label} must be 1 or greater")
        return n

    sp = as_page(start_page, "Start page")
    ep = as_page(end_page, "End page")
    if cat.get("pages") == "none":
        sp = ep = None
    if sp and ep and ep < sp:
        raise HTTPException(400, "End page cannot be before the start page")

    # front-page fields (Daily CA English only); empty = extract from the file
    cover = None
    if cat.get("cover"):
        cover = {"date": (cover_date or "").strip(),
                 "index": (cover_index or "").strip(),
                 "shorts": (cover_shorts or "").strip(),
                 "qr": (cover_qr or "").strip(),
                 "index2": (cover_index2 or "").strip()}

    if file is not None:
        job = Job(category_id, "upload", "", "", start_page=sp, end_page=ep,
                  cover=cover)
        job_dir = os.path.join(WORK, job.id)
        name, path = _save_upload(job_dir, file, await file.read())
        job.display_name, job.source = name, path
        if needs_two and file2 is not None:
            name2, path2 = _save_upload(job_dir, file2, await file2.read())
            job.display_name2, job.source2 = name2, path2
    else:
        job = Job(category_id, "link", link.strip(), "link",
                  start_page=sp, end_page=ep, cover=cover)
        if needs_two and link2:
            job.source2 = link2.strip()
            job.display_name2 = "link"

    manager.submit(job)
    return job.to_dict()


@app.get("/api/jobs")
def list_jobs():
    return manager.list()


@app.post("/api/jobs/{job_id}/stop")
def stop_job(job_id: str):
    if not manager.stop(job_id):
        raise HTTPException(404, "Job not found or not stoppable")
    return {"ok": True}


@app.delete("/api/jobs/{job_id}")
def remove_job(job_id: str):
    if not manager.remove(job_id):
        raise HTTPException(400, "Job not found or still running (stop it first)")
    return {"ok": True}


@app.get("/api/jobs/{job_id}/download")
def download(job_id: str, fmt: str = ""):
    """fmt='docx'/'pdf' serves the companion format when the job made both."""
    job = manager.jobs.get(job_id)
    if job is None or job.status != "done" or not job.output_path:
        raise HTTPException(404, "Output not available")
    path = job.output_path
    fmt = fmt.lower().lstrip(".")
    if fmt in ("docx", "pdf") and not path.lower().endswith("." + fmt):
        path = job.companion("." + fmt)
        if not path:
            raise HTTPException(404, f"No {fmt.upper()} file for this job")
    name = os.path.basename(path)
    media = ("application/vnd.openxmlformats-officedocument.wordprocessingml.document"
             if name.lower().endswith(".docx") else "application/pdf")
    return FileResponse(path, media_type=media, filename=name)


@app.exception_handler(Exception)
async def unhandled(request, exc):
    return JSONResponse(status_code=500, content={"detail": str(exc)})


if os.path.isdir(FRONTEND_DIST):
    app.mount("/", StaticFiles(directory=FRONTEND_DIST, html=True), name="static")
