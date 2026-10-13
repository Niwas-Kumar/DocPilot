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
from . import updater, version
from .categories import CATEGORIES, get_category
from .jobs import manager, Job, WORK

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FRONTEND_DIST = os.path.join(ROOT, "frontend", "dist")

app = FastAPI(title="DocPilot", version=version.VERSION)


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
                     cover_index2: str = Form(None),
                     ts_series: str = Form(None),
                     ts_testno: str = Form(None),
                     ts_type: str = Form(None),
                     ts_subject: str = Form(None),
                     ts_code: str = Form(None),
                     ts_marks: str = Form(None),
                     ts_permark: str = Form(None),
                     ts_time: str = Form(None),
                     ts_year: str = Form(None)):
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

    # Test Series details typed on the form; empty = file name / template
    details = {}
    if cat.get("test_fields"):
        typed = {"series": ts_series, "testno": ts_testno,
                 "type": ts_type, "subject": ts_subject,
                 "code": ts_code, "marks": ts_marks, "permark": ts_permark,
                 "time": ts_time,
                 "year": ts_year}
        details = {k: " ".join((v or "").split()) for k, v in typed.items()
                   if k in cat["test_fields"] and (v or "").strip()}
        for key, label in (("marks", "Total marks"), ("time", "Time")):
            if key in details:
                as_page(details[key], label)       # whole number, 1 or more
                details[key] = str(int(details[key]))
        if "permark" in details:
            # whole or half marks: 2, 1, 2.5
            if not re.fullmatch(r"\d{1,2}(\.\d)?", details["permark"]) \
                    or float(details["permark"]) == 0:
                raise HTTPException(400, "Marks per question must be a number "
                                         "like 2 or 2.5")
            details["permark"] = f"{float(details['permark']):g}"
        if "testno" in details:
            if not re.fullmatch(r"\d{1,3}", details["testno"]):
                raise HTTPException(400, "Test number must be a number like 08")
            details["testno"] = details["testno"].zfill(2)
        if "year" in details and not re.fullmatch(r"20\d\d", details["year"]):
            raise HTTPException(400, "Target year must be a year like 2027")
        if details.get("type") not in (None, "Sectional", "Full Length"):
            raise HTTPException(400, "Test series type must be Sectional or "
                                     "Full Length")
        if "code" in details:
            details["code"] = details["code"].upper()

    if file is not None:
        job = Job(category_id, "upload", "", "", start_page=sp, end_page=ep,
                  cover=cover, details=details)
        job_dir = os.path.join(WORK, job.id)
        name, path = _save_upload(job_dir, file, await file.read())
        job.display_name, job.source = name, path
        if needs_two and file2 is not None:
            name2, path2 = _save_upload(job_dir, file2, await file2.read())
            job.display_name2, job.source2 = name2, path2
    else:
        job = Job(category_id, "link", link.strip(), "link",
                  start_page=sp, end_page=ep, cover=cover, details=details)
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


# ------------------------------------------------------------------ updates
@app.get("/api/update")
def update_check():
    """Current version, and whether the Drive manifest offers a newer one."""
    return updater.check()


@app.get("/api/update/status")
def update_status():
    return updater.status()


@app.post("/api/update/download")
def update_download():
    try:
        updater.start_download()
    except updater.UpdateError as e:
        raise HTTPException(400, str(e))
    return updater.status()


@app.post("/api/update/install")
def update_install():
    busy = [j for j in manager.list() if j["status"] in ("queued", "running")]
    if busy:
        raise HTTPException(409, f"{len(busy)} job(s) still running or queued - "
                                 "let them finish (or stop them), then install.")
    try:
        updater.install()
    except updater.UpdateError as e:
        raise HTTPException(400, str(e))
    return {"state": "installing"}


@app.exception_handler(Exception)
async def unhandled(request, exc):
    return JSONResponse(status_code=500, content={"detail": str(exc)})


@app.middleware("http")
async def fresh_page(request, call_next):
    """The page itself is never cached: a browser holding yesterday's
    index.html keeps loading the old form (its scripts are hash-named, so
    they need no such rule) and silently drops the fields added since."""
    response = await call_next(request)
    if request.url.path in ("/", "/index.html"):
        response.headers["Cache-Control"] = "no-cache"
    return response


if os.path.isdir(FRONTEND_DIST):
    app.mount("/", StaticFiles(directory=FRONTEND_DIST, html=True), name="static")
