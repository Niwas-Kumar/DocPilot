# -*- coding: utf-8 -*-
"""In-memory job queue with a single worker (Word COM must be serialized)."""
import os
import shutil
import threading
import time
import traceback
import uuid
from queue import Queue, Empty

from .categories import get_category
from .drive import download_docx, DriveError

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "work", "jobs")


def _pipeline(name):
    """Imported lazily so one broken pipeline cannot take the server down."""
    from .pipeline import daily_ca
    if name == "daily_ca":
        return daily_ca
    if name == "daily_ca_hindi":
        from .pipeline import daily_ca_hindi
        return daily_ca_hindi
    if name == "daily_ca_xylem":
        from .pipeline import daily_ca_xylem
        return daily_ca_xylem
    if name == "prelims_ab":
        from .pipeline import prelims_ab
        return prelims_ab
    if name == "prelims_qb":
        from .pipeline import prelims_qb
        return prelims_qb
    if name == "mains_ab":
        from .pipeline import mains_ab
        return mains_ab
    if name == "dea_english":
        from .pipeline import dea_english
        return dea_english
    if name == "dea_hindi":
        from .pipeline import dea_hindi
        return dea_hindi
    if name == "combo_english":
        from .pipeline import combo_english
        return combo_english
    if name == "combo_hindi":
        from .pipeline import combo_hindi
        return combo_hindi
    raise ValueError(f"Unknown pipeline: {name}")


class Job:
    def __init__(self, category_id, source_kind, source, display_name,
                 start_page=None, end_page=None, source2=None, display_name2=None,
                 cover=None):
        self.id = uuid.uuid4().hex[:12]
        self.category_id = category_id
        self.source_kind = source_kind        # 'upload' | 'link'
        self.source = source                  # local path | url
        self.source2 = source2                # second file (bilingual)
        self.display_name = display_name
        self.display_name2 = display_name2
        self.start_page = start_page
        self.end_page = end_page
        self.cover = cover or {}              # {'date','index','shorts','qr','index2'} strings
        self.status = "queued"                # queued|running|done|error|stopped
        self.stage = "Waiting in queue"
        self.progress = 0
        self.error = None
        self.notes = []
        self.output_path = None
        self.created = time.time()
        self.started = None
        self.finished = None
        self._cancel = threading.Event()

    def cancel(self):
        self._cancel.set()

    @property
    def cancelled(self):
        return self._cancel.is_set()

    def companion(self, ext):
        """The job's other deliverable: pipelines save both formats beside
        each other under the same stem (PDF categories add a .docx, the Test
        Series ones add a .pdf). Returns None when it is the main output or
        was not produced."""
        if not self.output_path:
            return None
        ext = ext if ext.startswith(".") else "." + ext
        if self.output_path.lower().endswith(ext):
            return None
        candidate = os.path.splitext(self.output_path)[0] + ext
        return candidate if os.path.exists(candidate) else None

    @property
    def companion_docx(self):
        return self.companion(".docx")

    @property
    def companion_pdf(self):
        return self.companion(".pdf")

    def to_dict(self):
        return {
            "id": self.id,
            "category_id": self.category_id,
            "source_kind": self.source_kind,
            "display_name": self.display_name,
            "status": self.status,
            "stage": self.stage,
            "progress": self.progress,
            "error": self.error,
            "notes": self.notes,
            "output_name": os.path.basename(self.output_path) if self.output_path else None,
            "docx_name": os.path.basename(self.companion_docx) if self.companion_docx else None,
            "pdf_name": os.path.basename(self.companion_pdf) if self.companion_pdf else None,
            "display_name2": self.display_name2,
            "start_page": self.start_page,
            "end_page": self.end_page,
            "created": self.created,
            "started": self.started,
            "finished": self.finished,
        }


class JobManager:
    def __init__(self):
        self.jobs: dict[str, Job] = {}
        self.order: list[str] = []
        self.queue: Queue[str] = Queue()
        self.lock = threading.Lock()
        self.worker = threading.Thread(target=self._run, daemon=True)
        self.worker.start()

    def submit(self, job: Job):
        with self.lock:
            self.jobs[job.id] = job
            self.order.append(job.id)
        self.queue.put(job.id)
        return job

    def stop(self, job_id):
        job = self.jobs.get(job_id)
        if job and job.status in ("queued", "running"):
            job.cancel()
            if job.status == "queued":
                job.status = "stopped"
                job.stage = "Stopped before start"
            return True
        return False

    def remove(self, job_id):
        job = self.jobs.get(job_id)
        if not job:
            return False
        if job.status == "running":
            return False
        with self.lock:
            self.jobs.pop(job_id, None)
            if job_id in self.order:
                self.order.remove(job_id)
        shutil.rmtree(os.path.join(WORK, job_id), ignore_errors=True)
        return True

    def list(self):
        with self.lock:
            return [self.jobs[i].to_dict() for i in self.order if i in self.jobs]

    # ------------------------------------------------------------ worker

    def _run(self):
        while True:
            try:
                jid = self.queue.get(timeout=1)
            except Empty:
                continue
            job = self.jobs.get(jid)
            if job is None or job.status != "queued":
                continue
            if job.cancelled:
                job.status = "stopped"
                continue
            self._process(job)

    def _process(self, job: Job):
        job.status = "running"
        job.started = time.time()
        workdir = os.path.join(WORK, job.id)
        os.makedirs(workdir, exist_ok=True)

        def progress(stage, pct):
            job.stage = stage
            job.progress = pct
            if stage.startswith(("Cover page added", "No cover page",
                                 "Cover added to the PDF only")):
                if stage not in job.notes:
                    job.notes.append(stage)

        try:
            cat = get_category(job.category_id)
            if cat is None:
                raise ValueError("Unknown category")
            if not cat["ready"]:
                raise ValueError("This category's template is not configured yet - "
                                 "share a sample input & output pair to enable it.")

            # 1) acquire input(s)
            if job.source_kind == "link":
                progress("Downloading from Google Drive", 8)
                input_path = download_docx(job.source, workdir, cancelled=lambda: job.cancelled)
                if job.display_name in (None, "", "link"):
                    job.display_name = os.path.basename(input_path)
            else:
                input_path = job.source

            second_path = None
            if cat.get("files", 1) == 2:
                if not job.source2:
                    slots = cat.get("slots") or ["English paper", "Hindi paper"]
                    raise ValueError(f"This category needs two files "
                                     f"({slots[0]} and {slots[1]}).")
                if job.source_kind == "link":
                    progress("Downloading second file", 12)
                    second_path = download_docx(job.source2, workdir,
                                                cancelled=lambda: job.cancelled)
                else:
                    second_path = job.source2
            if job.cancelled:
                raise InterruptedError("stopped")

            # 2) run the category pipeline.
            # Daily CA (English) is frozen and its process() takes no page or
            # second-file arguments, so only pass what each pipeline accepts.
            mod = _pipeline(cat["pipeline"])
            kwargs = {
                "variant": cat["variant"],
                "progress": progress,
                "cancelled": lambda: job.cancelled,
            }
            extras = {
                "start_page": job.start_page,
                "end_page": job.end_page,
                "second_path": second_path,
                "cover_date": job.cover.get("date"),
                "cover_index": job.cover.get("index"),
                "cover_shorts": job.cover.get("shorts"),
                "cover_qr": job.cover.get("qr"),
                "cover_index2": job.cover.get("index2"),
            }
            import inspect
            sig = inspect.signature(mod.process)
            takes_kwargs = any(p.kind is inspect.Parameter.VAR_KEYWORD
                               for p in sig.parameters.values())
            for key, value in extras.items():
                if takes_kwargs or key in sig.parameters:
                    kwargs[key] = value

            result = mod.process(input_path, workdir, **kwargs)
            # Daily CA (English) returns just a path; the newer ones add notes
            if isinstance(result, tuple):
                out, notes = result
                job.notes = list(notes or [])
            else:
                out = result
            job.output_path = out
            job.status = "done"
            job.stage = "Completed"
            job.progress = 100
        except InterruptedError:
            job.status = "stopped"
            job.stage = "Stopped"
        except DriveError as e:
            job.status = "error"
            job.error = str(e)
            job.stage = "Failed"
        except Exception as e:
            traceback.print_exc()
            job.status = "error"
            job.error = f"{type(e).__name__}: {e}"
            job.stage = "Failed"
        finally:
            job.finished = time.time()


manager = JobManager()
