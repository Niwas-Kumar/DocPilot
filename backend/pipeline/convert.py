# -*- coding: utf-8 -*-
"""DOCX -> PDF via Microsoft Word COM automation (single-flight)."""
import os
import threading
import time

_word_lock = threading.Lock()

WD_EXPORT_PDF = 17
WD_DO_NOT_SAVE = 0

ATTEMPTS = 3            # a transient COM failure must not discard a whole job
RETRY_PAUSE = 2.0       # seconds; Word needs a moment after a bad exit


def _kill_orphan_word():
    """Drop a Word process left behind by a failed attempt.

    A half-dead WINWORD holds the automation server in a state where every
    following DispatchEx returns an instance that refuses to export, so the
    retry would fail for the same reason as the first try. Best effort only -
    if the user has Word open for their own work this simply does nothing
    useful, and the retry still gets its chance.
    """
    try:
        import subprocess
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"],
                       capture_output=True, timeout=15,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception:
        pass


def _export_once(docx_path, pdf_path):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False)
        doc.ExportAsFixedFormat(
            OutputFileName=pdf_path,
            ExportFormat=WD_EXPORT_PDF,
            OpenAfterExport=False,
            OptimizeFor=0,           # print quality
            CreateBookmarks=0,
            DocStructureTags=False,
            BitmapMissingFonts=True,
        )
    finally:
        try:
            if doc is not None:
                doc.Close(WD_DO_NOT_SAVE)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def docx_to_pdf(docx_path: str, pdf_path: str) -> str:
    """Convert with Word. Serialized: Word COM is not concurrency-safe.

    Retried, because Word fails transiently. Every pipeline here renders many
    times per job - the Prelims booklet does it dozens of times - so a single
    flaky export used to throw away a quarter of an hour of work with
    "Word did not produce the PDF output". The document is almost never the
    problem: the same file converts on the next attempt. A stale PDF from a
    half-finished export is removed first so a partial file can never be
    mistaken for success.
    """
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path)
    last = None
    with _word_lock:
        for attempt in range(1, ATTEMPTS + 1):
            try:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
            except OSError:
                pass
            try:
                _export_once(docx_path, pdf_path)
                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                    return pdf_path
                last = "Word reported success but wrote no PDF"
            except Exception as exc:
                last = f"{type(exc).__name__}: {exc}"
            if attempt < ATTEMPTS:
                _kill_orphan_word()
                time.sleep(RETRY_PAUSE * attempt)
    raise RuntimeError(
        f"Word could not convert {os.path.basename(docx_path)} after "
        f"{ATTEMPTS} attempts ({last}). Close any open Word windows and run "
        f"the job again; if it keeps happening, send this file.")
