# -*- coding: utf-8 -*-
"""Daily Current Affairs Summary pipeline (English / Hindi variants)."""
import os
import shutil
from .styler import restyle_docx
from .convert import docx_to_pdf
from .fitter import fit_document
from .postprocess import finalize_pdf
from .cover import attach_cover


def output_name(input_stem: str, variant: str) -> str:
    return f"{input_stem}_CA Analysis_UPSC_{variant}.pdf"


def process(input_docx: str, workdir: str, variant: str = "English",
            progress=lambda stage, pct: None, cancelled=lambda: False,
            cover_date=None, cover_index=None, cover_shorts=None, **kw):
    """Run the full transformation. Returns the final PDF path (inside workdir).

    `cover_*` are the optional front-page fields (Daily CA English only): an
    empty one is extracted from the raw file, a filled one is used verbatim.
    """
    os.makedirs(workdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(input_docx))[0]

    progress("Styling document", 25)
    styled = os.path.join(workdir, f"{stem}_styled.docx")
    restyle_docx(input_docx, styled, banner_text="Topics")
    if cancelled():
        raise InterruptedError("stopped")

    # measured pagination: fill page bottoms, pull small topic tails back,
    # keep topic heads with their bodies (client mark-up, 2026-08-26).
    # The pass converts through Word as it measures, so its last export
    # already matches the final docx state.
    progress("Fitting pages", 40)
    raw_pdf = fit_document(styled, workdir, progress=progress)
    if cancelled():
        raise InterruptedError("stopped")

    if not os.path.exists(raw_pdf):
        progress("Converting to PDF (Word)", 70)
        raw_pdf = os.path.join(workdir, f"{stem}_raw.pdf")
        docx_to_pdf(styled, raw_pdf)
    if cancelled():
        raise InterruptedError("stopped")

    progress("Branding pages & promo", 85)
    final_pdf = os.path.join(workdir, output_name(stem, variant))
    finalize_pdf(raw_pdf, final_pdf, first_number=2)

    # editable Word companion, named after the PDF so callers can find it by
    # convention (page numbers and the promo page are stamped into the PDF
    # only - the .docx is the branded, fitted source)
    companion = os.path.splitext(final_pdf)[0] + ".docx"
    try:
        shutil.copy2(styled, companion)
    except OSError:
        pass

    # front page (client request 2026-08-31): built from the official cover
    # template and prepended to BOTH deliverables. A separate stage on
    # purpose - everything above is the signed-off body pipeline.
    if variant == "English":
        progress("Building cover page", 88)
        notes = []
        attach_cover(input_docx, final_pdf, companion, workdir,
                     final_pdf, companion,
                     date_in=cover_date, index_in=cover_index,
                     shorts_in=cover_shorts, progress=progress, notes=notes)
        for n in notes:
            progress(n, 96)

    progress("Done", 100)
    return final_pdf
