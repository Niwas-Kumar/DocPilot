# -*- coding: utf-8 -*-
"""Daily Current Affairs + Daily Editorial Analysis (English) - one booklet.

Specification: ``Output_DCA+DEA english 17th August 2026_CA Analysis+Daily
Editorial Summary_UPSC_English_.pdf`` (22 pages). Its structure, measured:

    p1       combined cover - DAILY NEWS ANALYSIS, date, INDEX box of the CA
             topics, DAILY EDITORIAL SUMMARY, a box of the editorial topics
    p2-15    the Daily CA booklet exactly as that category makes it, closing
             promo page included (p15)
    p16-22   the Daily Editorial booklet without its cover: Topics banner,
             topics, PYQ banner in the space left, SRIJAN closing page

So each half is made by its OWN pipeline's stages, untouched: the CA half by
styler -> fitter -> postprocess (the same three calls daily_ca.process
makes, minus its cover), the editorial half by dea_english.build_body.

One deliberate difference from the sample: its editorial half restarts at
page 2 (the two PDFs were joined by hand). Here numbering runs on across the
join, so no two pages of one booklet carry the same number.
"""
import os
import shutil

import fitz

from . import cover as CA_COVER
from . import cover_kit as K
from . import dea_common as C
from . import dea_english as E
from .convert import docx_to_pdf
from .fitter import fit_document
from .postprocess import finalize_pdf
from .styler import BLUE, restyle_docx
from .template import LINK_TELEGRAM, LINK_WHATSAPP

ASSETS = C.ASSETS
PLATE = os.path.join(ASSETS, "en_combo_cover_plate.jpg")
PILL = (160.4, 213.2, 297.1, 47.8)
BOX_X, BOX_W = 53.5, 504.2
TICK_INSET, TEXT_HANG, TOP_INSET = 8.2, 27.8, 5.2    # ticks 61.7pt, text 90.0pt as published
CA_TOP, CA_FLOOR = 323.5, 511.0          # the editorial title starts at 519.8
DEA_TOP, DEA_FLOOR = 581.3, 700.0        # the buttons start at 705.5
# (size, pitch) - the published boxes are Times Bold 15pt on 24.7pt
STEPS = [(15.0, 24.7), (14.0, 23.0), (13.0, 21.4), (12.0, 19.8), (11.0, 18.1), (10.0, 16.5)]
BTN_W = (56.0, 707.0, 130.0, 40.0)
BTN_T = (427.5, 705.5, 133.0, 41.0)


def output_name(stem, variant="English"):
    return f"{stem}_CA Analysis+Daily Editorial Summary_UPSC_{variant}.pdf"


# ------------------------------------------------------------------ cover

RIGHT_INSET = 0.0
TNR_LINE = 1.15          # Times New Roman single line = 1.15 x its size


def _box_height(lines, pitch):
    # the last entry's space-after is the box's bottom padding, as published
    # (4 entries: 5.2 + 4 x 24.7 = 104.0 against the sample's 103.8)
    return TOP_INSET + lines * pitch


def _font():
    return fitz.Font(fontfile=E.TNR_BOLD) if os.path.exists(E.TNR_BOLD) else fitz.Font("tibo")


TEXT_W = BOX_W - TICK_INSET - TEXT_HANG - RIGHT_INSET


def _lines_for(t, size, font):
    """Lines an entry takes: one if a slight condense keeps it on one line."""
    if K.fit_spacing(t, size, TEXT_W, font):
        return 1
    return E._wrap_lines(t, size, TEXT_W, font)


def _fit_box(entries, top, floor):
    """Largest step whose box, grown to its content, still ends above floor."""
    font = _font()
    for size, pitch in STEPS:
        n = sum(_lines_for(t, size, font) for t in entries)
        if top + _box_height(n, pitch) <= floor:
            return size, pitch, n
    size, pitch = STEPS[-1]
    return size, pitch, sum(_lines_for(t, size, font) for t in entries)


def _tick_box(cv, entries, top, size, pitch, height, framed, name):
    # single spacing + space after, NOT exact spacing: with an exact line
    # Word puts the extra leading ABOVE the glyphs (6.8pt low here), and a
    # title that wraps should read as one entry, not as two
    after = max(0.0, pitch - TNR_LINE * size)
    font = _font()
    paras = [K.para([K.sym("Wingdings", "F0FC", size, "000000"), K.tab(),
                     K.run(t, "Times New Roman", size, "000000", bold=True,
                           spacing_pt=K.fit_spacing(t, size, TEXT_W, font))],
                    line=None, after=after, left=TEXT_HANG, hanging=TEXT_HANG,
                    tabs=[("left", TEXT_HANG, None)]) for t in entries]
    cv.textbox((BOX_X, top, BOX_W, height), paras,
               fill="FFFFFF" if framed else None, line=(1.0, BLUE) if framed else None,
               insets=(TICK_INSET, TOP_INSET, RIGHT_INSET, 0.0), name=name)


def build_cover(path, date, boxes, framed):
    """`boxes` = [(entries, top, size, pitch, height, name)]."""
    cv = K.Cover()
    cv.background(PLATE)
    if date:
        runs = [K.run(t, "Arial", 32, "FFFFFF", bold=True, sup=sup)
                for t, sup in CA_COVER.date_runs(date)]
        cv.textbox(PILL, [K.para(runs, jc="center")], anchor="ctr", name="Date")
    for entries, top, size, pitch, height, name in boxes:
        if entries:
            _tick_box(cv, entries, top, size, pitch, height, framed, name)
    cv.picture(os.path.join(ASSETS, "en_combo_btn_whatsapp.png"), BTN_W, link=LINK_WHATSAPP,
               name="Join WhatsApp")
    cv.picture(os.path.join(ASSETS, "en_combo_btn_telegram.png"), BTN_T, link=LINK_TELEGRAM,
               name="Join Telegram")
    return cv.save(path)


def _measure(pdf_path, entries, top, floor):
    """(every entry present, bottom of the last line) inside one box region."""
    lines = K.content_lines(pdf_path, (BOX_X - 2, top, BOX_X + BOX_W, floor + 200))
    ticks = [l for l in lines if "wingdings" in l[4].lower() and l[2] < BOX_X + TICK_INSET + 6]
    return len(ticks) >= len(entries), max((l[1] for l in lines), default=top)


def render_cover(workdir, date, ca_topics, dea_topics, progress=lambda s, p: None):
    """Measure, then frame.

    Pass 1 renders each list ON ITS OWN in a tall, unframed box - nothing
    can be clipped, and no other list shares the page to be mistaken for its
    overflow - and reads where its text really ends; a list that runs past
    its floor steps down a size. Pass 2 draws every frame to its measured
    height (plus the published bottom padding). So Word never gets the chance
    to hide a line inside a frame that was sized from an estimate."""
    cover_docx = os.path.join(workdir, "_cover_combo.docx")
    cover_pdf = os.path.join(workdir, "_cover_combo.pdf")
    probe_docx = os.path.join(workdir, "_cover_combo_probe.docx")
    probe_pdf = os.path.join(workdir, "_cover_combo_probe.pdf")
    specs = [["CA index", ca_topics, CA_TOP, CA_FLOOR], ["Editorial index", dea_topics, DEA_TOP, DEA_FLOOR]]
    chosen, bottoms, fits = {}, {}, True
    for name, entries, top, floor in specs:
        if not entries:
            continue
        size, _pitch, _n = _fit_box(entries, top, floor)
        i = next(k for k, st in enumerate(STEPS) if st[0] == size)
        while True:
            progress("Building cover page (measuring the index)", 89)
            size, pitch = STEPS[i]
            build_cover(probe_docx, "", [(entries, top, size, pitch, 792 - top, name)], framed=False)
            docx_to_pdf(probe_docx, probe_pdf)
            present, bottom = _measure(probe_pdf, entries, top, 792)
            pad = pitch - TNR_LINE * size
            if present and bottom + pad <= floor:
                break
            if i == len(STEPS) - 1:
                fits = False
                break
            i += 1
        chosen[name], bottoms[name] = i, bottom
    boxes = []
    for name, entries, top, floor in specs:
        if not entries:
            continue
        size, pitch = STEPS[chosen[name]]
        pad = pitch - TNR_LINE * size
        boxes.append((entries, top, size, pitch, (bottoms[name] - top) + pad + 0.8, name))
    progress("Building cover page", 90)
    build_cover(cover_docx, date, boxes, framed=True)
    docx_to_pdf(cover_docx, cover_pdf)
    for name, entries, top, floor in specs:              # the frame lost nothing
        if entries and not _measure(cover_pdf, entries, top, floor + 4)[0]:
            fits = False
    step = lambda n: STEPS[chosen.get(n, 0)]
    return cover_pdf, cover_docx, step("CA index"), step("Editorial index"), fits


# ------------------------------------------------------------------ pipeline

def process(input_path, workdir, variant="English", progress=lambda s, p: None,
            cancelled=lambda: False, second_path=None, cover_date=None, cover_index=None,
            cover_index2=None, **kw):
    """input_path = the Daily CA raw file, second_path = the Daily Editorial one."""
    if not second_path:
        raise ValueError("This category needs two files: the Daily Current Affairs "
                         "file and the Daily Editorial Analysis file.")
    os.makedirs(workdir, exist_ok=True)
    notes = []
    input_path, second_path = C.order_pair(input_path, second_path, notes, hindi=False)
    stem = os.path.splitext(os.path.basename(input_path))[0]

    # ---- 1) the Daily CA half: its own three stages, nothing changed
    progress("Styling Daily CA", 10)
    styled_ca = os.path.join(workdir, f"{stem}_styled.docx")
    restyle_docx(input_path, styled_ca, banner_text="Topics")
    if cancelled():
        raise InterruptedError("stopped")
    progress("Fitting Daily CA pages", 20)
    raw_ca = fit_document(styled_ca, workdir, progress=lambda s, p: progress(s + " (Daily CA)", 25))
    if not os.path.exists(raw_ca):
        raw_ca = os.path.join(workdir, f"{stem}_raw.pdf")
        docx_to_pdf(styled_ca, raw_ca)
    if cancelled():
        raise InterruptedError("stopped")
    ca_part = os.path.join(workdir, f"{stem}_part_ca.pdf")
    finalize_pdf(raw_ca, ca_part, first_number=2)
    with fitz.open(ca_part) as d:
        ca_pages = len(d)                       # body + its closing promo page

    # ---- 2) the Daily Editorial half, numbered on from the CA half
    styled_dea, dea_part, info, _ = E.build_body(
        second_path, workdir, 2 + ca_pages, notes,
        progress=lambda s, p: progress(s, 55), cancelled=cancelled)
    if cancelled():
        raise InterruptedError("stopped")

    # ---- 3) one booklet
    progress("Joining the two halves", 80)
    final_pdf = os.path.join(workdir, output_name(stem, variant))
    out = fitz.open()
    for part in (ca_part, dea_part):
        with fitz.open(part) as d:
            out.insert_pdf(d)
    out.save(final_pdf, garbage=3, deflate=True)
    out.close()
    companion = os.path.splitext(final_pdf)[0] + ".docx"
    try:
        C.merge_docx(styled_ca, styled_dea, companion)
    except Exception as exc:
        notes.append(f"Word companion not produced ({type(exc).__name__}: {exc}); the PDF is complete.")
        companion = None
    notes.append(f"Daily CA: pages 2-{1 + ca_pages} (its closing promo page included); "
                 f"Daily Editorial: pages {2 + ca_pages}-{1 + ca_pages + _pages(dea_part)}.")

    # ---- 4) the combined cover
    progress("Building cover page", 88)
    try:
        from .cover import _lines, prepend_docx, prepend_pdf
        ca_date, ca_index, _shorts = CA_COVER.extract_from_raw(input_path)
        date = (cover_date or "").strip() or ca_date
        if not date:
            dated = sorted([d for d in info["days"] if d.date], key=lambda d: d.date)
            date = C.english_date(dated[-1].date) if dated else ""
        ca_topics = _lines(cover_index) or ca_index
        dea_topics = _lines(cover_index2) or info["topics"]
        cover_pdf, cover_docx, ca_step, dea_step, fits = render_cover(
            workdir, date, ca_topics, dea_topics, progress)
        prepend_pdf(cover_pdf, final_pdf, final_pdf)
        if companion:
            try:
                prepend_docx(cover_docx, companion, companion)
            except Exception as exc:
                notes.append(f"Cover added to the PDF only ({type(exc).__name__}).")
        bits = [("date typed" if (cover_date or "").strip() else
                 ("date from the file" if date else "date MISSING - type it in the Date block")),
                f"{len(ca_topics)} Daily CA topic(s) at {ca_step[0]:g}pt",
                f"{len(dea_topics)} editorial topic(s) at {dea_step[0]:g}pt"]
        if not fits:
            bits.append("an index is still too long for its box - shorten it in the Index block")
        notes.append("Cover page added: " + ", ".join(bits) + ".")
        video = sorted({l for d in info["days"] for l, _b in d.qr})
        if video:
            notes.append("The editorial file's video QR (" + ", ".join(video) + ") is not "
                         "printed: the combined edition's cover carries none, as published.")
    except InterruptedError:
        raise
    except Exception as exc:
        notes.append(f"No cover page ({type(exc).__name__}: {exc}); the booklet is unchanged.")
    progress("Done", 100)
    return final_pdf, notes


def _pages(pdf):
    with fitz.open(pdf) as d:
        return len(d)
