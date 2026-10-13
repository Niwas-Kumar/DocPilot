# -*- coding: utf-8 -*-
"""Daily Editorial Analysis (Hindi) - "दैनिक संपादकीय सारांश" booklet.

Specification: ``Output 17 Aug 2026_Daily Editorial Summary_UPSC_Hindi.pdf``.

The body is the Hindi Daily CA design, measured: the red icon + peach 38.7pt
topic banner with its Baloo 2 13pt #CB202F title, the right-aligned GS line,
Adobe Devanagari 11.5pt body (Baloo 2 when that commercial face is absent),
red Wingdings dots on a 12pt hang, logo + centre watermark header, the three
brand buttons and the black page-number box. So the styling primitives are
daily_ca_hindi's own (imported, never edited); the reading of the raw file is
dea_common's (plain paragraphs, colour-coded titles, multi-day files).

What differs from Daily CA, all taken from the samples:

* topics flow; a new DAY starts a new page (the combined sample)
* a promo is set into the space a section leaves at its foot - the
  FOUNDATION COURSES banner after the last topic, the SRIJAN banner where an
  earlier day ends - scaled to the space but never below 70%
* the closing page is the published full-bleed "अन्य पुस्तकें एवं कार्यक्रम"
* the cover: title, date (from the file's date line, or typed), the video QR
  in its frame (from the file, or from a typed link), and the विषय सूची index
  with page numbers read from Word on the very document that made the PDF
"""
import os
import shutil

import fitz
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from . import daily_ca_hindi as H
from . import dea_common as C
from . import cover_kit as K
from .convert import docx_to_pdf

ASSETS = C.ASSETS
PAGE_FLOOR = 716.0                 # the footer buttons start at 722pt

FOUNDATION = os.path.join(ASSETS, "hi_foundation_banner.png")
FOUNDATION_W, FOUNDATION_CX = 540.0, 306.0
SRIJAN = os.path.join(ASSETS, "hi_srijan_banner.png")
SRIJAN_W, SRIJAN_CX = 334.8, 311.8
BANNER_MIN_SCALE = 0.70
BANNER_GAP = 14.0
BANNER_TOP = 80.0                 # below the logo / combined section badges
# the published closing page, flattened to a picture (see make_dea_assets)
CLOSING_PAGE = os.path.join(ASSETS, "hi_last_page.jpg")


# ================================================================== restyle

def _heading_builder(fonts):
    def make(text):
        p = parse_xml(f'<w:p {C.W_NS}><w:r><w:t xml:space="preserve"></w:t></w:r></w:p>')
        p.find('.//' + qn('w:t')).text = text
        H._style_heading(p, fonts)
        return p
    return make


def restyle(in_path, out_path, notes, fonts):
    """Raw DEA Hindi docx -> branded docx. Returns the info dict the cover
    and the page stage need (days in print order, topics, day starts)."""
    doc = Document(in_path)
    H._apply_page_template(doc)
    body = doc.element.body
    H._cleanup_source(body)
    tabs, breaks = H._drop_tab_artifacts(body, doc)
    if tabs or breaks:
        notes.append(f"Google Docs tab front matter removed ({tabs} line(s), {breaks} page break(s)).")

    days = C.split_days(doc, body)
    order = C.print_order(days)
    if len(days) > 1 and [id(d) for d in order] != [id(d) for d in days]:
        notes.append("The file holds " + ", ".join(d.date_text or "undated" for d in days)
                     + " - printed oldest first, each day from a new page.")
    C.reassemble(body, order)

    make_head = _heading_builder(fonts)
    topics, day_first_topic, day_first_block = [], [], []
    counts = {"headings": 0, "leadins": 0, "bullets": 0, "tables": 0, "empty": 0}
    for day in order:
        day_first_topic.append(len(topics))
        first_block = None
        blocks = day.blocks
        body_sz = C.body_size_of(doc, blocks)
        awaiting_gs = 0
        for i, el in enumerate(list(blocks)):
            if el.getparent() is None:
                continue
            if el.tag == qn('w:tbl'):
                H._style_table(el, doc, fonts)
                counts["tables"] += 1
                first_block = first_block if first_block is not None else el
                awaiting_gs = 0
                continue
            text = C.clean(C.para_text(el))
            if not text and not C.has_drawing(el):
                el.getparent().remove(el)
                counts["empty"] += 1
                continue
            nxt = next((b for b in blocks[i + 1:]
                        if b.tag != qn('w:p') or C.clean(C.para_text(b))), None)
            next_is_list = (nxt is not None and nxt.tag == qn('w:p')
                            and C.list_level(nxt, doc) is not None)
            level = C.list_level(el, doc)
            if C.is_topic_title(el, doc, body_sz):
                tbl = H._banner_table(el, doc, fonts)
                topics.append(text)
                first_block = first_block if first_block is not None else tbl
                awaiting_gs = 2
                continue
            first_block = first_block if first_block is not None else el
            if awaiting_gs:
                awaiting_gs -= 1
                if C.is_gs_line(text):
                    H._style_gs_line(el, fonts)
                    C.set_flag(el, 'w:keepNext', True)
                    awaiting_gs = 0
                    continue
            if level is None and C.strip_literal_bullet(el):
                level = 0
            if level is not None:
                H._style_list_item(el, level + 1, fonts)
                counts["bullets"] += 1
                continue
            if C.split_leadin(el, make_head) is not None:
                H._style_body(el, fonts)
                counts["leadins"] += 1
                continue
            if C.is_heading(el, doc, next_is_list):
                H._style_heading(el, fonts)
                counts["headings"] += 1
                continue
            H._style_body(el, fonts)
        day_first_block.append(first_block)

    floated = H._float_body_images(body, H.COL_W_PT * 0.40)
    for el in day_first_block[1:]:
        if el is not None:
            C.break_before(el)

    blocks = [e for e in body if e.tag in (qn('w:p'), qn('w:tbl'))]
    if blocks and blocks[-1].tag == qn('w:tbl'):
        blocks[-1].addnext(parse_xml(
            f'<w:p {C.W_NS}><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" '
            f'w:lineRule="exact"/></w:pPr></w:p>'))

    if not topics:
        notes.append("No topic titles found (a bold red or large line) - the file "
                     "may not be a Daily Editorial Analysis document.")
    notes.append(f"{len(topics)} editorial topic(s); {counts['headings']} heading(s), "
                 f"{counts['leadins']} संदर्भ/निष्कर्ष lead-in(s), {counts['bullets']} bullet(s)"
                 + (f", {counts['tables']} table(s)" if counts['tables'] else "")
                 + (f", {floated} picture(s) docked right" if floated else "")
                 + f"; {counts['empty']} empty paragraph(s) removed.")
    doc.save(out_path)
    return {"days": order, "topics": topics, "day_first_topic": day_first_topic}


# ================================================================== pages

def stamp_and_fill(doc, first_number, day_end_pages=(), end_banner=FOUNDATION,
                   end_banner_w=FOUNDATION_W, end_banner_cx=FOUNDATION_CX, notes=None,
                   label="last page"):
    """Frame + number every page; SRIJAN into the space where an earlier day
    ends; `end_banner` into the space after the last line. Returns the next
    page number."""
    notes = notes if notes is not None else []
    n = first_number
    for page in doc:
        H._stamp_page(page, n)
        n += 1
    for pno in sorted(set(day_end_pages)):
        if 0 <= pno < len(doc) - 1:
            page = doc[pno]
            C.place_banner(page, SRIJAN, C.content_bottom(page, PAGE_FLOOR), PAGE_FLOOR,
                           SRIJAN_W, SRIJAN_CX, BANNER_MIN_SCALE, BANNER_GAP,
                           min_top=BANNER_TOP)
    if end_banner:
        last = doc[-1]
        rect = C.place_banner(last, end_banner, C.content_bottom(last, PAGE_FLOOR), PAGE_FLOOR,
                              end_banner_w, end_banner_cx, BANNER_MIN_SCALE, BANNER_GAP,
                              min_top=BANNER_TOP)
        if rect is not None:
            scale = rect.width / end_banner_w
            notes.append(f"Promo banner set into the space left on the {label}"
                         + ("" if scale > 0.995 else f" (at {scale:.0%} to fit)") + ".")
        else:
            notes.append(f"No promo banner on the {label}: the text fills it.")
    return n


def add_closing_page(doc):
    page = doc.new_page(width=612, height=792)
    page.insert_image(page.rect, filename=CLOSING_PAGE)
    return page


def finalize(raw_pdf, out_pdf, first_number, day_end_pages=(), notes=None):
    doc = fitz.open(raw_pdf)
    n = stamp_and_fill(doc, first_number, day_end_pages, notes=notes)
    add_closing_page(doc)
    doc.save(out_pdf, garbage=3, deflate=True)
    doc.close()
    return n + 1


# ================================================================== cover

COVER_PLATE = os.path.join(ASSETS, "hi_dea_cover_plate.jpg")
INK, WHITE, NAVY = H.INK, "FFFFFF", "080B4F"
DATE_BOX = (235.3, 317.4, 141.4, 35.9)
QR_RECT = (234.4, 475.4, 143.2, 143.2)
NO_QR_PATCH = (178.0, 396.0, 256.0, 236.0)     # caption + frame, painted navy
INDEX_X, INDEX_R = 34.8, 575.8
INDEX_TOP, INDEX_FLOOR = 668.0, 782.0
INDEX_NUM_TAB = 13.5
# (size, pitch) - the published index is Bakbak One 11.5pt on 18.1pt
INDEX_STEPS = [(11.5, 18.1), (10.5, 16.6), (9.5, 15.0), (8.5, 13.5)]
DATE_CHAIN = ("Karma", "Nirmala UI")
INDEX_CHAIN = ("Bakbak One", "Nirmala UI")


def cover_fonts():
    from .cover_hindi import _installed_families, _pick
    avail = _installed_families()
    return {"date": _pick(DATE_CHAIN, avail), "index": _pick(INDEX_CHAIN, avail)}


def hindi_date_text(typed, days):
    """Typed wins (any spelling, via the Daily CA cover's parser); else the
    latest day's date line; else ''."""
    from .cover_hindi import MONTHS_HI, hindi_date
    if (typed or "").strip():
        return hindi_date(typed)[0], "typed"
    dated = sorted([d for d in days if d.date], key=lambda d: d.date)
    if dated:
        d = dated[-1].date
        return f"{d.day} {MONTHS_HI[d.month - 1]} {d.year}", "from the file"
    return "", "missing"


def index_paragraphs(entries, numbers, size, pitch, font, color, num_tab, right_x, left_x,
                     numbered=True, tick=None, after=0.0):
    """One index line per entry: 'N.' (or a tick), tab, title, dot-leader
    right tab, page number. Wrapped lines hang at the title."""
    paras = []
    width = right_x - left_x
    for i, title in enumerate(entries, 1):
        runs = []
        if tick:
            runs.append(K.sym(tick[0], tick[1], tick[2], tick[3]))
        elif numbered:
            runs.append(K.run(f"{i}.", font, size, color))
        runs.append(K.tab())
        runs.append(K.run(title, font, size, color))
        num = numbers.get(title)
        runs.append(K.tab())
        runs.append(K.run(str(num) if num else "", font, size, color))
        paras.append(K.para(runs, line=pitch, left=num_tab, hanging=num_tab, after=after,
                            tabs=[("left", num_tab, None), ("right", width, "dot")]))
    return paras


def build_cover(path, date, topics, numbers, qr_png, step, fonts):
    size, pitch = step
    cv = K.Cover()
    cv.background(COVER_PLATE)
    if date:
        cv.textbox(DATE_BOX, [K.para([K.run(date, fonts["date"], 21, INK)], jc="center")],
                   anchor="ctr", name="Date")
    if qr_png:
        cv.picture(qr_png, QR_RECT, name="Video QR")
    else:
        cv.textbox(NO_QR_PATCH, [], fill=NAVY, name="No video")
    paras = index_paragraphs(topics, numbers, size, pitch, fonts["index"], WHITE,
                             INDEX_NUM_TAB, INDEX_R, INDEX_X)
    # the box runs to the page foot: Word hides whatever overflows a text box,
    # so the frame is never the limit - the measured floor is
    cv.textbox((INDEX_X, INDEX_TOP, INDEX_R - INDEX_X, 792.0 - INDEX_TOP), paras, name="Index")
    return cv.save(path)


def _index_ok(pdf_path, topics, top, floor, x0, x1):
    import re
    lines = [l for l in K.lines_in(pdf_path, (x0 - 4, top - 2, x1 + 4, floor)) if l[0] < floor]
    starts = [l for l in lines if re.match(r"^\s*\d+\s*\.", l[3]) or "wingdings" in l[4].lower()]
    bottom = max((l[1] for l in lines), default=0.0)
    return len(starts) >= len(topics) and bottom <= floor + 1, bottom


def render_cover(workdir, date, topics, numbers, qr_png, fonts, progress=lambda s, p: None):
    cover_docx = os.path.join(workdir, "_cover_dea_hi.docx")
    cover_pdf = os.path.join(workdir, "_cover_dea_hi.pdf")
    used = INDEX_STEPS[0]
    for k, step in enumerate(INDEX_STEPS):
        progress("Building cover page" + (" (fitting index)" if k else ""), 90)
        build_cover(cover_docx, date, topics, numbers, qr_png, step, fonts)
        docx_to_pdf(cover_docx, cover_pdf)
        used = step
        ok, _bottom = _index_ok(cover_pdf, topics, INDEX_TOP, INDEX_FLOOR, INDEX_X, INDEX_R)
        if ok:
            return cover_pdf, cover_docx, used, True
    return cover_pdf, cover_docx, used, False


def video_qr(days, typed, workdir, date=None):
    """(png or None, note). Typed link wins; else the QR of the day the
    cover is dated for (the latest day), else the first QR in the file."""
    typed = (typed or "").strip()
    path = os.path.join(workdir, "_video_qr.png")
    if typed:
        png, got = C.qr_from_link(typed, path)
        if png:
            return png, f"video QR made from the link typed ({got})"
        note = f"video link ignored: {got}"
    else:
        note = ""
    dated = sorted([d for d in days if d.date and d.qr], key=lambda d: d.date)
    pool = ([dated[-1].qr[0]] if dated else []) + [q for d in days for q in d.qr]
    if pool:
        link, _blob = pool[0]
        others = sorted({l for d in days for l, _b in d.qr} - {link})
        return C.render_qr(link, path), (
            (note + "; " if note else "") + f"video QR from the file -> {link}"
            + (f" (the file also holds {', '.join(others)}; type a link to choose another)"
               if others else ""))
    return None, (note + "; " if note else "") + (
        "no video QR in the file and no link typed - the QR panel is left out")


def lookup_pages(styled_docx, targets, first_number, pdf_pages, cancelled=lambda: False):
    """{title: printed page number}, from Word on the very styled document
    that produced the PDF. Returns ({}, warning) when Word's pagination does
    not match the PDF - a wrong page number is worse than none."""
    from .cover_hindi import page_lookup
    try:
        found, total = page_lookup(styled_docx, targets, cancelled)
    except InterruptedError:
        raise
    except Exception as exc:
        return {}, f"page numbers could not be read from Word ({type(exc).__name__})"
    if total != pdf_pages:
        return {}, (f"Word paginates the document to {total} pages but the PDF has "
                    f"{pdf_pages}; index printed without page numbers")
    missing = [t for t in targets if t not in found]
    warn = ("no page found for: " + "; ".join(f"'{m[:40]}'" for m in missing)) if missing else ""
    return {t: p + first_number - 1 for t, p in found.items()}, warn


def attach_cover(info, numbers, final_pdf, companion, workdir, date_in=None, index_in=None,
                 qr_in=None, progress=lambda s, p: None, notes=None):
    from .cover import _lines, prepend_docx, prepend_pdf
    notes = notes if notes is not None else []
    fonts = cover_fonts()
    date, date_src = hindi_date_text(date_in, info["days"])
    topics = _lines(index_in) or info["topics"]
    qr_png, qr_note = video_qr(info["days"], qr_in, workdir)
    cover_pdf, cover_docx, used, fits = render_cover(workdir, date, topics, numbers, qr_png,
                                                     fonts, progress)
    prepend_pdf(cover_pdf, final_pdf, final_pdf)
    extra = ""
    if companion and os.path.exists(companion):
        try:
            prepend_docx(cover_docx, companion, companion)
        except Exception as exc:
            extra = f"; the Word file could not take it ({type(exc).__name__})"
    bits = [f"date {date_src}" if date else "date MISSING - type it in the Date block",
            f"{len(topics)} index line(s) at {used[0]:g}pt",
            f"page numbers {len([t for t in topics if numbers.get(t)])}/{len(topics)}"]
    if not fits:
        bits.append("the index is still too long for the panel - shorten it in the Index block")
    notes.append("Cover page added: " + ", ".join(bits) + extra + ".")
    notes.append(f"Cover fonts: date '{fonts['date']}', index '{fonts['index']}'.")
    if qr_note:
        notes.append("Cover: " + qr_note + ".")
    return True


# ================================================================== pipeline

def output_name(stem, variant="Hindi"):
    return f"{stem}_Daily Editorial Summary_UPSC_{variant}.pdf"


def build_body(input_path, workdir, notes, fonts, progress=lambda s, p: None,
               cancelled=lambda: False):
    """Styled docx + its raw Word PDF + page lookups - no stamping yet, so the
    combined edition can number and fill it as part of the whole booklet."""
    stem = os.path.splitext(os.path.basename(input_path))[0]
    progress("Styling editorial (Hindi)", 20)
    styled = os.path.join(workdir, f"{stem}_styled_dea_hi.docx")
    info = restyle(input_path, styled, notes, fonts)
    if cancelled():
        raise InterruptedError("stopped")
    progress("Converting editorial to PDF (Word)", 45)
    raw_pdf = os.path.join(workdir, f"{stem}_raw_dea_hi.pdf")
    docx_to_pdf(styled, raw_pdf)
    if cancelled():
        raise InterruptedError("stopped")
    with fitz.open(raw_pdf) as d:
        pages = len(d)
    progress("Reading page numbers (Word)", 62)
    found, warn = lookup_pages(styled, info["topics"], 1, pages, cancelled)
    if warn:
        notes.append("Page numbers: " + warn + ".")
    info["pages"] = pages
    info["found"] = found                        # 1-based pages inside this part
    info["day_end_pages"] = []
    if len(info["days"]) > 1:
        for k in info["day_first_topic"][1:]:
            if k < len(info["topics"]) and found.get(info["topics"][k]):
                end_page = found[info["topics"][k]] - 2
                # A topic on page 1 has no previous editorial page. Empty
                # leading days or repeated titles can otherwise yield -1,
                # which the combined edition offsets into the CA section.
                if 0 <= end_page < pages - 1 and end_page not in info["day_end_pages"]:
                    info["day_end_pages"].append(end_page)
    return styled, raw_pdf, info


def process(input_path, workdir, variant="Hindi", progress=lambda s, p: None,
            cancelled=lambda: False, cover_date=None, cover_index=None, cover_qr=None, **kw):
    os.makedirs(workdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(input_path))[0]
    notes = []
    progress("Checking Devanagari fonts", 8)
    fonts = H.resolve_fonts(notes)
    styled, raw_pdf, info = build_body(input_path, workdir, notes, fonts, progress, cancelled)
    progress("Branding pages & promo", 75)
    final_pdf = os.path.join(workdir, output_name(stem, variant))
    finalize(raw_pdf, final_pdf, H.FIRST_PAGE_NUMBER, info["day_end_pages"], notes)
    companion = os.path.splitext(final_pdf)[0] + ".docx"
    try:
        shutil.copy2(styled, companion)
    except OSError:
        companion = None
    if cancelled():
        raise InterruptedError("stopped")
    numbers = {t: p + H.FIRST_PAGE_NUMBER - 1 for t, p in info["found"].items()}
    progress("Building cover page", 88)
    try:
        attach_cover(info, numbers, final_pdf, companion, workdir, cover_date, cover_index,
                     cover_qr, progress, notes)
    except InterruptedError:
        raise
    except Exception as exc:
        notes.append(f"No cover page ({type(exc).__name__}: {exc}); the booklet is unchanged.")
    progress("Done", 100)
    return final_pdf, notes
