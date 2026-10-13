# -*- coding: utf-8 -*-
"""Front page for the Daily CA (English Xylem) booklet.

The Xylem cover template is the English cover template's twin: the same two
text boxes (index with the tick / square list prototypes, the orange date
pill), the same anchors, a different first-page artwork, and NO first-page
footer - the published Xylem cover carries no WhatsApp / Telegram buttons.
Rendered untouched it lands exactly on the published cover, so this module
reuses the English cover's fill, measure and prepend machinery unchanged
(imported, not copied) and only owns the three things that differ: the
template path, the index-box height the client enlarged by hand, and the
box floor measured from that.

Same contract as English: a filled block is used verbatim, an empty one is
read from the raw file (date from the Title line, topics from Heading 1
outside tables, shorts from the news table).
"""
import os

from docx import Document

from .convert import docx_to_pdf
from . import cover as _en

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEMPLATE = os.path.join(ROOT, "assets", "cover", "daily_ca_xylem_cover.docx")

# measured from the published cover: box 56.2..559.5 x 353.2..654.0
INDEX_BOX_TOP_PT = 353.2
INDEX_BOX_H_PT = 300.8          # template ships 278.4; the published box is taller
INDEX_BOX_INSET_PT = 3.6
INDEX_BOX_BOTTOM_PT = INDEX_BOX_TOP_PT + INDEX_BOX_H_PT - INDEX_BOX_INSET_PT   # 650.4
INDEX_SIZE_STEPS = _en.INDEX_SIZE_STEPS

extract_from_raw = _en.extract_from_raw
resolve = _en.resolve


def _fill_index(tb, index, shorts, sizes):
    """Like the English fill, with a third prototype: the template's own
    "News in Shorts" paragraph. In this template that line carries explicit
    spacing (line 271 against the list style's 288/300); cloning the first
    tick paragraph for it, as the English fill does, put every square line
    2.3pt lower than the published cover."""
    from docx.oxml.ns import qn
    import copy
    paras = tb.findall(qn('w:p'))
    tick_proto = square_proto = news_proto = None
    for p in paras:
        style = p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
        sid = style.get(qn('w:val')) if style is not None else ""
        if not _en._text(p):
            continue
        if sid == "ListParagraph":
            if tick_proto is None:
                tick_proto = p
            if square_proto is None:
                news_proto = p                  # the last tick before the squares
        elif sid == "ListParagraph2" and square_proto is None:
            square_proto = p
    if tick_proto is None:
        return False
    # explicit None tests: lxml elements are falsy when childless
    if square_proto is None:
        square_proto = tick_proto
    if news_proto is None:
        news_proto = tick_proto

    def make(proto, text, half_pts):
        p = copy.deepcopy(proto)
        runs = [r for r in p.findall(qn('w:r'))]
        proto_run = next((r for r in runs if r.findall(qn('w:t'))),
                         runs[0] if runs else None)
        for r in runs:
            p.remove(r)
        if proto_run is not None:
            p.append(_en._clone_run(proto_run, text))
        for r in p.findall(qn('w:r')):
            rpr = r.find(qn('w:rPr'))
            if rpr is None:
                rpr = r.makeelement(qn('w:rPr'), {})
                r.insert(0, rpr)
            _en._force_size(rpr, half_pts)
        ppr = p.find(qn('w:pPr'))
        if ppr is not None:
            mrpr = ppr.find(qn('w:rPr'))
            if mrpr is None:
                mrpr = ppr.makeelement(qn('w:rPr'), {})
                ppr.append(mrpr)
            _en._force_size(mrpr, half_pts)
        return p

    tick_sz, square_sz = sizes
    for p in paras:
        tb.remove(p)
    for line in index:
        tb.append(make(tick_proto, line, tick_sz))
    if shorts:
        tb.append(make(news_proto, _en.SHORTS_HEADING, tick_sz))
        for line in shorts:
            tb.append(make(square_proto, line, square_sz))
    return True


def build_cover_docx(out_path, date, index, shorts, sizes):
    """Fill the Xylem template; both continuous sections are left as they are
    (the cover page draws its furniture from the first one)."""
    doc = Document(TEMPLATE)
    _en._strip_page_fields(doc)
    body = doc.element.body
    filled = {"date": 0, "index": 0}
    for tb in _en._textboxes(body):
        kind = _en._classify(tb)
        if kind == "date":
            _en._fill_date(tb, date)            # empty date -> empty pill, never the sample
            filled["date"] += 1 if date else 0
        elif kind == "index":
            _en._size_index_box(tb, INDEX_BOX_H_PT)
            if _fill_index(tb, index, shorts, sizes):
                filled["index"] += 1
    doc.save(out_path)
    return filled


def _index_metrics(pdf_path):
    """(lowest text edge, bullet-line count) inside the index box. Word hides
    what overflows a fixed box, so the count matters as much as the edge."""
    import fitz
    d = fitz.open(pdf_path)
    page = d[0]
    bottom, bullets = 0.0, 0
    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0:
            continue
        for line in b["lines"]:
            y0, y1 = line["bbox"][1], line["bbox"][3]
            if y0 < INDEX_BOX_TOP_PT - 4 or y1 > 700:
                continue
            text = "".join(sp["text"] for sp in line["spans"]).strip()
            if not text:
                continue
            bottom = max(bottom, y1)
            if text[0] in _en.BULLET_GLYPHS:
                bullets += 1
    d.close()
    return bottom, bullets


def render_cover(workdir, date, index, shorts, progress=lambda s, p: None):
    """Fill, render, and shrink the index only if it really overflows.
    Returns (cover_pdf, cover_docx, size_pair_used, fits)."""
    os.makedirs(workdir, exist_ok=True)
    cover_docx = os.path.join(workdir, "_cover_xylem.docx")
    cover_pdf = os.path.join(workdir, "_cover_xylem.pdf")
    expected = len(index) + ((1 + len(shorts)) if shorts else 0)
    used, fits = INDEX_SIZE_STEPS[0], False
    for step, sizes in enumerate(INDEX_SIZE_STEPS):
        progress("Building cover page" + (" (fitting index)" if step else ""), 90)
        build_cover_docx(cover_docx, date, index, shorts, sizes)
        docx_to_pdf(cover_docx, cover_pdf)
        used = sizes
        bottom, bullets = _index_metrics(cover_pdf)
        fits = bottom <= INDEX_BOX_BOTTOM_PT and bullets >= expected
        if fits:
            break
    return cover_pdf, cover_docx, used, fits


def attach_cover(raw_docx, body_pdf, body_docx, workdir, out_pdf, out_docx,
                 date_in=None, index_in=None, shorts_in=None,
                 progress=lambda s, p: None, notes=None):
    """The whole stage. Returns True when a cover was added."""
    notes = notes if notes is not None else []
    date, index, shorts, source = resolve(raw_docx, date_in, index_in, shorts_in)
    if not index and not shorts and not date:
        notes.append("No cover page: the file has no topic headings, no news "
                     "table and no date line, and none were typed in.")
        return False
    cover_pdf, cover_docx, used, fits = render_cover(workdir, date, index, shorts, progress)
    _en.prepend_pdf(cover_pdf, body_pdf, out_pdf)
    if body_docx and os.path.exists(body_docx):
        try:
            _en.prepend_docx(cover_docx, body_docx, out_docx)
        except Exception as exc:               # never lose the Word file
            notes.append(f"Cover added to the PDF only; the Word companion "
                         f"could not take it ({type(exc).__name__}: {exc}).")
    parts = [f"date {source['date']}" if date else
             "date NOT FOUND (no 'News In Shorts: <date>' title line) - the "
             "pill is left empty, type the date in the Date block",
             f"index {source['index']} ({len(index)})",
             f"shorts {source['shorts']} ({len(shorts)})"]
    step = INDEX_SIZE_STEPS.index(used)
    fit = "" if step == 0 else f"; index reduced to {used[0] // 2}pt to fit the box"
    if not fits:
        fit += ("; the index is still too long for the box even at that size - "
                "Word hides the lines past the bottom, shorten the index")
    notes.append("Cover page added: " + ", ".join(parts) + fit + ".")
    return True
