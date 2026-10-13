# -*- coding: utf-8 -*-
"""Daily Current Affairs + Daily Editorial Analysis (Hindi) - one booklet.

Specification: ``Output Daily CA+Daily Editorial_17 Aug_UPSC_Hindi.pdf`` (19
pages). Unlike the English combined edition (two booklets joined), this one
is designed as ONE booklet, measured:

    p1       combined cover: डेली न्यूज़ एनालिसिस + date + the CA index, then
             दैनिक संपादकीय सारांश + the editorial index, page numbers on both
    p2-12    the Daily CA half; its first page carries the section badge
             "डेली न्यूज़ एनालिसिस" in place of the logo, and the SRIJAN banner
             fills the space its last page leaves
    p13-18   the editorial half, badge "दैनिक संपादकीय सारांश" on its first
             page; its days oldest first, each from a new page, the SRIJAN
             banner where a day ends short of the foot
    p19      the "अन्य पुस्तकें एवं कार्यक्रम" closing page
    numbers  continuous, 2 to the end

Shared stylers identify roles. The combined layout then applies continuous
articles, measured leading, shaded shorts and section-ending advertisements.
Standalone category layout rules are kept separate.
"""
import os

import fitz

from . import cover_kit as K
from . import combo_hindi_layout as L
from . import daily_ca_hindi as H
from . import dea_common as C
from . import dea_hindi as D
from .convert import docx_to_pdf

ASSETS = C.ASSETS
BADGE_DNA = (os.path.join(ASSETS, "hi_badge_dna.png"), fitz.Rect(184.8, 24.6, 427.1, 67.4))
BADGE_DEA = (os.path.join(ASSETS, "hi_badge_dea.png"), fitz.Rect(175.9, 27.4, 436.1, 67.8))
LOGO_PATCH = fitz.Rect(255.0, 31.5, 360.0, 60.5)     # the header logo, whited out

PLATE = os.path.join(ASSETS, "hi_combo_cover_plate.jpg")
DATE_BOX = (235.3, 288.0, 141.4, 36.0)               # centred on the pill at y=306
IDX_X, IDX_R = 36.2, 575.8
IDX_TAB = 11.0
# floors = the next printed thing minus a small gap, measured against the
# LOWEST GLYPH (Karma's line boxes run taller than the published Utsaah's)
CA_TOP, CA_FLOOR = 324.0, 513.0                      # editorial title glyphs start at 519.8
DEA_TOP, DEA_FLOOR = 573.5, 695.5                    # the navy band starts at 697.9
# (size, pitch) - published: CA list 12pt on 29.2pt, editorial list 12pt on 24.1pt
CA_STEPS = [(12.0, 29.2), (11.5, 26.5), (11.0, 24.0), (10.0, 21.5), (9.0, 19.0)]
DEA_STEPS = [(12.0, 24.1), (11.0, 22.0), (10.0, 20.0), (9.0, 18.0)]
TICK = ("Wingdings 2", "F050", 7.0, H.RED)
TEXT_CHAIN = ("Utsaah", "Karma", "Nirmala UI")
DATE_CHAIN = ("Utsaah", "Karma", "Nirmala UI")


def output_name(stem, variant="Hindi"):
    return f"{stem}_CA Analysis+Daily Editorial Summary_UPSC_{variant}.pdf"


# ------------------------------------------------------------------ cover

def cover_fonts():
    from .cover_hindi import _installed_families, _pick
    avail = _installed_families()
    return {"text": _pick(TEXT_CHAIN, avail), "date": _pick(DATE_CHAIN, avail)}


def _list(cv, entries, numbers, top, step, font, name):
    size, pitch = step
    paras = D.index_paragraphs(entries, numbers, size, pitch, font, H.INK, IDX_TAB, IDX_R, IDX_X,
                               numbered=False, tick=TICK)
    cv.textbox((IDX_X, top, IDX_R - IDX_X, 792.0 - top), paras, name=name)


def build_cover(path, date, lists, fonts):
    """`lists` = [(entries, numbers, top, step, name)]."""
    cv = K.Cover()
    cv.background(PLATE)
    if date:
        cv.textbox(DATE_BOX, [K.para([K.run(date, fonts["date"], 19, H.INK, bold=True)], jc="center")],
                   anchor="ctr", name="Date")
    for entries, numbers, top, step, name in lists:
        if entries:
            _list(cv, entries, numbers, top, step, fonts["text"], name)
    return cv.save(path)


def _measure(pdf_path, entries, top):
    lines = K.content_lines(pdf_path, (IDX_X - 4, top - 2, IDX_R + 4, 792))
    ticks = [l for l in lines if "wingdings" in l[4].lower()]
    return len(ticks) >= len(entries), max((l[1] for l in lines), default=top)


def render_cover(workdir, date, ca, dea, fonts, progress=lambda s, p: None):
    """Each list measured alone (tall unframed box, nothing clipped), stepped
    down only when it runs past its floor; then the page drawn once."""
    cover_docx = os.path.join(workdir, "_cover_combo_hi.docx")
    cover_pdf = os.path.join(workdir, "_cover_combo_hi.pdf")
    probe_docx = os.path.join(workdir, "_cover_combo_hi_probe.docx")
    probe_pdf = os.path.join(workdir, "_cover_combo_hi_probe.pdf")
    plan, fits = [], True
    for (entries, numbers), top, floor, steps, name in ((ca, CA_TOP, CA_FLOOR, CA_STEPS, "CA index"),
                                                         (dea, DEA_TOP, DEA_FLOOR, DEA_STEPS, "Editorial index")):
        chosen = steps[0]
        if entries:
            for k, step in enumerate(steps):
                progress("Building cover page (measuring the index)", 89)
                build_cover(probe_docx, "", [(entries, numbers, top, step, name)], fonts)
                docx_to_pdf(probe_docx, probe_pdf)
                present, bottom = _measure(probe_pdf, entries, top)
                chosen = step
                if present and bottom <= floor:
                    break
                if k == len(steps) - 1:
                    fits = False
        plan.append((entries, numbers, top, chosen, name))
    progress("Building cover page", 90)
    build_cover(cover_docx, date, plan, fonts)
    docx_to_pdf(cover_docx, cover_pdf)
    return cover_pdf, cover_docx, plan[0][3], plan[1][3], fits


# ------------------------------------------------------------------ pages

def _badge(page, badge):
    path, rect = badge
    page.draw_rect(LOGO_PATCH, color=None, fill=(1, 1, 1))
    page.insert_image(rect, filename=path)


def _pages_of(styled_docx, targets, pdf_pages, notes, label, cancelled):
    found, warn = D.lookup_pages(styled_docx, targets, 1, pdf_pages, cancelled)
    if warn:
        notes.append(f"Page numbers ({label}): {warn}.")
    return found


# ------------------------------------------------------------------ pipeline

def process(input_path, workdir, variant="Hindi", progress=lambda s, p: None,
            cancelled=lambda: False, second_path=None, cover_date=None, cover_index=None,
            cover_index2=None, **kw):
    """input_path = the Daily CA raw file, second_path = the Daily Editorial one."""
    if not second_path:
        raise ValueError("This category needs two files: the Daily Current Affairs "
                         "file and the Daily Editorial Analysis file.")
    from .cover_hindi import SHORTS_HEADING, extract_from_raw
    from .cover import _lines, prepend_docx, prepend_pdf
    os.makedirs(workdir, exist_ok=True)
    notes = []
    # Missing decoding support must not silently leave cover-only QR images
    # in the combined body and add whole pages to the booklet.
    try:
        import zxingcpp
    except ImportError as exc:
        raise RuntimeError('Hindi combined formatting needs zxing-cpp. Install the app requirements and restart DocPilot.') from exc
    input_path, second_path = C.order_pair(input_path, second_path, notes, hindi=True)
    stem = os.path.splitext(os.path.basename(input_path))[0]
    progress("Checking Devanagari fonts", 5)
    fonts = H.resolve_fonts(notes)
    notes.append('Combined Hindi layout: continuous articles, compact template tables, '
                 'section-ending ads and editable Word page furniture. '
                 + ('Bundled Baloo 2 uses a measured 9pt body / 10pt heading profile.'
                    if fonts[1].replace(' ', '').lower() == 'baloo2'
                    else 'The detected body font retains the template size.'))

    # ---- 1) shared role styling, then this category's continuous layout
    progress("Styling Daily CA (Hindi)", 10)
    styled_ca = os.path.join(workdir, f"{stem}_styled_hi.docx")
    H.restyle_docx(input_path, styled_ca, fonts, notes)
    L.normalize(styled_ca, workdir, fonts)
    L.add_badge(styled_ca, *BADGE_DNA)
    if cancelled():
        raise InterruptedError("stopped")
    progress("Rendering continuous Daily CA pages (Word)", 20)
    raw_ca = os.path.join(workdir, f"{stem}_raw_hi.pdf")
    docx_to_pdf(styled_ca, raw_ca)
    if cancelled():
        raise InterruptedError("stopped")
    with fitz.open(raw_ca) as d:
        ca_pages = len(d)
    ca_index, bar, _shorts = extract_from_raw(input_path)
    ca_entries = list(ca_index) + ([bar or SHORTS_HEADING] if (bar or _shorts) else [])
    progress("Reading Daily CA page numbers (Word)", 35)
    ca_found = _pages_of(styled_ca, ca_entries, ca_pages, notes, "Daily CA", cancelled)

    # ---- 2) the editorial half
    progress("Styling continuous editorial pages (Hindi)", 45)
    styled_dea = os.path.join(workdir, 'combo_styled_dea_hi.docx')
    info = D.restyle(second_path, styled_dea, notes, fonts)
    L.normalize(styled_dea, workdir, fonts, editorial=True, topics=info['topics'],
                day_starts=info['day_first_topic'], first_number=ca_pages + 2)
    L.add_badge(styled_dea, *BADGE_DEA)
    raw_dea = os.path.join(workdir, 'combo_raw_dea_hi.pdf')
    docx_to_pdf(styled_dea, raw_dea)
    with fitz.open(raw_dea) as d:
        info['pages'] = len(d)
    info['found'] = _pages_of(styled_dea, info['topics'], info['pages'], notes, 'Editorial', cancelled)
    ca_markers = L.marker_pages(styled_ca, cancelled)
    dea_markers = L.marker_pages(styled_dea, cancelled)
    if cancelled():
        raise InterruptedError("stopped")

    # ---- 3) one booklet: number, badge, fill, close
    progress("Joining, numbering and badging the pages", 78)
    out = fitz.open(raw_ca)
    with fitz.open(raw_dea) as d:
        out.insert_pdf(d)
    placed = L.fill_promos(out, ca_pages, ca_markers, dea_markers)
    filled = [f"{key} (p{pno + 2})" for key, (_, _, pno) in placed.items()]
    L.add_promos(styled_ca, {k: v for k, v in placed.items() if k in ca_markers})
    L.add_promos(styled_dea, {k: v for k, v in placed.items() if k in dea_markers})
    D.add_closing_page(out)
    final_pdf = os.path.join(workdir, output_name(stem, variant))
    out.save(final_pdf, garbage=3, deflate=True)
    total = len(out) + 1                           # + the cover, prepended below
    out.close()
    notes.append(f"Daily CA: pages 2-{ca_pages + 1}; editorials: pages {ca_pages + 2}-"
                 f"{ca_pages + 1 + info['pages']}; closing page {total}. "
                 + ("Promo banner " + ", ".join(filled) + "." if filled
                    else "No promo banner: no section leaves enough space at its foot."))
    companion = os.path.splitext(final_pdf)[0] + ".docx"
    try:
        C.merge_docx(styled_ca, styled_dea, companion)
        L.add_closing(companion)
    except Exception as exc:
        notes.append(f"Word companion not produced ({type(exc).__name__}: {exc}); the PDF is complete.")
        companion = None

    # ---- 4) the combined cover
    progress("Building cover page", 88)
    try:
        cfonts = cover_fonts()
        date, date_src = D.hindi_date_text(cover_date, info["days"])
        ca_list = _lines(cover_index) or ca_entries
        dea_list = _lines(cover_index2) or info["topics"]
        ca_numbers = {t: p + 1 for t, p in ca_found.items()}
        dea_numbers = {t: ca_pages + 1 + p for t, p in info["found"].items()}
        cover_pdf, cover_docx, ca_step, dea_step, fits = render_cover(
            workdir, date, (ca_list, ca_numbers), (dea_list, dea_numbers), cfonts, progress)
        prepend_pdf(cover_pdf, final_pdf, final_pdf)
        if companion:
            try:
                prepend_docx(cover_docx, companion, companion)
            except Exception as exc:
                notes.append(f"Cover added to the PDF only ({type(exc).__name__}).")
        bits = [f"date {date_src}" if date else "date MISSING - type it in the Date block",
                f"{len(ca_list)} Daily CA line(s) at {ca_step[0]:g}pt "
                f"(page numbers {sum(1 for t in ca_list if ca_numbers.get(t))}/{len(ca_list)})",
                f"{len(dea_list)} editorial line(s) at {dea_step[0]:g}pt "
                f"(page numbers {sum(1 for t in dea_list if dea_numbers.get(t))}/{len(dea_list)})"]
        if not fits:
            bits.append("an index is still too long for its panel - shorten it in the Index block")
        notes.append("Cover page added: " + ", ".join(bits) + ".")
        notes.append(f"Cover fonts: text '{cfonts['text']}', date '{cfonts['date']}'.")
        video = sorted({l for d in info["days"] for l, _b in d.qr})
        if video:
            notes.append("Video QR codes in the editorial file (" + ", ".join(video) + ") are "
                         "not printed: the combined edition's cover carries none, as published.")
    except InterruptedError:
        raise
    except Exception as exc:
        notes.append(f"No cover page ({type(exc).__name__}: {exc}); the booklet is unchanged.")
    progress("Done", 100)
    return final_pdf, notes
