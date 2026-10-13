# -*- coding: utf-8 -*-
"""Daily Editorial Analysis (English) - "Daily Editorial Summary" booklet.

Specification: the published sample
``Output 17th August 2026_Daily Editorial Summary_UPSC_English.pdf``.

The body is the English Daily CA design to the point - measured, not assumed:
16pt #C00000 Times titles, 11pt #00B050 GS line, 11pt #0070C0 headings,
Wingdings bullets on a 21pt hang, the Cambria "Topics" banner, the blue
double page border, logo header, Join buttons and numbered blue circle. So the
styling primitives are the English styler's own (imported, never edited);
only the reading of the raw file is new (see dea_common: no heading styles,
colour-coded titles, multi-day files).

What differs from Daily CA, all measured from the sample:

* topics FLOW - they do not start on a new page; a new DAY does
* a third bullet level: ✓ (Wingdings F0FC) one 21pt step past ➢
* footer buttons 8pt lower (734.7pt) than the CA template's
* an end mark (❍❍❍❍, Wingdings F0A6) under the last line
* the PYQ banner is set into the space left on the last page; the closing
  page then carries the big SRIJAN promo. When the space is too short even
  for the banner at 72% (its QR stops scanning below that), the booklet falls
  back to the Daily CA closing page, which carries both promos.
* the cover (see build_cover): date pill, "Topics to be covered", the day's
  video QR from the raw file - or from a link the user typed.
"""
import os
import re
import shutil

import fitz
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from . import dea_common as C
from . import cover_kit as K
from .convert import docx_to_pdf
from .ooxml import drop_google_tab_frontmatter, get_or_add, pt_to_emu
from .styler import (BLUE, _cleanup_source, _make_banner_el, _normalize_floating_images,
                     _rebuild_news_table, _set_table_borders, _style_blue_heading,
                     _style_body, _style_cell_paragraph, _style_h1, _style_infobox,
                     _style_list_item, _style_news_table, _style_subject, _cap_drawing,
                     _cell_paragraphs, COL_W_PT)
from .template import LINK_TELEGRAM, LINK_WHATSAPP, apply_page_template
from . import postprocess as P

ASSETS = C.ASSETS
BANNER_TEXT = "Topics"

# footer buttons on the DEA pages (the CA template has them at 726.9pt)
BUTTON_Y_PT = 734.7
PAGE_FLOOR = 722.0              # nothing of ours goes below this on a page
# page geometry of the DEA booklet, measured: text runs 78.8 -> 725.7pt (the
# CA template's 76/78pt margins stop it at 714pt). The body pitch is the
# styler's own "at least 14pt": the sample's wobbles 13.8-14.4 (median 14.1)
MARGIN_TOP_PT, MARGIN_BOTTOM_PT = 77.2, 65.0
BANNER_BEFORE, BANNER_AFTER = "70", "166"   # banner -> first title, measured
END_MARK = "F0A6"               # ❍ x4, Wingdings 12pt, centred

# ------------------------------------------------------------------ closing matter
PYQ_BANNER = os.path.join(ASSETS, "en_pyq_banner.jpg")
PYQ_W, PYQ_CX = 552.2, 307.65   # natural width / centre, from the sample
PYQ_GAP = 6.2                   # end mark -> banner
PYQ_MIN_SCALE = 0.72
SRIJAN_FULL = os.path.join(ASSETS, "en_srijan_full.jpg")
SRIJAN_RECT = fitz.Rect(29.9, 82.9, 582.4, 602.5)
CS_ICON_RECT = fitz.Rect(88.8, 675.7, 143.4, 726.5)
WEBAPP_RECT = fitz.Rect(145.2, 678.1, 527.3, 728.4)
# NOT assets/template/promo/pw_webapp_line.png - that file is fully transparent
WEBAPP_LINE = os.path.join(ASSETS, "en_webapp_line.png")
CLOSING_LINKS = [(fitz.Rect(298.6, 678.2, 438.7, 696.5), "https://smart.link/7wwosivoicgd4"),
                 (fitz.Rect(425.6, 696.5, 517.1, 713.9), LINK_TELEGRAM)]
BTN_WHATSAPP = fitz.Rect(56.9, 734.7, 150.1, 757.9)
BTN_TELEGRAM = fitz.Rect(466.6, 734.7, 560.2, 758.0)


# ================================================================== restyle

def _dea_numbering(doc):
    """Bullets as the sample: lvl0 ❖, lvl1 ➢, lvl2+ ✓ - the English CA
    numbering (same indents: 21pt per level) with the third glyph added."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    try:
        root = doc.part.part_related_by(RT.NUMBERING).element
    except KeyError:
        return
    for absnum in root.findall(qn('w:abstractNum')):
        for lvl in absnum.findall(qn('w:lvl')):
            fmt = lvl.find(qn('w:numFmt'))
            if fmt is None or fmt.get(qn('w:val')) != 'bullet':
                continue
            ilvl = int(lvl.get(qn('w:ilvl'), '0'))
            char = {0: '\uf076', 1: '\uf0d8'}.get(ilvl, '\uf0fc')
            left = 420 if ilvl == 0 else 440 + 420 * ilvl
            hang = 420 if ilvl == 0 else 440
            get_or_add(lvl, 'w:lvlText').set(qn('w:val'), char)
            ppr = get_or_add(lvl, 'w:pPr')
            for old in ppr.findall(qn('w:ind')):
                ppr.remove(old)
            ind = ppr.makeelement(qn('w:ind'), {})
            ind.set(qn('w:left'), str(left))
            ind.set(qn('w:hanging'), str(hang))
            ppr.append(ind)
            rpr = lvl.find(qn('w:rPr'))
            if rpr is not None:
                lvl.remove(rpr)
            rpr = lvl.makeelement(qn('w:rPr'), {})
            rf = rpr.makeelement(qn('w:rFonts'), {})
            for a in ('w:ascii', 'w:hAnsi'):
                rf.set(qn(a), 'Wingdings')
            rf.set(qn('w:hint'), 'default')
            rpr.append(rf)
            lvl.append(rpr)


def _bullet_num_id(doc):
    """A numId whose level 0 is a bullet (for typed '•' bullets), creating a
    minimal bullet list definition when the file has none."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    numbering = doc.part.numbering_part.element
    bullet_abs = set()
    for absnum in numbering.findall(qn('w:abstractNum')):
        lvl0 = next((l for l in absnum.findall(qn('w:lvl')) if l.get(qn('w:ilvl')) == '0'), None)
        fmt = lvl0.find(qn('w:numFmt')) if lvl0 is not None else None
        if fmt is not None and fmt.get(qn('w:val')) == 'bullet':
            bullet_abs.add(absnum.get(qn('w:abstractNumId')))
    for num in numbering.findall(qn('w:num')):
        ref = num.find(qn('w:abstractNumId'))
        if ref is not None and ref.get(qn('w:val')) in bullet_abs:
            return num.get(qn('w:numId'))
    ids = [int(a.get(qn('w:abstractNumId'))) for a in numbering.findall(qn('w:abstractNum'))] or [0]
    nids = [int(n.get(qn('w:numId'))) for n in numbering.findall(qn('w:num'))] or [0]
    aid, nid = max(ids) + 1, max(nids) + 1
    levels = "".join(
        f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'<w:lvlText w:val="\u2022"/><w:lvlJc w:val="left"/></w:lvl>' for i in range(3))
    absnum = parse_xml(f'<w:abstractNum {C.W_NS} w:abstractNumId="{aid}">'
                       f'<w:multiLevelType w:val="hybridMultilevel"/>{levels}</w:abstractNum>')
    first_num = numbering.find(qn('w:num'))
    if first_num is not None:
        first_num.addprevious(absnum)
    else:
        numbering.append(absnum)
    numbering.append(parse_xml(f'<w:num {C.W_NS} w:numId="{nid}"><w:abstractNumId w:val="{aid}"/></w:num>'))
    return str(nid)


def _set_numpr(p_el, num_id, ilvl=0):
    ppr = C.ppr_of(p_el)
    old = ppr.find(qn('w:numPr'))
    if old is not None:
        ppr.remove(old)
    ppr.append(parse_xml(f'<w:numPr {C.W_NS}><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>'))
    C.order_ppr(ppr)


def _lower_buttons(doc):
    """The DEA footer buttons sit 8pt lower than the CA template puts them."""
    NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    footer = doc.sections[0].footer._element
    for dp in footer.iter(f"{{{NS_WP}}}docPr"):
        if dp.get("name") in ("btn_whatsapp", "btn_telegram"):
            anchor = dp.getparent()
            pv = anchor.find(f"{{{NS_WP}}}positionV")
            off = pv.find(f"{{{NS_WP}}}posOffset") if pv is not None else None
            if off is not None:
                off.text = str(pt_to_emu(BUTTON_Y_PT))


def _keep_with_next(p_el):
    C.set_flag(p_el, 'w:keepNext', True)


def _head(text):
    p = parse_xml(f'<w:p {C.W_NS}><w:r><w:rPr><w:b/><w:bCs/></w:rPr>'
                  f'<w:t xml:space="preserve"></w:t></w:r></w:p>')
    p.find('.//' + qn('w:t')).text = text
    _style_blue_heading(p)
    _keep_with_next(p)
    return p


def _end_mark():
    return parse_xml(
        f'<w:p {C.W_NS}><w:pPr><w:keepLines/><w:spacing w:before="40" w:after="0" w:line="240" '
        f'w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
        + "".join(f'<w:r><w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/>'
                  f'<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:sym w:font="Wingdings" '
                  f'w:char="{END_MARK}"/></w:r>' for _ in range(4))
        + '</w:p>')


def _style_table(el, doc):
    """Exactly the English styler's table rules."""
    for ps in list(el.iter(qn('w:pStyle'))):
        ps.getparent().remove(ps)
    rows = el.findall(qn('w:tr'))
    ncols = len(rows[0].findall(qn('w:tc'))) if rows else 0
    if len(rows) == 1 and ncols == 1:
        _style_infobox(el, doc)
        return el
    if ncols == 2:
        el = _rebuild_news_table(el)
        _style_news_table(el, doc)
        return el
    _set_table_borders(el, with_inside=True)
    for tc in el.iter(qn('w:tc')):
        for p in _cell_paragraphs(tc):
            _style_cell_paragraph(p, doc)
    return el


def restyle(in_path, out_path, notes, banner_text=BANNER_TEXT):
    """Raw DEA English docx -> branded docx. Returns an info dict:
    days (print order), topics (print order), day_first_topic (indices)."""
    doc = Document(in_path)
    apply_page_template(doc)
    from docx.shared import Pt
    for section in doc.sections:
        section.top_margin, section.bottom_margin = Pt(MARGIN_TOP_PT), Pt(MARGIN_BOTTOM_PT)
    _lower_buttons(doc)
    _dea_numbering(doc)
    body = doc.element.body
    _cleanup_source(body)
    tabs, breaks = drop_google_tab_frontmatter(body, lambda el: C.style_name(doc, el))
    if tabs or breaks:
        notes.append(f"Google Docs tab front matter removed ({tabs} line(s), {breaks} page break(s)).")

    days = C.split_days(doc, body)
    order = C.print_order(days)
    if len(days) > 1 and [id(d) for d in order] != [id(d) for d in days]:
        notes.append("The file holds " + ", ".join(d.date_text or "undated" for d in days)
                     + " - printed oldest first, each day from a new page.")
    C.reassemble(body, order)

    bullet_id = None
    topics, day_first_topic = [], []
    counts = {"headings": 0, "leadins": 0, "bullets": 0, "typed_bullets": 0, "tables": 0, "empty": 0}
    for day in order:
        day_first_topic.append(len(topics))
        blocks = day.blocks
        body_sz = C.body_size_of(doc, blocks)
        awaiting_gs = 0
        for i, el in enumerate(list(blocks)):
            if el.tag == qn('w:tbl'):
                new = _style_table(el, doc)
                if new is not el:
                    blocks[i] = new
                counts["tables"] += 1
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
                C.drop_pstyle(el)
                _style_h1(el)
                _keep_with_next(el)
                topics.append(text)
                awaiting_gs = 2
                continue
            if awaiting_gs:
                awaiting_gs -= 1
                if C.is_gs_line(text):
                    C.drop_pstyle(el)
                    _style_subject(el)
                    C.set_jc(el, "both")
                    _keep_with_next(el)
                    awaiting_gs = 0
                    continue
            if level is None and C.strip_literal_bullet(el):
                bullet_id = bullet_id or _bullet_num_id(doc)
                _set_numpr(el, bullet_id, 0)
                level = 0
                counts["typed_bullets"] += 1
            if level is not None:
                C.drop_pstyle(el)
                _style_list_item(el)
                C.set_jc(el, "both")
                counts["bullets"] += 1
                continue
            C.drop_pstyle(el)
            if C.split_leadin(el, _head) is not None:
                _style_body(el)
                C.set_jc(el, "both")
                counts["leadins"] += 1
                continue
            if C.is_heading(el, doc, next_is_list):
                _style_blue_heading(el)
                C.set_jc(el, "left")
                _keep_with_next(el)
                counts["headings"] += 1
                continue
            _style_body(el)
            C.set_jc(el, "both")
            for inline in el.iter(qn('wp:inline')):
                _cap_drawing(inline, COL_W_PT)

    _normalize_floating_images(body)
    # each later day from a new page - found AFTER styling: a rebuilt table or
    # a dropped empty paragraph would otherwise leave the break on an element
    # that is no longer in the document
    for day in order[1:]:
        first = next((el for el in day.blocks if el.getparent() is not None), None)
        if first is not None:
            C.break_before(first)

    blocks = [e for e in body if e.tag in (qn('w:p'), qn('w:tbl'))]
    if blocks:
        banner = _make_banner_el(banner_text)
        sp = banner.find(qn('w:pPr') + '/' + qn('w:spacing'))
        sp.set(qn('w:before'), BANNER_BEFORE)
        sp.set(qn('w:after'), BANNER_AFTER)
        blocks[0].addprevious(banner)
        body.find(qn('w:sectPr')).addprevious(_end_mark())

    if not topics:
        notes.append("No topic titles found (a bold red or large line) - the file "
                     "may not be a Daily Editorial Analysis document.")
    notes.append(f"{len(topics)} editorial topic(s); {counts['headings']} heading(s), "
                 f"{counts['leadins']} Context/Conclusion lead-in(s), {counts['bullets']} bullet(s)"
                 + (f" ({counts['typed_bullets']} typed)" if counts['typed_bullets'] else "")
                 + (f", {counts['tables']} table(s)" if counts['tables'] else "")
                 + f"; {counts['empty']} empty paragraph(s) removed.")
    doc.save(out_path)
    return {"days": order, "topics": topics, "day_first_topic": day_first_topic}


# ================================================================== pages

def title_pages(pdf_path, count):
    """0-based page of each topic title, read off the render: the titles are
    the only 16pt #C00000 lines (consecutive lines = one wrapped title)."""
    pages = []
    with fitz.open(pdf_path) as d:
        for pno, page in enumerate(d):
            prev_y = None
            for block in page.get_text("dict")["blocks"]:
                for line in block.get("lines", []):
                    spans = [s for s in line["spans"] if s["text"].strip()]
                    if not spans:
                        continue
                    s = spans[0]
                    if abs(s["size"] - 16) < 0.6 and s["color"] == 0xC00000:
                        if prev_y is None or line["bbox"][1] - prev_y > 26:
                            pages.append(pno)
                        prev_y = line["bbox"][1]
                    else:
                        prev_y = None if abs(s["size"] - 16) >= 0.6 else prev_y
    return pages[:count]


def closing_page(doc, number):
    """The DEA closing page: the big SRIJAN promo + the customer-service line,
    on the booklet's own furniture (border, logo, lowered buttons, number)."""
    page = doc.new_page(width=612, height=792)
    page.insert_image(fitz.Rect(0, 0, 612, 792),
                      filename=os.path.join(P.ASSETS, "watermark_page_background.jpeg"))
    for band in P.BORDER_BANDS:
        page.draw_rect(band, color=None, fill=P.BLUE)
    page.insert_image(P.LOGO_RECT, filename=os.path.join(P.ASSETS, "header_logo_pw_onlyias.png"))
    page.insert_image(SRIJAN_RECT, filename=SRIJAN_FULL)
    page.insert_image(CS_ICON_RECT, filename=os.path.join(P.ASSETS, "promo", "customer_service_icon.jpeg"))
    page.insert_image(WEBAPP_RECT, filename=WEBAPP_LINE)
    page.insert_image(BTN_WHATSAPP, filename=os.path.join(P.ASSETS, "btn_join_whatsapp.png"))
    page.insert_image(BTN_TELEGRAM, filename=os.path.join(P.ASSETS, "btn_join_telegram.png"))
    page.insert_link({"kind": fitz.LINK_URI, "from": BTN_WHATSAPP, "uri": LINK_WHATSAPP})
    page.insert_link({"kind": fitz.LINK_URI, "from": BTN_TELEGRAM, "uri": LINK_TELEGRAM})
    for rect, uri in CLOSING_LINKS:
        page.insert_link({"kind": fitz.LINK_URI, "from": rect, "uri": uri})
    P._stamp_page_number(page, number, None)
    return page


def finalize(raw_pdf, out_pdf, first_number, day_end_pages=(), notes=None):
    """Number the pages, set the PYQ banner into leftover space, add the
    closing page. `day_end_pages` = 0-based pages where an earlier day ends
    (a later day starts on the next page), which get a banner too if it fits.
    Returns the number the next page would take."""
    notes = notes if notes is not None else []
    doc = fitz.open(raw_pdf)
    n = first_number
    for page in doc:
        P._stamp_page_number(page, n, None)
        P._ensure_footer_links(page)
        n += 1
    for pno in day_end_pages:
        if 0 <= pno < len(doc) - 1:
            page = doc[pno]
            C.place_banner(page, PYQ_BANNER, C.content_bottom(page, PAGE_FLOOR), PAGE_FLOOR,
                           PYQ_W, PYQ_CX, PYQ_MIN_SCALE, gap=PYQ_GAP + 6)
    last = doc[-1]
    rect = C.place_banner(last, PYQ_BANNER, C.content_bottom(last, PAGE_FLOOR), PAGE_FLOOR,
                          PYQ_W, PYQ_CX, PYQ_MIN_SCALE, gap=PYQ_GAP)
    if rect is not None:
        scale = rect.width / PYQ_W
        closing_page(doc, n)
        notes.append("PYQ banner set into the space left on the last page"
                     + ("" if scale > 0.995 else f" (at {scale:.0%} to fit)")
                     + "; SRIJAN promo on the closing page.")
    else:
        P._build_promo_page(doc, n)
        notes.append("Too little space left on the last page for the PYQ banner - "
                     "both promos go on the closing page instead.")
    n += 1
    doc.save(out_pdf, garbage=3, deflate=True)
    doc.close()
    return n


# ================================================================== cover

COVER_PLATE = os.path.join(ASSETS, "en_dea_cover_plate.jpg")
PILL = (160.2, 228.8, 297.5, 47.7)
TOPICS_X, TOPICS_R = 85.0, 562.0
TOPIC_HANG = 28.3
HEADING_TOP = 317.2            # box top; Word sets the heading's glyphs 1.4pt below it
HEADING_SIZE, HEADING_LINE, HEADING_AFTER = 20.0, 24.0, 10.8  # ticks land at 353.0 as published
# (topic size, line pitch) - the published cover is 16pt on a 24.4pt pitch
TOPIC_STEPS = [(16.0, 24.4), (15.0, 22.9), (14.0, 21.4), (13.0, 19.9), (12.0, 18.3), (11.0, 16.8)]
CAPTION = "To access the video of Editorial Summary:"
CAPTION_TOP, CAPTION_H, CAPTION_GAP = 500.0, 12.6, 6.8
QR_SIDE, QR_MIN, QR_X_CENTER = 155.9, 118.0, 309.05
QR_FLOOR = 706.0                     # the buttons start at 715.1
TOPICS_FLOOR_NO_QR = 700.0
BTN_W = (57.5, 715.1, 93.1, 23.2)
BTN_T = (465.2, 715.1, 93.5, 23.3)
TNR_BOLD = r"C:\Windows\Fonts\timesbd.ttf"


def _wrap_lines(text, size, width, font):
    """How many lines `text` takes at `size` in `width` - word wrap."""
    words, lines, cur = text.split(), 1, 0.0
    space = font.text_length(" ", fontsize=size)
    for w in words:
        wl = font.text_length(w, fontsize=size)
        if cur and cur + space + wl > width:
            lines, cur = lines + 1, wl
        else:
            cur = cur + (space if cur else 0.0) + wl
    return lines


def _tnr_bold():
    try:
        return fitz.Font(fontfile=TNR_BOLD)
    except Exception:
        return fitz.Font("tibo")


TOPIC_TEXT_W = TOPICS_R - TOPICS_X - TOPIC_HANG
TNR_LINE = 1.15                      # Times New Roman single line = 1.15 x size


def _qr_layout(bottom):
    """Caption + QR under a topic list that ends at `bottom`: the QR moves
    down, then shrinks, but never below QR_MIN (it must still scan)."""
    cap_top = max(CAPTION_TOP, bottom + 12.0)
    qr_top = cap_top + CAPTION_H + CAPTION_GAP
    side = min(QR_SIDE, QR_FLOOR - qr_top)
    return {"cap_top": cap_top, "qr_top": qr_top, "qr_side": side}


def build_cover(path, date, topics, qr_png, lay):
    from .cover import date_runs
    cv = K.Cover()
    cv.background(COVER_PLATE)
    x, y, w, h = PILL
    if date:
        runs = [K.run(t, "Arial", 32, "FFFFFF", bold=True, sup=sup) for t, sup in date_runs(date)]
        cv.textbox((x, y, w, h), [K.para(runs, jc="center", line=None)], anchor="ctr", name="Date")
    size, pitch = lay["size"], lay["pitch"]
    font = _tnr_bold()
    paras = [K.para([K.run("Topics to be covered:", "Times New Roman", HEADING_SIZE, BLUE,
                           bold=True, italic=True)],
                    line=HEADING_LINE, after=HEADING_AFTER)]
    # single spacing + space after (a wrapped title reads as one entry), and a
    # title a hair too long for its line is tightened instead of wrapped
    for t in topics:
        paras.append(K.para([K.sym("Wingdings", "F0FC", size, "000000"), K.tab(),
                             K.run(t, "Times New Roman", size, "000000", bold=True,
                                   spacing_pt=K.fit_spacing(t, size, TOPIC_TEXT_W, font))],
                            line=None, after=max(0.0, pitch - TNR_LINE * size),
                            left=TOPIC_HANG, hanging=TOPIC_HANG,
                            tabs=[("left", TOPIC_HANG, None)]))
    box_bottom = (lay["cap_top"] - 2) if lay.get("cap_top") else 792.0
    cv.textbox((TOPICS_X, HEADING_TOP, TOPICS_R - TOPICS_X, box_bottom - HEADING_TOP), paras,
               name="Topics")
    if qr_png and lay.get("cap_top"):
        cv.textbox((60.0, lay["cap_top"], 492.0, CAPTION_H + 2),
                   [K.para([K.run(CAPTION, "Times New Roman", 11, BLUE, bold=True)], jc="center",
                           line=CAPTION_H)], name="Video caption")
        side = lay["qr_side"]
        cv.picture(qr_png, (QR_X_CENTER - side / 2, lay["qr_top"], side, side), name="Video QR")
    cv.picture(os.path.join(ASSETS, "en_btn_whatsapp.png"), BTN_W, link=LINK_WHATSAPP, name="Join WhatsApp")
    cv.picture(os.path.join(ASSETS, "en_btn_telegram.png"), BTN_T, link=LINK_TELEGRAM, name="Join Telegram")
    return cv.save(path)


def _measure_topics(pdf_path, topics):
    """(every topic's tick on the page, bottom edge of the list)."""
    lines = K.content_lines(pdf_path, (TOPICS_X - 4, HEADING_TOP + 20, TOPICS_R, 792))
    ticks = [l for l in lines if l[2] < TOPICS_X + 6 and "wingdings" in l[4].lower()]
    return len(ticks) >= len(topics), max((l[1] for l in lines), default=HEADING_TOP + 30)


def render_cover(workdir, date, topics, qr_png, progress=lambda s, p: None):
    """Measure the list, then place the QR under it.

    Each size is first rendered with the list alone in a box running to the
    page foot, so nothing is clipped and the true bottom edge can be read.
    With a QR the caption and code then go 12pt under that edge - moving
    down, and shrinking to no less than QR_MIN, before the list gives up a
    type size. Without one the list may run to TOPICS_FLOOR_NO_QR."""
    cover_docx = os.path.join(workdir, "_cover_dea.docx")
    cover_pdf = os.path.join(workdir, "_cover_dea.pdf")
    has_qr = bool(qr_png)
    fits, lay = False, None
    for k, (size, pitch) in enumerate(TOPIC_STEPS):
        progress("Building cover page" + (" (fitting topics)" if k else ""), 89)
        probe = {"size": size, "pitch": pitch}
        build_cover(cover_docx, date, topics, None, probe)
        docx_to_pdf(cover_docx, cover_pdf)
        present, bottom = _measure_topics(cover_pdf, topics)
        if not present:
            continue
        if has_qr:
            q = _qr_layout(bottom)
            if q["qr_side"] >= QR_MIN:
                lay, fits = dict(probe, **q), True
                break
        elif bottom <= TOPICS_FLOOR_NO_QR:
            lay, fits = probe, True
            break
    if lay is None:                               # smallest size, best effort
        size, pitch = TOPIC_STEPS[-1]
        lay = {"size": size, "pitch": pitch}
        if has_qr:
            lay.update(cap_top=CAPTION_TOP, qr_top=CAPTION_TOP + CAPTION_H + CAPTION_GAP,
                       qr_side=QR_MIN)
    build_cover(cover_docx, date, topics, qr_png, lay)
    docx_to_pdf(cover_docx, cover_pdf)
    if fits:                                       # the final page lost nothing
        fits = _measure_topics(cover_pdf, topics)[0]
    return cover_pdf, cover_docx, lay, fits


def cover_inputs(info, date_in, index_in, qr_in, workdir, notes):
    """Resolve what the cover shows: typed wins, else the file."""
    from .cover import _lines
    days = info["days"]
    latest = next((d for d in reversed(sorted(
        [d for d in days if d.date], key=lambda d: d.date))), None)
    date = (date_in or "").strip()
    if not date and latest is not None:
        date = C.english_date(latest.date)
    topics = _lines(index_in) or info["topics"]
    qr_png, qr_note = None, ""
    typed = (qr_in or "").strip()
    if typed:
        qr_png, got = C.qr_from_link(typed, os.path.join(workdir, "_video_qr.png"))
        qr_note = (f"video QR made from the link typed ({got})" if qr_png
                   else f"video link ignored: {got}")
    if qr_png is None:
        pool = ([(d, q) for d in ([latest] if latest else []) for q in d.qr]
                + [(d, q) for d in days for q in d.qr])
        if pool:
            day, (link, blob) = pool[0]
            qr_png = C.render_qr(link, os.path.join(workdir, "_video_qr.png"))
            others = sorted({l for d in days for l, _b in d.qr} - {link})
            qr_note = (f"video QR from the file -> {link}"
                       + (f" (the file also holds {', '.join(others)}; type a link to choose another)"
                          if others else ""))
        elif not typed:
            qr_note = "no video QR in the file and no link typed - the cover goes without one"
    return date, topics, qr_png, qr_note


def attach_cover(info, final_pdf, companion, workdir, date_in=None, index_in=None,
                 qr_in=None, progress=lambda s, p: None, notes=None):
    from .cover import prepend_docx, prepend_pdf
    notes = notes if notes is not None else []
    date, topics, qr_png, qr_note = cover_inputs(info, date_in, index_in, qr_in, workdir, notes)
    cover_pdf, cover_docx, lay, fits = render_cover(workdir, date, topics, qr_png, progress)
    prepend_pdf(cover_pdf, final_pdf, final_pdf)
    extra = ""
    if companion and os.path.exists(companion):
        try:
            prepend_docx(cover_docx, companion, companion)
        except Exception as exc:
            extra = f"; the Word file could not take it ({type(exc).__name__})"
    bits = [("date typed" if (date_in or "").strip() else
             ("date from the file" if date else "date MISSING - type it in the Date block")),
            f"{len(topics)} topic(s) at {lay['size']:g}pt"]
    if qr_png:
        bits.append(f"QR {lay['qr_side']:.0f}pt")
    if not fits:
        bits.append("the topic list is still too long for the page - shorten it in the Topics block")
    notes.append("Cover page added: " + ", ".join(bits) + extra + ".")
    if qr_note:
        notes.append("Cover: " + qr_note + ".")
    return True


# ================================================================== pipeline

def output_name(stem, variant="English"):
    return f"{stem}_Daily Editorial Summary_UPSC_{variant}.pdf"


def build_body(input_path, workdir, first_number, notes, progress=lambda s, p: None,
               cancelled=lambda: False, banner_text=BANNER_TEXT):
    """Styled docx + finished body PDF (numbered, promo placed, closing page)
    WITHOUT a cover - the combined edition calls this for its second half."""
    stem = os.path.splitext(os.path.basename(input_path))[0]
    progress("Styling editorial (English)", 20)
    styled = os.path.join(workdir, f"{stem}_styled_dea.docx")
    info = restyle(input_path, styled, notes, banner_text=banner_text)
    if cancelled():
        raise InterruptedError("stopped")
    progress("Converting editorial to PDF (Word)", 45)
    raw_pdf = os.path.join(workdir, f"{stem}_raw_dea.pdf")
    docx_to_pdf(styled, raw_pdf)
    if cancelled():
        raise InterruptedError("stopped")
    progress("Branding editorial pages & promo", 75)
    day_ends = []
    if len(info["days"]) > 1:
        pages = title_pages(raw_pdf, len(info["topics"]))
        for k in info["day_first_topic"][1:]:
            if k < len(pages) and pages[k] > 0:
                day_ends.append(pages[k] - 1)
    body_pdf = os.path.join(workdir, f"{stem}_body_dea.pdf")
    next_no = finalize(raw_pdf, body_pdf, first_number, day_ends, notes)
    return styled, body_pdf, info, next_no


def process(input_path, workdir, variant="English", progress=lambda s, p: None,
            cancelled=lambda: False, cover_date=None, cover_index=None, cover_qr=None, **kw):
    os.makedirs(workdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(input_path))[0]
    notes = []
    styled, body_pdf, info, _ = build_body(input_path, workdir, 2, notes, progress, cancelled)
    final_pdf = os.path.join(workdir, output_name(stem, variant))
    shutil.copy2(body_pdf, final_pdf)
    companion = os.path.splitext(final_pdf)[0] + ".docx"
    try:
        shutil.copy2(styled, companion)
    except OSError:
        companion = None
    if cancelled():
        raise InterruptedError("stopped")
    progress("Building cover page", 88)
    try:
        attach_cover(info, final_pdf, companion, workdir, cover_date, cover_index, cover_qr,
                     progress, notes)
    except InterruptedError:
        raise
    except Exception as exc:
        notes.append(f"No cover page ({type(exc).__name__}: {exc}); the booklet is unchanged.")
    progress("Done", 100)
    return final_pdf, notes
