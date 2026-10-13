# -*- coding: utf-8 -*-
"""Front page for the Daily CA (Hindi) booklet.

A separate stage, like the English cover: the Hindi body pipeline (restyle ->
fitter -> brand -> companion) is signed off and is not touched. This module
fills the client's cover template, renders it through Word, and prepends the
page to the finished PDF and to the Word companion.

The published cover (UPSC sample\\...\\Output_with_frontpage.pdf) is the
specification; the client's Word template is used for its artwork and
shapes, and every object is placed at the coordinates measured from that
PDF (the template itself was drawn on a 612x866pt page, the booklet is
Letter).

What varies day to day, and where it comes from:

    date     <- the Date block (the Hindi raw file carries no date line)
    index    <- the 'Heading 2' paragraphs outside tables - exactly the
                paragraphs the styler turns into topic banners
    bar      <- the 'Heading 1' paragraph ("संक्षेप में समाचार"), the last tick
    shorts   <- the left column of the 2-column news table after that bar
    numbers  <- the page each banner / news row lands on, asked of Word on
                the very fitted docx that produced the PDF - never guessed
                from text

The index is rendered and measured; its size steps down only when it
overflows the box (Word hides overflow, so the bullet lines are counted as
well as the last line's edge).
"""
import copy
import os
import re
import unicodedata

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from .convert import docx_to_pdf
from . import cover as _en

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEMPLATE = os.path.join(ROOT, "assets", "cover", "daily_ca_hindi_cover.docx")
WATERMARK = os.path.join(ROOT, "assets", "cover", "daily_ca_hindi_watermark.png")

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
W_NS = f'xmlns:w="{NS_W}"'

EMU_PT = 12700


def emu(pt):
    return str(int(round(pt * EMU_PT)))


# ------------------------------------------------------------ geometry (pt)
# Every rectangle below is (x, y, w, h) on the Letter page, measured from
# the published cover. Objects are anchored to the PAGE, so the template's
# paragraph flow plays no part.
PAGE_W, PAGE_H = 612.0, 792.0
BAND = (0.0, 0.0, 612.0, 175.7)              # navy band
LOGO = (143.9, 57.6, 125.0, 38.1)            # white pill + PW OnlyIAS logo
TITLE_BOX = (16.5, 82.5, 380.0, 84.0)        # "डेली न्यूज एनालिसिस", centred on x=206.5
DATE_PILL = (135.7, 165.1, 141.4, 20.9)      # white rounded pill
DATE_TEXT = (135.7, 158.0, 141.4, 35.0)      # transparent text box over it
INDEX_BOX = (18.0, 229.2, 364.0, 352.8)      # transparent text box
INDEX_FLOOR = 582.0                          # skyline picture starts at 584.9
WATERMARK_RECT = (82.1, 331.9, 447.8, 128.2)
PIC_GATE = (393.9, 39.7, 179.8, 264.7)
PIC_GLOBE = (392.4, 326.9, 182.8, 162.7)     # groups already carry their frame
PIC_PARL = (392.4, 513.0, 182.8, 248.8)
PIC_SKY = (54.1, 584.9, 295.4, 175.4)
FRAME_COLOR = "BFBDDA"                       # lavender 3pt frame
FRAME_W_EMU = "38100"
BAND_COLOR = "090C4F"                        # published indigo (template: 000066)

# ------------------------------------------------------------ typography
INK = "231F20"
RED = "CB202F"
WHITE = "FFFFFF"
INDEX_HEADING = "विषय सूची"
SHORTS_HEADING = "संक्षेप में समाचार"
TITLE_TEXT = "डेली न्यूज एनालिसिस"
TITLE_CHAIN = ("Rozha One", "Nirmala UI")
TEXT_CHAIN = ("Utsaah", "Nirmala UI")
HEADING_CHAIN = ("Baloo 2", "Nirmala UI")
# index text size (half-points) with its exact line height (twips); the
# published cover is 12pt on a 21.4pt line, steps keep that ratio
INDEX_STEPS = [(24, 428), (22, 392), (20, 357), (18, 321)]
TICK = ("Wingdings 2", "F050")               # thin red tick
ARROW = ("Wingdings", "F0D8")                # red arrow head
TAB_RIGHT_TW = 6942                          # page numbers end at x=374.8pt
IND_TOPIC = (416, 240)                       # left, hanging (twips)
IND_SHORT = (636, 240)
IND_RIGHT = 1200

DEV_DIGITS = str.maketrans("०१२३४५६७८९", "0123456789")
MONTHS_HI = ["जनवरी", "फ़रवरी", "मार्च", "अप्रैल", "मई", "जून",
             "जुलाई", "अगस्त", "सितंबर", "अक्टूबर", "नवंबर", "दिसंबर"]
MONTH_LOOKUP = {}
for _i, _m in enumerate(MONTHS_HI, 1):
    MONTH_LOOKUP[unicodedata.normalize("NFC", _m)] = _i
for _m, _i in {"फरवरी": 2, "सितम्बर": 9, "अक्तूबर": 10, "अक्टुबर": 10,
               "नवम्बर": 11, "दिसम्बर": 12}.items():
    MONTH_LOOKUP[unicodedata.normalize("NFC", _m)] = _i
for _i, _names in enumerate([("january", "jan"), ("february", "feb"), ("march", "mar"),
                             ("april", "apr"), ("may",), ("june", "jun"), ("july", "jul"),
                             ("august", "aug"), ("september", "sep", "sept"),
                             ("october", "oct"), ("november", "nov"),
                             ("december", "dec")], 1):
    for _n in _names:
        MONTH_LOOKUP[_n] = _i


# ================================================================ extraction

def extract_from_raw(raw_docx):
    """(index_lines, bar_text, shorts_lines) as the Hindi styler sees them:
    topics are 'Heading 2' outside tables, the news bar is 'Heading 1', the
    shorts are the left cells of the first 2-column table after the bar."""
    doc = Document(raw_docx)
    body = doc.element.body
    index, bar, shorts = [], "", []
    has_bar = any(_en._style_name(doc, p) == "heading 1" and _en._text(p)
                  for p in body.iter(qn('w:p'))
                  if not any(a.tag == qn('w:tc') for a in p.iterancestors()))
    seen_bar = False
    for block in body:
        if block.tag == qn('w:p'):
            style = _en._style_name(doc, block)
            text = _en._text(block)
            if not text:
                continue
            if style == "heading 2":
                index.append(text)
            elif style == "heading 1" and not bar:
                bar, seen_bar = text, True
        elif block.tag == qn('w:tbl') and not shorts and (seen_bar or not has_bar):
            rows = block.findall(qn('w:tr'))
            if len(rows) < 2 or any(len(r.findall(qn('w:tc'))) != 2 for r in rows):
                continue
            shorts = [_en._text(r.findall(qn('w:tc'))[0]) for r in rows]
            shorts = [s for s in shorts if s]
    return index, bar, shorts


def resolve(raw_docx, index_in=None, shorts_in=None):
    """Typed block wins, empty block is extracted (same contract as English)."""
    ex_index, bar, ex_shorts = extract_from_raw(raw_docx)
    index = _en._lines(index_in) or ex_index
    shorts = _en._lines(shorts_in) or ex_shorts
    source = {"index": "typed" if _en._lines(index_in) else "extracted",
              "shorts": "typed" if _en._lines(shorts_in) else "extracted"}
    return index, bar, shorts, source


# ================================================================ date

def hindi_date(text):
    """'17 August 2026' / '17th Aug 2026' / '17 अगस्त 2026' / '17/08/2026' /
    '2026-08-17' -> ('17 अगस्त 2026', True). Anything else is used verbatim
    (never invented) and flagged False."""
    t = unicodedata.normalize("NFC", (text or "").strip()).translate(DEV_DIGITS)
    t = " ".join(t.split())
    m = re.match(r"^(\d{1,2})(?:st|nd|rd|th)?[\s.,/-]+([A-Za-z]+|[ऀ-ॿ]+)[\s.,/-]+(\d{4})\.?$",
                 t, re.I)
    month = None
    if m:
        month = MONTH_LOOKUP.get(m.group(2).lower()) or MONTH_LOOKUP.get(unicodedata.normalize("NFC", m.group(2)))
        if month:
            return f"{int(m.group(1))} {MONTHS_HI[month - 1]} {m.group(3)}", True
    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", t)
    if m and 1 <= int(m.group(2)) <= 12:
        return f"{int(m.group(1))} {MONTHS_HI[int(m.group(2)) - 1]} {m.group(3)}", True
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", t)
    if m and 1 <= int(m.group(2)) <= 12:
        return f"{int(m.group(3))} {MONTHS_HI[int(m.group(2)) - 1]} {m.group(1)}", True
    return t, False


# ================================================================ fonts

def _installed_families():
    """Lower-cased font titles from the machine + per-user font registries."""
    names = set()
    try:
        import winreg
        key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                key = winreg.OpenKey(root, key_path)
            except OSError:
                continue
            i = 0
            while True:
                try:
                    value_name, _, _ = winreg.EnumValue(key, i)
                except OSError:
                    break
                i += 1
                label = re.sub(r"\s*\((TrueType|OpenType|VGA res)\)\s*$", "", value_name, flags=re.I)
                for part in label.split("&"):
                    names.add(part.strip().lower())
    except Exception:
        pass
    return names


def _pick(chain, available):
    """Same tolerant match as the body styler: exact name, "Family Style", or
    ignoring spaces (Adobe/Microsoft font providers are not consistent about
    them). Returns the name that actually matched - Word is case-insensitive
    about font names but not about a missing space."""
    squashed = {n.replace(" ", ""): n for n in available}
    for family in chain:
        low = family.lower()
        if any(n == low or n.startswith(low + " ") for n in available):
            return family
        hit = squashed.get(low.replace(" ", ""))
        if hit:
            return hit
    return chain[-1]


def resolve_fonts():
    """{'title','text','heading'} -> family, plus a human note."""
    avail = _installed_families()
    fonts = {"title": _pick(TITLE_CHAIN, avail), "text": _pick(TEXT_CHAIN, avail),
             "heading": _pick(HEADING_CHAIN, avail)}
    bits = []
    for key, chain in (("title", TITLE_CHAIN), ("text", TEXT_CHAIN), ("heading", HEADING_CHAIN)):
        bits.append(f"{key} '{fonts[key]}'" + ("" if fonts[key] == chain[0]
                                               else f" ('{chain[0]}' not installed)"))
    return fonts, "Cover fonts: " + ", ".join(bits) + "."


# ================================================================ page numbers

def _norm(s):
    s = unicodedata.normalize("NFC", s or "").replace("\r", " ").replace("\x07", " ")
    return " ".join(s.split())


def _skel(s):
    return "".join(ch for ch in _norm(s).lower() if ch.isalnum())


def page_lookup(fitted_docx, targets, cancelled=lambda: False):
    """Ask Word which page each target text sits on, scanning the cells of
    the top-level tables (topic banners, the news bar and the news grid are
    all tables). Returns ({target: page_1_based}, total_pages)."""
    import pythoncom
    import win32com.client
    from .docx_tools import _word_lock

    want = {}
    for t in targets:
        want.setdefault(_norm(t), t)
    by_skel = {}
    for t in targets:
        by_skel.setdefault(_skel(t), t)
    found = {}
    with _word_lock:
        pythoncom.CoInitialize()
        word = doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(os.path.abspath(fitted_docx), ReadOnly=True,
                                      AddToRecentFiles=False)
            total = int(doc.ComputeStatistics(2))
            for tbl in doc.Tables:
                for cell in tbl.Range.Cells:
                    if cancelled():
                        raise InterruptedError("stopped")
                    txt = _norm(cell.Range.Text)
                    if not txt:
                        continue
                    target = want.get(txt) or by_skel.get(_skel(txt))
                    if target is not None and target not in found:
                        found[target] = int(cell.Range.Information(3))
                if len(found) == len(set(targets)):
                    break
            return found, total
        finally:
            try:
                if doc is not None:
                    doc.Close(0)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()


# ================================================================ template

def _q(ns, tag):
    return f"{{{ns}}}{tag}"


def _anchor_of(body, name):
    for dp in body.iter(_q(NS_WP, "docPr")):
        if dp.get("name") == name:
            return next(a for a in dp.iterancestors() if a.tag == _q(NS_WP, "anchor"))
    return None


def _drop_fallbacks(body):
    """The template writes every shape twice (wps: + a VML fallback for Word
    2007). Only the wps: branch is edited here, so the stale fallback goes."""
    for fb in list(body.iter(_q(NS_MC, "Fallback"))):
        fb.getparent().remove(fb)


def _place(anchor, rect, resize=True):
    """Anchor to the page at (x, y) and, for shapes/pictures, set (w, h)."""
    x, y, w, h = rect
    for tag, off in (("positionH", x), ("positionV", y)):
        el = anchor.find(_q(NS_WP, tag))
        for c in list(el):
            el.remove(c)
        el.set("relativeFrom", "page")
        po = el.makeelement(_q(NS_WP, "posOffset"), {})
        po.text = emu(off)
        el.append(po)
    ext = anchor.find(_q(NS_WP, "extent"))
    if resize:
        ext.set("cx", emu(w))
        ext.set("cy", emu(h))
        for xfrm in anchor.iter(_q(NS_A, "xfrm")):
            parent = xfrm.getparent()
            if parent.tag in (_q(NS_WPS, "spPr"), _q(NS_PIC, "spPr")):
                e = xfrm.find(_q(NS_A, "ext"))
                if e is not None:
                    e.set("cx", emu(w))
                    e.set("cy", emu(h))
                break
    ee = anchor.find(_q(NS_WP, "effectExtent"))
    if ee is not None:
        for k in ("l", "t", "r", "b"):
            ee.set(k, "0")
    wrap = next((c for c in anchor if c.tag.startswith(_q(NS_WP, "wrap"))), None)
    if wrap is not None and wrap.tag != _q(NS_WP, "wrapNone"):
        wrap.tag = _q(NS_WP, "wrapNone")
        for c in list(wrap):
            wrap.remove(c)


def _sppr(anchor):
    for tag in (_q(NS_WPS, "spPr"), _q(NS_PIC, "spPr")):
        el = next(anchor.iter(tag), None)
        if el is not None:
            return el
    return None


def _set_line(sppr, width_emu, color):
    for ln in sppr.findall(_q(NS_A, "ln")):
        sppr.remove(ln)
    sppr.append(parse_xml(
        f'<a:ln xmlns:a="{NS_A}" w="{width_emu}"><a:solidFill><a:srgbClr val="{color}"/>'
        f'</a:solidFill></a:ln>'))


def _no_fill_no_line(sppr):
    for tag in ("solidFill", "gradFill", "blipFill", "pattFill", "noFill", "ln"):
        for el in sppr.findall(_q(NS_A, tag)):
            sppr.remove(el)
    # schema order inside spPr: xfrm, geometry, fill, ln
    geom = next((c for c in sppr if c.tag in (_q(NS_A, "prstGeom"), _q(NS_A, "custGeom"))), None)
    nofill = parse_xml(f'<a:noFill xmlns:a="{NS_A}"/>')
    noline = parse_xml(f'<a:ln xmlns:a="{NS_A}"><a:noFill/></a:ln>')
    if geom is not None:
        geom.addnext(nofill)
    else:
        sppr.append(nofill)
    nofill.addnext(noline)


def _solid_fill(sppr, color):
    for tag in ("solidFill", "gradFill", "blipFill", "pattFill", "noFill"):
        for el in sppr.findall(_q(NS_A, tag)):
            sppr.remove(el)
    geom = next((c for c in sppr if c.tag in (_q(NS_A, "prstGeom"), _q(NS_A, "custGeom"))), None)
    fill = parse_xml(f'<a:solidFill xmlns:a="{NS_A}"><a:srgbClr val="{color}"/></a:solidFill>')
    if geom is not None:
        geom.addnext(fill)
    else:
        sppr.insert(0, fill)


def _white_fill(sppr):
    for tag in ("solidFill", "gradFill", "blipFill", "pattFill", "noFill"):
        for el in sppr.findall(_q(NS_A, tag)):
            sppr.remove(el)
    geom = next((c for c in sppr if c.tag in (_q(NS_A, "prstGeom"), _q(NS_A, "custGeom"))), None)
    fill = parse_xml(f'<a:solidFill xmlns:a="{NS_A}"><a:srgbClr val="{WHITE}"/></a:solidFill>')
    if geom is not None:
        geom.addnext(fill)
    else:
        sppr.insert(0, fill)


def _body_pr(anc, **attrs):
    bp = next(anc.iter(_q(NS_WPS, "bodyPr")), None)
    if bp is not None:
        for k, v in attrs.items():
            bp.set(k, str(v))
    return bp


def _txbx(anchor):
    return next(anchor.iter(qn('w:txbxContent')), None)


def _clear(tb):
    for p in list(tb.findall(qn('w:p'))):
        tb.remove(p)


def _run(text, font, half_pts, color, bold=False, w_pct=None):
    b = "<w:b/><w:bCs/>" if bold else ""
    wv = f'<w:w w:val="{w_pct}"/>' if w_pct else ""
    r = parse_xml(
        f'<w:r {W_NS}><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
        f'{b}<w:color w:val="{color}"/>{wv}<w:sz w:val="{half_pts}"/><w:szCs w:val="{half_pts}"/>'
        f'<w:cs/><w:lang w:bidi="hi-IN"/></w:rPr><w:t xml:space="preserve"></w:t></w:r>')
    r.find(qn('w:t')).text = text
    return r


def _sym_run(font, char, half_pts, color):
    return parse_xml(
        f'<w:r {W_NS}><w:rPr><w:color w:val="{color}"/><w:sz w:val="{half_pts}"/>'
        f'<w:szCs w:val="{half_pts}"/></w:rPr><w:sym w:font="{font}" w:char="{char}"/></w:r>')


def _tab_run():
    return parse_xml(f'<w:r {W_NS}><w:tab/></w:r>')


def _fill_title(anchor, fonts):
    tb = _txbx(anchor)
    _clear(tb)
    font = fonts["title"]
    exact = font == TITLE_CHAIN[0]
    # Rozha One is a display face with no bold; the Nirmala fallback keeps the
    # template's bold and condenses to fit the band.
    #
    # The fallback numbers are MEASURED, not guessed. `lineRule="exact"` with a
    # line shorter than the face needs lets the glyphs overflow the box
    # UPWARDS: the old 1062/50pt pair put the title at y=71.2 - 20pt above the
    # box - straight through the PW logo, which is what the client saw on a
    # machine without Rozha One. 1600/48pt/72% lands the fallback at y=94.9
    # against Rozha One's 95.2, clearing the logo (90.9) and the date pill
    # (165.1) either way.
    line, half_pts, width = (1500, 95, None) if exact else (1600, 96, 72)
    p = parse_xml(
        f'<w:p {W_NS}><w:pPr><w:spacing w:before="0" w:after="0" '
        f'w:line="{line}" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr></w:p>')
    p.append(_run(TITLE_TEXT, font, half_pts, WHITE, bold=not exact, w_pct=width))
    tb.append(p)


def _fill_date(anchor, date, fonts):
    tb = _txbx(anchor)
    _clear(tb)
    p = parse_xml(
        f'<w:p {W_NS}><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
        f'w:lineRule="auto"/><w:jc w:val="center"/></w:pPr></w:p>')
    if date:
        exact = fonts["text"] == TEXT_CHAIN[0]
        p.append(_run(date, fonts["text"], 38, INK, bold=True, w_pct=None if exact else 92))
    tb.append(p)


def _index_para(kind, sz, line):
    left, hang = IND_TOPIC if kind == "topic" else IND_SHORT
    after = 26 if kind == "topic" else 10
    return parse_xml(
        f'<w:p {W_NS}><w:pPr>'
        f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{TAB_RIGHT_TW}"/></w:tabs>'
        f'<w:spacing w:before="0" w:after="{after}" w:line="{line}" w:lineRule="exact"/>'
        f'<w:ind w:left="{left}" w:right="{IND_RIGHT}" w:hanging="{hang}"/>'
        f'<w:jc w:val="left"/>'
        f'<w:rPr><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        f'</w:pPr></w:p>')


def _fill_index(anchor, index, bar, shorts, numbers, sizes, fonts):
    """The index box: red centred heading, ticks for the topics, a tick for
    the news bar with the shorts as arrows beneath, dot leaders and page
    numbers on every line."""
    tb = _txbx(anchor)
    _clear(tb)
    sz, line = sizes
    text_font = fonts["text"]
    head = parse_xml(
        f'<w:p {W_NS}><w:pPr><w:spacing w:before="0" w:after="110" w:line="400" '
        f'w:lineRule="exact"/><w:ind w:left="180"/><w:jc w:val="center"/></w:pPr></w:p>')
    head.append(_run(INDEX_HEADING, fonts["heading"], 26, RED, bold=True))
    tb.append(head)

    def add(kind, text):
        p = _index_para(kind, sz, line)
        font, char = TICK if kind == "topic" else ARROW
        p.append(_sym_run(font, char, 14, RED))
        p.append(_tab_run())
        p.append(_run(text, text_font, sz, INK))
        p.append(_tab_run())
        num = numbers.get(text)
        p.append(_run(str(num) if num else "", text_font, sz, INK))
        tb.append(p)

    for t in index:
        add("topic", t)
    if shorts or bar:
        add("topic", bar or SHORTS_HEADING)
    for s in shorts:
        add("short", s)


def _collapse(body):
    """One host paragraph holding every drawing; everything else goes. With
    page-anchored objects the paragraph flow is irrelevant, and a single
    paragraph can never spill onto a second page."""
    paras = [p for p in body if p.tag == qn('w:p')]
    host = paras[0]
    for p in paras[1:]:
        for r in list(p):
            if r.tag == qn('w:r') and (r.find('.//' + qn('w:drawing')) is not None
                                       or r.find('.//' + _q(NS_MC, "AlternateContent")) is not None):
                host.append(r)
        body.remove(p)
    ppr = host.find(qn('w:pPr'))
    if ppr is not None:
        host.remove(ppr)
    host.insert(0, parse_xml(
        f'<w:pPr {W_NS}><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
        f'<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr>'))
    return host


def _letter(doc):
    sect = doc.element.body.find(qn('w:sectPr'))
    pg = sect.find(qn('w:pgSz'))
    pg.set(qn('w:w'), "12240")
    pg.set(qn('w:h'), "15840")
    if pg.get(qn('w:code')) is not None:
        del pg.attrib[qn('w:code')]


def _add_watermark(doc, host):
    if not os.path.exists(WATERMARK):
        return
    rid, _img = doc.part.get_or_add_image(WATERMARK)
    x, y, w, h = WATERMARK_RECT
    host.append(parse_xml(
        f'<w:r {W_NS} xmlns:wp="{NS_WP}" xmlns:a="{NS_A}" xmlns:pic="{NS_PIC}" xmlns:r="{NS_R}">'
        f'<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{emu(x)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{emu(y)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{emu(w)}" cy="{emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:wrapNone/><wp:docPr id="901" name="Cover watermark"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="{NS_PIC}"><pic:pic>'
        f'<pic:nvPicPr><pic:cNvPr id="901" name="Cover watermark"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'))


def build_cover_docx(out_path, date, index, bar, shorts, numbers, sizes, fonts):
    doc = Document(TEMPLATE)
    body = doc.element.body
    _letter(doc)
    _drop_fallbacks(body)
    host = _collapse(body)

    a = _anchor_of(body, "Graphic 2")            # navy band
    _place(a, BAND)
    _solid_fill(_sppr(a), BAND_COLOR)
    _place(_anchor_of(body, "Group 2"), LOGO, resize=False)

    a = _anchor_of(body, "Textbox 9")            # title
    _place(a, TITLE_BOX)
    _no_fill_no_line(_sppr(a))
    _body_pr(a, lIns="0", tIns="0", rIns="0", bIns="0", anchor="t")
    _fill_title(a, fonts)

    a = _anchor_of(body, "Rectangle: Rounded Corners 4")
    _place(a, DATE_PILL)
    _white_fill(_sppr(a))
    _set_line(_sppr(a), "6350", INK)
    st = next(a.iter(_q(NS_WPS, "style")), None)
    if st is not None:
        st.getparent().remove(st)

    a = _anchor_of(body, "Text Box 5")           # date text
    _place(a, DATE_TEXT)
    _no_fill_no_line(_sppr(a))
    _body_pr(a, lIns="0", tIns="0", rIns="0", bIns="0", anchor="ctr")
    _fill_date(a, date, fonts)

    a = _anchor_of(body, "Text Box 1")           # index
    _place(a, INDEX_BOX)
    _no_fill_no_line(_sppr(a))
    _body_pr(a, lIns="91440", tIns="45720", rIns="91440", bIns="0", anchor="t")
    _fill_index(a, index, bar, shorts, numbers, sizes, fonts)

    a = _anchor_of(body, "Image 5")              # India Gate
    _place(a, PIC_GATE)
    _set_line(_sppr(a), FRAME_W_EMU, FRAME_COLOR)
    _place(_anchor_of(body, "Group 12"), PIC_GLOBE, resize=False)
    _place(_anchor_of(body, "Group 15"), PIC_PARL, resize=False)
    a = _anchor_of(body, "Image 19")             # skyline
    _place(a, PIC_SKY)
    _set_line(_sppr(a), FRAME_W_EMU, FRAME_COLOR)

    _add_watermark(doc, host)
    doc.save(out_path)
    return out_path


# ================================================================ measure / render

def _index_metrics(pdf_path):
    """(lowest text edge, bullet count) inside the index box on page 1."""
    import fitz
    d = fitz.open(pdf_path)
    page = d[0]
    bottom, bullets = 0.0, 0
    top = INDEX_BOX[1]
    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0:
            continue
        for line in b["lines"]:
            x0, y0, x1, y1 = line["bbox"]
            if y0 < top or x0 > 385 or y1 > INDEX_FLOOR + 60:
                continue
            text = "".join(sp["text"] for sp in line["spans"]).strip()
            if not text:
                continue
            bottom = max(bottom, y1)
            if any("wingdings" in sp["font"].lower() for sp in line["spans"]):
                bullets += 1
    d.close()
    return bottom, bullets


def render_cover(workdir, date, index, bar, shorts, numbers, fonts,
                 progress=lambda s, p: None, cancelled=lambda: False):
    os.makedirs(workdir, exist_ok=True)
    cover_docx = os.path.join(workdir, "_cover_hi.docx")
    cover_pdf = os.path.join(workdir, "_cover_hi.pdf")
    expected = len(index) + (1 + len(shorts) if (shorts or bar) else 0)
    used, fits = INDEX_STEPS[0], False
    for step, sizes in enumerate(INDEX_STEPS):
        if cancelled():
            raise InterruptedError("stopped")
        progress("Building cover page" + (" (fitting index)" if step else ""), 92)
        build_cover_docx(cover_docx, date, index, bar, shorts, numbers, sizes, fonts)
        docx_to_pdf(cover_docx, cover_pdf)
        used = sizes
        bottom, bullets = _index_metrics(cover_pdf)
        fits = bottom <= INDEX_FLOOR and bullets >= expected
        if fits:
            break
    return cover_pdf, cover_docx, used, fits


# ================================================================ attach

def attach_cover(raw_docx, fitted_docx, body_pdf, body_docx, workdir, first_number,
                 date_in=None, index_in=None, shorts_in=None,
                 progress=lambda s, p: None, cancelled=lambda: False, notes=None):
    """The whole stage. Returns True when a cover was added."""
    import fitz
    notes = notes if notes is not None else []
    index, bar, shorts, source = resolve(raw_docx, index_in, shorts_in)
    if not index and not shorts:
        notes.append("No cover page: the file has no 'Heading 2' topics and no "
                     "news table, and none were typed in.")
        return False
    date, parsed = hindi_date(date_in) if (date_in or "").strip() else ("", False)
    fonts, font_note = resolve_fonts()

    # page numbers, from Word, on the fitted document that made the PDF
    targets = list(index) + ([bar or SHORTS_HEADING] if (shorts or bar) else []) + list(shorts)
    numbers, warn = {}, []
    progress("Building cover page (page numbers)", 90)
    try:
        found, total = page_lookup(fitted_docx, targets, cancelled)
    except InterruptedError:
        raise
    except Exception as exc:
        found, total = {}, None
        warn.append(f"page numbers could not be read from Word ({type(exc).__name__})")
    with fitz.open(body_pdf) as d:
        pdf_pages = len(d)
    if total is not None and total != pdf_pages:
        warn.append(f"Word paginates the fitted file to {total} pages but the PDF "
                    f"has {pdf_pages}; index printed without page numbers")
    elif total is not None:
        numbers = {t: p + first_number - 1 for t, p in found.items()}
        missing = [t for t in targets if t not in numbers]
        if missing:
            warn.append("no page found for: " + "; ".join(f"'{m[:40]}'" for m in missing))

    cover_pdf, cover_docx, used, fits = render_cover(
        workdir, date, index, bar, shorts, numbers, fonts, progress, cancelled)
    _en.prepend_pdf(cover_pdf, body_pdf, body_pdf)
    if body_docx and os.path.exists(body_docx):
        try:
            _en.prepend_docx(cover_docx, body_docx, body_docx)
        except Exception as exc:                 # never lose the Word file
            warn.append(f"cover added to the PDF only; the Word companion could "
                        f"not take it ({type(exc).__name__}: {exc})")

    parts = []
    if not date:
        parts.append("date MISSING - the pill is empty, type it in the Date block")
    elif parsed:
        parts.append("date typed")
    else:
        parts.append(f"date used verbatim ('{date}' is not a recognised date)")
    parts.append(f"index {source['index']} ({len(index)})")
    parts.append(f"shorts {source['shorts']} ({len(shorts)})")
    if numbers:
        parts.append(f"page numbers from Word ({len(numbers)}/{len(targets)})")
    step = INDEX_STEPS.index(used)
    if step:
        parts.append(f"index reduced to {used[0] // 2}pt to fit the box")
    if not fits:
        parts.append("the index is still too long for the box - Word hides the "
                     "lines past the bottom, shorten the index")
    notes.append("Cover page added: " + ", ".join(parts) + ".")
    notes.append(font_note)
    for w in warn:
        notes.append("Cover page: " + w + ".")
    return True
