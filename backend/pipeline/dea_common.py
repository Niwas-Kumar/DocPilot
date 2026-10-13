# -*- coding: utf-8 -*-
"""Shared machinery for the Daily Editorial Analysis (DEA) editions.

Four categories are built on this: DEA English, DEA Hindi, and the two
combined "Daily CA + Daily Editorial" editions. Nothing here is imported by
the Daily CA or Test Series pipelines, which stay exactly as they are; this
module only READS their helpers.

What the raw DEA files look like (the samples of 17 Aug 2026, and what the
design has to survive on other days):

* No heading styles. A topic title is a plain paragraph marked only by its
  type - bold, red, 14-18pt. Heading styles are honoured when a typist does
  use them, but never required.
* A GS line under the title ("GS II: ...", "GS-III: ...", "Subject: ...").
* Section headings that are sometimes bold and sometimes not: the Hindi
  sample's "ONOE (...) के संबंध में चिंताएँ:" is plain text in the raw file and
  a heading in the published booklet. So a heading is recognised by its
  STRUCTURE too - a short line, no sentence ending, that introduces bullets.
* "Context:" / "संदर्भ:" / "Conclusion:" / "निष्कर्ष:" either on a line of
  their own or as a bold lead-in to the paragraph.
* A date line (style Title, "17 अगस्त, 2026") and a picture of the day's
  video QR code at the top. A file can hold SEVERAL days - the combined Hindi
  sample carries 15 Aug then 14 Aug - and the published booklet prints them
  oldest first, each day on a fresh page.

Nothing is ever invented: every glyph printed comes from the raw file or from
a field the user typed.
"""
import datetime as _dt
import io
import os
import re
import unicodedata

from docx.oxml import parse_xml
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ASSETS = os.path.join(ROOT, "assets", "dea")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

# ------------------------------------------------------------------ dates

MONTHS_EN = ["January", "February", "March", "April", "May", "June", "July",
             "August", "September", "October", "November", "December"]
DEV_DIGITS = str.maketrans("०१२३४५६७८९", "0123456789")


def _month_lookup():
    from .cover_hindi import MONTH_LOOKUP           # English + Hindi spellings
    return MONTH_LOOKUP


def parse_date(text):
    """A date written the ways the typists write it -> datetime.date, else
    None. '17th August 2026', 'August 17, 2026', '17 अगस्त, 2026',
    '14 अगस्त,2026', '17/08/2026', '2026-08-17'."""
    t = unicodedata.normalize("NFC", (text or "").strip()).translate(DEV_DIGITS)
    t = " ".join(t.replace(",", " , ").split())
    months = _month_lookup()

    def month_of(word):
        return months.get(word.lower()) or months.get(unicodedata.normalize("NFC", word))

    m = re.search(r"(\d{1,2})(?:st|nd|rd|th)?\s*(?:,\s*)?([A-Za-z]+|[ऀ-ॿ]+)\s*,?\s*(\d{4})", t, re.I)
    if m and month_of(m.group(2)):
        try:
            return _dt.date(int(m.group(3)), month_of(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    m = re.search(r"([A-Za-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?\s*,?\s*(\d{4})", t, re.I)
    if m and month_of(m.group(1)):
        try:
            return _dt.date(int(m.group(3)), month_of(m.group(1)), int(m.group(2)))
        except ValueError:
            pass
    m = re.search(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{4})\b", t)
    if m:
        try:
            return _dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    m = re.search(r"\b(\d{4})-(\d{1,2})-(\d{1,2})\b", t)
    if m:
        try:
            return _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    return None


WEEKDAYS = ("monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday",
            "सोमवार", "मंगलवार", "बुधवार", "गुरुवार", "शुक्रवार", "शनिवार", "रविवार")


def looks_like_date_line(text):
    """A paragraph that is a date and NOTHING else - a day marker. Once the
    day, month, year, ordinal, a weekday and punctuation are taken away no
    word may remain: "Published: 17 August 2026" inside an editorial is
    text, not the start of a new day."""
    t = " ".join((text or "").split())
    if not t or len(t) > 40 or parse_date(t) is None:
        return False
    months = _month_lookup()
    rest = unicodedata.normalize("NFC", t).translate(DEV_DIGITS).lower()
    for word in re.findall(r"[a-z]+|[ऀ-ॿ]+", rest):
        if word in months or word in WEEKDAYS or word in ("st", "nd", "rd", "th"):
            rest = rest.replace(word, " ", 1)
    return not re.search(r"[a-zऀ-ॿ]", re.sub(r"[\d\s,./\-:]", "", rest))


def english_date(d):
    return f"{d.day} {MONTHS_EN[d.month - 1]} {d.year}"


# ------------------------------------------------------------------ QR codes

def decode_qr(blob):
    """Text of the QR code in a picture, or None. The picture is tried as
    given and upscaled - typists paste anything from 120px to 1000px."""
    try:
        import zxingcpp
        from PIL import Image
        im = Image.open(io.BytesIO(blob)).convert("RGB")
    except Exception:
        return None
    for scale in (1, 2, 3):
        try:
            probe = im if scale == 1 else im.resize((im.width * scale, im.height * scale))
            for r in zxingcpp.read_barcodes(probe):
                if r.text:
                    return r.text
        except Exception:
            return None
    return None


def render_qr(text, path, px=900):
    """A crisp QR picture of `text` (black on white, 2-module quiet zone)."""
    import segno
    qr = segno.make(text, error="m")
    modules = qr.symbol_size(border=2)[0]
    qr.save(path, kind="png", scale=max(4, px // modules), border=2,
            dark="#000000", light="#FFFFFF")
    return path


def qr_from_link(link, path):
    """User typed a video link: render it. Returns (path, link) or (None, reason)."""
    link = (link or "").strip()
    if not link:
        return None, "empty"
    if not re.match(r"^(https?://|www\.)\S+$", link, re.I):
        return None, f"'{link[:60]}' is not a web link"
    return render_qr(link, path), link


# ------------------------------------------------------------------ paragraphs

def _runs(p_el):
    out = []
    for child in p_el:
        if child.tag == qn('w:r'):
            out.append(child)
        elif child.tag in (qn('w:hyperlink'), qn('w:smartTag'), qn('w:ins')):
            out.extend(ch for ch in child if ch.tag == qn('w:r'))
    return out


def run_text(r):
    parts = []
    for ch in r:
        if ch.tag == qn('w:t'):
            parts.append(ch.text or "")
        elif ch.tag == qn('w:tab'):
            parts.append("\t")
    return "".join(parts)


def para_text(p_el):
    return "".join(run_text(r) for r in _runs(p_el))


def clean(text):
    return " ".join((text or "").split())


def _on(el):
    return el is not None and el.get(qn('w:val')) not in ('0', 'false', 'off')


def run_bold(r, para_mark_bold=False):
    """w:b OR w:bCs - Devanagari is a complex script, and Word shows the
    complex-script bold for it; the raw Hindi files set both, but a file
    that only sets bCs must not read as regular."""
    rpr = r.find(qn('w:rPr'))
    if rpr is None:
        return para_mark_bold
    b, bcs = rpr.find(qn('w:b')), rpr.find(qn('w:bCs'))
    if b is None and bcs is None:
        return para_mark_bold
    return _on(b) or _on(bcs)


def _mark_bold(p_el):
    ppr = p_el.find(qn('w:pPr'))
    rpr = ppr.find(qn('w:rPr')) if ppr is not None else None
    return rpr is not None and (_on(rpr.find(qn('w:b'))) or _on(rpr.find(qn('w:bCs'))))


def fully_bold(p_el):
    seen = False
    for r in _runs(p_el):
        if not run_text(r).strip():
            continue
        seen = True
        if not run_bold(r):
            return False
    return seen


def lead_bold_text(p_el):
    """The bold text a paragraph opens with (its lead-in), or ''."""
    out = []
    for r in _runs(p_el):
        t = run_text(r)
        if not t.strip() and not out:
            continue
        if run_bold(r):
            out.append(t)
        else:
            break
    return "".join(out)


def run_size_pt(r, default=None):
    rpr = r.find(qn('w:rPr'))
    sz = rpr.find(qn('w:sz')) if rpr is not None else None
    if sz is None and rpr is not None:
        sz = rpr.find(qn('w:szCs'))
    try:
        return int(sz.get(qn('w:val'))) / 2.0 if sz is not None else default
    except (TypeError, ValueError):
        return default


def run_color(r):
    rpr = r.find(qn('w:rPr'))
    c = rpr.find(qn('w:color')) if rpr is not None else None
    v = c.get(qn('w:val')) if c is not None else None
    return v.upper() if v and re.fullmatch(r"[0-9A-Fa-f]{6}", v) else None


def is_red(hexcolor):
    if not hexcolor:
        return False
    r, g, b = int(hexcolor[0:2], 16), int(hexcolor[2:4], 16), int(hexcolor[4:6], 16)
    return r >= 0xB0 and g <= 0x50 and b <= 0x50


def has_drawing(el):
    return el.find('.//' + qn('w:drawing')) is not None or el.find('.//' + qn('w:pict')) is not None


def drawing_blobs(p_el, part):
    """Picture bytes of every drawing in a paragraph, in order."""
    out = []
    for blip in p_el.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
        rid = blip.get(qn('r:embed'))
        if not rid:
            continue
        try:
            out.append(part.related_parts[rid].blob)
        except KeyError:
            continue
    return out


# ------------------------------------------------------------------ roles

GS_RE = re.compile(
    r"^\s*(?:GS|G\.\s?S\.?|General\s+Studies)\s*[-‐-―:_.]*\s*(?:Paper\s*[-:]?\s*)?"
    r"(?:[IVX]+|\d+)\b", re.I)
SUBJECT_RE = re.compile(r"^\s*(?:Subject|Syllabus|Relevance|विषय|प्रासंगिकता|पाठ्यक्रम)\s*[:：]", re.I)
LEADIN_WORDS = (
    "Context", "Conclusion", "Introduction", "Background", "Why in News",
    "Why in the News", "Way Forward", "Summary",
    "संदर्भ", "सन्दर्भ", "निष्कर्ष", "पृष्ठभूमि", "परिचय", "प्रसंग", "चर्चा में क्यों",
    "आगे की राह",
)
LEADIN_RE = re.compile(r"^\s*(" + "|".join(re.escape(w) for w in LEADIN_WORDS) + r")\s*[:：]\s*", re.I)
SENTENCE_END_RE = re.compile(r"[.।!॥]\s*[\"'”’)\]]*\s*$")
LITERAL_BULLET_RE = re.compile(r"^\s*([•●▪■➢❖✓✔◦‣⁃\-\*•])\s+")


def style_name(doc, p_el):
    ppr = p_el.find(qn('w:pPr'))
    ps = ppr.find(qn('w:pStyle')) if ppr is not None else None
    if ps is None:
        return ""
    sid = ps.get(qn('w:val'))
    try:
        return (doc.styles.element.xpath(f'w:style[@w:styleId="{sid}"]/w:name')[0]
                .get(qn('w:val')) or sid).strip().lower()
    except Exception:
        return (sid or "").strip().lower()


def list_level(p_el, doc):
    """Source list level (0-based) from numPr or a list-bearing style."""
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        return None
    npr = ppr.find(qn('w:numPr'))
    if npr is not None:
        num = npr.find(qn('w:numId'))
        if num is not None and num.get(qn('w:val')) == "0":
            return None                          # numbering explicitly off
        il = npr.find(qn('w:ilvl'))
        try:
            return int(il.get(qn('w:val'))) if il is not None else 0
        except (TypeError, ValueError):
            return 0
    return None


def max_size(p_el, doc, default=11.0):
    sizes = [run_size_pt(r) for r in _runs(p_el) if run_text(r).strip()]
    sizes = [s for s in sizes if s]
    return max(sizes) if sizes else default


def topic_colour(p_el):
    """Majority colour of a paragraph's visible text."""
    counts = {}
    for r in _runs(p_el):
        t = run_text(r).strip()
        if t:
            c = run_color(r)
            counts[c] = counts.get(c, 0) + len(t)
    return max(counts, key=counts.get) if counts else None


def is_topic_title(p_el, doc, body_size):
    """A topic title: a heading style, or a bold line set in red or clearly
    larger than the body. Never a list item, never long prose."""
    text = clean(para_text(p_el))
    if not text or len(text) > 220 or list_level(p_el, doc) is not None:
        return False
    style = style_name(doc, p_el)
    if style in ("heading 1", "heading 2"):
        return True
    size = max_size(p_el, doc, body_size)
    red = is_red(topic_colour(p_el))
    if not fully_bold(p_el):
        # a typist who forgot the bold still marks a title by colour AND size
        return red and size >= body_size + 1.5 and not SENTENCE_END_RE.search(text)
    if red:
        return True
    return size >= body_size + 2.5


def is_gs_line(text):
    return bool(GS_RE.match(text) or SUBJECT_RE.match(text))


def is_heading(p_el, doc, next_is_list):
    """A section heading, by emphasis OR by structure (see module doc)."""
    text = clean(para_text(p_el))
    if not text or list_level(p_el, doc) is not None or len(text) > 140:
        return False
    if LEADIN_RE.match(text) and len(LEADIN_RE.sub("", text)) > 0:
        return False                              # a lead-in with its body
    if fully_bold(p_el) and not SENTENCE_END_RE.search(text):
        return True
    if SENTENCE_END_RE.search(text):
        return False
    # structure: a short line that introduces the bullets under it
    return len(text) <= 110 and (text.endswith((":", "?", "：")) or next_is_list)


def split_leadin(p_el, make_heading):
    """'Context: <body>' -> a heading paragraph ('Context:') inserted before
    the paragraph, and the lead-in words stripped from it. `make_heading(text)`
    builds the heading element. Returns the heading or None."""
    runs = _runs(p_el)
    full = "".join(run_text(r) for r in runs)
    m = LEADIN_RE.match(full)
    if m is None or not clean(full[m.end():]):
        return None
    # No bold test: the words in LEADIN_WORDS followed by a colon open a
    # section in this genre whatever their type - the combined Hindi raw file
    # types "संदर्भ: 2026 के ..." in plain text and the published booklet
    # still sets संदर्भ on its own heading line.
    word = m.group(1)
    head = make_heading(word + ":")
    p_el.addprevious(head)
    # strip the matched prefix across as many runs as it spans
    remaining = m.end()
    for r in runs:
        if remaining <= 0:
            break
        t = run_text(r)
        if len(t) <= remaining:
            remaining -= len(t)
            r.getparent().remove(r)
            continue
        keep = t[remaining:]
        remaining = 0
        for node in [c for c in r if c.tag in (qn('w:t'), qn('w:tab'))]:
            r.remove(node)
        node = r.makeelement(qn('w:t'), {})
        node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        node.text = keep.lstrip()
        r.append(node)
    return head


def strip_literal_bullet(p_el):
    """A typed '•'/'➢'/'-' bullet: remove the glyph, report True."""
    runs = _runs(p_el)
    for r in runs:
        t = run_text(r)
        if not t.strip():
            continue
        m = LITERAL_BULLET_RE.match(t)
        if not m:
            return False
        rest = t[m.end():]
        for node in [c for c in r if c.tag in (qn('w:t'), qn('w:tab'))]:
            r.remove(node)
        if rest:
            node = r.makeelement(qn('w:t'), {})
            node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            node.text = rest
            r.append(node)
        return True
    return False


def body_size_of(doc, blocks):
    """The typical body size of a raw file (the mode of its run sizes)."""
    counts = {}
    for el in blocks:
        if el.tag != qn('w:p'):
            continue
        for r in _runs(el):
            t = run_text(r).strip()
            s = run_size_pt(r)
            if t and s:
                counts[s] = counts.get(s, 0) + len(t)
    if not counts:
        return 11.0
    return max(counts, key=counts.get)


# ------------------------------------------------------------------ days

class Day:
    """One edition day inside a raw file: its date line, its video QR codes
    and its blocks (topics and everything under them)."""

    def __init__(self, date_text=""):
        self.date_text = date_text
        self.date = parse_date(date_text) if date_text else None
        self.qr = []                 # [(link_or_None, blob)]
        self.blocks = []

    @property
    def has_content(self):
        return any(clean(para_text(b)) if b.tag == qn('w:p') else True for b in self.blocks)


def split_days(doc, body):
    """Cut the body into Day objects. A day starts at a date line (style
    Title, or a paragraph that is a date and nothing else). The day's QR
    pictures (paragraphs holding only a picture that decodes as a QR code)
    are lifted out of the flow. Empty Title spacers are dropped. Returns the
    list of days in FILE order; the caller decides the print order."""
    days = [Day()]
    blocks = [el for el in body if el.tag in (qn('w:p'), qn('w:tbl'))]
    for el in blocks:
        if el.tag == qn('w:p'):
            text = clean(para_text(el))
            style = style_name(doc, el)
            if not text and not has_drawing(el):
                if style == "title":
                    el.getparent().remove(el)
                    continue
            if text and (style == "title" or looks_like_date_line(text)) and looks_like_date_line(text):
                day = Day(text)
                days.append(day)
                el.getparent().remove(el)
                continue
            if not text and has_drawing(el):
                blobs = drawing_blobs(el, doc.part)
                links = [decode_qr(b) for b in blobs]
                if blobs and all(links):
                    days[-1].qr.extend(zip(links, blobs))
                    el.getparent().remove(el)
                    continue
        days[-1].blocks.append(el)
    # a leading "day" with nothing in it is just the space before the first date
    return [d for d in days if d.blocks or d.qr or d.date_text]


def print_order(days):
    """Oldest first when every day carries a readable date (the published
    combined booklet prints 14 Aug before 15 Aug although the raw file has
    them the other way round); otherwise the file's own order."""
    dated = [d for d in days if d.has_content]
    if len(dated) >= 2 and all(d.date for d in dated):
        return sorted(days, key=lambda d: (d.date or _dt.date.max))
    return list(days)


def reassemble(body, days):
    """Put the days' blocks back into the body in `days` order, before the
    final sectPr. Returns the first block of every day after the first."""
    sect = body.find(qn('w:sectPr'))
    starts = []
    for i, day in enumerate(days):
        for j, el in enumerate(day.blocks):
            if sect is not None:
                sect.addprevious(el)
            else:
                body.append(el)
            if j == 0 and i > 0:
                starts.append(el)
    return starts


def break_before(el):
    """Start `el` on a new page (paragraph: its own property; table: a
    hairline spacer paragraph in front of it) - the fitter's helper, reused."""
    from .fitter import _ensure_page_break_before
    return _ensure_page_break_before(el)


PPR_ORDER = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr',
             'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd',
             'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct',
             'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd',
             'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
             'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment',
             'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr',
             'w:sectPr', 'w:pPrChange']
_PPR_RANK = {qn(t): i for i, t in enumerate(PPR_ORDER)}


def order_ppr(ppr):
    """Word silently drops pPr children that are out of schema order."""
    kids = list(ppr)
    kids.sort(key=lambda el: _PPR_RANK.get(el.tag, len(PPR_ORDER)))
    for el in kids:
        ppr.remove(el)
    for el in kids:
        ppr.append(el)


def ppr_of(p_el):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = p_el.makeelement(qn('w:pPr'), {})
        p_el.insert(0, ppr)
    return ppr


def set_flag(p_el, tag, on=True):
    ppr = ppr_of(p_el)
    el = ppr.find(qn(tag))
    if on and el is None:
        ppr.append(ppr.makeelement(qn(tag), {}))
    elif not on and el is not None:
        ppr.remove(el)
    order_ppr(ppr)


def set_jc(p_el, val):
    ppr = ppr_of(p_el)
    jc = ppr.find(qn('w:jc'))
    if jc is None:
        jc = ppr.makeelement(qn('w:jc'), {})
        ppr.append(jc)
    jc.set(qn('w:val'), val)
    order_ppr(ppr)


def drop_pstyle(p_el):
    """Formatting is explicit on every restyled paragraph, so the raw file's
    paragraph style is dropped: a Google Docs 'Heading 3' would otherwise
    bring its own spacing, keep-with-next chain and outline level along -
    and into another document when the two editions are merged."""
    ppr = p_el.find(qn('w:pPr'))
    if ppr is not None:
        ps = ppr.find(qn('w:pStyle'))
        if ps is not None:
            ppr.remove(ps)


# ------------------------------------------------------------------ PDF: space left on a page

def content_bottom(page, floor_y):
    """Lowest edge of the page's CONTENT - text, content pictures, table
    rules - ignoring the page furniture (full-page watermark, header logo,
    footer buttons, the Hindi centre watermark, the page border)."""
    import fitz
    bottom = 0.0
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            if not "".join(s["text"] for s in line["spans"]).strip():
                continue
            y0, y1 = line["bbox"][1], line["bbox"][3]
            if 60 < y0 and y1 < floor_y + 8:
                bottom = max(bottom, y1)
    for info in page.get_image_info():
        r = fitz.Rect(info["bbox"])
        if r.width > 550 and r.height > 650:
            continue                               # full-page watermark
        if abs(r.width - 447.9) < 3 and abs(r.height - 128.3) < 3:
            continue                               # Hindi centre watermark
        if r.y1 < 80 or r.y0 > floor_y:
            continue                               # logo / footer buttons
        bottom = max(bottom, r.y1)
    for dr in page.get_drawings():
        r = dr["rect"]
        if r.y0 < 60 or r.y1 > floor_y + 4 or r.x0 < 22 or r.x1 > 590:
            continue                               # page border, footer
        if r.width < 1 and r.height < 1:
            continue
        bottom = max(bottom, r.y1)
    return bottom


def place_banner(page, image_path, top_y, floor_y, natural_w, center_x,
                 min_scale=0.72, gap=12.0, *, min_top=0.0):
    """Drop a promo banner into the space between `top_y` (the content's
    bottom edge) and `floor_y` (just above the footer furniture).

    The banner is set at its natural width when it fits; when the space is
    shorter it is scaled down to the space - but never below `min_scale`,
    because these banners carry QR codes and small print that stop working
    below that. Returns the rectangle used, or None (the caller then puts
    the promo on the closing page instead). `min_top` reserves the header
    area even when the page has no content."""
    import fitz
    from PIL import Image
    with Image.open(image_path) as im:
        aspect = im.height / im.width
    w, h = natural_w, natural_w * aspect
    y0 = max(top_y + gap, min_top)
    avail = floor_y - y0
    if avail <= 0:
        return None
    scale = min(1.0, avail / h)
    if scale < min_scale:
        return None
    w, h = w * scale, h * scale
    rect = fitz.Rect(center_x - w / 2, y0, center_x + w / 2, y0 + h)
    page.insert_image(rect, filename=image_path)
    return rect


def append_page_keeping_group(dst, src_path):
    """Append every page of `src_path` to the open fitz document `dst`,
    keeping each page's transparency group (insert_pdf drops it, and the
    InDesign closing page blends in /DeviceCMYK - its colours shift without)."""
    import fitz
    with fitz.open(src_path) as src:
        start = len(dst)
        dst.insert_pdf(src)
        for i in range(len(src)):
            kind, val = src.xref_get_key(src.page_xref(i), "Group")
            if kind == "xref":
                val = src.xref_object(int(val.split()[0]), compressed=True)
            if kind in ("xref", "dict"):
                dst.xref_set_key(dst.page_xref(start + i), "Group", val)


def merge_docx(first_path, second_path, out_path):
    """The editable companion of a combined edition: the first part's
    document, then every block of the second part starting on a new page.
    Pictures, hyperlinks and list definitions come across with the blocks
    (docx_tools.copy_block_with_media); the page furniture is the first
    part's - both parts share one design per language."""
    from docx import Document
    from .docx_tools import copy_block_with_media
    first, second = Document(first_path), Document(second_path)
    blocks = [el for el in second.element.body if el.tag in (qn('w:p'), qn('w:tbl'))]
    copied = [copy_block_with_media(first, el, second) for el in blocks]
    if copied:
        break_before(copied[0])
    first.save(out_path)
    return out_path


def looks_like_ca(path, hindi=False):
    """How much a raw file looks like a DAILY CA file: its topic headings
    carry the heading style the CA stylers key on (English 'Heading 1',
    Hindi 'Heading 2'), where an editorial file marks titles by type only."""
    from docx import Document
    doc = Document(path)
    want = "heading 2" if hindi else "heading 1"
    n = 0
    for el in doc.element.body.iter(qn('w:p')):
        if any(a.tag == qn('w:tc') for a in el.iterancestors()):
            continue
        if style_name(doc, el) == want and clean(para_text(el)):
            n += 1
    return n


def order_pair(ca_path, dea_path, notes, hindi=False):
    """The two uploads of a combined edition, in the right slots. Files put
    in each other's slot are swapped (with a note) instead of producing a
    booklet whose halves are both wrong."""
    try:
        a, b = looks_like_ca(ca_path, hindi), looks_like_ca(dea_path, hindi)
    except Exception:
        return ca_path, dea_path
    if a == 0 and b > 0:
        notes.append("The two files looked swapped (the Daily Editorial file was in "
                     "the Daily CA slot) - used the other way round.")
        return dea_path, ca_path
    return ca_path, dea_path
