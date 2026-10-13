# -*- coding: utf-8 -*-
"""Front page for the Daily CA (English) booklet.

A separate stage, deliberately: the body pipeline (restyle -> fitter ->
finalize) is signed off and is not touched. This module fills the official
cover template, renders it through Word, and prepends that single page to
the finished PDF and to the Word companion.

The cover shows three things - the DATE in the orange pill, the INDEX (one
tick per topic) and the NEWS IN SHORTS items nested under a final "News in
Shorts" tick. Each of the three comes from the user's form field when it is
filled, otherwise from the raw file itself:

    date     <- the 'Title' paragraph  ("News In Shorts: 17 August 2026")
    index    <- the 'Heading 1' paragraphs, in document order
    shorts   <- the left column of the two-column news table, in order

That is the same detection the styler uses for the body, so the cover cannot
disagree with the pages behind it. The user's own text is used verbatim.

The index box is a fixed-height text box with autofit off (set to the
official cover's 303.6pt at build time), and Word HIDES whatever does not
fit, so a long index would lose its last lines silently. The cover is
therefore rendered and measured - both the lowest line and the number of
bullet lines actually on the page - and the two index sizes (tick / square)
are stepped down in lockstep only when the text really overflows.

The template's two continuous sections are kept as they are, and the Word
companion carries the cover's styles under `cov_` names: rendered untouched,
the template reproduces the official cover to the point, and every attempt
to "simplify" it moved the page onto the wrong furniture or styles.
"""
import copy
import os
import re

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from .convert import docx_to_pdf

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEMPLATE = os.path.join(ROOT, "assets", "cover", "daily_ca_english_cover.docx")

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

# --- geometry of the index text box, measured from a render of the template
INDEX_BOX_TOP_PT = 360.0                 # top edge of the box on the page
# The box is enlarged to the official cover's height at build time: the
# template ships a 279.9pt box, the published 17-Aug cover uses 303.6pt (its
# 11-line index runs to y=656.6pt, which the smaller box would hide).
INDEX_BOX_H_PT = 303.6
INDEX_BOX_INSET_PT = 3.6                 # bIns 45720 EMU
INDEX_BOX_BOTTOM_PT = INDEX_BOX_TOP_PT + INDEX_BOX_H_PT - INDEX_BOX_INSET_PT   # 660.0
# bullet glyphs as they come out of the rendered page (tick / square)
BULLET_GLYPHS = "\u2713\u2714\u25aa\u25a0\u2022\u25ab"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
VML_NS = "urn:schemas-microsoft-com:vml"
# official sizes (half-points): tick lines 16pt, square lines 15pt
INDEX_SIZE_STEPS = [(32, 30), (30, 28), (28, 26), (26, 24), (24, 22), (22, 20)]

SHORTS_HEADING = "News in Shorts"

ORDINAL_RE = re.compile(r"^(\d{1,2})(st|nd|rd|th)?\b(.*)$", re.I)
DATE_IN_TITLE_RE = re.compile(
    r"(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s+(\d{4})")


# ---------------------------------------------------------------- extraction

def _style_name(doc, p_el):
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


def _text(el):
    return " ".join("".join(t.text or "" for t in el.iter(qn('w:t'))).split())


def extract_from_raw(raw_docx):
    """(date, index_lines, shorts_lines) as found in the raw file; any of the
    three may be empty when the file does not carry it."""
    doc = Document(raw_docx)
    body = doc.element.body
    date, index, shorts = "", [], []
    for el in body.iter(qn('w:p')):
        style = _style_name(doc, el)
        text = _text(el)
        if not text:
            continue
        # the news grid's left cells are ALSO styled 'Heading 1' in the raw
        # files - those are the shorts, read from the table below, not topics
        in_table = any(anc.tag == qn('w:tc') for anc in el.iterancestors())
        if style == "heading 1" and not in_table:
            index.append(text)
        elif style == "title" and not date:
            m = DATE_IN_TITLE_RE.search(text)
            if m:
                date = f"{int(m.group(1))} {m.group(2)} {m.group(3)}"
    for tbl in body.iter(qn('w:tbl')):
        rows = tbl.findall(qn('w:tr'))
        if len(rows) < 2:
            continue
        if any(len(r.findall(qn('w:tc'))) != 2 for r in rows):
            continue
        for r in rows:
            left = _text(r.findall(qn('w:tc'))[0])
            if left:
                shorts.append(left)
        break                                  # the first N x 2 table is the grid
    return date, index, shorts


def _lines(value):
    """Form field -> clean list of lines (blank lines dropped)."""
    if not value:
        return []
    return [ln.strip() for ln in str(value).splitlines() if ln.strip()]


def resolve(raw_docx, date_in=None, index_in=None, shorts_in=None):
    """Apply the rule: a filled field wins, an empty one is extracted."""
    ex_date, ex_index, ex_shorts = extract_from_raw(raw_docx)
    date = (date_in or "").strip() or ex_date
    index = _lines(index_in) or ex_index
    shorts = _lines(shorts_in) or ex_shorts
    source = {
        "date": "typed" if (date_in or "").strip() else "extracted",
        "index": "typed" if _lines(index_in) else "extracted",
        "shorts": "typed" if _lines(shorts_in) else "extracted",
    }
    return date, index, shorts, source


# ---------------------------------------------------------------- date text

def date_runs(date):
    """'17 August 2026' -> [('17', False), ('th', True), (' August 2026', False)]
    - the ordinal becomes a superscript like the sample; a date typed with
    its own suffix keeps it (superscripted too), any other text is verbatim."""
    m = ORDINAL_RE.match(date.strip())
    if not m:
        return [(date.strip(), False)]
    day, suffix, rest = m.group(1), m.group(2), m.group(3)
    if not suffix:
        n = int(day)
        suffix = "th" if 11 <= n % 100 <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return [(day, False), (suffix.lower(), True), (rest, False)]


# ---------------------------------------------------------------- template

def _textboxes(body):
    """All text-box bodies in the template. Word writes each text box twice
    (the modern wps: branch and the legacy VML fallback), so a change must be
    applied to every copy or the two branches disagree."""
    return list(body.iter(qn('w:txbxContent')))


def _classify(tb):
    """'date' | 'index' | None by what the sample text in the box looks like."""
    paras = [p for p in tb.findall(qn('w:p')) if _text(p)]
    if not paras:
        return None
    if len(paras) == 1 and DATE_IN_TITLE_RE.search(_text(paras[0])):
        return "date"
    if len(paras) > 1:
        return "index"
    return None


def _clone_run(proto_run, text):
    r = copy.deepcopy(proto_run)
    for t in r.findall(qn('w:t')):
        r.remove(t)
    for br in r.findall(qn('w:br')):
        r.remove(br)
    node = r.makeelement(qn('w:t'), {})
    node.set(qn('xml:space'), 'preserve')
    node.text = text
    r.append(node)
    return r


def _force_size(container, half_pts):
    """Explicit size on a run's / paragraph mark's rPr (created if absent)."""
    for tag in ('w:sz', 'w:szCs'):
        el = container.find(qn(tag))
        if el is None:
            el = container.makeelement(qn(tag), {})
            container.append(el)
        el.set(qn('w:val'), str(half_pts))


def _fill_date(tb, date):
    paras = [p for p in tb.findall(qn('w:p')) if _text(p)]
    p = paras[0]
    proto = next((r for r in p.findall(qn('w:r')) if r.findall(qn('w:t'))), None)
    if proto is None:
        return
    for r in list(p.findall(qn('w:r'))):
        p.remove(r)
    for text, sup in date_runs(date):
        r = _clone_run(proto, text)
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            rpr = r.makeelement(qn('w:rPr'), {})
            r.insert(0, rpr)
        va = rpr.find(qn('w:vertAlign'))
        if sup:
            if va is None:
                va = rpr.makeelement(qn('w:vertAlign'), {})
                rpr.append(va)
            va.set(qn('w:val'), 'superscript')
        elif va is not None:
            rpr.remove(va)
        p.append(r)


def _fill_index(tb, index, shorts, sizes):
    """Rebuild the index box: one tick paragraph per index line, then the
    'News in Shorts' tick with the shorts as square items beneath it. The two
    paragraph prototypes come from the template itself, so the bullet glyphs,
    indents and fonts are exactly the official ones."""
    paras = tb.findall(qn('w:p'))
    tick_proto = square_proto = None
    for p in paras:
        style = p.find(qn('w:pPr') + '/' + qn('w:pStyle'))
        sid = style.get(qn('w:val')) if style is not None else ""
        if sid == "ListParagraph" and tick_proto is None and _text(p):
            tick_proto = p
        elif sid == "ListParagraph2" and square_proto is None and _text(p):
            square_proto = p
    if tick_proto is None:
        return False
    if square_proto is None:
        square_proto = tick_proto

    def make(proto, text, half_pts):
        p = copy.deepcopy(proto)
        runs = [r for r in p.findall(qn('w:r'))]
        proto_run = next((r for r in runs if r.findall(qn('w:t'))),
                         runs[0] if runs else None)
        for r in runs:
            p.remove(r)
        if proto_run is not None:
            p.append(_clone_run(proto_run, text))
        # an explicit size on the runs AND the paragraph mark, so the line
        # box follows the step even though the style says 16pt
        for r in p.findall(qn('w:r')):
            rpr = r.find(qn('w:rPr'))
            if rpr is None:
                rpr = r.makeelement(qn('w:rPr'), {})
                r.insert(0, rpr)
            _force_size(rpr, half_pts)
        ppr = p.find(qn('w:pPr'))
        if ppr is not None:
            mrpr = ppr.find(qn('w:rPr'))
            if mrpr is None:
                mrpr = ppr.makeelement(qn('w:rPr'), {})
                ppr.append(mrpr)
            _force_size(mrpr, half_pts)
        return p

    tick_sz, square_sz = sizes
    for p in paras:
        tb.remove(p)
    for line in index:
        tb.append(make(tick_proto, line, tick_sz))
    if shorts:
        tb.append(make(tick_proto, SHORTS_HEADING, tick_sz))
        for line in shorts:
            tb.append(make(square_proto, line, square_sz))
    return True


def _size_index_box(tb, height_pt):
    """Set the height of the text box that holds `tb`, in whichever branch
    this copy of the box lives (the wps: drawing or the VML fallback)."""
    emu = str(int(round(height_pt * 12700)))
    for anc in tb.iterancestors():
        if anc.tag == f"{{{WPS_NS}}}wsp":
            for ext in anc.iter(qn('a:ext')):
                if ext.get('cy') is not None:
                    ext.set('cy', emu)
        elif anc.tag in (qn('wp:anchor'), qn('wp:inline')):
            ext = anc.find(qn('wp:extent'))
            if ext is not None:
                ext.set('cy', emu)
        elif anc.tag == f"{{{VML_NS}}}shape":
            style = anc.get('style') or ""
            anc.set('style', re.sub(r"height:[\d.]+pt", f"height:{height_pt}pt", style))


def _strip_page_fields(doc):
    """The template's footers carry a PAGE field, but the published cover
    prints no page number (the body starts at 2). Remove the field's runs -
    begin/instr/separate/result/end - leaving the button drawings intact."""
    for section in doc.sections:
        for part in (section.footer, section.first_page_footer,
                     section.header, section.first_page_header):
            try:
                root = part._element
            except Exception:
                continue
            for p in root.iter(qn('w:p')):
                runs = list(p.findall(qn('w:r')))
                inside, drop = False, []
                for r in runs:
                    fc = r.find(qn('w:fldChar'))
                    kind = fc.get(qn('w:fldCharType')) if fc is not None else None
                    if kind == 'begin':
                        inside = True
                    if inside:
                        drop.append(r)
                    if kind == 'end':
                        inside = False
                for r in drop:
                    p.remove(r)
                for fs in list(p.findall(qn('w:fldSimple'))):
                    p.remove(fs)


def build_cover_docx(out_path, date, index, shorts, sizes):
    """Fill the template. Its two CONTINUOUS sections are left exactly as
    they are: rendered untouched, the template reproduces the official cover
    to the point (box top 360.0pt, buttons 715.1pt, no logo, no page
    border). Collapsing the sections put the page on the second section's
    furniture - logo header, blue page border, everything 5pt lower."""
    doc = Document(TEMPLATE)
    _strip_page_fields(doc)
    body = doc.element.body
    filled = {"date": 0, "index": 0}
    for tb in _textboxes(body):
        kind = _classify(tb)
        if kind == "date":
            # always: with no date at all the box is EMPTIED, never left
            # showing the template's sample date
            _fill_date(tb, date)
            filled["date"] += 1 if date else 0
        elif kind == "index":
            _size_index_box(tb, INDEX_BOX_H_PT)
            if _fill_index(tb, index, shorts, sizes):
                filled["index"] += 1
    doc.save(out_path)
    return filled


# ---------------------------------------------------------------- measure

def _index_metrics(pdf_path):
    """(lowest text edge, number of bullet lines) inside the index box region
    of the rendered page. The count matters as much as the edge: Word HIDES
    whatever overflows a fixed text box, so an index one line too long
    measures as fitting - the hidden line is simply not there to measure."""
    import fitz
    d = fitz.open(pdf_path)
    page = d[0]
    bottom, bullets = 0.0, 0
    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0:
            continue
        for line in b["lines"]:
            y0, y1 = line["bbox"][1], line["bbox"][3]
            if y0 < INDEX_BOX_TOP_PT - 4 or y1 > 700:   # below the date, above the buttons
                continue
            text = "".join(sp["text"] for sp in line["spans"]).strip()
            if not text:
                continue
            bottom = max(bottom, y1)
            if text[0] in BULLET_GLYPHS:
                bullets += 1
    d.close()
    return bottom, bullets


def render_cover(workdir, date, index, shorts, progress=lambda s, p: None):
    """Fill, render, and shrink the index only if it really overflows.
    Returns (cover_pdf, cover_docx, size_pair_used, fits)."""
    os.makedirs(workdir, exist_ok=True)
    cover_docx = os.path.join(workdir, "_cover.docx")
    cover_pdf = os.path.join(workdir, "_cover.pdf")
    expected = len(index) + ((1 + len(shorts)) if shorts else 0)
    used, fits = INDEX_SIZE_STEPS[0], False
    for step, sizes in enumerate(INDEX_SIZE_STEPS):
        progress("Building cover page" + (" (fitting index)" if step else ""), 88)
        build_cover_docx(cover_docx, date, index, shorts, sizes)
        docx_to_pdf(cover_docx, cover_pdf)
        used = sizes
        bottom, bullets = _index_metrics(cover_pdf)
        fits = bottom <= INDEX_BOX_BOTTOM_PT and bullets >= expected
        if fits:
            break
    return cover_pdf, cover_docx, used, fits


# ---------------------------------------------------------------- attach

def prepend_pdf(cover_pdf, body_pdf, out_pdf):
    """Cover page first, then the whole body. `out_pdf` may be the very same
    file as `body_pdf` (the pipeline rewrites its deliverable in place): the
    result is written beside it and swapped in only after the sources are
    closed, so the input is never overwritten while it is still open."""
    import fitz
    tmp = out_pdf + ".cover.tmp"
    out = fitz.open()
    cov = fitz.open(cover_pdf)
    body = fitz.open(body_pdf)
    try:
        out.insert_pdf(cov, from_page=0, to_page=0)
        out.insert_pdf(body)
        out.save(tmp, garbage=3, deflate=True)
    finally:
        out.close()
        cov.close()
        body.close()
    os.replace(tmp, out_pdf)
    return out_pdf


def prepend_docx(cover_docx, body_docx, out_docx):
    """Word companion: the cover's two sections first, then the body.

    Assembled the simple, robust way: the BODY document is opened and the
    cover's blocks are inserted at its top. The cover keeps both of its
    sections exactly as the template has them (the cover page draws its
    furniture from the first one, so it must survive), and the second one -
    the one that closes the cover - becomes a NEXT-PAGE break so the body
    starts on page 2 instead of flowing onto the cover. The body's own page
    furniture, numbering and pictures stay as they were; only the cover's
    pictures and header/footer parts are carried across.
    """
    from . import docx_tools as dt
    cover = Document(cover_docx)
    body = Document(body_docx)
    cbody = cover.element.body
    bbody = body.element.body
    seen = {}
    blocks = (qn('w:p'), qn('w:tbl'), qn('w:sdt'))
    _carry_styles(body, cover)

    first = next((e for e in bbody if e.tag in blocks), None)
    for el in [e for e in cbody if e.tag in blocks]:
        ppr = el.find(qn('w:pPr')) if el.tag == qn('w:p') else None
        sect = ppr.find(qn('w:sectPr')) if ppr is not None else None
        if sect is None:
            dt.copy_block_with_media(body, el, cover, before=first)
            continue
        # a section-break paragraph: copy it bare, then hang a re-homed copy
        # of its section on it - the header/footer ids belong to the cover
        # package and must not go through the picture re-linking
        bare = copy.deepcopy(el)
        bare.find(qn('w:pPr')).remove(bare.find(qn('w:pPr')).find(qn('w:sectPr')))
        new = dt.copy_block_with_media(body, bare, cover, before=first)
        carried = copy.deepcopy(sect)
        _carry_hf_parts(body, cover, carried, seen)
        new.find(qn('w:pPr')).append(carried)

    # the cover's closing section, kept CONTINUOUS as in the template: in
    # OOXML `w:type` says how the section it belongs to STARTS, and this
    # one holds the text-box paragraph - "nextPage" here sent the whole
    # index to page 2. It carries no furniture of its own (never displayed:
    # the cover page shows section 1's), so it inherits, which also keeps
    # the template's 11 MB promo picture out of the file.
    closing = copy.deepcopy(cbody.find(qn('w:sectPr')))
    for ref in [c for c in closing
                if c.tag in (qn('w:headerReference'), qn('w:footerReference'))]:
        closing.remove(ref)
    brk = parse_xml(f'<w:p {W_NS}><w:pPr></w:pPr></w:p>')
    brk.find(qn('w:pPr')).append(closing)
    if first is not None:
        first.addprevious(brk)
    else:
        bbody.append(brk)
    # the body must begin on a fresh page: its section has no w:type (the
    # default, next page); only an explicit "continuous" would flow it onto
    # the cover, and that one case is corrected here
    sects = list(bbody.iter(qn('w:sectPr')))
    for sp in sects[sects.index(closing) + 1:sects.index(closing) + 2]:
        typ = sp.find(qn('w:type'))
        if typ is not None and typ.get(qn('w:val')) == 'continuous':
            typ.set(qn('w:val'), 'nextPage')
    # same in-place guard as the PDF: never save over a package still open
    tmp = out_docx + ".cover.tmp"
    body.save(tmp)
    os.replace(tmp, out_docx)
    return out_docx


COV_PREFIX = "cov_"

# schema order of the children - Word DROPS properties that come out of order
STYLE_ORDER = ['w:name', 'w:aliases', 'w:basedOn', 'w:next', 'w:link',
               'w:autoRedefine', 'w:hidden', 'w:uiPriority', 'w:semiHidden',
               'w:unhideWhenUsed', 'w:qFormat', 'w:locked', 'w:personal',
               'w:personalCompose', 'w:personalReply', 'w:rsid', 'w:pPr',
               'w:rPr', 'w:tblPr', 'w:trPr', 'w:tcPr', 'w:tblStylePr']
PPR_ORDER = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore',
             'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers',
             'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku',
             'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE',
             'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid',
             'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
             'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment',
             'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle',
             'w:rPr', 'w:sectPr', 'w:pPrChange']
RPR_ORDER = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps',
             'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow',
             'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish',
             'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern',
             'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect',
             'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs',
             'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath']


def _reorder(el, order):
    """Stable re-sort of `el`'s children into schema order; extension
    elements (w14:...) keep their place at the end."""
    rank = {qn(t): i for i, t in enumerate(order)}
    kids = list(el)
    kids.sort(key=lambda c: rank.get(c.tag, len(order)))
    for c in kids:
        el.remove(c)
    for c in kids:
        el.append(c)


def _merge_under(props, defaults):
    """Add every child of `defaults` whose tag `props` does not carry."""
    have = {c.tag for c in props}
    for c in defaults:
        if c.tag not in have:
            props.append(copy.deepcopy(c))


def _carry_styles(dest_doc, src_doc):
    """Make the cover's paragraphs look in the body document exactly as they
    do in the cover document.

    Copied paragraphs keep their pStyle, but in the body that name resolves
    against the BODY's styles.xml - which has no ListParagraph (the tick
    bullets, 16pt, line 288), no ListParagraph2 (the squares), a different
    Heading 1 and Arial/276 document defaults. The result was an index
    without bullets on a tighter pitch and a date pill 27pt too low.

    So every paragraph and run of the cover (body, text boxes, headers and
    footers) is pointed at a `cov_`-prefixed copy of its style, and those
    copies - the whole basedOn chain - are added to the body's styles. The
    cover's document defaults are folded into each chain's root, because
    the body's defaults would otherwise show through underneath. List
    numbering named by a style is copied across the same way as it is for
    paragraphs. Must run BEFORE the blocks and header/footer parts are
    copied, since it edits the cover in memory."""
    from docx.opc.constants import CONTENT_TYPE as CT
    from . import docx_tools as dt

    src_styles = src_doc.styles.element
    dst_styles = dest_doc.styles.element
    by_id = {st.get(qn('w:styleId')): st
             for st in src_styles.findall(qn('w:style'))}
    default_para = next((st.get(qn('w:styleId')) for st in src_styles.findall(qn('w:style'))
                         if st.get(qn('w:type')) == 'paragraph' and st.get(qn('w:default')) == '1'),
                        'Normal')
    existing = {st.get(qn('w:styleId')) for st in dst_styles.findall(qn('w:style'))}

    roots = [src_doc.element.body]
    for part in src_doc.part.package.iter_parts():
        if part.content_type in (CT.WML_HEADER, CT.WML_FOOTER):
            roots.append(part.element)

    needed = set()
    for root in roots:
        for para in root.iter(qn('w:p')):
            ppr = para.find(qn('w:pPr'))
            if ppr is None:
                ppr = para.makeelement(qn('w:pPr'), {})
                para.insert(0, ppr)
            ps = ppr.find(qn('w:pStyle'))
            if ps is None:
                ps = ppr.makeelement(qn('w:pStyle'), {})
                ppr.insert(0, ps)              # pStyle is the first child
                sid = default_para
            else:
                sid = ps.get(qn('w:val'))
            if sid not in by_id:
                sid = default_para
            ps.set(qn('w:val'), COV_PREFIX + sid)
            needed.add(sid)
        for rs in root.iter(qn('w:rStyle')):
            sid = rs.get(qn('w:val'))
            if sid in by_id:
                rs.set(qn('w:val'), COV_PREFIX + sid)
                needed.add(sid)

    defaults = src_styles.find(qn('w:docDefaults'))
    ppr_def = rpr_def = None
    if defaults is not None:
        ppr_def = defaults.find(qn('w:pPrDefault') + '/' + qn('w:pPr'))
        rpr_def = defaults.find(qn('w:rPrDefault') + '/' + qn('w:rPr'))

    queue, done = list(needed), set()
    while queue:
        sid = queue.pop()
        if sid in done or sid not in by_id:
            continue
        done.add(sid)
        st = copy.deepcopy(by_id[sid])
        st.set(qn('w:styleId'), COV_PREFIX + sid)
        if st.get(qn('w:default')) is not None:
            del st.attrib[qn('w:default')]
        st.set(qn('w:customStyle'), '1')
        name = st.find(qn('w:name'))
        if name is not None:
            name.set(qn('w:val'), COV_PREFIX + (name.get(qn('w:val')) or sid))
        for tag in ('w:link', 'w:next', 'w:qFormat', 'w:uiPriority', 'w:semiHidden',
                    'w:unhideWhenUsed', 'w:rsid', 'w:autoRedefine', 'w:aliases'):
            for e in st.findall(qn(tag)):
                st.remove(e)
        based = st.find(qn('w:basedOn'))
        if based is not None and based.get(qn('w:val')) in by_id:
            queue.append(based.get(qn('w:val')))
            based.set(qn('w:val'), COV_PREFIX + based.get(qn('w:val')))
        else:
            if based is not None:
                st.remove(based)
            # root of a chain: the cover's document defaults sit under it
            if st.get(qn('w:type')) == 'paragraph' and ppr_def is not None:
                ppr = st.find(qn('w:pPr'))
                if ppr is None:
                    ppr = st.makeelement(qn('w:pPr'), {})
                    st.append(ppr)
                _merge_under(ppr, ppr_def)
                _reorder(ppr, PPR_ORDER)
            if rpr_def is not None:
                rpr = st.find(qn('w:rPr'))
                if rpr is None:
                    rpr = st.makeelement(qn('w:rPr'), {})
                    st.append(rpr)
                _merge_under(rpr, rpr_def)
                _reorder(rpr, RPR_ORDER)
            _reorder(st, STYLE_ORDER)
        for num in st.iter(qn('w:numId')):
            val = num.get(qn('w:val'))
            if val and val.isdigit():
                new = dt._copy_numbering_def(dest_doc, src_doc, int(val))
                if new is not None:
                    num.set(qn('w:val'), str(new))
        if COV_PREFIX + sid not in existing:
            dst_styles.append(st)
            existing.add(COV_PREFIX + sid)


def _carry_hf_parts(dest_doc, src_doc, sect_el, seen=None):
    """Re-home the cover's header/footer parts into the body package and
    point the section's references at the new ids.

    The parts are CLONED under fresh partnames rather than linked as-is:
    the cover's pictures are named image2/3/5.png just like the body's own
    media, and relating the original part would put two different files
    under one zip entry (the 'Duplicate name' warnings). Every part the
    footer references (its button images) is renamed the same way. A
    failure drops the reference, so the cover section inherits the body's
    furniture instead of breaking the file. `seen` (source partname ->
    clone) may be shared between calls so two sections referencing the same
    picture get one copy."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.opc.packuri import PackURI
    from docx.opc.part import Part
    pkg = dest_doc.part.package
    # Names handed out so far. A freshly cloned part is not reachable from
    # pkg.iter_parts() until its footer is related to the document, so two
    # pictures cloned for one footer would otherwise get the same name.
    taken = {str(p.partname) for p in pkg.iter_parts()}

    def fresh(template_name):
        stem, ext = template_name.rsplit(".", 1)
        stem = stem.rstrip("0123456789") or stem
        n = 1
        while f"{stem}{n}.{ext}" in taken:
            n += 1
        name = f"{stem}{n}.{ext}"
        taken.add(name)
        return PackURI(name)

    def clone(part, seen):
        if part.partname in seen:
            return seen[part.partname]
        new = Part(fresh(str(part.partname)), part.content_type, part.blob, pkg)
        seen[part.partname] = new
        for rel in list(part.rels.values()):
            if rel.is_external:
                new.rels.add_relationship(rel.reltype, rel.target_ref, rel.rId, True)
            else:
                target = clone(rel.target_part, seen)
                new.rels.add_relationship(rel.reltype, target, rel.rId)
        return new

    seen = {} if seen is None else seen
    for ref in list(sect_el):
        if ref.tag not in (qn('w:headerReference'), qn('w:footerReference')):
            continue
        try:
            rid = ref.get(qn('r:id'))
            part = src_doc.part.related_parts[rid]
            rtype = RT.HEADER if ref.tag == qn('w:headerReference') else RT.FOOTER
            new_rid = dest_doc.part.relate_to(clone(part, seen), rtype)
            ref.set(qn('r:id'), new_rid)
        except Exception:
            sect_el.remove(ref)


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
    prepend_pdf(cover_pdf, body_pdf, out_pdf)
    if body_docx and os.path.exists(body_docx):
        try:
            prepend_docx(cover_docx, body_docx, out_docx)
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
