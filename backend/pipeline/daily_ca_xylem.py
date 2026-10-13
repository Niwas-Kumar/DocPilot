# -*- coding: utf-8 -*-
"""Daily Current Affairs Summary (English Xylem) - the PW OnlyIAS x Xylem IAS
co-branded edition of the English booklet.

Same raw file, same body design, different page furniture and front page.
This is a SEPARATE pipeline on purpose: it imports the English building
blocks (styler, fitter, Word conversion, cover machinery) but changes none
of them. What it adds is a "furnish" step that swaps the English page
furniture for Xylem's, taken from section 2 of the client's cover template:

    - green (#00B050) triple-rule page border
    - empty header (no watermark, no logo), top margin 42.5pt
    - footer strip: rounded green box, OnlyIAS logo, page number in a green
      ring (Impact, a PAGE field), Xylem IAS logo - no links
    - page numbering starts at 2 (the cover is page 1)
    - no promo last page; numbers are Word's own, not stamped into the PDF

Rendered with a filler page, that section of the template matches the
published body pages to 0.1pt, so the furniture is carried across from the
template rather than reconstructed.

Carrying a part is not enough on its own: its shapes and paragraphs resolve
against the DESTINATION document's theme and styles. Both are pinned on the
way in - paragraph formatting by `_pin_formatting`, theme-relative outline
widths by `_pin_line_widths` - so a raw file's own theme can never restyle
the OnlyIAS / Xylem footer strip.
"""
import copy
import os
import re
import shutil

from docx import Document
from docx.oxml import parse_xml
from lxml import etree
from docx.oxml.ns import qn

from .convert import docx_to_pdf
from .fitter import (fit_document, _EnglishSpec, _positioned_tables, _unfloat_box,
                     _norm_text, _para_text)
from .styler import restyle_docx
from . import cover as _en
from . import cover_xylem

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEMPLATE = cover_xylem.TEMPLATE            # section 2 = the body furniture
FIRST_PAGE_NUMBER = 2
W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

HF_TAGS = (qn('w:headerReference'), qn('w:footerReference'))
# schema order of a sectPr's children, needed to insert pgNumType correctly
SECT_ORDER = ['w:headerReference', 'w:footerReference', 'w:footnotePr', 'w:endnotePr',
              'w:type', 'w:pgSz', 'w:pgMar', 'w:paperSrc', 'w:pgBorders', 'w:lnNumType',
              'w:pgNumType', 'w:cols', 'w:formProt', 'w:vAlign', 'w:noEndnote',
              'w:titlePg', 'w:textDirection', 'w:bidi', 'w:rtlGutter', 'w:docGrid',
              'w:printerSettings', 'w:sectPrChange']


class _XylemSpec(_EnglishSpec):
    """The English fitter rules on the Xylem text area: top margin 42.5pt,
    and the footer strip (56.9pt from the page edge, 24.6pt tall) pushes
    the body text up to 710.5pt - measured on the published pages."""
    name = "xylem"
    content_top = 42.5
    content_bottom = 710.5


XYLEM_SPEC = _XylemSpec()


def output_name(input_stem: str) -> str:
    return f"{input_stem}_CA Analysis_Xylem_UPSC_English.pdf"


# ---------------------------------------------------------------- side boxes
# Two side boxes with too little body text between them collide: Word pushes
# the second one down but flows its anchor paragraph against the slot it
# first asked for, so the text prints straight through the box; once the
# fitter tightens the topic, the boxes overprint each other as well. The
# English fitter only knows "box pushes its anchor text to the next page";
# this stage checks the rendered geometry itself and takes the offending box
# back into the text flow at full width.

BLUE_RGB = (0.0, 0.439, 0.753)          # styler.BLUE, the box border colour
BOX_LEFT_MIN, BOX_LEFT_MAX = 250.0, 400.0   # a right-docked box's left edge


def _side_boxes(page):
    """[x0, y0, x1, y1] of every right-docked side box on the page, read from
    its left border segment (the header bar sits ~6pt above that border)."""
    boxes = []
    for dr in page.get_drawings():
        r = dr["rect"]
        f = dr.get("fill")
        if not f or abs(f[0] - BLUE_RGB[0]) > 0.03 or abs(f[2] - BLUE_RGB[2]) > 0.03:
            continue
        if r.width <= 2.5 and r.height >= 15 and BOX_LEFT_MIN < r.x0 < BOX_LEFT_MAX:
            boxes.append([r.x0, r.y0, 560.0, r.y1])
    # a border may be drawn in several pieces (one per row in some
    # exports): pieces on the same x that TOUCH are one box. Pieces that
    # overlap by more than a hair are two boxes drawn on top of each other
    # - exactly the failure this check exists for - and stay separate.
    boxes.sort(key=lambda b: (round(b[0]), b[1]))
    merged = []
    for b in boxes:
        if merged and abs(merged[-1][0] - b[0]) < 1.5 and                 merged[-1][3] - 3.0 <= b[1] <= merged[-1][3] + 3.0:
            merged[-1][3] = max(merged[-1][3], b[3])
        else:
            merged.append(b)
    for b in merged:
        b[1] -= 7.0                      # the header bar sits above the border
    merged.sort(key=lambda b: b[1])
    return merged


def box_collisions(pdf_path):
    """[(page_index, box_header_norm, reason)] for side boxes that body text
    runs through or that start inside another box."""
    import fitz
    found = []
    d = fitz.open(pdf_path)
    for pi, page in enumerate(d):
        boxes = _side_boxes(page)
        if not boxes:
            continue
        lines = []
        for b in page.get_text("dict")["blocks"]:
            if b.get("type") != 0:
                continue
            for ln in b["lines"]:
                t = "".join(sp["text"] for sp in ln["spans"]).strip()
                if t:
                    lines.append((ln["bbox"], t))
        stacked = [any(o is not b and o[1] < b[1] < o[3] - 3.0 for o in boxes)
                   for b in boxes]                # starts inside an earlier box
        for bi, (bx0, by0, bx1, by1) in enumerate(boxes):
            inside = sorted((bb[1], t) for bb, t in lines
                            if bx0 - 2 <= bb[0] <= bx0 + 20 and by0 - 2 <= bb[1] <= by0 + 24)
            header = _norm_text(inside[0][1]) if inside else ""
            crossed = any(bb[0] < bx0 - 3 and bb[2] > bx0 + 3
                          and by0 < (bb[1] + bb[3]) / 2 < by1 for bb, _ in lines)
            # a box that a LATER box overprints is a victim, not the cause:
            # fixing the later one usually clears the crossing as well
            victim = any(stacked[j] and boxes[j][1] > by0 for j in range(len(boxes)) if j != bi)
            if stacked[bi]:
                found.append((pi, header, "it overprints the box above"))
            elif crossed and not victim:
                found.append((pi, header, "body text runs through it"))
    d.close()
    return found


def unfloat_colliding(styled_path, collisions, notes):
    """Take every collided side box back into the flow. Returns how many."""
    if not collisions:
        return 0
    doc = Document(styled_path)
    body = doc.element.body
    blocks = [e for e in body if e.tag in (qn('w:p'), qn('w:tbl'))]
    wanted = {h: reason for _, h, reason in collisions if h}
    changed = 0
    for i in _positioned_tables(blocks):
        bar = next((p for p in blocks[i].iter(qn('w:p')) if _para_text(p)), None)
        if bar is None:
            continue
        key = _norm_text(_para_text(bar))
        if key in wanted and _unfloat_box(blocks[i]):
            changed += 1
            notes.append(f"Side box '{_para_text(bar)[:60]}' shown full width: "
                         f"{wanted[key]} when docked right (two side boxes too close).")
    if changed:
        doc.save(styled_path)
    return changed


def _reorder_sect(sect):
    rank = {qn(t): i for i, t in enumerate(SECT_ORDER)}
    kids = sorted(list(sect), key=lambda c: rank.get(c.tag, len(SECT_ORDER)))
    for c in kids:
        sect.remove(c)
    for c in kids:
        sect.append(c)


NS_A_URI = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


def _a(tag):
    return f"{{{NS_A_URI}}}{tag}"


def _template_line_widths():
    """The template theme's `lnStyleLst` widths, in EMU, indexed from 1.

    The footer strip's shapes carry no line width of their own - they say
    `<a:lnRef idx="2">`, "use theme line style 2". That is resolved against
    whatever theme the DESTINATION document has, and the raw typed files
    define slot 2 as 2.00pt where this template defines it as 1.00pt. The
    strip and the page-number ring therefore came out at double weight, which
    is visible against the client's sample. Read the real widths here so the
    carried shapes can be pinned to what the template intended.
    """
    import zipfile
    try:
        with zipfile.ZipFile(TEMPLATE) as z:
            theme = z.read("word/theme/theme1.xml").decode("utf-8")
    except Exception:
        return {}
    lst = re.search(r"<a:lnStyleLst>(.*?)</a:lnStyleLst>", theme, re.S)
    if not lst:
        return {}
    return {i: int(w) for i, w in enumerate(re.findall(r'<a:ln w="(\d+)"', lst.group(1)), 1)}


def _pin_line_widths(root, widths):
    """Give every theme-referencing outline in a carried part an explicit
    width, so it can never be re-resolved against the booklet's own theme.

    Only shapes that do NOT already state a width are touched, and each one
    gets exactly what its own `lnRef idx` means in the template's theme.
    """
    if not widths:
        return 0
    pinned = 0
    for sp in root.iter(f"{{{NS_WPS}}}wsp"):
        ln = next((e for e in sp.iter(_a("ln")) if e.getparent().tag.endswith("}spPr")), None)
        if ln is None or ln.get("w"):
            continue
        ref = next((e for e in sp.iter(_a("lnRef"))), None)
        try:
            idx = int(ref.get("idx")) if ref is not None else 0
        except (TypeError, ValueError):
            idx = 0
        w = widths.get(idx)
        if w:
            ln.set("w", str(w))
            pinned += 1
    return pinned


def _pin_formatting(root, header):
    """The carried parts' paragraphs use the template's Header/Footer styles
    (Times New Roman 11, spacing 0/0, single). The styled document has no
    such styles, so they fell back to its Normal (Arial, line 276) and the
    footer strip rose 1.9pt while the header grew enough to push the body
    3.8pt down. Make the template's effective formatting explicit."""
    for p in root.iter(qn('w:p')):
        ppr = p.find(qn('w:pPr'))
        if ppr is None:
            ppr = parse_xml(f'<w:pPr {W_NS}/>')
            p.insert(0, ppr)
        for old in ppr.findall(qn('w:spacing')):
            ppr.remove(old)
        spacing = parse_xml(
            f'<w:spacing {W_NS} w:before="0" w:after="0" '
            + ('w:line="20" w:lineRule="exact"/>' if header else 'w:line="240" w:lineRule="auto"/>'))
        anchor = ppr.find(qn('w:rPr'))
        if anchor is not None:
            anchor.addprevious(spacing)
        else:
            ppr.append(spacing)
        targets = [ppr.find(qn('w:rPr'))] + [r.find(qn('w:rPr')) for r in p.findall(qn('w:r'))]
        holders = [ppr] + list(p.findall(qn('w:r')))
        for holder, rpr in zip(holders, targets):
            if rpr is None:
                rpr = parse_xml(f'<w:rPr {W_NS}/>')
                holder.insert(0, rpr) if holder is not ppr else holder.append(rpr)
            if rpr.find(qn('w:rFonts')) is None:
                rpr.insert(0, parse_xml(
                    f'<w:rFonts {W_NS} w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
                    f'w:cs="Times New Roman"/>'))
            for tag in ('w:sz', 'w:szCs'):
                if rpr.find(qn(tag)) is None:
                    el = parse_xml(f'<{tag} {W_NS} w:val="{2 if header else 22}"/>')
                    rpr.append(el)


def furnish(styled_path):
    """Replace the English page furniture in the styled docx with Xylem's.

    Every section of the styled file gets the template's section-2
    header/footer parts (cloned in under fresh partnames), its page border,
    margins and header/footer distances, numbering from 2, and no separate
    first page. The English header/footer parts are unlinked so they leave
    the package. Raises if the furniture could not be carried - a Xylem
    booklet with OnlyIAS furniture must never be produced silently."""
    doc = Document(styled_path)
    tpl = Document(TEMPLATE)
    line_widths = _template_line_widths()
    src = tpl.element.body.find(qn('w:sectPr'))          # the last = section 2
    if src is None or src.find(qn('w:pgBorders')) is None:
        raise RuntimeError("Xylem template has no body-furniture section")

    for sect in list(doc.element.body.iter(qn('w:sectPr'))):
        # 1) unlink the English header/footer parts
        for ref in [c for c in sect if c.tag in HF_TAGS]:
            rid = ref.get(qn('r:id'))
            sect.remove(ref)
            try:
                doc.part.drop_rel(rid)
            except Exception:
                pass
        # 2) the Xylem parts, re-homed into this package. Only the DEFAULT
        #    header/footer: the template's section also names a first-page
        #    pair (a 7 MB leftover watermark and the English link buttons)
        #    that Word would never show here and that must not travel
        refs = [copy.deepcopy(c) for c in src
                if c.tag in HF_TAGS and c.get(qn('w:type')) == 'default']
        for i, r in enumerate(refs):
            sect.insert(i, r)
        _en._carry_hf_parts(doc, tpl, sect)
        carried = [c for c in sect if c.tag in HF_TAGS]
        if len(carried) != len(refs):
            raise RuntimeError("Xylem header/footer parts could not be carried")
        for ref in carried:
            # the clones are plain opc Parts (a blob, no element tree)
            part = doc.part.related_parts[ref.get(qn('r:id'))]
            root = etree.fromstring(part.blob)
            _pin_formatting(root, header=ref.tag == qn('w:headerReference'))
            _pin_line_widths(root, line_widths)
            part._blob = etree.tostring(root, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)
        # 3) border, margins, distances
        old = sect.find(qn('w:pgBorders'))
        if old is not None:
            sect.remove(old)
        sect.append(copy.deepcopy(src.find(qn('w:pgBorders'))))
        pg_src = src.find(qn('w:pgMar'))
        pg = sect.find(qn('w:pgMar'))
        if pg is None:
            pg = parse_xml(f'<w:pgMar {W_NS}/>')
            sect.append(pg)
        for k in ('w:top', 'w:right', 'w:bottom', 'w:left', 'w:header', 'w:footer', 'w:gutter'):
            v = pg_src.get(qn(k))
            if v is not None:
                pg.set(qn(k), v)
        # 4) numbering from 2, single header for every page
        for t in sect.findall(qn('w:titlePg')):
            sect.remove(t)
        num = sect.find(qn('w:pgNumType'))
        if num is None:
            num = parse_xml(f'<w:pgNumType {W_NS}/>')
            sect.append(num)
        num.set(qn('w:start'), str(FIRST_PAGE_NUMBER))
        _reorder_sect(sect)
    doc.save(styled_path)
    return styled_path


def process(input_docx: str, workdir: str, variant: str = "English",
            progress=lambda stage, pct: None, cancelled=lambda: False,
            cover_date=None, cover_index=None, cover_shorts=None, **kw):
    """Raw typed English CA docx -> Xylem-branded PDF (+ Word companion).
    Returns ``(final_pdf, notes)``."""
    os.makedirs(workdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(input_docx))[0]
    notes = []

    progress("Styling document", 20)
    styled = os.path.join(workdir, f"{stem}_styled_xylem.docx")
    restyle_docx(input_docx, styled, banner_text="Topics")
    if cancelled():
        raise InterruptedError("stopped")

    progress("Applying Xylem page furniture", 30)
    furnish(styled)
    notes.append("Xylem furniture applied: green page border, footer strip with "
                 "OnlyIAS / Xylem IAS logos and ringed page numbers, no watermark, "
                 "no promo page.")
    if cancelled():
        raise InterruptedError("stopped")

    # side boxes that Word cannot lay out cleanly on this page geometry go
    # back into the flow BEFORE fitting (one cheap render), and once more
    # on the fitted result in case the tightening created a new collision
    progress("Checking side boxes", 34)
    probe = os.path.join(workdir, f"{stem}_xylem_probe.pdf")
    docx_to_pdf(styled, probe)
    unfloat_colliding(styled, box_collisions(probe), notes)
    if cancelled():
        raise InterruptedError("stopped")

    # measured pagination on the Xylem text area (same rules as English)
    progress("Fitting pages", 40)
    raw_pdf = fit_document(styled, workdir, spec=XYLEM_SPEC,
                           progress=progress, notes=notes)
    if cancelled():
        raise InterruptedError("stopped")
    if os.path.exists(raw_pdf) and unfloat_colliding(styled, box_collisions(raw_pdf), notes):
        progress("Fitting pages (after side-box fix)", 60)
        raw_pdf = fit_document(styled, workdir, spec=XYLEM_SPEC,
                               progress=progress, notes=notes)
        if cancelled():
            raise InterruptedError("stopped")

    if not os.path.exists(raw_pdf):
        progress("Converting to PDF (Word)", 70)
        raw_pdf = os.path.join(workdir, f"{stem}_raw_xylem.pdf")
        docx_to_pdf(styled, raw_pdf)
    if cancelled():
        raise InterruptedError("stopped")

    # no PDF stamping for this edition: the numbers are Word's PAGE field in
    # the footer strip, there are no footer links and no promo page
    progress("Finalizing", 85)
    final_pdf = os.path.join(workdir, output_name(stem))
    shutil.copy2(raw_pdf, final_pdf)
    companion = os.path.splitext(final_pdf)[0] + ".docx"
    try:
        shutil.copy2(styled, companion)
    except OSError:
        pass

    progress("Building cover page", 90)
    cover_xylem.attach_cover(input_docx, final_pdf, companion, workdir,
                             final_pdf, companion,
                             date_in=cover_date, index_in=cover_index,
                             shorts_in=cover_shorts, progress=progress, notes=notes)

    progress("Done", 100)
    return final_pdf, notes
