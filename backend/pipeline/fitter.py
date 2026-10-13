# -*- coding: utf-8 -*-
"""Measured pagination pass for the Daily CA pipelines (English + Hindi).

The client's marked-up samples ask for three things a human designer does by
hand:

  * a page must not end in dead space while its topic continues overleaf
    ("Space" / "better it starts here"),
  * the last line or two of a topic must not spill onto a fresh page
    ("last line adjust in above page"),
  * a topic head must never sit stranded at the bottom of a page
    ("next page").

This pass measures the real Word pages (export -> PyMuPDF) and fixes only
what is wrong, exactly the way the designer would: nudge the text of the
offending topic down half a point at a time (11 -> 10.5 -> 10pt body, never
smaller), and keep heads glued to their first body line so Word carries a
head that does not fit to the next page on its own.

Measured rules, re-checked after every change (max MAX_ROUNDS cycles)
  W. page ends more than GAP_LIMIT short and the same topic continues on
     the next page             -> squeeze that topic one step
  O. page opens with the stray end of the previous structure (indent
     returns to a shallower level, or a fully-bold section lead sits right
     below other text)         -> squeeze that topic one step
  T. topic tail occupies less than TAIL_LIMIT at the top of its last page
                               -> squeeze that topic one step
  F. the page break was forced by a paragraph-anchored float that did not
     fit (Word drags the whole ANCHOR paragraph to the next page - dead
     space no squeeze can fix) -> move the float's run into the following
     paragraph; structural moves always settle before any squeezing

Every pass is scored (dead space + orphans + stranded heads + page count)
and the best measured layout wins - a change that reshuffles Word's page
breaks for the worse is simply not kept.

The two editions differ only in geometry and detectors, captured in a
FitSpec: the English design marks topics with red 16pt H1 paragraphs and
24pt banners; the Hindi design marks them with banner TABLES (red icon cell
+ peach title cell, or the full-red news bar) and floats a centred watermark
inside the content area that must not count as content.
"""
import os
import shutil
from collections import defaultdict

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from .convert import docx_to_pdf

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


GAP_LIMIT = 30.0        # pt of dead space at a page bottom worth fixing
TAIL_LIMIT = 110.0      # pt of topic tail on its last page worth pulling back
ORPHAN_LIMIT = 45.0     # pt of stray continuation above a structure reset
ROW_STRAND = 90.0       # a News-in-Shorts row head this close to the page
                        # foot, with its row continuing overleaf, is stranded
MAX_STEPS = 2           # per-topic squeeze steps (11 -> 10.5 -> 10pt body);
                        # a spec may override (Hindi body starts at 11.5pt,
                        # so it gets one step more down to the same floor)
MAX_ROUNDS = 26         # measure cycles; the loop stops early once no rule
                        # fires, so the cap only bites while work remains
                        # (structural moves apply one per cycle and can be
                        # reverted, so give the search enough room)
MAX_FLOAT_MOVES = 6     # per-float anchor forwarding steps
MIN_HALF_PTS = 20       # 10pt floor ("a decent level")
NO_SQUEEZE_MIN = 28     # runs at/above this size are titles - never scaled

# Every topic on a fresh page (client, 2026-09-08: Word companion stays
# editable; costs a white tail after each topic) - or topics flowing on as in
# the published samples (client, 2026-10-12: match the sample's page breaks).
# ONE switch for both Daily CA editions; flip it back to True to return to
# the 8 Sept layout. The squeeze rules below only have work when it is False.
TOPIC_PAGE_BREAK = False

# state scoring: every pass is measured and the BEST layout wins
PENALTY_ORPHAN = 45.0   # a page opening with a stray last line of a structure
PENALTY_STRAND = 70.0   # a topic head sitting in the bottom zone of a page
PENALTY_PAGE = 25.0     # per page - fewer, fuller pages are better
# a page carrying only a sliver (one line: "News Source: ...") is the worst
# break there is - it must outweigh the page it would cost to avoid it
PENALTY_SLIVER = 120.0
# a stranded shorts row head: must outweigh the dead band left behind when
# the row is moved whole (at most ROW_STRAND of waste)
PENALTY_ROW = 150.0


# ---------------------------------------------------------------- xml helpers

def _runs(p_el):
    out = []
    for child in p_el:
        if child.tag == qn('w:r'):
            out.append(child)
        elif child.tag == qn('w:hyperlink'):
            out.extend(ch for ch in child if ch.tag == qn('w:r'))
    return out


def _para_text(p_el):
    return "".join(t.text or "" for t in p_el.iter(qn('w:t'))).strip()


def _run_props(r):
    rpr = r.find(qn('w:rPr'))
    if rpr is None:
        return None, None
    sz = rpr.find(qn('w:sz'))
    color = rpr.find(qn('w:color'))
    try:
        size = int(sz.get(qn('w:val'))) if sz is not None else None
    except (TypeError, ValueError):
        size = None
    return size, (color.get(qn('w:val')) or "").upper() if color is not None else None


def _para_max_size(p_el):
    best = 0
    for r in _runs(p_el):
        size, _ = _run_props(r)
        if size:
            best = max(best, size)
    return best


def _has_color(p_el, hexcolor):
    for r in _runs(p_el):
        _, color = _run_props(r)
        if color == hexcolor:
            return True
    return False


def _get_ppr(p_el):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = p_el.makeelement(qn('w:pPr'), {})
        p_el.insert(0, ppr)
    return ppr


def _ensure_keep_next(p_el):
    """w:keepNext must sit at the top of w:pPr (after w:pStyle) or Word
    silently ignores it."""
    ppr = _get_ppr(p_el)
    if ppr.find(qn('w:keepNext')) is not None:
        return
    kn = ppr.makeelement(qn('w:keepNext'), {})
    ps = ppr.find(qn('w:pStyle'))
    if ps is not None:
        ps.addnext(kn)
    else:
        ppr.insert(0, kn)


def _blocks(body):
    return [el for el in body if el.tag in (qn('w:p'), qn('w:tbl'))]


# ---------------------------------------------------------------- fit specs

class _EnglishSpec:
    """Red 16pt H1 titles + 24pt 'Topics'/'News In Shorts' banners."""
    name = "english"
    content_top = 76.0
    content_bottom = 792.0 - 78.0
    drop_empties = True
    strip_keeps = True        # the source export's keep flags fight pagination
    max_steps = MAX_STEPS     # 11 -> 10.5 -> 10pt body
    # Every topic opens a new page (client, 2026-09-08). With this on the
    # squeeze rules stop firing by themselves - they all describe a topic
    # that continues overleaf, which can no longer happen - so the body keeps
    # its designed size and the Word file becomes editable: one text size,
    # one line height, and an edit reflows only its own topic.
    # Since 2026-10-12 the switch lives in TOPIC_PAGE_BREAK (off: flow).
    topic_page_break = TOPIC_PAGE_BREAK

    RED = "C00000"
    GREEN = "00B050"
    BLUE = "0070C0"
    RED_INT = int(RED, 16)
    H1_HALF_PTS = 32
    BANNER_HALF_PTS = 48

    def is_boundary_block(self, el):
        if el.tag != qn('w:p') or not _para_text(el):
            return False
        for r in _runs(el):
            size, color = _run_props(r)
            if size == self.BANNER_HALF_PTS:
                return True
            if size == self.H1_HALF_PTS and color == self.RED:
                return True
        return False

    def boundary_line(self, line):
        for span in line["spans"]:
            if not span["text"].strip():
                continue
            if 14.5 <= span["size"] <= 17.5 and span["color"] == self.RED_INT:
                return True
            if 22.0 <= span["size"] <= 26.0:
                return True
        return False

    def skip_image(self, r):
        # the full-bleed page watermark
        return r.width > 500 and r.height > 650

    def row_head_line(self, line):
        """A News-in-Shorts row title: red 11pt bold, centred in the left
        cell (x < 300). Nothing else in the design is red at that size."""
        if line["bbox"][0] >= 300:
            return False
        spans = [s for s in line["spans"] if s["text"].strip()]
        return bool(spans) and all(
            s["color"] == self.RED_INT and (s["flags"] & 16)
            and 10.0 <= s["size"] <= 11.9 for s in spans)

    def add_keeps(self, body):
        """Topic heads (H1 titles, Subject lines, blue section headings) must
        never strand at a page bottom.

        NEVER inside a table cell: w:keepNext there chains the row together
        and stops Word breaking it across pages, so a tall table jumps to the
        next page whole and leaves a dead band behind - the same trap the
        styler's in-cell pStyle strip exists for.
        """
        for p in body.iter(qn('w:p')):
            if not _para_text(p):
                continue
            if p.getparent().tag == qn('w:tc'):
                continue                        # see the docstring
            if self.is_boundary_block(p) or _has_color(p, self.GREEN) \
                    or _has_color(p, self.BLUE):
                _ensure_keep_next(p)


class _HindiSpec:
    """Banner tables (red icon cell + peach title, or the full-red news bar)
    mark the topics; the centred watermark floats INSIDE the content area."""
    name = "hindi"
    content_top = 69.0
    content_bottom = 792.0 - 74.0
    drop_empties = True       # the styler already drops them; harmless again
    strip_keeps = False       # the Hindi styler SETS its keeps deliberately
    max_steps = 3             # 11.5 -> 11 -> 10.5 -> 10pt body (same floor)
    topic_page_break = TOPIC_PAGE_BREAK   # see _EnglishSpec

    RED = "CB202F"
    RED_INT = int(RED, 16)
    WHITE_INT = 0xFFFFFF

    def is_boundary_block(self, el):
        # banner = small table (title row + optional 2pt spacer row) with a
        # red-filled cell; the >=3-column content grids also carry a red
        # HEADER row but are excluded by their column count
        if el.tag != qn('w:tbl'):
            return False
        rows = el.findall(qn('w:tr'))
        if not rows or len(rows) > 2:
            return False
        if max(len(tr.findall(qn('w:tc'))) for tr in rows) > 2:
            return False
        for shd in el.iter(qn('w:shd')):
            if (shd.get(qn('w:fill')) or "").upper() == self.RED:
                return True
        return False

    def boundary_line(self, line):
        # topic title: Baloo 13pt in brand red on the peach bar;
        # news bar: white 13pt centred on red. Nothing else is 13pt.
        for span in line["spans"]:
            if not span["text"].strip():
                continue
            if 12.2 <= span["size"] <= 13.9 and span["color"] in (
                    self.RED_INT, self.WHITE_INT):
                return True
        return False

    def skip_image(self, r):
        # centred watermark (448x128 at y~332 on every page)
        return 440 <= r.width <= 456 and 120 <= r.height <= 136

    def row_head_line(self, line):
        return False           # the Hindi news table has its own layout

    def add_keeps(self, body):
        pass                   # the styler's keepNext flags already cover heads


ENGLISH_SPEC = _EnglishSpec()
HINDI_SPEC = _HindiSpec()


# ---------------------------------------------------------------- static pass

def _drop_empty_paragraphs(body):
    """Every empty paragraph wastes a full line box. Keep separators between
    two adjacent tables (Word merges the tables without one), the final
    paragraph, and any paragraph carrying a drawing or a section break."""
    blocks = _blocks(body)
    dropped = 0
    for i, el in enumerate(blocks):
        if el.tag != qn('w:p'):
            continue
        if i == len(blocks) - 1:
            continue
        if _para_text(el):
            continue
        if next(el.iter(qn('w:drawing')), None) is not None:
            continue
        ppr = el.find(qn('w:pPr'))
        if ppr is not None and ppr.find(qn('w:sectPr')) is not None:
            continue
        prev_tbl = i > 0 and blocks[i - 1].tag == qn('w:tbl')
        next_tbl = i + 1 < len(blocks) and blocks[i + 1].tag == qn('w:tbl')
        if prev_tbl and next_tbl:
            continue
        el.getparent().remove(el)
        dropped += 1
    return dropped


# w:pPr children must appear in schema order or Word silently drops them;
# pageBreakBefore belongs right after the keep flags.
_PBB_AFTER = ('w:pStyle', 'w:keepNext', 'w:keepLines')


def _ensure_page_break_before(el):
    """Make the block `el` start a new page. Returns True if it changed.

    A paragraph carries `w:pageBreakBefore` itself. A TABLE cannot - Word has
    no such property for tables - so a hairline empty paragraph carrying the
    break is inserted ahead of it (2pt line, no spacing, so it costs nothing
    visible at the top of the page).
    """
    if el.tag == qn('w:tbl'):
        prev = el.getprevious()
        if prev is not None and prev.tag == qn('w:p') and not _para_text(prev):
            ppr = prev.find(qn('w:pPr'))
            if ppr is not None and ppr.find(qn('w:pageBreakBefore')) is not None:
                return False                      # already breaking here
        spacer = parse_xml(
            f'<w:p {NS}><w:pPr><w:pageBreakBefore/>'
            f'<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
            f'<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>')
        el.addprevious(spacer)
        return True
    if el.tag != qn('w:p'):
        return False
    ppr = el.find(qn('w:pPr'))
    if ppr is None:
        ppr = parse_xml(f'<w:pPr {NS}/>')
        el.insert(0, ppr)
    if ppr.find(qn('w:pageBreakBefore')) is not None:
        return False
    brk = parse_xml(f'<w:pageBreakBefore {NS}/>')
    anchor = None
    for tag in _PBB_AFTER:
        found = ppr.find(qn(tag))
        if found is not None:
            anchor = found
    if anchor is not None:
        anchor.addnext(brk)
    else:
        ppr.insert(0, brk)
    return True


def _break_between_topics(body, spec):
    """Start every topic after the first on a fresh page.

    Asked for by the client so the Word companion can be edited safely: with
    each topic on its own page an edit reflows only that topic instead of
    shunting every page after it, and a topic head can never strand at a page
    bottom. It costs the white space at the end of each topic - that is the
    trade the client chose.
    """
    blocks = _blocks(body)
    starts = [i for i, el in enumerate(blocks) if spec.is_boundary_block(el)]
    added = 0
    for prev, i in zip(starts, starts[1:]):       # never the first topic
        # Two boundaries can sit back to back - the English design opens with
        # a "Topics" banner immediately followed by the first topic title.
        # Breaking between them would strand the banner alone on a page, so a
        # break is only worth it when real content separates the two.
        if not any(_para_text(b) or b.tag == qn('w:tbl')
                   for b in blocks[prev + 1:i]):
            continue
        if _ensure_page_break_before(blocks[i]):
            added += 1
    return added


def _static_fixes(body, spec):
    dropped = _drop_empty_paragraphs(body) if spec.drop_empties else 0
    if spec.strip_keeps:
        for tag in ('w:keepNext', 'w:keepLines'):
            for el in list(body.iter(qn(tag))):
                el.getparent().remove(el)
    if getattr(spec, "topic_page_break", False):
        _break_between_topics(body, spec)
    _separate_stacked_boxes(body)
    spec.add_keeps(body)
    return dropped


# ---------------------------------------------------------------- measurement

def _measure(pdf_path, spec):
    """Per-page fill + topic boundary positions, from the rendered PDF.

    Returns (pages, bounds, floats):
      pages  - per page: {'last', 'first', 'toplines'}
      bounds - (page_index, y0) per topic boundary in reading order; a
               wrapped 2-line title is one boundary
      floats - (page_index, y0, w_pt, h_pt) per content illustration
    """
    import fitz

    top, bottom = spec.content_top, spec.content_bottom
    pages = []
    bounds = []
    floats = []
    doc = fitz.open(pdf_path)
    for pno, page in enumerate(doc):
        last = 0.0
        last_left = 0.0                  # fill of the LEFT column: a side box
        first = bottom                   # can fill the right edge-to-edge
        toplines = []                    # while the text column runs dry
        head_lines = []
        first_lines = []                 # (y0, text) - min y wins
        row_heads = []                   # (y0, text) News-in-Shorts row titles
        d = page.get_text("dict")
        for block in d["blocks"]:
            if block["type"] == 1:
                r = fitz.Rect(block["bbox"])
                if spec.skip_image(r):
                    continue
                if r.y0 < top - 6 or r.y1 > bottom + 6:
                    continue
                if r.width < 40 and r.height < 40:
                    continue
                last = max(last, r.y1)
                last_left = max(last_left, r.y1)   # pictures fill legitimately
                first = min(first, r.y0)
                if r.width >= 100:
                    floats.append((pno, r.y0, r.width, r.height))
                continue
            for line in block["lines"]:
                y0, y1 = line["bbox"][1], line["bbox"][3]
                if y0 < top - 6 or y1 > bottom + 8:
                    continue
                text = "".join(s["text"] for s in line["spans"]).strip()
                if not text:
                    continue
                last = max(last, y1)
                if line["bbox"][0] < 310:          # text in the main column
                    last_left = max(last_left, y1)  # (side-box text is x>=317)
                first = min(first, y0)
                if y0 <= top + 90:
                    first_lines.append((y0, text))
                if spec.boundary_line(line):
                    head_lines.append(y0)
                    continue
                if spec.row_head_line(line):
                    # a wrapped title is one head: keep its first line only
                    if not row_heads or y0 - row_heads[-1][0] > 30:
                        row_heads.append((y0, text))
                    else:
                        row_heads[-1] = (row_heads[-1][0],
                                         row_heads[-1][1] + " " + text)
                # top-zone lines feed the orphan rule: indent + boldness tell
                # whether the page opens with the stray end of the previous
                # structure (the bullet glyph is its own Wingdings span)
                if y0 <= top + ORPHAN_LIMIT + 40:
                    body_spans = [s for s in line["spans"] if s["text"].strip()
                                  and not s["font"].startswith("Wingdings")]
                    bold_all = bool(body_spans) and all(
                        (s["flags"] & 16) and s["size"] <= 13.5
                        for s in body_spans)
                    toplines.append((y0, line["bbox"][0], bold_all))
        pages.append({'last': last, 'last_left': last_left, 'first': first,
                      'first_text': min(first_lines)[1] if first_lines else "",
                      'toplines': sorted(toplines),
                      'row_heads': sorted(row_heads)})
        # merge the lines of one wrapped title into a single boundary
        for y in sorted(head_lines):
            if bounds and bounds[-1][0] == pno and 0 <= y - bounds[-1][1] < 30:
                continue
            bounds.append((pno, y))
    doc.close()
    return pages, bounds, floats


def _topic_ranges(blocks, spec):
    """[(start_block_idx, end_block_idx_exclusive)] per boundary, in order."""
    starts = [i for i, el in enumerate(blocks) if spec.is_boundary_block(el)]
    ranges = []
    for j, s in enumerate(starts):
        e = starts[j + 1] if j + 1 < len(starts) else len(blocks)
        ranges.append((s, e))
    return ranges


def _topic_at(bounds, page, y):
    """Index of the topic that owns position (page, y)."""
    k = -1
    for i, (bp, by) in enumerate(bounds):
        if (bp, by) <= (page, y):
            k = i
        else:
            break
    return k


def _orphan_pages(pages, bounds, spec):
    """Pages that open with the stray end of the previous structure: a line
    or two before the indent returns to a shallower level, or before a
    fully-bold section lead."""
    out = []
    for p in range(1, len(pages)):
        if any(bp == p and by < spec.content_top + 40 for bp, by in bounds):
            continue                             # fresh topic starts this page
        tl = pages[p]['toplines']
        if len(tl) < 2:
            continue
        y00, x00, _ = tl[0]
        for y, x, bold in tl[1:]:
            if y - spec.content_top > ORPHAN_LIMIT:
                break
            if x <= x00 - 8 or (bold and y > y00 + 4):
                out.append(p)
                break
    return out


def _sliver_pages(pages, bounds, spec):
    """Pages whose whole content is a sliver at the top - typically one
    stray line ("News Source: ...") that the previous page could not hold.
    The orphan rule cannot see these: it reads the indent of the SECOND
    line, and a one-line page has none (17-Aug combo, page 9)."""
    out = []
    for p in range(1, len(pages)):
        if any(bp == p for bp, by in bounds):
            continue                             # a topic starts on this page
        last = pages[p]['last']
        if 0 < last and last - spec.content_top < ORPHAN_LIMIT:
            out.append(p)
    return out


def _stranded_rows(pages, bounds, spec):
    """News-in-Shorts rows whose title sits in the page's bottom zone while
    the row itself runs on overleaf: Word split the row right after its
    head. Returns (page, title) for rows short enough to move whole - a row
    taller than a page has to split somewhere."""
    page_h = spec.content_bottom - spec.content_top
    out = []
    for p in range(len(pages) - 1):
        heads = pages[p]['row_heads']
        if not heads:
            continue
        y, text = heads[-1]
        if y < spec.content_bottom - ROW_STRAND:
            continue
        if any(bp == p + 1 and by < spec.content_top + 40 for bp, by in bounds):
            continue                             # the table ended with the page
        nxt = pages[p + 1]['row_heads']
        if not nxt:
            continue                             # row still running: too tall
        height = (spec.content_bottom - y) + (nxt[0][0] - spec.content_top)
        if height < page_h - 20:
            out.append((p, text))
    return out


def _score(pages, bounds, spec):
    """Lower is better: dead space + orphans + stranded heads + page count.
    Dead space is judged on the TEXT column: a side box filling the right
    edge does not excuse an empty left column."""
    waste = 0.0
    for p in range(len(pages) - 1):
        if pages[p]['last'] > 0:
            waste += max(0.0, spec.content_bottom - pages[p]['last_left'] - 8)
    stranded = sum(1 for bp, by in bounds if by > spec.content_bottom - 80)
    return (waste
            + PENALTY_ORPHAN * len(_orphan_pages(pages, bounds, spec))
            + PENALTY_SLIVER * len(_sliver_pages(pages, bounds, spec))
            + PENALTY_ROW * len(_stranded_rows(pages, bounds, spec))
            + PENALTY_STRAND * stranded
            + PENALTY_PAGE * len(pages))


# ---------------------------------------------------------------- floats

def _anchor_map(blocks):
    """Text-wrapped floats in document order: [(block_idx, run_el, w_pt, h_pt)].

    A paragraph-relative float drags its whole anchor paragraph to the next
    page when the float does not fit below it - the classic cause of a dead
    band at a page bottom. Forwarding the anchor into the next paragraph
    frees the text to flow up while the picture stays beside its section.
    """
    out = []
    for i, el in enumerate(blocks):
        if el.tag != qn('w:p'):
            continue
        for anchor in el.iter(qn('wp:anchor')):
            if anchor.find(qn('wp:wrapNone')) is not None:
                continue                        # page furniture
            ext = anchor.find(qn('wp:extent'))
            if ext is None:
                continue
            try:
                w = int(ext.get('cx')) / 12700.0
                h = int(ext.get('cy')) / 12700.0
            except (TypeError, ValueError):
                continue
            run = anchor
            while run is not None and run.tag != qn('w:r'):
                run = run.getparent()
            if run is not None:
                out.append((i, run, w, h))
    return out


def _match_anchor(anchors, pdf_floats, float_idx, fw, fh):
    """Map a rendered picture back to its wp:anchor: by position when the
    counts line up AND the size agrees (keeps duplicate sizes unambiguous,
    e.g. two 218pt-capped Hindi floats), by size otherwise. Inline pictures
    in table cells also land in pdf_floats, so a candidate may match no
    anchor at all - the caller must keep scanning."""
    if len(anchors) == len(pdf_floats):
        entry = anchors[float_idx]
        if abs(entry[2] - fw) <= 3 and abs(entry[3] - fh) <= 3:
            return float_idx, entry
    for ai, entry in enumerate(anchors):
        if abs(entry[2] - fw) <= 3 and abs(entry[3] - fh) <= 3:
            return ai, entry
    return None, None


def _shrink_float(entry, factor=0.88):
    """Scale a pushed picture down one step (~12%) - when it cannot be
    forwarded past a table, a slightly smaller picture that FITS the page
    beats a full-size one that drags its anchor overleaf."""
    _, run, _, _ = entry
    anchor = next(run.iter(qn('wp:anchor')), None)
    if anchor is None:
        return False
    ext = anchor.find(qn('wp:extent'))
    if ext is None:
        return False
    try:
        cx = int(ext.get('cx'))
        cy = int(ext.get('cy'))
    except (TypeError, ValueError):
        return False
    ncx, ncy = int(cx * factor), int(cy * factor)
    ext.set('cx', str(ncx))
    ext.set('cy', str(ncy))
    for a_ext in anchor.iter(qn('a:ext')):
        a_ext.set('cx', str(ncx))
        a_ext.set('cy', str(ncy))
    return True


def _forward_anchor(blocks, entry, spec):
    """Move a float's run into the next body paragraph. True on success."""
    idx, run, _, _ = entry
    j = idx + 1
    while j < len(blocks):
        el = blocks[j]
        if el.tag == qn('w:tbl') or spec.is_boundary_block(el):
            return False                        # never cross a table or a head
        if _para_text(el):
            break
        j += 1
    else:
        return False
    nxt = blocks[j]
    ppr = nxt.find(qn('w:pPr'))
    nxt.insert(1 if ppr is not None else 0, run)
    return True


# ---------------------------------------------------------------- side boxes

def _norm_text(s):
    """Alphanumerics only: the rendered line carries the Wingdings bullet
    glyph and tabs that the paragraph's XML text does not."""
    return "".join(ch for ch in s.lower() if ch.isalnum())[:20]


def _positioned_tables(blocks):
    """Block indexes of floating (side-box) tables, in document order."""
    out = []
    for i, el in enumerate(blocks):
        if el.tag != qn('w:tbl'):
            continue
        tblpr = el.find(qn('w:tblPr'))
        if tblpr is not None and tblpr.find(qn('w:tblpPr')) is not None:
            out.append(i)
    return out


def _box_lead_texts(blocks, i):
    """Texts that appear FIRST when the box at block i opens a page: its own
    header bar, the body paragraph that follows it in the flow (a floating
    table drags that anchor paragraph along with it), and the paragraph just
    BEFORE it (a keep-with-next heading there gets carried along too)."""
    texts = []
    bar = next((p for p in blocks[i].iter(qn('w:p')) if _para_text(p)), None)
    if bar is not None:
        texts.append(_para_text(bar))
    j = i + 1
    while j < len(blocks):
        el = blocks[j]
        if el.tag == qn('w:tbl'):
            break
        if _para_text(el):
            texts.append(_para_text(el))
            break
        j += 1
    j = i - 1
    while j >= 0:
        el = blocks[j]
        if el.tag == qn('w:tbl'):
            break
        if _para_text(el):
            texts.append(_para_text(el))
            break
        j -= 1
    return texts


FULL_W = 10060              # dxa: the 503pt text column (styler._set_table_borders)


def _unfloat_box(tbl_el):
    """Terminal fallback: put a floating side box back in the text flow at
    full width. An in-flow table splits across pages, so it can never strand
    one - the layout costs the side-by-side look but never a dead band."""
    tblpr = tbl_el.find(qn('w:tblPr'))
    if tblpr is None:
        return False
    pos = tblpr.find(qn('w:tblpPr'))
    if pos is None:
        return False                    # already in the flow
    tblpr.remove(pos)
    overlap = tblpr.find(qn('w:tblOverlap'))
    if overlap is not None:
        tblpr.remove(overlap)
    tblw = tblpr.find(qn('w:tblW'))
    if tblw is not None:
        tblw.set(qn('w:w'), str(FULL_W))
        tblw.set(qn('w:type'), 'dxa')
    grid = tbl_el.find(qn('w:tblGrid'))
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        if len(cols) == 1:
            cols[0].set(qn('w:w'), str(FULL_W))
    for tc in tbl_el.iter(qn('w:tc')):
        tcpr = tc.find(qn('w:tcPr'))
        if tcpr is None:
            continue
        tcw = tcpr.find(qn('w:tcW'))
        if tcw is not None:
            tcw.set(qn('w:w'), str(FULL_W))
            tcw.set(qn('w:type'), 'dxa')
    return True


def _move_box_later(blocks, i, spec):
    """Shift a floating box one body paragraph later in the flow, so the
    paragraphs it was dragging can flow up and fill the page above. The
    box keeps its positioning (right dock, width, wrap) untouched."""
    j = i + 1
    while j < len(blocks):
        el = blocks[j]
        if el.tag == qn('w:tbl') or spec.is_boundary_block(el):
            return False                # never cross a table or a topic head
        if _para_text(el):
            break
        j += 1
    else:
        return False
    # never land the box directly in front of another table: two floats with
    # no text between them anchor to the same paragraph at the same spot and
    # Word prints one box ON TOP of the other (17-Aug Daily CA English p6,
    # "Regulation of Major & Minor Minerals" over "About DMF", 2026-10-12)
    k = j + 1
    while k < len(blocks) and blocks[k].tag == qn('w:p') \
            and not _para_text(blocks[k]):
        k += 1
    if k < len(blocks) and blocks[k].tag == qn('w:tbl'):
        return False
    blocks[j].addnext(blocks[i])
    return True


def _separate_stacked_boxes(body):
    """A side box whose previous text block is another floating table
    shares that box's anchor and position: Word overprints the two. Put the
    second back in the flow at full width (an input can arrive that way;
    the fitter no longer creates it). Returns the count."""
    blocks = _blocks(body)
    floats = set(_positioned_tables(blocks))
    fixed = 0
    for i in sorted(floats):
        j = i - 1
        while j >= 0 and blocks[j].tag == qn('w:p') and not _para_text(blocks[j]):
            j -= 1
        if j >= 0 and j in floats and _unfloat_box(blocks[i]):
            floats.discard(i)
            fixed += 1
    return fixed


# ---------------------------------------------------------------- squeezing

def _cant_split_row(blocks, title):
    """Keep the News-in-Shorts row titled `title` whole across pages
    (w:cantSplit). The styler strips every cantSplit so that tall rows CAN
    break; this puts one back, measured, on a row Word broke right after
    its head. Returns True when a row was changed."""
    want = _norm_text(title)
    for el in blocks:
        if el.tag != qn('w:tbl'):
            continue
        for tr in el.findall(qn('w:tr')):
            tcs = tr.findall(qn('w:tc'))
            if len(tcs) != 2:
                continue
            got = _norm_text(" ".join(_para_text(p) for p in tcs[0].iter(qn('w:p'))))
            if got != want:
                continue
            trpr = tr.find(qn('w:trPr'))
            if trpr is None:
                trpr = tr.makeelement(qn('w:trPr'), {})
                tr.insert(0, trpr)
            if trpr.find(qn('w:cantSplit')) is not None:
                return False
            trpr.insert(0, trpr.makeelement(qn('w:cantSplit'), {}))
            return True
    return False


def _scale_val(v, num, den, floor=None):
    nv = max(1, int(round(v * num / den)))
    if floor is not None:
        nv = max(floor, nv)
    return nv


def _squeeze_topic(blocks, rng):
    """One half-point step: run sizes down 0.5pt (10pt floor), line heights
    and after-spacing scaled to match. The boundary block itself (banner
    table / H1 title) and any run of title size are never touched."""
    s, e = rng
    for el in blocks[s + 1:e]:
        for p in el.iter(qn('w:p')):
            if _para_max_size(p) >= NO_SQUEEZE_MIN:
                continue
            for r in p.iter(qn('w:r')):
                rpr = r.find(qn('w:rPr'))
                if rpr is None:
                    continue
                for tag in ('w:sz', 'w:szCs'):
                    sz = rpr.find(qn(tag))
                    if sz is None:
                        continue
                    try:
                        v = int(sz.get(qn('w:val')))
                    except (TypeError, ValueError):
                        continue
                    if v > MIN_HALF_PTS:
                        sz.set(qn('w:val'), str(v - 1))
            ppr = p.find(qn('w:pPr'))
            if ppr is None:
                continue
            sp = ppr.find(qn('w:spacing'))
            if sp is None:
                continue
            line = sp.get(qn('w:line'))
            if line:
                try:
                    sp.set(qn('w:line'), str(_scale_val(int(line), 21, 22,
                                                        floor=200)))
                except ValueError:
                    pass
            after = sp.get(qn('w:after'))
            if after:
                try:
                    sp.set(qn('w:after'), str(_scale_val(int(after), 21, 22)))
                except ValueError:
                    pass


# ---------------------------------------------------------------- main

def fit_document(styled_path, workdir, spec=ENGLISH_SPEC,
                 progress=lambda stage, pct: None, notes=None):
    """Measured fit pass. Edits the styled docx in place and returns the path
    of the PDF that matches its final state (so the caller can skip its own
    conversion)."""
    notes = notes if notes is not None else []
    stem = os.path.splitext(os.path.basename(styled_path))[0]
    fit_pdf = os.path.join(workdir, f"{stem}_fit.pdf")

    doc = Document(styled_path)
    body = doc.element.body
    dropped = _static_fixes(body, spec)

    blocks = _blocks(body)
    ranges = _topic_ranges(blocks, spec)
    steps = defaultdict(int)
    float_moves = defaultdict(int)
    float_shrinks = defaultdict(int)
    box_moves = defaultdict(int)
    squeezed = set()
    moved_floats = 0
    moved_boxes = 0
    shrunk_floats = 0
    unfloated = 0
    kept_rows = 0
    best = None                                  # (score, docx_copy, pdf_copy)
    stale = 0
    blacklist = set()        # structural moves that made the layout worse
    # a move sequence explores several shifts of the SAME box/float before
    # judging: one shift often ties or dips before the pages settle (bounded
    # lookahead; revert to the sequence start if it never pays off)
    seq = None               # {'key','docx','score','tries'}

    def _revert_to(path):
        d = Document(path)
        return d, d.element.body

    for round_no in range(MAX_ROUNDS):
        doc.save(styled_path)
        pass_docx = os.path.join(workdir, f"{stem}_fit_pass{round_no}.docx")
        pass_pdf = os.path.join(workdir, f"{stem}_fit_pass{round_no}.pdf")
        shutil.copy2(styled_path, pass_docx)
        progress(f"Fitting pages (pass {round_no + 1})",
                 min(78, 40 + round_no * 5))
        docx_to_pdf(pass_docx, pass_pdf)
        try:
            pages, bounds, pdf_floats = _measure(pass_pdf, spec)
        except Exception:
            if best is None:
                best = (float("inf"), pass_docx, pass_pdf)
            break
        if len(bounds) != len(ranges) or not bounds:
            if best is None:                     # cannot map reliably - stop
                best = (float("inf"), pass_docx, pass_pdf)
            break

        score = _score(pages, bounds, spec)

        # judge a running move sequence: cleared once it beats its starting
        # score; reverted wholesale (and blacklisted) when the lookahead is
        # exhausted without ever paying off
        if seq is not None:
            if score < seq['score'] - 0.5:
                seq = None                       # exploration paid off
            elif seq['tries'] >= 4:
                if score > seq['score'] + 0.5:
                    blacklist.add(seq['key'])
                    doc, body = _revert_to(seq['docx'])
                    blocks = _blocks(body)
                    ranges = _topic_ranges(blocks, spec)
                    stale = 0
                    seq = None
                    continue                     # re-measure reverted state
                seq = None                       # all ties - keep the moves

        if best is None or score < best[0]:
            best = (score, pass_docx, pass_pdf)
            stale = 0
        else:
            stale += 1

        anchors = _anchor_map(blocks)
        actions = set()
        acted = False

        def note_seq(key):
            nonlocal seq
            if seq is None or seq['key'] != key:
                seq = {'key': key, 'docx': pass_docx,
                       'score': score, 'tries': 1}
            else:
                seq['tries'] += 1

        # W: dead space at a page bottom while the topic continues overleaf.
        # If the next page opens with a pushed float, forward its anchor -
        # squeezing cannot fix an anchor push. Otherwise squeeze the topic.
        # The trigger looks at the TEXT column (last_left): a side box can
        # fill the right edge while the text column runs dry.
        for p in range(len(pages) - 1):
            gap = spec.content_bottom - pages[p]['last']
            gap_left = spec.content_bottom - pages[p]['last_left']
            if gap_left <= GAP_LIMIT or pages[p]['last'] == 0:
                continue
            # a boundary at the very top of the next page means a fresh topic
            # starts there - that break is wanted ("or start with the next page")
            if any(bp == p + 1 and by < spec.content_top + 40
                   for bp, by in bounds):
                continue
            pushed = None
            for fi, (fp, fy, fw, fh) in enumerate(pdf_floats):
                if fp != p + 1 or fy > spec.content_top + 200:
                    continue
                # cell inline pictures land in pdf_floats too and match no
                # anchor - keep scanning until a real float is found
                ai, entry = _match_anchor(anchors, pdf_floats, fi, fw, fh)
                if entry is not None:
                    pushed = (ai, entry)
                    break
            if pushed:
                ai, entry = pushed
                fkey, skey = ('float', ai), ('shrink', ai)
                did = False
                if fkey not in blacklist and float_moves[ai] < MAX_FLOAT_MOVES \
                        and (seq is None or seq['key'] == fkey):
                    if _forward_anchor(blocks, entry, spec):
                        float_moves[ai] += 1
                        moved_floats += 1
                        note_seq(fkey)
                        did = True
                # forwarding blocked (table/boundary in the way, budget or
                # blacklist): a slightly smaller picture that fits the band
                # beats a full-size one pushed overleaf
                if not did and skey not in blacklist \
                        and float_shrinks[ai] < 4 \
                        and (seq is None or seq['key'] == skey):
                    if _shrink_float(entry):
                        float_shrinks[ai] += 1
                        shrunk_floats += 1
                        note_seq(skey)
                        did = True
                if did:
                    acted = True
                    break            # one structural move per measure cycle
            # a floating side box that did not fit drags its anchor text to
            # the next page exactly like a picture does - shift the box one
            # paragraph later so the stranded text can flow up (identified by
            # matching the next page's first line against the box's header
            # bar / anchor paragraph)
            nt = _norm_text(pages[p + 1]['first_text'])
            box_moved = False
            if nt:
                for oi, bi in enumerate(_positioned_tables(blocks)):
                    exhausted = (box_moves[oi] >= MAX_FLOAT_MOVES
                                 or ('box', oi) in blacklist)
                    if exhausted and ('unfloat', oi) not in blacklist:
                        # out of shifts: judge the unfloat fallback below
                        if any(_norm_text(t) == nt
                               for t in _box_lead_texts(blocks, bi) if t) \
                                and _unfloat_box(blocks[bi]):
                            unfloated += 1
                            acted = True
                            box_moved = True
                            blacklist.add(('unfloat', oi))
                            note_seq(('unfloat', oi))
                            break
                    if exhausted \
                            or (seq is not None
                                and seq['key'] not in (('box', oi),
                                                       ('unfloat', oi))):
                        continue
                    if any(_norm_text(t) == nt
                           for t in _box_lead_texts(blocks, bi) if t):
                        if _move_box_later(blocks, bi, spec):
                            box_moves[oi] += 1
                            moved_boxes += 1
                            acted = True
                            box_moved = True
                            note_seq(('box', oi))
                            blocks = _blocks(body)
                            ranges = _topic_ranges(blocks, spec)
                        elif ('unfloat', oi) not in blacklist \
                                and _unfloat_box(blocks[bi]):
                            # nowhere left to shift it: back into the flow at
                            # full width, where Word can split it
                            unfloated += 1
                            acted = True
                            box_moved = True
                            blacklist.add(('unfloat', oi))
                            note_seq(('unfloat', oi))
                        break
            if box_moved:
                break                # block indexes changed - re-measure
            # squeezing helps only when the WHOLE page ends short; an empty
            # left column beside a full-height box is a box problem, not a
            # text-size problem
            if gap > GAP_LIMIT:
                t = _topic_at(bounds, p, pages[p]['last'] + 0.1)
                if t >= 0:
                    actions.add(t)

        # R: a News-in-Shorts row split right after its title -> keep that
        # row whole (structural, one per cycle, judged by score like a move)
        if not acted:
            for p, title in _stranded_rows(pages, bounds, spec):
                key = ('row', _norm_text(title))
                if key in blacklist or (seq is not None and seq['key'] != key):
                    continue
                if _cant_split_row(blocks, title):
                    kept_rows += 1
                    acted = True
                    note_seq(key)
                    break

        # O: a page opening with the stray end of the previous structure
        # ("last line adjust in above page") -> squeeze the owning topic
        for p in (_orphan_pages(pages, bounds, spec)
                  + _sliver_pages(pages, bounds, spec)):
            t = _topic_at(bounds, p - 1, pages[p - 1]['last'] + 0.1)
            if t >= 0:
                actions.add(t)

        # T: a small topic tail alone at the top of its last page. The last
        # topic has no boundary after it: its tail is the last page itself.
        tails = [(bounds[k], bounds[k + 1]) for k in range(len(bounds) - 1)]
        if bounds and len(pages) > 1:
            lp = len(pages) - 1
            tails.append((bounds[-1], (lp, pages[lp]['last'])))
        for k, ((sp_, sy), (ep_, ey)) in enumerate(tails):
            if ep_ <= sp_:                       # same page - no tail
                continue
            # real tail only when content actually sits above the boundary
            # (a boundary at the very top of a page is a clean fresh start)
            if k < len(bounds) - 1 and pages[ep_]['first'] >= ey - 4:
                continue
            tail = ey - spec.content_top
            if 0 < tail < TAIL_LIMIT:
                actions.add(k)

        # a sequence that can no longer continue (no matching candidate this
        # round) is judged now: revert if it never paid off, else keep
        if seq is not None and not acted:
            if score > seq['score'] + 0.5:
                blacklist.add(seq['key'])
                doc, body = _revert_to(seq['docx'])
                blocks = _blocks(body)
                ranges = _topic_ranges(blocks, spec)
                stale = 0
                seq = None
                continue                         # re-measure reverted state
            seq = None

        if acted:
            actions = set()      # structure first: let anchor forwarding
            stale = 0            # settle before any text is squeezed; scores
                                 # lag behind structural moves, so the
                                 # plateau counter starts fresh afterwards
        else:
            actions = {t for t in actions if steps[t] < spec.max_steps}
        if not actions and not acted:
            break
        # plateau: squeeze-only rounds that stopped improving are not worth
        # more Word exports (never break while structure is still moving)
        if stale >= 4 and not acted:
            break
        for t in sorted(actions):
            _squeeze_topic(blocks, ranges[t])
            steps[t] += 1
            squeezed.add(t)
        if actions:
            # squeezing re-flows every page: a structural move that failed
            # in the OLD geometry may well work in the new one
            blacklist.clear()

    # keep the best measured layout - a change that made things worse is
    # simply discarded (edits applied after the last measurement included)
    if best is not None:
        shutil.copy2(best[1], styled_path)
        shutil.copy2(best[2], fit_pdf)

    if dropped:
        notes.append(f"{dropped} empty paragraph(s) removed.")
    if squeezed:
        notes.append(f"{len(squeezed)} topic(s) tightened by up to half a "
                     f"point to fill their pages.")
    if moved_floats:
        notes.append(f"{moved_floats} illustration anchor(s) shifted so text "
                     f"fills the page above.")
    if kept_rows:
        notes.append(f"{kept_rows} News-in-Shorts row(s) kept whole so a title "
                     f"is not left alone at a page foot.")
    if moved_boxes:
        notes.append(f"{moved_boxes} side box(es) re-anchored so text fills "
                     f"the page above.")
    if shrunk_floats:
        notes.append(f"{shrunk_floats} illustration(s) reduced ~12% to fit "
                     f"their page.")
    if unfloated:
        notes.append(f"{unfloated} box(es) shown full-width - too tall to sit "
                     f"beside the text without leaving a gap.")
    return fit_pdf
