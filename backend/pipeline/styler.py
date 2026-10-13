# -*- coding: utf-8 -*-
"""Restyle a raw typed Daily-CA docx into the OnlyIAS branded layout.

The transformation is purely deterministic (verified 1:1 against the official
sample): fonts, colors, bullets, banners, info-boxes, news table, images.
"""
import re
from xml.sax.saxutils import escape
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from .ooxml import (drop_google_tab_frontmatter,
                    get_or_add, set_run_fonts, set_color, set_bold, para_shading,
                    para_spacing, clear_indent, get_ilvl, pt_to_twips, pt_to_emu)
from .template import apply_page_template

RED = "C00000"
GREEN = "00B050"
BLUE = "0070C0"
BANNER_BG = "BDD7EE"
BODY_FONT = "Times New Roman"
BANNER_FONT = "Cambria"
CS_FONT = "Kokila"

COL_W_PT = 612 - 57 - 52          # text column width (503pt)
NEWS_LEFT_W = 2700                # twips (135pt)
NEWS_RIGHT_W = 7360               # twips (368pt)

W = qn('w:p'), qn('w:tbl')


# ---------------------------------------------------------------- run helpers

def _runs(p_el):
    """All runs of a paragraph incl. inside hyperlinks (direct descendants only)."""
    out = []
    for child in p_el:
        if child.tag == qn('w:r'):
            out.append(child)
        elif child.tag == qn('w:hyperlink'):
            out.extend(ch for ch in child if ch.tag == qn('w:r'))
    return out


def _run_text(r):
    return "".join(t.text or "" for t in r.findall(qn('w:t')))


def _para_text(p_el):
    return "".join(_run_text(r) for r in _runs(p_el))


def _is_bold(r):
    rpr = r.find(qn('w:rPr'))
    if rpr is None:
        return False
    b = rpr.find(qn('w:b'))
    return b is not None and b.get(qn('w:val')) not in ('0', 'false')


def _fully_bold(p_el):
    seen = False
    for r in _runs(p_el):
        t = _run_text(r)
        if not t.strip():
            continue
        seen = True
        if not _is_bold(r):
            return False
    return seen


def _get_rpr(r):
    rpr = r.find(qn('w:rPr'))
    if rpr is None:
        rpr = r.makeelement(qn('w:rPr'), {})
        r.insert(0, rpr)
    return rpr


def _style_runs(p_el, size=22, color=None, bold=None, font=BODY_FONT):
    for r in _runs(p_el):
        rpr = _get_rpr(r)
        set_run_fonts(rpr, ascii_font=font, cs_font=CS_FONT, size_half_pts=size)
        # typists' highlighter marks are not part of the official design
        for tag in ('w:highlight', 'w:shd'):
            hl = rpr.find(qn(tag))
            if hl is not None:
                rpr.remove(hl)
        if color is not None:
            set_color(rpr, color)
        if bold is True:
            set_bold(rpr, True)


def _get_ppr(p_el):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = p_el.makeelement(qn('w:pPr'), {})
        p_el.insert(0, ppr)
    return ppr


def _set_jc(ppr, val):
    jc = get_or_add(ppr, 'w:jc')
    jc.set(qn('w:val'), val)


def _style_name(p_el, doc):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        return None
    ps = ppr.find(qn('w:pStyle'))
    if ps is None:
        return None
    sid = ps.get(qn('w:val'))
    try:
        return doc.styles.element.xpath(f'w:style[@w:styleId="{sid}"]/w:name')[0].get(qn('w:val'))
    except Exception:
        return sid


def _clear_pstyle(p_el):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is not None:
        ps = ppr.find(qn('w:pStyle'))
        if ps is not None:
            ppr.remove(ps)


def _has_numpr(p_el):
    ppr = p_el.find(qn('w:pPr'))
    return ppr is not None and ppr.find(qn('w:numPr')) is not None


# ---------------------------------------------------------------- numbering

def rewrite_numbering(doc):
    """Every bullet level in the source becomes: lvl0 ❖ (Wingdings F076), lvl1+ ➢ (F0D8)."""
    try:
        num_part = doc.part.part_related_by(RT.NUMBERING)
    except KeyError:
        return
    root = num_part.element
    for absnum in root.findall(qn('w:abstractNum')):
        for lvl in absnum.findall(qn('w:lvl')):
            fmt = lvl.find(qn('w:numFmt'))
            if fmt is None or fmt.get(qn('w:val')) != 'bullet':
                continue
            ilvl = int(lvl.get(qn('w:ilvl'), '0'))
            char = '' if ilvl == 0 else ''
            left = 420 if ilvl == 0 else 440 + 420 * ilvl
            hang = 420 if ilvl == 0 else 440
            txt = get_or_add(lvl, 'w:lvlText')
            txt.set(qn('w:val'), char)
            # indent
            ppr = get_or_add(lvl, 'w:pPr')
            for old in ppr.findall(qn('w:ind')):
                ppr.remove(old)
            ind = ppr.makeelement(qn('w:ind'), {})
            ind.set(qn('w:left'), str(left))
            ind.set(qn('w:hanging'), str(hang))
            ppr.append(ind)
            # bullet font
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


# ---------------------------------------------------------------- banners

def _make_banner_el(text):
    return parse_xml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:pPr>'
        f'<w:shd w:val="clear" w:color="auto" w:fill="{BANNER_BG}"/>'
        '<w:spacing w:before="40" w:after="80" w:line="600" w:lineRule="atLeast"/>'
        '<w:jc w:val="center"/>'
        '</w:pPr>'
        '<w:r><w:rPr>'
        f'<w:rFonts w:ascii="{BANNER_FONT}" w:hAnsi="{BANNER_FONT}" w:cs="{CS_FONT}"/>'
        '<w:b/><w:bCs/><w:sz w:val="48"/><w:szCs w:val="48"/>'
        f'</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
        '</w:p>'
    )


def _to_banner(p_el):
    ppr = _get_ppr(p_el)
    _clear_pstyle(p_el)
    clear_indent(ppr)
    para_shading(ppr, BANNER_BG)
    para_spacing(ppr, before_pt=6, after_pt=4, line_pt=30, rule="atLeast")
    _set_jc(ppr, 'center')
    _style_runs(p_el, size=48, color="000000", bold=True, font=BANNER_FONT)


# ---------------------------------------------------------------- paragraph rules

def _style_h1(p_el):
    ppr = _get_ppr(p_el)
    clear_indent(ppr)
    para_spacing(ppr, before_pt=8, after_pt=2, line_pt=19, rule="atLeast")
    _set_jc(ppr, 'left')
    _style_runs(p_el, size=32, color=RED, bold=True)


def _style_subject(p_el):
    ppr = _get_ppr(p_el)
    para_spacing(ppr)
    _style_runs(p_el, size=22, color=GREEN, bold=True)


def _style_blue_heading(p_el, color=BLUE):
    ppr = _get_ppr(p_el)
    para_spacing(ppr)
    _style_runs(p_el, size=22, color=color, bold=True)


# Bold headings INSIDE tables (info-box bodies, News-in-Shorts right cells,
# grids) print black, not blue (client, 2026-10-12). Body headings outside
# tables stay blue. Set to BLUE to restore the old look.
CELL_HEAD_COLOR = "000000"


def _style_body(p_el):
    ppr = _get_ppr(p_el)
    para_spacing(ppr)
    _style_runs(p_el, size=22)


def _style_list_item(p_el):
    ppr = _get_ppr(p_el)
    clear_indent(ppr)  # numbering definition now controls indents
    para_spacing(ppr)
    _style_runs(p_el, size=22)


CONTEXT_RE = re.compile(r'^\s*(Context\s*:?)\s*$', re.I)


def _try_split_context(p_el, color=BLUE):
    """'Context: <body...>' inline lead-in -> own heading line + remainder."""
    runs = _runs(p_el)
    if not runs:
        return None
    first = runs[0]
    t = _run_text(first)
    m = re.match(r'^\s*(Context|संदर्भ)\s*:\s*', t)
    if m is None or not _is_bold(first):
        return None
    rest_in_first = t[m.end():]
    # remainder must exist somewhere (else the fully-bold rule handles it)
    remainder = rest_in_first + "".join(_run_text(r) for r in runs[1:])
    if not remainder.strip():
        return None
    # build the standalone heading paragraph
    head = parse_xml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:r><w:rPr><w:b/><w:bCs/></w:rPr>'
        f'<w:t xml:space="preserve">{escape(m.group(1))}:</w:t></w:r></w:p>'
    )
    _style_blue_heading(head, color)
    p_el.addprevious(head)
    # strip the lead-in from the original first run
    for tnode in first.findall(qn('w:t')):
        first.remove(tnode)
    if rest_in_first:
        tnode = first.makeelement(qn('w:t'), {})
        tnode.text = rest_in_first
        tnode.set(qn('xml:space'), 'preserve')
        first.append(tnode)
        set_bold(_get_rpr(first), False)
    else:
        first.getparent().remove(first)
    return head


# ---------------------------------------------------------------- tables

def _set_table_borders(tbl_el, with_inside):
    tblpr = tbl_el.find(qn('w:tblPr'))
    if tblpr is None:
        tblpr = tbl_el.makeelement(qn('w:tblPr'), {})
        tbl_el.insert(0, tblpr)
    # fit the Letter text column (503pt) and let tall rows break across pages
    tblw = tblpr.find(qn('w:tblW'))
    if tblw is not None:
        tblw.set(qn('w:w'), '10060')
        tblw.set(qn('w:type'), 'dxa')
    for cant in tbl_el.iter(qn('w:cantSplit')):
        cant.getparent().remove(cant)
    old = tblpr.find(qn('w:tblBorders'))
    if old is not None:
        tblpr.remove(old)
    edges = ['top', 'left', 'bottom', 'right'] + (['insideH', 'insideV'] if with_inside else [])
    xml = ('<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           + "".join(f'<w:{e} w:val="single" w:sz="8" w:space="0" w:color="{BLUE}"/>' for e in edges)
           + '</w:tblBorders>')
    tblpr.append(parse_xml(xml))


def _cell_paragraphs(tc):
    return [p for p in tc if p.tag == qn('w:p')]


SIDEBOX_W = 4800            # dxa: 240pt, ~48% of the 503pt column - measured
                            # from the official sample's side boxes
SIDEBOX_W_PT = SIDEBOX_W / 20.0
# A floating table cannot split across pages: when one is too tall for the
# space a page can offer, Word carries it AND its anchor text overleaf and
# leaves a dead band behind. So a box only floats when its rendered height
# fits comfortably inside the 638pt content area - taller boxes stay
# full-width, where Word splits them happily.
SIDEBOX_MAX_H_PT = 360.0
SIDEBOX_CHARS_PER_LINE = 44   # TNR 11pt across the ~230pt usable cell width
SIDEBOX_LINE_PT = 13.6


def _estimate_box_height(paras):
    """Rendered height of an info-box at side-box width, in points."""
    total = 12.0                                   # cell padding
    for p in paras:
        text = _para_text(p)
        lines = max(1, -(-len(text) // SIDEBOX_CHARS_PER_LINE))   # ceil
        total += lines * SIDEBOX_LINE_PT + 2.5     # + after-spacing
        for _ in p.iter(qn('wp:inline')):          # pictures inside the box
            total += 120.0
    return total


def _make_sidebox(tbl_el):
    """Float the info-box at the right margin, body text wrapping down its
    left - the official layout (client mark-up "SIDE BOX", 2026-08-26)."""
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    tblpr = tbl_el.find(qn('w:tblPr'))
    if tblpr is None:
        tblpr = tbl_el.makeelement(qn('w:tblPr'), {})
        tbl_el.insert(0, tblpr)
    for tag in ('w:tblpPr', 'w:tblOverlap'):
        el = tblpr.find(qn(tag))
        if el is not None:
            tblpr.remove(el)
    # w:tblpPr must precede w:tblW etc. (schema order) or Word ignores it
    pos = parse_xml(
        f'<w:tblpPr {ns} w:leftFromText="200" w:rightFromText="0" '
        'w:topFromText="30" w:bottomFromText="60" '
        'w:vertAnchor="text" w:horzAnchor="margin" '
        'w:tblpXSpec="right" w:tblpY="1"/>')
    style = tblpr.find(qn('w:tblStyle'))
    if style is not None:
        style.addnext(pos)
    else:
        tblpr.insert(0, pos)
    pos.addnext(parse_xml(f'<w:tblOverlap {ns} w:val="never"/>'))

    tblw = tblpr.find(qn('w:tblW'))
    if tblw is None:
        tblw = tblpr.makeelement(qn('w:tblW'), {})
        tblpr.append(tblw)
    tblw.set(qn('w:w'), str(SIDEBOX_W))
    tblw.set(qn('w:type'), 'dxa')
    grid = tbl_el.find(qn('w:tblGrid'))
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        if len(cols) == 1:
            cols[0].set(qn('w:w'), str(SIDEBOX_W))
    for tc in tbl_el.iter(qn('w:tc')):
        tcpr = tc.find(qn('w:tcPr'))
        if tcpr is None:
            tcpr = tc.makeelement(qn('w:tcPr'), {})
            tc.insert(0, tcpr)
        tcw = get_or_add(tcpr, 'w:tcW')
        tcw.set(qn('w:w'), str(SIDEBOX_W))
        tcw.set(qn('w:type'), 'dxa')
    # pictures inside the box must fit its narrow column now
    for inline in tbl_el.iter(qn('wp:inline')):
        _cap_drawing(inline, 225)


def _style_infobox(tbl_el, doc):
    """1x1 table -> blue info-box: header bar (first para(s)) + transparent
    body, floated as a right-hand side box when it fits one page."""
    _set_table_borders(tbl_el, with_inside=False)
    tr = tbl_el.find(qn('w:tr'))
    tc = tr.find(qn('w:tc')) if tr is not None else None
    if tc is None:
        return
    paras = _cell_paragraphs(tc)
    if not paras:
        return
    header = paras[0]
    # optional "(UPSC CSE Prelims 20xx)" tail -> its own right-aligned bar line
    txt = _para_text(header)
    m = re.search(r'\((UPSC[^)]*)\)\s*$', txt)
    extra_bar = None
    if m and m.start() > 0:
        title_part = txt[:m.start()].rstrip()
        tag_part = txt[m.start():].strip()
        # rebuild header as two clean paragraphs
        for r in list(_runs(header)):
            parent = r.getparent()
            parent.remove(r)
        t1 = parse_xml(
            '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f'<w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">{escape(title_part)}</w:t></w:r>')
        header.append(t1)
        extra_bar = parse_xml(
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            '<w:pPr><w:jc w:val="right"/></w:pPr>'
            f'<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">{escape(tag_part)}</w:t></w:r></w:p>')
        header.addnext(extra_bar)
    bar_paras = [header] + ([extra_bar] if extra_bar is not None else [])
    for i, bp in enumerate(bar_paras):
        ppr = _get_ppr(bp)
        clear_indent(ppr)
        para_shading(ppr, BLUE)
        para_spacing(ppr, before_pt=1, after_pt=1, line_pt=14, rule="atLeast")
        if bp is header:
            _set_jc(ppr, 'left')
        # strip list numbering if any snuck in
        npr = ppr.find(qn('w:numPr'))
        if npr is not None:
            ppr.remove(npr)
        _style_runs(bp, size=22, color="FFFFFF", bold=True)
    # body paragraphs
    for p in _cell_paragraphs(tc):
        if p is header or (extra_bar is not None and p is extra_bar):
            continue
        _style_cell_paragraph(p, doc)

    # the official design runs these boxes down the RIGHT side of the page
    # with the body text wrapping on their left ("SIDE BOX" mark-up) - but
    # only while the box is short enough to sit inside one page
    if _estimate_box_height(_cell_paragraphs(tc)) <= SIDEBOX_MAX_H_PT:
        _make_sidebox(tbl_el)


def _style_cell_paragraph(p_el, doc):
    if _has_numpr(p_el):
        _style_list_item(p_el)
    elif _fully_bold(p_el) and _para_text(p_el).strip():
        _style_blue_heading(p_el, CELL_HEAD_COLOR)
    else:
        _try_split_context(p_el, CELL_HEAD_COLOR)
        _style_body(p_el)


def _rebuild_news_table(tbl_el):
    """Rebuild the N x 2 table as a clean Word-native table.

    Google-Docs-exported tables carry constructs that stop Word from breaking
    rows across pages (verified: raw input straight through Word shows one row
    per page). Moving the cell contents into a fresh table restores native
    row-splitting behavior.
    """
    rows_content = []
    for tr in tbl_el.findall(qn('w:tr')):
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) != 2:
            return tbl_el  # unexpected shape, keep original
        rows_content.append(([p for p in tcs[0] if p.tag != qn('w:tcPr')],
                             [p for p in tcs[1] if p.tag != qn('w:tcPr')]))

    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    new_tbl = parse_xml(
        f'<w:tbl {ns}>'
        '<w:tblPr>'
        '<w:tblW w:w="10060" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/>'
        '</w:tblPr>'
        '<w:tblGrid>'
        f'<w:gridCol w:w="{NEWS_LEFT_W}"/><w:gridCol w:w="{NEWS_RIGHT_W}"/>'
        '</w:tblGrid>'
        '</w:tbl>'
    )
    for left_paras, right_paras in rows_content:
        tr = parse_xml(f'<w:tr {ns}></w:tr>')
        for width, paras in ((NEWS_LEFT_W, left_paras), (NEWS_RIGHT_W, right_paras)):
            tc = parse_xml(
                f'<w:tc {ns}>'
                '<w:tcPr>'
                f'<w:tcW w:w="{width}" w:type="dxa"/>'
                '<w:tcMar>'
                '<w:top w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>'
                '<w:bottom w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
                '</w:tcMar>'
                '<w:vAlign w:val="top"/>'
                '</w:tcPr>'
                '</w:tc>'
            )
            for p in paras:
                tc.append(p)  # moves the element
            tr.append(tc)
        new_tbl.append(tr)
    tbl_el.addprevious(new_tbl)
    tbl_el.getparent().remove(tbl_el)
    return new_tbl


def _style_news_table(tbl_el, doc):
    """N x 2 table -> News-In-Shorts layout."""
    _set_table_borders(tbl_el, with_inside=True)
    # fixed column widths
    grid = tbl_el.find(qn('w:tblGrid'))
    if grid is not None:
        cols = grid.findall(qn('w:gridCol'))
        if len(cols) == 2:
            cols[0].set(qn('w:w'), str(NEWS_LEFT_W))
            cols[1].set(qn('w:w'), str(NEWS_RIGHT_W))
    for tr in tbl_el.findall(qn('w:tr')):
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) != 2:
            continue
        # enforce cell widths
        for tc, w in zip(tcs, (NEWS_LEFT_W, NEWS_RIGHT_W)):
            tcpr = tc.find(qn('w:tcPr'))
            if tcpr is None:
                tcpr = tc.makeelement(qn('w:tcPr'), {})
                tc.insert(0, tcpr)
            tcw = get_or_add(tcpr, 'w:tcW')
            tcw.set(qn('w:w'), str(w))
            tcw.set(qn('w:type'), 'dxa')
        left, right = tcs
        for p in _cell_paragraphs(left):
            ppr = _get_ppr(p)
            clear_indent(ppr)
            npr = ppr.find(qn('w:numPr'))
            if npr is not None:
                ppr.remove(npr)
            para_spacing(ppr, before_pt=2, after_pt=2)
            _set_jc(ppr, 'center')
            _style_runs(p, size=22, color=RED, bold=True)
        for p in _cell_paragraphs(right):
            _style_cell_paragraph(p, doc)
    _cap_cell_images(tbl_el)


def _cap_cell_images(tbl_el):
    """Inline images inside table cells must fit their column."""
    for tr in tbl_el.findall(qn('w:tr')):
        tcs = tr.findall(qn('w:tc'))
        for idx, tc in enumerate(tcs):
            max_w_pt = 120 if (len(tcs) == 2 and idx == 0) else 340
            for inline in tc.iter(qn('wp:inline')):
                _cap_drawing(inline, max_w_pt)


def _cap_drawing(holder, max_w_pt):
    ext = holder.find(qn('wp:extent'))
    if ext is None:
        return
    cx = int(ext.get('cx')); cy = int(ext.get('cy'))
    max_emu = pt_to_emu(max_w_pt)
    if cx <= max_emu:
        return
    f = max_emu / cx
    ncx, ncy = int(cx * f), int(cy * f)
    ext.set('cx', str(ncx)); ext.set('cy', str(ncy))
    for xfrm_ext in holder.iter(qn('a:ext')):
        xfrm_ext.set('cx', str(ncx)); xfrm_ext.set('cy', str(ncy))


def _normalize_floating_images(body):
    """Floating content images: cap to ~52% of column, dock to the right margin."""
    col_emu = pt_to_emu(COL_W_PT)
    for anchor in body.iter(qn('wp:anchor')):
        if anchor.find(qn('wp:wrapNone')) is not None:
            continue  # template art, not content
        ext = anchor.find(qn('wp:extent'))
        if ext is None:
            continue
        # two illustrations anchored near each other will happily stack on
        # top of one another while overlap is allowed - Word pushes them
        # apart once it is off
        anchor.set('allowOverlap', '0')
        cx = int(ext.get('cx'))
        if cx > int(col_emu * 0.58):
            _cap_drawing(anchor, COL_W_PT * 0.52)
        # dock right, relative to text margin
        posH = anchor.find(qn('wp:positionH'))
        if posH is not None:
            posH.set('relativeFrom', 'margin')
            for ch in list(posH):
                posH.remove(ch)
            align = posH.makeelement(qn('wp:align'), {})
            align.text = 'right'
            posH.append(align)


# ---------------------------------------------------------------- main walk

_REVISION_TAGS = ('w:tblGridChange', 'w:tblPrChange', 'w:trPrChange', 'w:tcPrChange',
                  'w:pPrChange', 'w:rPrChange', 'w:sectPrChange')


def _cleanup_source(body):
    """Strip Google-Docs export artifacts that break Word pagination."""
    for tag in _REVISION_TAGS:
        for el in list(body.iter(qn(tag))):
            el.getparent().remove(el)
    for el in list(body.iter(qn('w:pageBreakBefore'))):
        el.getparent().remove(el)
    for ind in body.iter(qn('w:tblInd')):
        ind.set(qn('w:w'), '0')
        ind.set(qn('w:type'), 'dxa')


def restyle_docx(in_path, out_path, banner_text="Topics"):
    doc = Document(in_path)
    apply_page_template(doc)
    rewrite_numbering(doc)
    _cleanup_source(doc.element.body)
    # a Google Docs "Tabs" export writes the tab name (and a hard page break)
    # into the top of the body; both would otherwise print as a near-empty
    # first page - see drop_google_tab_frontmatter
    drop_google_tab_frontmatter(doc.element.body, lambda el: _style_name(el, doc))

    body = doc.element.body
    children = [el for el in body if el.tag in (qn('w:p'), qn('w:tbl'))]

    prev_was_h1 = False
    for el in children:
        if el.tag == qn('w:tbl'):
            # heading pStyles inside cells carry keepNext/keepLines chains that
            # stop Word from breaking rows across pages (bisect-verified)
            for ps in list(el.iter(qn('w:pStyle'))):
                ps.getparent().remove(ps)
            rows = el.findall(qn('w:tr'))
            ncols = len(rows[0].findall(qn('w:tc'))) if rows else 0
            if len(rows) == 1 and ncols == 1:
                _style_infobox(el, doc)
            elif ncols == 2:
                el = _rebuild_news_table(el)
                _style_news_table(el, doc)
            else:
                _set_table_borders(el, with_inside=True)
                for tc in el.iter(qn('w:tc')):
                    for p in _cell_paragraphs(tc):
                        _style_cell_paragraph(p, doc)
            prev_was_h1 = False
            continue

        sname = _style_name(el, doc) or ""
        text = _para_text(el).strip()

        if sname.lower().strip() == "heading 1":
            _style_h1(el)
            prev_was_h1 = True
            continue
        if sname == "Title":
            _to_banner(el)
            prev_was_h1 = False
            continue
        if prev_was_h1 and re.match(r'^(Subject|विषय)\s*:', text):
            _style_subject(el)
            prev_was_h1 = False
            continue
        prev_was_h1 = False

        if _has_numpr(el):
            _style_list_item(el)
        elif text and _fully_bold(el):
            _style_blue_heading(el)
        else:
            _try_split_context(el)
            _style_body(el)

    _normalize_floating_images(body)

    # "Topics" banner at the very top
    banner = _make_banner_el(banner_text)
    first = next((el for el in body if el.tag in (qn('w:p'), qn('w:tbl'))), None)
    if first is not None:
        first.addprevious(banner)
    else:
        body.insert(0, banner)

    doc.save(out_path)
    return out_path
