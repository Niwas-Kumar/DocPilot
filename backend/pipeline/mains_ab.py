# -*- coding: utf-8 -*-
"""Mains Test Series - Answer Booklet (Hindi).

Formatting only. The published sample also carries a cover page (title block,
syllabus box, question index with difficulty / "Must Attempt" ratings) and word
limits - none of which exist in the raw file, so this pipeline does not invent
them. The user pastes the cover in afterwards.

Input quirk: the raw file is a Google Docs export whose content sits inside
w:sdt content controls. python-docx skips those silently (it sees 1 of 6 tables
and 77 of 113 paragraphs), so the body is flattened first.
"""
import os
import re

from docx.oxml.ns import qn

from . import docx_tools as dt

TEMPLATE = "mains_ab_hindi.docx"

# "प्रश्न 1. <stem> 10 अंक"  /  "Q1. <stem> 10M"
Q_RE = re.compile(r"^\s*(?:प्रश्न|प्र\.|Q(?:uestion)?)\s*\.?\s*(\d{1,3})\s*[.)]?\s*", re.I)

# section headings inside an answer
HEAD_WORDS = ("परिचय", "मुख्य भाग", "निष्कर्ष", "आगे की राह", "निकर्ष",
              "introduction", "body", "conclusion", "way forward")
EXAMPLE_RE = re.compile(r"^\s*(उदाहरण|उदाहरणार्थ|for example|e\.g\.)", re.I)
DEMAND_RE = re.compile(r"(मुख्य\s*माँग|मुख्य\s*मांग|main\s*demand)", re.I)
PYQ_RE = re.compile(r"(PYQ|विगत\s*वर्ष)", re.I)

# Typists head the PYQ box inconsistently - "Related UPSC PYQs" (English) on one
# question and "UPSC से संबंधित PYQs" on the next. The published booklets use a
# single Hindi heading throughout.
PYQ_HEADING = "UPSC के विगत वर्षों के प्रश्न (PYQs)"

# box fills measured from the official booklet
PYQ_FILL = "E2EFD9"          # pale green PYQ box
BOX_BORDER = "000000"

# ---------------------------------------------------------------- cover page
# The published booklet opens with a cover the raw file does not contain. We
# build the LAYOUT and fill in everything the source actually provides (the
# paper code from the file name, and the question list); the few facts only a
# human knows are left as obvious placeholders to type over.
BUILD_COVER = True
COVER_BANNER_FILL = "DEEAF6"     # pale blue band behind the title block
SYLLABUS_FILL = "FFF2CC"         # cream syllabus box
INDEX_HEADER_FILL = "D9EAD3"     # pale green header row of the question index
COVER_DISPLAY_FONT = "Aptos Black"
COVER_TITLE_COLOUR = "920000"

PH_SERIES = "[SERIES / YEAR]"
PH_TEST = "[TEST NAME]"
PH_SUBJECT = "[SUBJECT]"
PH_SYLLABUS = "[syllabus topic]"
PH_REMARK = "[remark]"

# "Input SDM 2719 Hind (1).docx" -> "SDM 2719"
CODE_RE = re.compile(r"\b([A-Z]{2,5})\s*[- ]?\s*(\d{3,5})\b")

STYLE_QUESTION = "Questions"
STYLE_HEAD = "AboveSpacing"
STYLE_BULLET = "Bullet-1"
STYLE_EXAMPLE = "Shadding"
STYLE_BODY = "Normal-DK"
CHAR_BOX = "Boxcolour"

# A bare "उत्तर :" label the typist left behind between a question and its
# answer. It carries no content and the published booklets drop it.
STRAY_LABEL_RE = re.compile(r"^\s*(उत्तर|answer)\s*[:：]?\s*$", re.I)


def _text(el):
    return "".join(t.text or "" for t in el.iter(qn('w:t')))


def _set_style(p_el, style_id):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = p_el.makeelement(qn('w:pPr'), {})
        p_el.insert(0, ppr)
    old = ppr.find(qn('w:pStyle'))
    if old is not None:
        ppr.remove(old)
    ps = ppr.makeelement(qn('w:pStyle'), {})
    ps.set(qn('w:val'), style_id)
    ppr.insert(0, ps)


def _clear_direct_formatting(p_el):
    """Drop the source's fonts/sizes/indents so the template styles win."""
    ppr = p_el.find(qn('w:pPr'))
    if ppr is not None:
        for tag in ('w:ind', 'w:spacing', 'w:jc', 'w:numPr', 'w:shd'):
            for el in ppr.findall(qn(tag)):
                ppr.remove(el)
    for rpr in p_el.iter(qn('w:rPr')):
        for tag in ('w:rFonts', 'w:sz', 'w:szCs', 'w:color', 'w:lang', 'w:highlight'):
            for el in rpr.findall(qn(tag)):
                rpr.remove(el)


def _underline_to_box(p_el):
    """The published booklets replace underline emphasis with the cream
    Boxcolour character style."""
    for r in p_el.findall(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            continue
        u = rpr.find(qn('w:u'))
        if u is None:
            continue
        rpr.remove(u)
        old = rpr.find(qn('w:rStyle'))
        if old is not None:
            rpr.remove(old)
        rs = rpr.makeelement(qn('w:rStyle'), {})
        rs.set(qn('w:val'), CHAR_BOX)
        rpr.insert(0, rs)


def _cover_para(out, text, size, before, *, bold=True, colour=None,
                fill=None, font=None):
    from docx.oxml import parse_xml
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    rfonts = (f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="Nirmala UI"/>'
              if font else '<w:rFonts w:cs="Nirmala UI"/>')
    col = f'<w:color w:val="{colour}"/>' if colour else ""
    half = int(round(size * 2))
    p = parse_xml(
        f'<w:p {ns}><w:pPr><w:pStyle w:val="Normal-DK"/>{shd}'
        f'<w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr>{rfonts}{"<w:b/><w:bCs/>" if bold else ""}{col}'
        f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/></w:rPr>'
        f'<w:t xml:space="preserve">{_esc(text)}</w:t></w:r></w:p>')
    before.addprevious(p)
    return p


def _esc(s):
    from xml.sax.saxutils import escape
    return escape(s or "")


def _cover_table(out, before, cols, rows, *, header_fill=None, table_fill=None,
                 borders=True, widths=None):
    """Build a simple table for the cover and return the docx Table object."""
    from docx.oxml import parse_xml
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in (widths or [5000] * cols))
    shd = (f'<w:shd w:val="clear" w:color="auto" w:fill="{table_fill}"/>'
           if table_fill else "")
    edges = "".join(
        f'<w:{e} w:val="single" w:sz="8" w:space="0" w:color="{BOX_BORDER}"/>'
        for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
    borders_xml = f'<w:tblBorders>{edges}</w:tblBorders>' if borders else ""
    body = ""
    for r in range(rows):
        cells = ""
        for c in range(cols):
            fill = header_fill if (header_fill and r == 0) else None
            cshd = (f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
                    if fill else (f'<w:shd w:val="clear" w:color="auto" '
                                  f'w:fill="{table_fill}"/>' if table_fill else ""))
            cells += (f'<w:tc><w:tcPr>{cshd}</w:tcPr>'
                      f'<w:p><w:pPr><w:pStyle w:val="Normal-DK"/></w:pPr></w:p></w:tc>')
        body += f'<w:tr>{cells}</w:tr>'
    tbl = parse_xml(
        f'<w:tbl {ns}><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{shd}'
        f'{borders_xml}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>')
    before.addprevious(tbl)
    from docx.table import Table
    return Table(tbl, out)


def _cell_write(cell, text, *, bold=False, centre=False, style="Normal-DK"):
    p = cell.paragraphs[0]
    _set_style(p._p, style)
    if centre:
        _centre(p._p)
    run = p.add_run(text)
    run.bold = bold
    return p


def _collect_questions(blocks):
    """Question stems, in order, with the literal 'प्रश्न N.' prefix removed."""
    out = []
    for el in blocks:
        if el.tag != qn('w:p'):
            continue
        txt = dt.clean(_text(el))
        m = Q_RE.match(txt)
        if m and len(txt) > 25:
            out.append(dt.clean(txt[m.end():]))
    return out


def _paper_code(stem):
    m = CODE_RE.search(stem.upper())
    return f"({m.group(1)} {m.group(2)})" if m else "[PAPER CODE]"


def _build_cover(out, before, questions, stem, notes):
    """Cover page: title block, syllabus box, question index, page break."""
    _cover_para(out, PH_SERIES, 36, before, colour=COVER_TITLE_COLOUR,
                fill=COVER_BANNER_FILL, font=COVER_DISPLAY_FONT)
    _cover_para(out, PH_TEST, 28, before, fill=COVER_BANNER_FILL,
                font=COVER_DISPLAY_FONT)
    _cover_para(out, PH_SUBJECT, 26, before, fill=COVER_BANNER_FILL,
                font=COVER_DISPLAY_FONT)
    _cover_para(out, _paper_code(stem), 18, before)

    syl = _cover_table(out, before, 1, 1, table_fill=SYLLABUS_FILL, borders=False)
    cell = syl.cell(0, 0)
    _cell_write(cell, "पाठ्यक्रम", bold=True, centre=True)
    extra = cell.add_paragraph()
    _set_style(extra._p, STYLE_BULLET)
    extra.add_run(PH_SYLLABUS)

    idx = _cover_table(out, before, 3, len(questions) + 1,
                       header_fill=INDEX_HEADER_FILL,
                       widths=[776, 6135, 2226])
    for c, head in enumerate(("No.", "Question", "Remark")):
        _cell_write(idx.cell(0, c), head, bold=True, centre=True)
    for i, q in enumerate(questions, start=1):
        _cell_write(idx.cell(i, 0), f"Q{i}", bold=True, centre=True)
        _cell_write(idx.cell(i, 1), q)
        _cell_write(idx.cell(i, 2), PH_REMARK, centre=True)

    from docx.oxml import parse_xml
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    before.addprevious(parse_xml(
        f'<w:p {ns}><w:r><w:br w:type="page"/></w:r></w:p>'))

    notes.append(
        "Cover page added with the question index filled in. The series name, "
        "test name, subject, syllabus lines and the Remark column are not in "
        "the source file - they are marked [like this] for you to type over.")


def _centre(p_el):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = p_el.makeelement(qn('w:pPr'), {})
        p_el.insert(0, ppr)
    for old in ppr.findall(qn('w:jc')):
        ppr.remove(old)
    jc = ppr.makeelement(qn('w:jc'), {})
    jc.set(qn('w:val'), 'center')
    ppr.append(jc)


def _paint_box(tbl_el, fill=None):
    """Thin black rule around the box, optional cell fill."""
    from docx.oxml import parse_xml
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    tblpr = tbl_el.find(qn('w:tblPr'))
    if tblpr is None:
        tblpr = tbl_el.makeelement(qn('w:tblPr'), {})
        tbl_el.insert(0, tblpr)
    for old in tblpr.findall(qn('w:tblBorders')):
        tblpr.remove(old)
    edges = "".join(
        f'<w:{e} w:val="single" w:sz="8" w:space="0" w:color="{BOX_BORDER}"/>'
        for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
    tblpr.append(parse_xml(f'<w:tblBorders {ns}>{edges}</w:tblBorders>'))

    if fill:
        for tc in tbl_el.iter(qn('w:tc')):
            tcpr = tc.find(qn('w:tcPr'))
            if tcpr is None:
                tcpr = tc.makeelement(qn('w:tcPr'), {})
                tc.insert(0, tcpr)
            for old in tcpr.findall(qn('w:shd')):
                tcpr.remove(old)
            tcpr.append(parse_xml(
                f'<w:shd {ns} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def _retitle(p_el, text):
    """Replace a paragraph's text, keeping its first run's formatting."""
    runs = p_el.findall(qn('w:r'))
    if not runs:
        return
    keep = runs[0]
    for r in runs[1:]:
        p_el.remove(r)
    for t in keep.findall(qn('w:t')):
        keep.remove(t)
    for br in keep.findall(qn('w:br')):
        keep.remove(br)
    t = keep.makeelement(qn('w:t'), {})
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    keep.append(t)


def _strip_q_prefix(p_el):
    """Remove the literal 'प्रश्न 1.' - the Questions style auto-numbers."""
    for r in p_el.findall(qn('w:r')):
        for t in r.findall(qn('w:t')):
            if not (t.text or "").strip():
                continue
            m = Q_RE.match(t.text)
            if m:
                t.text = t.text[m.end():]
                t.set(qn('xml:space'), 'preserve')
            return


def _is_bulleted(p_el):
    ppr = p_el.find(qn('w:pPr'))
    return ppr is not None and ppr.find(qn('w:numPr')) is not None


def _stays_bold(p_el):
    """True when no text run switches bold OFF.

    These files come out of Google Docs with bold ON in docDefaults, so body
    text carries an explicit b=0 and the typist's headings simply inherit. That
    inheritance is what marks a heading - a keyword list misses the ones phrased
    as questions ("... पुनर्गठित हो रहा है, कैसे?").
    """
    saw_text = False
    for r in p_el.findall(qn('w:r')):
        if not "".join(t.text or "" for t in r.findall(qn('w:t'))).strip():
            continue
        saw_text = True
        rpr = r.find(qn('w:rPr'))
        b = rpr.find(qn('w:b')) if rpr is not None else None
        if b is not None and b.get(qn('w:val')) in ("0", "false"):
            return False
    return saw_text


def _looks_like_heading(txt, style_name, p_el=None):
    t = dt.clean(txt)
    if not t or len(t) > 110:
        return False
    if style_name and style_name.lower().startswith("heading 3"):
        return True
    low = t.lower().rstrip(":： ")
    if any(low == w or low.startswith(w) for w in HEAD_WORDS):
        return True
    if p_el is not None and not _is_bulleted(p_el) and not EXAMPLE_RE.match(t):
        return _stays_bold(p_el)
    return False


def _split_at_breaks(p_el):
    """Split a paragraph into several where a <w:br/> starts a new 'उदाहरण:'.

    In the raw file some example lines are line-breaks inside their parent
    bullet; the published booklet gives each its own shaded paragraph.
    """
    import copy as _copy

    runs = p_el.findall(qn('w:r'))
    cut_at = br_run = None
    for ri, r in enumerate(runs):
        br = r.find(qn('w:br'))
        if br is None:
            continue
        # The break usually sits at the END of its run, so the example text
        # begins in the NEXT run - reading from runs[ri:] picks up the tail of
        # the current sentence instead and never matches.
        after_in_run = ""
        seen = False
        for child in r:
            if child is br:
                seen = True
                continue
            if seen and child.tag == qn('w:t'):
                after_in_run += child.text or ""
        tail_text = after_in_run + "".join(
            "".join(t.text or "" for t in rr.findall(qn('w:t')))
            for rr in runs[ri + 1:])
        if EXAMPLE_RE.match(tail_text.lstrip()):
            cut_at, br_run = ri + 1, r
            if after_in_run.strip():        # text after the break in same run
                cut_at = ri                 # move the whole run instead
            break
    if cut_at is None:
        return []

    for br in br_run.findall(qn('w:br')):
        br_run.remove(br)
    if not "".join(t.text or "" for t in br_run.findall(qn('w:t'))).strip():
        p_el.remove(br_run)                 # empty leftover

    tail = p_el.makeelement(qn('w:p'), {})
    ppr = p_el.find(qn('w:pPr'))
    if ppr is not None:
        new_ppr = _copy.deepcopy(ppr)
        for npr in new_ppr.findall(qn('w:numPr')):
            new_ppr.remove(npr)             # the example is not a bullet
        tail.append(new_ppr)
    for r in runs[cut_at:]:
        if r.getparent() is p_el:
            p_el.remove(r)
            tail.append(r)
    p_el.addnext(tail)
    return [tail]


def _style_name_of(p_el, doc):
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        return ""
    ps = ppr.find(qn('w:pStyle'))
    if ps is None:
        return ""
    sid = ps.get(qn('w:val'))
    try:
        node = doc.styles.element.xpath(f'w:style[@w:styleId="{sid}"]/w:name')
        return node[0].get(qn('w:val')) if node else sid
    except Exception:
        return sid


def build(input_path, out_path, start_page=None, notes=None,
          progress=lambda s, p: None, cancelled=lambda: False):
    notes = notes if notes is not None else []

    progress("Reading source", 15)
    # slice_blocks already descends into w:sdt, so the block list survives the
    # flatten; reading the file once keeps the extra Word pagination pass away.
    doc_in, blocks = dt.slice_blocks(input_path, start_page=start_page, cancelled=cancelled)
    dt.strip_sdt(doc_in)
    if cancelled():
        raise InterruptedError("stopped")

    progress("Restyling answers", 55)
    out = dt.open_template(TEMPLATE)
    dt.clear_body_keep_sections(out)
    tail = out.element.body.find(qn('w:sectPr'))

    if BUILD_COVER:
        stems = _collect_questions(blocks)
        if stems:
            _build_cover(out, tail, stems,
                         os.path.splitext(os.path.basename(input_path))[0], notes)

    seen_questions = 0
    split_examples = 0
    for el in blocks:
        if cancelled():
            raise InterruptedError("stopped")
        if el.tag == qn('w:p') and STRAY_LABEL_RE.match(dt.clean(_text(el))):
            continue
        # copy_block_with_media carries the pictures over and rewrites their
        # relationship ids; a plain copy leaves them dangling.
        new = dt.copy_block_with_media(out, el, doc_in, before=tail)

        if new.tag == qn('w:tbl'):
            paras = list(new.iter(qn('w:p')))
            head_txt = dt.clean(_text(paras[0])) if paras else ""
            is_pyq = bool(PYQ_RE.search(head_txt))
            is_demand = bool(DEMAND_RE.search(head_txt))
            _paint_box(new, fill=PYQ_FILL if is_pyq else None)
            for i, p in enumerate(paras):
                _clear_direct_formatting(p)
                _underline_to_box(p)
                if i == 0 and is_pyq:
                    _retitle(p, PYQ_HEADING)
                if i == 0 and (is_pyq or is_demand):
                    _centre(p)              # both box headings are centred
                _set_style(p, STYLE_BULLET if _is_bulleted(p) else STYLE_BODY)
            continue

        txt = _text(new)
        sname = _style_name_of(el, doc_in)

        if Q_RE.match(dt.clean(txt)) and len(dt.clean(txt)) > 25:
            _strip_q_prefix(new)
            _clear_direct_formatting(new)
            _underline_to_box(new)
            _set_style(new, STYLE_QUESTION)
            seen_questions += 1
            continue

        # an example tacked onto a bullet with a line break becomes its own line
        extras = _split_at_breaks(new)
        split_examples += len(extras)

        bulleted = _is_bulleted(new)
        is_head = _looks_like_heading(txt, sname, new)
        _clear_direct_formatting(new)
        _underline_to_box(new)

        if EXAMPLE_RE.match(dt.clean(txt)):
            _set_style(new, STYLE_EXAMPLE)
        elif is_head:
            _set_style(new, STYLE_HEAD)
        elif bulleted:
            _set_style(new, STYLE_BULLET)
        else:
            _set_style(new, STYLE_BODY)

        for extra in extras:
            _clear_direct_formatting(extra)
            _underline_to_box(extra)
            _set_style(extra, STYLE_EXAMPLE)

    if seen_questions == 0:
        raise ValueError("No questions found (expected paragraphs starting "
                         "'प्रश्न 1.' or 'Q1.'). Check the start page.")

    progress("Saving", 90)
    out.save(out_path)
    return out_path, seen_questions


def process(input_path, workdir, variant="Hindi", start_page=None,
            progress=lambda s, p: None, cancelled=lambda: False, **kw):
    os.makedirs(workdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(input_path))[0]
    out_path = os.path.join(workdir, f"{stem}_Ans Booklet_{variant}.docx")
    notes = []
    _, n = build(input_path, out_path, start_page=start_page, notes=notes,
                 progress=progress, cancelled=cancelled)
    progress("Done", 100)
    notes.insert(0, f"{n} question(s) formatted.")
    notes.append("Word limits (e.g. '150 शब्द') are not in the source file, so "
                 "the marks are left exactly as typed.")
    return out_path, notes
