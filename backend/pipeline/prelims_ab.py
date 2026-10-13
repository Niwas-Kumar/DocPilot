# -*- coding: utf-8 -*-
"""Prelims Test Series - Answer Booklet (English / Hindi).

Formatting only: nothing is invented. The title lines are left empty because
they do not exist in the raw file - the user fills them in Word afterwards.

Input layout (from the page the user gives us):
    1.<tab>(C) ... 100.<tab>(C)      <- flat answer key
    1.<tab>Text Solution:            <- then one block per question
    Ans: C / Exp: ...
    [Source: ...] [PW ONLYIAS SUPERHINT ...]   (1x1 tables)

Output layout (official template):
    section 1, one column   : title lines (blank) + answer-key grid
    section 2, two columns  : 1.<tab>Ans: (c) / Exp: ... / boxes
"""
import copy
import math
import os
import re

from docx.oxml.ns import qn
from docx.shared import Pt

from . import docx_tools as dt

TEMPLATE = "prelims_ab.docx"

# The template's page header carries the sample's own test code ("PTS 2708").
# Every booklet would print that code regardless of the actual test, so it is
# replaced at build time: with a code found in the input file name when there
# is one, otherwise with an obvious placeholder to type over.
SAMPLE_HEADER_CODE = "PTS 2708"
# short enough to fit the header's green tab (the box clips longer text)
HEADER_PLACEHOLDER = "[CODE]"
# '_' is a word character, so NEITHER \b fires in '..._PTS 2708_Indian...' -
# the guards have to exclude alphanumerics explicitly
CODE_RE = re.compile(r"(?<![A-Za-z0-9])([A-Z]{2,5})\s*[- ]?\s*(\d{3,5})(?!\d)")

# body geometry measured from the official booklets
IND_LEFT = 446              # twips
IND_HANG = 446
SPACING = dict(before="30", after="30", line="283", lineRule="auto")
TAB_STOPS = (900, 1276)
# Lists, measured from the official booklets: a level-0 item puts its number
# or ❖ at 450 and its text at 900 (the first tab stop); a nested bullet sits
# on that 900 line with its text on the next stop, 1276. The raw files carry
# A4 indents (720/360, 1440/360) that wander off this grid in the column.
LIST_IND = {0: (900, 450)}
NESTED_IND = (1276, 376)
# Inside a box the list hugs the box edge: marker at the cell margin.
BOX_LIST_STEP = 360
MAX_IMG_PT = 218            # widest image in the official booklets
HINDI_FONT = "Nirmala UI"
HINDI_SZ = "20"             # half-points = 10 pt

# The template (built from the official booklet) defines the ❖ Wingdings
# bullet as numId 3. Source bullet lists (the typist's plain •) are remapped
# onto it so lists print exactly like the published booklets.
TEMPLATE_BULLET_NUMID = "3"
# Level 1 of that template list is a plain "–"; the published booklets nest
# ➢ (Wingdings F0D8) under a numbered or ❖ item, so nested bullets get a
# list of their own (added to the booklet once, see _arrow_num_id).
ARROW_CHAR = ""

# The published booklets carry the PW ONLYIAS SUPERHINT boxes but not the
# "Source: ..." citation boxes (the sample input has 76 boxes, the official
# output keeps 19 - all of them SuperHints). Source lines are an internal
# reference for the question setter, not student-facing. Flip this to False to
# keep them.
DROP_SOURCE_BOXES = True
SOURCE_RE = re.compile(r"^\s*(source|स्रोत)\s*[:\-]", re.I)

# "1.<tab>(C)"  /  "100. (d)"
KEY_RE = re.compile(r"^\s*(\d{1,3})[.)]\s*\(?\s*([A-Da-d])\s*\)?\s*$")
# "1.<tab>Text Solution:"  (the marker is English even in the Hindi papers)
SOL_RE = re.compile(r"^\s*(\d{1,3})[.)]\s*(Text\s*Solution|Solution)\s*:?\s*$", re.I)
# leading "Ans: C" / "उत्तर: C"
ANS_RE = re.compile(r"^\s*(Ans|Answer|उत्तर)\s*[:.\-]?\s*\(?\s*([A-Da-d])\s*\)?\s*$", re.I)
ANS_PREFIX_RE = re.compile(r"^\s*(Ans|Answer|उत्तर)\s*[:.\-]?\s*\(?\s*([A-Da-d])\s*\)?", re.I)


def _runs(p_el):
    out = []
    for child in p_el:
        if child.tag == qn('w:r'):
            out.append(child)
        elif child.tag == qn('w:hyperlink'):
            out.extend(c for c in child if c.tag == qn('w:r'))
    return out


def _run_text(r):
    return "".join(t.text or "" for t in r.findall(qn('w:t')))


def _el_text(el):
    return "".join(t.text or "" for t in el.iter(qn('w:t')))


def parse_input(blocks):
    """Split the sliced blocks into (answer_key, questions).

    answer_key : {qno: 'c'}          from the flat list
    questions  : [{'no':1,'ans':'c','blocks':[el, ...]}, ...]
    """
    answer_key = {}
    questions = []
    current = None

    for el in blocks:
        if el.tag == qn('w:p'):
            txt = dt.clean(_el_text(el))

            m = SOL_RE.match(txt)
            if m:
                current = {"no": int(m.group(1)), "ans": None, "blocks": []}
                questions.append(current)
                continue

            if current is None:
                m = KEY_RE.match(txt)
                if m:
                    answer_key[int(m.group(1))] = m.group(2).lower()
                continue

            # first paragraph of a solution carries "Ans: X" as its own run
            if current["ans"] is None:
                letter, stripped = _split_ans_run(el)
                if letter:
                    current["ans"] = letter
                    if not dt.clean(_el_text(el)):
                        continue          # nothing left in the paragraph
            current["blocks"].append(el)
        else:
            if current is None:
                continue
            if DROP_SOURCE_BOXES and SOURCE_RE.match(dt.clean(_el_text(el))):
                continue
            current["blocks"].append(el)

    return answer_key, questions


def _split_ans_run(p_el):
    """Remove a leading 'Ans: X' / 'उत्तर: X' from the paragraph.

    The label is frequently SPLIT ACROSS RUNS ('उत्तर' + ': A'), so matching a
    single run's text misses it and the answer prints twice. Match against the
    joined paragraph text, then consume that many characters across runs.
    Returns (letter, changed).
    """
    runs = _runs(p_el)
    joined = "".join(_run_text(r) for r in runs)
    stripped = joined.lstrip()
    lead_ws = len(joined) - len(stripped)
    m = ANS_PREFIX_RE.match(stripped)
    if not m:
        return None, False
    letter = m.group(2).lower()

    # consume the label plus any separator that follows it
    consume = lead_ws + m.end()
    rest = joined[consume:]
    consume += len(rest) - len(rest.lstrip(" \t:- "))

    remaining = consume
    for r in runs:
        if remaining <= 0:
            break
        for t in list(r.findall(qn('w:t'))):
            if remaining <= 0:
                break
            text = t.text or ""
            if len(text) <= remaining:
                remaining -= len(text)
                r.remove(t)
            else:
                t.text = text[remaining:]
                t.set(qn('xml:space'), 'preserve')
                remaining = 0
        if not _run_text(r).strip():
            br = r.find(qn('w:br'))
            if br is not None:
                r.remove(br)
            if len(r) == 0 or all(ch.tag == qn('w:rPr') for ch in r):
                parent = r.getparent()
                if parent is not None:
                    parent.remove(r)

    # the label was usually followed by a line break; the remainder should not
    # open with a blank line
    for r in _runs(p_el):
        if _run_text(r).strip():
            break
        br = r.find(qn('w:br'))
        if br is not None:
            r.remove(br)
        if len(r) == 0 or all(ch.tag == qn('w:rPr') for ch in r):
            parent = r.getparent()
            if parent is not None:
                parent.remove(r)
    return letter, True


def reconcile(answer_key, questions):
    """Cross-check the flat key against each solution's own 'Ans:' line."""
    issues = []
    resolved = {}
    numbers = sorted(set(list(answer_key.keys()) + [q["no"] for q in questions]))
    by_no = {q["no"]: q for q in questions}

    for n in numbers:
        from_key = answer_key.get(n)
        from_sol = by_no[n]["ans"] if n in by_no else None
        if from_key and from_sol and from_key != from_sol:
            issues.append(f"Q{n}: answer key says ({from_key}) but the solution "
                          f"says ({from_sol}) - key used, please verify")
            resolved[n] = from_key
        elif from_key:
            resolved[n] = from_key
        elif from_sol:
            resolved[n] = from_sol
            issues.append(f"Q{n}: missing from the answer-key list - taken from the solution")
        else:
            issues.append(f"Q{n}: no answer found")

    missing_sol = [n for n in answer_key if n not in by_no]
    if missing_sol:
        issues.append("No solution text for: " + ", ".join(f"Q{n}" for n in sorted(missing_sol)))
    return resolved, issues


# ------------------------------------------------------------------ building

def _set_borders(tbl_el, color="000000", sz=4):
    tblpr = tbl_el.find(qn('w:tblPr'))
    if tblpr is None:
        tblpr = tbl_el.makeelement(qn('w:tblPr'), {})
        tbl_el.insert(0, tblpr)
    old = tblpr.find(qn('w:tblBorders'))
    if old is not None:
        tblpr.remove(old)
    from docx.oxml import parse_xml
    edges = "".join(
        f'<w:{e} w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        for e in ("top", "left", "bottom", "right", "insideH", "insideV"))
    tblpr.append(parse_xml(
        '<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + edges + '</w:tblBorders>'))


def _bullet_num_ids(src_doc):
    """numIds in the SOURCE whose level-0 format is 'bullet'."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    out = set()
    try:
        numbering = src_doc.part.part_related_by(RT.NUMBERING).element
    except KeyError:
        return out
    abstract_fmt = {}
    for a in numbering.findall(qn('w:abstractNum')):
        lvl0 = next((l for l in a.findall(qn('w:lvl'))
                     if l.get(qn('w:ilvl')) == '0'), None)
        fmt = lvl0.find(qn('w:numFmt')) if lvl0 is not None else None
        abstract_fmt[a.get(qn('w:abstractNumId'))] = (
            fmt.get(qn('w:val')) if fmt is not None else None)
    for n in numbering.findall(qn('w:num')):
        ref = n.find(qn('w:abstractNumId'))
        if ref is not None and abstract_fmt.get(ref.get(qn('w:val'))) == 'bullet':
            out.add(n.get(qn('w:numId')))
    return out


def _arrow_num_id(doc):
    """numId of a ➢ bullet list in the booklet, created on first use."""
    from docx.oxml import parse_xml
    numbering = doc.part.numbering_part.element
    cached = numbering.get('docpilotArrow')
    if cached:
        return cached
    abs_ids = [int(a.get(qn('w:abstractNumId'))) for a in numbering.findall(qn('w:abstractNum'))]
    num_ids = [int(n.get(qn('w:numId'))) for n in numbering.findall(qn('w:num'))]
    abs_id, num_id = str(max(abs_ids, default=0) + 1), str(max(num_ids, default=0) + 1)
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    lvls = "".join(
        f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>'
        f'<w:lvlText w:val="{ARROW_CHAR}"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="{NESTED_IND[0]}" w:hanging="{NESTED_IND[1]}"/></w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr>'
        f'</w:lvl>' for i in range(9))
    abstract = parse_xml(f'<w:abstractNum {ns} w:abstractNumId="{abs_id}">'
                         f'<w:multiLevelType w:val="hybridMultilevel"/>{lvls}</w:abstractNum>')
    # schema: every abstractNum precedes every num
    first_num = numbering.find(qn('w:num'))
    if first_num is not None:
        first_num.addprevious(abstract)
    else:
        numbering.append(abstract)
    numbering.append(parse_xml(f'<w:num {ns} w:numId="{num_id}">'
                               f'<w:abstractNumId w:val="{abs_id}"/></w:num>'))
    numbering.set('docpilotArrow', num_id)
    return num_id


def _remap_bullets(doc, p_el, orig_num_id, bullet_ids):
    """Point a copied bullet paragraph at the published bullets: ❖ (the
    template's own list) at level 0, ➢ when nested."""
    if orig_num_id not in bullet_ids:
        return
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        return
    npr = ppr.find(qn('w:numPr'))
    if npr is None:
        return
    nid = npr.find(qn('w:numId'))
    if nid is None:
        return
    if _list_level(p_el) > 0:
        nid.set(qn('w:val'), _arrow_num_id(doc))
        npr.find(qn('w:ilvl')).set(qn('w:val'), '0')
    else:
        nid.set(qn('w:val'), TEMPLATE_BULLET_NUMID)


def _list_level(p_el):
    """ilvl of a list paragraph, or None when it is not in a list."""
    ppr = p_el.find(qn('w:pPr'))
    npr = ppr.find(qn('w:numPr')) if ppr is not None else None
    if npr is None:
        return None
    il = npr.find(qn('w:ilvl'))
    try:
        return int(il.get(qn('w:val'))) if il is not None else 0
    except (TypeError, ValueError):
        return 0


# pPr children in schema order. Word silently ignores a property written out
# of order (the Question Booklet lost its tab stops and indents that way), so
# every normalised paragraph is put back into this order.
PPR_ORDER = [qn(f'w:{t}') for t in (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange")]


def _order_ppr(ppr):
    rank = {tag: i for i, tag in enumerate(PPR_ORDER)}
    kids = sorted(ppr, key=lambda el: rank.get(el.tag, len(PPR_ORDER)))
    for el in kids:
        ppr.remove(el)
        ppr.append(el)


def _strip_trailing_breaks(p_el):
    """A line break that ends a paragraph only adds a blank line under it -
    under a list item, and under the PW ONLYIAS SUPERHINT header of a box."""
    for r in reversed(_runs(p_el)):
        if _run_text(r).strip() or r.find(qn('w:drawing')) is not None \
                or r.find(qn('w:tab')) is not None:
            return
        for br in r.findall(qn('w:br')):
            if br.get(qn('w:type')) in (None, 'textWrapping'):
                r.remove(br)
        if all(ch.tag == qn('w:rPr') for ch in r) and not _run_text(r):
            r.getparent().remove(r)
        else:
            return


def _orig_num_id(el):
    if el.tag != qn('w:p'):
        return None
    ppr = el.find(qn('w:pPr'))
    if ppr is None:
        return None
    npr = ppr.find(qn('w:numPr'))
    if npr is None:
        return None
    nid = npr.find(qn('w:numId'))
    return nid.get(qn('w:val')) if nid is not None else None


def _normalize_paragraph(p_el, variant, in_table=False):
    """Direct formatting instead of source styles.

    Copied paragraphs keep their pStyle, which resolves against the TEMPLATE's
    styles.xml - and there `optionindentation` is centre-aligned, so Hindi
    answers rendered centred. The official flattens everything to direct
    formatting; do the same.
    """
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = p_el.makeelement(qn('w:pPr'), {})
        p_el.insert(0, ppr)

    for tag in ('w:pStyle', 'w:jc', 'w:spacing', 'w:widowControl', 'w:keepNext',
                'w:keepLines', 'w:pageBreakBefore'):
        for old in ppr.findall(qn(tag)):
            ppr.remove(old)

    # The template's docDefaults switch widow control OFF, so a paragraph's
    # last line could strand alone at the top of the next column ("...Hence,"
    # / "Option (b) is correct."). The official booklet sets it on every
    # paragraph.
    ppr.append(ppr.makeelement(qn('w:widowControl'), {}))

    level = _list_level(p_el)

    jc = ppr.makeelement(qn('w:jc'), {})
    jc.set(qn('w:val'), 'both')
    ppr.append(jc)

    sp = ppr.makeelement(qn('w:spacing'), {})
    for k, v in SPACING.items():
        sp.set(qn(f'w:{k}'), v)
    ppr.append(sp)

    # Lists keep their own marker (number or bullet) but sit on the booklet's
    # grid; a direct w:ind overrides the A4 indent the numbering carries.
    if level is None and not in_table:
        left, hang = IND_LEFT, IND_HANG
    elif level is None:
        left = hang = None                   # box text keeps the cell margin
    elif in_table:
        left, hang = BOX_LIST_STEP * (level + 1), BOX_LIST_STEP
    else:
        left, hang = LIST_IND.get(level, NESTED_IND)
    if left is not None:
        for ind in ppr.findall(qn('w:ind')):
            ppr.remove(ind)
        ind = ppr.makeelement(qn('w:ind'), {})
        ind.set(qn('w:left'), str(left))
        ind.set(qn('w:hanging'), str(hang))
        ppr.append(ind)

    if not in_table:
        for tabs in ppr.findall(qn('w:tabs')):
            ppr.remove(tabs)
        tabs = ppr.makeelement(qn('w:tabs'), {})
        for pos in TAB_STOPS:
            t = tabs.makeelement(qn('w:tab'), {})
            t.set(qn('w:val'), 'left')
            t.set(qn('w:pos'), str(pos))
            tabs.append(t)
        ppr.append(tabs)

    if variant.lower().startswith("hi"):
        # every run - runs WITHOUT rPr would otherwise stay at the template
        # default (TNR 11pt), which shows up on all the inter-word space runs
        def _force(rpr):
            rf = rpr.find(qn('w:rFonts'))
            if rf is None:
                rf = rpr.makeelement(qn('w:rFonts'), {})
                rpr.insert(0, rf)
            for slot in ('w:ascii', 'w:hAnsi', 'w:cs'):
                rf.set(qn(slot), HINDI_FONT)
            for tag in ('w:sz', 'w:szCs'):
                el = rpr.find(qn(tag))
                if el is None:
                    el = rpr.makeelement(qn(tag), {})
                    rpr.append(el)
                el.set(qn('w:val'), HINDI_SZ)

        for r in p_el.iter(qn('w:r')):
            rpr = r.find(qn('w:rPr'))
            if rpr is None:
                rpr = r.makeelement(qn('w:rPr'), {})
                r.insert(0, rpr)
            _force(rpr)
        mark_rpr = ppr.find(qn('w:rPr'))
        if mark_rpr is None:
            mark_rpr = ppr.makeelement(qn('w:rPr'), {})
            ppr.append(mark_rpr)
        _force(mark_rpr)

    _order_ppr(ppr)


def _cap_images(block_el):
    """The official scales every picture into the column; the raw sizes run off
    the page (one 468 pt image ends 205 pt past the paper edge)."""
    max_emu = MAX_IMG_PT * 12700
    for ext in list(block_el.iter(qn('wp:extent'))):
        cx, cy = int(ext.get('cx')), int(ext.get('cy'))
        if cx <= max_emu:
            continue
        f = max_emu / cx
        ext.set('cx', str(int(cx * f)))
        ext.set('cy', str(int(cy * f)))
        holder = ext.getparent()
        for a_ext in holder.iter(qn('a:ext')):
            a_ext.set('cx', str(int(cx * f)))
            a_ext.set('cy', str(int(cy * f)))


def _fit_table_to_column(tbl_el):
    """Boxes copied from the A4 source are far too wide for a 2-column page."""
    from docx.oxml import parse_xml
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    tblpr = tbl_el.find(qn('w:tblPr'))
    if tblpr is None:
        tblpr = tbl_el.makeelement(qn('w:tblPr'), {})
        tbl_el.insert(0, tblpr)
    for tag in ('w:tblW', 'w:tblLayout', 'w:tblInd'):
        old = tblpr.find(qn(tag))
        if old is not None:
            tblpr.remove(old)
    tblpr.append(parse_xml(f'<w:tblW {ns} w:w="5000" w:type="pct"/>'))
    tblpr.append(parse_xml(f'<w:tblLayout {ns} w:type="autofit"/>'))
    tblpr.append(parse_xml(f'<w:tblInd {ns} w:w="0" w:type="dxa"/>'))
    # cell widths must follow, otherwise Word keeps the old fixed grid
    for tc in tbl_el.iter(qn('w:tc')):
        tcpr = tc.find(qn('w:tcPr'))
        if tcpr is None:
            continue
        old = tcpr.find(qn('w:tcW'))
        if old is not None:
            tcpr.remove(old)
        tcpr.insert(0, parse_xml(f'<w:tcW {ns} w:w="5000" w:type="pct"/>'))
    grid = tbl_el.find(qn('w:tblGrid'))
    if grid is not None:
        tbl_el.remove(grid)


def build_answer_grid(doc, answers, before):
    """25x8 style grid: four (No., Ans) column pairs reading down."""
    numbers = sorted(answers)
    if not numbers:
        return None
    groups = 4
    rows = math.ceil(len(numbers) / groups)

    tbl = doc.add_table(rows=rows, cols=groups * 2)
    _set_borders(tbl._tbl)
    tbl.autofit = True

    # official grid: 456-twip rows, vertically centred, letters regular 11pt
    from docx.oxml import parse_xml
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    for tr in tbl._tbl.findall(qn('w:tr')):
        trpr = tr.find(qn('w:trPr'))
        if trpr is None:
            trpr = tr.makeelement(qn('w:trPr'), {})
            tr.insert(0, trpr)
        trpr.append(parse_xml(f'<w:trHeight {ns} w:val="456"/>'))
    for tc in tbl._tbl.iter(qn('w:tc')):
        tcpr = tc.find(qn('w:tcPr'))
        if tcpr is None:
            tcpr = tc.makeelement(qn('w:tcPr'), {})
            tc.insert(0, tcpr)
        tcpr.append(parse_xml(f'<w:vAlign {ns} w:val="center"/>'))

    for gi in range(groups):
        chunk = numbers[gi * rows:(gi + 1) * rows]
        for ri, n in enumerate(chunk):
            c_no = tbl.cell(ri, gi * 2)
            c_an = tbl.cell(ri, gi * 2 + 1)
            for cell, text, bold in ((c_no, f"{n}.", True),
                                     (c_an, f"({answers[n]})", False)):
                p = cell.paragraphs[0]
                p.alignment = 1
                run = p.add_run(text)
                run.bold = bold

    before.addprevious(tbl._tbl)
    return tbl


# Line 1 of the title block is the product name and is the same in every
# published booklet (English and Hindi alike). Bump it when the cycle changes.
SERIES_TITLE = "Prelims Wallah Test Series 2027"
# '_' is a word character, so \b never fires in 'Test 08_A_PTS 2708_...'
TEST_NO_RE = re.compile(r"(?<![A-Za-z0-9])test\s*[-–—:]?\s*0*(\d{1,2})(?!\d)", re.I)
SUBJECT_HINT_RE = re.compile(
    r"polity|history|geograph|econom|science|environment|society|governance|"
    r"csat|current|ethics|culture", re.I)


def _test_line(source_path, issues):
    """'Test 08: Indian Polity (Part II)' read from the input file name.

    Production files are named like the published booklet
    ('Test 08_A_PTS 2708_Indian Polity Part II_FLT_...'); when the name says
    nothing, a placeholder is left and reported - the raw paper itself never
    names the test.
    """
    stem = os.path.splitext(os.path.basename(source_path))[0]
    parts = [p.strip() for p in re.split(r"[_|]", stem) if p.strip()]
    m = TEST_NO_RE.search(stem)
    subject = None
    for i, part in enumerate(parts):
        if TEST_NO_RE.fullmatch(part) and i + 2 < len(parts):
            cand = parts[i + 2]                # Test 08 | A | PTS 2708 | subject
            if len(cand) > 4 and not re.fullmatch(r"[A-Z]{2,5}\s*\d{3,5}", cand):
                subject = cand
            break
    if subject is None:
        subject = next((p for p in parts if SUBJECT_HINT_RE.search(p)), None)

    if m and subject:
        # file names flatten the published '(Part II)' to 'Part II'
        subject = re.sub(r"\s+part\s+([IVX0-9]+)\s*$", r" (Part \1)", subject,
                         flags=re.I)
        return f"Test {m.group(1).zfill(2)}: {subject}"
    issues.append("Title line left blank - the file name carries no "
                  "'Test 08 ... <subject>' to read it from (the raw paper does "
                  "not name the test). Type it on the ruled line in Word.")
    return ""


# The official title block, measured from the published booklets: line 1 in
# the Title style (whose 886-twip left indent is cancelled so it centres on
# the page), line 2 between two blue rules. Line 2 stays a ruled blank line
# when the test name is unknown, ready to be typed into.
TITLE_PPR = ('<w:pStyle w:val="Title"/><w:spacing w:before="40" w:after="100"/>'
             '<w:ind w:left="0"/><w:jc w:val="center"/>')
TEST_LINE_PPR = ('<w:pBdr>'
                 '<w:top w:val="single" w:sz="8" w:space="4" w:color="4472C4"/>'
                 '<w:bottom w:val="single" w:sz="8" w:space="4" w:color="4472C4"/>'
                 '</w:pBdr><w:spacing w:after="160"/><w:ind w:left="0"/>'
                 '<w:jc w:val="center"/>')


def _title_line(doc, text, before, ppr="", size_half=None):
    """Centred bold title-block line. The paragraph mark carries the run
    formatting too, so text typed into a blank line comes out the same."""
    from docx.oxml import parse_xml
    from xml.sax.saxutils import escape
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    sz = (f'<w:sz w:val="{size_half}"/><w:szCs w:val="{size_half}"/>'
          if size_half else "")
    rpr = f'<w:rPr><w:b/><w:bCs/>{sz}</w:rPr>'
    run = (f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'
           if text else "")
    p = parse_xml(f'<w:p {ns}><w:pPr>{ppr}{rpr}</w:pPr>{run}</w:p>')
    before.addprevious(p)
    return p


def _emit_ans_line(doc, qno, letter, before, label, variant):
    """'1.<tab>Ans: (c)' on the same hanging indent as the Exp text below it.

    The number sits at the margin and a real <w:tab/> carries 'Ans:' to the
    446-twip indent, exactly as in the official booklet. A tab typed as text
    into a bare paragraph fell to Word's default 0.5" stop instead, so 'Ans:'
    printed to the right of 'Exp:'. keepNext holds it to its explanation.
    """
    from docx.oxml import parse_xml
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    bold = '<w:rPr><w:b/><w:bCs/></w:rPr>'
    p = parse_xml(
        f'<w:p {ns}><w:pPr/>'
        f'<w:r>{bold}<w:t>{qno}.</w:t></w:r><w:r>{bold}<w:tab/></w:r>'
        f'<w:r>{bold}<w:t xml:space="preserve">{label} ({letter})</w:t></w:r></w:p>')
    _normalize_paragraph(p, variant)
    ppr = p.find(qn('w:pPr'))
    ppr.insert(0, ppr.makeelement(qn('w:keepNext'), {}))
    _order_ppr(ppr)
    before.addprevious(p)
    return p


def _no_stretch_before_line_breaks(doc):
    """Word's 'Don't expand character spaces on a line ending with
    SHIFT+RETURN'. The text is justified, and without this Word spreads a
    line that ends in a manual break across the whole width - the box header
    printed as 'PW          ONLYIAS          SUPERHINT'."""
    settings = doc.settings.element
    compat = settings.find(qn('w:compat'))
    if compat is None:
        compat = settings.makeelement(qn('w:compat'), {})
        settings.append(compat)
    if compat.find(qn('w:doNotExpandShiftReturn')) is not None:
        return
    flag = compat.makeelement(qn('w:doNotExpandShiftReturn'), {})
    # schema order: after ulTrailSpace and the other legacy flags before it
    before = [qn(f'w:{t}') for t in (
        "useSingleBorderforContiguousCells", "wpJustification", "noTabHangInd",
        "noLeading", "spaceForUL", "noColumnBalance",
        "balanceSingleByteDoubleByteWidth", "noExtraLineSpacing",
        "doNotLeaveBackslashAlone", "ulTrailSpace")]
    last = None
    for child in compat:
        if child.tag in before:
            last = child
    if last is not None:
        last.addnext(flag)
    else:
        compat.insert(0, flag)


def _replace_header_code(out, input_path, notes):
    """Swap the sample's test code in the page headers.

    The code sits inside a TEXTBOX in the header (and is duplicated in the
    mc:AlternateContent fallback), so section.header.paragraphs never sees it -
    walk every w:p in the header parts' raw XML instead.
    """
    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    m = CODE_RE.search(os.path.splitext(os.path.basename(input_path))[0].upper())
    code = f"{m.group(1)} {m.group(2)}" if m else HEADER_PLACEHOLDER
    want = SAMPLE_HEADER_CODE.replace(" ", "")

    hit = False
    for rel in out.part.rels.values():
        if rel.reltype not in (RT.HEADER, RT.FOOTER):
            continue
        root = rel.target_part.element
        for p in root.iter(qn('w:p')):
            ts = [t for t in p.iter(qn('w:t'))]
            joined = "".join(t.text or "" for t in ts).replace(" ", "")
            if want not in joined:
                continue
            first = True
            for t in ts:
                if not (t.text or "").strip():
                    continue
                t.text = code if first else ""
                first = False
            hit = True

    if hit:
        if m:
            notes.append(f"Header test code set to '{code}' (from the file name).")
        else:
            notes.append(f"Header test code left as '{HEADER_PLACEHOLDER}' - the "
                         "file name carries no code like 'PTS 2708'. Edit it in "
                         "Word (double-click the page header).")


def build(input_path, out_path, start_page=None, variant="English",
          progress=lambda s, p: None, cancelled=lambda: False):

    progress("Reading source", 15)
    doc_in, blocks = dt.slice_blocks(input_path, start_page=start_page, cancelled=cancelled)
    if cancelled():
        raise InterruptedError("stopped")

    progress("Parsing answers", 35)
    answer_key, questions = parse_input(blocks)
    if not questions:
        raise ValueError(
            "No solutions found. Check the start page - it should be the page "
            "where the answer key / solutions begin.")
    answers, issues = reconcile(answer_key, questions)

    progress("Building booklet", 60)
    out = dt.open_template(TEMPLATE)
    _replace_header_code(out, input_path, issues)
    anchors = dt.clear_body_keep_sections(out)
    sect_break = anchors[0] if anchors else None
    tail = out.element.body.find(qn('w:sectPr'))

    # --- section 1: title block + answer grid
    place = sect_break if sect_break is not None else tail
    _title_line(out, SERIES_TITLE, place, ppr=TITLE_PPR)
    _title_line(out, _test_line(input_path, issues), place, ppr=TEST_LINE_PPR,
                size_half=36)
    build_answer_grid(out, answers, before=place)
    dt.new_paragraph(out, "", before=place)
    _no_stretch_before_line_breaks(out)

    # --- section 2: the solutions
    label = "उत्तर:" if variant.lower().startswith("hi") else "Ans:"
    bullet_ids = _bullet_num_ids(doc_in)
    for q in questions:
        if cancelled():
            raise InterruptedError("stopped")
        letter = answers.get(q["no"], q["ans"] or "?")
        _emit_ans_line(out, q["no"], letter, tail, label, variant)
        for el in q["blocks"]:
            orig_num = _orig_num_id(el)
            # carries pictures + numbering across, rewriting their ids;
            # a plain copy leaves broken images and decimal bullets
            new = dt.copy_block_with_media(out, el, doc_in, before=tail)
            _cap_images(new)
            if new.tag == qn('w:tbl'):
                _fit_table_to_column(new)
                for p in new.iter(qn('w:p')):
                    _strip_trailing_breaks(p)
                    _normalize_paragraph(p, variant, in_table=True)
            else:
                _strip_trailing_breaks(new)
                # reads the source list level, so it runs before the remap
                _normalize_paragraph(new, variant)
                _remap_bullets(out, new, orig_num, bullet_ids)

    progress("Saving", 90)
    out.save(out_path)
    return out_path, issues


def process(input_path, workdir, variant="English", start_page=None,
            progress=lambda s, p: None, cancelled=lambda: False, **kw):
    os.makedirs(workdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(input_path))[0]
    out_name = f"{stem}_Answer Booklet_{variant}.docx"
    out_path = os.path.join(workdir, out_name)
    _, issues = build(input_path, out_path, start_page=start_page, variant=variant,
                      progress=progress, cancelled=cancelled)

    # PDF companion beside the DOCX, same stem: the UI offers both downloads
    progress("Making the PDF", 95)
    pdf = dt.make_pdf_companion(out_path, issues)
    if pdf:
        # Word re-encodes every JPEG on export; print the raw file's own
        dt.restore_original_jpegs(out_path, pdf)

    progress("Done", 100)
    return out_path, issues
