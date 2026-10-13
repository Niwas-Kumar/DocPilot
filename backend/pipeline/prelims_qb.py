# -*- coding: utf-8 -*-
"""Prelims Test Series - Question Booklet (Bilingual, English or Hindi).

Bilingual: two uploads (English + Hindi). The published booklet alternates by PAGE:
one page of Hindi questions, then the SAME questions in English, and so on -
each page being a two-column snaking layout. Hindi comes first.

The official sample achieves that by micro-tuning line spacing per question
(40 distinct leading values, some non-integer) and it stops lining up after
page 8. We instead group the questions and let Word paginate each block, then
force a break - the alignment then holds for the whole booklet.

English / Hindi: one upload, the same pages in that language alone - its
cover and instructions in front, no other-language pages (`_arrange`).

Formatting only: the title lines are left blank because the raw files do not
contain them.
"""
import copy
import math
import os
import re

from docx.oxml import parse_xml
from docx.oxml.ns import qn

from . import docx_tools as dt

TEMPLATE = "prelims_qb_bilingual.docx"
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

# A tab in OOXML is a <w:tab/> element, not text, so "1.<tab>Which..." reads
# back as "1.Which..." - the number must not require whitespace after it.
Q_RE = re.compile(r"^\s*(\d{1,3})\s*[.)]\s*(?=\S)")
# some questions are typed with the number alone on its own line
BARE_NUM_RE = re.compile(r"^\s*(\d{1,3})\s*[.)]\s*$")
OPTION_RE = re.compile(r"^\s*\(?\s*([A-Da-d])\s*\)")
# where the questions stop and the answer key begins
KEY_RE = re.compile(r"^\s*(\d{1,3})[.)]\s*\(?\s*[A-Da-d]\s*\)?\s*$")
SOL_RE = re.compile(r"^\s*\d{1,3}[.)]\s*(Text\s*Solution|Solution)\s*:?\s*$", re.I)

MAX_PER_PAGE = 12           # safety cap while measuring

# geometry measured from the official booklet (twips)
STEM_IND = (432, 432)       # left, hanging
OPT_IND = (851, 851)
# Statements ("1. In India, ...") measured from the official: number on the
# stem's text line (432 = 78.3 pt), text on the 851 stop (99.5 pt) - the same
# column as the option text. Nested levels move one stop right.
STATEMENT_IND = (851, 419)  # left, hanging -> first line starts at 432
STATEMENT_STEP = 425
# Space above every question that does not open its page: the official
# leaves 150-170 twips (8.8-10.4 pt measured), we left 30 (4.5 pt).
QUESTION_GAP = "170"
# (before, after, gap above a question) on the official's densest columns
TIGHT_SPACING = {"hi": ("0", "0", "140"), "en": ("10", "10", "110")}
ENGLISH_FONT = "Times New Roman"
TAB_STOPS = (426, 851, 1276)
SPACING = dict(before="30", after="30", line="271", lineRule="auto")
# The official booklet sets every question - Hindi AND English - in
# Nirmala UI 10pt (measured: 21,109 runs at sz 20, 21,552 in Nirmala UI).
BODY_FONT = "Nirmala UI"
BODY_SZ = "20"              # 10 pt
HINDI_FONT = BODY_FONT      # kept for older callers
HINDI_SZ = BODY_SZ


def _text(el):
    """Text with <w:tab/> rendered as a real tab, so numbering reads naturally."""
    out = []
    for node in el.iter():
        if node.tag == qn('w:t'):
            out.append(node.text or "")
        elif node.tag == qn('w:tab'):
            out.append("\t")
    return "".join(out)


def split_questions(blocks):
    """Group blocks into [{'no': 1, 'blocks': [...]}, ...], stopping at the key."""
    questions = []
    current = None
    for el in blocks:
        if el.tag != qn('w:p'):
            if current is not None:
                current["blocks"].append(el)
            continue
        txt = dt.clean(_text(el))
        if not txt:
            if current is not None:
                current["blocks"].append(el)
            continue
        # the answer key ("1. (C)") and the solutions end the question section
        if SOL_RE.match(txt) or (KEY_RE.match(txt) and len(questions) > 1):
            break
        bare = BARE_NUM_RE.match(txt)
        m = bare or Q_RE.match(txt)
        if m and not OPTION_RE.match(txt) and not KEY_RE.match(txt):
            n = int(m.group(1))
            # Question numbers only ever go up. Without this, a hard-typed
            # sub-statement ("1. First Past the Post System: UK") inside a
            # question would be mistaken for question 1 and split the paper.
            if not questions or n > questions[-1]["no"]:
                current = {"no": n, "blocks": [el]}
                questions.append(current)
                continue
        if current is not None:
            current["blocks"].append(el)
    return questions


def _title_line(doc, text, before, size_half=None):
    """Centred bold placeholder line for the title block."""
    from xml.sax.saxutils import escape
    sz = (f'<w:sz w:val="{size_half}"/><w:szCs w:val="{size_half}"/>'
          if size_half else "")
    p = parse_xml(
        f'<w:p {NS}><w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:rPr><w:b/><w:bCs/>{sz}</w:rPr>'
        f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>')
    before.addprevious(p)
    return p


def _strip_trailing_breaks(p_el):
    """Trailing line breaks add a blank line under the paragraph.

    The raw papers end many stems with a break - sometimes in a run of its
    own, but often in the SAME run as the stem's last words ("...Litigation
    (PIL):<w:br/>"), which a run-by-run check never saw: 64 stems of Test 08
    printed a blank line before their statements (the official has one).
    So walk the paragraph's content from the end and drop every
    line/text-wrapping break until real content - text, a tab, a picture -
    is reached; then drop runs left empty.
    """
    content = (qn('w:t'), qn('w:br'), qn('w:tab'), qn('w:drawing'),
               qn('w:pict'), qn('w:object'), qn('w:sym'), qn('w:cr'))
    for node in reversed(list(p_el.iter(*content))):
        if node.tag == qn('w:t') and not (node.text or "").strip():
            continue
        if node.tag in (qn('w:br'), qn('w:cr')) and \
                node.get(qn('w:type')) in (None, 'textWrapping'):
            node.getparent().remove(node)
            continue
        break
    for r in list(p_el.iter(qn('w:r'))):
        if all(c.tag == qn('w:rPr') for c in r) and r.getparent() is not None:
            r.getparent().remove(r)


def _leading(spec):
    """'250' -> ('250', False); '250t' -> ('250', True).

    A trailing 't' asks for the official's TIGHT paragraph spacing, which it
    uses on its densest columns (Q3-Q5 of Test 08: Hindi 0/0 with 140 above a
    question, English 10/10 with 110) instead of 30/30 with 170.
    """
    s = str(spec)
    return (s[:-1], True) if s.endswith("t") else (s, False)


def _normalize_qb(p_el, lang, line="271", statement=None, question_gap=False):
    """Direct formatting per the official booklet.

    The official sets custom tab stops on ~1490 of 1494 paragraphs; without
    them the (a)-option hanging indent collapses and continuation lines start
    70 pt left of the text they continue. `line` is the leading - the official
    varies it per page (240-384) to make five questions fit; we do the same
    per GROUP, stepping down only when that group genuinely spills.

    `statement` = (label, level) turns a raw auto-numbered list item into the
    official's typed "1.<tab>text" on the booklet grid; `question_gap` adds
    the official's space above a question that does not open its page.
    """
    if statement is not None:
        _type_statement_label(p_el, statement[0])
    txt = dt.clean(_text(p_el))
    is_option = bool(OPTION_RE.match(txt))
    is_stem = bool(Q_RE.match(txt)) and not is_option

    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = p_el.makeelement(qn('w:pPr'), {})
        p_el.insert(0, ppr)

    # pPr children are ORDER-SENSITIVE (numPr, tabs, spacing, ind, jc, rPr in
    # schema order). Word silently DROPS out-of-order properties - the tab
    # stops and indents vanished when jc was written first. Rebuild cleanly.
    num_pr = ppr.find(qn('w:numPr'))
    if num_pr is not None:
        ppr.remove(num_pr)
    if statement is not None:
        num_pr = None                     # its number is typed text now
    mark_rpr = ppr.find(qn('w:rPr'))
    if mark_rpr is not None:
        ppr.remove(mark_rpr)
    for child in list(ppr):
        ppr.remove(child)
    has_num = num_pr is not None

    if num_pr is not None:
        ppr.append(num_pr)

    tabs = ppr.makeelement(qn('w:tabs'), {})
    for pos in TAB_STOPS:
        t = tabs.makeelement(qn('w:tab'), {})
        t.set(qn('w:val'), 'left')
        t.set(qn('w:pos'), str(pos))
        tabs.append(t)
    ppr.append(tabs)

    sp = ppr.makeelement(qn('w:spacing'), {})
    vals = dict(SPACING)
    vals["line"], tight = _leading(line)
    if tight:
        vals["before"], vals["after"] = TIGHT_SPACING[lang][:2]
    if question_gap:
        vals["before"] = TIGHT_SPACING[lang][2] if tight else QUESTION_GAP
    for k, v in vals.items():
        sp.set(qn(f'w:{k}'), v)
    ppr.append(sp)

    if statement is not None:
        # official grid: number on the stem's text line (432), text on the
        # next stop (851) - the same column as the (a)-(d) option text
        level = statement[1]
        ind = ppr.makeelement(qn('w:ind'), {})
        ind.set(qn('w:left'), str(STATEMENT_IND[0] + STATEMENT_STEP * level))
        ind.set(qn('w:hanging'), str(STATEMENT_IND[1]))
        ppr.append(ind)
    elif not has_num:
        left, hang = OPT_IND if is_option else STEM_IND if is_stem else (STEM_IND[0], 0)
        ind = ppr.makeelement(qn('w:ind'), {})
        ind.set(qn('w:left'), str(left))
        if hang:
            ind.set(qn('w:hanging'), str(hang))
        ppr.append(ind)

    jc = ppr.makeelement(qn('w:jc'), {})
    jc.set(qn('w:val'), 'both')
    ppr.append(jc)

    if mark_rpr is not None:
        ppr.append(mark_rpr)

    # Only the Devanagari pages are forced to Nirmala UI 10pt. The published
    # booklet sets its ENGLISH pages in Times New Roman 11pt (measured from
    # the official PDF's embedded fonts) - and TNR is so much more compact for
    # Latin text that the official fits 8,711 characters on an English page
    # where Nirmala UI 10pt manages 5,995. Forcing one font on both languages
    # cost the booklet a third of its density.
    if lang == "hi":
        _force_hindi_font(p_el)
    else:
        _force_english_font(p_el)
    _no_bold(p_el)


# Symbol faces carry their own glyphs (Wingdings ticks, math signs); forcing
# Times New Roman on them would print the wrong character.
SYMBOL_FONTS = {"symbol", "wingdings", "wingdings 2", "wingdings 3", "webdings",
                "cambria math", "mt extra"}


def _force_english_font(p_el):
    """Times New Roman on every run of an English page.

    The official English pages are Times New Roman throughout. The raw
    paper occasionally carries a stray face - Test 08's Q14 option (b) was
    set in Gungsuh, a Korean font, and printed in it. Sizes are left alone.
    """
    for r in p_el.iter(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            continue
        rf = rpr.find(qn('w:rFonts'))
        if rf is None:
            continue
        current = (rf.get(qn('w:ascii')) or rf.get(qn('w:hAnsi')) or "").lower()
        if current in SYMBOL_FONTS:
            continue
        for attr in list(rf.attrib):
            if attr.endswith("Theme"):          # a theme font wins over a name
                del rf.attrib[attr]
        for slot in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
            rf.set(qn(slot), ENGLISH_FONT)


def _type_statement_label(p_el, label):
    """Put the list item's number in front of its text as a typed label and
    a tab, in the text's own (un-emphasised) run formatting."""
    runs = [r for r in p_el.iter(qn('w:r'))
            if "".join(t.text or "" for t in r.findall(qn('w:t'))).strip()]
    rpr = copy.deepcopy(runs[0].find(qn('w:rPr'))) if runs and \
        runs[0].find(qn('w:rPr')) is not None else None
    if rpr is not None:
        for tag in ('w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:u'):
            for el in rpr.findall(qn(tag)):
                rpr.remove(el)
    ppr = p_el.find(qn('w:pPr'))
    anchor = ppr if ppr is not None else None
    for piece in ("label", "tab"):
        r = p_el.makeelement(qn('w:r'), {})
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        if piece == "label":
            t = r.makeelement(qn('w:t'), {})
            t.set(qn('xml:space'), 'preserve')
            t.text = label
            r.append(t)
        else:
            r.append(r.makeelement(qn('w:tab'), {}))
        if anchor is not None:
            anchor.addnext(r)
        else:
            p_el.insert(0, r)
        anchor = r


def _list_labels(src_doc):
    """{source paragraph: (rendered label, level)} for every auto-numbered
    list item, counted the way Word counts them (verified against Word's own
    ListString on both Test 08 papers). Bullets and formats this does not
    render are left out, so those paragraphs keep their real numbering.
    Cached on the Document: probe and final booklet must agree."""
    cached = getattr(src_doc, "_dp_list_labels", None)
    if cached is not None:
        return cached
    labels = {}
    try:
        numbering = src_doc.part.numbering_part.element
    except Exception:
        src_doc._dp_list_labels = labels
        return labels
    abstract = {a.get(qn('w:abstractNumId')): a for a in numbering.findall(qn('w:abstractNum'))}
    nums = {}
    for n in numbering.findall(qn('w:num')):
        ref = n.find(qn('w:abstractNumId'))
        overrides = {}
        for ov in n.findall(qn('w:lvlOverride')):
            so = ov.find(qn('w:startOverride'))
            if so is not None:
                overrides[int(ov.get(qn('w:ilvl')))] = int(so.get(qn('w:val')))
        nums[n.get(qn('w:numId'))] = (ref.get(qn('w:val')) if ref is not None else None, overrides)

    def level_def(aid, ilvl):
        a = abstract.get(aid)
        if a is None:
            return None
        return next((l for l in a.findall(qn('w:lvl')) if l.get(qn('w:ilvl')) == str(ilvl)), None)

    def fmt(value, kind):
        if kind in ("decimal",):
            return str(value)
        if kind == "decimalZero":
            return f"{value:02d}"
        if kind in ("lowerLetter", "upperLetter"):
            s = chr(ord('a') + (value - 1) % 26) * ((value - 1) // 26 + 1)
            return s if kind == "lowerLetter" else s.upper()
        if kind in ("lowerRoman", "upperRoman"):
            out, n = "", value
            for v, r in ((1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
                         (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
                         (5, "v"), (4, "iv"), (1, "i")):
                while n >= v:
                    out, n = out + r, n - v
            return out if kind == "lowerRoman" else out.upper()
        return None

    counters = {}            # abstractNumId -> {level: value}
    seen_num = set()
    for p in src_doc.element.body.iter(qn('w:p')):
        ppr = p.find(qn('w:pPr'))
        npr = ppr.find(qn('w:numPr')) if ppr is not None else None
        if npr is None:
            continue
        nid_el, il_el = npr.find(qn('w:numId')), npr.find(qn('w:ilvl'))
        nid = nid_el.get(qn('w:val')) if nid_el is not None else None
        if nid in (None, "0") or nid not in nums:
            continue
        ilvl = int(il_el.get(qn('w:val'))) if il_el is not None else 0
        aid, overrides = nums[nid]
        lvl = level_def(aid, ilvl)
        if lvl is None:
            continue
        state = counters.setdefault(aid, {})
        if nid not in seen_num:
            seen_num.add(nid)
            for lv, start in overrides.items():
                state[lv] = start - 1
        start_el = lvl.find(qn('w:start'))
        start = int(start_el.get(qn('w:val'))) if start_el is not None else 0
        state[ilvl] = state.get(ilvl, start - 1) + 1
        for deeper in [k for k in state if k > ilvl]:
            del state[deeper]                   # a higher level restarts deeper ones
        kind_el = lvl.find(qn('w:numFmt'))
        kind = kind_el.get(qn('w:val')) if kind_el is not None else "decimal"
        text_el = lvl.find(qn('w:lvlText'))
        text = text_el.get(qn('w:val')) if text_el is not None else ""
        if kind == "bullet" or not text:
            continue
        label, ok = text, True
        for k in range(ilvl + 1):
            if f"%{k + 1}" not in label:
                continue
            k_def = level_def(aid, k)
            k_kind = k_def.find(qn('w:numFmt')).get(qn('w:val')) if k_def is not None and \
                k_def.find(qn('w:numFmt')) is not None else "decimal"
            k_start = int(k_def.find(qn('w:start')).get(qn('w:val'))) if k_def is not None and \
                k_def.find(qn('w:start')) is not None else 0
            rendered = fmt(state.get(k, k_start), k_kind)
            if rendered is None:
                ok = False
                break
            label = label.replace(f"%{k + 1}", rendered)
        if ok:
            labels[p] = (label, ilvl)
    src_doc._dp_list_labels = labels
    return labels


def _no_bold(el):
    """No bold anywhere in a question - stem, number, statements, options,
    table cells - as the client asked (2026-10-13).

    Bold is switched OFF explicitly (w:b w:val="0") on every run and on the
    paragraph mark rather than just removed: a raw paragraph or character
    style that is itself bold would otherwise still print bold.
    """
    after = (qn('w:rStyle'), qn('w:rFonts'))
    for r in el.iter(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            rpr = r.makeelement(qn('w:rPr'), {})
            r.insert(0, rpr)
        _bold_off(rpr, after)
    for ppr in el.iter(qn('w:pPr')):
        mark = ppr.find(qn('w:rPr'))
        if mark is not None:
            _bold_off(mark, (qn('w:ins'), qn('w:del'), qn('w:moveFrom'),
                             qn('w:moveTo'), qn('w:rStyle'), qn('w:rFonts')))


def _bold_off(rpr, after):
    """rPr children are order-sensitive: b and bCs go right after
    rStyle / rFonts (and the revision marks a paragraph mark may carry)."""
    for tag in ('w:b', 'w:bCs'):
        for old in rpr.findall(qn(tag)):
            rpr.remove(old)
    pos = sum(1 for c in rpr if c.tag in after)
    for tag in ('w:bCs', 'w:b'):          # inserted in reverse: b, then bCs
        off = rpr.makeelement(qn(tag), {})
        off.set(qn('w:val'), '0')
        rpr.insert(pos, off)


def _force_rpr(rpr):
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = rpr.makeelement(qn('w:rFonts'), {})
        rpr.insert(0, rf)
    for slot in ('w:ascii', 'w:hAnsi', 'w:cs'):
        rf.set(qn(slot), BODY_FONT)
    for tag in ('w:sz', 'w:szCs'):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        el.set(qn('w:val'), BODY_SZ)


def _force_hindi_font(p_el):
    """Nirmala UI 10pt on every run and on the paragraph mark."""
    for r in p_el.iter(qn('w:r')):
        rpr = r.find(qn('w:rPr'))
        if rpr is None:
            rpr = r.makeelement(qn('w:rPr'), {})
            r.insert(0, rpr)
        _force_rpr(rpr)
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = p_el.makeelement(qn('w:pPr'), {})
        p_el.insert(0, ppr)
    mark = ppr.find(qn('w:rPr'))
    if mark is None:
        mark = ppr.makeelement(qn('w:rPr'), {})
        ppr.append(mark)       # rPr is the last pPr child that matters here
    _force_rpr(mark)


# Text width of one booklet column (twips), from the question section:
# Letter 12240 wide, margins 1134 + 1041, a 720 gap between two columns.
COLUMN_W = (12240 - 1134 - 1041 - 720) // 2


def _fit_table_to_column(tbl_el, lang):
    """Size a copied table to exactly one column.

    The raw papers are single-column A4, so their tables carry a full-page
    grid (Test 13's Q85: 3960 + 3960 twips) that ran across the column rule
    or off the paper in the two-column booklet. The table's width is set to
    the column and the grid scaled down to it; cell widths are left to
    Word's autofit, so a wordy column still gets more room than a serial
    number (the raw files set no cell widths - fixing them at the equal grid
    squeezed Q57's 'recommendation' column into one word a line). Cells are
    left-aligned: justified text in a narrow cell stretches its short lines
    ('3.      Fiat'). A floating table is put back into the flow - a
    positioned table drags its anchor to the next page and cannot split,
    which would break the Hindi/English pairing.
    """
    from docx.oxml import parse_xml
    tblpr = tbl_el.find(qn('w:tblPr'))
    if tblpr is None:
        tblpr = tbl_el.makeelement(qn('w:tblPr'), {})
        tbl_el.insert(0, tblpr)
    for tag in ('w:tblpPr', 'w:tblOverlap', 'w:tblW', 'w:tblLayout', 'w:tblInd'):
        for old in tblpr.findall(qn(tag)):
            tblpr.remove(old)

    rows = tbl_el.findall(qn('w:tr'))
    grid = tbl_el.find(qn('w:tblGrid'))
    widths = [int(g.get(qn('w:w')) or 0) for g in grid.findall(qn('w:gridCol'))] \
        if grid is not None else []
    if not widths or not all(widths):
        # no usable grid: one equal share per cell of the widest row
        n = max((sum(int((tc.find(qn('w:tcPr')).find(qn('w:gridSpan')).get(qn('w:val'))
                          if tc.find(qn('w:tcPr')) is not None
                          and tc.find(qn('w:tcPr')).find(qn('w:gridSpan')) is not None
                          else 1))
                     for tc in r.findall(qn('w:tc'))) for r in rows), default=1)
        widths = [1] * n
    total = sum(widths)
    scaled = [COLUMN_W * w // total for w in widths]
    scaled[-1] += COLUMN_W - sum(scaled)            # exact column width

    if grid is not None:
        tbl_el.remove(grid)
    grid = parse_xml(f'<w:tblGrid {NS}>' + "".join(
        f'<w:gridCol w:w="{w}"/>' for w in scaled) + '</w:tblGrid>')
    tblpr.addnext(grid)

    tblpr.append(parse_xml(f'<w:tblW {NS} w:w="{COLUMN_W}" w:type="dxa"/>'))
    tblpr.append(parse_xml(f'<w:tblInd {NS} w:w="0" w:type="dxa"/>'))
    tblpr.append(parse_xml(f'<w:tblLayout {NS} w:type="autofit"/>'))
    # tblPr children are order-sensitive, like pPr: put them in schema order
    rank = {qn(f'w:{t}'): i for i, t in enumerate((
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
        "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
        "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook",
        "tblCaption", "tblDescription", "tblPrChange"))}
    for child in sorted(tblpr, key=lambda c: rank.get(c.tag, len(rank))):
        tblpr.remove(child)
        tblpr.append(child)

    for tr in rows:
        for tc in tr.findall(qn('w:tc')):
            tcpr = tc.find(qn('w:tcPr'))
            if tcpr is None:
                tcpr = tc.makeelement(qn('w:tcPr'), {})
                tc.insert(0, tcpr)
            # an A4 cell width would win over the table width; let Word size it
            for old in tcpr.findall(qn('w:tcW')):
                tcpr.remove(old)
            tcpr.insert(0, parse_xml(f'<w:tcW {NS} w:w="0" w:type="auto"/>'))

    for p in tbl_el.iter(qn('w:p')):
        ppr = p.find(qn('w:pPr'))
        if ppr is None:
            ppr = p.makeelement(qn('w:pPr'), {})
            p.insert(0, ppr)
        jc = ppr.find(qn('w:jc'))
        if jc is None:                    # justified styles reach cells too
            jc = ppr.makeelement(qn('w:jc'), {})
            jc.set(qn('w:val'), 'left')
            # jc precedes textDirection .. rPr / sectPr / pPrChange
            later = [c for c in ppr if c.tag in {qn(f'w:{t}') for t in (
                'textDirection', 'textAlignment', 'textboxTightWrap',
                'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange')}]
            if later:
                later[0].addprevious(jc)
            else:
                ppr.append(jc)
        elif jc.get(qn('w:val')) in ('both', 'distribute'):
            jc.set(qn('w:val'), 'left')
        if lang == "hi":
            _force_hindi_font(p)
        else:
            _force_english_font(p)


def _is_blank_paragraph(el):
    """A paragraph with no text and no picture / object (spaces and
    non-breaking spaces count as nothing)."""
    if el.tag != qn('w:p') or dt.clean(_text(el)):
        return False
    return not any(el.iter(qn('w:drawing'), qn('w:pict'), qn('w:object')))


def _join_bare_number(num_p, text_p):
    """Move `text_p`'s content behind the bare number in `num_p` (with a tab
    between, like every other stem) and drop `text_p`."""
    _strip_trailing_breaks(num_p)
    if num_p.find('.//' + qn('w:tab')) is None:
        num_p.append(parse_xml(f'<w:r {NS}><w:tab/></w:r>'))
    for child in list(text_p):
        if child.tag != qn('w:pPr'):
            num_p.append(child)
    text_p.getparent().remove(text_p)


def _column_break_paragraph():
    """A 1pt paragraph holding only a column break.

    A break typed at the start of a question's own first paragraph would
    leave that paragraph's first line behind and start the column on a
    continuation line (number indented to the text); a normal-sized break
    paragraph would cost a whole line at the foot of a full column.
    """
    return parse_xml(
        f'<w:p {NS}><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" '
        f'w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/>'
        f'</w:rPr></w:pPr><w:r><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/>'
        f'</w:rPr><w:br w:type="column"/></w:r></w:p>')


def _emit_group_blocks(out, table, src_doc, group, lang, tail, break_before,
                       line="271", split=None):
    """Emit one language block of one group, identically for probe and final.

    The page turn is pageBreakBefore on the group's first paragraph - an
    explicit break paragraph adds an empty paragraph that can strand a fully
    blank page when the break lands exactly on a natural page end.

    `split` = index in `group` of the question that opens the right-hand
    column (the official keeps every question whole inside one column);
    `line` is then (left leading, right leading). Without a split the page
    snakes freely and `line` is one leading.
    """
    lines = (line,) if isinstance(line, str) else tuple(line)
    first = True
    para_count = 0
    first_para_of = {}
    labels = _list_labels(src_doc)
    for qi, n in enumerate(group):
        right = bool(split) and qi >= split
        ln = lines[min(int(right), len(lines) - 1)]
        opens_column = qi == 0 or (bool(split) and qi == split)
        if split and qi == split:
            tail.addprevious(_column_break_paragraph())
            para_count += 1
        first_para_of[n] = None
        blocks = list(table[n]["blocks"])
        # blank paragraphs closing a question (Test 08 Q100: the gap before
        # the answer key) only push the column foot down
        while len(blocks) > 1 and _is_blank_paragraph(blocks[-1]):
            blocks.pop()
        # a raw stem typed as "67." in one paragraph and its text in the next
        # (English Test 08 Q67) prints as one stem, as in the official
        merge = (len(blocks) > 1 and blocks[0].tag == qn('w:p')
                 and blocks[1].tag == qn('w:p')
                 and BARE_NUM_RE.match(dt.clean(_text(blocks[0])))
                 and dt.clean(_text(blocks[1])))
        for bi, el in enumerate(blocks):
            if merge and bi == 1:
                continue
            new = dt.copy_block_with_media(out, el, src_doc, before=tail)
            if merge and bi == 0:
                _join_bare_number(new, dt.copy_block_with_media(
                    out, blocks[1], src_doc, before=tail))
            if new.tag == qn('w:p'):
                _strip_trailing_breaks(new)
                _lower_option_labels(new)
                # the question's first paragraph gets the official's gap
                # above it, unless it opens the page or the column
                _normalize_qb(new, lang, line=ln, statement=labels.get(el),
                              question_gap=first_para_of[n] is None
                              and not opens_column)
                if first and break_before:
                    ppr = new.find(qn('w:pPr'))
                    pb = ppr.makeelement(qn('w:pageBreakBefore'), {})
                    ppr.insert(0 if ppr.find(qn('w:pStyle')) is None else 1, pb)
                if first_para_of[n] is None:
                    first_para_of[n] = para_count
                para_count += 1
                first = False
            else:
                if new.tag == qn('w:tbl'):
                    _fit_table_to_column(new, lang)
                    _no_bold(new)
                # Word also counts each table row's end mark; counting only
                # the w:p put every question after the paper's first table a
                # few paragraphs out, so each one measured as "taller than a
                # page" (Test 13 failed from Q58 on, right after Q57's table)
                para_count += dt.word_paragraph_count(new)
                first = False
    return first_para_of, para_count


def _lower_option_labels(p_el):
    """The published booklet uses (a)/(b)/(c)/(d), the raw files use capitals."""
    for t in p_el.iter(qn('w:t')):
        if not t.text:
            continue
        t.text = re.sub(r"\(\s*([A-D])\s*\)",
                        lambda m: "(" + m.group(1).lower() + ")", t.text)


def _page_break(before):
    before.addprevious(parse_xml(
        f'<w:p {NS}><w:r><w:br w:type="page"/></w:r></w:p>'))


def _section_break_continuous(before, cols=2):
    """A continuous section break so Word re-balances the columns."""
    before.addprevious(parse_xml(
        f'<w:p {NS}><w:pPr><w:sectPr>'
        f'<w:type w:val="continuous"/>'
        f'<w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>'
        f'<w:pgMar w:top="1340" w:right="1041" w:bottom="1276" w:left="1134" '
        f'w:header="720" w:footer="432"/>'
        f'<w:cols w:num="{cols}" w:space="720" w:sep="1"/>'
        f'</w:sectPr></w:pPr></w:p>'))


def _write_probe(table, src_doc, lang, numbers, probe_path, groups=None,
                 layout=None, langs=("hi", "en")):
    """One language rendered exactly as the final booklet will render it -
    same copy, same normalization, same page-break mechanism - so the
    measurement cannot drift from the real layout.
    `layout` = {group index: (split, line)} as _emit_group_blocks takes them.
    Returns ({question_no: paragraph_index}, {group: last paragraph index}).
    """
    out = _open_booklet(langs)
    # The probe has to paginate exactly like the final booklet, so it keeps
    # the template's front matter and pours the questions into the same
    # sentinel. Word counts every paragraph, so the questions' indices start
    # after everything that precedes the insertion point.
    tail = _find_sentinel(out)
    if tail is not None:
        offset = _paragraph_offset(out, tail)
    else:
        anchors = dt.clear_body_keep_sections(out)
        place = anchors[0] if anchors else out.element.body.find(qn('w:sectPr'))
        tail = out.element.body.find(qn('w:sectPr'))
        _title_line(out, "[TEST SERIES]", place, size_half=40)
        _title_line(out, "[TEST NAME]", place, size_half=36)
        offset = 3           # two title lines + the section-break paragraph

    order = groups if groups else [list(numbers)]
    first_para_of = {}
    group_last = {}
    para_count = 0
    for gi, group in enumerate(order):
        split, line = (layout or {}).get(gi, (None, "271"))
        # In the bilingual booklet only the HINDI half of the first group
        # shares the title band's page; its English half follows on a fresh
        # page, so the English probe turns the page there too. A one-language
        # booklet opens straight under the band.
        fmap, cnt = _emit_group_blocks(out, table, src_doc, group, lang, tail,
                                       break_before=gi > 0 or (
                                           lang == "en" and len(langs) > 1),
                                       line=line, split=split)
        for n, i in fmap.items():
            first_para_of[n] = (i + para_count) if i is not None else None
        para_count += cnt
        group_last[gi] = para_count - 1
    if tail.tag == qn('w:p'):
        # finish like build(): the back cover on its own page, sentinel gone -
        # otherwise the last page measures the cover as part of a question
        _break_before_back_matter(tail)
        tail.getparent().remove(tail)
    out.save(probe_path)
    return ({n: (i + offset if i is not None else None)
             for n, i in first_para_of.items()},
            {gi: i + offset for gi, i in group_last.items()})


# Page geometry of the question pages, in points, measured on the rendered
# booklet: the header artwork ends above 52, the footer strip starts at 737.4
# and the column rule sits at the page centre.
BODY_Y = (60.0, 737.3)
MID_X = 306.0
TARGET_BOTTOM = 731.0      # where a filled column should end (official: ~725)
FILL_TOL = 5.0             # a column ending within this of the target is full
L_MIN, L_FIT, L_CAP = 240, 245, 384   # official leadings run 245..365 (384)
NAT = 13.0                 # one line at single spacing, either language
FLOW_LOSS = 14.0           # what a column turn inside a question can waste
MAX_Q = 6                  # the official booklet's fullest page
EMPTY_COLUMN = 1000.0      # "leading" charged for a column left empty
TIGHT_COST = 20000.0       # a tight column is used only where it is needed
FLOW_COST = 200000.0       # so is a page whose question snakes across columns


def _join(lang, tight):
    """Space after a question's last paragraph + the gap above the next."""
    if tight:
        return (int(TIGHT_SPACING[lang][1]) + int(TIGHT_SPACING[lang][2])) / 20.0
    return (int(SPACING["after"]) + int(QUESTION_GAP)) / 20.0


def _column_extents(page):
    """{top, lb, rb, rtop} of the question text on one PDF page (None = empty).

    Text lines, pictures and table rules all count; the vertical column rule,
    the header/footer furniture and anything crossing the page centre (the
    first page's full-width title band) do not.
    """
    boxes = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] == 0:
            for ln in b["lines"]:
                if "".join(s["text"] for s in ln["spans"]).strip():
                    boxes.append(tuple(ln["bbox"]))
        else:
            boxes.append(tuple(b["bbox"]))
    for d in page.get_drawings():
        r = d["rect"]
        if r.width > 2:
            boxes.append((r.x0, r.y0, r.x1, r.y1))
    boxes = [b for b in boxes if b[1] > BODY_Y[0] and b[3] < BODY_Y[1]]
    left = [b for b in boxes if b[2] < MID_X]
    right = [b for b in boxes if b[0] > MID_X]
    return {"top": min(b[1] for b in left) if left else None,
            "lb": max(b[3] for b in left) if left else None,
            "rtop": min(b[1] for b in right) if right else None,
            "rb": max(b[3] for b in right) if right else None}


def _render(table, src_doc, lang, numbers, workdir, groups, layout, cancelled,
            langs=("hi", "en")):
    """Render one language's probe. Per group: the page its first and last
    paragraphs land on, the page the right-hand column's first question
    lands on, and the column extents of that page (and of the next page when
    the group spilled onto it)."""
    import time
    import fitz
    probe = os.path.join(workdir, f"_probe_{lang}.docx")
    pdf = os.path.join(workdir, f"_probe_{lang}.pdf")
    idx, last_idx = _write_probe(table, src_doc, lang, numbers, probe,
                                 groups, layout, langs)
    wanted = ([i for i in idx.values() if i is not None]
              + list(last_idx.values()))
    for attempt in range(3):            # Word fails transiently now and then
        try:
            got, _total = dt.pages_of_paragraphs(probe, wanted, cancelled,
                                                 pdf_path=pdf)
            break
        except InterruptedError:
            raise
        except Exception:
            if attempt == 2:
                raise
            # A Word that hung on Open keeps the probe locked, so every retry
            # of the same file failed too: drop it (as docx_to_pdf's retry
            # does) and retry under a fresh name.
            from .convert import _kill_orphan_word
            _kill_orphan_word()
            time.sleep(3 * (attempt + 1))
            fresh = os.path.join(workdir, f"_probe_{lang}_{attempt + 1}.docx")
            os.replace(probe, fresh)
            probe = fresh
    doc = fitz.open(pdf)
    cache = {}

    def extents(pno):
        if pno not in cache:
            cache[pno] = (_column_extents(doc[pno - 1])
                          if pno and 1 <= pno <= len(doc) else None)
        return cache[pno]

    res = []
    for gi, group in enumerate(groups):
        split = layout[gi][0]
        first = got.get(idx.get(group[0]))
        last = got.get(last_idx[gi])
        rpage = got.get(idx.get(group[split])) if split else None
        res.append({"first": first, "last": last, "right_page": rpage,
                    "page": extents(first),
                    "next": extents(first + 1) if first and last != first else None})
    doc.close()
    for f in (probe, pdf):
        try:
            os.remove(f)
        except OSError:
            pass
    return res


LANG_LABEL = {"en": "English", "hi": "Hindi"}


def _oversize_message(oversized, single=False):
    """Say WHICH question cannot be made to fit, and what to do about it.

    "did not converge - please report this file" told the user nothing they
    could act on. The measurement only fails when a single question is taller
    than a printed page even at the official's tightest leading, so name it:
    that is almost always one huge table or picture, or a numbering slip in
    the raw file that made one question swallow the ones after it.
    """
    if not oversized:
        return ("Page-fit measurement could not settle. No single question is "
                "oversized, so this is a bug - please send the two files.")
    bits = []
    for n in sorted(oversized):
        langs = " and ".join(LANG_LABEL.get(l, l) for l in sorted(oversized[n]))
        bits.append(f"Q{n} ({langs})")
    which = "; ".join(bits)
    why = ("cannot be printed whole" if single
           else "cannot keep Hindi and English on matching pages")
    return (f"These questions are taller than one page, so the booklet {why}: "
            f"{which}. "
            f"Usually one of two things: the question really does carry a very "
            f"large table or image (shrink it in the source file), or the "
            f"question numbering slips there - a missing, repeated or "
            f"out-of-order number makes one question swallow the ones after "
            f"it, so check the numbering around {which.split(' ')[0]} in "
            f"{'the file' if single else 'both files'}.")


def _flow_height(e):
    """Height of the content on one page, reading the left column and then
    the right one (a question can snake into the right column)."""
    if not e or e["top"] is None:
        return 0.0
    h = e["lb"] - e["top"]
    if e["rb"] is not None:
        h += e["rb"] - e["rtop"]
    return h


def _lstar(m, i, j, cap, tight=False):
    """The leading at which questions i..j-1 stacked in one column end
    exactly `cap` points below the column top (heights are linear in the
    leading: a + b * L/240 per question, plus the gap between questions)."""
    k = j - i
    a = (sum(m["at" if tight else "a"][i:j])
         + (k - 1) * (m["joint" if tight else "join"] - NAT))
    b = sum(m["b"][i:j]) + (k - 1) * NAT
    return 240.0 * (cap - a) / b if b > 0 else math.inf


def _slope(m, i, j):
    """Points gained per unit of leading for questions i..j-1 in one column."""
    return (sum(m["b"][i:j]) + (j - i - 1) * NAT) / 240.0


def _fill_column(m, i, j, cap):
    """(leading, tight) that takes questions i..j-1 to the column foot:
    normal paragraph spacing whenever that fits at the official's minimum
    leading, its tight spacing only when it does not. (None, None) = no fit."""
    L = _lstar(m, i, j, cap)
    if L >= L_FIT:
        return L, False
    L = _lstar(m, i, j, cap, tight=True)
    return (L, True) if L >= L_FIT else (None, None)


def _plan_page(model, i, j):
    """Best layout of questions i..j-1 on one page pair, or None.

    Per language: the column split whose two columns need the most similar
    (and lowest) leading to reach the page foot - both columns then end at
    the bottom, as on every page of the official. A page holding a question
    taller than one column snakes freely instead.
    Returns (cost, {lang: (split or None, [leading], [slope], [tight])}),
    one entry per column.
    """
    def cost(L, tight):
        return L * L + (TIGHT_COST if tight else 0.0)

    total, choice = 0.0, {}
    for lang, m in model.items():
        cap = m["cap_first"] if i == 0 else m["cap"]
        # snaking: a question may carry on into the right-hand column. Needed
        # for a question taller than a column; otherwise only a fallback for
        # a page whose questions cannot be split into two whole columns
        # (the official manages that by hand-tuning to the half point).
        L, t = _fill_column(m, i, j, 2 * cap - FLOW_LOSS)
        flow = None if L is None else (
            2 * cost(L, t) + (0.0 if any(m["tall"][i:j]) else FLOW_COST),
            (None, [L], [_slope(m, i, j)], [t]))
        if any(m["tall"][i:j]):
            if flow is None:
                return None
            total += flow[0]
            choice[lang] = flow[1]
            continue
        best = None
        for s in range(i + 1, j + 1):
            Ll, tl = _fill_column(m, i, s, cap)
            if s < j:
                Lr, tr = _fill_column(m, s, j, cap)
            else:
                Lr, tr = EMPTY_COLUMN, False
            if Ll is None or Lr is None:
                continue
            c = cost(Ll, tl) + cost(Lr, tr)
            if best is None or c < best[0]:
                if s < j:
                    best = (c, (s - i, [Ll, Lr],
                                [_slope(m, i, s), _slope(m, s, j)], [tl, tr]))
                else:
                    best = (c, (None, [Ll], [_slope(m, i, s)], [tl]))
        if best is None:
            best = flow
        if best is None:
            return None
        total += best[0]
        choice[lang] = best[1]
    return total, choice


def _plan_pages(model, n, banned):
    """Fewest page pairs, then the most even fill: dynamic programming over
    where each page pair starts. Returns [(i, j, choice)] or None."""
    INF = (math.inf, math.inf)
    best = [INF] * (n + 1)
    best[0] = (0, 0.0)
    back = [None] * (n + 1)
    for j in range(1, n + 1):
        for i in range(max(0, j - MAX_Q), j):
            if best[i] == INF or (i, j) in banned:
                continue
            plan = _plan_page(model, i, j)
            if plan is None:
                continue
            cand = (best[i][0] + 1, best[i][1] + plan[0])
            if cand < best[j]:
                best[j], back[j] = cand, (i, plan[1])
    if best[n] == INF:
        return None
    pages, j = [], n
    while j:
        i, choice = back[j]
        pages.append((i, j, choice))
        j = i
    return pages[::-1]


def measure_groups(en_by, hi_by, en_doc, hi_doc, numbers, workdir, cancelled,
                   progress=None):
    """Group questions into page pairs and set each column's leading.

    A one-language booklet passes None for the other language's table and
    document: the same planning then runs on that language alone, and a
    "page pair" is simply a page.

    The official booklet keeps every question whole inside one column and
    tunes the leading column by column so that BOTH columns end at the page
    foot (Q41-42 at 329 beside Q43-45 at 250, ...). Same here, measured:

    1. Every question is rendered alone at two leadings, which gives its
       height as a straight line in the leading (wrapping does not change
       with leading, so this is exact up to rounding), and once with the
       official's tight paragraph spacing, used only on columns that need it.
       A page that cannot be split into two whole columns may let one
       question run on into the right column (snaking) rather than cost a
       page; the official avoids that only by hand-tuning to the half point.
    2. Dynamic programming picks the page pairs: fewest pages first, then
       the most even fill, so no page pair is left carrying a lone question.
       Each language picks its own column split, as the official does.
    3. Each column's leading is set to reach the page foot and verified by
       real renders; a column that spills or falls short is corrected from
       what was measured. Pages are independent (each opens with a page
       break, the columns with a column break), so the corrections never
       disturb one another.

    Returns (groups, layout) with layout[(group, lang)] = (split, lines).
    """
    os.makedirs(workdir, exist_ok=True)
    sources = tuple(s for s in (("en", en_by, en_doc), ("hi", hi_by, hi_doc))
                    if s[1] is not None)
    langs = tuple(l for l, _t, _d in sources)
    # the probes are arranged like the booklet being made
    booklet = ("hi", "en") if len(langs) > 1 else langs
    single = len(langs) == 1
    n = len(numbers)
    oversized = {}

    def tick(msg, pct):
        if cancelled():
            raise InterruptedError("stopped")
        if progress:
            progress(msg, pct)

    # 1. Heights. The render at the official's minimum leading is also the
    # pre-flight: a question that does not fit a page even there can never
    # be paired, and the user is told which one at once.
    singles = [[q] for q in numbers]
    model = {}
    for lang, table, src in sources:
        hs = {}
        # normal spacing at two leadings (the slope), tight spacing at one
        for spec in (str(L_MIN), "330", f"{L_MIN}t"):
            tick(f"Measuring {LANG_LABEL[lang]} questions", 60)
            lay = {gi: (None, spec) for gi in range(n)}
            hs[spec] = _render(table, src, lang, numbers, workdir, singles,
                               lay, cancelled, booklet)
        base = hs[str(L_MIN)]
        for q, r in zip(numbers, base):
            if not r["first"] or r["last"] != r["first"]:
                oversized.setdefault(q, set()).add(lang)
        if oversized:
            continue
        tops = [r["page"]["top"] for r in base if r["page"]["top"]]
        top_normal = sorted(tops)[len(tops) // 2]
        top_first = base[0]["page"]["top"] or top_normal
        cap = TARGET_BOTTOM - top_normal

        def height(r, ref):
            """Flow height measured from the reference render's top, so a
            first line pushed down by other spacing counts too."""
            return (_flow_height(r["page"])
                    + (r["page"]["top"] or 0) - (ref["page"]["top"] or 0))

        a, at, b, tall = [], [], [], []
        for r1, r2, rt in zip(base, hs["330"], hs[f"{L_MIN}t"]):
            h1 = _flow_height(r1["page"])
            if r2["first"] and r2["last"] == r2["first"]:
                h2 = height(r2, r1)
            else:                           # spilled at 330: assume it all scales
                h2 = h1 * 330 / L_MIN
            slope = max((h2 - h1) / ((330 - L_MIN) / 240.0), 1.0)
            b.append(slope)
            a.append(h1 - slope * L_MIN / 240.0)
            ht = height(rt, r1) if rt["first"] and rt["last"] == rt["first"] else h1
            ht = min(ht, h1)
            at.append(ht - slope * L_MIN / 240.0)
            # taller than a column even with tight spacing: the page must
            # snake (judged by height, not by reaching the right column)
            tall.append(ht > cap)
        model[lang] = {"a": a, "at": at, "b": b, "tall": tall,
                       "join": _join(lang, False), "joint": _join(lang, True),
                       "cap": cap, "cap_first": TARGET_BOTTOM - top_first}
    if oversized:
        raise RuntimeError(_oversize_message(oversized, single))

    # 2 + 3. Plan, render, correct. lo/hi bracket every column's leading
    # between the loosest verified to fit and the tightest seen to spill.
    banned = set()

    def start(pages):
        st = []
        for i, j, choice in pages:
            for lang in langs:
                split, Ls, slopes, tights = choice[lang]
                st.append({"i": i, "j": j, "lang": lang, "split": split,
                           "slopes": slopes, "tight": tights,
                           "try": [max(L_MIN, min(L_CAP, int(L) - 4)) for L in Ls],
                           "lo": [None] * len(Ls), "hi": [L_CAP + 1] * len(Ls)})
        return st

    def specs(s, key):
        return [f"{L}{'t' if t else ''}" for L, t in zip(s[key], s["tight"])]

    pages = _plan_pages(model, n, banned)
    if pages is None:
        raise RuntimeError(_oversize_message(oversized, single))
    state = start(pages)
    for rnd in range(10):
        tick(f"Fitting columns (round {rnd + 1}, {len(pages)} page pairs)",
             min(62 + 2 * rnd, 72))
        groups = [numbers[i:j] for i, j, _c in pages]
        results = {}
        for lang, table, src in sources:
            mine = [s for s in state if s["lang"] == lang]
            lay = {gi: (s["split"], specs(s, "try"))
                   for gi, s in enumerate(mine)}
            for s, r in zip(mine, _render(table, src, lang, numbers, workdir,
                                          groups, lay, cancelled, booklet)):
                results[id(s)] = r

        failed, moving = set(), False
        for s in state:
            r = results[id(s)]
            e = r["page"] or {}
            spilled = r["last"] != r["first"]
            left_over = bool(s["split"]) and r["right_page"] != r["first"]
            if spilled or left_over:
                col = 0 if (left_over or not s["split"]) else 1
                if col == 0 and s["split"]:
                    # the left column's tail sits at the top of the right one
                    over = (e["rb"] - e["rtop"]) if e.get("rb") else NAT
                else:
                    over = _flow_height(r["next"]) or NAT
                if col == 1:               # the left column did fit
                    s["lo"][0] = max(s["lo"][0] or 0, s["try"][0])
                s["hi"][col] = min(s["hi"][col], s["try"][col])
                if s["try"][col] <= L_MIN:
                    failed.add((s["i"], s["j"]))
                    continue
                step = math.ceil((over + 3) / s["slopes"][col])
                s["try"][col] = max(L_MIN, min(s["try"][col] - step,
                                               s["hi"][col] - 1))
                if s["lo"][col] is not None:
                    s["try"][col] = max(s["try"][col], s["lo"][col])
                    if s["try"][col] >= s["hi"][col]:
                        s["try"][col] = s["lo"][col]
                moving = True
                continue
            # fitted: every column's leading is now known to be safe
            for col in range(len(s["try"])):
                s["lo"][col] = max(s["lo"][col] or 0, s["try"][col])
                if s["split"]:
                    bottom = e.get("lb") if col == 0 else e.get("rb")
                else:
                    # one column, or one snaking flow: aim its end at the foot
                    bottom = e.get("rb") if e.get("rb") is not None else e.get("lb")
                slack = TARGET_BOTTOM - bottom if bottom is not None else 0
                if slack <= FILL_TOL or s["try"][col] >= L_CAP:
                    continue
                want = s["try"][col] + int((slack - 1.5) / s["slopes"][col])
                want = min(want, L_CAP, s["hi"][col] - 1)
                if want > s["try"][col]:
                    s["try"][col] = want
                    moving = True

        if failed:
            # the measured heights promised a fit the render did not give:
            # plan again without those page ranges
            banned |= failed
            pages = _plan_pages(model, n, banned)
            if pages is None:
                break
            state = start(pages)
            continue
        if not moving:
            break

    # Every column keeps the loosest leading that was verified to fit. A page
    # range never verified (budget spent, or no plan left) falls back to one
    # question a page at the minimum leading - the pre-flight render proved
    # each of those fits.
    groups, layout = [], {}
    by_page = {}
    for s in (state if pages else []):
        by_page.setdefault((s["i"], s["j"]), {})[s["lang"]] = s
    for (i, j), per in sorted(by_page.items()):
        if all(None not in s["lo"] for s in per.values()):
            gi = len(groups)
            groups.append(numbers[i:j])
            for lang, s in per.items():
                layout[(gi, lang)] = (s["split"], specs(s, "lo"))
        else:
            for q in numbers[i:j]:
                gi = len(groups)
                groups.append([q])
                for lang in langs:
                    layout[(gi, lang)] = (None, [str(L_MIN)])
    if not groups or sum(len(g) for g in groups) != n:
        raise RuntimeError(_oversize_message(oversized, single))
    return groups, layout


QUESTION_SENTINEL = "[[QUESTIONS]]"

# Strings the official sample carries in its cover / instruction pages. They
# name THAT test, so every booklet has to have them rewritten; the parts that
# are the same in every booklet (product name, marks, instructions) stay.
SAMPLE_TEST_NO = "08"
SAMPLE_SUBJECT = "Indian Polity (Part II)"
# The rest of the sample's identity, printed on the English and the Hindi
# cover alike. Each is rewritten only when the form supplies the value.
SAMPLE_SERIES = "Prelims Test Series 1.0"          # cover, line 1
SAMPLE_BAND_SERIES = "Prelims Test Series - 2027"  # title band over the questions
SAMPLE_TYPE = "Full Length Test"
TARGET_RE = re.compile(r"^TARGET\s*[-–—]\s*20\d\d$")
# 'Time: 120 Hours<tab>Total Marks: 200' / 'समय: 120 मिनट<tab>पूर्णांक: 200'
TIME_EN_RE = re.compile(r"^Time\s*:.*Total Marks\s*:", re.S)
TIME_HI_RE = re.compile(r"^समय\s*:.*पूर्णांक\s*:", re.S)
# instruction 2: 'Each Question carries two marks.' / '... के लिए दो अंक हैं।'
# - the number word is a bold run of its own in both languages
EN_NUMBER_WORDS = ["zero", "one", "two", "three", "four", "five", "six",
                   "seven", "eight", "nine", "ten"]
HI_NUMBER_WORDS = ["शून्य", "एक", "दो", "तीन", "चार", "पाँच", "छह", "सात",
                   "आठ", "नौ", "दस"]


def _number_word(value, words):
    """'2' -> 'two'; beyond ten, or a fraction like 2.5, stays in digits."""
    n = float(value)
    return words[int(n)] if n.is_integer() and n < len(words) else value


def _set_marks_per_question(p_el, text, per):
    """Rewrite the marks-per-question sentence of instruction 2, keeping the
    bold run the number sits in. Returns True when the paragraph was it."""
    runs = p_el.findall(qn('w:r'))

    def t_of(r):
        return "".join(t.text or "" for t in r.findall(qn('w:t')))

    def set_t(r, value):
        ts = r.findall(qn('w:t'))
        ts[0].text = value
        ts[0].set(qn('xml:space'), 'preserve')
        for t in ts[1:]:
            t.text = ""

    one = float(per) == 1
    if "Each Question carries" in text:
        # the sentence's own number - 'equal marks.' earlier is bold too
        lead = next(i for i, r in enumerate(runs)
                    if "Each Question carries" in t_of(r))
        for r in runs[lead + 1:]:
            if re.fullmatch(r"\s*\S+ marks?\.\s*", t_of(r)):
                word = _number_word(per, EN_NUMBER_WORDS)
                set_t(r, f" {word} mark{'' if one else 's'}.")
                return True
    elif "प्रत्येक प्रश्न के लिए" in text:
        # the last 'अंक' - the first is 'सभी प्रश्नों के अंक समान हैं'
        for i in range(len(runs) - 1, 1, -1):
            r = runs[i]
            if t_of(r).strip() == "अंक":
                # 'लिए', ' ', '<number>', ' ', 'अंक', ' ', 'हैं।'
                set_t(runs[i - 2], _number_word(per, HI_NUMBER_WORDS))
                for later in runs[i + 1:]:
                    if t_of(later).strip() in ("हैं।", "है।"):
                        set_t(later, "है।" if one else "हैं।")
                return True
    return False
# '_' counts as a word character, so \b never fires in 'Test 08_Q_...'
TEST_NO_RE = re.compile(r"(?<![A-Za-z0-9])test\s*[-–—:]?\s*0*(\d{1,2})(?!\d)", re.I)


def _shrink_to_hairline(p_el):
    """Make an empty paragraph 1pt high, keeping whatever else its pPr holds
    (a section break, here) in schema order."""
    ppr = p_el.find(qn('w:pPr'))
    if ppr is None:
        ppr = p_el.makeelement(qn('w:pPr'), {})
        p_el.insert(0, ppr)
    for tag in ('w:spacing', 'w:rPr'):
        for old in ppr.findall(qn(tag)):
            ppr.remove(old)
    sp = parse_xml(f'<w:spacing {NS} w:before="0" w:after="0" w:line="20" '
                   f'w:lineRule="exact"/>')
    rpr = parse_xml(f'<w:rPr {NS}><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>')
    # spacing sits before ind/jc/...; rPr before sectPr / pPrChange
    later = {qn(t) for t in ('w:ind', 'w:contextualSpacing', 'w:mirrorIndents',
                             'w:suppressOverlap', 'w:jc', 'w:textDirection',
                             'w:textAlignment', 'w:textboxTightWrap',
                             'w:outlineLvl', 'w:divId', 'w:cnfStyle',
                             'w:sectPr', 'w:pPrChange')}
    nxt = next((c for c in ppr if c.tag in later), None)
    if nxt is None:
        ppr.append(sp)
    else:
        nxt.addprevious(sp)
    tail = next((c for c in ppr if c.tag in (qn('w:sectPr'), qn('w:pPrChange'))), None)
    if tail is None:
        ppr.append(rpr)
    else:
        tail.addprevious(rpr)


def _break_before_back_matter(sentinel):
    """Start the Hindi cover + instructions on a page of their own.

    Only continuous section breaks separate them from the questions in the
    template, so the cover ran on under the last English questions (Test 08
    page 51) instead of opening the closing page as in the published booklet.
    """
    for el in sentinel.itersiblings():
        if el.tag != qn('w:p'):
            continue
        if not "".join(t.text or "" for t in el.iter(qn('w:t'))).strip():
            # the empty section-break paragraphs: 1pt high, or they cannot
            # sit under a last page whose columns are filled to the foot and
            # spill onto a blank page of their own
            _shrink_to_hairline(el)
            continue
        ppr = el.find(qn('w:pPr'))
        if ppr is None:
            ppr = el.makeelement(qn('w:pPr'), {})
            el.insert(0, ppr)
        if ppr.find(qn('w:pageBreakBefore')) is None:
            pos = sum(1 for c in ppr if c.tag in (qn('w:pStyle'), qn('w:keepNext'),
                                                   qn('w:keepLines')))
            ppr.insert(pos, ppr.makeelement(qn('w:pageBreakBefore'), {}))
        return


def _has_sect(el):
    ppr = el.find(qn('w:pPr')) if el.tag == qn('w:p') else None
    return ppr is not None and ppr.find(qn('w:sectPr')) is not None


def _drop_back_matter(doc, sentinel):
    """Remove the closing Hindi cover + instructions.

    The questions' own two-column section becomes the document's last one:
    its section-break paragraph (the first after the sentinel) gives its
    sectPr to the body, and every page after it goes. Returns the Hindi
    paragraphs that were taken out, in order, for `_arrange` to reuse.
    """
    body = doc.element.body
    close = next(el for el in sentinel.itersiblings() if _has_sect(el))
    taken = []
    for el in list(close.itersiblings()):
        if el.tag == qn('w:sectPr'):
            continue
        if el.tag == qn('w:p'):
            if not _has_sect(el):
                taken.append(el)
            body.remove(el)
    sect = close.find(qn('w:pPr')).find(qn('w:sectPr'))
    body.replace(body.find(qn('w:sectPr')), sect)   # moves it out of `close`
    body.remove(close)
    return taken


def _arrange(doc, langs):
    """Shape the bilingual template for the booklet being made.

    Bilingual (('hi', 'en')): as it is - English cover + instructions, title
    band, questions, Hindi cover + instructions on the closing page.
    English: the same without the closing Hindi page.
    Hindi: the Hindi cover + instructions take the English ones' place on
    page 1, in the same three continuous sections (cover / instructions /
    mail line), so page 1 and the band page lay out as in the bilingual one.
    """
    if len(langs) > 1:
        return
    sentinel = _find_sentinel(doc)
    if sentinel is None:
        return
    hindi = _drop_back_matter(doc, sentinel)
    if langs != ("hi",):
        return

    def text(el):
        return "".join(t.text or "" for t in el.iter(qn('w:t'))).strip()

    body = doc.element.body
    front = []
    for el in body:
        if el.tag == qn('w:p') and any(b.get(qn('w:type')) == 'page'
                                       for b in el.iter(qn('w:br'))):
            page_turn = el              # page 1 ends here; the band follows
            break
        front.append(el)
    breaks = [el for el in front if _has_sect(el)]      # cover | instructions
    for el in front:
        if el.tag == qn('w:p') and not _has_sect(el):
            body.remove(el)
    instr = next(i for i, p in enumerate(hindi) if text(p) == "अनुदेश")
    mail = next(i for i, p in enumerate(hindi)
                if text(p).startswith("If you have any queries"))
    for part, before in ((hindi[:instr], breaks[0]),
                         (hindi[instr:mail], breaks[1]),
                         (hindi[mail:], page_turn)):
        for p in part:
            before.addprevious(p)


def _open_booklet(langs):
    out = dt.open_template(TEMPLATE)
    _arrange(out, tuple(langs))
    return out


def _find_sentinel(doc):
    for p in doc.element.body.findall(qn('w:p')):
        if "".join(t.text or "" for t in p.iter(qn('w:t'))).strip() == QUESTION_SENTINEL:
            return p
    return None


def _paragraph_offset(doc, before_el):
    """How many paragraphs precede `before_el` in document order.

    Word numbers every w:p, including those nested inside tables, plus one
    end-of-row mark per table row (see dt.word_paragraph_count), and
    `pages_of_paragraphs` uses that same numbering - so this is what turns a
    probe-local paragraph index into a document one.
    """
    count = 0
    for el in doc.element.body.iter(qn('w:p'), qn('w:tr')):
        if el is before_el:
            break
        if any(a.tag == qn('w:txbxContent') for a in el.iterancestors()):
            continue
        count += 1
    return count


def _set_para_text(p_el, text):
    """Replace a paragraph's words, keeping its first run's formatting."""
    runs = [r for r in p_el.findall(qn('w:r'))]
    if not runs:
        return False
    for r in runs[1:]:
        p_el.remove(r)
    keep = runs[0]
    for t in keep.findall(qn('w:t')):
        keep.remove(t)
    for br in keep.findall(qn('w:tab')):
        keep.remove(br)
    node = keep.makeelement(qn('w:t'), {})
    node.set(qn('xml:space'), 'preserve')
    node.text = text
    keep.append(node)
    return True


def _set_tabbed_text(p_el, left, right):
    """Replace the words either side of a paragraph's tab, keeping the tab
    (the time / marks line is one paragraph with a right tab stop) and the
    first run's formatting on each side."""
    side = 0
    first = [True, True]
    # runs only: the pPr's tab-stop list is made of w:tab elements too
    for el in (c for r in p_el.findall(qn('w:r'))
               for c in r if c.tag in (qn('w:t'), qn('w:tab'))):
        if el.tag == qn('w:tab'):
            side = 1
            continue
        el.text = (left, right)[side] if first[side] else ""
        el.set(qn('xml:space'), 'preserve')
        first[side] = False
    return side == 1


def _fill_test_identity(doc, source_path, notes, details=None):
    """Rewrite the sample's identity in the front & back matter.

    Values typed on the form (`details`: series, type, subject, marks, time,
    year) win. Otherwise the test number and subject come from the input file
    name when it carries them (production files are named like
    'Test 08_Q_PTS 2708_Indian Polity Part II_...') or an obvious placeholder
    is left and reported - the raw question paper does not name the test -
    and the rest keeps the template's text.
    """
    details = details or {}
    stem = os.path.splitext(os.path.basename(source_path))[0]
    m = TEST_NO_RE.search(stem)
    test_no = details.get("testno") or (m.group(1).zfill(2) if m else None)
    subject = details.get("subject")
    parts = [p.strip() for p in re.split(r"[_|]", stem) if p.strip()]
    for i, part in enumerate(parts if subject is None else []):
        if TEST_NO_RE.fullmatch(part) and i + 2 < len(parts):
            cand = parts[i + 2]           # Test 08 | A | PTS 2708 | <subject>
            if len(cand) > 4 and not re.fullmatch(r"[A-Z]{2,5}\s*\d{3,5}", cand):
                subject = cand
            break
    if subject is None:
        for part in parts:
            if re.search(r"polity|history|geograph|econom|science|environment|"
                         r"society|governance|csat|current", part, re.I):
                subject = part
                break

    if subject and not details.get("subject"):
        # file names flatten the published '(Part II)' to 'Part II' - same
        # rule as the answer booklet's title line
        subject = re.sub(r"\s+part\s+([IVX0-9]+)\s*$", r" (Part \1)", subject,
                         flags=re.I)

    missing = []
    if test_no is None:
        test_no = "[NN]"
        missing.append("test number")
    if subject is None:
        subject = "[SUBJECT]"
        missing.append("subject")

    from .prelims_ab import series_line
    series = details.get("series")
    year = details.get("year")
    kind = details.get("type")
    time_ = details.get("time")
    marks = details.get("marks")
    per = details.get("permark")

    changed = 0
    for p in doc.element.body.findall(qn('w:p')):
        text = "".join(t.text or "" for t in p.iter(qn('w:t'))).strip()
        if not text:
            continue
        new = None
        if text.startswith("Test ") and SAMPLE_SUBJECT in text:
            new = re.sub(r"0*" + SAMPLE_TEST_NO + r"\b", test_no, text, count=1)
            new = new.replace(SAMPLE_SUBJECT, subject)
        elif text == SAMPLE_SUBJECT:
            new = subject
        elif text == SAMPLE_SERIES and series:
            new = series
        elif text == SAMPLE_BAND_SERIES:
            new = series_line(details, default=SAMPLE_BAND_SERIES, sep=" - ")
        elif text == SAMPLE_TYPE and kind:
            new = f"{kind} Test"
        elif TARGET_RE.match(text) and year:
            new = re.sub(r"20\d\d", year, text)
        elif per and ("Each Question carries" in text
                      or "प्रत्येक प्रश्न के लिए" in text):
            if _set_marks_per_question(p, text, per):
                changed += 1
            continue
        elif (time_ or marks) and (TIME_EN_RE.match(text) or TIME_HI_RE.match(text)):
            t = time_ or re.search(r":\s*(\d+)", text).group(1)
            mk = marks or re.search(r"(\d+)\s*$", text).group(1)
            if TIME_EN_RE.match(text):
                # the sample reads '120 Hours' - the paper is 120 minutes
                left, right = f"Time: {t} Minutes", f"Total Marks: {mk}"
            else:
                left, right = f"समय: {t} मिनट", f"पूर्णांक: {mk}"
            if _set_tabbed_text(p, left, right):
                changed += 1
            continue
        if new and new != text and _set_para_text(p, new):
            changed += 1
    if missing:
        notes.append("Cover and title lines: " + " and ".join(missing) +
                     " not typed and not found in the file name, left as a placeholder to "
                     "type over (the raw paper does not name the test).")
    elif changed:
        notes.append(f"Cover and title lines set to Test {test_no}: {subject} "
                     + ("(subject as typed)." if details.get("subject")
                        else "(read from the file name)."))
    typed = [label for key, label in (("series", "series name"),
                                      ("type", "test type"),
                                      ("marks", "total marks"),
                                      ("permark", "marks per question"),
                                      ("time", "time"), ("year", "target year"))
             if details.get(key)]
    if typed:
        notes.append("Cover set from the form: " + ", ".join(typed) + ".")
    return changed


def build(en_path, hi_path, out_path, start_page=None, end_page=None,
          progress=lambda s, p: None, cancelled=lambda: False, details=None):
    """Bilingual with both papers; an English or Hindi booklet when only
    that paper is given (the other path None)."""
    notes = []
    details = details or {}
    langs = tuple(l for l, p in (("hi", hi_path), ("en", en_path)) if p)
    single = len(langs) == 1

    def read(path, label, pct):
        """Read questions, honouring the page range but never losing content.

        The English and Hindi papers do NOT paginate alike (65 vs 71 pages in
        the sample), so one page range cannot fit both. The question section
        always ends at the answer key, so if the range starves a file we rescan
        it whole and say so. The Document is returned too - the block elements
        belong to it, and copying their pictures later needs its part.
        """
        progress(f"Reading {label} paper", pct)
        doc, blocks = dt.slice_blocks(path, start_page=start_page,
                                      end_page=end_page, cancelled=cancelled)
        found = split_questions(blocks)
        if (start_page or end_page) and not single:
            full_doc, all_blocks = dt.slice_blocks(path, cancelled=cancelled)
            full = split_questions(all_blocks)
            if len(full) > len(found):
                notes.append(
                    f"{label}: the page range covered {len(found)} questions but the "
                    f"file holds {len(full)} - used all of them "
                    f"(the two languages paginate differently).")
                return full_doc, full
        return doc, found

    en_doc, en = read(en_path, "English", 12) if en_path else (None, None)
    hi_doc, hi = read(hi_path, "Hindi", 28) if hi_path else (None, None)
    if cancelled():
        raise InterruptedError("stopped")

    progress("Matching questions", 45)
    if single:
        if not (en or hi):
            raise ValueError("Could not find numbered questions in the file. "
                             "Check the start/end pages.")
    elif not en or not hi:
        raise ValueError("Could not find numbered questions in one of the files. "
                         "Check the start/end pages.")

    en_by = {q["no"]: q for q in en} if en_path else None
    hi_by = {q["no"]: q for q in hi} if hi_path else None
    if single:
        numbers = sorted(en_by or hi_by)
    else:
        numbers = sorted(set(en_by) & set(hi_by))
        only_en = sorted(set(en_by) - set(hi_by))
        only_hi = sorted(set(hi_by) - set(en_by))
        if only_en:
            notes.append("Missing from the Hindi file: " + ", ".join(f"Q{n}" for n in only_en))
        if only_hi:
            notes.append("Missing from the English file: " + ", ".join(f"Q{n}" for n in only_hi))
        if not numbers:
            raise ValueError("The two files share no question numbers.")

    progress("Measuring page fit", 60)
    groups, layout = measure_groups(en_by, hi_by, en_doc, hi_doc, numbers,
                                      os.path.dirname(os.path.abspath(out_path)),
                                      cancelled, progress)

    progress("Building booklet", 75)
    out = _open_booklet(langs)
    source = en_path or hi_path
    # The template's page header still carries the sample's own code
    # ("PTS 2708"), so Test 13 printed Test 08's code on 43 of 44 pages. Same
    # rule as the answer booklet: the code from the file name, else [CODE].
    from .prelims_ab import _replace_header_code
    _replace_header_code(out, source, notes, typed=details.get("code"))

    # The template keeps the official front matter (English cover + Test
    # Instructions + title band) and back matter (the same in Hindi); the
    # questions replace the sentinel paragraph between them.
    tail = _find_sentinel(out)
    if tail is None:                     # template without front matter
        anchors = dt.clear_body_keep_sections(out)
        place = anchors[0] if anchors else out.element.body.find(qn('w:sectPr'))
        _title_line(out, details.get("series") or "[TEST SERIES]", place,
                    size_half=40)
        _title_line(out, details.get("subject") or "[TEST NAME]", place,
                    size_half=36)
        notes.append("Title lines are placeholders - the raw files have no "
                     "test series or test name.")
        tail = out.element.body.find(qn('w:sectPr'))
    else:
        _fill_test_identity(out, source, notes, details)

    first_block = True
    for gi, group in enumerate(groups):
        if cancelled():
            raise InterruptedError("stopped")
        for lang, table, src in (("hi", hi_by, hi_doc), ("en", en_by, en_doc)):
            if table is None:
                continue
            # each language keeps the column split and the per-column
            # leading its own pages were measured at
            split, line = layout.get((gi, lang), (None, "259"))
            _emit_group_blocks(out, table, src, group, lang, tail,
                               break_before=not first_block,
                               line=line, split=split)
            first_block = False

    sentinel = _find_sentinel(out)
    if sentinel is not None:             # its job is done
        _break_before_back_matter(sentinel)
        sentinel.getparent().remove(sentinel)

    progress("Saving", 90)
    out.save(out_path)
    # the cover's total should be questions x marks each; say so if not
    total, per = details.get("marks"), details.get("permark")
    if total and per and abs(float(total) - len(numbers) * float(per)) > 1e-6:
        notes.append(f"Check the marks: {len(numbers)} questions x {per} = "
                     f"{len(numbers) * float(per):g}, but Total marks says {total}.")
    sizes = sorted({len(g) for g in groups})
    if single:
        notes.insert(0, f"{len(numbers)} questions in {LANG_LABEL[langs[0]]} on "
                        f"{len(groups)} pages ({min(sizes)}-{max(sizes)} questions "
                        f"per page, each question whole in one column, line "
                        f"spacing set per column so both columns reach the page "
                        f"foot).")
    else:
        notes.insert(0, f"{len(numbers)} questions in Hindi and English across "
                        f"{len(groups)} page pairs ({min(sizes)}-{max(sizes)} questions "
                        f"per page, each question whole in one column, line spacing "
                        f"set per column so both columns reach the page foot).")
    return out_path, notes


def process(input_path, workdir, variant="Bilingual", start_page=None, end_page=None,
            second_path=None, progress=lambda s, p: None,
            cancelled=lambda: False, details=None, **kw):
    """Bilingual: input_path = English file, second_path = Hindi file.
    English / Hindi: input_path is that language's paper."""
    if variant == "English":
        en_path, hi_path = input_path, None
    elif variant == "Hindi":
        en_path, hi_path = None, input_path
    else:
        if not second_path:
            raise ValueError("This category needs two files: the English paper "
                             "and the Hindi paper.")
        en_path, hi_path = input_path, second_path
    os.makedirs(workdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(input_path))[0]
    out_path = os.path.join(workdir, f"{stem}_Question Booklet_{variant}.docx")
    _, notes = build(en_path, hi_path, out_path,
                     start_page=start_page, end_page=end_page,
                     progress=progress, cancelled=cancelled, details=details)

    # PDF companion beside the DOCX, same stem: the UI offers both downloads
    # (for this category the PDF is what the client actually prints)
    progress("Making the PDF", 95)
    pdf = dt.make_pdf_companion(out_path, notes)
    if pdf:
        # Word re-encodes every JPEG on export; print the raw file's own
        dt.restore_original_jpegs(out_path, pdf)

    progress("Done", 100)
    return out_path, notes
