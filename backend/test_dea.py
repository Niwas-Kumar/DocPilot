# -*- coding: utf-8 -*-
"""Daily Editorial Analysis - variant-input tests.

    python backend/test_dea.py [case ...]

Every case derives a raw file from the REAL samples (UPSC sample folder) by
mutating it the way typists' files actually differ from day to day, runs the
category's pipeline, and checks the result mechanically:

  * a PDF and an editable Word companion are produced
  * page 1 is the cover and carries every topic (and the date)
  * body pages are numbered 2, 3, 4 ... with no gap or repeat
  * a promo banner, when placed, never overlaps a line of text
  * the case's own expectations (day order, QR, headings, swap ...)

Needs Word (the pipelines render through it) and the samples; each case
takes 20 s - 4 min.
"""
import copy
import os
import re
import shutil
import sys
import time
import traceback
import unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, ROOT)

import fitz
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from backend.pipeline import dea_common as C

SAMPLES = os.environ.get("DOCPILOT_SAMPLES") or os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "UPSC sample")
WORK = os.path.join(ROOT, "work", "tests_dea")
EN_DEA = os.path.join(SAMPLES, "Daily Editorial Analysis (English)", "Input August 2026_ Editorial Summary.docx")
HI_DEA = os.path.join(SAMPLES, "Daily Editorial Analysis (Hindi)", "Input 17th August_2026_Hindi_DEA.docx")
EN_CA = os.path.join(SAMPLES, "Daily Current Affairs+Daily Editorial Analysis (English)", "Input_DCA english 17th August 2026 (2).docx")
HI_CA = os.path.join(SAMPLES, "Daily Current Affairs+Daily Editorial Analysis (Hindi)", "Input DCA_Input Hindi 17 August-Daily News Hindi (1).docx")
HI_DEA2 = os.path.join(SAMPLES, "Daily Current Affairs+Daily Editorial Analysis (Hindi)", "Input DEA_Input Hindi 17th August_Editorial Analysis_Hindi_1.docx")
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


# ------------------------------------------------------------------ mutations

def _paras(doc):
    return [el for el in doc.element.body if el.tag == qn('w:p')]


def _titles(doc):
    sz = C.body_size_of(doc, list(doc.element.body))
    return [p for p in _paras(doc) if C.is_topic_title(p, doc, sz)]


def _style_id(doc, name):
    for st in doc.styles:
        if st.name and st.name.lower() == name:
            return st.style_id
    return name.replace(" ", "").title()


def m_heading_styles(doc):
    """Titles typed as 'Heading 1' in plain black - no colour, no size."""
    sid = _style_id(doc, "heading 1")
    for p in _titles(doc):
        ppr = C.ppr_of(p)
        ppr.insert(0, parse_xml(f'<w:pStyle {W} w:val="{sid}"/>'))
        for r in p.iter(qn('w:r')):
            rpr = r.find(qn('w:rPr'))
            if rpr is not None:
                for tag in ('w:color', 'w:sz', 'w:szCs', 'w:b', 'w:bCs'):
                    for el in rpr.findall(qn(tag)):
                        rpr.remove(el)


def m_no_qr(doc):
    for p in _paras(doc):
        if not C.clean(C.para_text(p)) and C.has_drawing(p):
            p.getparent().remove(p)


def m_typed_bullets(doc):
    """Word lists replaced by '• ' typed at the start of plain paragraphs."""
    for p in _paras(doc):
        ppr = p.find(qn('w:pPr'))
        npr = ppr.find(qn('w:numPr')) if ppr is not None else None
        if npr is None:
            continue
        ppr.remove(npr)
        first = next((t for t in p.iter(qn('w:t')) if t.text), None)
        if first is not None:
            first.text = "• " + first.text


def m_not_bold_titles(doc):
    for p in _titles(doc):
        for r in p.iter(qn('w:r')):
            rpr = r.find(qn('w:rPr'))
            if rpr is not None:
                for tag in ('w:b', 'w:bCs'):
                    for el in rpr.findall(qn(tag)):
                        rpr.remove(el)


def m_tabs(doc):
    """Google Docs 'Tabs' export: 'Tab 1' + a hard page break up top."""
    body = doc.element.body
    first = next(el for el in body if el.tag in (qn('w:p'), qn('w:tbl')))
    first.addprevious(parse_xml(f'<w:p {W}><w:r><w:t>Tab 1</w:t></w:r></w:p>'))
    first.addprevious(parse_xml(f'<w:p {W}><w:r><w:br w:type="page"/></w:r></w:p>'))


def m_tables(doc):
    """A 3x3 comparison table and a 1x1 info box inside the first topic."""
    heads = [p for p in _paras(doc) if C.clean(C.para_text(p)).endswith(":") and len(C.clean(C.para_text(p))) < 90]
    anchor = heads[1] if len(heads) > 1 else _paras(doc)[5]
    cells = lambda r: "".join(f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>{t}</w:t></w:r></w:p></w:tc>' for t in r)
    rows = [["Aspect", "Before", "After"], ["Cycle", "Staggered", "Synchronised"], ["Cost", "High", "Lower"]]
    tbl = parse_xml(f'<w:tbl {W}><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid>'
                    + '<w:gridCol w:w="3000"/>' * 3 + '</w:tblGrid>'
                    + "".join(f"<w:tr>{cells(r)}</w:tr>" for r in rows) + '</w:tbl>')
    anchor.addprevious(tbl)
    box = parse_xml(f'<w:tbl {W}><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>'
                    '<w:tr><w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>About ONOE</w:t></w:r></w:p>'
                    '<w:p><w:r><w:t>A proposal to hold Lok Sabha and State Assembly elections together.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
    anchor.addprevious(box)


def m_triple(doc):
    """Three times the topics - a long day."""
    body = doc.element.body
    sect = body.find(qn('w:sectPr'))
    blocks = [el for el in body if el.tag in (qn('w:p'), qn('w:tbl'))]
    for k in (2, 3):
        for el in blocks:
            new = copy.deepcopy(el)
            if el in _titles(doc):
                t = next((t for t in new.iter(qn('w:t')) if t.text), None)
                if t is not None:
                    t.text = f"({k}) " + t.text
            sect.addprevious(new)


def m_two_days_reversed(doc, later="17 August 2026", earlier="16 August 2026"):
    """Two days in one file, the LATER first (as typists prepend): the
    booklet must print the earlier day first, each day from a new page."""
    body = doc.element.body
    sect = body.find(qn('w:sectPr'))
    blocks = [el for el in body if el.tag in (qn('w:p'), qn('w:tbl'))]
    title_sid = _style_id(doc, "title")
    first = blocks[0]
    first.addprevious(parse_xml(f'<w:p {W}><w:pPr><w:pStyle w:val="{title_sid}"/></w:pPr><w:r><w:t>{later}</w:t></w:r></w:p>'))
    sect.addprevious(parse_xml(f'<w:p {W}><w:pPr><w:pStyle w:val="{title_sid}"/></w:pPr><w:r><w:t>{earlier}</w:t></w:r></w:p>'))
    for el in blocks:
        new = copy.deepcopy(el)
        if el in _titles(doc):
            t = next((t for t in new.iter(qn('w:t')) if t.text), None)
            if t is not None:
                t.text = "EARLIER DAY " + t.text
        sect.addprevious(new)


def m_drop_date(doc):
    for p in _paras(doc):
        if C.style_name(doc, p) == "title":
            p.getparent().remove(p)


def m_junk(doc):
    body = doc.element.body
    sect = body.find(qn('w:sectPr'))
    for el in [e for e in body if e is not sect]:
        body.remove(el)
    for t in ("Some notes typed for later.", "Nothing here is an editorial."):
        sect.addprevious(parse_xml(f'<w:p {W}><w:r><w:t>{t}</w:t></w:r></w:p>'))


def make(src, name, *mutations):
    os.makedirs(WORK, exist_ok=True)
    out = os.path.join(WORK, f"in_{name}.docx")
    doc = Document(src)
    for m in mutations:
        m(doc)
    doc.save(out)
    return out


# ------------------------------------------------------------------ checks

def _norm(s):
    return re.sub(r"\s+", " ", unicodedata.normalize("NFC", s or "")).strip()


def _skel(s):
    return "".join(ch for ch in _norm(s).lower() if ch.isalnum())


def page_numbers(pdf, hindi):
    """The stamped number on each page (None where there is none)."""
    out = []
    with fitz.open(pdf) as d:
        for page in d:
            num = None
            for b in page.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    for s in l["spans"]:
                        t = s["text"].strip()
                        in_box = (289 < s["bbox"][0] < 316 and 748 < s["bbox"][1] < 772) if hindi \
                            else (290 < s["bbox"][0] < 322 and 728 < s["bbox"][1] < 752)
                        if t.isdigit() and in_box and s["color"] == 0xFFFFFF:
                            num = int(t)
            out.append(num)
    return out


def banner_clear(pdf):
    """No promo banner overlaps a text line (ignoring page furniture)."""
    bad = []
    with fitz.open(pdf) as d:
        for pno, page in enumerate(d):
            imgs = [fitz.Rect(i["bbox"]) for i in page.get_image_info()
                    if 300 < fitz.Rect(i["bbox"]).width < 600 and 150 < fitz.Rect(i["bbox"]).height < 500
                    and fitz.Rect(i["bbox"]).y0 > 90]
            if not imgs or pno == 0:
                continue
            for b in page.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    r = fitz.Rect(l["bbox"])
                    t = "".join(s["text"] for s in l["spans"]).strip()
                    if not t or r.y1 > 715 or r.y0 < 70:
                        continue
                    for im in imgs:
                        if r.intersects(im) and (r & im).height > 2:
                            bad.append((pno + 1, t[:30]))
    return bad


class Case:
    def __init__(self, name, module, inp, second=None, expect_topics=None, kwargs=None,
                 hindi=False, check=None):
        self.name, self.module, self.inp, self.second = name, module, inp, second
        self.expect_topics, self.kwargs = expect_topics, kwargs or {}
        self.hindi, self.check = hindi, check


def run(case):
    import importlib
    mod = importlib.import_module(f"backend.pipeline.{case.module}")
    wd = os.path.join(WORK, case.name)
    shutil.rmtree(wd, ignore_errors=True)
    os.makedirs(wd)
    src = shutil.copy2(case.inp, os.path.join(wd, os.path.basename(case.inp)))
    kw = dict(case.kwargs)
    if case.second:
        kw["second_path"] = shutil.copy2(case.second, os.path.join(wd, "second_" + os.path.basename(case.second)))
    t0 = time.time()
    out, notes = mod.process(src, wd, **kw)
    dt = time.time() - t0
    problems = []
    if not os.path.exists(out):
        return [f"no PDF"], notes, dt
    comp = os.path.splitext(out)[0] + ".docx"
    if not os.path.exists(comp):
        problems.append("no Word companion")
    else:
        try:
            Document(comp)
        except Exception as exc:
            problems.append(f"companion does not open ({exc})")
    nums = page_numbers(out, case.hindi)
    body = [n for n in nums[1:] if n is not None]
    if body and body != list(range(2, 2 + len(body))):
        problems.append(f"page numbers not 2..N: {nums}")
    if not any("Cover page added" in n for n in notes):
        problems.append("no cover")
    with fitz.open(out) as d:
        cover_text = _skel(d[0].get_text())
    if case.expect_topics and not case.hindi:
        for t in case.expect_topics:
            if _skel(t)[:25] not in cover_text:
                problems.append(f"cover lacks topic '{t[:40]}'")
    if case.hindi:
        # Devanagari comes back out of a PDF scrambled (matras reorder - the
        # published PDFs do it too), so count the index entries instead
        m = next((re.search(r"(\d+) (?:editorial )?topic", n) for n in notes
                  if re.search(r"\d+ editorial topic", n)), None)
        want = int(m.group(1)) if m else 0
        with fitz.open(out) as d:
            entries = [l for b in d[0].get_text("dict")["blocks"] for l in b.get("lines", [])
                       if l["spans"] and (re.match(r"^\s*\d+\s*\.", "".join(s["text"] for s in l["spans"]))
                                          or any("wingdings" in s["font"].lower() for s in l["spans"]))]
        if want and len(entries) < want:
            problems.append(f"cover index has {len(entries)} entries, expected >= {want}")
    overlaps = banner_clear(out)
    if overlaps:
        problems.append(f"banner overlaps text: {overlaps[:3]}")
    if case.check:
        try:
            problems += case.check(out, notes) or []
        except Exception as exc:
            problems.append(f"check crashed: {exc}")
    return problems, notes, dt


# ------------------------------------------------------------------ expectations

def _first_topic_pages(out, marker):
    with fitz.open(out) as d:
        return [i + 1 for i, p in enumerate(d) if marker in p.get_text()]


def check_two_days(out, notes):
    p = []
    if not any("oldest first" in n for n in notes):
        p.append("no day-order note")
    early = _first_topic_pages(out, "EARLIER DAY")
    with fitz.open(out) as d:
        later = [i + 1 for i, pg in enumerate(d) if "India’s Democracy needs" in pg.get_text()
                 and "EARLIER DAY" not in pg.get_text() and i > 0]
    if not early or not later or min(early) > min(later):
        p.append(f"earlier day not printed first (earlier on {early[:2]}, later on {later[:2]})")
    return p


def check_no_qr(out, notes):
    return [] if any("no video QR" in n for n in notes) else ["no note about the missing QR"]


def check_typed_link(out, notes):
    import zxingcpp
    from PIL import Image
    with fitz.open(out) as d:
        pix = d[0].get_pixmap(dpi=200)
    got = [r.text for r in zxingcpp.read_barcodes(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))]
    return [] if "https://youtu.be/TYPED-LINK-123" in got else [f"cover QR is {got}"]


def check_file_qr(out, notes):
    import zxingcpp
    from PIL import Image
    with fitz.open(out) as d:
        pix = d[0].get_pixmap(dpi=200)
    got = [r.text for r in zxingcpp.read_barcodes(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))]
    return [] if "https://youtu.be/fy3CfVE0J2g" in got else [f"cover QR is {got}"]


def check_swapped(out, notes):
    return [] if any("looked swapped" in n for n in notes) else ["swap not detected"]


def check_junk(out, notes):
    return [] if any("No topic titles found" in n for n in notes) else ["no warning for a non-editorial file"]


def check_date_missing(out, notes):
    return [] if any("date MISSING" in n for n in notes) else ["no missing-date note"]


def check_tables(out, notes):
    return [] if any("table(s)" in n for n in notes) else ["tables not reported"]


EN_TOPICS = ["India’s Democracy needs different Electoral Rhythms",
             "Key Sectors to Position India as a Global Manufacturing Hub",
             "India is losing West Asia, but it is not too Late"]
HI_TOPICS = ["भारतीय लोकतंत्र को एक भिन्न चुनावी प्रक्रियागत प्रणाली की आवश्यकता है"]
DATE = {"cover_date": "17 August 2026"}


def cases():
    return [
        Case("en_sample", "dea_english", EN_DEA, expect_topics=EN_TOPICS, kwargs=DATE, check=check_file_qr),
        Case("en_heading_styles", "dea_english", make(EN_DEA, "en_heading_styles", m_heading_styles),
             expect_topics=EN_TOPICS, kwargs=DATE),
        Case("en_no_qr", "dea_english", make(EN_DEA, "en_no_qr", m_no_qr), expect_topics=EN_TOPICS,
             kwargs=DATE, check=check_no_qr),
        Case("en_typed_link", "dea_english", EN_DEA, expect_topics=EN_TOPICS,
             kwargs=dict(DATE, cover_qr="https://youtu.be/TYPED-LINK-123"), check=check_typed_link),
        Case("en_typed_bullets", "dea_english", make(EN_DEA, "en_typed_bullets", m_typed_bullets),
             expect_topics=EN_TOPICS, kwargs=DATE),
        Case("en_not_bold_titles", "dea_english", make(EN_DEA, "en_not_bold", m_not_bold_titles),
             expect_topics=EN_TOPICS, kwargs=DATE),
        Case("en_tabs_tables", "dea_english", make(EN_DEA, "en_tabs_tables", m_tabs, m_tables),
             expect_topics=EN_TOPICS, kwargs=DATE, check=check_tables),
        Case("en_two_days", "dea_english", make(EN_DEA, "en_two_days", m_two_days_reversed),
             expect_topics=EN_TOPICS[:1], kwargs=DATE, check=check_two_days),
        Case("en_triple", "dea_english", make(EN_DEA, "en_triple", m_triple),
             expect_topics=EN_TOPICS, kwargs=DATE),
        Case("en_junk", "dea_english", make(EN_DEA, "en_junk", m_junk), kwargs=DATE, check=check_junk),
        Case("hi_sample", "dea_hindi", HI_DEA, expect_topics=HI_TOPICS, hindi=True),
        Case("hi_heading_styles", "dea_hindi", make(HI_DEA, "hi_heading_styles", m_heading_styles),
             expect_topics=HI_TOPICS, hindi=True),
        Case("hi_no_date", "dea_hindi", make(HI_DEA, "hi_no_date", m_drop_date), hindi=True,
             check=check_date_missing),
        Case("hi_tables", "dea_hindi", make(HI_DEA, "hi_tables", m_tables), hindi=True, check=check_tables),
        Case("hi_triple", "dea_hindi", make(HI_DEA, "hi_triple", m_triple), hindi=True),
        Case("combo_en_swapped", "combo_english", EN_DEA, second=EN_CA, check=check_swapped),
        Case("combo_hi_single_day", "combo_hindi", HI_CA, second=HI_DEA, hindi=True,
             kwargs={"cover_date": "17 August 2026"}),
    ]


def main():
    wanted = set(sys.argv[1:])
    results = []
    for case in cases():
        if wanted and case.name not in wanted:
            continue
        try:
            problems, notes, dt = run(case)
        except Exception as exc:
            problems, notes, dt = [f"CRASH {type(exc).__name__}: {exc}"], [], 0
            traceback.print_exc()
        ok = not problems
        results.append(ok)
        print(f"{'PASS' if ok else 'FAIL'}  {case.name:22s} {dt:5.0f}s" + ("" if ok else "  <- " + "; ".join(problems)),
              flush=True)
    print(f"\n{sum(results)} passed, {len(results) - sum(results)} failed")
    return 0 if all(results) else 1


if __name__ == "__main__":
    sys.exit(main())
