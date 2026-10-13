# -*- coding: utf-8 -*-
"""Bug-hunt suite for the Test Series categories.

Pipeline-level (no server needed):  python backend\test_series.py
"""
import os
import sys
import traceback

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx import Document
from docx.shared import Pt

from backend.pipeline import docx_tools as dt
from backend.pipeline import prelims_ab, prelims_qb, mains_ab

S = os.environ.get("DOCPILOT_SAMPLES") or os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "UPSC sample")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORK = os.path.join(ROOT, "work", "tests_series")
EDGE = os.path.join(WORK, "edge")
os.makedirs(EDGE, exist_ok=True)

PRE_EN = os.path.join(S, "Prelims Test Series (Answer Booklet English)", "Input English.docx")
PRE_HI = os.path.join(S, "Prelims Test Series (Answer Booklet Hindi)", "Input Hindi.docx")
QB_EN = os.path.join(S, "Prelims Test Series (Question Booklet Bilingual)", "Input English.docx")
QB_HI = os.path.join(S, "Prelims Test Series (Question Booklet Bilingual)", "Input Hindi.docx")
MAINS = os.path.join(S, "Mains Test Series (Answer Booklet Hindi)", "Input SDM 2719 Hind  (1).docx")

PASS, FAIL = [], []


def assert_numbering_resolves(path):
    """Every w:numId must exist in numbering.xml - a dangling id makes Word
    silently substitute decimal numbering, so bullets print as 1./2."""
    import re
    import zipfile
    z = zipfile.ZipFile(path)
    doc = z.read("word/document.xml").decode("utf8", "replace")
    num = z.read("word/numbering.xml").decode("utf8", "replace")
    used = set(re.findall(r'<w:numId w:val="(\d+)"', doc))
    defined = set(re.findall(r'<w:num w:numId="(\d+)"', num))
    dangling = sorted(used - defined, key=int)
    assert not dangling, f"dangling numIds {dangling[:6]}"


def assert_no_dangling_images(path):
    """Every body picture reference must resolve to a real relationship."""
    import re
    import zipfile
    z = zipfile.ZipFile(path)
    doc = z.read("word/document.xml").decode("utf8", "replace")
    embeds = set(re.findall(r'r:embed="([^"]+)"', doc))
    rels = z.read("word/_rels/document.xml.rels").decode("utf8", "replace")
    known = set(re.findall(r'Id="([^"]+)"', rels))
    dangling = sorted(e for e in embeds if e not in known)
    assert not dangling, f"{len(dangling)} dangling picture refs: {dangling[:4]}"
    return len(embeds)


def check(name, fn):
    try:
        fn()
        PASS.append(name)
        print(f"  PASS  {name}")
    except Exception as e:
        FAIL.append((name, e))
        print(f"  FAIL  {name}: {type(e).__name__}: {e}")
        traceback.print_exc(limit=2)


def make(path, build):
    doc = Document()
    build(doc)
    doc.save(path)
    return path


# ---------------------------------------------------------------- edge inputs

def _empty(doc):
    pass


def _no_solutions(doc):
    for i in range(1, 6):
        doc.add_paragraph(f"{i}.\tA question with no solutions section at all.")
        for opt in "ABCD":
            doc.add_paragraph(f"\t({opt})\tOption {opt}")


def _tricky_numbering(doc):
    """Hard-typed sub-statements that look like question numbers."""
    doc.add_paragraph("1.\tConsider the following pairs:")
    doc.add_paragraph("1. First Past the Post System: United Kingdom")
    doc.add_paragraph("2. Proportional Representation: Israel")
    doc.add_paragraph("How many pairs are correct?")
    for opt in "ABCD":
        doc.add_paragraph(f"\t({opt})\tOption {opt}")
    doc.add_paragraph("2.\tA genuine second question?")
    for opt in "ABCD":
        doc.add_paragraph(f"\t({opt})\tOption {opt}")


print("== 1. docx_tools ==")


def _sdt_awareness():
    blocks = dt.block_elements(Document(MAINS))
    plain = Document(MAINS)
    assert len(blocks) > len(plain.paragraphs) + len(plain.tables) - 5, \
        "sdt-aware walk should surface content python-docx hides"
    tables = sum(1 for k, _ in dt.iter_block_items(Document(MAINS)) if k == 't')
    assert tables >= 6, f"expected >=6 tables via sdt walk, saw {tables}"
check("sdt-aware walk finds the hidden Mains content", _sdt_awareness)


def _page_locate():
    found, total = dt.locate_pages(PRE_EN, [1, 23, 25])
    assert total > 50, total
    assert found[1] == 0, found
    assert 700 < found[23] < 800, found
    assert found[25] > found[23], found
check("locate_pages binary search", _page_locate)


def _slice_matches_pages():
    _, all_blocks = dt.slice_blocks(PRE_EN)
    _, from23 = dt.slice_blocks(PRE_EN, start_page=23)
    assert 0 < len(from23) < len(all_blocks), (len(from23), len(all_blocks))
check("slice_blocks trims to the page range", _slice_matches_pages)


print("== 2. Prelims Answer Booklet ==")


def _prelims_en():
    out, issues = prelims_ab.process(PRE_EN, os.path.join(WORK, "pre_en"),
                                     variant="English", start_page=23)
    assert os.path.exists(out) and out.endswith(".docx")
    # placeholder advisories are informational; answer mismatches are not
    # notes about facts the raw paper cannot supply (test code, title) are
    # expected - the sample inputs are named 'Input English.docx'
    real = [i for i in issues
            if "Header test code" not in i and "placeholders" not in i
            and "Title line is a placeholder" not in i
            and "Title line left blank" not in i]
    assert not real, real
    d = Document(out)
    grids = [t for t in d.tables if len(t.rows) > 5]
    assert len(grids) == 1, f"expected one answer grid, saw {len(grids)}"
    g = grids[0]
    assert (len(g.rows), len(g.columns)) == (25, 8), (len(g.rows), len(g.columns))
    assert g.cell(0, 0).text.strip() == "1."
    assert g.cell(24, 6).text.strip() == "100."
    text = "\n".join(p.text for p in d.paragraphs)
    assert "Ans: (c)" in text
    assert "Text Solution" not in text, "the raw marker should be gone"
    assert "Source:" not in text, "Source boxes should be dropped"
check("English booklet: grid 25x8, Ans: (c), markers removed", _prelims_en)


def _prelims_hi():
    out, issues = prelims_ab.process(PRE_HI, os.path.join(WORK, "pre_hi"),
                                     variant="Hindi", start_page=23)
    d = Document(out)
    text = "\n".join(p.text for p in d.paragraphs)
    assert "उत्तर: (c)" in text, "Hindi answer label missing"
    # notes about facts the raw paper cannot supply (test code, title) are
    # expected - the sample inputs are named 'Input English.docx'
    real = [i for i in issues
            if "Header test code" not in i and "placeholders" not in i
            and "Title line is a placeholder" not in i
            and "Title line left blank" not in i]
    assert not real, real
check("Hindi booklet uses उत्तर:", _prelims_hi)


def _images_survive():
    """The input's pictures must arrive intact - including Google-Docs JPEGs
    whose JFIF-less header python-docx's sniffer refuses."""
    en = os.path.join(WORK, "pre_en", "Input English_Answer Booklet_English.docx")
    hi = os.path.join(WORK, "pre_hi", "Input Hindi_Answer Booklet_Hindi.docx")
    n_en = assert_no_dangling_images(en)
    n_hi = assert_no_dangling_images(hi)
    assert n_en >= 13, f"expected >=13 picture refs in EN, saw {n_en}"
    assert n_hi >= 13, f"expected >=13 picture refs in HI, saw {n_hi}"
    assert_numbering_resolves(en)
    assert_numbering_resolves(hi)
check("pictures + list definitions survive with valid ids (EN+HI)", _images_survive)


def _header_code_replaced():
    import zipfile
    en = os.path.join(WORK, "pre_en", "Input English_Answer Booklet_English.docx")
    z = zipfile.ZipFile(en)
    headers = " ".join(z.read(n).decode("utf8", "replace")
                       for n in z.namelist() if "header" in n)
    assert "PTS 270" not in headers, "sample's own test code leaked into the output"
    assert "[CODE]" in headers, "placeholder missing from the header"
check("sample test code replaced in the page header", _header_code_replaced)


def _answers_match_official():
    """The generated grid must equal the official booklet's grid."""
    ref = Document(os.path.join(
        S, "Prelims Test Series (Answer Booklet English)",
        "Output Test 08_A_PTS 2708_Indian Polity Part II_FLT_"
        "Prelims Wallah Test Series_UPSC_English_Aniket (1).docx"))
    ref_grid = [t for t in ref.tables if len(t.rows) > 5][0]
    mine = Document(os.path.join(WORK, "pre_en", "Input English_Answer Booklet_English.docx"))
    my_grid = [t for t in mine.tables if len(t.rows) > 5][0]
    diffs = []
    for r in range(25):
        for c in range(8):
            a = ref_grid.cell(r, c).text.strip()
            b = my_grid.cell(r, c).text.strip()
            if a != b:
                diffs.append((r, c, a, b))
    assert not diffs, f"{len(diffs)} cells differ, first: {diffs[:3]}"
check("answer grid identical to the official booklet", _answers_match_official)


def _wrong_start_page():
    try:
        prelims_ab.process(PRE_EN, os.path.join(WORK, "pre_bad"),
                           variant="English", start_page=65)
        # page 65 is the last page and holds no 'N. Text Solution:' header
        # (Q100's is on 64). This used 60, which only "worked" because the
        # start page was resolved ~1 paragraph per table row too late; pages
        # 60-64 really hold Q87-Q100 (2026-10-11, verified in Word).
        raise AssertionError("should have refused")
    except ValueError as e:
        assert "start page" in str(e).lower(), e
check("start page past the solutions -> actionable error", _wrong_start_page)


def _no_solutions_file():
    p = make(os.path.join(EDGE, "no_solutions.docx"), _no_solutions)
    try:
        prelims_ab.process(p, os.path.join(WORK, "pre_nosol"), variant="English")
        raise AssertionError("should have refused")
    except ValueError:
        pass
check("file with no solutions -> clean error", _no_solutions_file)


def _empty_file():
    p = make(os.path.join(EDGE, "empty.docx"), _empty)
    try:
        prelims_ab.process(p, os.path.join(WORK, "pre_empty"), variant="English")
        raise AssertionError("should have refused")
    except ValueError:
        pass
check("empty file -> clean error", _empty_file)


print("== 3. Prelims Question Booklet (bilingual) ==")


def _qb_pair():
    out, notes = prelims_qb.process(QB_EN, os.path.join(WORK, "qb"),
                                    second_path=QB_HI, start_page=1, end_page=22)
    assert os.path.exists(out)
    d = Document(out)
    text = "\n".join(p.text for p in d.paragraphs)
    assert "(A)" not in text, "option labels should be lowercased"
    assert "(a)" in text
    assert notes and "100 questions" in notes[0], notes
    assert_numbering_resolves(out)
    assert_no_dangling_images(out)
check("bilingual: 100 questions, options lowercased", _qb_pair)


def _qb_missing_second():
    try:
        prelims_qb.process(QB_EN, os.path.join(WORK, "qb_bad"), second_path=None)
        raise AssertionError("should have refused")
    except ValueError as e:
        assert "two files" in str(e).lower(), e
check("bilingual without the second file -> clean error", _qb_missing_second)


def _qb_tricky_numbering():
    """Hard-typed sub-statements must not be read as new questions."""
    p = make(os.path.join(EDGE, "tricky.docx"), _tricky_numbering)
    _, blocks = dt.slice_blocks(p)
    qs = prelims_qb.split_questions(blocks)
    nums = [q["no"] for q in qs]
    assert nums == [1, 2], f"expected [1, 2], got {nums}"
check("sub-statements are not mistaken for questions", _qb_tricky_numbering)


print("== 4. Mains Answer Booklet ==")


def _mains():
    out, notes = mains_ab.process(MAINS, os.path.join(WORK, "mains"), variant="Hindi")
    assert os.path.exists(out)
    d = Document(out)
    text = "\n".join(p.text for p in d.paragraphs)
    assert "प्रश्न 1." not in text, "the literal prefix should be stripped"
    assert "शब्द" not in text, "word limits must NOT be invented"
    styles = {p.style.name for p in d.paragraphs}
    assert "Questions" in styles, styles
    assert any(s in styles for s in ("Bullet-1", "Normal-DK")), styles
    assert notes and "3 question" in notes[0], notes
    assert_no_dangling_images(out)
check("Mains: prefix stripped, styles applied, nothing invented", _mains)


def _mains_cover():
    out = os.path.join(WORK, "mains", "Input SDM 2719 Hind  (1)_Ans Booklet_Hindi.docx")
    d = Document(out)
    text = "\n".join(p.text for p in d.paragraphs)
    # layout is built; the human-only facts are obvious placeholders
    for ph in ("[SERIES / YEAR]", "[TEST NAME]", "[SUBJECT]"):
        assert ph in text, f"cover placeholder {ph} missing"
    assert "(SDM 2719)" in text, "paper code should be derived from the file name"
    # question index carries the three real stems
    idx = [t for t in d.tables if len(t.columns) == 3]
    assert idx, "question index table missing"
    body = idx[0].rows[1].cells[1].text
    assert "संयुक्त परिवार" in body, "Q1 stem missing from the index"
check("Mains: cover page with derived code + placeholders", _mains_cover)


def _mains_tables_survive():
    out = os.path.join(WORK, "mains", "Input SDM 2719 Hind  (1)_Ans Booklet_Hindi.docx")
    d = Document(out)
    assert len(d.tables) >= 6, f"the sdt-hidden tables were lost: {len(d.tables)}"
check("Mains: all 6 source tables survive the sdt flatten", _mains_tables_survive)


print()
print(f"{len(PASS)} passed, {len(FAIL)} failed")
if FAIL:
    print("FAILURES:")
    for n, e in FAIL:
        print(" -", n, "->", e)
    sys.exit(1)
