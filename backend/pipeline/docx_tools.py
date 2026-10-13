# -*- coding: utf-8 -*-
"""Shared helpers for the Test Series pipelines.

Nothing here is used by the Daily CA (English) pipeline - that one is frozen.
"""
import copy
import os
import re
import threading

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TEMPLATES = os.path.join(ROOT, "assets", "templates")

_word_lock = threading.Lock()


def make_pdf_companion(docx_path, notes=None):
    """Render `docx_path` to a PDF of the same stem beside it.

    The Test Series categories deliver a Word file; the PDF companion lets the
    UI offer both downloads. A conversion failure is reported as a note, never
    as a job failure - the DOCX is the primary deliverable.
    """
    from .convert import docx_to_pdf
    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
    try:
        docx_to_pdf(docx_path, pdf_path)
        return pdf_path
    except Exception as e:                       # Word busy / missing / crash
        if notes is not None:
            notes.append(f"PDF version could not be produced ({type(e).__name__}); "
                         "the Word file is complete - export it from Word if needed.")
        return None


def restore_original_jpegs(docx_path, pdf_path):
    """Put the raw file's own JPEG bytes back into Word's PDF export.

    Word re-encodes every JPEG when it exports a PDF - whatever the
    'do not compress' settings say - so each picture loses a generation of
    quality on top of an often small original. The pixel size is unchanged,
    so the re-encoded image can be matched to its source in the .docx and
    the original stream dropped back in (DCTDecode carries a JPEG file as
    is). Only plain RGB / greyscale JPEGs whose pixels visibly agree are
    swapped; anything else (cropped, CMYK, rotated) keeps Word's copy.
    Returns the number of pictures restored; never raises.
    """
    import io
    import zipfile
    try:
        import fitz
        from PIL import Image, ImageChops, ImageOps
    except ImportError:
        return 0

    def thumb(img):
        return ImageOps.grayscale(img.convert("RGB")).resize((32, 32))

    try:
        sources = []
        with zipfile.ZipFile(docx_path) as z:
            for name in z.namelist():
                if not name.startswith("word/media/"):
                    continue
                data = z.read(name)
                if data[:3] != b"\xff\xd8\xff":
                    continue
                try:
                    img = Image.open(io.BytesIO(data))
                    if img.mode not in ("RGB", "L"):
                        continue
                    if img.getexif().get(0x0112, 1) != 1:   # EXIF-rotated
                        continue
                    sources.append((img.size, img.mode, data, thumb(img)))
                except Exception:
                    continue
        if not sources:
            return 0

        doc = fitz.open(pdf_path)
        restored = 0
        for xref in range(1, doc.xref_length()):
            try:
                if doc.xref_get_key(xref, "Subtype")[1] != "/Image":
                    continue
                if "DCTDecode" not in doc.xref_get_key(xref, "Filter")[1]:
                    continue
                size = (int(doc.xref_get_key(xref, "Width")[1]),
                        int(doc.xref_get_key(xref, "Height")[1]))
                cands = [s for s in sources if s[0] == size]
                if not cands:
                    continue
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha not in (1, 3):
                    continue
                mode = "L" if pix.n - pix.alpha == 1 else "RGB"
                if pix.alpha:
                    pix = fitz.Pixmap(pix, 0)
                shown = thumb(Image.frombytes(mode, (pix.width, pix.height), pix.samples))
                best = min(cands, key=lambda s: sum(
                    ImageChops.difference(shown, s[3]).getdata()))
                diff = sum(ImageChops.difference(shown, best[3]).getdata()) / 1024
                if diff > 12:                        # not the same picture
                    continue
                doc.update_stream(xref, best[2], compress=False)
                doc.xref_set_key(xref, "Filter", "/DCTDecode")
                doc.xref_set_key(xref, "DecodeParms", "null")
                doc.xref_set_key(xref, "ColorSpace",
                                 "/DeviceGray" if best[1] == "L" else "/DeviceRGB")
                doc.xref_set_key(xref, "BitsPerComponent", "8")
                doc.xref_set_key(xref, "Decode", "null")
                restored += 1
            except Exception:
                continue
        if restored:
            # never save a PyMuPDF result over its own input
            tmp = pdf_path + ".tmp"
            doc.save(tmp, garbage=3, deflate=True)
            doc.close()
            os.replace(tmp, pdf_path)
        else:
            doc.close()
        return restored
    except Exception:
        return 0


# --------------------------------------------------------------- block access

def iter_block_items(parent):
    """Yield ('p', Paragraph) / ('t', Table) in true document order.

    Unlike python-docx's own iteration this descends into w:sdt content
    controls. Google Docs exports wrap large parts of the document in them and
    python-docx skips those silently - in the Mains sample that hides 5 of 6
    tables and a third of the paragraphs.
    """
    if hasattr(parent, "element"):
        root = parent.element.body if hasattr(parent.element, "body") else parent.element
        doc = parent
    else:
        root, doc = parent, None

    def walk(node):
        for child in node.iterchildren():
            if child.tag == qn('w:p'):
                yield 'p', Paragraph(child, doc)
            elif child.tag == qn('w:tbl'):
                yield 't', Table(child, doc)
            elif child.tag == qn('w:sdt'):
                content = child.find(qn('w:sdtContent'))
                if content is not None:
                    yield from walk(content)

    yield from walk(root)


def block_elements(doc):
    """Flat list of the underlying XML elements, sdt-unwrapped, in order."""
    return [obj._p if kind == 'p' else obj._tbl
            for kind, obj in iter_block_items(doc)]


def para_text(p):
    """Paragraph text including runs nested inside hyperlinks / sdt."""
    return "".join(t.text or "" for t in p._p.iter(qn('w:t')))


def clean(s):
    s = (s or "").replace("\xa0", " ").replace("​", "")
    return re.sub(r"\s+", " ", s).strip()


# --------------------------------------------------------------- page mapping

def locate_pages(docx_path, targets, cancelled=lambda: False):
    """For each page number in `targets`, find the first Word paragraph on it.

    Page numbers increase monotonically down the document, so this binary
    searches instead of asking Word about every paragraph - on a 1,200-paragraph
    booklet that is ~11 COM round-trips per target instead of 1,200.

    Returns (mapping {page: paragraph_index_0_based}, total_pages).
    """
    import pythoncom
    import win32com.client

    wdActiveEndPageNumber = 3
    with _word_lock:
        pythoncom.CoInitialize()
        word = doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True,
                                      AddToRecentFiles=False)
            n = doc.Paragraphs.Count
            total = doc.ComputeStatistics(2)
            cache = {}

            def page_of(i):                      # i is 1-based
                if i not in cache:
                    cache[i] = doc.Paragraphs(i).Range.Information(wdActiveEndPageNumber)
                return cache[i]

            found = {}
            for target in targets:
                if cancelled():
                    raise InterruptedError("stopped")
                lo, hi, ans = 1, n, n + 1
                while lo <= hi:
                    mid = (lo + hi) // 2
                    if page_of(mid) >= target:
                        ans, hi = mid, mid - 1
                    else:
                        lo = mid + 1
                found[target] = ans - 1          # to 0-based
            return found, total
        finally:
            try:
                if doc is not None:
                    doc.Close(0)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()


def pages_of_paragraphs(docx_path, indices, cancelled=lambda: False, pdf_path=None):
    """Page number for each 0-based paragraph index in `indices`.

    With `pdf_path`, the same Word session also exports the PDF, so a caller
    that needs both the page numbers and the geometry pays for one Word start
    and one layout instead of two.
    """
    import pythoncom
    import win32com.client

    wdActiveEndPageNumber = 3
    with _word_lock:
        pythoncom.CoInitialize()
        word = doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True,
                                      AddToRecentFiles=False)
            n = doc.Paragraphs.Count
            out = {}
            for i in indices:
                if cancelled():
                    raise InterruptedError("stopped")
                if 0 <= i < n:
                    out[i] = doc.Paragraphs(i + 1).Range.Information(wdActiveEndPageNumber)
            total = doc.ComputeStatistics(2)
            if pdf_path:
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
                doc.ExportAsFixedFormat(OutputFileName=os.path.abspath(pdf_path),
                                        ExportFormat=17, OpenAfterExport=False,
                                        OptimizeFor=0, CreateBookmarks=0,
                                        DocStructureTags=False,
                                        BitmapMissingFonts=True)
            return out, total
        finally:
            try:
                if doc is not None:
                    doc.Close(0)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            pythoncom.CoUninitialize()


def word_paragraph_count(el):
    """How many entries a table adds to Word's Document.Paragraphs: every w:p
    in it, plus one end-of-row mark per row (text-box paragraphs live in
    another story and are not counted)."""
    count = 0
    for node in el.iter(qn('w:p'), qn('w:tr')):
        if any(a.tag == qn('w:txbxContent') for a in node.iterancestors()):
            continue
        count += 1
    return count


def _para_index_to_block_index(blocks, para_index):
    """Word counts every paragraph, including those inside tables - and each
    table row's end mark. Missing the row marks put a start page ~1 paragraph
    per table row too late: Test 13's answer key lost its first 21 lines."""
    seen = 0
    for bi, el in enumerate(blocks):
        if el.tag == qn('w:p'):
            if seen >= para_index:
                return bi
            seen += 1
        else:
            inner = word_paragraph_count(el)
            if seen + inner > para_index:
                return bi
            seen += inner
    return len(blocks)


def slice_blocks(docx_path, start_page=None, end_page=None, cancelled=lambda: False):
    """Open the doc and return (Document, [block elements] within the range)."""
    doc = Document(docx_path)
    blocks = block_elements(doc)
    if not start_page and not end_page:
        return doc, blocks

    targets = [p for p in (start_page, (end_page + 1) if end_page else None) if p]
    found, total = locate_pages(docx_path, targets, cancelled)

    lo = _para_index_to_block_index(blocks, found[start_page]) if start_page else 0
    if end_page:
        hi = _para_index_to_block_index(blocks, found[end_page + 1])
    else:
        hi = len(blocks)
    return doc, blocks[lo:hi]


# --------------------------------------------------------------- template use

def open_template(name):
    path = os.path.join(TEMPLATES, name)
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Template {name} is missing. Run: python tools\\make_templates.py")
    return Document(path)


def template_body(doc):
    return doc.element.body


def clear_body_keep_sections(doc):
    """Empty the template body but keep its section breaks as anchors.

    Returns the list of section-break paragraphs, in order, so a pipeline can
    place content before/after each one.
    """
    body = doc.element.body
    anchors = []
    for child in list(body):
        if child.tag == qn('w:p'):
            ppr = child.find(qn('w:pPr'))
            if ppr is not None and ppr.find(qn('w:sectPr')) is not None:
                anchors.append(child)
            else:
                body.remove(child)
        elif child.tag in (qn('w:tbl'), qn('w:sdt')):
            body.remove(child)
    return anchors


def append_block(doc, element, before=None):
    """Deep-copy an element into the document body (or before an anchor)."""
    new = copy.deepcopy(element)
    if before is not None:
        before.addprevious(new)
    else:
        sect = doc.element.body.find(qn('w:sectPr'))
        if sect is not None:
            sect.addprevious(new)
        else:
            doc.element.body.append(new)
    return new


R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

_CTYPES = {
    "jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
    "gif": "image/gif", "bmp": "image/bmp", "tif": "image/tiff",
    "tiff": "image/tiff", "emf": "image/x-emf", "wmf": "image/x-wmf",
}


def _adopt_image_raw(doc, blob, src_partname):
    """Add an image to `doc` as a raw part, bypassing python-docx's sniffer.

    Google-Docs exports contain JPEGs whose header is a bare FF D8 FF DB (no
    JFIF/EXIF APP0 marker). Word renders them fine, but python-docx's
    Image.from_blob refuses them, so get_or_add_image raises and the picture
    would be dropped. Deduped by content hash per document.
    """
    import hashlib

    from docx.opc.packuri import PackURI
    from docx.opc.part import Part

    cache = getattr(doc, "_docpilot_img_cache", None)
    if cache is None:
        cache = {}
        doc._docpilot_img_cache = cache
    key = hashlib.sha1(blob).hexdigest()
    if key in cache:
        return cache[key]

    ext = str(src_partname).rsplit(".", 1)[-1].lower() or "png"
    ctype = _CTYPES.get(ext, "image/png")
    package = doc.part.package
    used = {str(p.partname) for p in package.iter_parts()}
    n = 1
    while f"/word/media/docpilot{n}.{ext}" in used:
        n += 1
    part = Part(PackURI(f"/word/media/docpilot{n}.{ext}"), ctype,
                blob=blob, package=package)
    cache[key] = part
    return part


def _copy_numbering_def(doc, src_doc, num_id):
    """Bring one list definition (w:num + its w:abstractNum) across.

    Copied paragraphs keep their w:numId, but that id belongs to the SOURCE
    document's numbering.xml. In the destination it is undefined - and Word
    silently substitutes decimal numbering, so diamond bullets print as
    "1./2." inside answers. Returns the new numId, deduped per document pair.
    """
    import copy as _copy

    from docx.opc.constants import RELATIONSHIP_TYPE as RT

    cache = getattr(doc, "_docpilot_num_map", None)
    if cache is None:
        cache = {}
        doc._docpilot_num_map = cache
    key = (id(src_doc), num_id)
    if key in cache:
        return cache[key]

    try:
        src_numbering = src_doc.part.part_related_by(RT.NUMBERING).element
        dst_numbering = doc.part.part_related_by(RT.NUMBERING).element
    except KeyError:
        cache[key] = None
        return None

    src_num = next((n for n in src_numbering.findall(qn('w:num'))
                    if n.get(qn('w:numId')) == str(num_id)), None)
    if src_num is None:
        cache[key] = None
        return None
    abs_ref = src_num.find(qn('w:abstractNumId'))
    src_abs = None
    if abs_ref is not None:
        src_abs = next((a for a in src_numbering.findall(qn('w:abstractNum'))
                        if a.get(qn('w:abstractNumId')) == abs_ref.get(qn('w:val'))),
                       None)

    def _max(tag, attr):
        vals = [int(e.get(qn(attr))) for e in dst_numbering.findall(qn(tag))
                if (e.get(qn(attr)) or "").lstrip("-").isdigit()]
        return max(vals) if vals else 0

    new_abs_id = _max('w:abstractNum', 'w:abstractNumId') + 1
    new_num_id = _max('w:num', 'w:numId') + 1

    if src_abs is not None:
        abs_copy = _copy.deepcopy(src_abs)
        abs_copy.set(qn('w:abstractNumId'), str(new_abs_id))
        # abstractNum elements must precede every w:num
        first_num = dst_numbering.find(qn('w:num'))
        if first_num is not None:
            first_num.addprevious(abs_copy)
        else:
            dst_numbering.append(abs_copy)

    num_copy = _copy.deepcopy(src_num)
    num_copy.set(qn('w:numId'), str(new_num_id))
    ref = num_copy.find(qn('w:abstractNumId'))
    if ref is not None and src_abs is not None:
        ref.set(qn('w:val'), str(new_abs_id))
    dst_numbering.append(num_copy)

    cache[key] = new_num_id
    return new_num_id


def copy_block_with_media(doc, element, src_doc, before=None):
    """Copy a block into `doc` and bring its pictures AND list definitions.

    A plain deep-copy keeps the drawing XML but its `r:embed` still names a
    relationship id from the SOURCE document. In the destination that id either
    does not exist (Word shows a broken-image placeholder) or points at a
    different picture. This pulls each referenced image part across and rewrites
    the ids. External hyperlinks are re-registered the same way, and w:numId
    references are remapped onto copied numbering definitions so bullets do not
    degrade into decimal numbering.
    """
    import io

    new = append_block(doc, element, before=before)
    src_part, dst_part = src_doc.part, doc.part
    remapped = {}

    for num_ref in new.iter(qn('w:numId')):
        old = num_ref.get(qn('w:val'))
        if old is None or not str(old).isdigit():
            continue
        new_id = _copy_numbering_def(doc, src_doc, int(old))
        if new_id is not None:
            num_ref.set(qn('w:val'), str(new_id))

    for node in new.iter():
        for attr in (f"{{{R_NS}}}embed", f"{{{R_NS}}}link", f"{{{R_NS}}}id"):
            rid = node.get(attr)
            if not rid:
                continue
            if rid in remapped:
                node.set(attr, remapped[rid])
                continue
            try:
                rel = src_part.rels[rid]
            except KeyError:
                continue
            try:
                if rel.is_external:
                    from docx.opc.constants import RELATIONSHIP_TYPE as RT
                    new_rid = dst_part.relate_to(rel.target_ref, RT.HYPERLINK,
                                                 is_external=True)
                else:
                    blob = rel.target_part.blob
                    try:
                        new_rid, _img = dst_part.get_or_add_image(io.BytesIO(blob))
                    except Exception:
                        # sniffer rejected the blob (e.g. JFIF-less JPEG) -
                        # adopt it as a raw part instead of dropping the picture
                        from docx.opc.constants import RELATIONSHIP_TYPE as RT
                        part = _adopt_image_raw(doc, blob, rel.target_part.partname)
                        new_rid = dst_part.relate_to(part, RT.IMAGE)
            except Exception:
                continue
            remapped[rid] = new_rid
            node.set(attr, new_rid)
    return new


def new_paragraph(doc, text="", style=None, before=None):
    p = doc.element.body.makeelement(qn('w:p'), {})
    if style:
        ppr = p.makeelement(qn('w:pPr'), {})
        ps = p.makeelement(qn('w:pStyle'), {})
        ps.set(qn('w:val'), style)
        ppr.append(ps)
        p.append(ppr)
    if text:
        r = p.makeelement(qn('w:r'), {})
        t = p.makeelement(qn('w:t'), {})
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        r.append(t)
        p.append(r)
    if before is not None:
        before.addprevious(p)
    else:
        sect = doc.element.body.find(qn('w:sectPr'))
        if sect is not None:
            sect.addprevious(p)
        else:
            doc.element.body.append(p)
    return Paragraph(p, doc)


def strip_sdt(doc):
    """Flatten every w:sdt wrapper so downstream code sees plain content."""
    body = doc.element.body
    changed = True
    while changed:
        changed = False
        for sdt in list(body.iter(qn('w:sdt'))):
            content = sdt.find(qn('w:sdtContent'))
            parent = sdt.getparent()
            if parent is None:
                continue
            if content is not None:
                for child in list(content):
                    sdt.addprevious(child)
            parent.remove(sdt)
            changed = True
