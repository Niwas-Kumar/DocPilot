# -*- coding: utf-8 -*-
"""Extract the Daily Editorial Analysis (DEA) artwork from the published samples.

    python tools/make_dea_assets.py [UPSC sample folder]

Writes assets/dea/. Every file there is cut from a PW-published booklet, so
the new editions use the client's own artwork pixel for pixel:

  cover plates   the published cover with every day-to-day element taken off
                 (date, topic lists, QR, buttons) and rendered to a picture;
                 the pipelines lay those elements back on as live Word text
  banners        the promo pictures the booklets drop into space left over at
                 the end of a section
  badges         the Hindi combined edition's section-title badges
  last pages     the closing promo page of each edition

Re-run it only when the client sends new published samples - the pipelines
never read the samples themselves.
"""
import io
import json
import os
import sys

import fitz
from PIL import Image

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "assets", "dea")
DEFAULT_SAMPLES = os.environ.get("DOCPILOT_SAMPLES") or os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "UPSC sample")

SRC = {
    "dea_en": r"Daily Editorial Analysis (English)\Output 17th August 2026_Daily Editorial Summary_UPSC_English.pdf",
    "dea_hi": r"Daily Editorial Analysis (Hindi)\Output 17 Aug 2026_Daily Editorial Summary_UPSC_Hindi.pdf",
    "combo_en": r"Daily Current Affairs+Daily Editorial Analysis (English)\Output_DCA+DEA english 17th August 2026_CA Analysis+Daily Editorial Summary_UPSC_English_.pdf",
    "combo_hi": r"Daily Current Affairs+Daily Editorial Analysis (Hindi)\Output Daily CA+Daily Editorial_17 Aug_UPSC_Hindi.pdf",
}

PLATE_DPI = 300
CROP_DPI = 400


def _open(samples, key):
    return fitz.open(os.path.join(samples, SRC[key]))


def _image_rects(page):
    return [(info["xref"], fitz.Rect(info["bbox"]), info["width"], info["height"])
            for info in page.get_image_info(xrefs=True)]


def _find_image(page, w, h, tol=2.0, min_y=None):
    """xref of the picture drawn at size (w, h) points - sizes are stable
    between exports where xref numbers are not."""
    for xref, r, _pw, _ph in _image_rects(page):
        if abs(r.width - w) <= tol and abs(r.height - h) <= tol:
            if min_y is None or r.y0 >= min_y:
                return xref, r
    raise LookupError(f"no {w}x{h} picture on page {page.number + 1}")


def _save_image(doc, xref, path, flatten=True):
    """Write a picture at its native resolution. JPEGs are written as-is (no
    re-compression); anything else goes through a Pixmap, keeping its mask."""
    info = doc.extract_image(xref)
    if info.get("ext") in ("jpeg", "jpg") and not info.get("smask"):
        with open(path, "wb") as fh:
            fh.write(info["image"])
        return
    pix = fitz.Pixmap(doc, xref)
    if info.get("smask") and not pix.alpha:      # newer PyMuPDF applies it already
        pix = fitz.Pixmap(pix, fitz.Pixmap(doc, info["smask"]))
    if pix.colorspace and pix.colorspace.n > 3:
        pix = fitz.Pixmap(fitz.csRGB, pix)
    im = Image.open(io.BytesIO(pix.tobytes("png")))
    if flatten and im.mode in ("RGBA", "LA", "P"):
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im.convert("RGBA"), mask=im.convert("RGBA").split()[-1])
        im = bg
    if path.lower().endswith((".jpg", ".jpeg")):
        im.convert("RGB").save(path, quality=92, optimize=True)
    else:
        im.save(path, optimize=True)


def _text_rects(page, predicate, pad=1.2):
    """Bounding boxes of the text lines `predicate(text, span)` selects."""
    rects = []
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            spans = [s for s in line["spans"] if s["text"].strip()]
            if not spans:
                continue
            text = "".join(s["text"] for s in spans)
            if predicate(text, spans[0], fitz.Rect(line["bbox"])):
                rects.append(fitz.Rect(line["bbox"]) + (-pad, -pad, pad, pad))
    return rects


def _redact_text(page, rects):
    for r in rects:
        page.add_redact_annot(r)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                          graphics=fitz.PDF_REDACT_LINE_ART_NONE)


def _redact_line_art(page, rects):
    for r in rects:
        page.add_redact_annot(r)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                          graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED,
                          text=fitz.PDF_REDACT_TEXT_NONE)


def _remove_shapes(page, rects):
    """Vector shapes lying wholly inside `rects` go; pictures and text stay."""
    for r in rects:
        page.add_redact_annot(r)
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE,
                          graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                          text=fitz.PDF_REDACT_TEXT_NONE)


def _render(page, path, dpi=PLATE_DPI, clip=None):
    pix = page.get_pixmap(dpi=dpi, clip=clip, alpha=False)
    im = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    if path.lower().endswith((".jpg", ".jpeg")):
        im.save(path, quality=90, optimize=True, subsampling=0)
    else:
        im.save(path, optimize=True)
    return im.size


def _single_page(src_doc, page_no):
    """A one-page copy that keeps the page's transparency group. insert_pdf
    drops /Group, and the InDesign pages blend in /DeviceCMYK: without it the
    navy renders (8,11,80) instead of the published (27,22,77)."""
    out = fitz.open()
    out.insert_pdf(src_doc, from_page=page_no, to_page=page_no)
    copy_group(src_doc, page_no, out, 0)
    return out


def copy_group(src_doc, src_pno, dst_doc, dst_pno):
    kind, val = src_doc.xref_get_key(src_doc.page_xref(src_pno), "Group")
    if kind == "xref":
        val = src_doc.xref_object(int(val.split()[0]), compressed=True)
    if kind in ("xref", "dict"):
        dst_doc.xref_set_key(dst_doc.page_xref(dst_pno), "Group", val)


# ------------------------------------------------------------------ English

def english(samples, meta):
    # --- DEA cover plate: take off date, topics block, video caption, QR, buttons
    d = _open(samples, "dea_en")
    p = _single_page(d, 0)
    page = p[0]
    qr_x, qr_r = _find_image(page, 155.9, 155.9)
    # the two buttons are the same size to within 0.4pt - tell them apart by side
    buttons = sorted((r.x0, x, r) for x, r, _w, _h in _image_rects(page)
                     if abs(r.width - 93.3) < 1.5 and abs(r.height - 23.25) < 1.0)
    (_, wa_x, wa_r), (_, tg_x, tg_r) = buttons[0], buttons[-1]
    # buttons first, from the untouched page, so they sit on the same pixels
    _render(page, os.path.join(OUT, "en_btn_whatsapp.png"), CROP_DPI, wa_r)
    _render(page, os.path.join(OUT, "en_btn_telegram.png"), CROP_DPI, tg_r)
    pill_x, pill_r = _find_image(page, 297.5, 47.7)
    rects = _text_rects(page, lambda t, s, r: r.y0 > 200 and r.y1 < 700)
    _redact_text(page, rects)
    # the orange pill on top is a picture with the date baked into its pixels;
    # the plate's own pill underneath is blank
    for x in (qr_x, wa_x, tg_x, pill_x):
        page.delete_image(x)
    size = _render(page, os.path.join(OUT, "en_dea_cover_plate.jpg"))
    meta["en_dea_cover_plate"] = {"px": size, "buttons": {"whatsapp": list(wa_r), "telegram": list(tg_r)},
                                  "qr": list(qr_r), "pill": list(pill_r)}

    # --- end-of-content banner and closing page picture
    x, r = _find_image(d[6], 552.2, 310.4)
    _save_image(d, x, os.path.join(OUT, "en_pyq_banner.jpg"))
    meta["en_pyq_banner"] = {"rect": list(r)}
    x, r = _find_image(d[7], 552.5, 519.6)
    _save_image(d, x, os.path.join(OUT, "en_srijan_full.jpg"))
    meta["en_srijan_full"] = {"rect": list(r)}
    # the "PW Web/App ... / subscribe to our Telegram Channel" line. NOT the
    # file in assets/template/promo: that one is fully transparent (alpha 0
    # everywhere), so it prints nothing - kept apart so Daily CA is untouched
    # it is a stencil mask painted in colour, so it is rendered, not extracted
    _x, r = _find_image(d[7], 382.1, 52.6)
    r = fitz.Rect(r.x0, r.y0, r.x1, 728.4)       # stop above the page-number circle
    _render(d[7], os.path.join(OUT, "en_webapp_line.png"), CROP_DPI, r)
    meta["en_webapp_line"] = {"rect": list(r)}

    # --- combined cover: plate keeps both titles and INDEX, loses the lists,
    # the two box outlines, the date text and the button groups
    c = _open(samples, "combo_en")
    p = _single_page(c, 0)
    page = p[0]
    groups = {"whatsapp": fitz.Rect(56.0, 707.0, 186.0, 747.0),
              "telegram": fitz.Rect(427.5, 705.5, 560.5, 746.5)}
    for name, r in groups.items():
        _render(page, os.path.join(OUT, f"en_combo_btn_{name}.png"), CROP_DPI, r)
    title_x, title_r = _find_image(page, 503.1, 39.7)
    _save_image(p, title_x, os.path.join(OUT, "en_combo_dea_title.png"))
    pill_x, pill_r = _find_image(page, 297.1, 47.8)
    rects = _text_rects(page, lambda t, s, r: r.y0 > 205)
    _redact_text(page, rects)
    # the index boxes are a blue shape under a white one - both go, whole
    _remove_shapes(page, [fitz.Rect(50.0, 319.0, 562.0, 432.0), fitz.Rect(50.0, 577.0, 562.0, 666.0)])
    # the button outlines are strokes: take them by their edges
    edges = []
    for b in groups.values():
        edges += [fitz.Rect(b.x0, b.y0, b.x1, b.y0 + 3), fitz.Rect(b.x0, b.y1 - 3, b.x1, b.y1),
                  fitz.Rect(b.x0, b.y0, b.x0 + 3, b.y1), fitz.Rect(b.x1 - 3, b.y0, b.x1, b.y1)]
    _redact_line_art(page, edges)
    for xref, r, _w, _h in _image_rects(page):
        if r.y0 > 700 and r.width < 120:
            page.delete_image(xref)
    page.delete_image(pill_x)
    size = _render(page, os.path.join(OUT, "en_combo_cover_plate.jpg"))
    meta["en_combo_cover_plate"] = {"px": size, "dea_title": list(title_r), "pill": list(pill_r),
                                    "buttons": {k: list(v) for k, v in groups.items()}}


# ------------------------------------------------------------------ Hindi

FOUNDATION_RECT = fitz.Rect(36.0, 334.5, 576.0, 699.5)    # DEA p6: artwork 339.7-696.4
SRIJAN_RECT = fitz.Rect(144.4, 505.5, 479.2, 718.0)       # combined p16: artwork 510.2-714.3


def _banner_crop(doc, page_no, rect, path):
    """Render one promo banner on its own: the footer buttons and the page
    watermark are taken off a copy of the page first, so the crop carries
    the banner and nothing of the page around it."""
    tmp = _single_page(doc, page_no)
    page = tmp[0]
    for xref, r, _w, _h in _image_rects(page):
        if r.y0 > 715 or (abs(r.width - 447.9) < 2 and abs(r.height - 128.3) < 2):
            page.delete_image(xref)                       # buttons' icons, watermark
    footer = fitz.Rect(0, 718.0, 612, 792)
    _redact_text(page, [footer])
    _remove_shapes(page, [footer])
    _render(page, path, CROP_DPI, rect)

def hindi(samples, meta):
    d = _open(samples, "dea_hi")
    p = _single_page(d, 0)
    page = p[0]
    qr_x, qr_r = _find_image(page, 143.2, 147.4)
    # the date (inside its outlined box) and the index lines under विषय सूची
    # whole bands, not line boxes: the broken Devanagari text layer leaves
    # glyphs sitting between two line boxes
    _redact_text(page, [fitz.Rect(237.0, 319.0, 375.0, 352.0),      # date, inside its outline
                        fitz.Rect(24.0, 664.0, 588.0, 752.0)])      # index lines under विषय सूची
    page.delete_image(qr_x)
    size = _render(page, os.path.join(OUT, "hi_dea_cover_plate.jpg"))
    meta["hi_dea_cover_plate"] = {"px": size, "qr": list(qr_r), "date_box": [235.3, 317.4, 376.7, 353.3]}

    # promo banners are a paper texture with vector artwork on top, so they
    # are rendered from their region (page furniture taken off first)
    r = FOUNDATION_RECT
    _banner_crop(d, 5, r, os.path.join(OUT, "hi_foundation_banner.png"))
    meta["hi_foundation_banner"] = {"rect": list(r)}
    # FLATTENED to a picture: the master page's button captions and page
    # number lie under the full-bleed artwork - invisible, but in the text
    # layer (a stray "7" in every booklet), and redacting them also cut the
    # visible address line that overlaps them. A 300dpi render carries
    # exactly what is seen and nothing else.
    last = _single_page(d, 6)
    _render(last[0], os.path.join(OUT, "hi_last_page.jpg"))
    stale = os.path.join(OUT, "hi_last_page.pdf")
    if os.path.exists(stale):
        os.remove(stale)

    c = _open(samples, "combo_hi")
    p = _single_page(c, 0)
    page = p[0]
    _redact_text(page, [fitz.Rect(236.5, 296.5, 375.5, 315.5),      # date, inside its outline
                        fitz.Rect(24.0, 322.0, 588.0, 470.0),       # daily news index
                        fitz.Rect(24.0, 572.0, 588.0, 695.0)])      # editorial index
    size = _render(page, os.path.join(OUT, "hi_combo_cover_plate.jpg"))
    meta["hi_combo_cover_plate"] = {"px": size, "date_box": [235.3, 295.5, 376.7, 316.5]}

    badges = {"dna": (1, fitz.Rect(184.8, 24.6, 427.1, 67.4)),
              "dea": (12, fitz.Rect(175.9, 27.4, 436.1, 67.8))}
    for name, (pno, r) in badges.items():
        _render(c[pno], os.path.join(OUT, f"hi_badge_{name}.png"), CROP_DPI, r)
        meta[f"hi_badge_{name}"] = {"rect": list(r)}
    r = SRIJAN_RECT
    _banner_crop(c, 15, r, os.path.join(OUT, "hi_srijan_banner.png"))
    meta["hi_srijan_banner"] = {"rect": list(r)}


def main():
    samples = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SAMPLES
    os.makedirs(OUT, exist_ok=True)
    meta = {}
    english(samples, meta)
    hindi(samples, meta)
    with open(os.path.join(OUT, "geometry.json"), "w", encoding="utf-8") as fh:
        json.dump(meta, fh, indent=1, default=lambda o: [round(v, 2) for v in o])
    for name in sorted(os.listdir(OUT)):
        print(f"{os.path.getsize(os.path.join(OUT, name)):>10,}  {name}")


if __name__ == "__main__":
    main()
