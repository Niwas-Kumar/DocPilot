# -*- coding: utf-8 -*-
"""PDF post-processing: page-number circles, footer links, promo last page."""
import os
import fitz

ASSETS = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "assets", "template")

BLUE = (0x00 / 255, 0x70 / 255, 0xC0 / 255)
WHITE = (1, 1, 1)

CIRCLE_CENTER = (306, 741.8)
CIRCLE_R = 11.5
NUM_BASELINE_Y = 745.3

LINK_WHATSAPP = "https://whatsapp.com/channel/0029Van2ybPIt5rqKRSCzK14"
LINK_TELEGRAM = "https://t.me/pw_upsc"
RECT_WHATSAPP = fitz.Rect(57.0, 726.9, 150.2, 750.1)
RECT_TELEGRAM = fitz.Rect(466.6, 726.7, 560.2, 750.0)

SEGOE_BOLD = r"C:\Windows\Fonts\segoeuib.ttf"

# promo page geometry measured from the official sample (page 612x792)
PROMO_IMAGES = [
    ("promo/promo_srijan.jpeg", fitz.Rect(57.5, 76.3, 560.1, 323.1)),
    ("promo/promo_pyq32.jpeg", fitz.Rect(29.2, 328.7, 582.8, 666.2)),
    ("promo/customer_service_icon.jpeg", fitz.Rect(89.3, 671.6, 144.0, 722.4)),
    ("promo/pw_webapp_line.png", fitz.Rect(145.8, 674.0, 527.8, 726.6)),
]
PROMO_LINKS = [
    (fitz.Rect(299, 674, 439, 692), "https://smart.link/7wwosivoicgd4"),
    (fitz.Rect(426, 692, 518, 710), LINK_TELEGRAM),
]
LOGO_RECT = fitz.Rect(223.5, 32.4, 393.5, 74.0)

# page border bands (thick outer + thin inner, blue)
BORDER_BANDS = [
    fitz.Rect(24.0, 24.0, 588.1, 27.0),     # top outer
    fitz.Rect(24.0, 28.4, 588.1, 29.9),     # top inner
    fitz.Rect(24.0, 765.1, 588.1, 768.1),   # bottom outer
    fitz.Rect(24.0, 762.2, 588.1, 763.7),   # bottom inner
    fitz.Rect(24.0, 24.0, 27.0, 768.1),     # left outer
    fitz.Rect(28.4, 28.4, 29.9, 763.7),     # left inner
    fitz.Rect(585.1, 24.0, 588.1, 768.1),   # right outer
    fitz.Rect(582.2, 28.4, 583.7, 763.7),   # right inner
]


_segoe_font = None


def _get_segoe():
    global _segoe_font
    if _segoe_font is None:
        try:
            _segoe_font = fitz.Font(fontfile=SEGOE_BOLD)
        except Exception:
            _segoe_font = fitz.Font("hebo")  # Helvetica-Bold fallback
    return _segoe_font


def _stamp_page_number(page, number, fontbuf=None):
    page.draw_circle(fitz.Point(*CIRCLE_CENTER), CIRCLE_R, color=None, fill=BLUE)
    label = str(number)
    font = _get_segoe()
    tw = font.text_length(label, fontsize=10)
    tw_pt = fitz.Point(CIRCLE_CENTER[0] - tw / 2, NUM_BASELINE_Y)
    if font.name != "Helvetica-Bold":
        page.insert_text(tw_pt, label, fontfile=SEGOE_BOLD, fontname="SegoeUIBold",
                         fontsize=10, color=WHITE)
    else:
        page.insert_text(tw_pt, label, fontname="hebo", fontsize=10, color=WHITE)


def _ensure_footer_links(page):
    uris = {l.get("uri") for l in page.get_links()}
    if LINK_WHATSAPP not in uris:
        page.insert_link({"kind": fitz.LINK_URI, "from": RECT_WHATSAPP, "uri": LINK_WHATSAPP})
    if LINK_TELEGRAM not in uris:
        page.insert_link({"kind": fitz.LINK_URI, "from": RECT_TELEGRAM, "uri": LINK_TELEGRAM})


def _build_promo_page(doc, number):
    page = doc.new_page(width=612, height=792)
    page.insert_image(fitz.Rect(0, 0, 612, 792),
                      filename=os.path.join(ASSETS, "watermark_page_background.jpeg"))
    for band in BORDER_BANDS:
        page.draw_rect(band, color=None, fill=BLUE)
    page.insert_image(LOGO_RECT, filename=os.path.join(ASSETS, "header_logo_pw_onlyias.png"))
    for rel, rect in PROMO_IMAGES:
        page.insert_image(rect, filename=os.path.join(ASSETS, rel))
    page.insert_image(RECT_WHATSAPP, filename=os.path.join(ASSETS, "btn_join_whatsapp.png"))
    page.insert_image(RECT_TELEGRAM, filename=os.path.join(ASSETS, "btn_join_telegram.png"))
    _ensure_footer_links(page)
    for rect, uri in PROMO_LINKS:
        page.insert_link({"kind": fitz.LINK_URI, "from": rect, "uri": uri})
    _stamp_page_number(page, number, None)
    return page


def finalize_pdf(in_pdf: str, out_pdf: str, first_number: int = 2) -> str:
    """Stamp page numbers (sequence starts at `first_number`), ensure footer
    links, append the rebuilt promo page with the continuing number."""
    doc = fitz.open(in_pdf)
    n = first_number
    for page in doc:
        _stamp_page_number(page, n, None)
        _ensure_footer_links(page)
        n += 1
    _build_promo_page(doc, n)
    doc.save(out_pdf, garbage=3, deflate=True)
    doc.close()
    return out_pdf
