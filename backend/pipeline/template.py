# -*- coding: utf-8 -*-
"""Apply the OnlyIAS page template (Letter, borders, watermark header, footer buttons)."""
import os
from docx.shared import Inches, Pt
from docx.enum.section import WD_HEADER_FOOTER
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from .ooxml import add_anchored_image

ASSETS = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "assets", "template")

PAGE_W_PT = 612.0
PAGE_H_PT = 792.0

LINK_WHATSAPP = "https://whatsapp.com/channel/0029Van2ybPIt5rqKRSCzK14"
LINK_TELEGRAM = "https://t.me/pw_upsc"

BORDER_XML = (
    '<w:pgBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'w:offsetFrom="page">'
    '<w:top w:val="thinThickSmallGap" w:sz="24" w:space="24" w:color="0070C0"/>'
    '<w:left w:val="thinThickSmallGap" w:sz="24" w:space="24" w:color="0070C0"/>'
    '<w:bottom w:val="thickThinSmallGap" w:sz="24" w:space="24" w:color="0070C0"/>'
    '<w:right w:val="thickThinSmallGap" w:sz="24" w:space="24" w:color="0070C0"/>'
    '</w:pgBorders>'
)


def _setup_section_geometry(section):
    section.page_width = Pt(PAGE_W_PT)
    section.page_height = Pt(PAGE_H_PT)
    section.left_margin = Pt(57)
    section.right_margin = Pt(52)
    section.top_margin = Pt(76)
    section.bottom_margin = Pt(78)
    section.header_distance = Pt(30)
    section.footer_distance = Pt(30)
    section.different_first_page_header_footer = False

    # page border (thick-thin double rule, blue, inset 24pt from page edge)
    sectPr = section._sectPr
    old = sectPr.find(qn('w:pgBorders'))
    if old is not None:
        sectPr.remove(old)
    pgMar = sectPr.find(qn('w:pgMar'))
    border = parse_xml(BORDER_XML)
    pgMar.addnext(border)


def apply_page_template(doc):
    """Mutates doc: Letter page, margins, page border, watermark+logo header, footer buttons."""
    # every section gets the geometry + border; later sections inherit the
    # first section's header/footer via Word's link-to-previous default
    for i, section in enumerate(doc.sections):
        _setup_section_geometry(section)
        if i > 0:
            section.header.is_linked_to_previous = True
            section.footer.is_linked_to_previous = True

    section = doc.sections[0]
    # header: full-bleed watermark background (behind text) + centered logo
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.text = ""
    add_anchored_image(hp, header.part,
                       os.path.join(ASSETS, "watermark_page_background.jpeg"),
                       0, 0, PAGE_W_PT, PAGE_H_PT, behind=True, name="bg_watermark")
    add_anchored_image(hp, header.part,
                       os.path.join(ASSETS, "header_logo_pw_onlyias.png"),
                       223.5, 32.4, 170.0, 41.6, name="pw_onlyias_logo")

    # footer: join buttons, hyperlinked, hugging page edges
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.text = ""
    add_anchored_image(fp, footer.part,
                       os.path.join(ASSETS, "btn_join_whatsapp.png"),
                       57.0, 726.9, 93.2, 23.2, hyperlink=LINK_WHATSAPP, name="btn_whatsapp")
    add_anchored_image(fp, footer.part,
                       os.path.join(ASSETS, "btn_join_telegram.png"),
                       466.6, 726.7, 93.6, 23.3, hyperlink=LINK_TELEGRAM, name="btn_telegram")
