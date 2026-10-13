# -*- coding: utf-8 -*-
"""A small absolute-layout kit for building a one-page cover in Word.

The Daily Editorial covers have no client Word template - only the published
PDFs - so each cover is assembled here from a background plate (the published
artwork with the day-to-day elements taken off, see tools/make_dea_assets.py)
plus live Word objects laid exactly where the published page has them:

    picture   page-anchored, optionally hyperlinked (reuses ooxml)
    text box  page-anchored, fixed size, no autofit, optional fill / outline

Every object is anchored to the PAGE and wrapped "none", and all of them hang
off ONE host paragraph with a hairline line height - so the paragraph flow
plays no part and the cover can never spill onto a second page.

The document is built through Word (like every other cover in DocPilot), so
the same file is both the PDF's first page and the editable Word companion's.
"""
import os

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Pt
from xml.sax.saxutils import escape

from .ooxml import add_anchored_image, next_docpr_id, pt_to_emu, pt_to_twips

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_NS = f'xmlns:w="{NS_W}"'
NSDECL = (
    f'xmlns:w="{NS_W}" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

PAGE_W, PAGE_H = 612.0, 792.0


# ------------------------------------------------------------------ runs & paragraphs

def run(text, font, size_pt, color="000000", bold=False, italic=False, sup=False,
        cs_font=None, spacing_pt=0.0):
    """One formatted run. Devanagari needs the complex-script slot and the
    cs flags too, or Word renders it in the wrong face. `spacing_pt` is
    character spacing (negative = condensed), see fit_spacing()."""
    cs = cs_font or font
    half = int(round(size_pt * 2))
    b = "<w:b/><w:bCs/>" if bold else '<w:b w:val="0"/><w:bCs w:val="0"/>'
    i = "<w:i/><w:iCs/>" if italic else ""
    va = '<w:vertAlign w:val="superscript"/>' if sup else ""
    sp = f'<w:spacing w:val="{int(round(spacing_pt * 20))}"/>' if spacing_pt else ""
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{cs}" w:eastAsia="{font}"/>'
            f'{b}{i}<w:color w:val="{color}"/>{sp}<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>{va}'
            f'</w:rPr><w:t xml:space="preserve">{escape(text)}</w:t></w:r>')


MAX_CONDENSE = 0.04          # a title may be tightened by up to 4% ...
MAX_TRACK_PT = 0.40          # ... and never by more than 0.4pt a character


def fit_spacing(text, size_pt, width_pt, font):
    """Character spacing that keeps `text` on ONE line of `width_pt`, when it
    overflows by no more than MAX_CONDENSE - what the published covers do by
    hand (the combined English cover sets "Mines and Minerals ... Bill, 2026"
    2.4% tight to keep it on one line). 0.0 when it fits or would need more:
    then it wraps normally. `font` is a fitz.Font for measuring."""
    import math
    natural = font.text_length(text, fontsize=size_pt)
    if natural <= width_pt - 1.0 or natural > width_pt * (1 + MAX_CONDENSE):
        return 0.0
    # Word measures a little wider than the font's advances and keeps spacing
    # in 1/20pt steps: aim 3pt inside the line and round the step UP
    per_char = (natural - (width_pt - 3.0)) / max(1, len(text) - 1)
    twips = min(MAX_TRACK_PT * 20, math.ceil(per_char * 20))
    return -twips / 20.0


def sym(font, char, size_pt, color="000000"):
    """A symbol-font glyph (Wingdings tick etc.) - written as w:sym so the
    glyph is pinned regardless of the paragraph's text font."""
    half = int(round(size_pt * 2))
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:hint="default"/>'
            f'<w:color w:val="{color}"/><w:sz w:val="{half}"/><w:szCs w:val="{half}"/>'
            f'</w:rPr><w:sym w:font="{font}" w:char="{char}"/></w:r>')


def tab():
    return '<w:r><w:tab/></w:r>'


def para(runs, jc="left", before=0.0, after=0.0, line=None, rule="exact",
         left=0.0, hanging=0.0, right=0.0, tabs=(), keep_lines=True):
    """A paragraph with explicit geometry (points). `line` None = single.
    `tabs` = [(kind, pos_pt, leader)], e.g. ('right', 541, 'dot')."""
    tab_xml = ""
    if tabs:
        tab_xml = "<w:tabs>" + "".join(
            f'<w:tab w:val="{k}" w:pos="{pt_to_twips(p)}"' + (f' w:leader="{ld}"' if ld else "") + "/>"
            for k, p, ld in tabs) + "</w:tabs>"
    if line is None:
        spacing = f'<w:spacing w:before="{pt_to_twips(before)}" w:after="{pt_to_twips(after)}" w:line="240" w:lineRule="auto"/>'
    else:
        spacing = (f'<w:spacing w:before="{pt_to_twips(before)}" w:after="{pt_to_twips(after)}" '
                   f'w:line="{pt_to_twips(line)}" w:lineRule="{rule}"/>')
    ind = (f'<w:ind w:left="{pt_to_twips(left)}" w:right="{pt_to_twips(right)}"'
           + (f' w:hanging="{pt_to_twips(hanging)}"' if hanging else ' w:firstLine="0"') + "/>")
    keep = "<w:keepLines/>" if keep_lines else ""
    return (f'<w:p {W_NS}><w:pPr>{keep}<w:widowControl w:val="0"/>{tab_xml}{spacing}{ind}'
            f'<w:jc w:val="{jc}"/></w:pPr>{"".join(runs)}</w:p>')


# ------------------------------------------------------------------ the document

class Cover:
    """One Letter page; add() objects, then save()."""

    def __init__(self):
        self.doc = Document()
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Pt(PAGE_W), Pt(PAGE_H)
        for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(sec, side, Pt(0))
        sec.header_distance = sec.footer_distance = Pt(0)
        body = self.doc.element.body
        for p in list(body.findall(qn('w:p'))):
            body.remove(p)
        self.host = parse_xml(
            f'<w:p {W_NS}><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" '
            f'w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>')
        body.find(qn('w:sectPr')).addprevious(self.host)
        from docx.text.paragraph import Paragraph
        self._host_para = Paragraph(self.host, self.doc._body)

    def picture(self, path, rect, behind=False, link=None, name="picture"):
        x, y, w, h = rect
        add_anchored_image(self._host_para, self.doc.part, path, x, y, w, h,
                           behind=behind, hyperlink=link, name=name)

    def background(self, path):
        self.picture(path, (0.0, 0.0, PAGE_W, PAGE_H), behind=True, name="cover plate")

    def textbox(self, rect, paragraphs, anchor="t", fill=None, line=None,
                insets=(0.0, 0.0, 0.0, 0.0), name="text"):
        """`paragraphs` are XML strings from para(). `line` = (width_pt, hex)."""
        x, y, w, h = rect
        did = next_docpr_id()
        fill_xml = (f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill
                    else '<a:noFill/>')
        line_xml = (f'<a:ln w="{pt_to_emu(line[0])}"><a:solidFill><a:srgbClr val="{line[1]}"/>'
                    f'</a:solidFill><a:miter lim="800000"/></a:ln>' if line
                    else '<a:ln><a:noFill/></a:ln>')
        li, ti, ri, bi = (pt_to_emu(v) for v in insets)
        body = "".join(paragraphs) or f'<w:p {W_NS}/>'
        xml = f'''<w:r {NSDECL}><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>
  <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{251660000 + did}"
             behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="page"><wp:posOffset>{pt_to_emu(x)}</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="page"><wp:posOffset>{pt_to_emu(y)}</wp:posOffset></wp:positionV>
    <wp:extent cx="{pt_to_emu(w)}" cy="{pt_to_emu(h)}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapNone/>
    <wp:docPr id="{did}" name="{escape(name)}"/>
    <wp:cNvGraphicFramePr/>
    <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <wps:wsp>
        <wps:cNvSpPr txBox="1"/>
        <wps:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="{pt_to_emu(w)}" cy="{pt_to_emu(h)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          {fill_xml}{line_xml}
        </wps:spPr>
        <wps:txbx><w:txbxContent>{body}</w:txbxContent></wps:txbx>
        <wps:bodyPr rot="0" vert="horz" wrap="square" lIns="{li}" tIns="{ti}" rIns="{ri}" bIns="{bi}"
                    anchor="{anchor}" anchorCtr="0" upright="1"><a:noAutofit/></wps:bodyPr>
      </wps:wsp>
    </a:graphicData></a:graphic>
  </wp:anchor>
</w:drawing></mc:Choice></mc:AlternateContent></w:r>'''
        self.host.append(parse_xml(xml))

    def save(self, path):
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        self.doc.save(path)
        return path


# ------------------------------------------------------------------ measuring a render

def lines_in(pdf_path, rect, page_no=0):
    """[(y0, y1, x0, text, first_font)] of every text line inside `rect` of
    a rendered page - how a cover's fitting loop checks what Word did."""
    import fitz
    out = []
    with fitz.open(pdf_path) as d:
        page = d[page_no]
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                spans = [s for s in line["spans"] if s["text"].strip()]
                if not spans:
                    continue
                x0, y0, x1, y1 = line["bbox"]
                if y0 >= rect[1] - 1 and y1 <= rect[3] + 40 and x0 >= rect[0] - 2 and x0 <= rect[2]:
                    out.append((y0, y1, x0, "".join(s["text"] for s in spans), spans[0]["font"]))
    out.sort()
    return out


def content_lines(pdf_path, rect, page_no=0):
    """Text lines inside `rect` (x0, y0, x1, y1) of a render, top to bottom."""
    return [l for l in lines_in(pdf_path, rect, page_no) if l[0] < rect[3]]
