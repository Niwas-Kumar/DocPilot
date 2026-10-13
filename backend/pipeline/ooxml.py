# -*- coding: utf-8 -*-
"""Low-level OOXML helpers shared by the template builder and styler."""
import re

from docx.oxml.ns import qn, nsmap
from docx.oxml import parse_xml
from lxml import etree

EMU_PER_PT = 12700

# A Google Doc that uses the Tabs feature exports the tab's NAME as a real
# paragraph at the top of the body ("Tab 1"), usually followed by a hard page
# break so the tab's content starts fresh. Neither belongs in a booklet: the
# client reported a page carrying nothing but "Tab 1", with the first topic
# pushed to the page after it.
GOOGLE_TAB_RE = re.compile(r"^tab\s*\d+$", re.I)

# Styles that mean "real content has started" - the scan never passes one.
CONTENT_STYLES = {"heading 1", "heading 2", "heading 3", "title"}


def drop_google_tab_frontmatter(body, style_of):
    """Remove a Google Docs tabs artifact from the FRONT of a document.

    `style_of(paragraph_element)` returns that paragraph's style name (any
    case) or None. Only the region before the first real content is touched -
    the first table, picture, heading/title, or any ordinary text - so a
    genuine "Tab 2" inside an article can never be swallowed.

    `w:pageBreakBefore` is stripped elsewhere by both stylers, but a
    `<w:br w:type="page"/>` RUN is not, and that is what actually split the
    page; deleting the artifact paragraph takes its break with it.

    Returns (paragraphs removed, stray page breaks removed).
    """
    removed = breaks = 0
    for el in [e for e in body if e.tag in (qn('w:p'), qn('w:tbl'))]:
        if el.tag == qn('w:tbl'):
            break                                   # real content
        if next(el.iter(qn('w:drawing')), None) is not None:
            break
        if (style_of(el) or "").strip().lower() in CONTENT_STYLES:
            break
        text = "".join(t.text or "" for t in el.iter(qn('w:t'))).strip()
        if GOOGLE_TAB_RE.match(text):
            el.getparent().remove(el)               # takes its page break too
            removed += 1
            continue
        if text:
            break                                   # ordinary front matter
        for br in el.findall(qn('w:r') + '/' + qn('w:br')):
            if br.get(qn('w:type')) == 'page':
                br.getparent().remove(br)
                breaks += 1
    return removed, breaks

NSDECL = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"'
)


def pt_to_emu(pt: float) -> int:
    return int(round(pt * EMU_PER_PT))


def pt_to_twips(pt: float) -> int:
    return int(round(pt * 20))


_docpr_id = [9000]


def next_docpr_id() -> int:
    _docpr_id[0] += 1
    return _docpr_id[0]


def make_anchored_image_xml(rid: str, x_pt: float, y_pt: float, w_pt: float, h_pt: float,
                            behind: bool = False, hlink_rid: str | None = None,
                            name: str = "img") -> str:
    """wp:anchor drawing positioned relative to the PAGE, no text wrap."""
    did = next_docpr_id()
    hlink = f'<a:hlinkClick xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:id="{hlink_rid}"/>' if hlink_rid else ""
    return f'''<w:drawing {NSDECL}>
  <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
             relativeHeight="{251658240 + did}" behindDoc="{1 if behind else 0}" locked="0"
             layoutInCell="1" allowOverlap="1">
    <wp:simplePos x="0" y="0"/>
    <wp:positionH relativeFrom="page"><wp:posOffset>{pt_to_emu(x_pt)}</wp:posOffset></wp:positionH>
    <wp:positionV relativeFrom="page"><wp:posOffset>{pt_to_emu(y_pt)}</wp:posOffset></wp:positionV>
    <wp:extent cx="{pt_to_emu(w_pt)}" cy="{pt_to_emu(h_pt)}"/>
    <wp:effectExtent l="0" t="0" r="0" b="0"/>
    <wp:wrapNone/>
    <wp:docPr id="{did}" name="{name}">{hlink}</wp:docPr>
    <wp:cNvGraphicFramePr>
      <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
    </wp:cNvGraphicFramePr>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:nvPicPr>
            <pic:cNvPr id="{did}" name="{name}"/>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="{rid}"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="{pt_to_emu(w_pt)}" cy="{pt_to_emu(h_pt)}"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:anchor>
</w:drawing>'''


def add_anchored_image(paragraph, part, image_path, x_pt, y_pt, w_pt, h_pt,
                       behind=False, hyperlink=None, name="img"):
    """Insert a page-anchored image into `paragraph` living in `part` (header/footer/body)."""
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    rid, _ = part.get_or_add_image(image_path)
    hlink_rid = None
    if hyperlink:
        hlink_rid = part.relate_to(hyperlink, RT.HYPERLINK, is_external=True)
    xml = make_anchored_image_xml(rid, x_pt, y_pt, w_pt, h_pt, behind, hlink_rid, name)
    run = paragraph.add_run()
    run._r.append(parse_xml(xml))
    return run


def set_run_fonts(rpr, ascii_font="Times New Roman", cs_font="Kokila", size_half_pts=22):
    """Force fonts + size on an existing w:rPr element (creates children as needed)."""
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = rpr.makeelement(qn('w:rFonts'), {})
        rpr.insert(0, rfonts)
    rfonts.set(qn('w:ascii'), ascii_font)
    rfonts.set(qn('w:hAnsi'), ascii_font)
    rfonts.set(qn('w:cs'), cs_font)
    if rfonts.get(qn('w:eastAsia')):
        rfonts.set(qn('w:eastAsia'), ascii_font)
    for tag, val in ((qn('w:sz'), str(size_half_pts)), (qn('w:szCs'), str(size_half_pts))):
        el = rpr.find(tag)
        if el is None:
            el = rpr.makeelement(tag, {})
            rpr.append(el)
        el.set(qn('w:val'), val)


def get_or_add(parent, tag, insert_first=False):
    el = parent.find(qn(tag))
    if el is None:
        el = parent.makeelement(qn(tag), {})
        if insert_first:
            parent.insert(0, el)
        else:
            parent.append(el)
    return el


def set_color(rpr, hexcolor):
    c = get_or_add(rpr, 'w:color')
    c.set(qn('w:val'), hexcolor)


def set_bold(rpr, on=True):
    for tag in ('w:b', 'w:bCs'):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.append(el)
        if on:
            if el.get(qn('w:val')):
                del el.attrib[qn('w:val')]
        else:
            el.set(qn('w:val'), '0')


def para_shading(ppr, fill_hex):
    shd = get_or_add(ppr, 'w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_hex)


def para_spacing(p_pr, before_pt=0, after_pt=2, line_pt=14, rule="atLeast"):
    sp = get_or_add(p_pr, 'w:spacing')
    sp.set(qn('w:before'), str(pt_to_twips(before_pt)))
    sp.set(qn('w:after'), str(pt_to_twips(after_pt)))
    if line_pt:
        sp.set(qn('w:line'), str(pt_to_twips(line_pt)))
        sp.set(qn('w:lineRule'), rule)


def clear_indent(ppr):
    ind = ppr.find(qn('w:ind'))
    if ind is not None:
        ppr.remove(ind)


def get_numpr(p):
    ppr = p.find(qn('w:pPr'))
    if ppr is None:
        return None
    return ppr.find(qn('w:numPr'))


def get_ilvl(p):
    npr = get_numpr(p)
    if npr is None:
        return None
    ilvl = npr.find(qn('w:ilvl'))
    return int(ilvl.get(qn('w:val'))) if ilvl is not None else 0
