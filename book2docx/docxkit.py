# -*- coding: utf-8 -*-
"""python-docx 底层 XML 辅助：浮动图片、透明文本框、页面设置。"""
from __future__ import annotations

import re

from docx.enum.section import WD_SECTION
from docx.oxml import parse_xml
from docx.oxml.ns import nsmap, qn
from docx.shared import Emu

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def set_page(section, width_pt: float, height_pt: float,
             margins_pt: tuple[float, float, float, float] = (0, 0, 0, 0)):
    """设置页面尺寸与页边距（pt）。margins = (上, 下, 左, 右)。"""
    sec_pr = section._sectPr
    pg_sz = sec_pr.find(qn("w:pgSz"))
    if pg_sz is None:
        pg_sz = sec_pr.makeelement(qn("w:pgSz"), {})
        sec_pr.append(pg_sz)
    pg_sz.set(qn("w:w"), str(int(round(width_pt * 20))))
    pg_sz.set(qn("w:h"), str(int(round(height_pt * 20))))
    t, b, l, r = margins_pt
    pg_mar = sec_pr.find(qn("w:pgMar"))
    if pg_mar is None:
        pg_mar = sec_pr.makeelement(qn("w:pgMar"), {})
        sec_pr.append(pg_mar)
    pg_mar.set(qn("w:top"), str(int(round(t * 20))))
    pg_mar.set(qn("w:bottom"), str(int(round(b * 20))))
    pg_mar.set(qn("w:left"), str(int(round(l * 20))))
    pg_mar.set(qn("w:right"), str(int(round(r * 20))))
    pg_mar.set(qn("w:header"), "0")
    pg_mar.set(qn("w:footer"), "0")
    pg_mar.set(qn("w:gutter"), "0")


_id_counter = [1000]


def _next_id() -> int:
    _id_counter[0] += 1
    return _id_counter[0]


def add_background_image(paragraph, image_path: str | bytes, width_pt: float,
                         height_pt: float, rid: str | None = None):
    """在段落中插入浮动图片：锚定页面左上角、衬于文字下方、不换行。"""
    cx, cy = int(round(width_pt * 12700)), int(round(height_pt * 12700))
    doc_pr_id = _next_id()
    embed_rid = rid or f"rBg{doc_pr_id}"
    xml = f"""<w:r xmlns:w="{W_NS}"><w:drawing>
<wp:anchor xmlns:wp="{WP_NS}" distT="0" distB="0" distL="0" distR="0"
 simplePos="0" relativeHeight="0" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/>
<wp:docPr id="{doc_pr_id}" name="PageImage{doc_pr_id}"/>
<wp:cNvGraphicFramePr/>
<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="{PIC_NS}">
<pic:pic xmlns:pic="{PIC_NS}">
<pic:nvPicPr><pic:cNvPr id="{doc_pr_id}" name="PageImage{doc_pr_id}"/><pic:cNvPicPr/></pic:nvPicPr>
<pic:blipFill><a:blip r:embed="{embed_rid}" xmlns:r="{R_NS}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"""
    run = paragraph.add_run()
    run._r.append(parse_xml(xml))
    return embed_rid


def add_transparent_textbox(paragraph, x_pt: float, y_pt: float, w_pt: float, h_pt: float,
                            content_builder, name: str = "TextLayer"):
    """插入绝对定位透明文本框（页面锚定、无填充无边框）。

    content_builder(txbx_content_element) 向 w:txbxContent 填充段落。
    """
    x, y = int(round(x_pt * 12700)), int(round(y_pt * 12700))
    cx, cy = int(round(w_pt * 12700)), int(round(h_pt * 12700))
    doc_pr_id = _next_id()
    xml = f"""<w:r xmlns:w="{W_NS}"><w:drawing>
<wp:anchor xmlns:wp="{WP_NS}" distT="0" distB="0" distL="0" distR="0"
 simplePos="0" relativeHeight="251650000" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
<wp:simplePos x="0" y="0"/>
<wp:positionH relativeFrom="page"><wp:posOffset>{x}</wp:posOffset></wp:positionH>
<wp:positionV relativeFrom="page"><wp:posOffset>{y}</wp:posOffset></wp:positionV>
<wp:extent cx="{cx}" cy="{cy}"/>
<wp:effectExtent l="0" t="0" r="0" b="0"/>
<wp:wrapNone/>
<wp:docPr id="{doc_pr_id}" name="{name}{doc_pr_id}"/>
<wp:cNvGraphicFramePr/>
<a:graphic xmlns:a="{A_NS}"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
<wps:cNvSpPr txBox="1"/><wps:spPr>
<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>
<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>
<a:ln><a:noFill/></a:ln></wps:spPr>
<wps:txbx><w:txbxContent/></wps:txbx>
<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"
 anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>
</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"""
    run = paragraph.add_run()
    el = parse_xml(xml)
    run._r.append(el)
    txbx = el.find(f".//{qn('w:txbxContent')}")
    content_builder(txbx)
    return doc_pr_id


def make_paragraph(parent, spacing_before=0, spacing_after=0, line=None):
    """在 txbxContent 等父元素里造一个空段落。line: 行距(240=单倍)。"""
    p = parent.makeelement(qn("w:p"), {})
    ppr = p.makeelement(qn("w:pPr"), {})
    spac = p.makeelement(qn("w:spacing"), {})
    spac.set(qn("w:before"), str(int(spacing_before * 20)))
    spac.set(qn("w:after"), str(int(spacing_after * 20)))
    if line:
        spac.set(qn("w:line"), str(int(line)))
        spac.set(qn("w:lineRule"), "auto")
    ppr.append(spac)
    p.append(ppr)
    parent.append(p)
    return p


def add_hidden_run(p, text: str, size_half_pt: int = 2, color: str = "FFFFFF"):
    r = p.makeelement(qn("w:r"), {})
    rpr = p.makeelement(qn("w:rPr"), {})
    sz = p.makeelement(qn("w:sz"), {})
    sz.set(qn("w:val"), str(size_half_pt))
    rpr.append(sz)
    szcs = p.makeelement(qn("w:szCs"), {})
    szcs.set(qn("w:val"), str(size_half_pt))
    rpr.append(szcs)
    c = p.makeelement(qn("w:color"), {})
    c.set(qn("w:val"), color)
    rpr.append(c)
    r.append(rpr)
    t = p.makeelement(qn("w:t"), {})
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p.append(r)
    return r


def sanitize_xml_text(s: str) -> str:
    """去掉 XML 1.0 非法控制字符。"""
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]", "", s)
