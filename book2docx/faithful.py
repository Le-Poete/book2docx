# -*- coding: utf-8 -*-
"""faithful 管线：整页高清图 + 隐形 OCR 文字层。

100% 保留原书版式外观；文字层置于透明文本框中，可搜索、可复制、可批注。
"""
from __future__ import annotations

import io
import os

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.shared import Pt

from .common import PageInfo, PageKind, classify_page
from .docxkit import (add_background_image, add_hidden_run, add_transparent_textbox,
                      make_paragraph, sanitize_xml_text, set_page)


def _extract_page_image(page: fitz.Page, dpi: int = 250):
    """提取页面图像：优先取原始嵌入扫描图（零重编码），失败则按 dpi 渲染。"""
    infos = page.get_image_info(xrefs=True)
    if len(infos) == 1:
        b = infos[0]["bbox"]
        cov = (b[2] - b[0]) * (b[3] - b[1]) / (page.rect.width * page.rect.height)
        if cov > 0.9:
            try:
                raw = page.parent.extract_image(infos[0]["xref"])
                if raw["image"] and raw["width"] >= page.rect.width:  # 分辨率足够
                    return raw["image"], raw["ext"]
            except Exception:
                pass
    pix = page.get_pixmap(dpi=dpi)
    return pix.tobytes("png"), "png"


def _layer_lines(page: fitz.Page, pad: float = 2.0):
    """从 OCR 文字层取行级 (bbox, text)。"""
    d = page.get_text("dict")
    lines = []
    for blk in d["blocks"]:
        if blk["type"] != 0:
            continue
        for line in blk["lines"]:
            text = "".join(sp["text"] for sp in line["spans"]).strip()
            if text:
                lines.append((fitz.Rect(line["bbox"]), sanitize_xml_text(text)))
    return lines


def convert(doc: fitz.Document, out_path: str, pages: range | None = None,
            image_dpi: int = 250, progress=None):
    """faithful 转换。pages: 0-based 页码范围（默认全部）。"""
    rng = pages if pages is not None else range(doc.page_count)
    document = Document()

    # 默认样式：隐藏层用宋体
    style = document.styles["Normal"]
    style.font.name = "Times New Roman"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    first = True
    for i in rng:
        page = doc[i]
        info: PageInfo = classify_page(page)
        sec = document.sections[0] if first else document.add_section(WD_SECTION.NEW_PAGE)
        if first:
            first = False
        set_page(sec, page.rect.width, page.rect.height, (0, 0, 0, 0))

        img_bytes, ext = _extract_page_image(page, dpi=image_dpi)
        p = document.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.0

        # 注册图片关系并作为背景图插入
        rid, _ = document.part.get_or_add_image(io.BytesIO(img_bytes))
        add_background_image(p, img_bytes, page.rect.width, page.rect.height, rid=rid)

        # 隐形文字层（OCR 层质量尚可时才放）
        if info.kind != PageKind.SCAN_RAW and info.ocr_quality > 0.5:
            lines = _layer_lines(page)
            if lines:
                def builder(txbx, lines=lines):
                    for rect, text in lines:
                        para = make_paragraph(txbx, line=240)
                        add_hidden_run(para, text)

                add_transparent_textbox(p, 0, 0, page.rect.width, page.rect.height,
                                        builder, name=f"TextLayer{i}")
        if progress:
            progress(i, len(rng))
    document.save(out_path)
    return out_path
