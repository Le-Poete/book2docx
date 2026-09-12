# -*- coding: utf-8 -*-
"""digital 管线：原生文字层 PDF → Word 结构化重建。

保留字体（中英文分开映射）、字号、颜色、粗斜体、段落对齐、行距；
图片提取嵌入；复杂矢量结构（表格）回退为区域裁图。
"""
from __future__ import annotations

import io

import fitz
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Emu, Pt, RGBColor

from .common import map_font, span_flags
from .docxkit import sanitize_xml_text, set_page


def _garbled_ratio(text: str) -> float:
    """异常字符占比：私有区/错映射区字符（Word 导出 PDF 的公式编码病灶）。"""
    if not text:
        return 0.0
    bad = ok = 0
    for c in text:
        cp = ord(c)
        if 0xE000 <= cp <= 0xF8FF:
            bad += 1
        elif (c.isascii() or 0x4E00 <= cp <= 0x9FFF or 0x3000 <= cp <= 0x303F
              or 0xFF00 <= cp <= 0xFFEF or 0x0370 <= cp <= 0x03FF
              or c in "≤≥≠∑∈√πμα°±×÷"):
            ok += 1
        else:
            bad += 1
    return bad / max(1, bad + ok)


def _detect_formula_fonts(page: fitz.Page) -> set[str]:
    """按字体统计字符构成：匿名 CID 字体若几乎不含 CJK/常规文本，
    判定为公式字体（Word/LaTeX 导出 PDF 的无 ToUnicode 子集字体）。"""
    stats: dict[str, list[int]] = {}
    d = page.get_text("dict")
    for blk in d["blocks"]:
        if blk["type"] != 0:
            continue
        for line in blk["lines"]:
            for sp in line["spans"]:
                t = sp["text"].strip()
                if not t:
                    continue
                f = sp["font"]
                st = stats.setdefault(f, [0, 0])
                for c in t:
                    cp = ord(c)
                    st[1] += 1
                    if (0x4E00 <= cp <= 0x9FFF or c.isascii() and c.isalnum()
                            or c in "，。；：、！？（）《》""''"):
                        st[0] += 1
    out = set()
    for f, (good, total) in stats.items():
        if total >= 20 and good / total < 0.3:
            out.add(f)
    return out


def _render_clip(page: fitz.Page, bbox, dpi: int = 300) -> bytes:
    import cv2
    import numpy as np
    r = fitz.Rect(max(0, bbox[0] - 1), max(0, bbox[1] - 1),
                  min(page.rect.width, bbox[2] + 1),
                  min(page.rect.height, bbox[3] + 1))
    pix = page.get_pixmap(clip=r, matrix=fitz.Matrix(dpi / 72, dpi / 72))
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
    if pix.n == 4:
        img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
    ok, buf = cv2.imencode(".png", img)
    return buf.tobytes() if ok else b""


def _collect_lines(page: fitz.Page):
    d = page.get_text("dict")
    lines = []
    for blk in d["blocks"]:
        if blk["type"] == 1:
            lines.append({"type": "image", "bbox": list(blk["bbox"]), "img": blk["image"]})
            continue
        for line in blk["lines"]:
            spans = [sp for sp in line["spans"] if sp["text"].strip()]
            if not spans:
                continue
            lines.append({"type": "line", "bbox": list(line["bbox"]), "spans": spans})
    lines.sort(key=lambda l: (l["bbox"][1], l["bbox"][0]))
    return lines


def _merge_paragraphs(lines, page_height: float):
    """行合并成段落：y 间距小且 x 范围连续 → 同段。"""
    paras = []
    cur: list[dict] = []
    for ln in lines:
        if ln["type"] == "image":
            if cur:
                paras.append(cur)
                cur = []
            paras.append([ln])
            continue
        if cur:
            prev = cur[-1]
            gap = ln["bbox"][1] - prev["bbox"][3]
            h = prev["bbox"][3] - prev["bbox"][1]
            prev_x1 = prev["bbox"][2]
            x_overlap = ln["bbox"][0] < prev_x1 - 20 or abs(ln["bbox"][0] - prev["bbox"][0]) < 3
            if gap < h * 0.9 and x_overlap:
                cur.append(ln)
                continue
            paras.append(cur)
            cur = [ln]
        else:
            cur = [ln]
    if cur:
        paras.append(cur)
    return paras


def _collapse_matrix_clusters(lines: list[dict]) -> list[dict]:
    """把散落的短行簇（矩阵/行列式的元素列）合并为整块裁图行。

    判据：连续 ≥3 行，每行 ≤4 个字符，行距连续，x 中心集中。
    """
    def is_short(ln):
        if ln["type"] != "line":
            return False
        t = "".join(sp["text"] for sp in ln["spans"]).strip()
        return 0 < len(t) <= 4

    out: list[dict] = []
    i = 0
    while i < len(lines):
        if not is_short(lines[i]):
            out.append(lines[i])
            i += 1
            continue
        j = i
        while j + 1 < len(lines) and is_short(lines[j + 1]):
            h = lines[j]["bbox"][3] - lines[j]["bbox"][1]
            gap = lines[j + 1]["bbox"][1] - lines[j]["bbox"][3]
            cx_a = (lines[j]["bbox"][0] + lines[j]["bbox"][2]) / 2
            cx_b = (lines[j + 1]["bbox"][0] + lines[j + 1]["bbox"][2]) / 2
            if gap > 2.2 * max(1, h) or abs(cx_b - cx_a) > 3.5 * max(1, h):
                break
            j += 1
        if j - i + 1 >= 3:
            bb = [min(l["bbox"][0] for l in lines[i:j + 1]),
                  min(l["bbox"][1] for l in lines[i:j + 1]),
                  max(l["bbox"][2] for l in lines[i:j + 1]),
                  max(l["bbox"][3] for l in lines[i:j + 1])]
            out.append({"type": "line", "bbox": bb, "spans": [],
                        "matrix_cluster": True})
        else:
            out.extend(lines[i:j + 1])
        i = j + 1
    return out


def build_page_digital(document: Document, page: fitz.Page, first: bool = False):
    pr = page.rect
    lines = _collect_lines(page)
    body = [l for l in lines
            if l["bbox"][1] > pr.height * 0.04 and l["bbox"][3] < pr.height * 0.96]
    body = _collapse_matrix_clusters(body)
    if body:
        mx0 = min(l["bbox"][0] for l in body)
        my0 = min(l["bbox"][1] for l in body)
        mx1 = max(l["bbox"][2] for l in body)
        my1 = max(l["bbox"][3] for l in body)
        margins = (max(24, my0 - 2), max(24, pr.height - my1 - 2),
                   max(24, mx0 - 3), max(24, pr.width - mx1 - 3))
    else:
        margins = (72, 72, 72, 72)

    sec = document.sections[0] if first else document.add_section(WD_SECTION.NEW_PAGE)
    set_page(sec, pr.width, pr.height, margins)

    paras = _merge_paragraphs(body, pr.height)
    formula_fonts = _detect_formula_fonts(page)
    for para in paras:
        if len(para) == 1 and para[0]["type"] == "image":
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run()
            bb = para[0]["bbox"]
            try:
                run.add_picture(fitz.Pixmap(para[0]["img"]).tobytes("png"),
                                width=Emu(int((bb[2] - bb[0]) * 12700)))
            except Exception:
                pass
            continue

        p = document.add_paragraph()
        # 对齐推断
        x0s = [l["bbox"][0] for l in para]
        x1s = [l["bbox"][2] for l in para]
        blk_x0, blk_x1 = min(x0s), max(x1s)
        w = blk_x1 - blk_x0
        if len(para) == 1 and w < (pr.width - margins[2] - margins[3]) * 0.6:
            lm = para[0]["bbox"][0] - blk_x0
            rm = blk_x1 - para[0]["bbox"][2]
            if lm > 20 and abs(lm - rm) < 15:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # 首行缩进
        indent = para[0]["bbox"][0] - blk_x0
        if indent > 8:
            p.paragraph_format.first_line_indent = Pt(indent)
        # 行距
        if len(para) > 1:
            gaps = [para[i + 1]["bbox"][1] - para[i]["bbox"][1] for i in range(len(para) - 1)]
            g = sorted(gaps)[len(gaps) // 2]
            if g > 4:
                p.paragraph_format.line_spacing = Pt(g)

        for ln in para:
            if ln.get("matrix_cluster"):
                p2 = document.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.paragraph_format.space_before = Pt(0)
                p2.paragraph_format.space_after = Pt(0)
                run2 = p2.add_run()
                data = _render_clip(page, ln["bbox"])
                if data:
                    bb = ln["bbox"]
                    run2.add_picture(io.BytesIO(data),
                                     width=Emu(int((bb[2] - bb[0]) * 12700)))
                continue
            ln_text = "".join(sp["text"] for sp in ln.get("spans", []))
            font_bad = any(sp["font"] in formula_fonts for sp in ln.get("spans", []))
            if ln["type"] == "line" and (font_bad or _garbled_ratio(ln_text) > 0.34):
                # 公式字体/私有编码行（Word 导出 PDF 常见）→ 高清裁图保真
                p2 = document.add_paragraph()
                p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p2.paragraph_format.space_before = Pt(0)
                p2.paragraph_format.space_after = Pt(0)
                run2 = p2.add_run()
                data = _render_clip(page, ln["bbox"])
                if data:
                    bb = ln["bbox"]
                    run2.add_picture(io.BytesIO(data),
                                     width=Emu(int((bb[2] - bb[0]) * 12700)))
                continue
            for sp in ln["spans"]:
                text = sanitize_xml_text(sp["text"])
                if not text:
                    continue
                cn, west = map_font(sp["font"])
                fl = span_flags(sp["font"], sp.get("flags", 0))
                run = p.add_run(text)
                run.font.name = west
                run.font.size = Pt(round(sp["size"], 1))
                run.font.bold = fl["bold"]
                run.font.italic = fl["italic"]
                c = sp.get("color", 0)
                if c:
                    run.font.color.rgb = RGBColor((c >> 16) & 255, (c >> 8) & 255, c & 255)
                run._element.rPr.rFonts.set(qn("w:eastAsia"), cn)


def convert(doc: fitz.Document, out_path: str, pages: range | None = None):
    rng = pages if pages is not None else range(doc.page_count)
    document = Document()
    st = document.styles["Normal"]
    st.font.name = "Times New Roman"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    for k, i in enumerate(rng):
        build_page_digital(document, doc[i], first=(k == 0))
    document.save(out_path)
    return out_path
