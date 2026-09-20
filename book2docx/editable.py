# -*- coding: utf-8 -*-
"""editable 管线：扫描书 → 可编辑 Word。

正文 OCR 重建为真文字（字号/缩进/对齐/行距按几何还原），
公式/插图区域高清裁图保真，表格按网格线重建为原生 Word 表格。
"""
from __future__ import annotations

import io
from pathlib import Path

import cv2
import fitz
import numpy as np
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Emu, Pt, RGBColor

from .common import PageKind, classify_page
from .docxkit import set_page, sanitize_xml_text
from .segment import (Block, BlockKind, Line, W, binarize, classify_formula_line,
                      cluster_blocks, cluster_lines, dedup_lines, detect_grid_tables,
                      detect_ink_regions, deskew, estimate_skew, extract_words_pt,
                      get_words, ink_stats, render_gray, split_lines_by_gap,
                      table_grid, transform_pt_bbox)

_CODE_RE = __import__("re").compile(
    r"[A-Za-z_][A-Za-z0-9_]{2,}\s*\(|'[^']*'|\"[^\"]*\"|\bfor\b|\bif\b|\bwhile\b"
    r"|\bend\b|==|<=|>=|\[[^\]]*\]\s*[=,;)]")


def _code_hint_line(t: str) -> bool:
    if not t:
        return False
    n_ascii = sum(1 for c in t if c.isascii())
    if n_ascii < len(t) * 0.6:        # 中文主导行不可能是代码
        return False
    if _CODE_RE.search(t):
        return True
    return t.count("*") >= 3          # MATLAB 行的典型乘号密度


# ---------------------------------------------------------------- 单页分析

def _words_from_full_ocr(img: np.ndarray, zoom: float) -> list[W]:
    """在校正后的整页图上跑 RapidOCR，词块坐标为校正系 pt。"""
    from .segment import rapidocr
    result, _ = rapidocr()(img)
    words: list[W] = []
    for pts, text, conf in result or []:
        t = str(text).strip()
        if not t:
            continue
        xs = [q[0] for q in pts]
        ys = [q[1] for q in pts]
        words.append(W(min(xs) / zoom, min(ys) / zoom, max(xs) / zoom,
                       max(ys) / zoom, t, (max(ys) - min(ys)) / zoom * 0.85,
                       float(conf)))
    return words


def analyze_page(page: fitz.Page, dpi: int = 300, reocr: bool = False) -> dict:
    pr = page.rect
    zoom0 = dpi / 72.0
    if reocr:
        # 全页 RapidOCR 重识别：在原始渲染图上检测文字，随后随 deskew 矩阵
        # 变换坐标（与 PDF 层词块同一处理路径）
        img0 = render_gray(page, dpi)
        angle = estimate_skew(img0)
        if abs(angle) > 0.05:
            img, M = deskew(img0, angle)
        else:
            img, M = img0, np.array([[1.0, 0, 0], [0, 1.0, 0]])
        words = _words_from_full_ocr(img, zoom0)
        source = "rapidocr_full"
        words0 = words
    else:
        words0, source = get_words(page)

    if not words0:
        return {"blocks": [], "margins": (72, 72, 72, 72), "base_h": 12.0,
                "img": None, "page": pr, "words": [], "base_ink": 0.08,
                "angle": 0.0, "source": source}

    zoom = dpi / 72.0
    if not reocr:
        img0 = render_gray(page, dpi)
        # 倾斜校正：扫描页常带 ~0.5° 旋转，斜表格线会破坏逐行检测
        angle = estimate_skew(img0)
        if abs(angle) > 0.05:
            img, M = deskew(img0, angle)
        else:
            img, M = img0, np.array([[1.0, 0, 0], [0, 1.0, 0]])
        words = []
        for wd in words0:
            nb = transform_pt_bbox([wd.x0, wd.y0, wd.x1, wd.y1], M, zoom)
            words.append(W(nb[0], nb[1], nb[2], nb[3], wd.text, wd.size, wd.conf))
        words0 = words

    bw = binarize(img)

    lines = cluster_lines(words)
    med_h = float(np.median([l.h for l in lines]))
    base_h = med_h
    lines = split_lines_by_gap(lines, base_h)
    lines = dedup_lines(lines, base_h)

    # 页眉页脚：页面上下 5.5% 内的行 → SKIP
    head_zone = pr.height * 0.055
    keep_lines, skip_lines = [], []
    for ln in lines:
        (skip_lines if (ln.bbox[3] < head_zone or ln.bbox[1] > pr.height - head_zone)
         else keep_lines).append(ln)

    blocks = cluster_blocks(keep_lines, base_h)

    # 表格区域（投影线检测，基于灰度图）
    tables, _lines = detect_grid_tables(img, zoom, pr)
    used_lines: set[int] = set()
    out_blocks: list[Block] = []
    table_bboxes = []
    for tb in tables:
        inside = [i for i, ln in enumerate(keep_lines)
                  if _inside(ln.bbox, tb, 0.6)]
        if len(inside) >= 2:
            table_bboxes.append(tb)
    # 表格行从普通块中剔除；块内因剔除产生的碎片重新聚类，但保留原块边界
    for blk in blocks:
        remaining = [ln for ln in blk.lines
                     if not any(_inside(ln.bbox, tb, 0.5) for tb in table_bboxes)]
        if not remaining:
            continue
        if len(remaining) == len(blk.lines):
            blk.bbox = _bbox_of_lines(remaining)
            out_blocks.append(blk)
        else:
            out_blocks.extend(cluster_blocks(remaining, base_h))
    for tb in table_bboxes:
        out_blocks.append(Block(BlockKind.TABLE, tb, meta={"grid": True}))

    # 插图区域
    for rg in detect_ink_regions(bw, zoom, words, pr.width * pr.height):
        if any(_overlap_ratio(rg, tb) > 0.3 for tb in table_bboxes):
            continue
        if _overlap_ratio(rg, [0, 0, pr.width, head_zone]) > 0.6:
            continue
        out_blocks.append(Block(BlockKind.IMAGE, rg))

    # OCR 行框交叉产生的假分块：合并高度重叠的块
    out_blocks = _merge_overlapping_blocks(out_blocks, base_h)

    # 公式/代码行分类 → 连续同类行重组为纯子块
    base_ink = _body_ink(bw, zoom, keep_lines, base_h)
    rebuilt: list[Block] = []
    for blk in out_blocks:
        if blk.kind != BlockKind.TEXT:
            rebuilt.append(blk)
            continue
        # 组标记：f=公式(裁图)  c=代码(裁图)  t=正文(可编辑)
        groups: list[tuple[str, list[Line]]] = []
        for ln in blk.lines:
            if classify_formula_line(ln, base_h, bw, zoom):
                k = "f"
            elif _code_hint_line(ln.text):
                k = "c"
            else:
                k = "t"
            if groups and groups[-1][0] == k:
                groups[-1][1].append(ln)
            else:
                groups.append((k, [ln]))
        # 夹在公式组之间的短正文碎片（OCR 把公式行读出几个汉字）→ 并入公式
        i = 1
        while i < len(groups) - 1:
            k, ls = groups[i]
            if (k == "t" and len("".join(l.text for l in ls).strip()) <= 14
                    and groups[i - 1][0] == "f" and groups[i + 1][0] == "f"):
                groups[i - 1][1].extend(ls)
                del groups[i]
                groups[i - 1][1].sort(key=lambda l: l.bbox[1])
                continue
            i += 1
        # 夹在正文组之间的短公式组 → 回退为正文（行内公式的正文行，
        # 如"设 i=1,2,3 分别表示…"；OCR 文本可读。乱码公式保持裁图）
        i = 1
        while i < len(groups) - 1:
            k, ls = groups[i]
            joined = "".join(l.text for l in ls)
            n_ascii = sum(1 for c in joined if c.isascii())
            if (k == "f" and len(ls) <= 2
                    and groups[i - 1][0] == "t" and groups[i + 1][0] == "t"
                    and max(len(l.text) for l in ls) <= 55
                    and n_ascii >= len(joined) * 0.7):
                groups[i] = ("t", ls)
                if groups[i - 1][0] == "t":
                    groups[i - 1][1].extend(ls)
                    groups[i - 1][1].sort(key=lambda l: l.bbox[1])
                    del groups[i]
                continue
            i += 1
        # 代码传播：紧邻代码组的行组若是碎片（OCR 读碎的代码行）→ 并入代码
        i = 0
        while i < len(groups):
            k, ls = groups[i]
            if k == "t" and len(ls) <= 3:
                n_code = sum(1 for ln in ls if _code_hint_line(ln.text))
                near_code = ((i > 0 and groups[i - 1][0] == "c")
                             or (i + 1 < len(groups) and groups[i + 1][0] == "c"))
                if near_code and (n_code >= 1 or len("".join(l.text for l in ls)) < 25):
                    groups[i] = ("c", ls)
            i += 1
        # 组 → 块
        for k, ls in groups:
            bb = _bbox_of_lines(ls)
            if k == "f":
                rebuilt.append(Block(BlockKind.FORMULA, bb, ls))
            elif k == "c":
                rebuilt.append(Block(BlockKind.TEXT, bb, ls, meta={"code": True}))
            else:
                rebuilt.append(Block(BlockKind.TEXT, bb, ls))
    out_blocks = rebuilt

    # 相邻公式块合并（碎片行拆开或 expand 互侵的公式组回归整块）
    merged: list[Block] = []
    for blk in out_blocks:
        if (blk.kind == BlockKind.FORMULA and merged
                and merged[-1].kind == BlockKind.FORMULA):
            prev = merged[-1]
            x_ov = min(prev.bbox[2], blk.bbox[2]) - max(prev.bbox[0], blk.bbox[0])
            gap = blk.bbox[1] - prev.bbox[3]
            if x_ov > -base_h and gap < base_h * 1.8:
                prev.lines.extend(blk.lines)
                prev.lines.sort(key=lambda l: l.bbox[1])
                prev.bbox = _bbox_of_lines(prev.lines)
                continue
        merged.append(blk)
    out_blocks = merged

    # 碎片文本块夹在代码块之间 → 也是代码（OCR 把代码行读碎成独立小块）
    changed = True
    while changed:
        changed = False
        for i, blk in enumerate(out_blocks):
            if (blk.kind == BlockKind.TEXT and not blk.meta.get("code")
                    and len(blk.lines) <= 3):
                prev_c = (i > 0 and out_blocks[i - 1].kind == BlockKind.TEXT
                          and out_blocks[i - 1].meta.get("code"))
                next_c = (i + 1 < len(out_blocks)
                          and out_blocks[i + 1].kind == BlockKind.TEXT
                          and out_blocks[i + 1].meta.get("code"))
                if prev_c or next_c:
                    n_code = sum(1 for ln in blk.lines
                                 if _code_hint_line(ln.text))
                    if n_code >= 1:
                        blk.meta["code"] = True
                        changed = True

    # 正文块内与公式区重叠的短碎片行 → 拨给邻近公式块（裁图保真）
    def _near_formula(blk):
        best = None
        for fb in out_blocks:
            if fb is blk or fb.kind != BlockKind.FORMULA:
                continue
            gap = max(blk.bbox[1] - fb.bbox[3], fb.bbox[1] - blk.bbox[3])
            x_ov = min(fb.bbox[2], blk.bbox[2]) - max(fb.bbox[0], blk.bbox[0])
            if x_ov > 0 and gap < base_h * 1.3:
                if best is None or gap < best[1]:
                    best = (fb, gap)
        return best[0] if best else None

    for blk in list(out_blocks):
        if blk.kind != BlockKind.TEXT or blk.meta.get("code"):
            continue
        if _near_formula(blk) is None:
            continue
        moved = [ln for ln in blk.lines if len(ln.text.strip()) <= 12]
        if moved and len(moved) < len(blk.lines):
            fb = _near_formula(blk)
            fb.lines.extend(moved)
            fb.lines.sort(key=lambda l: l.bbox[1])
            fb.bbox = _bbox_of_lines(fb.lines)
            blk.lines = [ln for ln in blk.lines if ln not in moved]
            if blk.lines:
                blk.bbox = _bbox_of_lines(blk.lines)
            else:
                out_blocks.remove(blk)

    # 黑体短块（如 "解"、"证"）
    for blk in out_blocks:
        if (blk.kind == BlockKind.TEXT and len(blk.lines) == 1
                and 0 < len(blk.text.strip()) <= 4
                and ink_bold_ratio(bw, zoom, blk.bbox, base_ink) > 1.35):
            blk.kind = BlockKind.HEADING

    # 跳过的页眉页脚块
    if skip_lines:
        out_blocks.append(Block(BlockKind.SKIP, _bbox_of_lines(skip_lines), skip_lines))

    out_blocks.sort(key=lambda b: b.bbox[1])

    # 正文区边距（排除页眉页脚/页码）
    body = [b for b in out_blocks if b.kind != BlockKind.SKIP]
    if body:
        bx0 = min(b.bbox[0] for b in body)
        by0 = min(b.bbox[1] for b in body)
        bx1 = max(b.bbox[2] for b in body)
        by1 = max(b.bbox[3] for b in body)
        margins = (max(12, by0 - 4), max(12, pr.height - by1 - 4),
                   max(12, bx0 - 4), max(12, pr.width - bx1 - 4))
    else:
        margins = (72, 72, 72, 72)

    # 页面高度预算：Word 渲染的图段/表格行/段落间距都比原书排版略高，
    # 累积导致溢出分页。估算渲染高度并对行距做整体压缩（下限 0.85）。
    spacing_scale = 1.0
    if body:
        avail_h = pr.height - margins[0] - margins[1]
        est_h = 0.0
        for b in body:
            bh = b.bbox[3] - b.bbox[1]
            if b.kind in (BlockKind.FORMULA, BlockKind.IMAGE):
                est_h += bh + 5          # 图片行 leading + 段距
            elif b.kind == BlockKind.TABLE:
                est_h += bh * 1.22       # 单元格 padding 累积
            else:
                est_h += bh + len(b.lines) * 1.2
        if est_h > 0:
            spacing_scale = min(1.0, max(0.85, avail_h * 0.97 / est_h))

    return {"blocks": out_blocks, "margins": margins, "base_h": base_h,
            "source": source, "page": pr, "img": img, "angle": angle,
            "words": words, "base_ink": base_ink,
            "spacing_scale": spacing_scale}


def _merge_overlapping_blocks(blocks: list[Block], base_h: float) -> list[Block]:
    """合并 bbox 重叠大的块（OCR 行框交叉造成的假分块）。"""
    def overlap(a, b):
        ix0, iy0 = max(a[0], b[0]), max(a[1], b[1])
        ix1, iy1 = min(a[2], b[2]), min(a[3], b[3])
        if ix1 <= ix0 or iy1 <= iy0:
            return 0.0
        inter = (ix1 - ix0) * (iy1 - iy0)
        union = (a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - inter
        return inter / union

    changed = True
    while changed:
        changed = False
        blocks.sort(key=lambda b: b.bbox[1])
        for i in range(len(blocks)):
            done = False
            for j in range(i + 1, len(blocks)):
                if overlap(blocks[i].bbox, blocks[j].bbox) > 0.45:
                    a, b = blocks[i], blocks[j]
                    a.lines.extend(b.lines)
                    a.lines.sort(key=lambda l: l.bbox[1])
                    a.bbox = _bbox_of_lines(a.lines)
                    if b.meta.get("code"):
                        a.meta["code"] = True
                    blocks.pop(j)
                    changed = done = True
                    break
            if done:
                break
    return blocks


def _body_ink(bw, zoom, lines, base_h) -> float:
    """正文行平均墨迹密度（粗体检测基准）。"""
    vals = []
    for ln in lines[:60]:
        if len(ln.text) >= 6:
            v, _ = ink_stats(bw, zoom, ln.bbox)
            vals.append(v)
    return float(np.median(vals)) if vals else 0.08


def ink_bold_ratio(bw, zoom, bbox, base_ink) -> float:
    ink, _ = ink_stats(bw, zoom, bbox)
    return ink / base_ink if base_ink > 0 else 1.0


def _inside(a, b, ratio=0.6) -> bool:
    """bbox a 是否大部分在 b 内。"""
    ix0, iy0 = max(a[0], b[0]), max(a[1], b[1])
    ix1, iy1 = min(a[2], b[2]), min(a[3], b[3])
    if ix1 <= ix0 or iy1 <= iy0:
        return False
    inter = (ix1 - ix0) * (iy1 - iy0)
    area = max(1e-6, (a[2] - a[0]) * (a[3] - a[1]))
    return inter / area >= ratio


def _overlap_ratio(a, b) -> float:
    ix0, iy0 = max(a[0], b[0]), max(a[1], b[1])
    ix1, iy1 = min(a[2], b[2]), min(a[3], b[3])
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    inter = (ix1 - ix0) * (iy1 - iy0)
    area = max(1e-6, (a[2] - a[0]) * (a[3] - a[1]))
    return inter / area


def _bbox_of_lines(lines):
    return [min(l.bbox[0] for l in lines), min(l.bbox[1] for l in lines),
            max(l.bbox[2] for l in lines), max(l.bbox[3] for l in lines)]


# ---------------------------------------------------------------- Word 重建

def _join_lines(prev: str, nxt: str) -> str:
    if prev and nxt:
        a, b = prev[-1], nxt[0]
        if a.isascii() and a.isalnum() and b.isascii() and b.isalnum():
            return prev + " " + nxt
    return prev + nxt


def expand_to_ink(img: np.ndarray, bbox_pt: list, zoom: float,
                  max_h: float = 10.0, max_v: float = 2.0) -> list:
    """公式 bbox 向四周扩展到墨迹边界（限幅），补回 OCR 漏识别的符号。"""
    H, W = img.shape
    dark = img < 185
    x0, y0, x1, y1 = [int(v * zoom) for v in bbox_pt]
    mx, my = int(max_h * zoom), int(max_v * zoom)
    xa, xb = max(0, x0 - mx), min(W, x1 + mx)
    ya, yb = max(0, y0 - my), min(H, y1 + my)
    region = dark[ya:yb, xa:xb]
    ys, xs = np.flatnonzero(region.any(axis=1)), np.flatnonzero(region.any(axis=0))
    if len(ys) == 0 or len(xs) == 0:
        return bbox_pt
    return [max(xa + int(xs[0]), 0) / zoom, max(ya + int(ys[0]), 0) / zoom,
            min(xa + int(xs[-1]) + 1, W) / zoom, min(ya + int(ys[-1]) + 1, H) / zoom]


def _trim_partial_lines(roi: np.ndarray) -> np.ndarray:
    """裁掉上下边缘的残行碎片（相邻正文的下半截被行 bbox 带入）。"""
    H, W = roi.shape
    if H < 12:
        return roi
    inked = (roi < 185).sum(axis=1) > 2
    rows = np.flatnonzero(inked)
    if len(rows) == 0:
        return roi
    # 墨迹行分组（间隔 > 4px 切组）
    groups = [[rows[0]]]
    for r in rows[1:]:
        if r - groups[-1][-1] <= 4:
            groups[-1].append(r)
        else:
            groups.append([r])
    # 残行判据：高度明显小于组高中位数（被 bbox 切断的半行）
    heights = [g[-1] - g[0] + 1 for g in groups]
    med = sorted(heights)[len(heights) // 2]
    top, bot = 0, H
    if len(groups) >= 2 and heights[0] < 0.55 * med:
        top = groups[1][0]
    if len(groups) >= 2 and heights[-1] < 0.55 * med:
        bot = groups[-2][-1] + 1
    return roi[top:bot] if bot > top else roi


def _crop_from(img: np.ndarray, bbox_pt: list, zoom: float, pad_pt: float = 1.0,
               trim: bool = False,
               exclude_pt: list[list] | None = None) -> bytes:
    """从（已校正的）页面灰度图按 pt bbox 裁图，白底化。

    exclude_pt: 需要涂白的行框列表（pt）——用于抹除裁图区域内
    属于其他行/块的墨迹，防止裁图与正文文本重复显示。
    """
    x0, y0, x1, y1 = [int(v * zoom) for v in bbox_pt]
    pad = int(pad_pt * zoom)
    x0, y0 = max(0, x0 - pad), max(0, y0 - pad)
    x1 = min(img.shape[1], x1 + pad)
    y1 = min(img.shape[0], y1 + pad)
    if x1 <= x0 or y1 <= y0:
        return b""
    roi = img[y0:y1, x0:x1].copy()
    if exclude_pt:
        h, w = roi.shape
        for eb in exclude_pt:
            ex0 = max(0, int(eb[0] * zoom) - x0 - 1)
            ey0 = max(0, int(eb[1] * zoom) - y0 - 1)
            ex1 = min(w, int(eb[2] * zoom) - x0 + 1)
            ey1 = min(h, int(eb[3] * zoom) - y0 + 1)
            if ex1 > ex0 and ey1 > ey0:
                roi[ey0:ey1, ex0:ex1] = 255
    roi = np.where(roi > 190, 255, roi).astype(np.uint8)
    if trim:
        roi = _trim_partial_lines(roi)
    ok, buf = cv2.imencode(".png", roi)
    return buf.tobytes() if ok else b""


def _set_font(run, cn_font: str, size_pt: float, bold=False):
    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)


def _para_spacing(p, before_pt=0, after_pt=0, line_mult=None, exact_pt=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    if exact_pt:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(exact_pt)
    elif line_mult:
        pf.line_spacing = line_mult


def build_table(document, blk: Block, img, zoom: float, base_h: float,
                words: list[W]):
    from .segment import rapidocr
    row_ys, col_xs = table_grid(img, zoom, blk.bbox)
    if len(row_ys) < 2 or len(col_xs) < 2:
        # 网格提取失败 → 退化为裁图
        return False
    rows, cols = len(row_ys) - 1, len(col_xs) - 1
    table = document.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    _table_borders(table)

    # 列宽
    tbl_pr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)
    for j in range(cols):
        w_pt = col_xs[j + 1] - col_xs[j]
        for cell in table.columns[j].cells:
            cell.width = Emu(int(w_pt * 12700))

    # RapidOCR 对表格区域重新识别（印刷体准确率远高于扫描仪内置 OCR 层）
    x0, y0 = int(blk.bbox[0] * zoom), int(blk.bbox[1] * zoom)
    x1, y1 = int(blk.bbox[2] * zoom), int(blk.bbox[3] * zoom)
    roi = img[max(0, y0):y1, max(0, x0):x1]
    cell_words: dict[tuple, list[W]] = {}
    if roi.size:
        # 宽扁表格会令检测模型失效，上下补白边
        pad_v = max(30, int(roi.shape[0] * 0.5))
        padded = cv2.copyMakeBorder(roi, pad_v, pad_v, 20, 20,
                                    cv2.BORDER_CONSTANT, value=255)
        result, _ = rapidocr()(padded)
        if result:
            inv = 1.0 / zoom
            for pts, text, conf in result:
                t = str(text).strip()
                if not t:
                    continue
                xs = [q[0] for q in pts]
                ys = [q[1] for q in pts]
                gx0 = (x0 - 20 + min(xs)) * inv
                gy0 = (y0 - pad_v + min(ys)) * inv
                gx1 = (x0 - 20 + max(xs)) * inv
                gy1 = (y0 - pad_v + max(ys)) * inv
                ri = _span_index((gy0 + gy1) / 2, row_ys)
                ci = _span_index((gx0 + gx1) / 2, col_xs)
                if 0 <= ri < rows and 0 <= ci < cols:
                    cell_words.setdefault((ri, ci), []).append(
                        W(gx0, gy0, gx1, gy1, t, (gy1 - gy0) * 0.85, float(conf)))

    # RapidOCR 漏识别的空单元格 → 用 PDF 文字层词兜底
    for (i, j), ws in list(cell_words.items()):
        pass
    for i in range(rows):
        for j in range(cols):
            if (i, j) in cell_words and cell_words[(i, j)]:
                continue
            cx0, cy0 = col_xs[j], row_ys[i]
            cx1, cy1 = col_xs[j + 1], row_ys[i + 1]
            fallback = [wd for wd in words
                        if cx0 <= (wd.x0 + wd.x1) / 2 < cx1
                        and cy0 <= (wd.y0 + wd.y1) / 2 < cy1]
            if fallback:
                cell_words[(i, j)] = sorted(fallback,
                                            key=lambda w: (w.y0, w.x0))

    for i in range(rows):
        for j in range(cols):
            ws = cell_words.get((i, j), [])
            cell = table.cell(i, j)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _para_spacing(p, 0, 0,
                          exact_pt=max(12.0, (row_ys[i + 1] - row_ys[i]) * 0.85))
            if ws:
                ws.sort(key=lambda w: (round(w.y0), w.x0))
                text = ""
                for w in ws:
                    text = _join_lines(text, w.text)
                run = p.add_run(sanitize_xml_text(text))
                sz = float(np.median([w.h for w in ws])) * 0.9
                _set_font(run, "宋体", max(8, min(sz, base_h * 0.95)))
    return True


def _span_index(v, edges) -> int:
    for i in range(len(edges) - 1):
        if edges[i] <= v < edges[i + 1]:
            return i
    # 边界容错：落在端点附近的取最近单元格（OCR 框中心常压线）
    if edges:
        best_i, best_d = -1, 1e9
        for i, e in enumerate(edges):
            d = abs(v - e)
            if d < best_d:
                best_d, best_i = d, i
        if best_d < 5:
            idx = best_i - 1 if v < edges[best_i] else best_i
            return max(0, min(idx, len(edges) - 2))
    return -1


def _table_borders(table):
    tbl_pr = table._tbl.tblPr
    borders = parse_xml(
        '<w:tblBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        + "".join(
            f'<w:{edge} w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            for edge in ("top", "left", "bottom", "right", "insideH", "insideV"))
        + "</w:tblBorders>")
    tbl_pr.append(borders)


def build_page(document: Document, page: fitz.Page, analysis: dict,
               dpi: int = 300, first: bool = False,
               formula_omml: bool = False):
    pr = analysis["page"]
    margins = analysis["margins"]
    base_h = analysis["base_h"]
    spacing_scale = analysis.get("spacing_scale", 1.0)
    sec = document.sections[0] if first else document.add_section(WD_SECTION.NEW_PAGE)
    set_page(sec, pr.width, pr.height, margins)

    if not analysis["blocks"]:
        document.add_paragraph()
        return

    zoom = dpi / 72.0
    img = analysis.get("img")
    if img is None:                       # 空页兜底
        img = render_gray(page, dpi)
    bw_cache = {}

    def get_bw():
        if "bw" not in bw_cache:
            bw_cache["bw"] = binarize(img)
        return bw_cache["bw"]

    def crop(bbox):
        return _crop_from(img, bbox, zoom)

    # 裁图排除区：所有其他 TEXT/SKIP 块的行框（涂白，防止裁图
    # 罩住相邻正文墨迹造成"文本+图"重复显示）
    def excl_for(bbox):
        excl = []
        for ob in all_blocks:
            if ob is blk or ob.kind not in (BlockKind.TEXT, BlockKind.SKIP,
                                            BlockKind.HEADING):
                continue
            for ln in ob.lines:
                if _overlap_ratio(ln.bbox, bbox) > 0:
                    excl.append(ln.bbox)
        return excl

    body_size = _body_font_size(analysis, base_h)
    all_blocks = analysis["blocks"]

    for bi, blk in enumerate(all_blocks):
        if blk.kind == BlockKind.SKIP:
            continue

        if blk.kind == BlockKind.TABLE:
            ok = build_table(document, blk, img, zoom, base_h,
                             analysis.get("words", []))
            nxt = all_blocks[bi + 1] if bi + 1 < len(all_blocks) else None
            if nxt and nxt.kind == BlockKind.TABLE:
                document.add_paragraph()   # 防止 Word 合并相邻表格
            if not ok:
                p = document.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _para_spacing(p, 0, 0, line_mult=1.0)
                run = p.add_run()
                data = crop(blk.bbox)
                if data:
                    run.add_picture(io.BytesIO(data),
                                    width=Emu(int((blk.bbox[2] - blk.bbox[0]) * 12700)))
            continue

        if blk.kind == BlockKind.IMAGE:
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _para_spacing(p, 0, 0, line_mult=1.0)
            run = p.add_run()
            data = _crop_from(img, blk.bbox, zoom, trim=True,
                              exclude_pt=excl_for(blk.bbox))
            if data:
                run.add_picture(io.BytesIO(data),
                                width=Emu(int((blk.bbox[2] - blk.bbox[0]) * 12700)))
            continue

        if blk.kind == BlockKind.FORMULA:
            # 公式 OCR → 原生 OMML 公式（路线图项）；失败回退高清裁图
            omml_root = None
            if formula_omml:
                fb = expand_to_ink(img, blk.bbox, zoom)
                data = _crop_from(img, fb, zoom, trim=True,
                                  exclude_pt=excl_for(fb))
                if data:
                    try:
                        from PIL import Image as _PILImage
                        from .omml import (append_omml_to_paragraph,
                                           image_to_latex, latex_plausible,
                                           latex_to_omml_xml)
                        im = _PILImage.open(io.BytesIO(data))
                        latex = image_to_latex(im)
                        if not latex_plausible(latex):
                            raise ValueError("公式识别结果不合理")
                        omml_root = latex_to_omml_xml(latex)
                    except Exception:
                        omml_root = None
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _para_spacing(p, 0, 0, line_mult=1.0)
            if omml_root is not None:
                append_omml_to_paragraph(p, omml_root, body_size)
                continue
            fb = expand_to_ink(img, blk.bbox, zoom)
            data = _crop_from(img, fb, zoom, trim=True,
                              exclude_pt=excl_for(fb))
            run = p.add_run()
            if data:
                run.add_picture(io.BytesIO(data),
                                width=Emu(int((fb[2] - fb[0]) * 12700)))
            continue

        # 代码块：整块高清裁图（OCR 错字会破坏代码语义，如 8olve/solve）
        if blk.meta.get("code"):
            p = document.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            _para_spacing(p, 0, 0, line_mult=1.0)
            run = p.add_run()
            data = _crop_from(img, blk.bbox, zoom, trim=False,
                              exclude_pt=excl_for(blk.bbox))
            if data:
                run.add_picture(io.BytesIO(data),
                                width=Emu(int((blk.bbox[2] - blk.bbox[0]) * 12700)))
            continue

        # TEXT / HEADING：拆段输出
        for para_lines in _split_paragraphs(blk, base_h):
            p = document.add_paragraph()
            indent = para_lines[0].bbox[0] - blk.bbox[0]
            sz = body_size
            if blk.kind == BlockKind.HEADING:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                sz = body_size * 1.05
            else:
                p.alignment = _guess_align(para_lines, blk, base_h)
            if indent > base_h * 0.8 and blk.kind == BlockKind.TEXT:
                p.paragraph_format.first_line_indent = Pt(indent)
            # 行距：段内行 pitch 中位数，按页面高度预算整体压缩
            line_pt = _line_spacing(para_lines, base_h) * spacing_scale
            _para_spacing(p, 0, 0, exact_pt=line_pt)

            bold = blk.kind == BlockKind.HEADING
            if not bold:
                bold = ink_bold_ratio(get_bw(), zoom, _bbox_of_lines(para_lines),
                                      analysis.get("base_ink", 0.08)) > 1.4
            text = ""
            for ln in para_lines:
                text = _join_lines(text, ln.text)
            run = p.add_run(sanitize_xml_text(text))
            _set_font(run, "宋体", sz, bold=bold)


def _body_font_size(analysis, base_h) -> float:
    """正文字号(pt)：行高中位数换算。"""
    hs = [ln.h for blk in analysis["blocks"] if blk.kind == BlockKind.TEXT
          for ln in blk.lines if len(ln.text) > 8]
    med = float(np.median(hs)) if hs else base_h
    return max(8.0, round(med * 0.80 * 2) / 2)


def _split_paragraphs(blk: Block, base_h: float) -> list[list[Line]]:
    """块内按首行缩进/居中特征切分段落。"""
    if blk.kind == BlockKind.HEADING:
        return [blk.lines]
    main_x0 = _dominant_x0(blk)
    main_x1 = float(np.median([round(l.bbox[2]) for l in blk.lines]))
    blk_w = blk.bbox[2] - blk.bbox[0]
    paras, cur = [], []
    for ln in blk.lines:
        indent = ln.bbox[0] - main_x0
        ln_w = ln.bbox[2] - ln.bbox[0]
        # 居中短行（表标题等）独立成段
        is_centered = (ln_w < blk_w * 0.7 and blk_w > base_h * 6
                       and ln.bbox[0] - blk.bbox[0] > base_h
                       and blk.bbox[2] - ln.bbox[2] > base_h)
        if cur and (indent > base_h * 0.8 or is_centered):
            paras.append(cur)
            cur = [ln]
        else:
            cur.append(ln)
    if cur:
        paras.append(cur)
    return paras


def _dominant_x0(blk: Block) -> float:
    xs = sorted(round(l.bbox[0]) for l in blk.lines)
    if not xs:
        return 0.0
    # 众数
    best, best_n = xs[0], 0
    for x in xs:
        n = sum(1 for y in xs if abs(y - x) <= 3)
        if n > best_n:
            best, best_n = x, n
    return float(best)


def _guess_align(lines: list[Line], blk: Block, base_h):
    ln = lines[-1]
    w_blk = blk.bbox[2] - blk.bbox[0]
    ln_w = ln.bbox[2] - ln.bbox[0]
    if len(lines) == 1:
        lm = ln.bbox[0] - blk.bbox[0]
        rm = blk.bbox[2] - ln.bbox[2]
        if w_blk > base_h * 4 and abs(lm - rm) < base_h * 1.2 and lm > base_h:
            return WD_ALIGN_PARAGRAPH.CENTER
    # 短行不两端对齐（避免字距拉伸）
    if ln_w < w_blk * 0.85:
        return WD_ALIGN_PARAGRAPH.LEFT
    return WD_ALIGN_PARAGRAPH.JUSTIFY


def _line_spacing(lines: list[Line], base_h: float) -> float:
    if len(lines) < 2:
        return base_h * 1.25
    gaps = [lines[i + 1].bbox[1] - lines[i].bbox[1] for i in range(len(lines) - 1)]
    g = float(np.median([x for x in gaps if x > 0]) if gaps else base_h * 1.25)
    return max(10.0, g)


# ---------------------------------------------------------------- 主入口

def _build_document(doc: fitz.Document, pages: range, dpi: int, reocr: bool,
                    scale: float, progress=None,
                    analysis_cache: dict | None = None,
                    formula_omml: bool = False) -> Document:
    """按指定行距压缩系数构建完整 Document（一遍）。"""
    document = Document()
    st = document.styles["Normal"]
    st.font.name = "Times New Roman"
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    st.font.size = Pt(12)

    for k, i in enumerate(pages):
        page = doc[i]
        info = classify_page(page)
        if info.kind == PageKind.DIGITAL:
            from .digital import build_page_digital
            build_page_digital(document, page, first=(k == 0))
        else:
            cache_key = ("scan", i)
            if analysis_cache is not None and cache_key in analysis_cache:
                analysis = analysis_cache[cache_key]
            else:
                analysis = analyze_page(page, dpi=dpi, reocr=reocr)
                if analysis_cache is not None:
                    analysis_cache[cache_key] = analysis
            analysis["spacing_scale"] = min(
                analysis.get("spacing_scale", 1.0), scale)
            build_page(document, page, analysis, dpi=dpi, first=(k == 0),
                       formula_omml=formula_omml)
        if progress:
            progress(i, len(pages))
    return document


def count_docx_pages(docx_path: str) -> int:
    """用 Word COM 实测 docx 页数（不可用时返回 -1）。"""
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            d = word.Documents.Open(str(Path(docx_path).resolve()), ReadOnly=True)
            n = d.ComputeStatistics(2)          # wdStatisticPages
            d.Close(False)
            return int(n)
        finally:
            word.Quit()
    except Exception:
        return -1


def convert(doc: fitz.Document, out_path: str, pages: range | None = None,
            dpi: int = 300, progress=None, reocr: bool = False,
            fit_target: int | None = None, fit_rounds: int = 2,
            formula_omml: bool = False):
    """editable 转换主入口。

    fit_target: 期望页数（通常=原书页数）。给出时启用二遍排版：
    每轮构建后用 Word 实测页数，超页则按实测比例回填行距压缩系数，
    再用缓存的分析结果快速重建（不重跑 OCR）。
    """
    rng = pages if pages is not None else range(doc.page_count)
    scale = 1.0
    document = None
    analysis_cache: dict = {}
    attempts = (1 + max(0, fit_rounds)) if fit_target else 1
    for attempt in range(attempts):
        document = _build_document(doc, rng, dpi, reocr, scale,
                                   progress, analysis_cache,
                                   formula_omml=formula_omml)
        document.save(out_path)
        if not fit_target:
            break
        actual = count_docx_pages(out_path)
        if actual < 0 or actual <= fit_target:
            if actual > 0:
                print(f"[book2docx] 二遍排版：实测 {actual} 页 ≤ 目标 "
                      f"{fit_target} 页，无需压缩")
            break
        new_scale = scale * max(0.88, (fit_target / actual) ** 0.9)
        if new_scale < 0.72 or abs(new_scale - scale) < 0.01:
            print(f"[book2docx] 二遍排版：实测 {actual} 页，压缩到下限，停止")
            break
        print(f"[book2docx] 二遍排版：实测 {actual} 页 > 目标 {fit_target} 页，"
              f"行距系数 {scale:.2f} → {new_scale:.2f}，重建…")
        scale = new_scale
    return out_path
