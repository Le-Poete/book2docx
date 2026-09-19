# -*- coding: utf-8 -*-
"""扫描页版面分割：词块 → 行 → 块，表格/插图/公式区域检测与分类。"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

import cv2
import fitz
import numpy as np

from .common import math_density


class BlockKind(Enum):
    TEXT = "text"
    HEADING = "heading"
    FORMULA = "formula"      # 独立公式行/块 → 高清裁图
    TABLE = "table"          # 有网格线 → 重建 Word 表格
    IMAGE = "image"          # 插图/大图形 → 裁图
    SKIP = "skip"            # 页码/页眉等


@dataclass
class W:
    """一个 OCR 词块。坐标均为 pt（PDF 坐标系，y 向下）。"""
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    size: float = 0.0        # 字号(pt)，来自 PDF 层或估计
    conf: float = 1.0

    @property
    def cx(self):
        return (self.x0 + self.x1) / 2

    @property
    def cy(self):
        return (self.y0 + self.y1) / 2

    @property
    def h(self):
        return self.y1 - self.y0

    @property
    def w(self):
        return self.x1 - self.x0


@dataclass
class Line:
    bbox: list                # [x0,y0,x1,y1]
    words: list[W] = field(default_factory=list)
    ink_ratio: float = 0.0    # 墨迹密度（图像校验用）
    v_span: float = 0.0       # 词块垂直跨度（上下标检测）

    @property
    def text(self):
        return "".join(w.text for w in sorted(self.words, key=lambda t: t.x0))

    @property
    def h(self):
        return self.bbox[3] - self.bbox[1]


@dataclass
class Block:
    kind: BlockKind
    bbox: list                # pt
    lines: list[Line] = field(default_factory=list)
    meta: dict = field(default_factory=dict)

    @property
    def text(self):
        return "\n".join(ln.text for ln in self.lines)


# ---------------------------------------------------------------- 词块提取

def extract_words_pt(page: fitz.Page) -> list[W]:
    """PDF 自带 OCR 文字层 → 词块（pt）。"""
    words = []
    for x0, y0, x1, y1, text, *_ in page.get_text("words"):
        text = text.strip()
        if not text:
            continue
        h = y1 - y0
        words.append(W(x0, y0, x1, y1, text, size=h))
    return words


def dedup_lines(lines: list[Line], base_h: float) -> list[Line]:
    """去掉 OCR 层的重复行（保守判据：垂直高度重叠 > 60%、水平重叠 > 50%，
    且文本完全相同或短行为长行的子串。不使用模糊相似度，避免误删公式行。）"""
    out: list[Line] = []
    for ln in lines:
        dup = False
        for prev in out:
            iy = min(ln.bbox[3], prev.bbox[3]) - max(ln.bbox[1], prev.bbox[1])
            ix = min(ln.bbox[2], prev.bbox[2]) - max(ln.bbox[0], prev.bbox[0])
            if iy <= 0:
                continue
            h_min = min(ln.bbox[3] - ln.bbox[1], prev.bbox[3] - prev.bbox[1])
            w_min = min(ln.bbox[2] - ln.bbox[0], prev.bbox[2] - prev.bbox[0])
            if iy / h_min < 0.6 or ix / w_min < 0.5:
                continue
            ta, tb = ln.text.strip(), prev.text.strip()
            if not ta or not tb:
                continue
            if ta == tb:
                dup = True
                break
            short, long = (ta, tb) if len(ta) < len(tb) else (tb, ta)
            if len(short) >= 10 and short in long:
                dup = True
                break
            if len(short) >= 12:
                import difflib
                if difflib.SequenceMatcher(None, ta, tb).ratio() >= 0.8:
                    dup = True
                    break
        if not dup:
            out.append(ln)
    return out


def extract_words_ocr(page: fitz.Page, dpi: int = 200) -> list[W]:
    """RapidOCR 识别（无文字层页面）→ 词块（pt）。"""
    from rapidocr_onnxruntime import RapidOCR
    ocr = RapidOCR()
    zoom = dpi / 72
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csGRAY)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
    result, _ = ocr(img)
    words = []
    if result:
        for pts, text, conf in result:
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            words.append(W(min(xs) / zoom, min(ys) / zoom, max(xs) / zoom, max(ys) / zoom,
                           text, size=(max(ys) - min(ys)) / zoom * 0.85, conf=float(conf)))
    return words


def get_words(page: fitz.Page, min_quality: float = 0.5) -> tuple[list[W], str]:
    """统一入口：优先 PDF 层，质量差或缺失时跑 OCR。返回 (words, source)。"""
    ws = extract_words_pt(page)
    if ws:
        return ws, "pdf_layer"
    return extract_words_ocr(page), "rapidocr"


# ---------------------------------------------------------------- 聚类

def cluster_lines(words: list[W], ytol_ratio: float = 0.55) -> list[Line]:
    """按垂直中线聚类成行。上下标词（与已有行水平重叠、y 偏移大）并入最近行。"""
    if not words:
        return []
    med_h = float(np.median([w.h for w in words]))
    ws = sorted(words, key=lambda t: (t.cy, t.x0))
    lines: list[Line] = []

    for w in ws:
        placed = False
        # 尝试并入现有行：中线接近，或垂直重叠足够
        for ln in reversed(lines[-6:]):
            lc_y = (ln.bbox[1] + ln.bbox[3]) / 2
            overlap = min(ln.bbox[2], w.x1) - max(ln.bbox[0], w.x0)
            if abs(w.cy - lc_y) < med_h * ytol_ratio and overlap > -med_h * 2:
                ln.words.append(w)
                _reline(ln)
                placed = True
                break
        if not placed:
            lines.append(Line(bbox=[w.x0, w.y0, w.x1, w.y1], words=[w]))
    lines.sort(key=lambda l: l.bbox[1])
    for ln in lines:
        ln.words.sort(key=lambda t: t.x0)
        ln.v_span = max(w.y1 for w in ln.words) - min(w.y0 for w in ln.words)
    return lines


def _reline(ln: Line):
    xs0 = [w.x0 for w in ln.words]
    ys0 = [w.y0 for w in ln.words]
    xs1 = [w.x1 for w in ln.words]
    ys1 = [w.y1 for w in ln.words]
    ln.bbox = [min(xs0), min(ys0), max(xs1), max(ys1)]


def cluster_blocks(lines: list[Line], base_h: float) -> list[Block]:
    """行间距聚类成块：垂直间隔 < 0.75*base_h 视为同块。"""
    blocks: list[Block] = []
    cur: list[Line] = []
    for ln in lines:
        if cur:
            gap = ln.bbox[1] - cur[-1].bbox[3]
            if gap > base_h * 0.95:
                blocks.append(Block(BlockKind.TEXT, _bbox_of(cur), cur))
                cur = [ln]
            else:
                cur.append(ln)
        else:
            cur = [ln]
    if cur:
        blocks.append(Block(BlockKind.TEXT, _bbox_of(cur), cur))
    return blocks


def split_lines_by_gap(lines: list[Line], base_h: float,
                       factor: float = 1.15) -> list[Line]:
    """行内按大水平空隙切分：公式与正文同行粘连时拆开分别分类。"""
    out: list[Line] = []
    for ln in lines:
        ws = sorted(ln.words, key=lambda t: t.x0)
        segs: list[list[W]] = [[ws[0]]]
        for w in ws[1:]:
            if w.x0 - segs[-1][-1].x1 > factor * base_h:
                segs.append([w])
            else:
                segs[-1].append(w)
        if len(segs) == 1:
            out.append(ln)
            continue
        for sg in segs:
            xs0 = min(w.x0 for w in sg)
            ys0 = min(w.y0 for w in sg)
            xs1 = max(w.x1 for w in sg)
            ys1 = max(w.y1 for w in sg)
            out.append(Line(bbox=[xs0, ys0, xs1, ys1], words=sg,
                            v_span=ys1 - ys0))
    out.sort(key=lambda l: l.bbox[1])
    return out


_OCR = None


def dedup_lines(lines: list[Line], base_h: float) -> list[Line]:
    """去掉 OCR 层的重复行（行框重叠且文字互为子串/高度相似）。"""
    out: list[Line] = []

    def _similar(ta: str, tb: str) -> bool:
        if not ta or not tb:
            return False
        if ta == tb:
            return True
        short, long = (ta, tb) if len(ta) < len(tb) else (tb, ta)
        if len(short) >= 6 and short in long:
            return True
        import difflib
        return difflib.SequenceMatcher(None, ta, tb).ratio() > 0.85

    for ln in lines:
        dup = False
        for prev in out:
            ix = min(ln.bbox[2], prev.bbox[2]) - max(ln.bbox[0], prev.bbox[0])
            iy = min(ln.bbox[3], prev.bbox[3]) - max(ln.bbox[1], prev.bbox[1])
            if ix <= 0 or iy <= 0:
                continue
            area_a = (ln.bbox[2] - ln.bbox[0]) * (ln.bbox[3] - ln.bbox[1])
            area_b = (prev.bbox[2] - prev.bbox[0]) * (prev.bbox[3] - prev.bbox[1])
            if ix * iy / max(1.0, min(area_a, area_b)) < 0.45:
                continue
            if _similar(ln.text.strip(), prev.text.strip()):
                dup = True
                break
        if not dup:
            out.append(ln)
    return out


def rapidocr():
    global _OCR
    if _OCR is None:
        from rapidocr_onnxruntime import RapidOCR
        _OCR = RapidOCR()
    return _OCR


def _bbox_of(lines: list[Line]) -> list:
    return [
        min(l.bbox[0] for l in lines), min(l.bbox[1] for l in lines),
        max(l.bbox[2] for l in lines), max(l.bbox[3] for l in lines),
    ]


# ---------------------------------------------------------------- 图像分析

def render_gray(page: fitz.Page, dpi: int) -> np.ndarray:
    zoom = dpi / 72
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), colorspace=fitz.csGRAY)
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
    return img


def estimate_skew(img_gray: np.ndarray) -> float:
    """投影剖面法估计倾斜角（度）：在 ±3° 内搜索行投影最尖锐的角度。

    文本页在正确校正角度下，行投影呈现最尖锐的峰谷结构（差分能量最大）。
    """
    small = cv2.resize(img_gray, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA)
    dark = (small < 170).astype(np.float32)
    h, w = dark.shape

    def sharpness(a: float) -> float:
        M = cv2.getRotationMatrix2D((w / 2, h / 2), a, 1.0)
        rot = cv2.warpAffine(dark, M, (w, h), flags=cv2.INTER_NEAREST,
                             borderMode=cv2.BORDER_CONSTANT, borderValue=0)
        prof = rot.sum(axis=1)
        return float(np.square(np.diff(prof)).sum())

    best_a, best_v = 0.0, -1.0
    for a in np.arange(-3.0, 3.01, 0.1):
        v = sharpness(float(a))
        if v > best_v:
            best_v, best_a = v, float(a)
    for a in np.arange(best_a - 0.1, best_a + 0.11, 0.02):
        v = sharpness(float(a))
        if v > best_v:
            best_v, best_a = v, float(a)
    return round(best_a, 2)


def deskew(img_gray: np.ndarray, angle: float):
    """旋转校正。返回 (校正图, 变换矩阵 M)：原图 px → 校正图 px。"""
    h, w = img_gray.shape
    M = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    cos, sin = abs(M[0, 0]), abs(M[0, 1])
    nw, nh = int(h * sin + w * cos) + 2, int(h * cos + w * sin) + 2
    M[0, 2] += nw / 2 - w / 2
    M[1, 2] += nh / 2 - h / 2
    out = cv2.warpAffine(img_gray, M, (nw, nh), flags=cv2.INTER_CUBIC,
                         borderMode=cv2.BORDER_CONSTANT, borderValue=255)
    return out, M


def transform_pt_bbox(bbox_pt: list, M: np.ndarray, zoom: float) -> list:
    """把原图 pt bbox 经 deskew 矩阵映射为校正图 pt bbox。"""
    pts = np.array([
        [bbox_pt[0] * zoom, bbox_pt[1] * zoom, 1.0],
        [bbox_pt[2] * zoom, bbox_pt[1] * zoom, 1.0],
        [bbox_pt[0] * zoom, bbox_pt[3] * zoom, 1.0],
        [bbox_pt[2] * zoom, bbox_pt[3] * zoom, 1.0],
    ])
    tp = (M @ pts.T).T
    return [float(tp[:, 0].min() / zoom), float(tp[:, 1].min() / zoom),
            float(tp[:, 0].max() / zoom), float(tp[:, 1].max() / zoom)]


def binarize(img: np.ndarray) -> np.ndarray:
    """自适应二值化，墨迹=255。blockSize 需小于字符尺寸。"""
    bw = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_MEAN_C,
                               cv2.THRESH_BINARY_INV, 41, 12)
    # 去除孤立噪点
    bw = cv2.medianBlur(bw, 3)
    return bw


def ink_stats(bw: np.ndarray, zoom: float, bbox_pt: list) -> tuple[float, float]:
    """区域墨迹占比和最大连通行高。返回 (ink_ratio, )"""
    x0, y0, x1, y1 = [int(v * zoom) for v in bbox_pt]
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(bw.shape[1], x1), min(bw.shape[0], y1)
    if x1 <= x0 or y1 <= y0:
        return 0.0, 0.0
    roi = bw[y0:y1, x0:x1]
    ratio = float(roi.mean() / 255.0)
    return ratio, 0.0


def _solid_run(mask_1d: np.ndarray, gap: int = 1):
    """一维掩码的最大连续段：返回 (长度, 起点, 终点, 实心度)。

    实心度 = 段内暗像素数/段长，用于排除表格线与正文笔画
    以 1px 间隙链式连通形成的假长线。
    """
    idx = np.flatnonzero(mask_1d)
    if len(idx) == 0:
        return 0, 0, 0, 0.0
    best = (1, int(idx[0]), int(idx[0]), 1.0)
    run, start, prev, npx = 1, idx[0], idx[0], 1
    for v in idx[1:]:
        if v - prev <= gap:
            run += 1
            npx += 1
            if run > best[0]:
                best = (run, int(start), int(v), npx / run)
        else:
            run, start, npx = 1, v, 1
        prev = v
    return best


def _all_runs(mask_1d: np.ndarray, gap: int = 1):
    """一维掩码的全部连续段：返回 [(长度, 起点, 终点, 实心度)]。"""
    idx = np.flatnonzero(mask_1d)
    if len(idx) == 0:
        return []
    runs = []
    start = prev = idx[0]
    npx = 1
    for v in idx[1:]:
        if v - prev <= gap:
            prev = v
            npx += 1
        else:
            runs.append((prev - start + 1, int(start), int(prev), npx / (prev - start + 1)))
            start = prev = v
            npx = 1
    runs.append((prev - start + 1, int(start), int(prev), npx / (prev - start + 1)))
    return runs


def detect_table_lines(img_gray: np.ndarray, dark_thr: int = 170,
                       h_min_ratio: float = 0.25, v_min_ratio: float = 0.03):
    """扫描页表格线检测（投影 + 实心连续段双条件）。

    返回 (hlines, vlines)，元素为 (pos_px, a_px, b_px, len_px)。
    同一列/行上的多条分离线段都会被提取（如上下排列的两个表格）。
    h_min_ratio/v_min_ratio: 相对图宽/高的最短线长比例。
    """
    h, w = img_gray.shape
    dark = img_gray < dark_thr
    h_min = int(h_min_ratio * w)
    v_min = int(v_min_ratio * h)

    def _group(pairs, gap):
        pairs.sort()
        out = []
        for pos, a, b, ln in pairs:
            if out and pos - out[-1][0] <= gap:
                po, pa, pb, pl = out[-1]
                out[-1] = (pos, min(pa, a), max(pb, b), max(pl, ln))
            else:
                out.append((pos, a, b, ln))
        return out

    hlines, vlines = [], []
    row_ratio = dark.mean(axis=1)
    for r in np.flatnonzero(row_ratio > 0.30):
        for ln, a, b, solid in _all_runs(dark[r], gap=1):
            if ln < h_min or solid < 0.55:
                continue
            # 核心区间收缩：去掉两端稀疏脏点（用细条带列密度判定）
            band = dark[max(0, r - 1):r + 2, a:b]
            prof = band.mean(axis=0)          # bool 数组，mean 即密度 0~1
            dense = np.flatnonzero(prof > 0.4)
            if len(dense) > h_min and dense[-1] - dense[0] > h_min:
                hlines.append((int(r), a + int(dense[0]),
                               a + int(dense[-1]) + 1, int(dense[-1] - dense[0]) + 1))
    col_ratio = dark.mean(axis=0)
    for c in np.flatnonzero(col_ratio > 0.02):
        for ln, a, b, solid in _all_runs(dark[:, c], gap=1):
            if ln >= v_min and solid >= 0.7:
                vlines.append((int(c), a, b, ln))

    hg = _group(hlines, gap=6)
    vg = _group(vlines, gap=6)
    return hg, vg


def detect_grid_tables(img_gray: np.ndarray, zoom: float, page_rect: fitz.Rect):
    """横线驱动的表格区域定位。返回 (tables, (hg, vg))，坐标 pt。

    表格必然有 ≥2 条横线：把 x 范围重叠的横线聚为一族，族内按 y 间隙
    切分（上下排列的两个表），再用区间内竖线数验证。
    """
    hg, vg = detect_table_lines(img_gray)
    inv = 1.0 / zoom
    tables = []
    if len(hg) < 2:
        return tables, (hg, vg)

    fams: list[list] = []
    for line in hg:
        placed = False
        for fam in fams:
            fa0 = min(l[1] for l in fam)
            fb1 = max(l[2] for l in fam)
            ov = min(fb1, line[2]) - max(fa0, line[1])
            if ov > 0.5 * min(fb1 - fa0, line[2] - line[1]):
                fam.append(line)
                placed = True
                break
        if not placed:
            fams.append([line])

    for fam in fams:
        fam.sort()
        clusters = [[fam[0]]]
        for line in fam[1:]:
            if line[0] - clusters[-1][-1][0] < 35 * zoom:   # 行高上限 ~25pt
                clusters[-1].append(line)
            else:
                clusters.append([line])
        for cl in clusters:
            if len(cl) < 2:
                continue
            y0 = min(l[0] for l in cl)
            y1 = max(l[0] for l in cl)
            x0 = min(l[1] for l in cl)
            x1 = max(l[2] for l in cl)
            n_v = 0
            for vx, vy0, vy1, _ in vg:
                if x0 - 8 <= vx <= x1 + 8:
                    ov = min(vy1, y1) - max(vy0, y0)
                    if ov > 0.5 * (y1 - y0):
                        n_v += 1
            if n_v >= 2:
                tables.append([x0 * inv, y0 * inv, x1 * inv, y1 * inv])
    tables.sort(key=lambda t: t[1])
    return tables, (hg, vg)


def table_grid(img_gray: np.ndarray, zoom: float, bbox_pt: list):
    """在表格 bbox 内求行列分割线（pt）。返回 (row_ys, col_xs)。

    行 = 内部横线 + 上下边界；列 = 内部竖线 + 左右边界（原书表格常无左右边框）。
    """
    x0, y0, x1, y1 = [int(v * zoom) for v in bbox_pt]
    x0, y0 = max(0, x0), max(0, y0)
    roi = img_gray[y0:y1, x0:x1]
    if roi.size == 0:
        return [], []
    rh, rw = roi.shape
    # ROI 内按表格几何给更严的线长阈值：横线跨大半表宽，竖线跨大半表高
    hg, vg = detect_table_lines(roi, h_min_ratio=0.6, v_min_ratio=0.5)

    row_ys = sorted(p for p, a, b, l in hg)
    col_xs = sorted(p for p, a, b, l in vg)
    if row_ys and row_ys[0] > 5 * zoom:
        row_ys.insert(0, 0)
    if row_ys and (rh - row_ys[-1]) > 5 * zoom:
        row_ys.append(rh - 1)
    if col_xs and col_xs[0] > 5 * zoom:
        col_xs.insert(0, 0)
    if col_xs and (rw - col_xs[-1]) > 5 * zoom:
        col_xs.append(rw - 1)
    if len(row_ys) < 2:                       # 无横线：上下边界切一行
        row_ys = [0, rh - 1]
    if len(col_xs) < 2:
        col_xs = [0, rw - 1]
    inv = 1.0 / zoom
    return ([y0 * inv + v * inv for v in row_ys],
            [x0 * inv + v * inv for v in col_xs])


def _peaks(profile: np.ndarray, thr: float, min_gap: int) -> list[int]:
    idx = np.where(profile > thr)[0]
    groups = []
    for i in idx:
        if groups and i - groups[-1][-1] <= 3:
            groups[-1].append(i)
        else:
            groups.append([i])
    lines = [int(np.mean(g)) for g in groups]
    # 合并过近的线
    merged: list[int] = []
    for v in lines:
        if merged and v - merged[-1] < min_gap:
            merged[-1] = (merged[-1] + v) // 2
        else:
            merged.append(v)
    return merged


def detect_ink_regions(bw: np.ndarray, zoom: float, words: list[W],
                       page_area_pt: float):
    """检测 OCR 词块稀疏但墨迹密集的大块区域（插图）。"""
    # 去掉小连通域（文字），剩下的大连通域即图形
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    dense = cv2.morphologyEx(bw, cv2.MORPH_CLOSE, kernel, iterations=2)
    n, labels, stats, _ = cv2.connectedComponentsWithStats(dense, 8)
    inv = 1.0 / zoom
    regions = []
    for i in range(1, n):
        x, y, ww, hh, area = stats[i]
        if ww < 90 or hh < 50 or area < 2500:
            continue
        bbox_pt = [x * inv, y * inv, (x + ww) * inv, (y + hh) * inv]
        if (bbox_pt[2] - bbox_pt[0]) * (bbox_pt[3] - bbox_pt[1]) < page_area_pt * 0.01:
            continue
        # 词块覆盖率：插图区域词少
        n_words = sum(1 for wd in words
                      if wd.cx > bbox_pt[0] and wd.cx < bbox_pt[2]
                      and wd.cy > bbox_pt[1] and wd.cy < bbox_pt[3])
        w_area = sum(wd.w * wd.h for wd in words
                     if wd.cx > bbox_pt[0] and wd.cx < bbox_pt[2]
                     and wd.cy > bbox_pt[1] and wd.cy < bbox_pt[3])
        box_area = (bbox_pt[2] - bbox_pt[0]) * (bbox_pt[3] - bbox_pt[1])
        if n_words <= 6 or w_area / box_area < 0.05:
            regions.append(bbox_pt)
    return regions


# ---------------------------------------------------------------- 分类

import re as _re

_CODE_HINT = _re.compile(
    r"[A-Za-z_][A-Za-z0-9_]*\s*\(|'[^']*'|\"[^\"]*\"|\bfor\b|\bif\b|\bend\b"
    r"|==|<=|>=|\[[^\]]*\]\s*[=,;)]|;\s*$")
_STRONG_MATH = set("∑∏∫≤≥≠∈√±×÷∞→←↔⇒⇔⊂⊃⊆⊇∪∩∉∀∂∇")


def _cjk_ratio(t: str) -> float:
    if not t:
        return 0.0
    return sum(1 for c in t if 0x4E00 <= ord(c) <= 0x9FFF) / len(t)


def classify_formula_line(ln: Line, base_h: float, bw: np.ndarray, zoom: float) -> bool:
    """独立行是否为公式行（需裁图）。

    数学教材正文行的特征是汉字主导；公式行以数字/符号/单字母变量为主。
    判成公式的代价只是该行不可编辑（视觉无损），漏判的代价是乱码文本，
    因此阈值取向"宁判公式"。
    """
    t = ln.text.strip()
    if not t:
        return False
    cr = _cjk_ratio(t)

    # 过短行（如"件可得"等段落尾词）不判公式：这类行被判公式会把
    # formula 块 bbox 顶进相邻正文区域，造成裁图与文本重复
    if len(t) < 5:
        return False
    # 纯中文行：正文（含"解""假设"等短标题）
    if cr > 0.85:
        return False
    # 代码行特征：引号/函数调用/关键字 → 保留为文本
    if _CODE_HINT.search(t):
        return False
    # 强数学符号
    if any(c in _STRONG_MATH for c in t):
        return True
    # 等式/算式结构：短行含 = 直接判公式（含"正文+行内公式"粘连行，
    # 其 OCR 多为碎片，裁图比文本更可读）
    if ("=" in t) and len(t) <= 45:
        return True
    # 墨迹-文本不匹配：公式符号被 OCR 漏识别，墨迹远多于文字
    # （对中文碎片与公式粘连的行也能命中）
    ink, _ = ink_stats(bw, zoom, ln.bbox)
    if ink > 0.04:
        est = sum(0.16 if ord(c) > 0x2E7F else 0.10 for c in t)
        expected = est * max(base_h, ln.h) * 0.9 / max(1.0, (ln.bbox[2] - ln.bbox[0]))
        if ink > expected * 1.8:
            return True
    # 中文主导：正文
    if cr > 0.65:
        return False
    # 上下标/分式结构：词块垂直跨度明显超过行高
    if ln.v_span > base_h * 1.45 and len(t) < 40:
        return True
    # 数字/字母与标点混排的短行（如 "x₁=1200,x₂=230" 或 OCR 破碎的公式）
    if len(t) <= 45 and cr < 0.3:
        n_alpha = sum(1 for c in t if c.isascii() and c.isalpha())
        n_digit = sum(1 for c in t if c.isdigit())
        if n_digit >= 2 and (n_alpha + n_digit) >= len(t) * 0.5:
            return True
    return False


def ink_boldness(bw: np.ndarray, zoom: float, bbox_pt: list, base_ink: float) -> float:
    """区域墨迹密度相对正文的倍数（粗体/黑体笔画更密）。"""
    ink, _ = ink_stats(bw, zoom, bbox_pt)
    if base_ink <= 0:
        return 1.0
    return ink / base_ink
