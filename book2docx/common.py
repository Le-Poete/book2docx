# -*- coding: utf-8 -*-
"""公共工具：页面类型判定、几何换算、字体映射、数学符号检测。"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import Enum

import fitz

PT_PER_CM = 28.3465
EMU_PER_PT = 12700


class PageKind(Enum):
    DIGITAL = "digital"      # 原生文字层（真实字体信息）
    SCAN_OCR = "scan_ocr"    # 扫描页 + OCR 隐形文字层
    SCAN_RAW = "scan_raw"    # 纯扫描页，无文字层


@dataclass
class PageInfo:
    index: int
    kind: PageKind
    width: float            # pt
    height: float           # pt
    text_chars: int = 0
    n_fonts: int = 0
    image_coverage: float = 0.0
    ocr_quality: float = 0.0  # 0~1，扫描页文字层的可信度


# ---------------------------------------------------------------- 页面分析

def classify_page(page: fitz.Page) -> PageInfo:
    """判定页面类型：数字页 / 带 OCR 层扫描页 / 纯扫描页。"""
    pw, ph = page.rect.width, page.rect.height
    area = pw * ph
    info = PageInfo(index=page.number, kind=PageKind.SCAN_RAW,
                    width=pw, height=ph)

    img_cov = 0.0
    for im in page.get_image_info():
        b = im["bbox"]
        cov = max(0.0, min(1.0, (b[2] - b[0]) * (b[3] - b[1]) / area))
        img_cov += cov
    info.image_coverage = min(1.0, img_cov)

    d = page.get_text("dict")
    n_spans = 0
    n_real_fonts = 0
    fonts = set()
    for blk in d["blocks"]:
        if blk["type"] != 0:
            continue
        for line in blk["lines"]:
            for sp in line["spans"]:
                t = sp["text"].strip()
                if t:
                    n_spans += 1
                    info.text_chars += len(t)
                    fonts.add(sp["font"])
    info.n_fonts = len(fonts)

    full_text = page.get_text().strip()
    if info.text_chars < 8:
        info.kind = PageKind.SCAN_RAW
    else:
        # 扫描页特征：一张几乎全页覆盖的底图 + 单一 OCR 字体 + render mode 3
        single_ocr_font = info.n_fonts <= 2
        if info.image_coverage > 0.85 and single_ocr_font and _is_ocr_layer(page):
            info.kind = PageKind.SCAN_OCR
            info.ocr_quality = estimate_ocr_quality(page)
        elif info.image_coverage > 0.85:
            info.kind = PageKind.SCAN_OCR
            info.ocr_quality = estimate_ocr_quality(page)
        else:
            info.kind = PageKind.DIGITAL
    _ = n_spans, n_real_fonts, full_text
    return info


def _is_ocr_layer(page: fitz.Page) -> bool:
    """检查文字是否为隐形渲染（OCR 层通常 render mode=3）。"""
    try:
        raw = page.get_text("rawdict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        for blk in raw["blocks"]:
            if blk["type"] != 0:
                continue
            for line in blk["lines"]:
                for sp in line["spans"]:
                    if sp.get("char_flags", 0) & 1:  # rendered as invisible
                        return True
                    return False
    except Exception:
        pass
    return False


def estimate_ocr_quality(page: fitz.Page) -> float:
    """粗估 OCR 文字层质量：中文字符占比 + 乱码字符占比。

    扫描书 OCR 层的典型病灶：公式区错字、生僻私码区字符（PUA/E0-FF 区）。
    """
    text = page.get_text()
    if not text.strip():
        return 0.0
    n_total = 0
    n_good = 0
    n_bad = 0
    for ch in text:
        cp = ord(ch)
        if ch.isspace():
            continue
        n_total += 1
        if 0xE000 <= cp <= 0xF8FF or 0xFFF0 <= cp <= 0xFFFF:
            n_bad += 1
        elif (
            0x4E00 <= cp <= 0x9FFF or ch.isascii() and (ch.isalnum() or ch in ".,;:!?()[]{}<>=+-*/%\"'")
            or 0x3000 <= cp <= 0x303F or 0xFF00 <= cp <= 0xFFEF
            or ch in "≤≥≠∑∈∫√πμα≠→←↑↓"
        ):
            n_good += 1
    if n_total == 0:
        return 0.0
    return max(0.0, min(1.0, n_good / n_total))


def analyze_document(doc: fitz.Document, sample: int = 12) -> dict:
    """全文抽样分析，决定默认管线。"""
    step = max(1, doc.page_count // sample)
    pages = [classify_page(doc[i]) for i in range(0, doc.page_count, step)]
    kinds = [p.kind for p in pages]
    n_scan = sum(1 for k in kinds if k != PageKind.DIGITAL)
    q = [p.ocr_quality for p in pages if p.kind == PageKind.SCAN_OCR]
    return {
        "page_count": doc.page_count,
        "scan_ratio": n_scan / len(kinds),
        "mean_ocr_quality": sum(q) / len(q) if q else 0.0,
        "size": (doc[0].rect.width, doc[0].rect.height),
    }


# ---------------------------------------------------------------- 几何/单位

def pt_to_cm(pt: float) -> float:
    return pt / PT_PER_CM


def cm_to_pt(cm: float) -> float:
    return cm * PT_PER_CM


def pt_to_emu(pt: float) -> int:
    return int(round(pt * EMU_PER_PT))


# ---------------------------------------------------------------- 字体映射

# PDF 字体名（小写包含匹配）→ (Word 中文字体, Word 西文字体)
FONT_MAP = [
    (r"simsun|songti|song|宋|fangsong|仿宋|fs", ("宋体", "Times New Roman")),
    (r"simhei|hei|黑|heiti", ("黑体", "Arial")),
    (r"kaiti|kai|楷", ("楷体", "Times New Roman")),
    (r"fangzheng|fz|方正", ("宋体", "Times New Roman")),
    (r"msyh|yahei|雅", ("微软雅黑", "Segoe UI")),
    (r"times|serif|roman|cmr|ptm", ("宋体", "Times New Roman")),
    (r"arial|helv|sans|phv", ("黑体", "Arial")),
    (r"courier|mono|cmtt|pcm|consol", ("宋体", "Consolas")),
    (r"cmmi|cmsy|cmex|msam|msbm|math|cambria|symbol", ("Cambria Math", "Cambria Math")),
]


def map_font(pdf_font: str) -> tuple[str, str]:
    """PDF 字体名 → (中文字体, 西文字体)。"""
    name = (pdf_font or "").lower()
    for pat, res in FONT_MAP:
        if re.search(pat, name):
            return res
    return ("宋体", "Times New Roman")


ITALIC_HINT = re.compile(r"italic|oblique|cmti|pti|cmmi", re.I)
BOLD_HINT = re.compile(r"bold|cmb|ptb|-b\b", re.I)


def span_flags(font: str, flags: int) -> dict:
    """从字体名 + PyMuPDF flags 推断粗斜体。"""
    bold = bool(flags & 16) or bool(BOLD_HINT.search(font or ""))
    italic = bool(flags & 2) or bool(ITALIC_HINT.search(font or ""))
    return {"bold": bold, "italic": italic}


# ---------------------------------------------------------------- 数学检测

MATH_CHARS = set("∑∏∫∑∈∀∂∇≤≥≠≈±×÷√∞μαβγλθφωΩΔΣΠ⊂⊃⊆⊇∪∩∉←→↔⇒⇔∘⊗⊕⋅")
MATH_RE = re.compile(r"[∑∏∫∈∀∂∇≤≥≠≈±×÷√∞→←↔⇒⇔⊗⊕⊂⊃⊆⊇∪∩∉]")
SUBSUP_RE = re.compile(r"[₀₁₂₃₄₅₆₇₈₉⁰¹²³⁴⁵⁶⁷⁸⁹ⁿᵢⱼ]")


def math_density(text: str) -> float:
    """文本中数学符号密度，用于公式区域判定。"""
    if not text:
        return 0.0
    n = sum(1 for c in text if c in MATH_CHARS or MATH_RE.match(c) or SUBSUP_RE.match(c))
    return n / max(1, len(text))
