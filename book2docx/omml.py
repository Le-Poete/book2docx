# -*- coding: utf-8 -*-
"""LaTeX → Word 原生公式（OMML）转换。

链路：pix2tex(LaTeX-OCR) 识别公式图 → LaTeX → MathML(latex2mathml)
→ XSLT(Office 自带 MML2OMML.XSL) → OMML → 插入 docx 段落。
任一环节失败由调用方回退为高清裁图。
"""
from __future__ import annotations

import os
import re
from pathlib import Path

_LATEX2MML = None
_XSLT = None
_XSL_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
]


def _latex2mathml():
    global _LATEX2MML
    if _LATEX2MML is None:
        from latex2mathml.converter import convert as _c
        _LATEX2MML = _c
    return _LATEX2MML


def _xslt():
    global _XSLT
    if _XSLT is None:
        from lxml import etree
        xsl_path = os.environ.get("MML2OMML_XSL")
        if not xsl_path:
            for cand in _XSL_CANDIDATES:
                if Path(cand).exists():
                    xsl_path = cand
                    break
        if not xsl_path:
            raise RuntimeError("未找到 MML2OMML.XSL（需要安装 Microsoft Office）")
        # 禁用外部实体与网络访问（防 XXE）
        parser = etree.XMLParser(resolve_entities=False, no_network=True,
                                 load_dtd=False)
        xsl_root = etree.parse(xsl_path, parser)
        _XSLT = etree.XSLT(xsl_root)
    return _XSLT


def latex_available() -> bool:
    """环境是否具备公式转换条件（惰性探测并缓存结果）。"""
    global _LATEX2MML_OK, _XSL_OK
    try:
        _latex2mathml()
        _LATEX2MML_OK = True
    except Exception:
        _LATEX2MML_OK = False
    try:
        _xslt()
        _XSL_OK = True
    except Exception:
        _XSL_OK = False
    return _LATEX2MML_OK and _XSL_OK


def latex_plausible(tex: str) -> bool:
    """识别结果的合理性门控：必须有数学结构特征（=/^/_/rac 等），
    且不含 CJK 乱码。不满足则回退高清裁图。"""
    if not tex or len(tex) < 4:
        return False
    if any(0x4E00 <= ord(c) <= 0x9FFF for c in tex):
        return False
    # 数学结构特征：等号/上下标/分数/求和/积分/根号
    if not re.search(r"[=^_]|\frac|\sum|\int|\sqrt", tex):
        return False
    good = sum(1 for c in tex
               if c.isalnum() or c in "+-={}^_().,|/<>[] !;:'")
    return good / len(tex) >= 0.85


def clean_latex(tex: str) -> str:
    """清理 pix2tex 输出中 Word 公式不友好的包装。"""
    tex = tex.strip()
    # 去掉整段 array/overbrace 包装（pix2tex 对行间公式的常见输出）
    tex = re.sub(r"^\\begin\{array\}\{[a-z]+\}|\\end\{array\}$", "", tex).strip()
    tex = tex.replace("\\overbrace{", "{").replace("\\boldsymbol", "\\mathbf")
    # 尾部多余括号配平由 latex2mathml 容错
    if tex.startswith("{") and tex.endswith("}"):
        tex = tex[1:-1]
    return tex


def latex_to_omml_xml(tex: str):
    """LaTeX → OMML 元素（oMathPara 外层）。失败抛异常。"""
    from lxml import etree
    tex = clean_latex(tex)
    if not tex or len(tex) > 800:
        raise ValueError("公式过长或为空")
    mathml = _latex2mathml()(tex)
    # MathML → OMML（禁用外部实体，防 XXE；输入为内部生成的 MathML）
    parser = etree.XMLParser(resolve_entities=False, no_network=True,
                             load_dtd=False)
    mml_tree = etree.fromstring(mathml.encode(), parser)
    omml = _xslt()(mml_tree)
    root = omml.getroot()
    # XSLT 输出可能是 <oMath> 或 <oMathPara>，规范为 oMathPara 包裹
    if root.tag.endswith("}oMath"):
        # 构造 oMathPara 外层
        m_ns = "http://schemas.openxmlformats.org/officeDocument/2006/math"
        para = root.makeelement(f"{{{m_ns}}}oMathPara", {})
        para.append(root)
        root = para
    return root


def append_omml_to_paragraph(paragraph, omml_root, font_size_pt: float = None):
    """把 OMML 元素挂到 docx 段落的 run 里。"""
    r = paragraph.add_run()
    r._r.append(omml_root)
    if font_size_pt:
        from docx.shared import Pt
        from docx.oxml.ns import qn as _qn
        rpr = r._r.get_or_add_rPr()
        sz = rpr.makeelement(_qn("w:sz"), {})
        sz.set(_qn("w:val"), str(int(round(font_size_pt * 2))))
        rpr.append(sz)


_PIX2TEX = None


def pix2tex_model():
    """惰性加载 LaTeX-OCR（pix2tex）模型；不可用时抛异常。"""
    global _PIX2TEX
    if _PIX2TEX is None:
        from pix2tex.cli import LatexOCR
        _PIX2TEX = LatexOCR()
    return _PIX2TEX


def image_to_latex(img) -> str:
    """PIL 图片 → LaTeX（失败抛异常）。调用方需预热 pix2tex_model()。"""
    return pix2tex_model()(img)
