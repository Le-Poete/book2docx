# -*- coding: utf-8 -*-
"""验证：Word COM 将 docx 转 PDF，渲染页面并与原 PDF 页并排对比。"""
from __future__ import annotations

import os

import fitz


def docx_to_pdf(docx_path: str, pdf_path: str | None = None) -> str:
    import pythoncom
    import win32com.client
    pdf_path = pdf_path or docx_path.rsplit(".", 1)[0] + ".pdf"
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        d = word.Documents.Open(os.path.abspath(docx_path))
        d.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        d.Close(False)
    finally:
        word.Quit()
    return pdf_path


def side_by_side(orig_pdf: str, conv_pdf: str, page_idx: int, out_png: str,
                 dpi: int = 90, conv_page_idx: int | None = None):
    a = fitz.open(orig_pdf)
    b = fitz.open(conv_pdf)
    pa = a[page_idx]
    pb = b[conv_page_idx if conv_page_idx is not None else page_idx]
    m = fitz.Matrix(dpi / 72, dpi / 72)
    ia = pa.get_pixmap(matrix=m)
    ib = pb.get_pixmap(matrix=m)
    from PIL import Image
    img_a = Image.frombytes("RGB", (ia.width, ia.height),
                            fitz.Pixmap(fitz.csRGB, ia).samples)
    img_b = Image.frombytes("RGB", (ib.width, ib.height),
                            fitz.Pixmap(fitz.csRGB, ib).samples)
    combo = Image.new("RGB", (img_a.width + img_b.width + 12,
                              max(img_a.height, img_b.height)), (180, 180, 180))
    combo.paste(img_a, (0, 0))
    combo.paste(img_b, (img_a.width + 12, 0))
    combo.save(out_png)
    return out_png
