# -*- coding: utf-8 -*-
"""命令行入口：python -m book2docx input.pdf [-o out.docx] [--mode auto] [--pages 1-20]"""
from __future__ import annotations

import argparse
import sys
import time

import fitz

from .common import PageKind, analyze_document, classify_page


def parse_pages(s: str | None, total: int) -> range | None:
    if not s:
        return None
    rng = range(total)
    out = []
    for part in s.split(","):
        if "-" in part:
            a, b = part.split("-")
            out.extend(range(max(0, int(a) - 1), min(int(b), total)))
        else:
            out.append(max(0, min(int(part) - 1, total - 1)))
    return [i for i in out if i in rng] and range(min(out), max(out) + 1) or None


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="book2docx", description="书本 PDF 高保真转 Word")
    ap.add_argument("input", help="输入 PDF 路径")
    ap.add_argument("-o", "--output", help="输出 docx 路径（默认同目录同名）")
    ap.add_argument("-m", "--mode", choices=["auto", "editable", "faithful", "digital"],
                    default="auto", help="转换模式")
    ap.add_argument("-p", "--pages", help="页码范围，如 10-20 或 5,8,11（1-based）")
    ap.add_argument("--dpi", type=int, default=300, help="裁图/渲染 DPI（默认 300）")
    args = ap.parse_args(argv)

    doc = fitz.open(args.input)
    pages = parse_pages(args.pages, doc.page_count)
    out = args.output or (args.input.rsplit(".", 1)[0] + ".docx")

    ana = analyze_document(doc)
    print(f"[book2docx] 页数={ana['page_count']}  扫描页占比={ana['scan_ratio']:.0%}  "
          f"OCR质量均值={ana['mean_ocr_quality']:.2f}")
    mode = args.mode
    if mode == "auto":
        if ana["scan_ratio"] > 0.5:
            mode = "editable" if ana["mean_ocr_quality"] > 0.6 else "faithful"
        else:
            mode = "digital"
    print(f"[book2docx] 模式: {mode}")

    t0 = time.time()
    last_report = [0]

    def progress(i, n):
        pct = int((last_report[0] + 1) / n * 100)
        if pct >= last_report[0] + 10 or last_report[0] + 1 == n:
            print(f"  ... {last_report[0]+1}/{n} 页")
            last_report[0] = last_report[0] + 1

    if mode == "faithful":
        from .faithful import convert as conv
    elif mode == "digital":
        from .digital import convert as conv
    else:
        from .editable import convert as conv

    if pages is not None:
        conv(doc, out, pages=pages, progress=progress, **({"dpi": args.dpi}
           if mode in ("editable", "faithful") else {}))
    else:
        conv(doc, out, progress=progress, **({"dpi": args.dpi}
            if mode in ("editable", "faithful") else {}))

    import os
    print(f"[book2docx] 完成: {out}  ({os.path.getsize(out)/1e6:.1f} MB, "
          f"耗时 {time.time()-t0:.0f}s)")


if __name__ == "__main__":
    main()
