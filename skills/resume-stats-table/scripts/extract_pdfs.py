#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""把目录下所有 PDF 简历提取成 UTF-8 文本，用于后续人工/模型解析。
依赖: PyMuPDF (pymupdf)。用法:
    python extract_pdfs.py <简历目录> <输出目录>
输出: <输出目录>/<每份PDF同名>.txt
"""
import sys, os, glob

def extract(pdf_dir, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            sys.exit("需要 PyMuPDF，请先 pip install pymupdf（见 SKILL.md 依赖安装，注意 TMPDIR 被占用问题）")
    pdfs = sorted(glob.glob(os.path.join(pdf_dir, "*.pdf")))
    if not pdfs:
        print("未找到 PDF 文件:", pdf_dir); return 0
    n = 0
    for p in pdfs:
        base = os.path.splitext(os.path.basename(p))[0]
        try:
            d = pymupdf.open(p)
            parts = []
            for i, page in enumerate(d):
                parts.append(f"===== PAGE {i+1} =====\n")
                parts.append(page.get_text())
            text = "\n".join(parts)
            with open(os.path.join(out_dir, base + ".txt"), "w", encoding="utf-8") as f:
                f.write(text)
            print(f"OK  {base} ({d.page_count} pages)")
            n += 1
        except Exception as e:
            print(f"ERR {base}: {e}")
    print(f"DONE {n}/{len(pdfs)}")
    return n

if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit("用法: python extract_pdfs.py <简历目录> <输出目录>")
    extract(sys.argv[1], sys.argv[2])
