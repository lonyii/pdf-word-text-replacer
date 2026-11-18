#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Add an invisible, searchable/selectable OCR text layer to scanned PDF pages.

Uses the built-in Windows OCR engine (Simplified Chinese) - fully silent,
no dialogs, no mouse. This is the scripted alternative to Acrobat's
"Recognize Text" button (which cannot be driven silently via Acrobat IAC).

How it works, per page:
  1. Render the page to a 300-DPI image (PyMuPDF).
  2. Windows OCR (winrt) returns each recognized word with its pixel box.
  3. Word boxes are mapped back to PDF points and the text is inserted with
     render_mode=3 (invisible) using an embedded CJK font - so the page looks
     unchanged but its text becomes searchable and selectable.

Usage:
  # One file -> <name>_ocr.pdf next to source
  python pdf_ocr_overlay.py --input scan.pdf

  # Batch a folder -> results in a Desktop timestamped folder
  python pdf_ocr_overlay.py --input "D:\\scans" --out-dir

  # Custom output path / DPI / pages
  python pdf_ocr_overlay.py --input scan.pdf --out out.pdf --dpi 300 --pages 1,3-5

Dependencies (Python 3.12 in this environment):
  python -m pip install pymupdf winrt-runtime winrt-Windows.Foundation \
      winrt-Windows.Globalization winrt-Windows.Graphics.Imaging \
      winrt-Windows.Media.Ocr winrt-Windows.Storage winrt-Windows.Storage.Streams
"""
import os
import sys
import asyncio
import argparse
import datetime
import tempfile
import glob

import fitz  # PyMuPDF

from winrt.windows.media.ocr import OcrEngine
from winrt.windows.graphics.imaging import BitmapDecoder
from winrt.windows.storage import StorageFile, FileAccessMode
from winrt.windows.globalization import Language

# CJK font for invisible text (simhei subsets well -> smaller output)
CJK_FONT = r"C:\Windows\Fonts\simhei.ttf"


def log(msg):
    print(msg, flush=True)


def desktop_path():
    try:
        import win32com.shell.shell as shell
        import win32com.shell.shellcon as shellcon
        return shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


def parse_pages(spec, total):
    if not spec:
        return list(range(total))
    out = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a, b = int(a), int(b)
            if a > b:
                a, b = b, a
            for n in range(a, b + 1):
                if 1 <= n <= total:
                    out.add(n - 1)
        else:
            n = int(part)
            if 1 <= n <= total:
                out.add(n - 1)
    return sorted(out)


async def _ocr_png_async(png_bytes, lang_tag):
    """Return (words, pixel_width, pixel_height); words = (text,x,y,w,h)."""
    fd, tmp = tempfile.mkstemp(suffix=".png")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(png_bytes)
        file = await StorageFile.get_file_from_path_async(tmp)
        stream = await file.open_async(FileAccessMode.READ)
        decoder = await BitmapDecoder.create_async(stream)
        bitmap = await decoder.get_software_bitmap_async()
        try:
            engine = OcrEngine.try_create_from_language(Language(lang_tag))
        except Exception:
            engine = None
        if engine is None:
            engine = OcrEngine.try_create_from_user_profile_languages()
        result = await engine.recognize_async(bitmap)
        words = []
        for line in result.lines:
            for w in line.words:
                b = w.bounding_rect
                words.append((w.text, float(b.x), float(b.y),
                              float(b.width), float(b.height)))
        pw, ph = bitmap.pixel_width, bitmap.pixel_height
        stream.close()
        return words, pw, ph
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


def collect_inputs(paths):
    found, seen = [], set()
    for p in paths:
        p = p.strip().strip('"')
        if os.path.isdir(p):
            cands = [f for f in glob.glob(os.path.join(p, "**", "*.pdf"), recursive=True)
                     if os.path.isfile(f)]
        elif any(c in p for c in "*?"):
            cands = [f for f in glob.glob(p, recursive=True) if os.path.isfile(f)]
        elif os.path.isfile(p):
            cands = [p]
        else:
            log("WARN|not found: " + p)
            cands = []
        for f in cands:
            if os.path.splitext(f)[1].lower() != ".pdf":
                continue
            rp = os.path.normcase(os.path.abspath(f))
            if rp not in seen:
                seen.add(rp)
                found.append(f)
    return found


def overlay_one(src, out, dpi, lang_tag, pages_spec, verbose=True):
    doc = fitz.open(src)
    indices = parse_pages(pages_spec, len(doc))
    total_words = 0
    for idx in indices:
        page = doc[idx]
        pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, alpha=False)
        words, pw, ph = asyncio.run(_ocr_png_async(pix.tobytes("png"), lang_tag))
        pr = page.rect
        sx, sy = pr.width / pw, pr.height / ph
        for text, x, y, w, h in words:
            if not text:
                continue
            fontsize = max(4.0, h * sy * 0.92)
            try:
                page.insert_text((x * sx, (y + h * 0.82) * sy), text,
                                 fontname="ocr_cjk", fontfile=CJK_FONT,
                                 fontsize=fontsize, render_mode=3, overlay=True)
            except Exception:
                pass
        total_words += len(words)
        if verbose:
            log(f"PAGE_OK|{os.path.basename(src)}|p{idx+1}|words={len(words)}")
    doc.save(out, garbage=3, deflate=True)
    doc.close()
    return total_words


def main():
    ap = argparse.ArgumentParser(description="Add invisible OCR text layer to scanned PDFs")
    ap.add_argument("--input", nargs="+", required=True, help="PDF file(s), folder(s), or glob")
    ap.add_argument("--out", default="", help="Single output PDF path (only for one input)")
    ap.add_argument("--out-dir", action="store_true",
                    help="Write <name>_ocr.pdf into a Desktop timestamped folder")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--lang", default="zh-Hans-CN")
    ap.add_argument("--pages", default="", help="Pages to OCR, e.g. '1,3-5' (default all)")
    args = ap.parse_args()

    files = collect_inputs(args.input)
    log(f"INPUTS|{len(files)}")
    if not files:
        log("DONE|0")
        return

    outdir = ""
    if args.out_dir:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        outdir = os.path.join(desktop_path(), f"OCR文字层_{ts}")
        os.makedirs(outdir, exist_ok=True)
        log(f"OUTDIR|{outdir}")

    results = []
    for f in files:
        if args.out:
            out = args.out
        elif outdir:
            out = os.path.join(outdir, os.path.splitext(os.path.basename(f))[0] + "_ocr.pdf")
        else:
            out = os.path.splitext(f)[0] + "_ocr.pdf"
        try:
            n = overlay_one(f, out, args.dpi, args.lang, args.pages)
            log(f"OK|{f}|{out}|words={n}|size={os.path.getsize(out)}")
            results.append(out)
        except Exception as e:
            log(f"FAIL|{f}|{type(e).__name__}: {e}")

    final_dir = outdir or (os.path.dirname(results[0]) if results else "")
    log(f"DONE|{len(results)}|{final_dir}")


if __name__ == "__main__":
    main()
