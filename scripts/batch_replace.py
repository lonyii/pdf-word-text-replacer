#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Batch PDF->Word conversion + find/replace for a folder / file list.

Workflow:
  1. Collect input files (folders are scanned recursively) of .pdf/.docx/.doc.
  2. Convert PDFs to .docx via Acrobat (acrobat_export.ps1, silent, no mouse).
  3. Word documents: find & replace across body / headers / footers / tables /
     text boxes using an independent Word COM instance (preserves formatting).
  4. Save results into a NEW folder on the Desktop named by timestamp.

Usage:
  python batch_replace.py --input <path> [<path> ...] \
      --rules '[{"find":"old","replace":"new"}, ...]' \
      [--out-dir D:\\custom] [--no-convert-pdf]

  Rules can also be a JSON file path passed to --rules (a .json file).
"""
import os
import sys
import json
import glob
import shutil
import argparse
import tempfile
import datetime
import subprocess

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ACROBAT_PS1 = os.path.join(SCRIPT_DIR, "acrobat_export.ps1")
WORD_EXTS = (".docx", ".doc")
PDF_EXTS = (".pdf",)


def log(msg):
    print(msg, flush=True)


def desktop_path():
    """Return the real Desktop path (handles OneDrive redirection)."""
    try:
        import win32com.shell.shell as shell
        import win32com.shell.shellcon as shellcon
        return shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


def collect_inputs(paths, recursive=True):
    """Expand folders / globs / files into a de-duplicated list of target files."""
    found = []
    seen = set()
    for p in paths:
        p = p.strip().strip('"')
        if not p:
            continue
        candidates = []
        if os.path.isdir(p):
            pattern = os.path.join(p, "**", "*") if recursive else os.path.join(p, "*")
            candidates = [f for f in glob.glob(pattern, recursive=recursive)
                          if os.path.isfile(f)]
        elif any(ch in p for ch in "*?"):
            candidates = [f for f in glob.glob(p, recursive=recursive) if os.path.isfile(f)]
        elif os.path.isfile(p):
            candidates = [p]
        else:
            log("WARN|Path not found: " + p)
        for f in candidates:
            ext = os.path.splitext(f)[1].lower()
            if ext in WORD_EXTS + PDF_EXTS:
                rp = os.path.normcase(os.path.abspath(f))
                if rp not in seen:
                    seen.add(rp)
                    found.append(f)
    return found


def pdf_to_docx(pdf, workdir):
    """Silently convert one PDF to .docx using the Acrobat IAC script."""
    base = os.path.splitext(os.path.basename(pdf))[0]
    out = os.path.join(workdir, base + ".docx")
    # Avoid collision if two PDFs share a base name
    n = 1
    while os.path.exists(out):
        out = os.path.join(workdir, f"{base}_{n}.docx")
        n += 1
    cmd = [
        "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", ACROBAT_PS1,
        "-PdfPath", pdf,
        "-Format", "docx",
        "-OutPath", out,
    ]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        stdout = (r.stdout or "").strip()
        if os.path.exists(out) and os.path.getsize(out) > 0:
            log(f"CONVERT_OK|{pdf}|{out}")
            return out
        log(f"CONVERT_FAIL|{pdf}|{stdout}|{(r.stderr or '').strip()[:300]}")
    except Exception as e:
        log(f"CONVERT_FAIL|{pdf}|{type(e).__name__}: {e}")
    return None


def unique_target(outdir, name):
    """Return a non-colliding output path inside outdir."""
    target = os.path.join(outdir, name)
    stem, ext = os.path.splitext(name)
    n = 1
    while os.path.exists(target):
        target = os.path.join(outdir, f"{stem}_{n}{ext}")
        n += 1
    return target


def replace_in_docx(src, dst, rules, word):
    """Open src in the given Word instance, replace text everywhere, SaveAs dst."""
    import win32com.client as win32
    doc = None
    try:
        # AddToRecentFiles=False: do not pollute the user's recent docs list.
        doc = word.Documents.Open(src, False, True, False)
        # Search all story ranges: body, headers, footers, text boxes, footnotes...
        for story in doc.StoryRanges:
            apply_rules(story, rules)
            linked = getattr(story, "NextStoryRange", None)
            while linked is not None:
                apply_rules(linked, rules)
                linked = getattr(linked, "NextStoryRange", None)
        # wdFormatXMLDocument = 12 (.docx), independent of source format.
        doc.SaveAs2(dst, 12)
        size = os.path.getsize(dst) if os.path.exists(dst) else 0
        log(f"REPLACE_OK|{src}|{dst}|{size}")
        return True
    except Exception as e:
        log(f"REPLACE_FAIL|{src}|{type(e).__name__}: {e}")
        return False
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass


def apply_rules(story, rules):
    """Run Word Find/Replace on one story range, preserving formatting.

    IMPORTANT: win32com requires the FULL POSITIONAL argument list for
    Find.Execute. Passing only Replace=2 as a named argument silently does
    nothing. Signature (Word Find.Execute):
        Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards,
                MatchSoundsLike, MatchAllWordForms, Forward, Wrap,
                Format, ReplaceWith, Replace)
    """
    for rule in rules:
        find_text = rule.get("find", "")
        repl_text = rule.get("replace", "")
        if find_text == "":
            continue
        find = story.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        match_case = bool(rule.get("match_case", False))
        whole_word = bool(rule.get("whole_word", False))
        wildcards = bool(rule.get("wildcards", False))
        # wdFindContinue = 1, wdReplaceAll = 2
        find.Execute(
            find_text,          # FindText
            match_case,         # MatchCase
            whole_word,         # MatchWholeWord
            wildcards,          # MatchWildcards
            False,              # MatchSoundsLike
            False,              # MatchAllWordForms
            True,               # Forward
            1,                  # Wrap = wdFindContinue
            False,              # Format
            repl_text,          # ReplaceWith
            2,                  # Replace = wdReplaceAll
        )


def load_rules(raw):
    """Accept either a JSON string or a path to a .json file."""
    if os.path.isfile(raw):
        with open(raw, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
    else:
        data = json.loads(raw)
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list):
        raise ValueError("--rules must be a list of {find, replace} objects")
    clean = []
    for item in data:
        if not isinstance(item, dict) or "find" not in item:
            raise ValueError("Each rule must be an object with a 'find' key")
        clean.append({
            "find": str(item["find"]),
            "replace": str(item.get("replace", "")),
            "match_case": bool(item.get("match_case", False)),
            "whole_word": bool(item.get("whole_word", False)),
            "wildcards": bool(item.get("wildcards", False)),
        })
    return clean


def main():
    ap = argparse.ArgumentParser(description="PDF->Word + batch find/replace")
    ap.add_argument("--input", nargs="+", required=True,
                    help="Input file(s), folder(s), or glob pattern(s)")
    ap.add_argument("--rules", required=True,
                    help='JSON string or .json file path: [{"find":"a","replace":"b"}]')
    ap.add_argument("--out-dir", default="",
                    help="Output folder (default: new timestamped folder on Desktop)")
    ap.add_argument("--no-convert-pdf", action="store_true",
                    help="Skip PDF conversion (only process Word files)")
    args = ap.parse_args()

    rules = load_rules(args.rules)
    log(f"RULES|{json.dumps(rules, ensure_ascii=False)}")

    files = collect_inputs(args.input)
    log(f"INPUTS|{len(files)}|" + json.dumps(files, ensure_ascii=False))
    if not files:
        log("DONE|0 files to process")
        return

    # Output folder on Desktop with timestamp
    if args.out_dir:
        outdir = args.out_dir
    else:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        outdir = os.path.join(desktop_path(), f"Word替换_{ts}")
    os.makedirs(outdir, exist_ok=True)
    log(f"OUTDIR|{outdir}")

    workdir = tempfile.mkdtemp(prefix="pdfword_replace_")
    docs_to_process = []   # list of (path, is_temp)
    try:
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            if ext in PDF_EXTS:
                if args.no_convert_pdf:
                    log(f"SKIP_PDF|{f}")
                    continue
                conv = pdf_to_docx(f, workdir)
                if conv:
                    docs_to_process.append((conv, True))
            else:
                docs_to_process.append((f, False))

        if not docs_to_process:
            log("DONE|0 Word documents to process")
            return

        import win32com.client as win32
        word = None
        try:
            # DispatchEx = a NEW independent Word instance.
            # We only ever close documents WE opened and Quit OUR OWN instance;
            # any Word the user already had open is never touched.
            word = win32.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            for src, is_temp in docs_to_process:
                name = os.path.splitext(os.path.basename(src))[0] + ".docx"
                dst = unique_target(outdir, name)
                replace_in_docx(src, dst, rules, word)
        finally:
            if word is not None:
                try:
                    word.Quit()  # affects only our own DispatchEx instance
                except Exception:
                    pass

        log(f"DONE|{len(docs_to_process)}|{outdir}")
    finally:
        # Always clean up temp conversion dir
        shutil.rmtree(workdir, ignore_errors=True)


if __name__ == "__main__":
    main()
