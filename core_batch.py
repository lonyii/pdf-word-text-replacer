#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Batch pipeline: codes -> Everything search -> PDF->Word -> replace -> open output.

Config is embedded as JSON by default (no external file). If batch_config.json
exists next to the EXE it overrides the embedded defaults. Users can "release"
the embedded config to that file for editing.
"""
import os
import sys
import json
import subprocess

CONFIG_NAME = "batch_config.json"

# Embedded default config (JSON string, kept verbatim so it can be released).
DEFAULT_CONFIG_JSON = """{
  "es_exe": "bin\\\\es.exe",
  "search_template": "{code}  线上 .pdf  jszt-pc2 \\\\",
  "fallback_template": "{code}-线上资料.pdf",
  "host_keyword": "jszt-pc2",
  "prefer_unc": true,
  "replace_rules": [
    {"find": "中石化工建设有限公司", "replace": "江苏中天市政工程有限公司"}
  ],
  "output_prefix": "批处理"
}"""


def _app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def config_path():
    return os.path.join(_app_dir(), CONFIG_NAME)


def load_config():
    """Return (config_dict, is_external). External file wins if present."""
    cfg = json.loads(DEFAULT_CONFIG_JSON)
    p = config_path()
    if os.path.isfile(p):
        try:
            with open(p, "r", encoding="utf-8-sig") as fh:
                ext = json.load(fh)
            cfg.update(ext)
            return cfg, True
        except Exception:
            pass
    return cfg, False


def release_config():
    """Write the embedded default config next to the EXE for editing."""
    p = config_path()
    with open(p, "w", encoding="utf-8") as fh:
        fh.write(DEFAULT_CONFIG_JSON + "\n")
    return p


def resolve_es(cfg):
    """Resolve es.exe path (absolute, or relative to the app directory)."""
    es = cfg.get("es_exe", "")
    if not os.path.isabs(es):
        es = os.path.join(_app_dir(), es)
    return es


def _decode_es(data):
    for enc in ("gbk", "cp936", "utf-8"):
        try:
            return data.decode(enc)
        except Exception:
            continue
    return data.decode("utf-8", "replace")


def _es_query(es, query, n=30):
    p = subprocess.run([es, "-n", str(n), query], capture_output=True, timeout=30)
    return [ln.strip() for ln in _decode_es(p.stdout).splitlines() if ln.strip()]


def search_code(code, cfg, log):
    """Search one code via es.exe (plain Everything query, not regex).

    The configured search_template is tried first. Some Everything CLI builds
    fail on multi-word queries; in that case fall back to a single continuous
    query. Results are filtered locally to .pdf / preferred host paths.
    """
    es = resolve_es(cfg)
    if not os.path.isfile(es):
        raise RuntimeError(f"找不到 Everything 命令行工具：{es}")
    query = cfg.get("search_template", "{code}  线上 .pdf  jszt-pc2 \\").format(code=code)
    try:
        lines = _es_query(es, query)
        used = "配置模板"
        if not lines:
            # Fallback 1: continuous single-token query (works on all CLI builds)
            fb = cfg.get("fallback_template", "{code}-线上资料.pdf").format(code=code)
            if fb and fb != query:
                lines = _es_query(es, fb)
                used = "兜底1"
        if not lines:
            # Fallback 2: bare code query, filter locally for 线上 PDFs.
            # Covers naming like "VX26F002-b.线上资料.pdf".
            lines = _es_query(es, code, n=100)
            used = "兜底2"
    except Exception as e:
        raise RuntimeError(f"Everything 搜索失败：{type(e).__name__}: {e}")
    if not lines:
        log(f"  [{code}] 未找到匹配文件")
        return None
    pdfs = [ln for ln in lines if os.path.splitext(ln)[1].lower() == ".pdf"]
    if used == "兜底2":
        # bare-code query can return anything; keep only 线上 PDFs
        online = [x for x in pdfs if "线上" in os.path.basename(x)]
        if online:
            pdfs = online
    if not pdfs:
        log(f"  [{code}] 找到 {len(lines)} 项，但无 PDF")
        return None
    # Prefer paths containing the configured host keyword (e.g. jszt-pc2),
    # then UNC paths, then the first result.
    host = cfg.get("host_keyword", "jszt-pc2")
    if host:
        hits = [x for x in pdfs if host.lower() in x.lower()]
        if hits:
            log(f"  [{code}] 命中 {host} 路径 PDF（{used}，共 {len(pdfs)} 个 PDF）")
            return hits[0]
    if cfg.get("prefer_unc", True):
        unc = [x for x in pdfs if x.startswith("\\\\")]
        if unc:
            log(f"  [{code}] 命中 UNC 路径 PDF（{used}，共 {len(pdfs)} 个 PDF）")
            return unc[0]
    log(f"  [{code}] 命中 PDF（{used}，共 {len(pdfs)} 个）")
    return pdfs[0]


def search_all(codes, cfg, log):
    """Return list of (code, pdf_path); codes with no PDF are skipped."""
    found = []
    for code in codes:
        code = code.strip()
        if not code:
            continue
        try:
            pdf = search_code(code, cfg, log)
        except Exception as e:
            log(f"  [{code}] 搜索异常：{type(e).__name__}: {e}")
            continue
        if pdf:
            found.append((code, pdf))
    return found


def _replace_docx(src, dst, rules):
    """Open a .docx in Word, run find/replace over all stories, SaveAs dst."""
    import pythoncom
    import win32com.client as win32
    import office_core as core
    pythoncom.CoInitialize()
    word = doc = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(src, False, True, False)
        for story in doc.StoryRanges:
            core._apply_rules_story(story, rules)
            linked = getattr(story, "NextStoryRange", None)
            while linked is not None:
                core._apply_rules_story(linked, rules)
                linked = getattr(linked, "NextStoryRange", None)
        if os.path.normcase(os.path.abspath(dst)) == os.path.normcase(os.path.abspath(src)):
            doc.Save()  # in-place overwrite
        else:
            doc.SaveAs2(dst, 12)  # wdFormatXMLDocument
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def process_pdfs(items, log, cfg=None):
    """Process (label, pdf_path) pairs: PDF->Word -> replace -> output.
    Returns (outdir, ok, total).
    """
    import office_core as core
    if cfg is None:
        cfg, _ = load_config()
    if not items:
        log("没有可处理的 PDF，批处理结束")
        return None, 0, 0
    outdir = core.make_output_dir(cfg.get("output_prefix", "批处理"))
    log(f"输出目录：{outdir}")
    rules = cfg.get("replace_rules", [])
    ok = 0
    for label, pdf in items:
        stem = os.path.splitext(os.path.basename(pdf))[0]
        tmp_docx = os.path.join(outdir, stem + ".docx")
        log(f"[{label}] PDF 转 Word…")
        try:
            if not core._acrobat_convert(pdf, tmp_docx):
                log(f"  [失败] PDF 转 Word 未生成输出")
                continue
            try:
                core.fix_docx_header_overlap(tmp_docx)
            except Exception:
                pass
        except Exception as e:
            log(f"  [失败] PDF 转 Word：{type(e).__name__}: {e}")
            continue
        log(f"[{label}] 替换文本…")
        replaced = os.path.join(outdir, stem + "_replaced.docx")
        try:
            _replace_docx(tmp_docx, replaced, rules)
            os.replace(replaced, tmp_docx)   # atomically overwrite the original
        except Exception as e:
            log(f"  [失败] 替换文本：{type(e).__name__}: {e}")
            try:
                if os.path.exists(replaced):
                    os.remove(replaced)
            except Exception:
                pass
            continue
        log(f"  [OK] -> {os.path.basename(tmp_docx)}")
        ok += 1
    log(f"完成：{ok}/{len(items)} 个成功")
    try:
        os.startfile(outdir)  # open output folder in Explorer
    except Exception:
        pass
    return outdir, ok, len(items)


def run_batch(codes, log, cfg=None):
    """Search codes then process. Kept for compatibility."""
    if cfg is None:
        cfg, _ = load_config()
    log("正在搜索工程编码…")
    found = search_all(codes, cfg, log)
    return process_pdfs(found, log, cfg)


def run_batch_files(files, log, cfg=None):
    """Process an explicit list of PDF files (from the batch tab's file list)."""
    if cfg is None:
        cfg, _ = load_config()
    items = [(os.path.basename(f), f) for f in files
             if os.path.splitext(f)[1].lower() == ".pdf"]
    return process_pdfs(items, log, cfg)
