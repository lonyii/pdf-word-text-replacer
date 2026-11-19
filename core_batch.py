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
  "search_template": "{code}-线上资料.pdf",
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


def search_code(code, cfg, log):
    """Search one code via es.exe; return best PDF path or None."""
    es = resolve_es(cfg)
    if not os.path.isfile(es):
        raise RuntimeError(f"找不到 Everything 命令行工具：{es}")
    query = cfg.get("search_template", "{code}-线上资料.pdf").format(code=code)
    try:
        p = subprocess.run([es, "-n", "30", query], capture_output=True, timeout=30)
    except Exception as e:
        raise RuntimeError(f"Everything 搜索失败：{type(e).__name__}: {e}")
    text = _decode_es(p.stdout)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not lines:
        log(f"  [{code}] 未找到匹配文件")
        return None
    pdfs = [ln for ln in lines if os.path.splitext(ln)[1].lower() == ".pdf"]
    if not pdfs:
        log(f"  [{code}] 找到 {len(lines)} 项，但无 PDF")
        return None
    # Prefer UNC (\\\\server\\share) paths, then first result
    if cfg.get("prefer_unc", True):
        unc = [x for x in pdfs if x.startswith("\\\\")]
        if unc:
            log(f"  [{code}] 命中 UNC 路径 PDF（共 {len(pdfs)} 个 PDF）")
            return unc[0]
    log(f"  [{code}] 命中 PDF（共 {len(pdfs)} 个）")
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


def run_batch(codes, log, cfg=None):
    """Full pipeline. Returns (outdir, success_count, total_found)."""
    import shutil
    import office_core as core
    if cfg is None:
        cfg, _ = load_config()
    log("正在搜索工程编码…")
    found = search_all(codes, cfg, log)
    if not found:
        log("没有找到任何可处理的 PDF，批处理结束")
        return None, 0, 0
    outdir = core.make_output_dir(cfg.get("output_prefix", "批处理"))
    log(f"输出目录：{outdir}")
    rules = cfg.get("replace_rules", [])
    ok = 0
    for code, pdf in found:
        stem = os.path.splitext(os.path.basename(pdf))[0]
        tmp_docx = os.path.join(outdir, stem + ".docx")
        log(f"[{code}] PDF 转 Word…")
        try:
            if not core._acrobat_convert(pdf, tmp_docx):
                log(f"  [失败] PDF 转 Word 未生成输出")
                continue
        except Exception as e:
            log(f"  [失败] PDF 转 Word：{type(e).__name__}: {e}")
            continue
        log(f"[{code}] 替换文本…")
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
    log(f"完成：{ok}/{len(found)} 个成功")
    try:
        os.startfile(outdir)  # open output folder in Explorer
    except Exception:
        pass
    return outdir, ok, len(found)
