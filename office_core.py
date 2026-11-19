#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Core logic for OfficeToolkit GUI. All long-running functions take a
``log(msg)`` callback so the GUI can display progress. No input()/print()."""
import os
import glob
import datetime


def desktop_path():
    try:
        import win32com.shell.shell as shell
        import win32com.shell.shellcon as shellcon
        return shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


def make_output_dir(prefix, custom=""):
    if custom:
        d = custom.strip().strip('"')
    else:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        d = os.path.join(desktop_path(), f"{prefix}_{ts}")
    if d:
        os.makedirs(d, exist_ok=True)
    return d


def unique_target(outdir, name):
    target = os.path.join(outdir, name)
    stem, ext = os.path.splitext(name)
    n = 1
    while os.path.exists(target):
        target = os.path.join(outdir, f"{stem}_{n}{ext}")
        n += 1
    return target


def expand_paths(paths, exts):
    """Expand files/folders into a de-duplicated list, preserving order."""
    found, seen = [], set()
    for p in paths:
        p = p.strip().strip('"')
        if not p:
            continue
        cands = []
        if os.path.isdir(p):
            cands = [f for f in glob.glob(os.path.join(p, "**", "*"), recursive=True)
                     if os.path.isfile(f)]
        elif any(c in p for c in "*?"):
            cands = [f for f in glob.glob(p, recursive=True) if os.path.isfile(f)]
        elif os.path.isfile(p):
            cands = [p]
        for f in cands:
            if os.path.splitext(f)[1].lower() not in exts:
                continue
            rp = os.path.normcase(os.path.abspath(f))
            if rp not in seen:
                seen.add(rp)
                found.append(f)
    return found


# ---------------------------------------------------------------------------
# PDF -> Word via Acrobat (pythonnet in-process .NET reflection; no subprocess)
# ---------------------------------------------------------------------------
_acrobat_ready = None


def _acrobat_convert(pdf, out):
    global _acrobat_ready
    if _acrobat_ready is None:
        try:
            import clr  # noqa: F401
            _acrobat_ready = True
        except Exception:
            _acrobat_ready = False
    if not _acrobat_ready:
        raise RuntimeError("pythonnet (clr) 不可用")

    import System
    from System.Reflection import BindingFlags
    M = BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance

    def arr(items):
        if items is None:
            return None
        a = System.Array.CreateInstance(System.Object, len(items))
        for i, v in enumerate(items):
            a[i] = v
        return a

    def inv(obj, name, args=None):
        return obj.GetType().InvokeMember(name, M, None, obj, arr(args))

    t = System.Type.GetTypeFromProgID("AcroExch.AVDoc")
    if t is None:
        raise RuntimeError("未安装 Acrobat（找不到 AcroExch.AVDoc）")
    av = System.Activator.CreateInstance(t)
    pd = js = None
    try:
        if not inv(av, "Open", [pdf, ""]):
            raise RuntimeError("Acrobat 无法打开该 PDF")
        pd = inv(av, "GetPDDoc")
        js = inv(pd, "GetJSObject")
        if js is None:
            raise RuntimeError("无法获取 JSObject（需要 Acrobat Pro）")
        inv(js, "saveAs", [out, "com.adobe.acrobat.docx"])
        return os.path.exists(out) and os.path.getsize(out) > 0
    finally:
        for obj, meth, args in ((js, "closeDoc", [True]),
                                (pd, "Close", None),
                                (av, "Close", [1])):
            if obj is not None:
                try:
                    inv(obj, meth, args)
                except Exception:
                    pass


def core_pdf2word(files, outdir, log):
    # .NET/COM calls from a worker thread need an initialized COM apartment.
    import pythoncom
    pythoncom.CoInitialize()
    ok = 0
    try:
        for i, f in enumerate(files, 1):
            name = os.path.splitext(os.path.basename(f))[0] + ".docx"
            out = unique_target(outdir, name)
            log(f"[{i}/{len(files)}] 正在转换：{os.path.basename(f)} …")
            try:
                if _acrobat_convert(f, out):
                    log(f"  [OK] -> {name}")
                    ok += 1
                else:
                    log(f"  [失败] Acrobat 未生成输出")
            except Exception as e:
                log(f"  [失败] {type(e).__name__}: {e}")
    finally:
        pythoncom.CoUninitialize()
    log(f"完成：{ok}/{len(files)} 个成功")
    return ok


# ---------------------------------------------------------------------------
# Word find/replace
# ---------------------------------------------------------------------------
def _apply_rules_story(story, rules):
    for rule in rules:
        text = rule.get("find", "")
        if text == "":
            continue
        f = story.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        # Find.Execute needs the full 11 positional args in win32com
        f.Execute(
            text,
            bool(rule.get("match_case", False)),
            bool(rule.get("whole_word", False)),
            bool(rule.get("wildcards", False)),
            False, False, True, 1, False,
            rule.get("replace", ""),
            2,  # wdReplaceAll
        )


def core_replace(files, rules, outdir, log):
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    word = None
    ok = 0
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for src in files:
            doc = None
            try:
                doc = word.Documents.Open(src, False, True, False)
                for story in doc.StoryRanges:
                    _apply_rules_story(story, rules)
                    linked = getattr(story, "NextStoryRange", None)
                    while linked is not None:
                        _apply_rules_story(linked, rules)
                        linked = getattr(linked, "NextStoryRange", None)
                name = os.path.splitext(os.path.basename(src))[0] + ".docx"
                dst = unique_target(outdir, name)
                doc.SaveAs2(dst, 12)
                log(f"[OK] {os.path.basename(src)} -> {name}")
                ok += 1
            except Exception as e:
                log(f"[失败] {os.path.basename(src)}：{type(e).__name__}: {e}")
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
    finally:
        if word is not None:
            try:
                word.Quit()  # only our own DispatchEx instance
            except Exception:
                pass
        pythoncom.CoUninitialize()
    log(f"完成：{ok}/{len(files)} 个成功")
    return ok


# ---------------------------------------------------------------------------
# Word -> PDF
# ---------------------------------------------------------------------------
def core_word2pdf(files, outdir, log):
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    word = None
    ok = 0
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for src in files:
            doc = None
            try:
                doc = word.Documents.Open(src, False, True, False)
                name = os.path.splitext(os.path.basename(src))[0] + ".pdf"
                dst = unique_target(outdir, name)
                doc.ExportAsFixedFormat(dst, 17)  # wdExportFormatPDF
                log(f"[OK] {os.path.basename(src)} -> {name}")
                ok += 1
            except Exception as e:
                log(f"[失败] {os.path.basename(src)}：{type(e).__name__}: {e}")
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    log(f"完成：{ok}/{len(files)} 个成功")
    return ok
