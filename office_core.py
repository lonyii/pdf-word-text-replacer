#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Core logic for OfficeToolkit GUI. All long-running functions take a
``log(msg)`` callback so the GUI can display progress. No input()/print()."""
import os
import sys
import glob
import shutil
import hashlib
import subprocess
import datetime
import asyncio
import tempfile

# LAN update source: published release lives on the team share
SHARE_DIR = r"\\Jsztnas\团队文件\内部应用\OfficeToolkit-PDF-Word工具"
SHARE_EXE = os.path.join(SHARE_DIR, "OfficeToolkit.exe")


def _md5(path, chunk=1024 * 1024):
    h = hashlib.md5()
    with open(path, "rb") as f:
        while True:
            b = f.read(chunk)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


def local_exe_path():
    """Path to the running EXE (None when running from source)."""
    exe = getattr(sys, "frozen", False) and sys.executable or None
    return exe


def check_update():
    """Compare local EXE with the LAN share. Returns a status dict."""
    result = {"available": False, "msg": "", "remote_mtime": None}
    if not os.path.isfile(SHARE_EXE):
        result["msg"] = "无法访问局域网更新源：\n" + SHARE_EXE
        return result
    remote_mtime = os.path.getmtime(SHARE_EXE)
    result["remote_mtime"] = remote_mtime
    exe = local_exe_path()
    if not exe or not os.path.isfile(exe):
        result["msg"] = "当前不是 EXE 运行模式，无法自更新（请在打包后的程序中检查）"
        return result
    try:
        if _md5(exe) == _md5(SHARE_EXE):
            result["msg"] = "已是最新版本"
        else:
            result["available"] = True
            result["msg"] = "发现新版本，是否更新？"
    except Exception as e:
        result["msg"] = f"检查更新失败：{type(e).__name__}: {e}"
    return result


def apply_update():
    """Download the share EXE over the local one via a helper batch that runs
    after this process exits (a running EXE cannot be overwritten in place).
    Raises on error; returns True when the update has been scheduled.
    """
    exe = local_exe_path()
    if not exe or not os.path.isfile(SHARE_EXE):
        raise RuntimeError("更新源或本地程序不可用")
    local_dir = os.path.dirname(exe)
    new_exe = os.path.join(local_dir, "_ot_update_new.exe")
    bat = os.path.join(local_dir, "_ot_update.bat")
    shutil.copy2(SHARE_EXE, new_exe)
    # GBK encoding for cmd.exe Chinese support; wait for this process to exit,
    # replace, then relaunch.
    script = (
        "@echo off\r\n"
        "chcp 936 >nul\r\n"
        "timeout /t 3 /nobreak >nul\r\n"
        f'move /Y "{new_exe}" "{exe}"\r\n'
        f'start "" "{exe}"\r\n'
        f'del "%~f0"\r\n'
    )
    with open(bat, "w", encoding="gbk", newline="") as fh:
        fh.write(script)
    subprocess.Popen(["cmd", "/c", bat], cwd=local_dir,
                     creationflags=subprocess.CREATE_NO_WINDOW)
    return True


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


# ---------------------------------------------------------------------------
# Scanned PDF -> searchable double-layer PDF (WeChat OCR via wcocr + PyMuPDF)
# ---------------------------------------------------------------------------
CJK_FONT = r"C:/Windows/Fonts/simhei.ttf"
OCR_DPI = 300
_wcocr = None
_wcocr_ready = False


def _wechat_dirs():
    """Locate WeChatOCR.exe and the dir holding mmmojo_64.dll."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
        candidates = [os.path.join(base, "WeChatOCR"),
                      os.path.join(base, "_internal", "WeChatOCR")]
    else:
        base = os.path.dirname(os.path.abspath(__file__))
        candidates = [os.path.join(base, "vendor", "_old_WeChatOCR_unused"),
                      os.path.join(base, "vendor", "WeChatOCR")]
    for vendor in candidates:
        ocr_exe = os.path.join(vendor, "WeChatOCR", "7079", "extracted", "WeChatOCR.exe")
        if os.path.isfile(ocr_exe):
            return ocr_exe, vendor
    return os.path.join(candidates[0], "WeChatOCR", "7079", "extracted", "WeChatOCR.exe"), candidates[0]


def _init_wcocr(log):
    global _wcocr, _wcocr_ready
    if _wcocr_ready:
        return _wcocr
    if not getattr(sys, "frozen", False):
        vp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor")
        for d in (vp, os.path.join(vp, "_old_WeChatOCR_unused")):
            if d not in sys.path:
                sys.path.append(d)
    import wcocr
    ocr_exe, wechat_dir = _wechat_dirs()
    if not os.path.isfile(ocr_exe):
        raise RuntimeError(f"缺少微信 OCR 组件：{ocr_exe}")
    log("正在启动微信 OCR 引擎…")
    wcocr.init(ocr_exe, wechat_dir)
    _wcocr = wcocr
    _wcocr_ready = True
    return wcocr


def _ocr_one(src, out, wc, log):
    import fitz
    doc = fitz.open(src)
    total = 0
    try:
        for i, page in enumerate(doc, 1):
            pix = page.get_pixmap(dpi=OCR_DPI, colorspace=fitz.csGRAY, alpha=False)
            fd, tmp = tempfile.mkstemp(suffix=".png")
            os.close(fd)
            try:
                with open(tmp, "wb") as fh:
                    fh.write(pix.tobytes("png"))
                result = wc.ocr(tmp)
            finally:
                try:
                    os.remove(tmp)
                except Exception:
                    pass
            if result.get("errcode") not in (0, None):
                raise RuntimeError(f"微信 OCR 错误码 {result.get('errcode')}")
            r = page.rect
            sx, sy = r.width / result["width"], r.height / result["height"]
            items = result.get("ocr_response") or []
            for it in items:
                text = it.get("text", "")
                if not text:
                    continue
                left, top = float(it["left"]), float(it["top"])
                bottom = float(it["bottom"])
                fontsize = max(4.0, (bottom - top) * sy * 0.95)
                try:
                    page.insert_text((left * sx, bottom * sy), text,
                                     fontname="ocr_cjk", fontfile=CJK_FONT,
                                     fontsize=fontsize, render_mode=3, overlay=True)
                except Exception:
                    pass
            total += len(items)
            log(f"  第 {i}/{len(doc)} 页：识别 {len(items)} 行")
        doc.save(out, garbage=3, deflate=True)
    finally:
        doc.close()
    return total


def core_ocr(files, rules, outdir, log):
    """OCR scanned PDFs into double-layer (image + invisible text) PDFs,
    using WeChat OCR. Engine is initialized once and reused."""
    global _wcocr, _wcocr_ready
    os.makedirs(outdir, exist_ok=True)
    ok = 0
    wc = None
    try:
        wc = _init_wcocr(log)
        for f in files:
            name = os.path.splitext(os.path.basename(f))[0] + "_ocr.pdf"
            out = unique_target(outdir, name)
            log(f"正在 OCR：{os.path.basename(f)}")
            try:
                n = _ocr_one(f, out, wc, log)
                log(f"  [OK] -> {name}（文字 {n} 行）")
                ok += 1
            except Exception as e:
                log(f"  [失败] {type(e).__name__}: {e}")
    except Exception as e:
        log(f"  [失败] 引擎加载：{type(e).__name__}: {e}")
    finally:
        if _wcocr_ready and wc is not None:
            try:
                wc.destroy()
            except Exception:
                pass
        _wcocr_ready = False
        _wcocr = None
    log(f"完成：{ok}/{len(files)} 个成功")
    return ok
