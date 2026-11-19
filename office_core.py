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
import threading


class JobQueue:
    """Thread-safe work queue shared between the GUI and a worker thread.

    The GUI can add files while processing, and remove files that are either
    still pending or currently being processed. Removing the active file marks
    it cancelled: the worker finishes the current (non-interruptible) COM / OCR
    call, discards its result, and moves on to the next file.
    """
    def __init__(self, files):
        self._lock = threading.Lock()
        self._pending = list(files)
        self._cancelled = set()
        self._current = None
        self._closed = False           # no more files will be added
        self._empty_since = None
        self.idle_timeout = 0.0        # >0: auto-finish after idle seconds

    def take(self):
        """Return a file to process, None if busy/empty. Caller should loop."""
        with self._lock:
            if self._current is not None:
                return None            # worker still busy
            if not self._pending:
                return None
            self._current = self._pending.pop(0)
            return self._current

    def is_done(self, idle_timeout=0.0):
        """True when closed and drained. If idle_timeout > 0, also finish
        after the queue has been empty for that many seconds (lets late adds
        arrive during processing without hanging forever)."""
        import time
        with self._lock:
            if self._closed and not self._pending and self._current is None:
                return True
            if idle_timeout > 0 and not self._pending and self._current is None:
                if self._empty_since is None:
                    self._empty_since = time.time()
                elif time.time() - self._empty_since >= idle_timeout:
                    return True
            else:
                self._empty_since = None
            return False

    def close(self):
        with self._lock:
            self._closed = True

    def release(self):
        with self._lock:
            self._current = None

    def is_cancelled(self, path):
        with self._lock:
            return path in self._cancelled

    def add(self, paths):
        added = 0
        with self._lock:
            for p in paths:
                if p not in self._pending and p != self._current and p not in self._cancelled:
                    self._pending.append(p)
                    added += 1
        return added

    def remove(self, paths):
        """Remove pending files; mark the active one cancelled."""
        removed, cancelled = 0, 0
        with self._lock:
            for p in paths:
                if p in self._pending:
                    self._pending.remove(p)
                    removed += 1
                elif p == self._current:
                    self._cancelled.add(p)
                    cancelled += 1
        return removed, cancelled

    def pending_count(self):
        with self._lock:
            return len(self._pending)

    def snapshot(self):
        with self._lock:
            return list(self._pending), self._current, set(self._cancelled)


def iter_job(job, files=None):
    """Yield files one at a time from a JobQueue, or from a plain list.
    For the queue case, waits for new additions until job.close() and drained.
    Yields (file, cancelled) so callers can skip removed-in-flight files.
    """
    import time
    if job is None:
        for f in (files or []):
            yield f, False
        return
    while True:
        f = job.take()
        if f is not None:
            cancelled = job.is_cancelled(f)
            try:
                yield f, cancelled
            finally:
                job.release()
            continue
        if job.is_done(job.idle_timeout):
            break
        time.sleep(0.1)


# LAN update source: published release is a ZIP archive on the team share
SHARE_DIR = r"\\Jsztnas\团队文件\内部应用\OfficeToolkit-PDF-Word工具"
SHARE_ZIP = os.path.join(SHARE_DIR, "OfficeToolkit.zip")
VERSION_FILE = "version.txt"


def _version_tuple(text):
    """'2026.11.19.1' -> (2026, 11, 19, 1); invalid -> (0,)."""
    s = (text or "").strip().lstrip("vV")
    try:
        parts = tuple(int(x) for x in s.replace("-", ".").split(".") if x != "")
        return parts or (0,)
    except Exception:
        return (0,)


def _read_local_version():
    exe = local_exe_path()
    if not exe:
        return None
    p = os.path.join(os.path.dirname(exe), VERSION_FILE)
    try:
        with open(p, "r", encoding="utf-8-sig") as fh:
            return fh.read().strip()
    except Exception:
        return None


def _read_zip_version(zip_path):
    import zipfile
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            names = [n for n in zf.namelist()
                     if os.path.basename(n).lower() == VERSION_FILE.lower()]
            if not names:
                return None
            return zf.read(names[0]).decode("utf-8-sig").strip()
    except Exception:
        return None


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
    """Compare local version.txt with the one inside the share ZIP.
    Returns {"available", "msg", "local_version", "remote_version"}.
    """
    result = {"available": False, "msg": "", "local_version": None,
              "remote_version": None}
    if not os.path.isfile(SHARE_ZIP):
        result["msg"] = "无法访问局域网更新源：\n" + SHARE_ZIP
        return result
    exe = local_exe_path()
    if not exe or not os.path.isfile(exe):
        result["msg"] = "当前不是 EXE 运行模式，无法自更新（请在打包后的程序中检查）"
        return result
    local_v = _read_local_version()
    remote_v = _read_zip_version(SHARE_ZIP)
    result["local_version"] = local_v
    result["remote_version"] = remote_v
    if not remote_v:
        result["msg"] = "更新包缺少版本信息，已跳过更新"
        return result
    if local_v and _version_tuple(local_v) >= _version_tuple(remote_v):
        result["msg"] = f"已是最新版本（{local_v}）"
        return result
    result["available"] = True
    result["msg"] = f"发现新版本：{local_v or '未知'} → {remote_v}，是否更新？"
    return result


def apply_update():
    """Copy the share ZIP locally, then run a helper batch after this process
    exits that extracts it over the program directory and relaunches the app.
    """
    exe = local_exe_path()
    if not exe or not os.path.isfile(SHARE_ZIP):
        raise RuntimeError("更新源或本地程序不可用")
    local_dir = os.path.dirname(exe)
    zip_local = os.path.join(local_dir, "_ot_update.zip")
    bat = os.path.join(local_dir, "_ot_update.bat")
    shutil.copy2(SHARE_ZIP, zip_local)
    # Expand-Archive -Force overwrites files; PowerShell handles Chinese names
    # and onedir's many subfolders. Wait for this process to exit first.
    # NOTE: cmd.exe on Chinese Windows defaults to code page 936 (GBK).
    # Keep the batch file GBK-encoded and do NOT switch code page, otherwise
    # Chinese folder names in the paths get mangled (UTF-8 bytes read as GBK).
    # Use %~dp0 (the batch file's own directory) instead of hard-coded paths,
    # so Chinese folder names never need to be encoded inside the .bat file.
    script = (
        "@echo off\r\n"
        "timeout /t 3 /nobreak >nul\r\n"
        "powershell -NoProfile -Command "
        "\"Expand-Archive -LiteralPath (Join-Path $pwd '_ot_update.zip') "
        "-DestinationPath $pwd -Force\"\r\n"
        'del /F /Q "%~dp0_ot_update.zip"\r\n'
        'start "" "%~dp0OfficeToolkit.exe"\r\n'
        'del "%~f0"\r\n'
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


def fix_docx_header_overlap(docx_path):
    """Fix two Acrobat-export issues in converted DOCX:
    1) Large headings sit at the same vertical position as floating header
       objects (logo / form number), so Word clips their top. Add space-before.
    2) Tiny page margins (< 20pt) make Word warn "margins are very small"
       on open/print. Raise them to a safe minimum.
    """
    from docx import Document
    from docx.shared import Pt
    doc = Document(docx_path)

    # --- raise too-small margins to avoid Word's "margins very small" prompt
    min_margin = Pt(20)
    for sec in doc.sections:
        for attr in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
            v = getattr(sec, attr)
            if v is not None and v < min_margin:
                setattr(sec, attr, min_margin)

    def eff_size(p):
        szs = [r.font.size.pt for r in p.runs if r.font.size is not None]
        if szs:
            return max(szs)
        try:
            if p.style.font.size is not None:
                return p.style.font.size.pt
        except Exception:
            pass
        return None

    def iter_paras(d):
        for p in d.paragraphs:
            yield p
        for t in d.tables:
            for row in t.rows:
                for c in row.cells:
                    for p in c.paragraphs:
                        yield p

    n = 0
    for p in iter_paras(doc):
        if not p.text.strip():
            continue
        if p.paragraph_format.space_before is not None:
            continue
        sz = eff_size(p)
        if sz is not None and sz >= 16:
            p.paragraph_format.space_before = Pt(10)
            n += 1
    if n:
        doc.save(docx_path)
    return n


def core_pdf2word(files, outdir, log, job=None):
    # .NET/COM calls from a worker thread need an initialized COM apartment.
    import pythoncom
    pythoncom.CoInitialize()
    ok = 0
    processed = 0
    try:
        for f, cancelled in iter_job(job, files):
            processed += 1
            name = os.path.splitext(os.path.basename(f))[0] + ".docx"
            out = unique_target(outdir, name)
            log(f"[{processed}] 正在转换：{os.path.basename(f)} …")
            if cancelled:
                log(f"  [已取消] {os.path.basename(f)}")
                continue
            result = None
            try:
                result = _acrobat_convert(f, out)
            except Exception as e:
                result = e
            # Re-check cancellation after the (non-interruptible) conversion:
            # deleting the active file marks it cancelled mid-flight.
            if job is not None and job.is_cancelled(f):
                log(f"  [已取消] {os.path.basename(f)}（已中断，不计入结果）")
                if os.path.exists(out):
                    try:
                        os.remove(out)
                    except Exception:
                        pass
            elif isinstance(result, Exception):
                log(f"  [失败] {type(result).__name__}: {result}")
            elif result:
                try:
                    fix_docx_header_overlap(out)
                except Exception:
                    pass
                log(f"  [OK] -> {name}")
                ok += 1
            else:
                log(f"  [失败] Acrobat 未生成输出")
    finally:
        pythoncom.CoUninitialize()
    log(f"完成：{ok}/{processed} 个成功")
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


def core_replace(files, rules, outdir, log, job=None):
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    word = None
    ok = 0
    processed = 0
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for src, cancelled in iter_job(job, files):
            processed += 1
            if cancelled:
                log(f"[已取消] {os.path.basename(src)}")
                continue
            doc = None
            name = os.path.splitext(os.path.basename(src))[0] + ".docx"
            dst = unique_target(outdir, name)
            err = None
            try:
                doc = word.Documents.Open(src, False, True, False)
                for story in doc.StoryRanges:
                    _apply_rules_story(story, rules)
                    linked = getattr(story, "NextStoryRange", None)
                    while linked is not None:
                        _apply_rules_story(linked, rules)
                        linked = getattr(linked, "NextStoryRange", None)
                doc.SaveAs2(dst, 12)
            except Exception as e:
                err = e
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
            if job is not None and job.is_cancelled(src):
                log(f"[已取消] {os.path.basename(src)}（已中断，不计入结果）")
                if os.path.exists(dst):
                    try:
                        os.remove(dst)
                    except Exception:
                        pass
            elif err is not None:
                log(f"[失败] {os.path.basename(src)}：{type(err).__name__}: {err}")
            else:
                log(f"[OK] {os.path.basename(src)} -> {name}")
                ok += 1
    finally:
        if word is not None:
            try:
                word.Quit()  # only our own DispatchEx instance
            except Exception:
                pass
        pythoncom.CoUninitialize()
    log(f"完成：{ok}/{processed} 个成功")
    return ok


# ---------------------------------------------------------------------------
# Word -> PDF
# ---------------------------------------------------------------------------
def core_word2pdf(files, outdir, log, job=None):
    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    word = None
    ok = 0
    processed = 0
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for src, cancelled in iter_job(job, files):
            processed += 1
            if cancelled:
                log(f"[已取消] {os.path.basename(src)}")
                continue
            doc = None
            name = os.path.splitext(os.path.basename(src))[0] + ".pdf"
            dst = unique_target(outdir, name)
            err = None
            try:
                doc = word.Documents.Open(src, False, True, False)
                doc.ExportAsFixedFormat(dst, 17)  # wdExportFormatPDF
            except Exception as e:
                err = e
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
            if job is not None and job.is_cancelled(src):
                log(f"[已取消] {os.path.basename(src)}（已中断，不计入结果）")
                if os.path.exists(dst):
                    try:
                        os.remove(dst)
                    except Exception:
                        pass
            elif err is not None:
                log(f"[失败] {os.path.basename(src)}：{type(err).__name__}: {err}")
            else:
                log(f"[OK] {os.path.basename(src)} -> {name}")
                ok += 1
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    log(f"完成：{ok}/{processed} 个成功")
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


def core_ocr(files, rules, outdir, log, job=None):
    """OCR scanned PDFs into double-layer (image + invisible text) PDFs,
    using WeChat OCR. Engine is initialized once and reused."""
    global _wcocr, _wcocr_ready
    os.makedirs(outdir, exist_ok=True)
    ok = 0
    processed = 0
    wc = None
    try:
        wc = _init_wcocr(log)
        for f, cancelled in iter_job(job, files):
            processed += 1
            name = os.path.splitext(os.path.basename(f))[0] + "_ocr.pdf"
            out = unique_target(outdir, name)
            log(f"正在 OCR：{os.path.basename(f)}")
            if cancelled:
                log(f"  [已取消] {os.path.basename(f)}")
                continue
            err = None
            try:
                _ocr_one(f, out, wc, log)
            except Exception as e:
                err = e
            if job is not None and job.is_cancelled(f):
                log(f"  [已取消] {os.path.basename(f)}（已中断，不计入结果）")
                if os.path.exists(out):
                    try:
                        os.remove(out)
                    except Exception:
                        pass
            elif err is not None:
                log(f"  [失败] {type(err).__name__}: {err}")
            else:
                log(f"  [OK] -> {name}")
                ok += 1
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
    log(f"完成：{ok}/{processed} 个成功")
    return ok
