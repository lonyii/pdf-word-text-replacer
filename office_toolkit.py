#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Office Toolkit - PDF/Word batch utilities (menu-driven EXE).

Functions:
  1. PDF -> Word   (via Adobe Acrobat Pro IAC, silent)
  2. Batch find/replace text in Word (.doc/.docx)
  3. Word -> PDF   (via Microsoft Word COM)

All functions support batch input: a file, a folder (recursive), or a glob.
Run the exe directly for an interactive menu, or drop files onto it.
"""
import os
import sys
import glob
import json
import shutil
import tempfile
import subprocess
import datetime

APP_TITLE = "Office Toolkit  v1.0  (PDF/Word batch utilities)"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def resource_path(rel):
    """Path to a bundled resource (works in PyInstaller onefile)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, rel)


def prompt(msg, default=""):
    suffix = f" [{default}]" if default else ""
    try:
        val = input(f"{msg}{suffix}: ").strip().strip('"')
    except EOFError:
        return default
    return val if val else default


def desktop_path():
    try:
        import win32com.shell.shell as shell
        import win32com.shell.shellcon as shellcon
        return shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
    except Exception:
        return os.path.join(os.path.expanduser("~"), "Desktop")


def collect_inputs(raw_list, exts):
    """Expand files/folders/globs into a de-duplicated list."""
    found, seen = [], set()
    for p in raw_list:
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
        else:
            print(f"  [warn] not found: {p}")
        for f in cands:
            if os.path.splitext(f)[1].lower() not in exts:
                continue
            rp = os.path.normcase(os.path.abspath(f))
            if rp not in seen:
                seen.add(rp)
                found.append(f)
    return found


def ask_paths(prompt_text, exts):
    """Prompt for one or more paths (comma separated), supporting drag&drop."""
    raw = prompt(prompt_text)
    if not raw:
        return []
    parts = [x for x in raw.split(",") if x.strip()]
    return collect_inputs(parts, exts)


def ask_out_mode(default_folder_prefix):
    """Return an output directory: desktop timestamp folder, or a custom path."""
    print("  输出位置：")
    print("    1) 桌面新文件夹（按时间戳命名）")
    print("    2) 自定义文件夹")
    choice = prompt("    选择 [1]", "1")
    if choice == "2":
        d = prompt("    输入输出文件夹完整路径")
        if not d:
            print("  [取消] 未指定输出路径")
            return None
    else:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        d = os.path.join(desktop_path(), f"{default_folder_prefix}_{ts}")
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


# ---------------------------------------------------------------------------
# Feature 1: PDF -> Word via Acrobat
# ---------------------------------------------------------------------------
_acrobat_ready = None


def acrobat_convert(pdf, out):
    """Convert one PDF to .docx via Acrobat IAC, using in-process .NET
    reflection (pythonnet). This is the same engine as Acrobat's GUI
    'Export To Word', but silent - and needs NO external powershell process.
    Returns True on success.
    """
    global _acrobat_ready
    if _acrobat_ready is None:
        try:
            import clr  # noqa: F401  (initializes .NET runtime)
            _acrobat_ready = True
        except Exception:
            _acrobat_ready = False
    if not _acrobat_ready:
        raise RuntimeError("pythonnet (clr) unavailable")

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
        raise RuntimeError("Acrobat not installed (AcroExch.AVDoc unavailable)")
    av = System.Activator.CreateInstance(t)
    pd = js = None
    try:
        if not inv(av, "Open", [pdf, ""]):
            raise RuntimeError("Acrobat could not open the PDF")
        pd = inv(av, "GetPDDoc")
        js = inv(pd, "GetJSObject")
        if js is None:
            raise RuntimeError("JSObject unavailable (Acrobat Pro required)")
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


def feature_pdf2word():
    print("\n" + "=" * 60)
    print("功能 1：PDF 转 Word（调用 Acrobat Pro 静默导出）")
    print("=" * 60)
    files = ask_paths("  输入 PDF 文件/文件夹/通配符（多个用逗号分隔）", (".pdf",))
    if not files:
        print("  [取消] 没有找到 PDF 文件")
        return
    print(f"  找到 {len(files)} 个 PDF")
    outdir = ask_out_mode("PDF转Word")
    if not outdir:
        return

    ok = 0
    for f in files:
        name = os.path.splitext(os.path.basename(f))[0] + ".docx"
        out = unique_target(outdir, name)
        try:
            if acrobat_convert(f, out):
                print(f"  [OK] {os.path.basename(f)} -> {name} ({os.path.getsize(out)} bytes)")
                ok += 1
            else:
                print(f"  [失败] {os.path.basename(f)} :: Acrobat 未生成输出文件")
        except Exception as e:
            print(f"  [失败] {os.path.basename(f)} : {type(e).__name__}: {e}")
    print(f"\n  完成：{ok}/{len(files)} 个转换成功")
    print(f"  输出目录：{outdir}")


# ---------------------------------------------------------------------------
# Feature 2: batch find/replace in Word
# ---------------------------------------------------------------------------
def apply_rules_story(story, rules):
    for rule in rules:
        find_text = rule.get("find", "")
        if find_text == "":
            continue
        f = story.Find
        f.ClearFormatting()
        f.Replacement.ClearFormatting()
        f.Execute(
            find_text,
            bool(rule.get("match_case", False)),
            bool(rule.get("whole_word", False)),
            bool(rule.get("wildcards", False)),
            False, False, True, 1, False,
            rule.get("replace", ""),
            2,  # wdReplaceAll
        )


def feature_replace():
    print("\n" + "=" * 60)
    print("功能 2：批量替换 Word 文本（正文/页眉页脚/表格/文本框）")
    print("=" * 60)
    files = ask_paths("  输入 Word 文件/文件夹/通配符（多个用逗号分隔）", (".doc", ".docx"))
    if not files:
        print("  [取消] 没有找到 Word 文件")
        return
    print(f"  找到 {len(files)} 个 Word 文件")

    rules = []
    print("\n  请输入替换规则（find -> replace），输入空行结束：")
    while True:
        find = prompt("    要查找的文本（回车=结束）")
        if find == "":
            break
        repl = prompt("    替换为")
        rules.append({"find": find, "replace": repl})
    if not rules:
        print("  [取消] 没有替换规则")
        return
    print(f"  规则数：{len(rules)}")
    for i, r in enumerate(rules, 1):
        print(f"    {i}. '{r['find']}' -> '{r['replace']}'")

    outdir = ask_out_mode("文本替换")
    if not outdir:
        return

    import win32com.client as win32
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
                    apply_rules_story(story, rules)
                    linked = getattr(story, "NextStoryRange", None)
                    while linked is not None:
                        apply_rules_story(linked, rules)
                        linked = getattr(linked, "NextStoryRange", None)
                name = os.path.splitext(os.path.basename(src))[0] + ".docx"
                dst = unique_target(outdir, name)
                doc.SaveAs2(dst, 12)
                print(f"  [OK] {os.path.basename(src)} -> {name}")
                ok += 1
            except Exception as e:
                print(f"  [失败] {os.path.basename(src)} : {type(e).__name__}: {e}")
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
    print(f"\n  完成：{ok}/{len(files)} 个替换成功")
    print(f"  输出目录：{outdir}")


# ---------------------------------------------------------------------------
# Feature 3: Word -> PDF via Word COM
# ---------------------------------------------------------------------------
def feature_word2pdf():
    print("\n" + "=" * 60)
    print("功能 3：Word 转 PDF（Microsoft Word 高质量导出）")
    print("=" * 60)
    files = ask_paths("  输入 Word 文件/文件夹/通配符（多个用逗号分隔）", (".doc", ".docx"))
    if not files:
        print("  [取消] 没有找到 Word 文件")
        return
    print(f"  找到 {len(files)} 个 Word 文件")
    outdir = ask_out_mode("Word转PDF")
    if not outdir:
        return

    import win32com.client as win32
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
                # wdExportFormatPDF = 17
                doc.ExportAsFixedFormat(dst, 17)
                print(f"  [OK] {os.path.basename(src)} -> {name}")
                ok += 1
            except Exception as e:
                print(f"  [失败] {os.path.basename(src)} : {type(e).__name__}: {e}")
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
    print(f"\n  完成：{ok}/{len(files)} 个转换成功")
    print(f"  输出目录：{outdir}")


# ---------------------------------------------------------------------------
# main menu
# ---------------------------------------------------------------------------
def main():
    print("=" * 60)
    print(APP_TITLE)
    print("=" * 60)
    # Drag & drop support: argv[1] can pre-fill, but menu still guides the user
    if len(sys.argv) > 1:
        dropped = [a for a in sys.argv[1:] if os.path.exists(a)]
        if dropped:
            print("  检测到拖入的文件：")
            for f in dropped:
                print("    -", f)
            print("  （请在对应功能中粘贴/输入路径，或用逗号分隔多个）")

    actions = {
        "1": feature_pdf2word,
        "2": feature_replace,
        "3": feature_word2pdf,
    }
    while True:
        print("\n请选择功能：")
        print("  1. PDF 转 Word（调用 Acrobat）")
        print("  2. 批量替换 Word 文本")
        print("  3. Word 转 PDF")
        print("  0. 退出")
        choice = prompt("  输入选项 [0]", "0")
        if choice == "0":
            print("  再见！")
            break
        fn = actions.get(choice)
        if fn:
            try:
                fn()
            except Exception as e:
                print(f"\n  [错误] 功能执行异常：{type(e).__name__}: {e}")
        else:
            print("  无效选项，请重新输入")
        prompt("\n按回车键继续...")


if __name__ == "__main__":
    main()
