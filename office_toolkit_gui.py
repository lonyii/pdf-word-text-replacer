#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OfficeToolkit - Windows GUI for batch PDF/Word utilities.

Three tabs:
  1. PDF -> Word (via Acrobat Pro, silent)
  2. Batch find/replace in Word
  3. Word -> PDF

All work runs in a background thread; logs are marshalled back to the UI
through a queue. The user never touches a command line.
"""
import os
import sys
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import office_core as core

APP_TITLE = "OfficeToolkit  ·  PDF / Word 批量处理工具"
FONT = ("Microsoft YaHei UI", 10)
FONT_B = ("Microsoft YaHei UI", 10, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 16, "bold")


class TabFrame(ttk.Frame):
    """Base frame holding: file list, output chooser, start button, log."""

    def __init__(self, master, title, file_exts, output_prefix, with_rules=False):
        super().__init__(master, padding=12)
        self.file_exts = file_exts
        self.output_prefix = output_prefix
        self.with_rules = with_rules
        self.files = []            # list of file paths
        self.rules = []            # list of {find, replace}
        self.task = None           # callable(op_log) for the current job
        self.running = False

        # ---- top title ----
        ttk.Label(self, text=title, font=FONT_B).pack(anchor="w", pady=(0, 8))

        # ---- file section ----
        file_box = ttk.LabelFrame(self, text=" ① 选择文件 / 文件夹 ", padding=8)
        file_box.pack(fill="both", expand=True, pady=(0, 8))

        btn_bar = ttk.Frame(file_box)
        btn_bar.pack(fill="x", pady=(0, 6))
        ttk.Button(btn_bar, text="添加文件…", command=self.add_files).pack(side="left", padx=(0, 6))
        ttk.Button(btn_bar, text="添加文件夹…", command=self.add_folder).pack(side="left", padx=(0, 6))
        ttk.Button(btn_bar, text="清空列表", command=self.clear_files).pack(side="left")

        list_frame = ttk.Frame(file_box)
        list_frame.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(list_frame, font=FONT, selectmode="extended",
                                  activestyle="dotbox")
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # ---- rules section (replace tab only) ----
        if with_rules:
            rule_box = ttk.LabelFrame(self, text=" ② 替换规则（查找 → 替换） ", padding=8)
            rule_box.pack(fill="both", expand=True, pady=(0, 8))
            rb_top = ttk.Frame(rule_box)
            rb_top.pack(fill="x", pady=(0, 6))
            self.find_var = tk.StringVar()
            self.repl_var = tk.StringVar()
            ttk.Label(rb_top, text="查找：").pack(side="left")
            self.find_entry = ttk.Entry(rb_top, textvariable=self.find_var, width=24)
            self.find_entry.pack(side="left", padx=(0, 8))
            ttk.Label(rb_top, text="替换为：").pack(side="left")
            self.repl_entry = ttk.Entry(rb_top, textvariable=self.repl_var, width=24)
            self.repl_entry.pack(side="left", padx=(0, 8))
            ttk.Button(rb_top, text="添加规则", command=self.add_rule).pack(side="left", padx=(0, 6))
            ttk.Button(rb_top, text="删除选中", command=self.del_rule).pack(side="left")
            rf = ttk.Frame(rule_box)
            rf.pack(fill="both", expand=True)
            self.rule_tree = ttk.Treeview(rf, columns=("find", "replace"),
                                          show="headings", height=5)
            self.rule_tree.heading("find", text="查找")
            self.rule_tree.heading("replace", text="替换为")
            self.rule_tree.column("find", width=260)
            self.rule_tree.column("replace", width=260)
            rsb = ttk.Scrollbar(rf, orient="vertical", command=self.rule_tree.yview)
            self.rule_tree.configure(yscrollcommand=rsb.set)
            self.rule_tree.pack(side="left", fill="both", expand=True)
            rsb.pack(side="right", fill="y")

        # ---- output section ----
        step = "③" if with_rules else "②"
        out_box = ttk.LabelFrame(self, text=f" {step} 输出位置 ", padding=8)
        out_box.pack(fill="x", pady=(0, 8))
        self.out_mode = tk.StringVar(value="desktop")
        ttk.Radiobutton(out_box, text="桌面新建文件夹（按时间戳命名）",
                        variable=self.out_mode, value="desktop",
                        command=self.toggle_out).pack(anchor="w")
        row = ttk.Frame(out_box)
        row.pack(fill="x", pady=(4, 0))
        ttk.Radiobutton(row, text="自定义文件夹：", variable=self.out_mode,
                        value="custom", command=self.toggle_out).pack(side="left")
        self.out_var = tk.StringVar()
        self.out_entry = ttk.Entry(row, textvariable=self.out_var, state="disabled")
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(4, 6))
        self.out_btn = ttk.Button(row, text="浏览…", command=self.pick_out,
                                  state="disabled")
        self.out_btn.pack(side="left")

        # ---- start button + progress ----
        step = "④" if with_rules else "③"
        action = ttk.Frame(self)
        action.pack(fill="x", pady=(0, 8))
        self.start_btn = ttk.Button(action, text=f"{step} 开始处理",
                                    command=self.on_start)
        self.start_btn.pack(side="left")
        self.progress = ttk.Progressbar(action, mode="indeterminate", length=200)
        self.progress.pack(side="left", padx=(12, 0))

        # ---- log ----
        log_box = ttk.LabelFrame(self, text=" 处理日志 ", padding=6)
        log_box.pack(fill="both", expand=True)
        lf = ttk.Frame(log_box)
        lf.pack(fill="both", expand=True)
        self.log_text = tk.Text(lf, height=8, font=("Consolas", 9),
                                state="disabled", wrap="none")
        lsb = ttk.Scrollbar(lf, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=lsb.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        lsb.pack(side="right", fill="y")

    # ---------- file list ----------
    def add_files(self):
        ftypes = []
        for e in self.file_exts:
            ftypes.append((f"{e} 文件", f"*{e}"))
        ftypes.append(("所有文件", "*.*"))
        paths = filedialog.askopenfilenames(title="选择文件", filetypes=ftypes)
        self._append_files(paths)

    def add_folder(self):
        d = filedialog.askdirectory(title="选择文件夹（含子文件夹）")
        if d:
            self._append_files([d])

    def _append_files(self, paths):
        added = 0
        for p in paths:
            files = core.expand_paths([p], self.file_exts)
            for f in files:
                if f not in self.files:
                    self.files.append(f)
                    self.listbox.insert("end", f)
                    added += 1
        if added:
            self._log(f"已添加 {added} 个文件（支持格式：{'  '.join(self.file_exts)}）")

    def clear_files(self):
        self.files.clear()
        self.listbox.delete(0, "end")
        self._log("已清空文件列表")

    # ---------- rules ----------
    def add_rule(self):
        f = self.find_var.get().strip()
        r = self.repl_var.get()
        if not f:
            messagebox.showwarning("提示", "请填写要查找的文本", parent=self)
            return
        self.rules.append({"find": f, "replace": r})
        self.rule_tree.insert("", "end", values=(f, r))
        self.find_var.set("")
        self.repl_var.set("")
        self.find_entry.focus_set()
        self._log(f"已添加规则：'{f}' → '{r}'")

    def del_rule(self):
        sel = self.rule_tree.selection()
        if not sel:
            return
        for item in sel:
            idx = self.rule_tree.index(item)
            self.rule_tree.delete(item)
            if 0 <= idx < len(self.rules):
                self.rules.pop(idx)

    # ---------- output ----------
    def toggle_out(self):
        custom = self.out_mode.get() == "custom"
        self.out_entry.configure(state="normal" if custom else "disabled")
        self.out_btn.configure(state="normal" if custom else "disabled")

    def pick_out(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.out_var.set(d)

    # ---------- logging ----------
    def _log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", str(msg) + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    # ---------- start ----------
    def on_start(self):
        if self.running:
            return
        if not self.files:
            messagebox.showwarning("提示", "请先添加要处理的文件", parent=self)
            return
        if self.with_rules and not self.rules:
            messagebox.showwarning("提示", "请至少添加一条替换规则", parent=self)
            return
        custom = self.out_var.get() if self.out_mode.get() == "custom" else ""
        try:
            outdir = core.make_output_dir(self.output_prefix, custom)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出文件夹：{e}", parent=self)
            return

        self._log(f"开始处理 {len(self.files)} 个文件，输出：{outdir}")
        self.running = True
        self.start_btn.configure(state="disabled", text="处理中…")
        self.progress.start(12)

        files = list(self.files)
        rules = list(self.rules) if self.with_rules else None

        def worker(q, task, files, rules, outdir):
            def op_log(m):
                q.put(("log", str(m)))
            try:
                task(files, rules, outdir, op_log)
            except Exception:
                q.put(("log", "发生异常：\n" + traceback.format_exc()))
            finally:
                q.put(("done", None))

        # NB: self.master is the Notebook, not the Tk root. The queue lives on
        # the App root - use winfo_toplevel() to reach it from any widget.
        app_queue = self.winfo_toplevel().q
        t = threading.Thread(target=worker,
                             args=(app_queue, self.task, files, rules, outdir),
                             daemon=True)
        t.start()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("860x680")
        self.minsize(760, 560)
        self.q = queue.Queue()

        style = ttk.Style(self)
        try:
            style.theme_use("vista")
        except Exception:
            pass

        # header
        header = ttk.Frame(self, padding=(14, 12, 14, 4))
        header.pack(fill="x")
        ttk.Label(header, text="OfficeToolkit", font=FONT_TITLE).pack(side="left")
        ttk.Label(header, text="  批量 PDF / Word 处理  ·  无需命令行",
                  font=FONT).pack(side="left", pady=(8, 0))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=8)

        self.pdf_tab = TabFrame(nb, "PDF 转 Word（调用 Acrobat Pro）",
                                (".pdf",), "PDF转Word")
        self.rep_tab = TabFrame(nb, "批量替换 Word 文本（正文 / 页眉页脚 / 表格 / 文本框）",
                                (".docx", ".doc"), "文本替换", with_rules=True)
        self.w2p_tab = TabFrame(nb, "Word 转 PDF",
                                (".docx", ".doc"), "Word转PDF")

        self.pdf_tab.task = self._task_pdf2word
        self.rep_tab.task = self._task_replace
        self.w2p_tab.task = self._task_word2pdf

        nb.add(self.pdf_tab, text="  PDF 转 Word  ")
        nb.add(self.rep_tab, text="  批量替换文本  ")
        nb.add(self.w2p_tab, text="  Word 转 PDF  ")

        ttk.Label(self, text="提示：处理前请关闭正在编辑的 Word / Acrobat 文件；中文路径与文本均支持。",
                  font=("Microsoft YaHei UI", 9), foreground="#666").pack(
                      anchor="w", padx=16, pady=(0, 10))

        self.after(120, self._poll_queue)

    # task adapters: core signatures differ (replace takes rules)
    def _task_pdf2word(self, files, rules, outdir, log):
        core.core_pdf2word(files, outdir, log)

    def _task_replace(self, files, rules, outdir, log):
        core.core_replace(files, rules, outdir, log)

    def _task_word2pdf(self, files, rules, outdir, log):
        core.core_word2pdf(files, outdir, log)

    def _poll_queue(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "log":
                    for tab in (self.pdf_tab, self.rep_tab, self.w2p_tab):
                        if tab.running:
                            tab._log(val)
                            break
                elif kind == "done":
                    for tab in (self.pdf_tab, self.rep_tab, self.w2p_tab):
                        if tab.running:
                            tab.running = False
                            tab.start_btn.configure(state="normal", text="开始处理")
                            tab.progress.stop()
                            tab._log("— 处理结束 —")
                            break
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)


def _selftest():
    """Hidden mode: run a real PDF->Word conversion, write log to a file.
    Usage: OfficeToolkit.exe --selftest <pdf> <outdir>"""
    import traceback
    args = sys.argv[2:]
    pdf = args[0] if len(args) > 0 else r"D:\Documents\桌面\新建文件夹 (8)\NC26E444-b.线上资料.pdf"
    outdir = args[1] if len(args) > 1 else os.path.join(os.environ.get("TEMP", "C:\\Temp"), "ot_selftest")
    logfile = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "ot_selftest_log.txt")
    lines = []
    def log(m):
        lines.append(str(m))
        with open(logfile, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines))
    try:
        os.makedirs(outdir, exist_ok=True)
        log("SELFTEST START")
        core.core_pdf2word([pdf], outdir, log)
        log("SELFTEST DONE files=" + repr(os.listdir(outdir)))
    except Exception:
        log("SELFTEST ERR\n" + traceback.format_exc())


def main():
    if "--selftest" in sys.argv:
        _selftest()
        return
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
