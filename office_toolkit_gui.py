#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""OfficeToolkit - Windows GUI for batch PDF/Word utilities.

Three tabs share ONE bottom control bar (output + start + log), so content
can never be pushed off-screen. The user never touches a command line.
"""
import os
import sys
import json
import queue
import subprocess
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from tkinterdnd2 import TkinterDnD
    _DND = True
except Exception:
    _DND = False

import office_core as core
import core_batch

APP_TITLE = "OfficeToolkit  ·  PDF / Word 批量处理工具"
FONT = ("Microsoft YaHei UI", 10)
FONT_B = ("Microsoft YaHei UI", 10, "bold")
FONT_TITLE = ("Microsoft YaHei UI", 15, "bold")


# ===========================================================================
# Tab page: only the parts unique to each function (file list, and rules)
# ===========================================================================
class TabPage(ttk.Frame):
    def __init__(self, master, file_exts, with_rules=False):
        super().__init__(master, padding=(10, 8))
        self.file_exts = file_exts
        self.with_rules = with_rules
        self.files = []
        self.rules = []
        self.output_prefix = ""

        self.columnconfigure(0, weight=1)
        row = 0

        # ---- file list ----
        frm = ttk.LabelFrame(self, text=" 文件 / 文件夹 ", padding=6)
        frm.grid(row=row, column=0, sticky="nsew")
        self.rowconfigure(row, weight=1)
        row += 1
        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(1, weight=1)

        bar = ttk.Frame(frm)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        ttk.Button(bar, text="添加文件", command=self.add_files, width=11).pack(side="left", padx=(0, 5))
        ttk.Button(bar, text="添加文件夹", command=self.add_folder, width=11).pack(side="left", padx=(0, 5))
        ttk.Button(bar, text="清空", command=self.clear_files, width=7).pack(side="left")

        lf = ttk.Frame(frm)
        lf.grid(row=1, column=0, sticky="nsew")
        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(0, weight=1)
        self.listbox = tk.Listbox(lf, font=FONT, selectmode="extended",
                                  activestyle="dotbox", height=4)
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self._setup_listbox_bindings()

        # ---- rules (replace tab only) ----
        if with_rules:
            rf = ttk.LabelFrame(self, text=" 替换规则（查找 → 替换为） ", padding=6)
            rf.grid(row=row, column=0, sticky="nsew", pady=(8, 0))
            self.rowconfigure(row, weight=1)
            row += 1
            rf.columnconfigure(1, weight=1)
            rf.columnconfigure(3, weight=1)
            rf.rowconfigure(1, weight=1)   # let the rule table grow with the window

            ttk.Label(rf, text="查找：").grid(row=0, column=0, sticky="w")
            self.find_var = tk.StringVar()
            ttk.Entry(rf, textvariable=self.find_var).grid(row=0, column=1, sticky="ew", padx=(2, 8))
            ttk.Label(rf, text="替换为：").grid(row=0, column=2, sticky="w")
            self.repl_var = tk.StringVar()
            ttk.Entry(rf, textvariable=self.repl_var).grid(row=0, column=3, sticky="ew", padx=(2, 8))
            ttk.Button(rf, text="+ 添加", command=self.add_rule, width=9).grid(row=0, column=4)

            tf = ttk.Frame(rf)
            tf.grid(row=1, column=0, columnspan=5, sticky="nsew", pady=(6, 0))
            tf.columnconfigure(0, weight=1)
            tf.rowconfigure(0, weight=1)
            self.rule_tree = ttk.Treeview(tf, columns=("find", "replace"),
                                          show="headings", height=3)
            self.rule_tree.heading("find", text="查找")
            self.rule_tree.heading("replace", text="替换为")
            self.rule_tree.column("find", width=300, stretch=True)
            self.rule_tree.column("replace", width=300, stretch=True)
            rsb = ttk.Scrollbar(tf, orient="vertical", command=self.rule_tree.yview)
            self.rule_tree.configure(yscrollcommand=rsb.set)
            self.rule_tree.grid(row=0, column=0, sticky="nsew")
            rsb.grid(row=0, column=1, sticky="ns")
            self.rule_tree.bind("<Delete>", lambda e: self.del_rule())
            self.rule_tree.bind("<Button-3>", self._show_rule_menu)

    # ---------- listbox interactions ----------
    def _setup_listbox_bindings(self):
        lb = self.listbox
        lb.bind("<Control-v>", lambda e: self._on_paste())
        lb.bind("<Delete>", lambda e: self._delete_selected())
        lb.bind("<Button-3>", self._show_context_menu)
        if _DND:
            try:
                lb.drop_target_register("DND_Files")
                lb.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:
                pass

    @staticmethod
    def _parse_dnd(data):
        """Parse tkinterdnd2 file drop payload -> list of paths."""
        paths, buf, in_brace = [], "", False
        for ch in data:
            if ch == "{":
                in_brace = True
            elif ch == "}":
                in_brace = False
                if buf:
                    paths.append(buf); buf = ""
            elif ch.isspace() and not in_brace:
                if buf:
                    paths.append(buf); buf = ""
            else:
                buf += ch
        if buf:
            paths.append(buf)
        return paths

    def _on_drop(self, event):
        self._append(self._parse_dnd(event.data))

    def _on_paste(self):
        try:
            data = self.clipboard_get()
        except Exception:
            return
        # Windows Explorer copies files as newline-separated paths (or CF_HDROP
        # rendered to text). Accept each non-empty line as a path.
        paths = [ln.strip().strip('"') for ln in data.splitlines() if ln.strip()]
        if paths:
            self._append(paths)

    def _show_context_menu(self, event):
        selected = set(self.listbox.curselection())
        i = self.listbox.nearest(event.y)
        # Only collapse to a single selection if the right-clicked row was not
        # already selected; otherwise preserve the user's multi-selection.
        if 0 <= i < self.listbox.size() and i not in selected:
            self.listbox.selection_clear(0, "end")
            self.listbox.selection_set(i)
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="添加文件…", command=self.add_files)
        menu.add_command(label="添加文件夹…", command=self.add_folder)
        if self.listbox.curselection():
            menu.add_separator()
            menu.add_command(label="删除选中", command=self._delete_selected)
        menu.add_separator()
        menu.add_command(label="清空列表", command=self.clear_files)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _delete_selected(self):
        sel = list(self.listbox.curselection())
        if not sel:
            return
        paths = [self.files[i] for i in sel if 0 <= i < len(self.files)]
        for i in reversed(sel):
            self.listbox.delete(i)
        for p in paths:
            if p in self.files:
                self.files.remove(p)
        app = self.winfo_toplevel()
        if hasattr(app, "notify_files_removed"):
            app.notify_files_removed(self, paths)
        self.log(f"已删除 {len(paths)} 项")

    # ---------- file helpers ----------
    def add_files(self):
        ftypes = [(f"{e} 文件", f"*{e}") for e in self.file_exts]
        ftypes.append(("所有文件", "*.*"))
        self._append(filedialog.askopenfilenames(title="选择文件", filetypes=ftypes))

    def add_folder(self):
        d = filedialog.askdirectory(title="选择文件夹（含子文件夹）")
        if d:
            self._append([d])

    def _append(self, paths):
        added = []
        for p in paths:
            for f in core.expand_paths([p], self.file_exts):
                if f not in self.files:
                    self.files.append(f)
                    self.listbox.insert("end", f)
                    added.append(f)
        if added:
            app = self.winfo_toplevel()
            if hasattr(app, "notify_files_added"):
                app.notify_files_added(self, added)
            self.log(f"已添加 {len(added)} 个文件")

    def clear_files(self):
        old = list(self.files)
        self.files.clear()
        self.listbox.delete(0, "end")
        app = self.winfo_toplevel()
        if hasattr(app, "notify_files_removed"):
            app.notify_files_removed(self, old)
        self.log("已清空文件列表")

    # ---------- rule helpers ----------
    def add_rule(self):
        f, r = self.find_var.get().strip(), self.repl_var.get()
        if not f:
            messagebox.showwarning("提示", "请填写要查找的文本", parent=self)
            return
        self.rules.append({"find": f, "replace": r})
        self.rule_tree.insert("", "end", values=(f, r))
        self.find_var.set("")
        self.repl_var.set("")
        self.log(f"规则：'{f}' → '{r}'")

    def del_rule(self):
        for item in self.rule_tree.selection():
            i = self.rule_tree.index(item)
            self.rule_tree.delete(item)
            if 0 <= i < len(self.rules):
                self.rules.pop(i)

    def clear_rules(self):
        self.rules.clear()
        for item in self.rule_tree.get_children():
            self.rule_tree.delete(item)

    def _show_rule_menu(self, event):
        i = self.rule_tree.identify_row(event.y)
        current = set(self.rule_tree.selection())
        if i and i not in current:
            self.rule_tree.selection_set(i)
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="添加规则…", command=self._focus_rule_input)
        if self.rule_tree.selection():
            menu.add_separator()
            menu.add_command(label="删除选中", command=self.del_rule)
        if self.rule_tree.get_children():
            menu.add_separator()
            menu.add_command(label="清空全部规则", command=self.clear_rules)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _focus_rule_input(self):
        self.find_var.set("")
        self.repl_var.set("")
        # focus the "查找" entry so the user can type a new rule immediately
        for w in self.winfo_children():
            for c in w.winfo_children():
                if isinstance(c, ttk.Entry):
                    c.focus_set()
                    return

    def log(self, msg):
        self.winfo_toplevel().push_log(str(msg))


# ===========================================================================
# Batch tab: codes -> Everything -> PDF2Word -> replace -> open output
# ===========================================================================
class BatchPage(ttk.Frame):
    """Left: code input + search-import button. Right: file list (reuses
    TabPage's list with drag/drop, paste, context menu, Delete). Config is
    embedded and hidden; right-click the code box to release it."""
    def __init__(self, master):
        super().__init__(master, padding=(10, 8))
        self.output_prefix = "批处理"
        self.with_rules = False
        self.columnconfigure(0, weight=0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        # ---- left: codes + search/import button ----
        left = ttk.LabelFrame(self, text=" 工程编码（一行一个） ", padding=6)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        ttk.Label(left, text="每行一个编码，点「搜索导入」将结果加入右侧列表。"
                  ).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.code_text = tk.Text(left, font=FONT, height=8, width=16, wrap="none")
        csb = ttk.Scrollbar(left, orient="vertical", command=self.code_text.yview)
        self.code_text.configure(yscrollcommand=csb.set)
        self.code_text.grid(row=1, column=0, sticky="nsew")
        csb.grid(row=1, column=1, sticky="ns")
        self.code_text.bind("<Button-3>", self._show_menu)
        ttk.Button(left, text="搜索导入", command=self.search_import, width=14
                   ).grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        # ---- right: file list (same interactions as other tabs) ----
        self.file_page = TabPage(self, (".pdf",))
        self.file_page.grid(row=0, column=1, sticky="nsew")
        self.files = self.file_page.files   # alias for common start/queue logic

    def get_codes(self):
        return [ln.strip() for ln in self.code_text.get("1.0", "end").splitlines()
                if ln.strip()]

    def search_import(self, _=None):
        codes = self.get_codes()
        if not codes:
            messagebox.showwarning("提示", "请先输入工程编码（一行一个）", parent=self)
            return
        cfg, _ = core_batch.load_config()
        if not os.path.isfile(core_batch.resolve_es(cfg)):
            messagebox.showerror("错误", "找不到 Everything 命令行工具 es.exe", parent=self)
            return
        app = self.winfo_toplevel()
        app.push_log(f"正在搜索 {len(codes)} 个编码…")

        def worker():
            added = []
            try:
                for code in codes:
                    try:
                        pdf = core_batch.search_code(code, cfg, app.push_log)
                    except Exception as e:
                        app.push_log(f"  [{code}] 搜索异常：{e}")
                        continue
                    if pdf:
                        added.append(pdf)
            finally:
                app.q.put(("batch_import", added))

        threading.Thread(target=worker, daemon=True).start()

    def _show_menu(self, event):
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="释放配置文件（高级）", command=self._release_config)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _release_config(self):
        try:
            p = core_batch.release_config()
            self.log(f"配置文件已释放：{p}")
        except Exception as e:
            self.log(f"释放配置失败：{type(e).__name__}: {e}")

    def log(self, msg):
        self.winfo_toplevel().push_log(str(msg))


class App(TkinterDnD.Tk if _DND else tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("860x630")
        self.minsize(700, 500)
        self.q = queue.Queue()
        self.running = False
        self.current = None

        try:
            ttk.Style(self).theme_use("vista")
        except Exception:
            pass

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)      # notebook expands
        # row 2 (bottom bar) fixed height

        # header
        hd = ttk.Frame(self, padding=(12, 8, 12, 0))
        hd.grid(row=0, column=0, sticky="ew")
        ttk.Label(hd, text="OfficeToolkit", font=FONT_TITLE).pack(side="left")
        ttk.Label(hd, text="   PDF / Word 批量处理  ·  无需命令行",
                  font=FONT, foreground="#666").pack(side="left", pady=(6, 0))

        # notebook (middle)
        nb = ttk.Notebook(self)
        nb.grid(row=1, column=0, sticky="nsew", padx=10, pady=8)
        self.pdf = TabPage(nb, (".pdf",))
        self.rep = TabPage(nb, (".docx", ".doc"), with_rules=True)
        self.w2p = TabPage(nb, (".docx", ".doc"))
        self.ocr = TabPage(nb, (".pdf",))
        self.batch = BatchPage(nb)
        self.pdf.output_prefix = "PDF转Word"
        self.rep.output_prefix = "文本替换"
        self.w2p.output_prefix = "Word转PDF"
        self.ocr.output_prefix = "OCR双层PDF"
        self.pages = [self.pdf, self.rep, self.w2p, self.ocr, self.batch]
        # normalize all tasks to (files, rules, outdir, log, job)
        self.tasks = [
            lambda f, r, o, log, job: core.core_pdf2word(f, o, log, job),
            lambda f, r, o, log, job: core.core_replace(f, r, o, log, job),
            lambda f, r, o, log, job: core.core_word2pdf(f, o, log, job),
            lambda f, r, o, log, job: core.core_ocr(f, r, o, log, job),
            lambda f, r, o, log, job: core_batch.run_batch(f, log),
        ]
        for p, name in zip(self.pages,
                           (" PDF 转 Word ", " 批量替换文本 ", " Word 转 PDF ",
                            " 扫描件 OCR ", " 批处理 ")):
            nb.add(p, text=name)
        self.nb = nb
        self.current = self.pdf
        nb.bind("<<NotebookTabChanged>>", self._tab_changed)

        # ---- bottom fixed control bar ----
        bottom = ttk.LabelFrame(self, text=" 输出与执行 ", padding=(10, 6))
        bottom.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 6))
        bottom.columnconfigure(2, weight=1)   # give the path entry the extra space

        # output row
        ttk.Label(bottom, text="输出：").grid(row=0, column=0, sticky="w")
        self.out_mode = tk.StringVar(value="desktop")
        ttk.Radiobutton(bottom, text="桌面新建文件夹", variable=self.out_mode,
                        value="desktop", command=self._toggle_out).grid(row=0, column=1, sticky="w")
        self.out_var = tk.StringVar()
        self.out_entry = ttk.Entry(bottom, textvariable=self.out_var, state="disabled", width=60)
        self.out_entry.grid(row=0, column=2, sticky="ew", padx=(8, 4))
        self.out_btn = ttk.Button(bottom, text="浏览…", command=self._pick_out, state="disabled", width=7)
        self.out_btn.grid(row=0, column=3)

        # action row
        act = ttk.Frame(bottom)
        act.grid(row=1, column=0, columnspan=4, sticky="ew", pady=(6, 4))
        self.start_btn = ttk.Button(act, text="开始处理", command=self.on_start, width=14)
        self.start_btn.pack(side="left")
        self.progress = ttk.Progressbar(act, mode="indeterminate", length=140)
        self.progress.pack(side="left", padx=(10, 0))
        self.update_btn = ttk.Button(act, text="检查更新", command=self.on_check_update, width=12)
        self.update_btn.pack(side="left", padx=(8, 0))
        ttk.Label(act, text="处理前请关闭正在编辑的文件", foreground="#888").pack(side="right")

        # log (fixed, always visible)
        lf = ttk.Frame(bottom)
        lf.grid(row=2, column=0, columnspan=4, sticky="ew")
        lf.columnconfigure(0, weight=1)
        self.log_text = tk.Text(lf, height=5, font=("Consolas", 9),
                                state="disabled", wrap="none")
        lsb = ttk.Scrollbar(lf, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=lsb.set)
        self.log_text.grid(row=0, column=0, sticky="ew")
        lsb.grid(row=0, column=1, sticky="ns")

        self.after(100, self._poll)

    # ---------- tab switching ----------
    def _tab_changed(self, _=None):
        self.current = self.pages[self.nb.index(self.nb.select())]

    # ---------- shared output helpers ----------
    def _toggle_out(self):
        c = self.out_mode.get() == "custom"
        self.out_entry.configure(state="normal" if c else "disabled")
        self.out_btn.configure(state="normal" if c else "disabled")

    def _pick_out(self):
        d = filedialog.askdirectory(title="选择输出文件夹")
        if d:
            self.out_var.set(d)

    def push_log(self, msg):
        self.q.put(("log", str(msg)))

    # ---------- dynamic file queue during processing ----------
    def notify_files_added(self, page, paths):
        if not self.running or getattr(self, "job_page", None) is not page:
            return
        job = getattr(self, "job", None)
        if job is not None:
            job.add(paths)
            self.push_log(f"已加入处理队列：{len(paths)} 个文件")
        elif page is self.ocr and getattr(self, "_ocr_task_file", None):
            self._ocr_queue_add(paths)

    def notify_files_removed(self, page, paths):
        if not self.running or getattr(self, "job_page", None) is not page:
            return
        job = getattr(self, "job", None)
        if job is not None:
            removed, cancelled = job.remove(paths)
            if cancelled:
                self.push_log(f"已取消正在处理的文件：{cancelled} 个（处理完当前后跳过）")
            elif removed:
                self.push_log(f"已从队列移除：{removed} 个文件")
        elif page is self.ocr and getattr(self, "_ocr_task_file", None):
            self._ocr_queue_remove(paths)

    def _ocr_queue_add(self, paths):
        try:
            with open(self._ocr_task_file, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            q = data.setdefault("queue_add", [])
            for p in paths:
                if p not in q:
                    q.append(p)
            self._write_ocr_task(data)
        except Exception:
            pass

    def _ocr_queue_remove(self, paths):
        try:
            with open(self._ocr_task_file, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            q = data.setdefault("queue_remove", [])
            for p in paths:
                if p not in q:
                    q.append(p)
            self._write_ocr_task(data)
        except Exception:
            pass

    def _write_ocr_task(self, data):
        tmp = self._ocr_task_file + ".gui.tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False)
        os.replace(tmp, self._ocr_task_file)

    # ---------- batch pipeline ----------
    def _start_batch(self, page):
        files = list(page.file_page.files)
        if not files:
            messagebox.showwarning("提示",
                                   "请先点「搜索导入」导入文件，或向右侧列表添加 PDF",
                                   parent=self)
            return
        cfg, external = core_batch.load_config()
        self.push_log(f"开始批处理 {len(files)} 个文件" +
                      ("（使用外部配置）" if external else ""))
        self.running = True
        page.running = True
        self.job_page = page
        self.start_btn.configure(state="disabled", text="处理中…")
        self.progress.start(12)

        def worker(files, cfg):
            def op_log(m):
                self.q.put(("log", str(m)))
            try:
                core_batch.run_batch_files(files, op_log, cfg)
            except Exception:
                self.q.put(("log", "异常：\n" + traceback.format_exc()))
            finally:
                self.q.put(("done", None))

        threading.Thread(target=worker, args=(files, cfg), daemon=True).start()

    # ---------- start ----------
    def on_start(self):
        if self.running:
            return
        page = self.current
        if page is self.batch:
            self._start_batch(page)
            return
        if not page.files:
            messagebox.showwarning("提示", "请先添加要处理的文件", parent=self)
            return
        if page.with_rules and not page.rules:
            messagebox.showwarning("提示", "请至少添加一条替换规则", parent=self)
            return
        custom = self.out_var.get() if self.out_mode.get() == "custom" else ""
        try:
            outdir = core.make_output_dir(page.output_prefix, custom)
        except Exception as e:
            messagebox.showerror("错误", f"无法创建输出文件夹：{e}", parent=self)
            return

        idx = self.pages.index(page)
        files, rules = list(page.files), (list(page.rules) if page.with_rules else None)

        self.push_log(f"开始处理 {len(files)} 个文件 → {outdir}")
        self.running = True
        page.running = True
        self.job_page = page
        self.start_btn.configure(state="disabled", text="处理中…")
        self.progress.start(12)

        if page is self.ocr:
            # WeChat OCR is a C extension that holds the GIL for seconds at a
            # time; running it in a thread freezes the GUI. Run it in a
            # separate process and poll a small JSON status file instead.
            import json
            task_file = os.path.join(os.environ.get("TEMP", "C:\\Temp"),
                                     "ot_ocr_task.json")
            payload = {"files": files, "outdir": outdir, "logs": [], "done": False}
            with open(task_file, "w", encoding="utf-8") as fh:
                json.dump(payload, fh, ensure_ascii=False)
            try:
                flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
                # In the frozen EXE, sys.executable IS the app. From source,
                # it is python.exe, so pass this script path explicitly.
                if getattr(sys, "frozen", False):
                    cmd = [sys.executable, "--ocr-worker", task_file]
                else:
                    cmd = [sys.executable, os.path.abspath(__file__),
                           "--ocr-worker", task_file]
                subprocess.Popen(cmd, creationflags=flags)
            except Exception as e:
                self.push_log("无法启动 OCR 进程：" + str(e))
                self.q.put(("done", None))
            self._ocr_task_file = task_file
            self._ocr_log_pos = 0
        else:
            task = self.tasks[idx]
            job = core.JobQueue(files)
            job.idle_timeout = 2.0   # allow files added during processing
            self.job = job

            def worker(q, task, files, rules, outdir, job):
                def op_log(m):
                    q.put(("log", str(m)))
                try:
                    task(files, rules, outdir, op_log, job)
                except Exception:
                    q.put(("log", "异常：\n" + traceback.format_exc()))
                finally:
                    q.put(("done", None))

            threading.Thread(target=worker,
                             args=(self.q, task, files, rules, outdir, job),
                             daemon=True).start()

    # ---------- LAN update ----------
    def on_check_update(self):
        if self.running:
            messagebox.showinfo("提示", "请先等待当前处理完成", parent=self)
            return
        self.update_btn.configure(state="disabled", text="检查中…")
        self._append_log("正在检查局域网更新…")

        def worker():
            try:
                info = core.check_update()
            except Exception as e:
                info = {"available": False, "msg": f"检查失败：{type(e).__name__}: {e}"}
            self.q.put(("update", info))

        threading.Thread(target=worker, daemon=True).start()

    def _on_update_result(self, info):
        self.update_btn.configure(state="normal", text="检查更新")
        msg = info.get("msg", "")
        self._append_log(msg)
        if not info.get("available"):
            messagebox.showinfo("检查更新", msg, parent=self)
            return
        # available -> confirm
        if not messagebox.askyesno("发现新版本", msg + "\n\n更新需要重启程序，是否现在更新？",
                                   parent=self):
            return
        self.update_btn.configure(state="disabled", text="更新中…")
        self._append_log("正在下载并安装更新…")

        def worker():
            try:
                core.apply_update()
                self.q.put(("update_done", None))
            except Exception as e:
                self.q.put(("update_err", f"{type(e).__name__}: {e}"))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_ocr_task(self):
        """Read OCR progress from the worker process's JSON file (non-blocking)."""
        tf = getattr(self, "_ocr_task_file", None)
        if not tf or not os.path.isfile(tf):
            return
        try:
            with open(tf, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            logs = data.get("logs", [])
            for msg in logs[getattr(self, "_ocr_log_pos", 0):]:
                self._append_log(msg)
            self._ocr_log_pos = len(logs)
            if data.get("done"):
                self.q.put(("done", None))
                self._ocr_task_file = None
        except Exception:
            pass  # file being rewritten by worker; try again next tick

    def _poll(self):
        self._poll_ocr_task()
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "log":
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", val + "\n")
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
                elif kind == "batch_import":
                    self.batch.file_page._append(val or [])
                    self._append_log(f"搜索导入完成：新增 {len(val or [])} 个文件")
                elif kind == "done":
                    self.running = False
                    self.job = None
                    self.job_page = None
                    self._ocr_task_file = None
                    for p in self.pages:
                        p.running = False
                    self.start_btn.configure(state="normal", text="开始处理")
                    self.progress.stop()
                    self._append_log("— 处理结束 —")
                elif kind == "update":
                    self._on_update_result(val)
                elif kind == "update_done":
                    self._append_log("更新已安装，程序即将重启…")
                    self.destroy()
                elif kind == "update_err":
                    self.update_btn.configure(state="normal", text="检查更新")
                    self._append_log("更新失败：" + val)
                    messagebox.showerror("更新失败", val, parent=self)
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _append_log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


def _selftest():
    import traceback
    args = sys.argv[2:]
    mode = args[0] if args else "pdf2word"
    logfile = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "ot_selftest_log.txt")
    lines = []
    def log(m):
        lines.append(str(m))
        with open(logfile, "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines))
    try:
        if mode == "ocr":
            pdf = args[1] if len(args) > 1 else ""
            outdir = args[2] if len(args) > 2 else os.path.join(os.environ.get("TEMP", "C:\\Temp"), "ot_ocr_selftest")
            os.makedirs(outdir, exist_ok=True)
            log("SELFTEST OCR START")
            core.core_ocr([pdf], None, outdir, log)
            log("SELFTEST OCR DONE " + repr(os.listdir(outdir)))
        else:
            pdf = args[1] if len(args) > 1 else r"D:\Documents\桌面\新建文件夹 (8)\NC26E444-b.线上资料.pdf"
            outdir = args[2] if len(args) > 2 else os.path.join(os.environ.get("TEMP", "C:\\Temp"), "ot_selftest")
            os.makedirs(outdir, exist_ok=True)
            log("SELFTEST START")
            core.core_pdf2word([pdf], outdir, log)
            log("SELFTEST DONE " + repr(os.listdir(outdir)))
    except Exception:
        log("SELFTEST ERR\n" + traceback.format_exc())


def _ocr_worker(task_file):
    """Run OCR in a standalone process; stream progress into the JSON file.
    A small watcher thread applies dynamic add/remove requests from the GUI.
    """
    import time
    data = {"files": [], "outdir": "", "logs": [], "done": False}
    logs = []
    wlock = threading.Lock()

    def save():
        tmp = task_file + ".tmp"
        with open(tmp, "w", encoding="utf-8") as out:
            json.dump(data, out, ensure_ascii=False)
        os.replace(tmp, task_file)

    def log(msg):
        with wlock:
            logs.append(str(msg))
            data["logs"] = logs
            save()

    try:
        with open(task_file, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        logs = data.setdefault("logs", [])
        job = core.JobQueue(data.get("files", []))
        job.idle_timeout = 2.0
        stop = threading.Event()

        def watcher():
            while not stop.is_set():
                time.sleep(0.3)
                try:
                    with open(task_file, "r", encoding="utf-8") as fh:
                        cur = json.load(fh)
                    changed = False
                    adds = cur.pop("queue_add", []) or []
                    rems = cur.pop("queue_remove", []) or []
                    if adds:
                        job.add(adds); changed = True
                        with wlock:
                            logs.append(f"已加入处理队列：{len(adds)} 个文件")
                    if rems:
                        removed, cancelled = job.remove(rems); changed = True
                        with wlock:
                            if cancelled:
                                logs.append(f"已取消正在处理的文件：{cancelled} 个")
                            elif removed:
                                logs.append(f"已从队列移除：{removed} 个文件")
                    if changed:
                        data["logs"] = logs
                        with wlock:
                            save()
                except Exception:
                    pass

        wt = threading.Thread(target=watcher, daemon=True)
        wt.start()
        try:
            core.core_ocr(data.get("files", []), None, data["outdir"], log, job)
        finally:
            stop.set()
    except Exception:
        try:
            with wlock:
                logs.append("OCR 进程异常：\n" + traceback.format_exc())
                save()
        except Exception:
            pass
    finally:
        try:
            with wlock:
                data["done"] = True
                save()
        except Exception:
            pass
        except Exception:
            pass


def main():
    if "--selftest" in sys.argv:
        _selftest()
        return
    if "--ocr-worker" in sys.argv:
        i = sys.argv.index("--ocr-worker")
        if i + 1 < len(sys.argv):
            _ocr_worker(sys.argv[i + 1])
        return
    App().mainloop()


if __name__ == "__main__":
    main()
