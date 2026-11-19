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

import office_core as core

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
        n = 0
        for p in paths:
            for f in core.expand_paths([p], self.file_exts):
                if f not in self.files:
                    self.files.append(f)
                    self.listbox.insert("end", f)
                    n += 1
        if n:
            self.log(f"已添加 {n} 个文件")

    def clear_files(self):
        self.files.clear()
        self.listbox.delete(0, "end")
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

    def log(self, msg):
        self.winfo_toplevel().push_log(str(msg))


# ===========================================================================
# Main window
# ===========================================================================
class App(tk.Tk):
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
        self.pdf.output_prefix = "PDF转Word"
        self.rep.output_prefix = "文本替换"
        self.w2p.output_prefix = "Word转PDF"
        self.ocr.output_prefix = "OCR双层PDF"
        self.pages = [self.pdf, self.rep, self.w2p, self.ocr]
        # normalize all tasks to (files, rules, outdir, log)
        self.tasks = [
            lambda f, r, o, log: core.core_pdf2word(f, o, log),
            lambda f, r, o, log: core.core_replace(f, r, o, log),
            lambda f, r, o, log: core.core_word2pdf(f, o, log),
            lambda f, r, o, log: core.core_ocr(f, r, o, log),
        ]
        for p, name in zip(self.pages,
                           (" PDF 转 Word ", " 批量替换文本 ", " Word 转 PDF ", " 扫描件 OCR ")):
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

    # ---------- start ----------
    def on_start(self):
        if self.running:
            return
        page = self.current
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

            def worker(q, task, files, rules, outdir):
                def op_log(m):
                    q.put(("log", str(m)))
                try:
                    task(files, rules, outdir, op_log)
                except Exception:
                    q.put(("log", "异常：\n" + traceback.format_exc()))
                finally:
                    q.put(("done", None))

            threading.Thread(target=worker,
                             args=(self.q, task, files, rules, outdir),
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
                elif kind == "done":
                    self.running = False
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
    """Run OCR in a standalone process; stream progress into the JSON file."""
    data = {"files": [], "outdir": "", "logs": [], "done": False}
    logs = []
    try:
        with open(task_file, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        logs = data.setdefault("logs", [])

        def log(msg):
            logs.append(str(msg))
            data["logs"] = logs
            tmp = task_file + ".tmp"
            with open(tmp, "w", encoding="utf-8") as out:
                json.dump(data, out, ensure_ascii=False)
            os.replace(tmp, task_file)

        core.core_ocr(data["files"], None, data["outdir"], log)
    except Exception:
        try:
            logs.append("OCR 进程异常：\n" + traceback.format_exc())
            data["logs"] = logs
            with open(task_file, "w", encoding="utf-8") as out:
                json.dump(data, out, ensure_ascii=False)
        except Exception:
            pass
    finally:
        try:
            data["done"] = True
            tmp = task_file + ".tmp"
            with open(tmp, "w", encoding="utf-8") as out:
                json.dump(data, out, ensure_ascii=False)
            os.replace(tmp, task_file)
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
