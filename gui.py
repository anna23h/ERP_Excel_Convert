#!/usr/bin/env python3
"""VO 拉单工具 · 图形界面（给办公室员工，零命令行）。

打包成 Windows exe 后双击运行：选文件 → 点按钮 → 出结果。
阶段一：选 ERP + 天猫导出 → 生成「拣货表+面单」打印给仓库。
阶段二：选仓库返回文件(+账单模板) → 生成 B/C/D + 缺货记录。
"""
import os
import sys
import threading
import traceback
import subprocess
from datetime import date

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# 让 PyInstaller 单文件运行时也能 import 同目录模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_excel  # noqa: E402
import stage2       # noqa: E402

EXCEL_TYPES = [("Excel / CSV", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]

if getattr(sys, "frozen", False):           # 打包后
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def open_folder(path):
    """跨平台打开文件夹。"""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)              # noqa: P204
        elif sys.platform == "darwin":
            subprocess.run(["open", path])
        else:
            subprocess.run(["xdg-open", path])
    except Exception:
        pass


class App:
    def __init__(self, root):
        self.root = root
        root.title("VO 拉单工具")
        root.geometry("720x620")

        self.erp = tk.StringVar()
        self.tmall = tk.StringVar()
        self.outdir = tk.StringVar(value=os.path.join(BASE_DIR, "输出"))
        self.returned = tk.StringVar()
        self.billing = tk.StringVar()
        self.mmdd = tk.StringVar(value=date.today().strftime("%m%d"))
        self._buttons = []

        self._build_ui()

    # ---------- UI ----------
    def _file_row(self, parent, label, var, optional=False):
        fr = ttk.Frame(parent)
        fr.pack(fill="x", pady=3)
        ttk.Label(fr, text=label, width=14, anchor="e").pack(side="left")
        ttk.Entry(fr, textvariable=var).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(fr, text="选择…",
                   command=lambda: self._pick_file(var)).pack(side="left")
        if optional:
            ttk.Label(fr, text="选填", foreground="#888").pack(side="left", padx=2)

    def _build_ui(self):
        pad = dict(padx=10, pady=6)

        common = ttk.LabelFrame(self.root, text="① 输入文件（两阶段共用）")
        common.pack(fill="x", **pad)
        self._file_row(common, "ERP 导出:", self.erp)
        self._file_row(common, "天猫全量:", self.tmall)
        fr = ttk.Frame(common); fr.pack(fill="x", pady=3)
        ttk.Label(fr, text="输出目录:", width=14, anchor="e").pack(side="left")
        ttk.Entry(fr, textvariable=self.outdir).pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(fr, text="选择…", command=self._pick_dir).pack(side="left")

        s1 = ttk.LabelFrame(self.root, text="② 阶段一 → 打印给仓库")
        s1.pack(fill="x", **pad)
        b1 = ttk.Button(s1, text="生成「拣货表 + 面单」", command=self._run_stage1)
        b1.pack(side="left", padx=10, pady=8)
        self._buttons.append(b1)

        s2 = ttk.LabelFrame(self.root, text="③ 阶段二 → 仓库返回后")
        s2.pack(fill="x", **pad)
        self._file_row(s2, "返回文件:", self.returned)
        self._file_row(s2, "账单模板:", self.billing, optional=True)
        fr2 = ttk.Frame(s2); fr2.pack(fill="x", pady=3)
        ttk.Label(fr2, text="日期(MMDD):", width=14, anchor="e").pack(side="left")
        ttk.Entry(fr2, textvariable=self.mmdd, width=10).pack(side="left", padx=4)
        b2 = ttk.Button(s2, text="生成 发货表 / 账单 / 缺货记录", command=self._run_stage2)
        b2.pack(side="left", padx=10, pady=8)
        self._buttons.append(b2)

        logfr = ttk.LabelFrame(self.root, text="运行日志")
        logfr.pack(fill="both", expand=True, **pad)
        self.log = scrolledtext.ScrolledText(logfr, height=12, state="disabled")
        self.log.pack(fill="both", expand=True, padx=6, pady=6)
        ttk.Button(self.root, text="打开输出文件夹",
                   command=lambda: open_folder(self.outdir.get())).pack(pady=4)

    # ---------- helpers ----------
    def _pick_file(self, var):
        p = filedialog.askopenfilename(filetypes=EXCEL_TYPES)
        if p:
            var.set(p)

    def _pick_dir(self):
        p = filedialog.askdirectory()
        if p:
            self.outdir.set(p)

    def _write(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_busy(self, busy):
        for b in self._buttons:
            b.configure(state="disabled" if busy else "normal")

    def _bg(self, work):
        """后台线程执行 work()，结果/异常通过 root.after 回主线程。"""
        self._set_busy(True)

        def task():
            try:
                lines = work()
                self.root.after(0, lambda: self._done(lines))
            except Exception as e:
                tb = traceback.format_exc()
                self.root.after(0, lambda: self._fail(e, tb))
        threading.Thread(target=task, daemon=True).start()

    def _done(self, lines):
        for ln in lines:
            self._write(ln)
        self._write("✅ 完成\n")
        self._set_busy(False)

    def _fail(self, e, tb):
        self._write("❌ 出错：" + str(e))
        self._write(tb)
        self._set_busy(False)
        messagebox.showerror("出错", str(e))

    # ---------- actions ----------
    def _run_stage1(self):
        if not self.erp.get() or not self.tmall.get():
            messagebox.showwarning("缺少文件", "请先选择 ERP 导出 和 天猫全量 文件")
            return
        self._write("【阶段一】生成拣货表+面单…")

        def work():
            out, st = build_excel.build(self.erp.get(), self.tmall.get(),
                                        outdir=self.outdir.get())
            return [f"已生成: {out}",
                    f"  拣货表 {st['sku']} SKU；面单 {st['lines']} 行 / {st['orders']} 单；"
                    f"多品 {st['multi']} 行；标黄 {st['highlight']} 格"]
        self._bg(work)

    def _run_stage2(self):
        if not self.erp.get() or not self.tmall.get():
            messagebox.showwarning("缺少文件", "请先选择 ERP 导出 和 天猫全量 文件")
            return
        if not self.returned.get():
            messagebox.showwarning("缺少文件", "请选择仓库返回文件（带『无货勾选』页）")
            return
        if not self.mmdd.get().strip():
            messagebox.showwarning("缺少日期", "请填写日期 MMDD（如 0611）")
            return
        self._write("【阶段二】生成 B/C/D/缺货记录…")

        def work():
            return stage2.run(self.mmdd.get().strip(), self.erp.get(), self.tmall.get(),
                              nogoods=self.returned.get(),
                              billing=self.billing.get() or None,
                              outdir=self.outdir.get())
        self._bg(work)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
