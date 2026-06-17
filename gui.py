#!/usr/bin/env python3
"""VO 拉单工具 · 图形界面（给办公室员工，零命令行）。

打包成 Windows exe 后双击运行：选文件 → 点按钮 → 出结果。
阶段一：选 ERP + 完整天猫导出 → 生成「拣货表+面单」打印给仓库。
阶段二：选 销售ERP + 仓库返回文件(有货/无货) + 出库数据 → 生成 系统履约单号/发货表/账单/出库。
两阶段在界面上各自独立输入，互不依赖；阶段二无需天猫数据。
"""
import os
import sys
import threading
import traceback
import subprocess
from datetime import date

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter import font as tkfont

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
        root.geometry("1060x720")
        root.minsize(900, 600)

        self.erp = tk.StringVar()          # 阶段一 ERP
        self.full = tk.StringVar()         # 阶段一 完整天猫导出
        self.erp2 = tk.StringVar()         # 阶段二 销售 ERP(与阶段一独立)
        self.outdir = tk.StringVar(value=os.path.join(BASE_DIR, "输出"))
        self.shipped = tk.StringVar()      # 有货入口
        self.nogoods = tk.StringVar()      # 无货勾选入口
        self.picking = tk.StringVar()
        self.forwarder = tk.StringVar()    # 货代合并：N 份发货表
        self.shipdate = tk.StringVar(value=date.today().strftime("%Y%m%d"))
        self.mmdd = tk.StringVar(value=date.today().strftime("%m%d"))
        self._buttons = []

        self._build_ui()

    # ---------- UI ----------
    LABEL_W = 13   # 标签列统一宽度，左侧对齐

    def _file_row(self, parent, label, var, optional=False, multi=False):
        fr = ttk.Frame(parent)
        fr.pack(fill="x", pady=4)
        ttk.Label(fr, text=label, width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        pick = self._pick_files if multi else self._pick_file
        ttk.Button(fr, text="选择…", width=7,
                   command=lambda: pick(var)).pack(side="left")
        tag = "选填" if optional else "必选"
        ttk.Label(fr, text=tag, width=4, font=("", 9),
                  foreground="#9aa0a6" if optional else "#2563eb").pack(side="left", padx=(4, 0))

    def _hint(self, parent, text):
        ttk.Label(parent, text=text, style="Hint.TLabel", wraplength=860,
                  justify="left").pack(anchor="w", padx=(self.LABEL_W * 7, 0), pady=(0, 4))

    def _section(self, parent, title):
        """统一外观的区块：带标题、内边距的 LabelFrame。"""
        lf = ttk.LabelFrame(parent, text=title, style="Card.TLabelframe", padding=12)
        lf.pack(fill="x", padx=12, pady=(8, 0))
        return lf

    def _init_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")          # clam 下配色稳定生效(含 Windows)
        except tk.TclError:
            pass
        # 字号层级：区块标题 14 粗 > 字段标签 11 > 灰提示 9；按钮 11 粗(强调色)
        style.configure("Action.TButton",
                        font=("", 11, "bold"), padding=(18, 8),
                        foreground="white", background="#2563eb", borderwidth=0)
        style.map("Action.TButton",
                  background=[("active", "#1d4ed8"), ("disabled", "#b6c2d6")])
        style.configure("Card.TLabelframe.Label", font=("", 14, "bold"),
                        foreground="#111827")
        style.configure("Field.TLabel", font=("", 11), foreground="#111827")
        style.configure("Hint.TLabel", font=("", 9), foreground="#6b7280")
        style.configure("TNotebook.Tab", font=("", 11), padding=(14, 7))

    def _build_ui(self):
        self._init_styles()

        # 底部固定操作条(始终可见)
        bottom = ttk.Frame(self.root)
        bottom.pack(side="bottom", fill="x", padx=12, pady=8)
        ttk.Button(bottom, text="📂 打开输出文件夹",
                   command=lambda: open_folder(self.outdir.get())).pack(side="right")

        # ① 共用输出目录(三个标签页都写到这里)，固定在顶部
        common = self._section(self.root, "① 输出目录（共用）")
        fr = ttk.Frame(common); fr.pack(fill="x", pady=4)
        ttk.Label(fr, text="输出目录:", width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr, textvariable=self.outdir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(fr, text="选择…", width=7, command=self._pick_dir).pack(side="left")
        ttk.Label(fr, text="", width=4).pack(side="left", padx=(4, 0))

        # 左右分隔：左=分阶段标签页 / 右=运行日志，日志常驻可见、可拖宽
        pw = ttk.PanedWindow(self.root, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=12, pady=(8, 0))
        nb = ttk.Notebook(pw)
        logfr = ttk.LabelFrame(pw, text="运行日志", style="Card.TLabelframe", padding=6)
        pw.add(nb, weight=3)
        pw.add(logfr, weight=2)

        self.log = scrolledtext.ScrolledText(
            logfr, width=34, state="disabled", wrap="word",
            font=tkfont.nametofont("TkFixedFont"),
            background="white", foreground="#1f2937", insertbackground="#1f2937")
        self.log.pack(fill="both", expand=True)

        # 阶段一标签页(自带 ERP + 天猫输入，与阶段二完全独立)
        t1 = ttk.Frame(nb, padding=14); nb.add(t1, text="  阶段一 · 打印给仓库  ")
        self._file_row(t1, "ERP 导出:", self.erp, multi=True)
        self._file_row(t1, "完整天猫导出:", self.full)
        self._hint(t1, "ERP 可多选(VO/GW)。完整天猫导出：唯一天猫输入，自动定发货范围(履约+面单)并识别取消/无运单。")
        ttk.Label(t1, style="Hint.TLabel", justify="left", wraplength=520,
                  text="自动分流，生成：拣货表+面单 / 今日预计发货总获单清单 / "
                       "新订单获单清单 / 回传ERP上传表 / 已补运单清单。").pack(anchor="w", pady=(6, 10))
        b1 = ttk.Button(t1, text="▶  开始生成",
                        style="Action.TButton", command=self._run_stage1)
        b1.pack(anchor="w")
        self._buttons.append(b1)

        # 阶段二标签页(自带销售 ERP 输入，无需天猫，与阶段一完全独立)
        t2 = ttk.Frame(nb, padding=14); nb.add(t2, text="  阶段二 · 仓库返回后  ")
        self._file_row(t2, "销售ERP导出:", self.erp2, multi=True)
        self._hint(t2, "阶段二自带销售 ERP(可多选 VO/GW)，与阶段一互不影响；账单上传直接取 ERP 里的 External ID 列，无需额外文件。")
        self._file_row(t2, "有货订单清单:", self.shipped, optional=True, multi=True)
        self._file_row(t2, "无货勾选返回:", self.nogoods, optional=True, multi=True)
        self._hint(t2, "入口二选一(都填优先有货)：有货清单=真实发货单号；无货勾选=仓库标无货的返回文件。")
        self._file_row(t2, "出库原始数据:", self.picking, optional=True, multi=True)
        self._hint(t2, "出库单：选 stock picking 全量导出(可多选)，自动按发货订单过滤、拆 VO/GW。")
        fr2 = ttk.Frame(t2); fr2.pack(fill="x", pady=4)
        ttk.Label(fr2, text="日期(MMDD):", width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr2, textvariable=self.mmdd, width=10).pack(side="left", padx=6)
        ttk.Label(fr2, text="发货日期(YYYYMMDD):",
                  style="Field.TLabel").pack(side="left", padx=(16, 0))
        ttk.Entry(fr2, textvariable=self.shipdate, width=12).pack(side="left", padx=6)
        self._hint(t2, "生成：系统履约单号 / 发货表 / 账单上传 / 出库单。"
                       "四个产出彼此独立：某产出因缺数据无法生成则跳过并提示，不影响其他产出。")
        b2 = ttk.Button(t2, text="▶  开始生成",
                        style="Action.TButton", command=self._run_stage2)
        b2.pack(anchor="w", pady=(10, 0))
        self._buttons.append(b2)

        # 货代合并标签页
        t3 = ttk.Frame(nb, padding=14); nb.add(t3, text="  货代合并  ")
        self._file_row(t3, "发货表(可多份):", self.forwarder, multi=True)
        fr3 = ttk.Frame(t3); fr3.pack(fill="x", pady=4)
        ttk.Label(fr3, text="发货日期(YYYYMMDD):", width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr3, textvariable=self.shipdate, width=12).pack(side="left", padx=6)
        self._hint(t3, "把当天各店、各次拉单产生的『发货表』全选进来，合并去重成一张给货代核对的清单"
                       "（IHTCTGMBH+IH日期+单数.xlsx）。发货日期与阶段二同步。")
        b3 = ttk.Button(t3, text="▶  合并发货表",
                        style="Action.TButton", command=self._run_forwarder)
        b3.pack(anchor="w", pady=(10, 0))
        self._buttons.append(b3)

    # ---------- helpers ----------
    def _pick_file(self, var):
        p = filedialog.askopenfilename(filetypes=EXCEL_TYPES)
        if p:
            var.set(p)

    def _pick_files(self, var):
        ps = filedialog.askopenfilenames(filetypes=EXCEL_TYPES)
        if ps:
            var.set("; ".join(ps))

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
                self.root.after(0, lambda e=e, tb=tb: self._fail(e, tb))
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
    def _erp_list(self):
        return [p.strip() for p in self.erp.get().split(";") if p.strip()]

    def _run_stage1(self):
        if not self.erp.get() or not self.full.get():
            messagebox.showwarning("缺少文件", "请先选择 ERP 导出 和 完整天猫导出")
            return
        self._write("【阶段一】分流 + 生成拣货表+面单 / 新订单获单清单 / 回传ERP销售上传表 / 已补运单清单…")
        erp = self._erp_list()
        full = self.full.get() or None

        def work():
            log, _ = build_excel.build(erp, full, outdir=self.outdir.get())
            return log
        self._bg(work)

    def _run_stage2(self):
        if not self.erp2.get():
            messagebox.showwarning("缺少文件", "请先选择 销售ERP导出")
            return
        if not self.shipped.get() and not self.nogoods.get():
            messagebox.showwarning("缺少文件", "请选择『有货订单清单』或『无货勾选返回』(二选一)")
            return
        if not self.mmdd.get().strip():
            messagebox.showwarning("缺少日期", "请填写日期 MMDD（如 0611）")
            return
        self._write("【阶段二】生成 系统履约单号 / 发货表 / 账单上传 / 出库单…")
        picking = [p.strip() for p in self.picking.get().split(";") if p.strip()] or None
        shipdate = self.shipdate.get().strip() or None
        erp = [p.strip() for p in self.erp2.get().split(";") if p.strip()]
        shipped = [p.strip() for p in self.shipped.get().split(";") if p.strip()] or None
        nogoods = [p.strip() for p in self.nogoods.get().split(";") if p.strip()] or None

        def work():
            return stage2.run(self.mmdd.get().strip(), erp, shipped, nogoods,
                              outdir=self.outdir.get(),
                              picking=picking, shipdate=shipdate)
        self._bg(work)

    def _run_forwarder(self):
        files = [p.strip() for p in self.forwarder.get().split(";") if p.strip()]
        if not files:
            messagebox.showwarning("缺少文件", "请选择当天的『发货表』(可多份)")
            return
        self._write("【货代合并】合并当天发货表 → 货代清单…")
        shipdate = self.shipdate.get().strip() or None
        outdir = self.outdir.get()

        def work():
            p, n, conf = stage2.build_forwarder(files, outdir, shipdate)
            lines = [f"货代合并发货表 已生成: {p}  ({n} 单)"]
            for ref, old, new in conf:
                lines.append(f"⚠ 运单冲突 {ref}: {old} vs {new}(已保留先出现的)")
            return lines
        self._bg(work)


def _enable_dpi_awareness():
    """Windows 高分屏防模糊：进程声明 DPI 感知，避免被位图拉伸糊化。
    必须在创建任何 Tk 窗口之前调用。"""
    if not sys.platform.startswith("win"):
        return
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)    # System DPI aware (Win 8.1+)
    except Exception:
        try:
            from ctypes import windll
            windll.user32.SetProcessDPIAware()     # Vista+ 回落
        except Exception:
            pass


def _apply_dpi_scaling(root):
    """按实际 DPI 放大 Tk 字体/控件，避免 DPI 感知后界面整体偏小。"""
    if not sys.platform.startswith("win"):
        return
    try:
        dpi = root.winfo_fpixels("1i")             # 每英寸像素(DPI感知后反映真实缩放)
        if dpi and dpi > 0:
            root.tk.call("tk", "scaling", dpi / 72.0)
    except Exception:
        pass


def main():
    _enable_dpi_awareness()
    root = tk.Tk()
    _apply_dpi_scaling(root)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
