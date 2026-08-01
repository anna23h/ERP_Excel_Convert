#!/usr/bin/env python3
"""VO 拉单工具 · 图形界面（给办公室员工，零命令行）。

打包成 Windows exe 后双击运行：选文件 → 点按钮 → 出结果。
阶段一：选 ERP + 完整天猫导出 → 生成「拣货表+面单」打印给仓库。
阶段二：选 销售ERP + 仓库返回文件(有货/无货) + 出库数据 → 生成 系统履约单号/发货表/账单/出库。
两阶段在界面上各自独立输入，互不依赖；阶段二无需天猫数据。
另有「货代合并」与「箱单」两个标签页，与拉单两阶段互不相干，各自独立使用。
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
# 仓库根也进 path：箱单在隔壁 packing_list/ 包里
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import build_excel  # noqa: E402
import stage2       # noqa: E402
from packing_list import packing_list  # noqa: E402

EXCEL_TYPES = [("Excel / CSV", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]

if getattr(sys, "frozen", False):           # 打包后
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def open_folder(path):
    """跨平台打开文件夹。目录不存在则先创建(默认输出目录首次运行前尚不存在)。"""
    try:
        os.makedirs(path, exist_ok=True)
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
        # 基础尺寸按 DPI 缩放(Windows 高分屏控件放大后固定像素窗口会装不下)，并封顶不超屏
        scale = 1.0
        try:
            dpi = root.winfo_fpixels("1i")
            if dpi and dpi > 96:
                scale = dpi / 96.0
        except tk.TclError:
            pass
        w, h = int(1180 * scale), int(820 * scale)
        w = min(w, root.winfo_screenwidth() - 40)
        # 底边留够 任务栏+标题栏 的空间(随 DPI 放大)，保证窗口整体落在工作区内，
        # 否则 Windows 上标签页底部的操作行(含「打开输出文件夹」)会被任务栏遮住
        h = min(h, root.winfo_screenheight() - int(140 * scale))
        root.geometry(f"{w}x{h}")
        root.minsize(min(1000, w), min(700, h))

        self.erp = tk.StringVar()          # 阶段一 ERP
        self.full = tk.StringVar()         # 阶段一 完整天猫导出
        self.po = tk.StringVar()           # 阶段一 采购单导出(选填，补货预判清单采购参考)
        self.erp2 = tk.StringVar()         # 阶段二 销售 ERP(与阶段一独立)
        # 默认按日期结构化归档：output/YYYYMMDD(运行当天)，免操作员每次手动改
        # 统一用英文 output/，与 CLI/build 默认一致(消除中英文分裂)
        self.outdir = tk.StringVar(
            value=os.path.join(BASE_DIR, "output", date.today().strftime("%Y%m%d")))
        self.shipped = tk.StringVar()      # 有货入口
        self.nogoods = tk.StringVar()      # 无货勾选入口
        self.picking = tk.StringVar()
        self.cancel_list = tk.StringVar()  # 取消订单清单(生成取消出库单)
        self.forwarder = tk.StringVar()    # 货代合并：N 份发货表
        self.shipdate = tk.StringVar(value=date.today().strftime("%Y%m%d"))
        self.mmdd = tk.StringVar(value=date.today().strftime("%m%d"))
        self._buttons = []

        # 箱单标签：SO 导出 → 箱单半成品
        self.pl_so = tk.StringVar()            # sale.order 导出(一份，可含多张 SO)
        self.pl_spare = tk.StringVar(value="2")  # 每 SKU 预留空行

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

    def _action_row(self, parent, text, command, **pack):
        """页内底部操作行：动作按钮靠左，「打开输出文件夹」靠右并排。
        放进各标签页(而非窗口最底), 避免 Windows 任务栏遮挡最底部控件。"""
        row = ttk.Frame(parent)
        row.pack(fill="x", **pack)
        b = ttk.Button(row, text=text, style="Action.TButton", command=command)
        b.pack(side="left")
        ttk.Button(row, text="📂 打开输出文件夹",
                   command=lambda: open_folder(self.outdir.get())).pack(side="right")
        self._buttons.append(b)

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
            logfr, width=28, state="disabled", wrap="word",
            font=tkfont.nametofont("TkFixedFont"),
            background="white", foreground="#1f2937", insertbackground="#1f2937")
        self.log.pack(fill="both", expand=True)

        # 阶段一标签页(自带 ERP + 天猫输入，与阶段二完全独立)
        t1 = ttk.Frame(nb, padding=14); nb.add(t1, text="  阶段一 · 打印给仓库  ")
        self._file_row(t1, "ERP 导出:", self.erp, multi=True)
        self._file_row(t1, "完整天猫导出:", self.full)
        self._file_row(t1, "采购单导出:", self.po, optional=True)
        self._hint(t1, "ERP 导出可多选(VO/GW 各一份)。完整天猫导出必传。"
                       "采购单导出选填——无补货需求则跳过。")
        ttk.Label(t1, style="Hint.TLabel", justify="left", wraplength=520,
                  text="点「开始生成」自动生成：拣货表+面单 / 新订单获单清单 / 回传ERP上传表。"
                       "若填了采购单导出，再多一张补货预判清单。").pack(anchor="w", pady=(6, 10))
        self._action_row(t1, "▶  开始生成", self._run_stage1)

        # 阶段二标签页(自带销售 ERP 输入，无需天猫，与阶段一完全独立)
        t2 = ttk.Frame(nb, padding=14); nb.add(t2, text="  阶段二 · 仓库返回后  ")
        self._file_row(t2, "销售ERP导出:", self.erp2, multi=True)
        self._hint(t2, "销售ERP导出可多选(VO/GW 各一份)。账单上传所需信息已在 ERP 里，无需额外文件。")
        self._file_row(t2, "有货订单清单:", self.shipped, optional=True, multi=True)
        self._file_row(t2, "无货勾选返回:", self.nogoods, optional=True, multi=True)
        self._hint(t2, "两者填其一即可(都填以有货为准)：有货订单清单=仓库已发货的单号；无货勾选返回=仓库标了缺货的文件。")
        self._file_row(t2, "出库原始数据:", self.picking, optional=True, multi=True)
        self._hint(t2, "选 ERP 导出的出库单文件(可多选)，程序自动筛出本次发货的、按店(VO/GW)拆开。")
        self._file_row(t2, "取消订单清单:", self.cancel_list, optional=True, multi=True)
        self._hint(t2, "选填。传阶段一产出的『取消订单清单』(回传天猫后把后到的取消单补录进去)，"
                       "配上面『出库原始数据』→ 多产出一张『取消出库单』(运单号统一写『订单取消』、"
                       "合并不分店)，供你在 ERP 里筛出批量取消。不影响其余四张产出。")
        fr2 = ttk.Frame(t2); fr2.pack(fill="x", pady=4)
        ttk.Label(fr2, text="日期(MMDD):", width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr2, textvariable=self.mmdd, width=10).pack(side="left", padx=6)
        ttk.Label(fr2, text="发货日期(YYYYMMDD):",
                  style="Field.TLabel").pack(side="left", padx=(16, 0))
        ttk.Entry(fr2, textvariable=self.shipdate, width=12).pack(side="left", padx=6)
        self._hint(t2, "生成：系统履约单号 / 发货表 / 账单上传 / 出库单。"
                       "四张各自独立，缺哪份数据就跳过哪张，不影响其他。")
        self._action_row(t2, "▶  开始生成", self._run_stage2, pady=(10, 0))

        # 货代合并标签页
        t3 = ttk.Frame(nb, padding=14); nb.add(t3, text="  货代合并  ")
        self._file_row(t3, "发货表(可多份):", self.forwarder, multi=True)
        fr3 = ttk.Frame(t3); fr3.pack(fill="x", pady=4)
        ttk.Label(fr3, text="发货日期(YYYYMMDD):", width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr3, textvariable=self.shipdate, width=12).pack(side="left", padx=6)
        self._hint(t3, "把当天各店、各次生成的『发货表』都选进来，合并去重成一张给货代核对的清单。"
                       "发货日期与阶段二填的一致。")
        self._action_row(t3, "▶  合并发货表", self._run_forwarder, pady=(10, 0))

        # 箱单标签页(SO 导出 → 箱单半成品，机器可知的列填好、现场才知道的留空)
        t4 = ttk.Frame(nb, padding=14); nb.add(t4, text="  箱单  ")
        self._file_row(t4, "SO 导出:", self.pl_so)
        self._hint(t4, "选 ERP 导出的 sale.order 行明细。要几张 SO 合并成一张箱单，"
                       "就在 ERP 里把它们一次性导到同一份文件里——脚本按『Order Reference』自动分单。")
        fr4 = ttk.Frame(t4); fr4.pack(fill="x", pady=4)
        ttk.Label(fr4, text="每SKU预留空行:", width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr4, textvariable=self.pl_spare, width=10).pack(side="left", padx=6)
        self._hint(t4, "每个 SKU 底下多留几行空行，供仓库拆批次号/保质期(默认 2)。")
        ttk.Label(t4, style="Hint.TLabel", justify="left", wraplength=520,
                  text="产出照抄成品箱单排版：品名/SKU/条码/HS/原产国由脚本填，"
                       "托盘号/批次号/箱号/箱规/箱数/保质期/毛重/尺寸/体积重留空给仓库手填。"
                       "Quantity total 写成公式 =箱规×箱数，仓库填完自动出数。"
                       "条码在 ERP 里为空的行会标黄。").pack(anchor="w", pady=(6, 10))
        self._action_row(t4, "▶  生成箱单", self._run_packing_list)

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
        self._write("【阶段一】分流 + 生成拣货表+面单 / 新订单获单清单 / 回传ERP销售上传表…")
        erp = self._erp_list()
        full = self.full.get() or None

        def work():
            log, _ = build_excel.build(erp, full, outdir=self.outdir.get(),
                                       po_path=self.po.get().strip() or None)
            return log
        self._bg(work)

    def _run_stage2(self):
        if not self.erp2.get():
            messagebox.showwarning("缺少文件", "请先选择 销售ERP导出")
            return
        # 仅取消模式(取消订单清单+出库原始数据)可不填有货/无货；否则二选一
        cancel_only = (self.cancel_list.get() and self.picking.get()
                       and not self.shipped.get() and not self.nogoods.get())
        if not cancel_only and not self.shipped.get() and not self.nogoods.get():
            messagebox.showwarning("缺少文件",
                                   "请选择『有货订单清单』或『无货勾选返回』(二选一)；"
                                   "或只填『取消订单清单』+『出库原始数据』单独生成取消出库单")
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
        cancel = [p.strip() for p in self.cancel_list.get().split(";") if p.strip()] or None

        def work():
            return stage2.run(self.mmdd.get().strip(), erp, shipped, nogoods,
                              outdir=self.outdir.get(),
                              picking=picking, shipdate=shipdate, cancel_list=cancel)
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
            p, n, conf, warns, rp, rn = stage2.build_forwarder(files, outdir, shipdate)
            lines = [f"货代合并发货表 已生成: {p}  ({n} 单)",
                     f"天猫回执(系统履约单号) 已生成: {rp}  ({rn} 单)"]
            for w in warns:
                lines.append(w)
            for ref, old, new in conf:
                lines.append(f"⚠ 运单冲突 {ref}: {old} vs {new}(已保留先出现的)")
            return lines
        self._bg(work)

    # ---------- 箱单 ----------
    def _run_packing_list(self):
        path = self.pl_so.get().strip()
        if not path:
            messagebox.showwarning("缺少文件", "请先选择 ERP 导出的 SO 文件")
            return
        try:
            spare = int(self.pl_spare.get().strip() or 2)
        except ValueError:
            messagebox.showwarning("预留空行填错", "『每SKU预留空行』要填整数，如 2")
            return
        if spare < 0:
            messagebox.showwarning("预留空行填错", "『每SKU预留空行』不能是负数")
            return
        self._write("【箱单】读 SO 导出 → 生成箱单半成品…")
        outdir = self.outdir.get()

        def work():
            out, lines = packing_list.run(path, outdir, spare)
            return [f"箱单已生成: {out}"] + lines
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
