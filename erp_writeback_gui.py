#!/usr/bin/env python3
"""ERP 回写工具 · 图形界面（销售分析 + FS 回写）。

**刻意独立于 VOTool（vo_orders/gui.py），不是忘了合并。**
2026-07-08 立 fs_writeback 时定过：ERP 回写是个人月频维护动作，不进给同事用的界面——
同事没有入口就不会误触。2026-08-01 加这个 GUI 时用户复核并维持该决定：
另起一个应用，VOTool 里仍然没有任何 ERP 回写入口。

两个标签页都**只产出 Odoo 导入文件，不直接写 ERP**。上传始终是人工动作，
上传前请看产出里的对照信息复核。

Mac 上跑源码即可（个人工具，不打包 Windows exe）：
    python3 erp_writeback_gui.py
或双击 `ERP回写-Mac.command`。
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)                          # common/ 与两个包
from sales_insight import sales_insight as si  # noqa: E402
from vo_orders import fs_writeback as fw  # noqa: E402
from common import vendor as vd  # noqa: E402

EXCEL_TYPES = [("Excel / CSV", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]


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
    LABEL_W = 14

    def __init__(self, root):
        self.root = root
        root.title("ERP 回写工具 · 销售分析 / FS 回写")
        root.geometry("1180x800")
        root.minsize(980, 680)

        # 两个标签页都要产品主数据，故提到共用区——省得同一份文件选两遍
        self.products = tk.StringVar()
        self.outdir = tk.StringVar(
            value=os.path.join(BASE_DIR, "output", date.today().strftime("%Y%m%d")))

        self.si_sales = tk.StringVar()
        self.si_safety = tk.StringVar()
        self.si_weeks = tk.StringVar()                 # 留空 = 从表头自动数
        self.si_cover = tk.StringVar(value="2")
        self.si_test = tk.StringVar()

        self.fw_po = tk.StringVar()
        self.fw_sample = tk.StringVar(value="0")       # 0 = 全量

        self._buttons = []
        self._build_ui()

    # ---------- UI 骨架（与 VOTool 同一套观感） ----------
    def _file_row(self, parent, label, var, optional=False):
        fr = ttk.Frame(parent)
        fr.pack(fill="x", pady=4)
        ttk.Label(fr, text=label, width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(fr, text="选择…", width=7,
                   command=lambda: self._pick_file(var)).pack(side="left")
        tag = "选填" if optional else "必选"
        ttk.Label(fr, text=tag, width=4, font=("", 9),
                  foreground="#9aa0a6" if optional else "#2563eb").pack(side="left", padx=(4, 0))

    def _num_row(self, parent, label, var, hint):
        fr = ttk.Frame(parent)
        fr.pack(fill="x", pady=4)
        ttk.Label(fr, text=label, width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr, textvariable=var, width=10).pack(side="left", padx=6)
        ttk.Label(fr, text=hint, style="Hint.TLabel").pack(side="left", padx=(6, 0))

    def _hint(self, parent, text):
        ttk.Label(parent, text=text, style="Hint.TLabel", wraplength=560,
                  justify="left").pack(anchor="w", padx=(self.LABEL_W * 7, 0), pady=(0, 4))

    def _action_row(self, parent, text, command, **pack):
        row = ttk.Frame(parent)
        row.pack(fill="x", **pack)
        b = ttk.Button(row, text=text, style="Action.TButton", command=command)
        b.pack(side="left")
        ttk.Button(row, text="📂 打开输出文件夹",
                   command=lambda: open_folder(self.outdir.get())).pack(side="right")
        self._buttons.append(b)

    def _section(self, parent, title):
        lf = ttk.LabelFrame(parent, text=title, style="Card.TLabelframe", padding=12)
        lf.pack(fill="x", padx=12, pady=(8, 0))
        return lf

    def _init_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Action.TButton", font=("", 11, "bold"), padding=(18, 8),
                        foreground="white", background="#2563eb", borderwidth=0)
        style.map("Action.TButton",
                  background=[("active", "#1d4ed8"), ("disabled", "#b6c2d6")])
        style.configure("Card.TLabelframe.Label", font=("", 14, "bold"), foreground="#111827")
        style.configure("Field.TLabel", font=("", 11), foreground="#111827")
        style.configure("Hint.TLabel", font=("", 9), foreground="#6b7280")
        style.configure("Warn.TLabel", font=("", 10, "bold"), foreground="#b45309")
        style.configure("TNotebook.Tab", font=("", 11), padding=(14, 7))

    def _build_ui(self):
        self._init_styles()

        # 顶部横幅：这两个工具产的是 ERP 导入文件，不是普通报表——先说清楚
        ttk.Label(self.root, style="Warn.TLabel", justify="left", wraplength=1120,
                  text="⚠ 本工具产出的是 ERP 导入文件（不直接写 ERP）。上传前请看「对照」信息复核；"
                       "首次导入务必先用试水（销售分析填试水 SKU / FS 回写填试水行数）。"
                       ).pack(anchor="w", padx=14, pady=(10, 0))

        common = self._section(self.root, "① 共用输入")
        self._file_row(common, "产品主数据:", self.products)
        self._hint(common, "ERP 的 product.product 导出，两个标签页共用同一份。"
                           "**筛选条件只勾 `can be sold`**——加别的条件会漏货"
                           "（实测 `VO active=true` 只有 4575 行、漏掉 9 个在管商品）。\n"
                           "列请勾上 External ID / Quantity On Hand / Safety Stock / "
                           "Supply Remark / FS —— 缺哪列只是降级，日志会明确告警。")
        fr = ttk.Frame(common); fr.pack(fill="x", pady=4)
        ttk.Label(fr, text="输出目录:", width=self.LABEL_W, anchor="e",
                  style="Field.TLabel").pack(side="left")
        ttk.Entry(fr, textvariable=self.outdir).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(fr, text="选择…", width=7, command=self._pick_dir).pack(side="left")
        ttk.Label(fr, text="", width=4).pack(side="left", padx=(4, 0))

        pw = ttk.PanedWindow(self.root, orient="horizontal")
        pw.pack(fill="both", expand=True, padx=12, pady=(8, 10))
        nb = ttk.Notebook(pw)
        logfr = ttk.LabelFrame(pw, text="运行日志", style="Card.TLabelframe", padding=6)
        pw.add(nb, weight=3)
        pw.add(logfr, weight=2)
        self.log = scrolledtext.ScrolledText(
            logfr, width=30, state="disabled", wrap="word",
            font=tkfont.nametofont("TkFixedFont"),
            background="white", foreground="#1f2937", insertbackground="#1f2937")
        self.log.pack(fill="both", expand=True)

        self._tab_sales(nb)
        self._tab_fs(nb)

    def _tab_sales(self, nb):
        t = ttk.Frame(nb, padding=14); nb.add(t, text="  销售分析 · 安全库存  ")
        self._file_row(t, "销售数据:", self.si_sales)
        self._hint(t, "Odoo「Sales Analysis」透视导出。**建议在 ERP 里按周分组导出**"
                      "（筛好日期区间、行维度按周）——这样期间周数脚本自己数，不用你填。")
        self._file_row(t, "安全库存表:", self.si_safety, optional=True)
        self._hint(t, "运营手工维护那份。不给的话所有 SKU 都走脚本推算，"
                      "而推算值**不进回写表**，回写表会是空的。")
        self._num_row(t, "期间周数:", self.si_weeks, "留空 = 从表头自动数（推荐）")
        self._hint(t, "只有两种情况要填：导出是旧的整期累计格式（没有周信息，必须填）；"
                      "或你要覆盖自动值（会告警提示与表头不符）。填错了周均和推算值全错。")
        self._num_row(t, "覆盖周数:", self.si_cover, "安全库存按几周量推算（默认 2）")
        self._file_row(t, "试水 SKU:", self.si_test, optional=True)
        self._hint(t, "首次导入试水：填一个 SKU（如 Femibion_15199958），"
                      "回写表就只出这一条，另加一份导入前快照供事后逐字段比对。"
                      "这一栏直接填 SKU 文本，不用点「选择…」。")
        ttk.Label(t, style="Hint.TLabel", justify="left", wraplength=560,
                  text="产出四张：销量排名（含累计占比做 ABC）/ 补货提醒（缺口降序）/ "
                       "安全库存回写表（只含运营人工审过的值）/ 安全库存候选值（推算值，待运营审阅，"
                       "不进回写表）。").pack(anchor="w", pady=(6, 10))
        self._action_row(t, "▶  生成销售分析", self._run_sales)

    def _tab_fs(self, nb):
        t = ttk.Frame(nb, padding=14); nb.add(t, text="  FS 回写  ")
        self._file_row(t, "采购单导出:", self.fw_po)
        self._hint(t, "ERP 的 purchase.order 行明细。必须含 Order Lines/Total Quantity 与 "
                      "Created on 两列，缺了会直接报错——导出时记得勾。")
        self._num_row(t, "试水行数:", self.fw_sample, "0 = 全量；首次导入建议填 24")
        self._hint(t, "填 N 则按覆盖面挑 N 行（每行 FS 值互不相同，铺开各种形态与新增/覆盖），"
                      "不是切前 N 行。导完回 ERP 抽查几行确认无误，再改回 0 跑全量。")
        alias = ("已加载 %d 条供应商代号对照" % len(vd.VENDOR_ALIAS)) if vd.VENDOR_ALIAS \
            else "⚠ 没读到代号对照（config.py 的 VENDOR_ALIAS），FS 会写成供应商真名"
        ttk.Label(t, style="Hint.TLabel", justify="left", wraplength=560,
                  text="只写 FS，**不碰 Supply Remark**——那个字段留给运营同事自己维护。"
                       "FS 现值看着像人写的采购判断（如「首选AEP 不在Phoenix订」）时整行跳过，"
                       "不拿聚合结果盖掉。\n" + alias).pack(anchor="w", pady=(6, 10))
        self._action_row(t, "▶  生成 FS 回写表", self._run_fs)

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
        """后台线程执行 work()，结果/异常通过 root.after 回主线程。

        ⚠ 两个 run() 出错都抛 ValueError 而非 SystemExit——后者是 BaseException，
        这里的 except Exception 抓不到，界面会永远卡在「运行中」按钮禁用态。
        """
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

    def _need(self, *pairs):
        """检查必填项，缺了就弹窗并返回 False。"""
        missing = [name for name, var in pairs if not var.get().strip()]
        if missing:
            messagebox.showwarning("缺少文件", "请先选择：" + " / ".join(missing))
            return False
        return True

    def _num(self, var, name, default=None, allow_blank=False):
        """输入框取数。填错就弹窗并抛 ValueError 由调用方中止。"""
        s = var.get().strip()
        if not s:
            if allow_blank:
                return default
            return default
        try:
            v = float(s)
        except ValueError:
            messagebox.showwarning(f"{name}填错", f"『{name}』要填数字，当前是 {s!r}")
            raise
        if v < 0:
            messagebox.showwarning(f"{name}填错", f"『{name}』不能是负数")
            raise ValueError(name)
        return v

    # ---------- actions ----------
    def _run_sales(self):
        if not self._need(("销售数据", self.si_sales), ("产品主数据", self.products)):
            return
        try:
            weeks = self._num(self.si_weeks, "期间周数", None, allow_blank=True)
            cover = self._num(self.si_cover, "覆盖周数", 2.0)
        except ValueError:
            return
        sales, prods = self.si_sales.get().strip(), self.products.get().strip()
        safety = self.si_safety.get().strip() or None
        test = self.si_test.get().strip() or None
        outdir = self.outdir.get()
        self._write("【销售分析】销量排名 / 补货提醒 / 安全库存回写表 / 候选值…")

        def work():
            _, lines = si.run(sales, prods, safety, weeks, cover, outdir, test)
            return lines
        self._bg(work)

    def _run_fs(self):
        if not self._need(("采购单导出", self.fw_po), ("产品主数据", self.products)):
            return
        try:
            k = int(self._num(self.fw_sample, "试水行数", 0))
        except ValueError:
            return
        po, prods, outdir = self.fw_po.get().strip(), self.products.get().strip(), self.outdir.get()
        self._write("【FS 回写】读采购单 → 生成 FS 回写导入文件…"
                    + ("（试水 %d 行）" % k if k else "（全量）"))

        def work():
            _, lines = fw.run(po, prods, outdir, k)
            return lines
        self._bg(work)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
