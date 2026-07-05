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
import jd_export    # noqa: E402

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
        h = min(h, root.winfo_screenheight() - 80)
        root.geometry(f"{w}x{h}")
        root.minsize(min(1000, w), min(700, h))

        self.erp = tk.StringVar()          # 阶段一 ERP
        self.full = tk.StringVar()         # 阶段一 完整天猫导出
        self.erp2 = tk.StringVar()         # 阶段二 销售 ERP(与阶段一独立)
        # 默认按日期结构化归档：output/YYYYMMDD(运行当天)，免操作员每次手动改
        # 统一用英文 output/，与 CLI/build 默认一致(消除中英文分裂)
        self.outdir = tk.StringVar(
            value=os.path.join(BASE_DIR, "output", date.today().strftime("%Y%m%d")))
        self.shipped = tk.StringVar()      # 有货入口
        self.nogoods = tk.StringVar()      # 无货勾选入口
        self.picking = tk.StringVar()
        self.forwarder = tk.StringVar()    # 货代合并：N 份发货表
        self.shipdate = tk.StringVar(value=date.today().strftime("%Y%m%d"))
        self.mmdd = tk.StringVar(value=date.today().strftime("%m%d"))
        self._buttons = []

        # 京东标签：通用选列导出
        self.jd_raw = tk.StringVar()           # 京东原始导出
        self.jd_dedup = tk.BooleanVar(value=False)
        self.jd_outname = tk.StringVar(value="京东导出")
        self.jd_presets = jd_export.load_presets(BASE_DIR)

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
            logfr, width=28, state="disabled", wrap="word",
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
                       "新订单获单清单 / 回传ERP上传表 / 已补运单清单 / "
                       "补货预判清单(ERP 勾上 FS/安全库存/备注 三列时)。").pack(anchor="w", pady=(6, 10))
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

        # 京东标签页(通用选列导出：勾选+排序原始列 → 出表；可存预设)
        t4 = ttk.Frame(nb, padding=14); nb.add(t4, text="  京东  ")
        self._file_row(t4, "京东原始导出:", self.jd_raw)
        self._hint(t4, "选京东后台导出的原始 xlsx(如「复核历史查询-SKU汇总」)，点『读取列名』后在下方挑选并排序要输出的列。")
        pr = ttk.Frame(t4); pr.pack(fill="x", pady=(6, 2))
        ttk.Button(pr, text="读取列名", width=9,
                   command=self._jd_read_columns).pack(side="left")
        ttk.Label(pr, text="预设:", style="Field.TLabel").pack(side="left", padx=(16, 4))
        self.jd_preset_cb = ttk.Combobox(pr, state="readonly", width=16,
                                          values=[p["name"] for p in self.jd_presets])
        self.jd_preset_cb.pack(side="left")
        ttk.Button(pr, text="应用预设", width=9,
                   command=self._jd_apply_preset).pack(side="left", padx=(4, 0))
        ttk.Button(pr, text="存为预设", width=9,
                   command=self._jd_save_preset).pack(side="left", padx=(4, 0))

        # 双列表：左=可选列 / 右=已选列(顺序即输出列序)。cols 最后 pack，填充中部
        cols = ttk.Frame(t4)
        left = ttk.LabelFrame(cols, text="可选列", padding=4)
        mid = ttk.Frame(cols)
        right = ttk.LabelFrame(cols, text="已选列(自上而下=输出列序)", padding=4)
        left.pack(side="left", fill="both", expand=True)
        mid.pack(side="left", padx=6)
        right.pack(side="left", fill="both", expand=True)
        self.jd_avail = tk.Listbox(left, selectmode="extended", height=8,
                                   exportselection=False, activestyle="none")
        self.jd_avail.pack(fill="both", expand=True)
        self.jd_sel = tk.Listbox(right, selectmode="extended", height=8,
                                 exportselection=False, activestyle="none")
        self.jd_sel.pack(fill="both", expand=True)
        for txt, cmd in [("加入 →", self._jd_add), ("← 移除", self._jd_remove),
                         ("↑ 上移", self._jd_up), ("↓ 下移", self._jd_down)]:
            ttk.Button(mid, text=txt, width=8, command=cmd).pack(pady=3)

        # 底部控件钉住(side=bottom)：保证「生成表格」在任何窗口高度下都不被裁掉
        b4 = ttk.Button(t4, text="▶  生成表格",
                        style="Action.TButton", command=self._run_jd)
        b4.pack(side="bottom", anchor="w", pady=(8, 0))
        self._buttons.append(b4)
        ttk.Label(t4, style="Hint.TLabel", justify="left", wraplength=520,
                  text="输出文件名：YYYY年MM月DD日{单数}单 {输出名}.xlsx，写到上方共用输出目录。"
                       "长数字列(订单号/运单号)自动锁文本，防精度丢失。"
                  ).pack(side="bottom", anchor="w", pady=(6, 0))
        opt = ttk.Frame(t4); opt.pack(side="bottom", fill="x", pady=(4, 2))
        ttk.Checkbutton(opt, text="输出前对整行去重", variable=self.jd_dedup).pack(side="left")
        ttk.Label(opt, text="输出名:", style="Field.TLabel").pack(side="left", padx=(16, 4))
        ttk.Entry(opt, textvariable=self.jd_outname, width=16).pack(side="left")

        cols.pack(side="top", fill="both", expand=True, pady=(6, 4))

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
            p, n, conf, warns, rp, rn = stage2.build_forwarder(files, outdir, shipdate)
            lines = [f"货代合并发货表 已生成: {p}  ({n} 单)",
                     f"天猫回执(系统履约单号) 已生成: {rp}  ({rn} 单)"]
            for w in warns:
                lines.append(w)
            for ref, old, new in conf:
                lines.append(f"⚠ 运单冲突 {ref}: {old} vs {new}(已保留先出现的)")
            return lines
        self._bg(work)

    # ---------- 京东 ----------
    def _jd_read_columns(self):
        path = self.jd_raw.get().strip()
        if not path:
            messagebox.showwarning("缺少文件", "请先选择京东原始导出 xlsx")
            return
        try:
            allcols = jd_export.read_columns(path)
        except Exception as e:                              # noqa: BLE001
            messagebox.showerror("读取失败", str(e))
            return
        sel = list(self.jd_sel.get(0, "end"))              # 保留已选(仍存在的)
        self.jd_avail.delete(0, "end")
        for c in allcols:
            if c not in sel:
                self.jd_avail.insert("end", c)
        self._jd_reconcile_selected(allcols)
        self._write(f"【京东】读到 {len(allcols)} 列。")

    def _jd_reconcile_selected(self, allcols):
        """已选列里剔除原始数据已不存在的，避免出表时才发现缺列。"""
        kept, dropped = [], []
        for c in self.jd_sel.get(0, "end"):
            (kept if c in allcols else dropped).append(c)
        if dropped:
            self.jd_sel.delete(0, "end")
            for c in kept:
                self.jd_sel.insert("end", c)
            self._write("【京东】已选列中这些在当前数据里没有，已移除：" + "、".join(dropped))

    def _jd_apply_preset(self):
        name = self.jd_preset_cb.get()
        if not name:
            messagebox.showwarning("未选预设", "请先在下拉里选一个预设")
            return
        p = next((x for x in self.jd_presets if x["name"] == name), None)
        if not p:
            return
        allcols = list(self.jd_avail.get(0, "end")) + list(self.jd_sel.get(0, "end"))
        self.jd_sel.delete(0, "end")
        missing = []
        for c in p["columns"]:
            if allcols and c not in allcols:               # 已读列名才能校验缺失
                missing.append(c)
                continue
            self.jd_sel.insert("end", c)
        # 可选列 = 全部 - 已选
        chosen = set(self.jd_sel.get(0, "end"))
        self.jd_avail.delete(0, "end")
        for c in allcols:
            if c not in chosen:
                self.jd_avail.insert("end", c)
        self.jd_dedup.set(bool(p.get("dedup")))
        self.jd_outname.set(p.get("out_name", name))
        msg = f"【京东】已应用预设「{name}」，选中 {self.jd_sel.size()} 列。"
        if missing:
            msg += " 原始数据缺列(已跳过)：" + "、".join(missing)
        self._write(msg)

    def _jd_save_preset(self):
        cols = list(self.jd_sel.get(0, "end"))
        if not cols:
            messagebox.showwarning("没有已选列", "先把要输出的列加到右侧，再存为预设")
            return
        from tkinter import simpledialog
        name = simpledialog.askstring("存为预设", "预设名称：", parent=self.root)
        if not name:
            return
        preset = {"name": name.strip(), "columns": cols,
                  "dedup": self.jd_dedup.get(), "aggregate": None,
                  "out_name": self.jd_outname.get().strip() or name.strip()}
        try:
            jd_export.save_preset(BASE_DIR, preset)
        except Exception as e:                              # noqa: BLE001
            messagebox.showerror("保存失败", str(e))
            return
        self.jd_presets = jd_export.load_presets(BASE_DIR)
        self.jd_preset_cb.configure(values=[p["name"] for p in self.jd_presets])
        self.jd_preset_cb.set(name.strip())
        self._write(f"【京东】预设「{name.strip()}」已保存。")

    def _jd_move(self, delta):
        lb = self.jd_sel
        picks = list(lb.curselection())
        if not picks:
            return
        picks = picks if delta < 0 else list(reversed(picks))
        for i in picks:
            j = i + delta
            if 0 <= j < lb.size():
                v = lb.get(i)
                lb.delete(i); lb.insert(j, v)
                lb.selection_set(j)

    def _jd_up(self):
        self._jd_move(-1)

    def _jd_down(self):
        self._jd_move(1)

    def _jd_add(self):
        for i in self.jd_avail.curselection():
            v = self.jd_avail.get(i)
            if v not in self.jd_sel.get(0, "end"):
                self.jd_sel.insert("end", v)
        for i in reversed(self.jd_avail.curselection()):
            self.jd_avail.delete(i)

    def _jd_remove(self):
        for i in reversed(self.jd_sel.curselection()):
            v = self.jd_sel.get(i)
            self.jd_sel.delete(i)
            self.jd_avail.insert("end", v)

    def _run_jd(self):
        path = self.jd_raw.get().strip()
        cols = list(self.jd_sel.get(0, "end"))
        if not path:
            messagebox.showwarning("缺少文件", "请先选择京东原始导出 xlsx")
            return
        if not cols:
            messagebox.showwarning("没有已选列", "请把要输出的列加到右侧『已选列』")
            return
        self._write(f"【京东】导出 {len(cols)} 列 → {self.jd_outname.get()} …")
        outdir = self.outdir.get()
        dedup = self.jd_dedup.get()
        outname = self.jd_outname.get().strip() or "京东导出"

        def work():
            op, n, warns = jd_export.export(path, cols, outdir, outname, dedup=dedup)
            return [f"已生成：{op}  ({n} 行)"] + warns
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
