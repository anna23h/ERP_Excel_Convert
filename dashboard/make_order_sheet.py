#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""采购订单跟踪表生成器 · 转置录入 + 行式总览 (ISSUES: [dashboard] K)

一个工作簿两张脸，靠公式实时联动，没有脚本刷新、没有服务端、没有数据库：

  录入   完全照搬 Jürgen 的 `260825 Order template.xlsx`——字段在 A 列，每个产品占一整列，
         供应商纵向堆三块。**写的人一个习惯都不用改。**
  总览   一行一个产品，全部公式从「录入」横着拉过来，可筛选。**看的人（销售）用这张。**

与 J 条的 `make_board_xlsx.py`（行式录入 + 五档阶段 + 泳道看板）的区别，是那条路要求
Jürgen 转 90 度改用新布局——没谈成，于是整件东西没落地（见 ISSUES K 的根因复核）。
本脚本把代价挪回到「看」的一侧：布局不动，多生成一张只读的行式视图。

字段 = Jürgen 原表的一格不动，外加三项销售要的手填项（要货日 / 提出人 / 备注，
2026-09-22 补；他的原表没有，但销售要知道「谁提的、什么时候要」）。

自动算的只有三样（2026-09-22 用户圈定，其余一律不做）：
  · 已配 = 三家 Quantity 之和；缺口 = order quantity − 已配；已到 = 三家 received quant. 之和
  · 品名（德）与品名（中）= PZN → 产品字典 VLOOKUP。德语名那格**可以手打覆盖**
    （非药房品没有 PZN），覆盖只影响那一列
**不做**阶段推导、泳道看板、超时报警——那是看板形态自带的包袱，不是这张表里的信息。

坑（都是 J 条实跑踩过的，同构，直接继承）：
  · **「这一列在不在用」的守卫必须比值、不能用 `COUNTA`**：`COUNTA` 把「有公式但结果是
    空字符串」的格子算作非空，而 Product / 品名（中）行本身就是公式，于是守卫恒为真、
    空列的已配与已到一律显示成 `SUM(空格子)=0`。用 `AND(B2="",B3="")`。
  · VLOOKUP 尾部必须接 `&""`：字典里中文名为空时返回空单元格，Excel 显示成 `0`。
  · `ETA` 不设日期格式：真实数据里 `requested` / `??` 比日期还多。
  · 不用 `FILTER` 等动态数组：经 openpyxl 写出要带 `_xlfn.` 前缀，且老桌面版不认。
  · ⚠ 真实供应商名与进货价**不进公开库**：脚本里只有占位符，要真名单放本地种子
    `dashboard/data/board_seed.json`（该目录已 gitignore）。
"""
import argparse, json, os, sys

if hasattr(sys.stdout, "reconfigure"):      # Windows 控制台默认 cp1252，打不出中文
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.formatting.rule import CellIsRule, FormulaRule

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DICT = os.path.join(HERE, "data", "product_dict.min.json")
FULL_DICT = os.path.join(HERE, "data", "product_dict.json")
SEED = os.path.join(HERE, "data", "board_seed.json")
DEFAULT_OUT = os.path.join(os.path.dirname(HERE), "results", "采购订单跟踪表.xlsx")

N_PRODUCTS = 40          # 录入表预铺多少个产品列（= 总览多少行）
N_SUPPLIER = 3           # 每个产品堆几块供应商。Jürgen 现表就是 3 块

TEAL = "0C5A61"          # 沿用既有配色，不重造
INK = "1B2B2A"
MUTED = "6B7F7C"
CALC_BG = "F2F6F5"       # 公式格底色：提示「别手填」
INPUT_BG = "FFFFFF"

# ⚠ 占位符。真实名单放 dashboard/data/board_seed.json（gitignore 内）。
SUPPLIERS = ["供应商 A", "供应商 B", "供应商 C", "供应商 D", "供应商 E"]

# 样例：PZN / 需求量 / 参考价 / 下单日 / 要货日 / 提出人 / 备注 /
#       [(供应商, 数量, 单价, ETA, 实收), ...]      —— 与 board_seed.json 的形态逐项对齐
# 价格取整、供应商用占位名——只为演示联动，不是真实数据。
SAMPLE = [
    ("04100371", 2200, None, "2026-08-25", "2026-09-30", "Lisa", "客户催得紧，缺口先补第一家", [
        ("供应商 A", 140, 10.00, "2026-08-28", None),
        ("供应商 B", 1000, None, "requested", None),
        ("供应商 C", 1000, None, "requested", None)]),
    ("16233255", 2000, 6.00, "2026-08-25", "", "JD", "", [
        ("供应商 B", 200, 5.00, "2026-08-26", 200),
        ("供应商 B", 800, 5.00, "requested", None),
        ("供应商 D", 500, 5.00, "??", None)]),
]

# ── 录入表（转置）行定义 ────────────────────────────────────────────────
# 字段名与顺序照搬 Jürgen 现表；只插进三行自动值（品名中 / 已配·缺口 / 已到）。
# (行号, 标签, 是否公式格)
R_TITLE = 1
R_PRODUCT = 2       # Product          手填（德语名）
R_PZN = 3           # PZN              手填
R_NAME_ZH = 4       # 品名（中）        自动
R_AEP = 5           # AEP / HAP        手填
R_ODATE = 6         # Order date       手填
R_QTY = 7           # order quantity   手填
# 下面三行不在 Jürgen 原表里，是 2026-09-22 用户要求补的（销售要知道「谁提的、什么时候要」）。
# 位置定在需求量之后、自动值之前——手填的归手填，灰底的归灰底，中间不再插花。
R_NEEDBY = 8        # 要货日            手填
R_WHO = 9           # 提出人            手填
R_NOTE = 10         # 备注              手填
R_ALLOC = 11        # 已配              自动
R_GAP = 12          # 缺口              自动
R_RECV = 13         # 已到              自动
R_BLOCK0 = 14       # 第一块供应商起始行
BLOCK_H = 5         # 供应商 / Quantity / price / ETA / received quant.

HEAD_ROWS = [
    (R_PRODUCT, "Product", False),
    (R_PZN, "PZN", False),
    (R_NAME_ZH, "品名（中）\nProdukt CN", True),
    (R_AEP, "AEP / HAP", False),
    (R_ODATE, "Order date", False),
    (R_QTY, "order quantity", False),
    (R_NEEDBY, "要货日 Bedarf bis", False),
    (R_WHO, "提出人 Anforderer", False),
    (R_NOTE, "备注 Notiz", False),
    (R_ALLOC, "已配 zugeteilt", True),
    (R_GAP, "缺口 fehlt", True),
    (R_RECV, "已到 erhalten", True),
]
BLOCK_ROWS = ["Supplier {}", "Quantity", "price", "ETA", "received quant."]


def block_rows(n):
    """第 n 块（1-based）供应商的五个行号 → dict"""
    base = R_BLOCK0 + (n - 1) * BLOCK_H
    return {"supplier": base, "qty": base + 1, "price": base + 2,
            "eta": base + 3, "recv": base + 4}


def last_row():
    return R_BLOCK0 + N_SUPPLIER * BLOCK_H - 1


def load_dict(path):
    """收两种格式：min 版 {"v":1,"d":{PZN:[de,zh]}} 与全量版 {"products":[...]}。

    J 条默认读全量版（1.5MB），但那份是 gitignore 的产出、本机可能压根没有；
    min 版 457KB 就带着本表要的全部两个字段。两种都认，省得先跑一趟导出。
    """
    d = json.load(open(path, encoding="utf-8"))
    if "d" in d:                       # min 版
        return [(pzn, v[0] or "", (v[1] if len(v) > 1 else "") or "")
                for pzn, v in d["d"].items()]
    return [((p.get("pzn") or ""), p.get("nameDe") or "", p.get("nameZh") or "")
            for p in d.get("products", [])]


def load_seed():
    """本地种子（真实供应商名单与样例）覆盖占位符。没有就用占位符，不报错。"""
    if not os.path.exists(SEED):
        return False
    d = json.load(open(SEED, encoding="utf-8"))
    if d.get("suppliers"):
        SUPPLIERS[:] = d["suppliers"]
    if d.get("sample"):
        # 种子形态 [pzn, 需求量, 参考价, 下单日, 要货日, 提出人, 备注, allocs]，与 SAMPLE 逐项同构
        SAMPLE[:] = [tuple(r[:7]) + ([tuple(a) for a in r[7]],) for r in d["sample"]]
    return True


# ── 录入表 ──────────────────────────────────────────────────────────────
def build_input_sheet(ws):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 20
    for i in range(N_PRODUCTS):
        ws.column_dimensions[get_column_letter(2 + i)].width = 22

    t = ws.cell(R_TITLE, 1, "采购订单跟踪 · 录入 / Erfassung")
    t.font = Font(bold=True, size=12, color=TEAL)
    tip = ws.cell(R_TITLE, 2,
                  "只在这张表填。一个产品一整列，往右接着加。灰底行是自动算的，别手填。"
                  "／ Nur hier eintragen: eine Spalte pro Produkt. Graue Zeilen rechnen sich selbst.")
    tip.font = Font(size=9, color=MUTED)
    ws.row_dimensions[R_TITLE].height = 22

    head_fill = PatternFill("solid", fgColor=TEAL)
    sup_fill = PatternFill("solid", fgColor="3E7C7F")
    calc_fill = PatternFill("solid", fgColor=CALC_BG)
    head_font = Font(color="FFFFFF", bold=True, size=10)
    thin = Side(style="thin", color="C8D6D3")

    rows = list(HEAD_ROWS)
    for n in range(1, N_SUPPLIER + 1):
        b = block_rows(n)
        rows.append((b["supplier"], BLOCK_ROWS[0].format(n), False))
        for j, lab in enumerate(BLOCK_ROWS[1:], start=1):
            rows.append((b["supplier"] + j, lab, False))

    # A 列字段名
    for r, label, is_calc in rows:
        c = ws.cell(r, 1, label)
        c.fill = sup_fill if r >= R_BLOCK0 else head_fill
        c.font = head_font
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        # Product 与 品名（中）两行要容两行字（德语名、中文名都长），其余一行够了
        ws.row_dimensions[r].height = 32 if r in (R_PRODUCT, R_NAME_ZH, R_NOTE) else 18
    # 每块供应商的第一行（Supplier N）加粗一点，视觉上把三块分开
    for n in range(1, N_SUPPLIER + 1):
        ws.cell(block_rows(n)["supplier"], 1).font = Font(color="FFFFFF", bold=True, size=11)

    qty_rows = [block_rows(n)["qty"] for n in range(1, N_SUPPLIER + 1)]
    recv_rows = [block_rows(n)["recv"] for n in range(1, N_SUPPLIER + 1)]

    for i in range(N_PRODUCTS):
        L = get_column_letter(2 + i)
        # 整列是否在用。**必须比值，不能用 COUNTA**：COUNTA 把「有公式但结果是空字符串」
        # 的格子也算作非空，而 Product 行本身就是公式，于是守卫永远为真、空列的
        # 已配/已到 一律显示成 SUM(空格子)=0（2026-09-23 用户实测报出）。
        used = 'AND({L}{p}="",{L}{z}="")'.format(L=L, p=R_PRODUCT, z=R_PZN)
        pzn = "{L}{r}".format(L=L, r=R_PZN)

        # Product（德语名）：也从 PZN 查出来。原先留空让人手打，但填了 PZN 却不出德语名
        # 是反直觉的（2026-09-23 用户实测提出）。**这格仍可手打覆盖**——非药房品没有 PZN，
        # 覆盖掉公式正是预期用法，只影响这一列。
        ws["%s%d" % (L, R_PRODUCT)] = (
            '=IF({pzn}="","",IFERROR(VLOOKUP(TEXT({pzn},"00000000"),'
            '\'产品字典\'!$A:$C,2,FALSE)&"",""))'.format(pzn=pzn))
        # 品名（中）：空 PZN 不查（TEXT("") 会变成 00000000 撞上真 PZN）。
        # 尾部 &"" 是必须的：字典里中文名为空时 VLOOKUP 返回空单元格，Excel 显示成 0。
        ws["%s%d" % (L, R_NAME_ZH)] = (
            '=IF({pzn}="","",IFERROR(VLOOKUP(TEXT({pzn},"00000000"),'
            '\'产品字典\'!$A:$C,3,FALSE)&"",""))'.format(pzn=pzn))
        ws["%s%d" % (L, R_ALLOC)] = '=IF({u},"",SUM({cells}))'.format(
            u=used, cells=",".join("%s%d" % (L, r) for r in qty_rows))
        ws["%s%d" % (L, R_GAP)] = '=IF(OR({u},{q}=""),"",MAX(0,{q}-N({a})))'.format(
            u=used, q="%s%d" % (L, R_QTY), a="%s%d" % (L, R_ALLOC))
        ws["%s%d" % (L, R_RECV)] = '=IF({u},"",SUM({cells}))'.format(
            u=used, cells=",".join("%s%d" % (L, r) for r in recv_rows))

        for r, _label, is_calc in rows:
            c = ws.cell(r, 2 + i)
            c.fill = PatternFill("solid", fgColor=CALC_BG if is_calc else INPUT_BG)
            c.border = Border(bottom=thin, right=thin)
            c.alignment = Alignment(horizontal="left", vertical="center")
        # PZN 补零成 8 位显示：字典里是 8 位，Jürgen 手打常是 7 位，
        # 不补零两边看着就对不上（查找本身走 TEXT(...,"00000000")，不受影响）
        ws.cell(R_PZN, 2 + i).number_format = "00000000"
        for r in (R_PRODUCT, R_NAME_ZH):      # 德语名与中文名都长，换行显示别被邻列截断
            ws.cell(r, 2 + i).alignment = Alignment(
                horizontal="left", vertical="center", wrap_text=True)
        ws.cell(R_NOTE, 2 + i).alignment = Alignment(     # 备注也会长
            horizontal="left", vertical="center", wrap_text=True)
        ws.cell(R_AEP, 2 + i).number_format = "0.00"
        for r in (R_ODATE, R_NEEDBY):
            ws.cell(r, 2 + i).number_format = "yyyy-mm-dd"
        for r in (R_QTY, R_ALLOC, R_GAP, R_RECV):
            ws.cell(r, 2 + i).number_format = "0"
        for n in range(1, N_SUPPLIER + 1):
            b = block_rows(n)
            ws.cell(b["qty"], 2 + i).number_format = "0"
            ws.cell(b["price"], 2 + i).number_format = "0.00"
            ws.cell(b["recv"], 2 + i).number_format = "0"
            # ETA 不设日期格式：requested / ?? 比日期还多

    span = "B%d:%s%d" % (R_GAP, get_column_letter(N_PRODUCTS + 1), R_GAP)
    ws.conditional_formatting.add(span, CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color="B3261E", bold=True)))

    dv_sup = DataValidation(type="list", formula1="='供应商'!$A$2:$A$60",
                            allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv_sup)
    for n in range(1, N_SUPPLIER + 1):
        r = block_rows(n)["supplier"]
        dv_sup.add("B%d:%s%d" % (r, get_column_letter(N_PRODUCTS + 1), r))

    # ── 输入约束 ────────────────────────────────────────────────────────
    # ⚠ 一个格子只能挂**一条**数据验证。PZN 原先挂的是「不在字典里就警告」，
    # 现在把硬约束（必须是正整数）给验证、把「不在字典里」降级成条件格式变色——
    # 两者可以共存，且分工更对：垃圾输入直接挡，陌生新品只提示不拦
    # （字典是药房产品快照，新品不该被挡住，这条从 J 条起就成立）。
    last_col = get_column_letter(N_PRODUCTS + 1)
    dv_pzn = DataValidation(
        type="whole", operator="between", formula1="1", formula2="99999999",
        errorStyle="stop", allow_blank=True, showErrorMessage=True,
        error="PZN 只能是数字（1~99999999 的整数）。别带字母、空格或连字符。",
        errorTitle="PZN 只能输数字",
        prompt="填 8 位 PZN，德语名与中文名会自动出来。不在字典里的会标成橙色，但不拦你。",
        promptTitle="PZN")
    ws.add_data_validation(dv_pzn)
    dv_pzn.add("B%d:%s%d" % (R_PZN, last_col, R_PZN))

    # PZN 不在字典里 → 橙色提示（与上面的验证并行不悖）
    # ⚠ 条件格式的填充走**差异格式 dxf**，颜色必须写 `bgColor`：写成平时的
    # `PatternFill("solid", fgColor=…)` 规则会照样命中，但画出来是「无填充」，
    # 看上去就是这条规则没生效（2026-09-23 实测：COM 读回底色是 0x0）。
    ws.conditional_formatting.add(
        "B{r}:{c}{r}".format(r=R_PZN, c=last_col),
        FormulaRule(formula=['AND(B%d<>"",ISNA(MATCH(TEXT(B%d,"00000000"),'
                             '\'产品字典\'!$A:$A,0)))' % (R_PZN, R_PZN)],
                    fill=PatternFill(bgColor="FFE0B2")))

    # 数量与金额只能是数：量为非负整数，价为非负小数。
    # **日期与 ETA 不设验证**——ETA 真实数据里 requested / ?? 比日期还多（J 条已确认），
    # 下单日/要货日 实测也常被人填成文本，卡死只会逼人去关验证。
    qty_rows_all = [R_QTY] + [block_rows(n)["qty"] for n in range(1, N_SUPPLIER + 1)] \
        + [block_rows(n)["recv"] for n in range(1, N_SUPPLIER + 1)]
    dv_qty = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="0",
                            errorStyle="stop", allow_blank=True, showErrorMessage=True,
                            error="数量只能是 0 或正整数。", errorTitle="只能填数量")
    ws.add_data_validation(dv_qty)
    for r in qty_rows_all:
        dv_qty.add("B%d:%s%d" % (r, last_col, r))
    dv_price = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                              errorStyle="stop", allow_blank=True, showErrorMessage=True,
                              error="单价只能是数字。谈价中就先空着，别写字。",
                              errorTitle="只能填数字")
    ws.add_data_validation(dv_price)
    for r in [R_AEP] + [block_rows(n)["price"] for n in range(1, N_SUPPLIER + 1)]:
        dv_price.add("B%d:%s%d" % (r, last_col, r))

    # ── 锁：公式格锁死，其余放开 ─────────────────────────────────────────
    # Excel 的锁是两段式：格子的 locked 属性**只在工作表被保护后才生效**，
    # 而所有格子默认 locked=True，所以这里要反过来把「人要填的」逐个解锁。
    # ⚠ Product 行**不锁**：它虽是公式，但设计上允许手打覆盖（非药房品没有 PZN）。
    open_rows = [R_PRODUCT, R_PZN, R_AEP, R_ODATE, R_QTY, R_NEEDBY, R_WHO, R_NOTE]
    for n in range(1, N_SUPPLIER + 1):
        open_rows += list(block_rows(n).values())
    for r in open_rows:
        for i in range(N_PRODUCTS):
            ws.cell(r, 2 + i).protection = Protection(locked=False)
    # 无密码：目的是防手滑，不是防人。要加字段就「撤消工作表保护」点一下，改完再保护。
    ws.protection = SheetProtection(
        sheet=True, password=None,
        formatCells=False, formatColumns=False, formatRows=False,   # 调格式/列宽随意
        insertRows=True, insertColumns=True, deleteRows=True, deleteColumns=True,
        sort=True, autoFilter=False, selectLockedCells=False, selectUnlockedCells=False)

    ws.freeze_panes = "B2"


def fill_sample(ws):
    for i, (pzn, need, ref, odate, needby, who, note, allocs) in enumerate(SAMPLE[:N_PRODUCTS]):
        L = get_column_letter(2 + i)
        # Product 行不填：它现在是公式，填 PZN 就自动出德语名（覆盖掉反而看不出这条）
        ws["%s%d" % (L, R_PZN)] = int(pzn)
        ws["%s%d" % (L, R_QTY)] = need
        if ref is not None:
            ws["%s%d" % (L, R_AEP)] = ref
        ws["%s%d" % (L, R_ODATE)] = odate
        ws["%s%d" % (L, R_NEEDBY)] = needby
        ws["%s%d" % (L, R_WHO)] = who
        ws["%s%d" % (L, R_NOTE)] = note
        for j, (sup, q, price, eta, rec) in enumerate(allocs[:N_SUPPLIER], start=1):
            b = block_rows(j)
            ws["%s%d" % (L, b["supplier"])] = sup
            ws["%s%d" % (L, b["qty"])] = q
            if price is not None:
                ws["%s%d" % (L, b["price"])] = price
            ws["%s%d" % (L, b["eta"])] = eta      # 文本，不转日期
            if rec is not None:
                ws["%s%d" % (L, b["recv"])] = rec


# ── 总览表（行式，全公式） ───────────────────────────────────────────────
OV_FIXED = [
    ("PZN", 11, "00000000"),
    ("品名（德）Produkt", 40, None),
    ("品名（中）", 32, None),
    ("AEP / HAP", 10, "0.00"),
    ("Order date", 12, "yyyy-mm-dd"),
    ("需求量\norder qty", 11, "0"),
    ("已配\nzugeteilt", 11, "0"),
    ("缺口\nfehlt", 11, "0"),
    ("已到\nerhalten", 11, "0"),
    # 这三列排在供应商块之前、冻结线之后：冻结的是 A~I，再往左塞会把冻结块撑得没法看
    ("要货日\nBedarf bis", 12, "yyyy-mm-dd"),
    ("提出人\nAnforderer", 13, None),
    ("备注\nNotiz", 28, None),
]
# 列宽按**德文表头**定，不按数字定：筛选箭头还要再吃掉约 2 个字符宽，
# 9 宽下 zugeteilt / erhalten / received quant. 全被切（2026-09-22 Excel 实拍看出来的）
OV_BLOCK = [("Supplier {}", 16, None), ("Quantity", 11, "0"),
            ("price", 10, "0.00"), ("ETA", 12, None), ("received\nquant.", 11, "0")]
OV_HEAD = 2      # 表头行
OV_FIRST = 3     # 第一条数据行


def build_overview_sheet(ws):
    ws.sheet_view.showGridLines = False
    t = ws.cell(1, 1, "采购订单跟踪 · 总览 / Übersicht")
    t.font = Font(bold=True, size=12, color=TEAL)
    tip = ws.cell(1, 4,
                  "只读：每一行自动取自「录入」的一列，那边一改这边就动。"
                  "要排序请先复制→选择性粘贴为「值」，直接排会把公式排乱。空行是预留产品位，用筛选去掉。")
    tip.font = Font(size=9, color=MUTED)
    ws.row_dimensions[1].height = 22

    head_fill = PatternFill("solid", fgColor=TEAL)
    sup_fill = PatternFill("solid", fgColor="3E7C7F")
    head_font = Font(color="FFFFFF", bold=True, size=10)
    calc_fill = PatternFill("solid", fgColor=CALC_BG)
    thin = Side(style="thin", color="C8D6D3")

    cols = [(lab, w, fmt, False) for lab, w, fmt in OV_FIXED]
    for n in range(1, N_SUPPLIER + 1):
        for lab, w, fmt in OV_BLOCK:
            cols.append((lab.format(n), w, fmt, True))

    for i, (lab, w, _fmt, is_sup) in enumerate(cols, start=1):
        c = ws.cell(OV_HEAD, i, lab)
        c.fill = sup_fill if is_sup else head_fill
        c.font = head_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[OV_HEAD].height = 30

    def pull(L, row):
        """从录入表某列某行拉一个值；空就留空（别让 0 冒出来）"""
        return '=IF(\'录入\'!{L}${r}="","",\'录入\'!{L}${r})'.format(L=L, r=row)

    src_rows = [R_PZN, None, R_NAME_ZH, R_AEP, R_ODATE, R_QTY, R_ALLOC, R_GAP, R_RECV,
                R_NEEDBY, R_WHO, R_NOTE]          # 顺序必须与 OV_FIXED 一一对齐
    for n in range(1, N_SUPPLIER + 1):
        b = block_rows(n)
        src_rows += [b["supplier"], b["qty"], b["price"], b["eta"], b["recv"]]

    for i in range(N_PRODUCTS):
        r = OV_FIRST + i
        L = get_column_letter(2 + i)           # 对应录入表的产品列
        for j, src in enumerate(src_rows, start=1):
            cell = ws.cell(r, j)
            if src is None:
                # 品名（德）：Product 没填就用字典里的德语名顶上，别留一格空白
                cell.value = (
                    '=IF(\'录入\'!{L}${p}<>"",\'录入\'!{L}${p},'
                    'IF(\'录入\'!{L}${z}="","",IFERROR(VLOOKUP(TEXT(\'录入\'!{L}${z},"00000000"),'
                    '\'产品字典\'!$A:$C,2,FALSE)&"","")))'.format(L=L, p=R_PRODUCT, z=R_PZN))
            else:
                cell.value = pull(L, src)
            cell.fill = calc_fill
            cell.border = Border(bottom=thin)
            fmt = cols[j - 1][2]
            if fmt:
                cell.number_format = fmt
            cell.alignment = Alignment(
                horizontal="center" if fmt in ("0", "0.00", "00000000") else "left",
                vertical="center")

    gap_col = get_column_letter(8)
    ws.conditional_formatting.add(
        "{c}{a}:{c}{b}".format(c=gap_col, a=OV_FIRST, b=OV_FIRST + N_PRODUCTS - 1),
        CellIsRule(operator="greaterThan", formula=["0"],
                   font=Font(color="B3261E", bold=True)))
    ws.freeze_panes = "D%d" % OV_FIRST
    ws.auto_filter.ref = "A%d:%s%d" % (OV_HEAD, get_column_letter(len(cols)),
                                       OV_FIRST + N_PRODUCTS - 1)
    # 整张表都是公式，全锁。**筛选放开、排序锁死**——表头那句「直接排会把公式排乱」
    # 从此不只是一句提醒，Excel 会真的拦住。
    ws.protection = SheetProtection(
        sheet=True, password=None, autoFilter=False, sort=True,
        formatCells=False, formatColumns=False, formatRows=False,
        selectLockedCells=False, selectUnlockedCells=False)


# ── 旁挂两张表 ──────────────────────────────────────────────────────────
def build_dict_sheet(ws, products):
    ws.append(["PZN", "品名（德）Produkt", "品名（中）"])
    for c, w in zip("ABC", (12, 60, 40)):
        ws.column_dimensions[c].width = w
        ws["%s1" % c].fill = PatternFill("solid", fgColor=TEAL)
        ws["%s1" % c].font = Font(color="FFFFFF", bold=True, size=10)
    for pzn, de, zh in products:
        ws.append([pzn, de, zh])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:C%d" % (len(products) + 1)
    # 字典是 export_product_dict.py 的产出快照，不该在这里改（改了也会被下次导出冲掉）
    ws.protection = SheetProtection(sheet=True, password=None, autoFilter=False,
                                    selectLockedCells=False, selectUnlockedCells=False)


def build_supplier_sheet(ws):
    ws.append(["供应商 / Lieferant"])
    ws["A1"].fill = PatternFill("solid", fgColor=TEAL)
    ws["A1"].font = Font(color="FFFFFF", bold=True, size=10)
    ws.column_dimensions["A"].width = 28
    for s in SUPPLIERS:
        ws.append([s])
    n = len(SUPPLIERS) + 3
    ws["A%d" % n] = "↑ 往上面接着加就行，录入表的下拉会自动跟上（最多 58 家）"
    ws["A%d" % n].font = Font(size=9, color=MUTED)


def main():
    global N_PRODUCTS, N_SUPPLIER
    ap = argparse.ArgumentParser(
        description="生成采购订单跟踪表（转置录入 + 行式总览，Excel Online 共享版）")
    ap.add_argument("--dict", default=None,
                    help="产品字典 JSON（min 版或全量版都认，默认找 data/product_dict.min.json）")
    ap.add_argument("--out", default=DEFAULT_OUT, help="输出 xlsx 路径")
    ap.add_argument("--products", type=int, default=N_PRODUCTS, help="预铺多少个产品列")
    ap.add_argument("--suppliers", type=int, default=N_SUPPLIER, help="每个产品堆几块供应商")
    ap.add_argument("--no-sample", action="store_true", help="不写样例数据，出一张空表")
    a = ap.parse_args()

    N_PRODUCTS, N_SUPPLIER = a.products, a.suppliers

    path = a.dict or (DEFAULT_DICT if os.path.exists(DEFAULT_DICT) else FULL_DICT)
    if not os.path.exists(path):
        sys.exit("产品字典不存在: %s\n先跑 python3 dashboard/export_product_dict.py" % path)
    products = load_dict(path)
    seeded = load_seed()

    wb = Workbook()
    ws_in = wb.active
    ws_in.title = "录入"
    ws_ov = wb.create_sheet("总览")
    ws_dict = wb.create_sheet("产品字典")
    ws_sup = wb.create_sheet("供应商")

    build_input_sheet(ws_in)
    if not a.no_sample:
        fill_sample(ws_in)
    build_overview_sheet(ws_ov)
    build_dict_sheet(ws_dict, products)
    build_supplier_sheet(ws_sup)
    wb.active = wb.index(ws_ov)          # 打开默认落在总览（看的人多，填的人少）

    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    wb.save(a.out)
    print("已生成: %s" % a.out)
    print("  录入 %d 个产品列 × %d 块供应商（第 %d~%d 行）"
          % (N_PRODUCTS, N_SUPPLIER, R_PRODUCT, last_row()))
    print("  总览 %d 行 × %d 列，全部公式" % (N_PRODUCTS, len(OV_FIXED) + N_SUPPLIER * len(OV_BLOCK)))
    print("  产品字典 %d 条（有中文名 %d）· 供应商 %d 家"
          % (len(products), sum(1 for p in products if p[2]), len(SUPPLIERS)))
    print("  种子: %s" % ("dashboard/data/board_seed.json（真实名单/样例）" if seeded
                          else "占位符（真实供应商与价格不进公开库；放种子文件可覆盖）"))


if __name__ == "__main__":
    main()
