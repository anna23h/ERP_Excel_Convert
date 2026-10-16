"""SO / 出库单 -> Packing List 半成品。

吃 ERP 导出的 sale.order 行明细，产出照抄成品箱单格式的 Excel：
机器可知的列（品名/SKU/条码/HS/原产国）填好，仓库现场才知道的列
（托盘号/批次号/箱号/箱规/箱数/保质期/毛重/尺寸/体积重）留空给手填。

也吃 stock.picking 出库单的「详细操作」导出（按表头自动认）：这时批次号与
保质期由 ERP 预留的批次带出，一个批次一行；Quantity total 取各批次数量之和。

Quantity total（L 列）预填 SO 订购量（出库单则为批次数量之和）；
P 列 =SUMPRODUCT(箱规*箱数)，仓库填完与 L 列对拍。

用法:
    python packing_list.py "sale.order.xlsx" [-o 输出目录] [--spare 2]
    python packing_list.py "stock.picking.xlsx" [-o 输出目录] [--spare 2]
"""

import argparse
import datetime as dt
import os
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # 让 common/ 可导入
from common.xlsx import BORDER, YELLOW as WARN_FILL, unique_path  # noqa: E402

# ---- 成品箱单的固定排版 ----
COMPANY = "IHTCT GmbH \nHansaallee. 189, 40549 Düsseldorf \n"
TITLE = "   Packing List"
HEADERS = [
    "Item", "Products Name", "SKU 系统产品号", "Bar Code", "HS Code",
    "批次号", "箱号", "箱规", "箱数", "保质期", "Origin\n原产国",
    "Quantity total", "Gross Weight (kg)\n毛重",
    "Paket Measurements\n(L)*(W)*(D)cm\n包裹尺寸", " Size weight/体积重",
    "核对\n箱规×箱数",   # P 列：不在成品箱单的 15 列里，是对拍用的辅助列
]
COL_WIDTHS = {"B": 36.375, "C": 22.625, "D": 18.875, "F": 18.625,
              "G": 9.0, "J": 9.0, "N": 23.625, "P": 11.0}
ORIGIN = "DE"
DEFAULT_HS = "30049000"  # ERP 主数据常缺；成品箱单里全是这个
HEADER_ROW = 6
FIRST_DATA_ROW = 7

# 脚本填的列 / 仓库手填的列（1-based）
COL_ITEM, COL_NAME, COL_SKU, COL_BARCODE, COL_HS = 1, 2, 3, 4, 5
COL_LOT, COL_BOXSPEC, COL_BOXCOUNT, COL_EXPIRY = 6, 8, 9, 10
COL_ORIGIN, COL_QTY = 11, 12
COL_CHECK = 16  # P 列：=SUMPRODUCT(箱规×箱数)，与 L 列的 SO 订购量对拍

# ERP 导出的列名。Odoo 导出会跟界面语言走：英文一套、德文一套（用户机器语言不定，
# 德文导出会给 Auftragsreferenz/... 这类列名）。每个字段列出「英文名, 德文名」别名，
# read_orders 按别名逐一在表头里找，找到哪套算哪套。加新语言就往元组里再补一个名。
F_ORDER = ("Order Reference", "Auftragsreferenz")
F_NAME = ("Order Lines/Product/Name", "Auftragszeilen/Produkt/Name")
F_SKU = ("Order Lines/Product/Internal Reference",
         "Auftragszeilen/Produkt/Interne Referenz")
F_BARCODE = ("Order Lines/Product/Barcode", "Auftragszeilen/Produkt/Strichcode")
F_HS = ("Order Lines/Product/HS Code", "Auftragszeilen/Produkt/HS-Code")
F_QTY = ("Order Lines/Quantity", "Auftragszeilen/Menge")

# stock.picking 出库单「详细操作」导出：一行 = 一条 move line = 一个批次。
# 只有英文名——德文界面的列名还没拿到实导出，见 ISSUES [箱单] C。
# 批次必须是 Lot/Serial Number（lot_id，出库预留的现成批次）；
# Lot/Serial Number Name 是 lot_name，入库手打新批次用，出库单上恒空。
P_ORDER = ("Source Document",)
P_NAME = ("Operations/Product/Display Name",)   # "[SKU] 品名"
P_SKU = ("Operations/Product/Internal Reference",)
P_BARCODE = ("Operations/Product/Barcode",)
P_HS = ("Operations/Product/HS Code",)
P_LOT = ("Operations/Lot/Serial Number",)
P_EXPIRY = ("Operations/Lot/Serial Number/Expiration Date",)
# 逐批次数量。不能用 Stock Move/Quantity Reserved——那是整条 move 的合计，
# 一个产品分两批时会每行各算一遍。
P_RESERVED = ("Operations/Reserved",)
P_DONE = ("Operations/Done",)   # 可选：拣货/验证后 Reserved 会归零，这时取 Done
EXPIRY_FORMAT = "mm.yyyy"       # 2026-10-16 用户定：只写到月（显示 10.2028）


def _clean(value):
    """ERP 导出里空单元格有时是 False/空串，统一成 None。"""
    if value is None or value is False or value == "":
        return None
    return str(value).strip() or None


def _col(header, aliases):
    """在表头里按别名(英/德…)找列位；都找不到抛 ValueError。
    用 ValueError 而非 SystemExit——后者是 BaseException，会逃出 GUI 后台线程的
    `except Exception` 导致界面卡在「运行中」(run() docstring 记的就是这个坑)。"""
    for name in aliases:
        if name in header:
            return header.index(name)
    raise ValueError(aliases[0])


def _is_picking(header):
    """出库单导出的列都挂在 Operations/ 下；SO 导出没有这个前缀。"""
    return any(h and h.startswith("Operations/") for h in header)


def read_orders(path):
    """读 ERP 导出 → [{order, kind, lines}]，按表头自动认 SO / 出库单。

    同单内同 SKU 合并求和；跨单的同 SKU 不合并——不同 SO 的货是分开打包的。
    每行产品带 lots 列表（SO 导出恒为空，出库单一个批次一项）。
    """
    ws = openpyxl.load_workbook(path, data_only=True).active
    header = [_clean(c.value) for c in ws[1]]
    if _is_picking(header):
        return _read_picking(ws, header)

    try:
        idx = {f: _col(header, f) for f in
               (F_ORDER, F_NAME, F_SKU, F_BARCODE, F_HS, F_QTY)}
    except ValueError as e:
        raise ValueError(f"导出缺列（英文/德文列名都没找到）：{e}\n实际表头：{header}")

    orders = []
    current = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        order_no = _clean(row[idx[F_ORDER]])
        if order_no:
            current = {"order": order_no, "kind": "so", "lines": {}}
            orders.append(current)
        sku = _clean(row[idx[F_SKU]])
        if not sku or current is None:
            continue  # 小计行 / 空行
        qty = row[idx[F_QTY]] or 0
        line = current["lines"].get(sku)
        if line:
            line["qty"] += qty
        else:
            current["lines"][sku] = _new_line(
                sku, _clean(row[idx[F_NAME]]), _clean(row[idx[F_BARCODE]]),
                _clean(row[idx[F_HS]]), qty)
    return [o for o in orders if o["lines"]]


def _new_line(sku, name, barcode, hs, qty):
    return {
        "sku": sku,
        "name": name,
        "barcode": barcode,
        "hs": hs or DEFAULT_HS,
        "hs_defaulted": not hs,
        "qty": qty,
        "lots": [],   # [{"lot", "expiry", "qty"}]，只有出库单导出才有
    }


def _expiry(value):
    """Odoo 导出的效期是 datetime（时间部分是时区换算的残留，如 09:00），只留日期。"""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    return None


def _read_picking(ws, header):
    """出库单「详细操作」导出：一行一个批次。

    `Source Document`（SO 号）只写在每张出库单的第一行，后续行留空、挂在上一张下。
    同一张 SO 拆成多张出库单（欠单）时归到同一张箱单。
    """
    try:
        idx = {f: _col(header, f) for f in
               (P_ORDER, P_NAME, P_SKU, P_BARCODE, P_HS, P_LOT, P_EXPIRY,
                P_RESERVED)}
    except ValueError as e:
        hint = ""
        if str(e) == P_LOT[0] and "Operations/Lot/Serial Number Name" in header:
            hint = ("\n导出里选的是『Lot/Serial Number Name』——那是入库手打新批次用的，"
                    "出库单上恒空；要选『Lot/Serial Number』。")
        raise ValueError(f"出库单导出缺列：{e}{hint}\n实际表头：{header}")
    done_col = P_DONE[0] in header and header.index(P_DONE[0])

    by_no = {}
    current = None
    for row in ws.iter_rows(min_row=2, values_only=True):
        order_no = _clean(row[idx[P_ORDER]])
        if order_no:
            current = by_no.setdefault(
                order_no, {"order": order_no, "kind": "picking", "lines": {},
                           "used_done": False})
        sku = _clean(row[idx[P_SKU]])
        if not sku:
            continue
        if current is None:
            raise ValueError("出库单导出第一行产品前没有『Source Document』(SO 号)，"
                             "无法判断属于哪张单")
        qty = row[idx[P_RESERVED]] or 0
        done = row[done_col] if done_col is not False else 0
        if done:
            qty = done
            current["used_done"] = True

        line = current["lines"].get(sku)
        if not line:
            name = re.sub(r"^\[[^\]]*\]\s*", "", _clean(row[idx[P_NAME]]) or "")
            line = current["lines"][sku] = _new_line(
                sku, name or None, _clean(row[idx[P_BARCODE]]),
                _clean(row[idx[P_HS]]), 0)
        line["qty"] += qty
        # 同一批次分在两条 move line（不同包裹）时合成一行
        lot_no = _clean(row[idx[P_LOT]])
        lot = next((l for l in line["lots"] if l["lot"] == lot_no), None)
        if lot:
            lot["qty"] += qty
        else:
            line["lots"].append({"lot": lot_no,
                                 "expiry": _expiry(row[idx[P_EXPIRY]]),
                                 "qty": qty})
    return [o for o in by_no.values() if o["lines"]]


def _write_head(ws, made_on, so_no):
    ws["A1"] = COMPANY
    ws["A2"] = TITLE
    # 合并到 P（表宽随对拍列扩了一列），否则标题会偏左半格
    ws.merge_cells("A1:P1")
    ws.merge_cells("A2:P2")
    ws["A1"].font = Font(name="Arial", size=16, bold=True)
    ws["A2"].font = Font(name="Arial", size=28, bold=True)
    for coord in ("A1", "A2"):
        ws[coord].alignment = Alignment(horizontal="center", vertical="center",
                                        wrap_text=True)
    ws.row_dimensions[1].height = 20.25
    ws.row_dimensions[2].height = 35.25

    # SO 号：左上角 A4:C5 一个大格（3 列宽 × 2 行高）。
    # 原先放在 J5（Invoice 左边一格），单格装不下多张 SO 的拼接串（S02881+2882+2886），
    # 会被右邻的 Invoice 挡住显示不全；且这张表是打印给仓库的，SO 号要一眼看见。
    # 描边必须在 merge 之前：合并后除左上角外都是 MergedCell，赋 border 会被静默忽略，
    # 打出来只有左边半个框。先给 6 个真实单元格描边，再合并（openpyxl 合并时会把
    # 左上角的边框铺到整个区域外沿）。
    for row in ws["A4:C5"]:
        for cell in row:
            cell.border = BORDER
    ws.merge_cells("A4:C5")
    ws["A4"] = so_no
    ws["A4"].font = Font(name="Arial", size=16, bold=True)
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    ws["A4"].number_format = "@"

    ws["K5"] = "Invoice："  # 发票号人工填
    ws["M5"] = f"Date: {made_on:%d.%m.%Y}"
    for coord in ("K5", "M5"):
        ws[coord].font = Font(name="Arial", size=11, bold=True)
        ws[coord].border = BORDER
        ws[coord].number_format = "@"
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 22   # 两行合计 44pt，装得下 16 号字

    for col, text in enumerate(HEADERS, start=1):
        cell = ws.cell(HEADER_ROW, col, text)
        cell.font = Font(name="Arial", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[HEADER_ROW].height = 60

    for col, width in COL_WIDTHS.items():
        ws.column_dimensions[col].width = width


def _write_line(ws, row, line, spare):
    """一个 SKU 占 max(批次数, 1) 行 + spare 行空行（供仓库按箱拆行）。

    出库单导出带批次时，前几行逐行写批次号/保质期；SO 导出没有批次，
    就是原来的 1 主行 + spare 行，产出与加批次之前逐格一致。
    """
    last = row + max(len(line["lots"]), 1) - 1 + spare
    for i, lot in enumerate(line["lots"]):
        cell = ws.cell(row + i, COL_LOT, lot["lot"])
        cell.number_format = "@"   # 28211 这类纯数字批次号别被当数
        if lot["expiry"]:
            cell = ws.cell(row + i, COL_EXPIRY, lot["expiry"])
            cell.number_format = EXPIRY_FORMAT
    ws.cell(row, COL_NAME, line["name"])
    ws.cell(row, COL_SKU, line["sku"])
    ws.cell(row, COL_BARCODE, line["barcode"])
    ws.cell(row, COL_HS, line["hs"])
    # L 列预填 SO 订购量（2026-08-29 用户拍板）。原先这里是 =SUMPRODUCT(箱规×箱数)、
    # 刻意不预填以免与实发不符；现改为订购量打底，实发不符靠 P 列对拍列人工发现。
    ws.cell(row, COL_QTY, line["qty"])
    # P 列：仓库填完箱规/箱数后自动出数，与左边的订购量对照
    spec = get_column_letter(COL_BOXSPEC)
    count = get_column_letter(COL_BOXCOUNT)
    ws.cell(row, COL_CHECK,
            f"=SUMPRODUCT({spec}{row}:{spec}{last},{count}{row}:{count}{last})")

    for col in (COL_NAME, COL_SKU, COL_BARCODE, COL_HS, COL_QTY, COL_CHECK):
        if last > row:
            ws.merge_cells(start_row=row, start_column=col,
                           end_row=last, end_column=col)
        # ERP 主数据缺失 → 标黄，提醒回 ERP 补维护而不是在这里手补
        if col in (COL_BARCODE, COL_HS) and not ws.cell(row, col).value:
            ws.cell(row, col).fill = WARN_FILL

    for r in range(row, last + 1):
        for col in range(1, len(HEADERS) + 1):
            cell = ws.cell(r, col)
            cell.border = BORDER
            cell.font = Font(name="宋体", size=11)
            if col == COL_NAME:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="center",
                                           vertical="center")
        ws.row_dimensions[r].height = 30
    return last + 1


def so_label(orders):
    """S02881+2882+2886 —— 首个 SO 写全，其余省略共同前缀。
    文件名与表头 J5 的 SO 号共用这一口径，两处必须一致。"""
    nums = [o["order"] for o in orders]
    return "+".join([nums[0]] + [re.sub(r"^S0*", "", n) for n in nums[1:]])


def build(orders, spare, made_on):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet2"
    _write_head(ws, made_on, so_label(orders))

    row = FIRST_DATA_ROW
    for order in orders:
        for line in order["lines"].values():
            row = _write_line(ws, row, line, spare)
    last = row - 1

    if last >= FIRST_DATA_ROW:
        # 原产国全表恒为 DE；托盘号/毛重/尺寸/体积重按托盘合并，
        # 但托盘是仓库现场划分的，脚本不猜 → 留空不合并
        ws.merge_cells(start_row=FIRST_DATA_ROW, start_column=COL_ORIGIN,
                       end_row=last, end_column=COL_ORIGIN)
        ws.cell(FIRST_DATA_ROW, COL_ORIGIN, ORIGIN)
        ws.cell(FIRST_DATA_ROW, COL_ORIGIN).alignment = Alignment(
            horizontal="center", vertical="center")
    return wb


def run(export, outdir="output", spare=2):
    """CLI 与 GUI 共用的入口：出表并返回 (输出路径, 摘要行)。

    摘要行是给人对 SO 单用的，CLI 打到终端、GUI 打到日志区，两边同一份。
    没有产品行时抛 ValueError（不用 SystemExit——它是 BaseException，
    在 GUI 的后台线程里不会被捕获，界面会卡在「运行中」）。
    """
    orders = read_orders(export)
    if not orders:
        raise ValueError("导出里没有产品行")

    # 箱单是打包当天产生的单据，日期取制单日（当下），与 SO 何时下的无关
    made_on = dt.date.today()
    wb = build(orders, spare, made_on)

    name = f"{so_label(orders)}箱单{made_on:%d.%m}.xlsx"

    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)
    # 防覆盖：文件名只含 SO 号 + 制单日，同一批 SO 当天重跑必然重名。
    # 旧的那份仓库可能已经在填了，不能悄悄盖掉 → 重名自动加 (1)/(2)。
    path = Path(unique_path(str(out / name)))
    wb.save(path)
    renamed = path.name != name

    all_lines = [l for o in orders for l in o["lines"].values()]
    blanks = sum(1 for l in all_lines if not l["barcode"])
    hs_defaulted = sum(1 for l in all_lines if l["hs_defaulted"])
    # L 列已预填订购量（出库单为批次数量之和），这里再报一次总数供人工对单
    lines = []
    for o in orders:
        qty = sum(l["qty"] for l in o["lines"].values())
        if o["kind"] == "picking":
            what = "已拣" if o["used_done"] else "预留"
            n_lots = sum(1 for l in o["lines"].values() for lot in l["lots"] if lot["lot"])
            lines.append(f"  {o['order']}: {len(o['lines'])} 行 SKU / {n_lots} 个批次"
                         f" / {what} {qty:g} 件（出库单导出）")
        else:
            lines.append(f"  {o['order']}: {len(o['lines'])} 行 SKU / 订购 {qty:g} 件")
    lines.append(f"  共 {len(orders)} 张 SO，每行预留 {spare} 行空行")
    no_lot = sum(1 for o in orders if o["kind"] == "picking"
                 for l in o["lines"].values() for lot in l["lots"] if not lot["lot"])
    if no_lot:
        lines.append(f"  · {no_lot} 行在出库单上没有批次（产品不追踪批次或还没预留），"
                     f"批次号/保质期留空")
    if renamed:
        lines.append(f"  · 同名文件 {name} 已存在，本次另存为 {path.name}（不覆盖）")
    if blanks:
        lines.append(f"  ⚠ {blanks} 行的 Bar Code 在 ERP 里为空，已标黄"
                     f"——请回 ERP 补维护产品主数据")
    if hs_defaulted:
        lines.append(f"  · {hs_defaulted} 行的 HS Code 在 ERP 里为空，已按默认"
                     f" {DEFAULT_HS} 填入")
    return path, lines


def main():
    ap = argparse.ArgumentParser(description="SO / 出库单导出 → 箱单半成品")
    ap.add_argument("export", help="ERP 导出的 sale.order 或 stock.picking xlsx（按表头自动认）")
    ap.add_argument("-o", "--outdir", default="output", help="输出目录")
    ap.add_argument("--spare", type=int, default=2,
                    help="每个 SKU 预留几行空行给仓库拆批次（默认 2）")
    args = ap.parse_args()

    try:
        path, lines = run(args.export, args.outdir, args.spare)
    except ValueError as e:
        raise SystemExit(str(e))
    print(path)
    for ln in lines:
        print(ln)


if __name__ == "__main__":
    main()
