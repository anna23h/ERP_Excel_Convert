#!/usr/bin/env python3
"""入库批次录入：Odoo 入库单导出 → 操作员录入表 → Odoo 导入表。（ISSUES: [入库批次] R）

一个入口两用，按文件内容自动判断：
  1. 喂 Odoo `stock.picking` 导出（含行 ID / move ID / 应收量）→ 生成 `Lot_Entry_*.xlsx`
     （按入库单分组、每产品一行标题 + 2 行空录入行，操作员只填 批次 / MHD / 数量）。
  2. 喂填好的录入表 → 生成 `Lot_Import_*.xlsx`（直接导回 Odoo）+ `Lot_Check_*.xlsx`（核对清单）。
产出写 `output/YYYYMMDD/`。

导入格式（2026-10-11 实测导入成功）的要点，改之前先读 ISSUES 的 Q/R 两条：
  - 新批次必须写 `lot_name`；`Lot/Serial Number` 是多对一，只认已存在的批次。
  - 已有行带 `move_line_ids/id` → 更新；同产品第 2 个起的批次另起行、行 ID 留空、抄 move ID。
  - 新行必填 product / uom / location / location_dest，缺一个就报一个。
  - 导入表带上所选单的**全部**已有行（没录的原值写回）：不赌 Odoo 对未列出的 o2m 行是保留还是替换。

用户裁决：MHD 只写月份 → 当月最后一天；合计 ≠ 应收只警告；产品名里的 MHD 与录入不符只警告；
「Keine Charge」= 包装上本就没有批次，该产品跳过。
"""
import calendar
import os
import re
import sys
from collections import OrderedDict
from datetime import date, datetime

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.datavalidation import DataValidation

META_SHEET = "_meta"
META_KIND = "lot_entry_v1"
ENTRY_ROWS = 2  # 每产品预留的空录入行（用户 2026-10-11 定）

# Odoo 导出列：技术名（勾了导入兼容）与英文标签两种都认。
# 注意 `ID`（大写，数据库数字 ID）**不是** `id`（External ID），不收。
FIELDS = OrderedDict([
    ("picking_id", ("id", "External ID")),
    ("line_id", ("move_line_ids/id", "Operations/ID", "Operations/External ID")),
    ("move_id", ("move_line_ids/move_id/id", "Operations/Stock Move/ID",
                 "Operations/Stock Move/External ID")),
    ("product", ("move_line_ids/product_id", "Operations/Product")),
    ("demand", ("move_line_ids/move_id/product_uom_qty", "Operations/Stock Move/Demand")),
    ("uom", ("move_line_ids/product_uom_id", "Operations/Unit of Measure")),
    ("loc_from", ("move_line_ids/location_id", "Operations/From")),
    ("loc_to", ("move_line_ids/location_dest_id", "Operations/To")),
])
OPTIONAL = OrderedDict([
    ("ref", ("name", "Reference")),
    ("partner", ("partner_id", "Contact")),
    ("origin", ("origin", "Source Document")),
])
# 单头字段：一对多导出里只在每张单第一行出现，后续行要向下继承
PICKING_KEYS = ("picking_id", "ref", "partner", "origin")

IMPORT_HEADER = [
    "id", "move_line_ids/id", "move_line_ids/move_id/id", "move_line_ids/product_id",
    "move_line_ids/lot_name", "move_line_ids/qty_done", "move_line_ids/product_uom_id",
    "move_line_ids/location_id", "move_line_ids/location_dest_id",
    "move_line_ids/expiration_date",
]

LOT_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-/.]{2,19}$")


class InputError(Exception):
    pass


# --------------------------------------------------------------------------
# 解析小件
# --------------------------------------------------------------------------
def month_end(y, m):
    return date(y, m, calendar.monthrange(y, m)[1])


def parse_mhd(val):
    """→ (date, None) 或 (None, 错误说明)。空值返回 (None, None)。

    认：MM.YYYY / M.YYYY / DD.MM.YYYY / YYYY-MM-DD / YYYY-MM，分隔符 . / - 都行。
    只有月份 → 当月最后一天。两位年份（05.28）分不清是 月.年 还是 日.月，不猜，报错。
    """
    if val is None or (isinstance(val, str) and not val.strip()):
        return None, None
    if isinstance(val, datetime):
        return val.date(), None
    if isinstance(val, date):
        return val, None
    if isinstance(val, (int, float)):
        return None, "MHD was stored as a number - type it as text, e.g. 11.2027"
    s = re.sub(r"\s+", "", str(val))
    parts = re.split(r"[./\-]", s)
    if not all(p.isdigit() for p in parts):
        return None, "MHD not recognised: %r" % val
    nums = [int(p) for p in parts]
    lens = [len(p) for p in parts]
    try:
        if len(parts) == 2 and lens[1] == 4 and lens[0] <= 2:          # MM.YYYY
            return month_end(nums[1], nums[0]), None
        if len(parts) == 2 and lens[0] == 4 and lens[1] <= 2:          # YYYY-MM
            return month_end(nums[0], nums[1]), None
        if len(parts) == 3 and lens[2] == 4 and max(lens[:2]) <= 2:    # DD.MM.YYYY
            return date(nums[2], nums[1], nums[0]), None
        if len(parts) == 3 and lens[0] == 4 and max(lens[1:]) <= 2:    # YYYY-MM-DD
            return date(nums[0], nums[1], nums[2]), None
    except ValueError:
        return None, "MHD is not a real date: %r" % val
    return None, "MHD not recognised (use MM.YYYY or DD.MM.YYYY): %r" % val


def name_mhd_months(name):
    """产品名里自带的 MHD（如 `MHD 2027/10`、`MHD05.28`、`MHD28.07`）→ 可能的 (年, 月) 集合。

    两段数字的先后各家写法不一，两种读法都试，合法月份的都收；集合为空表示名字里没写。
    """
    out = set()
    for a, b in re.findall(r"MHD\s*:?\s*(\d{2,4})\s*[./\-]\s*(\d{2,4})", name or "", re.I):
        for y, m in ((a, b), (b, a)):
            yy = int(y) + 2000 if len(y) == 2 else int(y)
            if 1 <= int(m) <= 12 and len(m) <= 2 and 2000 <= yy <= 2099:
                out.add((yy, int(m)))
    return out


def split_product(display):
    """`[Code] Name` → (code, name)，name 里的换行压成空格（Odoo 产品名偶有硬换行）。"""
    text = re.sub(r"\s+", " ", str(display or "")).strip()
    m = re.match(r"^\[([^\]]+)\]\s*(.*)$", text)
    return (m.group(1), m.group(2)) if m else ("", text)


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def clean_num(x):
    return int(x) if isinstance(x, float) and x.is_integer() else x


# --------------------------------------------------------------------------
# 1. 读 Odoo 导出
# --------------------------------------------------------------------------
def _map_header(header):
    pos = {}
    for key, aliases in list(FIELDS.items()) + list(OPTIONAL.items()):
        for a in aliases:
            if a in header:
                pos[key] = header.index(a)
                break
    missing = [k for k in FIELDS if k not in pos]
    if missing:
        need = ["%s (e.g. %s)" % (k, FIELDS[k][-1]) for k in missing]
        raise InputError("Odoo export is missing columns: " + "; ".join(need))
    return pos


def read_export(ws):
    """→ OrderedDict picking_id → {ref, partner, origin, moves: OrderedDict move_id → move}。"""
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    pos = _map_header(header)
    pickings = OrderedDict()
    cur = {}
    for r in rows[1:]:
        if not any(v not in (None, "") for v in r):
            continue
        get = lambda k: r[pos[k]] if k in pos else None
        if get("picking_id"):
            cur = {k: get(k) for k in PICKING_KEYS}
        if not cur:
            raise InputError("First data row has no picking External ID")
        if not get("line_id") or not get("move_id"):
            continue
        pk = pickings.setdefault(cur["picking_id"], {
            "ref": cur.get("ref"), "partner": cur.get("partner"),
            "origin": cur.get("origin"), "moves": OrderedDict()})
        mv = pk["moves"].setdefault(get("move_id"), {
            "product": get("product"), "demand": num(get("demand")) or 0,
            "uom": get("uom"), "loc_from": get("loc_from"), "loc_to": get("loc_to"),
            "lines": []})
        mv["lines"].append(get("line_id"))
    if not pickings:
        raise InputError("No receipt lines found in the export")
    return pickings


# --------------------------------------------------------------------------
# 2. 生成录入表
# --------------------------------------------------------------------------
# 可见列
C_CODE, C_NAME, C_SOLL, C_LOT, C_MHD, C_QTY, C_NOLOT, C_STAT = range(1, 9)
# 隐藏技术列（只在产品行上有值）
TECH = ["_type", "move_id", "line_ids", "picking_id", "product", "uom", "loc_from", "loc_to"]
C_TECH0 = 10  # J 起；I 留空作隔离

THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_HEAD = PatternFill("solid", fgColor="1F3864")
F_PROD = PatternFill("solid", fgColor="DDE4EE")
F_INPUT = PatternFill("solid", fgColor="FFFFFF")
F_OK = PatternFill("solid", fgColor="C6EFCE")
F_BAD = PatternFill("solid", fgColor="FFC7CE")
F_WARN = PatternFill("solid", fgColor="FFEB9C")
UNLOCK = Protection(locked=False)


def build_entry(pickings, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Erfassung"
    heads = ["Produkt", "Bezeichnung", "Soll", "Charge / Lot", "MHD", "Menge",
             "Keine Charge", "Σ / Status"]
    widths = [26, 52, 7, 18, 13, 8, 9, 16]
    for i, (h, w) in enumerate(zip(heads, widths), start=1):
        c = ws.cell(1, i, h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        ws.column_dimensions[c.column_letter].width = w
    for j, h in enumerate(TECH):
        ws.cell(1, C_TECH0 + j, h)
        ws.column_dimensions[ws.cell(1, C_TECH0 + j).column_letter].hidden = True
    ws.column_dimensions["I"].hidden = True
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    dv_nolot = DataValidation(type="list", formula1='"x"', allow_blank=True,
                              error="Only x (or empty)", showErrorMessage=True)
    dv_qty = DataValidation(type="decimal", operator="greaterThan", formula1="0",
                            allow_blank=True, error="Quantity must be a number > 0",
                            showErrorMessage=True)
    ws.add_data_validation(dv_nolot)
    ws.add_data_validation(dv_qty)

    r = 2
    status_cells = []
    for pid, pk in pickings.items():
        ref = pk["ref"] or pk["origin"] or pid
        title = " · ".join(str(x) for x in (ref, pk["origin"] if pk["ref"] else None,
                                              pk["partner"]) if x)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=C_STAT)
        c = ws.cell(r, 1, title)
        c.font = Font(bold=True, size=13, color="FFFFFF")
        c.fill = F_HEAD
        c.alignment = Alignment(vertical="center")
        ws.row_dimensions[r].height = 24
        ws.cell(r, C_TECH0, "H")
        r += 1
        for mid, mv in pk["moves"].items():
            code, name = split_product(mv["product"])
            prod_row = r
            vals = {C_CODE: code, C_NAME: name, C_SOLL: clean_num(mv["demand"])}
            for col in range(1, C_STAT + 1):
                c = ws.cell(r, col, vals.get(col))
                c.fill = F_PROD
                c.border = BOX
                c.alignment = Alignment(vertical="center", wrap_text=(col == C_NAME))
            ws.cell(r, C_CODE).font = Font(bold=True, size=12)
            ws.cell(r, C_SOLL).font = Font(bold=True, size=12)
            ws.cell(r, C_SOLL).alignment = Alignment(horizontal="center", vertical="center")
            nolot = ws.cell(r, C_NOLOT)
            nolot.fill = F_INPUT
            nolot.protection = UNLOCK
            nolot.alignment = Alignment(horizontal="center", vertical="center")
            dv_nolot.add(nolot)
            tech = ["P", mid, "|".join(mv["lines"]), pid, mv["product"], mv["uom"],
                    mv["loc_from"], mv["loc_to"]]
            for j, v in enumerate(tech):
                ws.cell(r, C_TECH0 + j, v)
            r += 1
            for _ in range(ENTRY_ROWS):
                for col in range(1, C_STAT + 1):
                    ws.cell(r, col).border = BOX
                for col in (C_LOT, C_MHD, C_QTY):
                    c = ws.cell(r, col)
                    c.protection = UNLOCK
                    c.fill = F_INPUT
                    if col != C_QTY:
                        c.number_format = "@"  # 文本：防 11.2027 变小数、1E5 变科学计数
                dv_qty.add(ws.cell(r, C_QTY))
                ws.cell(r, C_CODE, "   ↳").font = Font(color="999999")
                ws.cell(r, C_TECH0, "E")
                r += 1
            # 合计区间一直延到下一产品行/单头行：在它正上方插行也会被公式区间吃进去
            a, b = prod_row + 1, r
            L, Q, S = (ws.cell(1, x).column_letter for x in (C_LOT, C_QTY, C_SOLL))
            n = "COUNTA(%s%d:%s%d)" % (L, a, L, b)
            s = "SUM(%s%d:%s%d)" % (Q, a, Q, b)
            G = ws.cell(1, C_NOLOT).column_letter
            f = ('=IF({g}{p}<>"","keine Charge",IF({n}=0,"",IF(AND({n}=1,{s}=0),"✓ = Soll",'
                 'IF({s}={c}{p},"✓ "&{s},"≠ "&{s}&" / "&{c}{p}))))').format(
                g=G, p=prod_row, n=n, s=s, c=S)
            st = ws.cell(prod_row, C_STAT, f)
            st.alignment = Alignment(horizontal="center", vertical="center")
            status_cells.append(st.coordinate)
    end = r
    ws.cell(end, C_TECH0, "END")

    col = ws.cell(1, C_STAT).column_letter
    rng = "%s2:%s%d" % (col, col, end)
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(%s2,1)="✓"' % col], fill=F_OK))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(%s2,1)="≠"' % col], fill=F_WARN))

    ws.protection.sheet = True
    ws.protection.insertRows = False   # False = 允许：多批次时操作员自己插行
    ws.protection.deleteRows = False
    ws.protection.formatCells = False
    ws.protection.autoFilter = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:1"

    meta = wb.create_sheet(META_SHEET)
    meta["A1"], meta["B1"] = "kind", META_KIND
    meta["A2"], meta["B2"] = "created", datetime.now().strftime("%Y-%m-%d %H:%M")
    meta.sheet_state = "hidden"
    wb.save(path)
    return sum(len(p["moves"]) for p in pickings.values())


# --------------------------------------------------------------------------
# 3. 录入表 → 导入表 + 核对清单
# --------------------------------------------------------------------------
def read_entry(ws):
    """→ 产品块列表：{tech..., code, name, soll, nolot, entries:[(row, lot, mhd, qty)]}。

    行类型看隐藏列 `_type`；操作员插进来的行该列为空，归到上方最近的产品块。
    """
    blocks, cur = [], None
    for r in range(2, ws.max_row + 1):
        t = ws.cell(r, C_TECH0).value
        if t == "H" or t == "END":
            cur = None
            continue
        if t == "P":
            tech = {TECH[j]: ws.cell(r, C_TECH0 + j).value for j in range(1, len(TECH))}
            cur = dict(tech, row=r, code=ws.cell(r, C_CODE).value,
                       name=ws.cell(r, C_NAME).value, soll=num(ws.cell(r, C_SOLL).value) or 0,
                       nolot=bool(str(ws.cell(r, C_NOLOT).value or "").strip()), entries=[])
            blocks.append(cur)
            continue
        if cur is None:
            continue
        lot = ws.cell(r, C_LOT).value
        mhd = ws.cell(r, C_MHD).value
        qty = ws.cell(r, C_QTY).value
        lot = re.sub(r"\s+", " ", str(lot)).strip() if lot not in (None, "") else ""
        if lot or mhd not in (None, "") or qty not in (None, ""):
            cur["entries"].append((r, lot, mhd, qty))
    return blocks


def check_block(b):
    """→ (lots, issues)。lots = [(lot, mhd_date, qty)]；issues = [(level, row, msg)]，level ∈ ERROR/WARN。"""
    issues, lots = [], []
    E = lambda row, m: issues.append(("ERROR", row, m))
    W = lambda row, m: issues.append(("WARN", row, m))
    entries = b["entries"]
    if b["nolot"]:
        if any(lot for _, lot, _, _ in entries):
            E(b["row"], "'Keine Charge' is ticked but a lot was entered")
        elif any(m not in (None, "") for _, _, m, _ in entries):
            W(b["row"], "No lot: the MHD cannot be stored in Odoo without a lot, skipped")
        return [], issues
    if not entries:
        return [], issues
    seen = set()
    name_months = name_mhd_months(b["name"])
    for row, lot, mhd, qty in entries:
        if not lot:
            E(row, "Lot is empty (tick 'Keine Charge' if the pack has no lot)")
            continue
        if not LOT_RE.match(lot):
            E(row, "Lot %r: only letters, digits and - / . allowed, 3-20 characters" % lot)
        if lot.upper() in seen:
            E(row, "Lot %r entered twice for this product" % lot)
        seen.add(lot.upper())
        d, err = parse_mhd(mhd)
        if err:
            E(row, err)
        elif d is None:
            W(row, "No MHD entered for lot %r" % lot)
        else:
            if d <= date.today():
                E(row, "MHD %s is not in the future" % d.strftime("%d.%m.%Y"))
            if name_months and (d.year, d.month) not in name_months:
                W(row, "MHD %s differs from the MHD in the product name" % d.strftime("%m.%Y"))
        q = num(qty)
        if qty not in (None, "") and (q is None or q <= 0):
            E(row, "Quantity %r is not a number > 0" % qty)
        lots.append((lot, d, q))
    if any(i[0] == "ERROR" for i in issues):
        return lots, issues
    if len(lots) == 1 and lots[0][2] is None:
        lots[0] = (lots[0][0], lots[0][1], b["soll"])
    elif any(q is None for _, _, q in lots):
        E(b["row"], "Several lots: every lot needs a quantity")
        return lots, issues
    total = sum(q for _, _, q in lots)
    if abs(total - b["soll"]) > 1e-9:
        W(b["row"], "Lots add up to %s, expected %s" % (clean_num(total), clean_num(b["soll"])))
    return lots, issues


def build_import(blocks, import_path, check_path):
    by_picking = OrderedDict()
    for b in blocks:
        by_picking.setdefault(b["picking_id"], []).append(b)

    out, check = [], []
    n_written = 0
    for pid, bl in by_picking.items():
        results = [(b,) + check_block(b) for b in bl]
        if not any(lots and not any(i[0] == "ERROR" for i in iss) for _, lots, iss in results):
            for b, lots, iss in results:
                check.extend(_check_rows(b, lots, iss))
            continue  # 这张单一个能写的都没有：整单不进导入表
        first = True
        for b, lots, iss in results:
            check.extend(_check_rows(b, lots, iss))
            ok = lots and not any(i[0] == "ERROR" for i in iss)
            lines = (b["line_ids"] or "").split("|")
            common = [b["product"], None, None, b["uom"], b["loc_from"], b["loc_to"], None]
            if ok:
                n_written += 1
                for k, (lot, d, q) in enumerate(lots):
                    line = lines[k] if k < len(lines) else None
                    row = [pid if first else None, line, b["move_id"]] + common
                    row[4], row[5] = lot, clean_num(q)
                    row[9] = datetime(d.year, d.month, d.day) if d else None
                    out.append(row)
                    first = False
                spare = lines[len(lots):]
            else:
                spare = lines
            for line in spare:  # 没录 / 标红的已有行：原值写回，不动它
                row = [pid if first else None, line, b["move_id"]] + common
                row[5] = 0
                out.append(row)
                first = False

    if out:
        wb = Workbook()
        ws = wb.active
        ws.title = "Import"
        ws.append(IMPORT_HEADER)
        for row in out:
            ws.append(row)
        for c in ws["J"][1:]:
            c.number_format = "yyyy-mm-dd hh:mm:ss"
        for col, w in zip("ABCDEFGHIJ", (42, 44, 40, 60, 14, 9, 10, 26, 12, 20)):
            ws.column_dimensions[col].width = w
        wb.save(import_path)
    _write_check(check, check_path)
    return n_written, check


def _check_rows(b, lots, issues):
    rows = []
    if not b["entries"] and not b["nolot"]:
        return [("EMPTY", b["code"], "", "", "", "Nothing entered - line left unchanged")]
    for level, row, msg in issues:
        rows.append((level, b["code"], "", "", "", "row %d: %s" % (row, msg)))
    if not any(i[0] == "ERROR" for i in issues):
        for lot, d, q in lots:
            rows.append(("OK", b["code"], lot, d.strftime("%d.%m.%Y") if d else "",
                         clean_num(q), ""))
    else:
        rows.append(("SKIP", b["code"], "", "", "", "Product NOT in the import file - fix the red rows"))
    if b["nolot"] and not any(i[0] == "ERROR" for i in issues):
        rows.append(("SKIP", b["code"], "", "", "", "Keine Charge - not in the import file"))
    return rows


def _write_check(check, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Check"
    ws.append(["Status", "Produkt", "Charge", "MHD (Odoo)", "Menge", "Hinweis"])
    fills = {"ERROR": F_BAD, "SKIP": F_BAD, "WARN": F_WARN, "OK": F_OK}
    for row in check:
        ws.append(list(row))
        f = fills.get(row[0])
        if f:
            ws.cell(ws.max_row, 1).fill = f
    for c in ws[1]:
        c.font = Font(bold=True)
    for col, w in zip("ABCDEF", (9, 26, 16, 12, 8, 80)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


# --------------------------------------------------------------------------
# 入口
# --------------------------------------------------------------------------
def out_dir():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    d = os.path.join(root, "output", date.today().strftime("%Y%m%d"))
    os.makedirs(d, exist_ok=True)
    return d


def is_entry_sheet(wb):
    return META_SHEET in wb.sheetnames and wb[META_SHEET]["B1"].value == META_KIND


def run(path, dest=None):
    """→ (kind, [产出路径], 摘要文字)。"""
    dest = dest or out_dir()
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    wb = load_workbook(path)
    if is_entry_sheet(wb):
        blocks = read_entry(wb["Erfassung"])
        imp = os.path.join(dest, "Lot_Import_%s.xlsx" % stamp)
        chk = os.path.join(dest, "Lot_Check_%s.xlsx" % stamp)
        n, check = build_import(blocks, imp, chk)
        cnt = {k: sum(1 for c in check if c[0] == k) for k in ("ERROR", "WARN", "SKIP")}
        lines = ["Products written to import file: %d / %d" % (n, len(blocks)),
                 "Errors: %d   Warnings: %d" % (cnt["ERROR"], cnt["WARN"])]
        lines += ["  %-5s %s  %s" % (c[0], c[1], c[5]) for c in check if c[0] in ("ERROR", "WARN")]
        files = ([imp] if n else []) + [chk]
        if not n:
            lines.append("Nothing valid to import - no import file written.")
        return "import", files, "\n".join(lines)
    pickings = read_export(wb.active)
    entry = os.path.join(dest, "Lot_Entry_%s.xlsx" % stamp)
    n = build_entry(pickings, entry)
    no_ref = [pid for pid, p in pickings.items() if not p["ref"]]
    lines = ["Receipts: %d   Products: %d" % (len(pickings), n)]
    if no_ref:
        lines.append("Note: export has no 'Reference' column - receipts are labelled by "
                     "Source Document. Add Reference to the export template.")
    return "entry", [entry], "\n".join(lines)


def _pick_file():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    p = filedialog.askopenfilename(
        title="Odoo export (-> entry sheet) or filled entry sheet (-> import file)",
        filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
    root.destroy()
    return p


def _open(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: P204
        elif sys.platform == "darwin":
            os.system('open "%s"' % path)
    except Exception:
        pass


def main(argv):
    no_open = "--no-open" in argv
    args = [a for a in argv[1:] if a != "--no-open"]
    path = args[0] if args else _pick_file()
    if not path:
        print("No file selected.")
        return 1
    try:
        kind, files, summary = run(path)
    except InputError as e:
        print("ERROR: %s" % e)
        return 1
    print("Input: %s" % path)
    print("Mode: %s" % ("Odoo export -> entry sheet" if kind == "entry"
                        else "entry sheet -> import file + check list"))
    print(summary)
    for f in files:
        print("Written: %s" % f)
    if not no_open:
        _open(os.path.dirname(files[0]))
        _open(files[0])
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
