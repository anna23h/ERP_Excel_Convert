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
一个批次都没写 = 包装上本就没有批次，该产品跳过（核对清单里列为 NO LOT 供复查）。
版式（R2）：单内按品名排序；第 1 个批次与产品同行，多批次同列往下续写；Σ / Status 只在屏幕上、不进打印区域。
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
# 存成导出模板后 `External ID` 列名会变成 `ID`，与数据库数字 `ID` 同名、可能两列并存（ISSUES R5），
# 所以 ID 类字段不按列名定，按值形态挑：只认 `模块.名称` 形态的 External ID，纯数字列不收。
FIELDS = OrderedDict([
    ("picking_id", ("id", "External ID", "ID")),
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
# 值是 External ID 的字段（写进导入表作匹配键，写成数字 ID 导回去会新建单/行）
XID_KEYS = ("picking_id", "line_id", "move_id")
XID_RE = re.compile(r"^\w+\.\w+$")  # __export__.stock_picking_433727_9a35dba9

IMPORT_HEADER = [
    "id", "move_line_ids/id", "move_line_ids/move_id/id", "move_line_ids/product_id",
    "move_line_ids/lot_name", "move_line_ids/qty_done", "move_line_ids/product_uom_id",
    "move_line_ids/location_id", "move_line_ids/location_dest_id",
    "move_line_ids/expiration_date",
]

# MHD 合理性上限：防 `05.82` 这类笔误被当成 2082 年（ISSUES R4）
MHD_MAX_YEARS = 10
MHD_MAX = lambda today: today.replace(year=today.year + MHD_MAX_YEARS, day=min(today.day, 28))
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

    认：MM.YY（操作员习惯）/ MM.YYYY / DD.MM.YYYY / DD.MM.YY / YYYY-MM-DD / YYYY-MM，分隔符 . / - 都行，
    日、月可一位。只有月份 → 当月最后一天。

    MM.YY 无歧义（ISSUES R4）：MHD 在未来，两位年 ≥ 26 > 12，不可能被读成「日.月」的月份。
    YY.MM（`28.05`）同样无歧义，但不是操作员习惯——用户定**报错**并提示改写，只认一种习惯。
    DD.MM.YY（照抄包装全日期，ISSUES R6）：三段两位理论上也能读成 YY.MM.DD，但德式只有日.月.年一种，
    用户定**只按 DD.MM.YY 读、不提示**；笔误靠已过期 / 超 10 年检查兜底。
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
        if len(parts) == 2 and lens[1] == 2 and lens[0] <= 2:          # MM.YY
            if 1 <= nums[0] <= 12:
                return month_end(2000 + nums[1], nums[0]), None
            if 1 <= nums[1] <= 12:
                return None, "MHD %r looks like YY.MM - write MM.YY, e.g. %02d.%02d" % (
                    val, nums[1], nums[0])
        if len(parts) == 2 and lens[0] == 4 and lens[1] <= 2:          # YYYY-MM
            return month_end(nums[0], nums[1]), None
        if len(parts) == 3 and lens[2] == 4 and max(lens[:2]) <= 2:    # DD.MM.YYYY
            return date(nums[2], nums[1], nums[0]), None
        if len(parts) == 3 and max(lens) <= 2:                         # DD.MM.YY
            return date(2000 + nums[2], nums[1], nums[0]), None
        if len(parts) == 3 and lens[0] == 4 and max(lens[1:]) <= 2:    # YYYY-MM-DD
            return date(nums[0], nums[1], nums[2]), None
    except ValueError:
        return None, "MHD is not a real date: %r" % val
    return None, "MHD not recognised (use MM.YY, MM.YYYY, DD.MM.YY or DD.MM.YYYY): %r" % val


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
def _xid_col(cands, rows):
    """同名候选列里挑值为 External ID 的那一列：非空值全是 `模块.名称` 形态。都不是 → None。

    候选列全空（无数据可判）→ 取第一列，由后面的「没有入库行」报错。
    """
    filled = False
    for i in cands:
        vals = [str(r[i]).strip() for r in rows[1:] if i < len(r) and r[i] not in (None, "")]
        filled = filled or bool(vals)
        if vals and all(XID_RE.match(v) for v in vals):
            return i
    return None if filled else cands[0]


def _map_header(header, rows):
    pos, numeric = {}, []
    for key, aliases in list(FIELDS.items()) + list(OPTIONAL.items()):
        cands = [i for a in aliases for i, h in enumerate(header) if h == a]
        if not cands:
            continue
        if key in XID_KEYS:
            i = _xid_col(cands, rows)
            if i is None:
                numeric.append("%s (column %s)" % (key, " / ".join(sorted({header[c] for c in cands}))))
                continue
            pos[key] = i
        else:
            pos[key] = cands[0]
    if numeric:
        raise InputError("These columns hold database numbers, not External IDs: " + "; ".join(numeric)
                         + ". Add 'External ID' to the export template (it may show up as 'ID').")
    missing = [k for k in FIELDS if k not in pos]
    if missing:
        need = ["%s (e.g. %s)" % (k, FIELDS[k][-1]) for k in missing]
        raise InputError("Odoo export is missing columns: " + "; ".join(need))
    return pos


def read_export(ws):
    """→ OrderedDict picking_id → {ref, partner, origin, moves: OrderedDict move_id → move}。"""
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    pos = _map_header(header, rows)
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
# 可见列（2026-10-11 R2：品名打头作查找键；第 1 个批次与产品同行；去掉 Keine Charge 列）
C_NAME, C_CODE, C_SOLL, C_LOT, C_MHD, C_QTY, C_STAT = range(1, 8)
C_PRINT_LAST = C_QTY  # 打印区域到 Menge 为止；Σ / Status 只在屏幕上看
# 隐藏技术列（只在产品行上有值）。H 列留空作隔离。
TECH = ["_type", "move_id", "line_ids", "picking_id", "product", "uom", "loc_from", "loc_to", "code"]
C_TECH0 = 9  # I 起
ROW_H = 26   # 打印出来要能手写

THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_HEAD = PatternFill("solid", fgColor="1F3864")
F_PROD = PatternFill("solid", fgColor="DDE4EE")
F_INPUT = PatternFill("solid", fgColor="FFFFFF")
F_OK = PatternFill("solid", fgColor="C6EFCE")
F_BAD = PatternFill("solid", fgColor="FFC7CE")
F_WARN = PatternFill("solid", fgColor="FFEB9C")
UNLOCK = Protection(locked=False)


def _status_formula(row, end):
    """产品行上的 Σ / Status。产品块 = 本行起、到下一个有类型标记的行（下一产品 / 单头 / END）之前。

    续行与操作员插的行类型列都为空，所以块的长度用 MATCH 现找，不写死区间：
    在块内任何位置插行（含紧贴下一产品的上方）都会被算进来。
    """
    t = ws_col(C_TECH0)
    k = 'MATCH(TRUE,INDEX(${t}{n}:${t}${e}<>"",0),0)'.format(t=t, n=row + 1, e=end)
    rng = lambda c: "{c}{r}:INDEX({c}:{c},ROW()+{k}-1)".format(c=ws_col(c), r=row, k=k)
    n, s, soll = "COUNTA(%s)" % rng(C_LOT), "SUM(%s)" % rng(C_QTY), "%s%d" % (ws_col(C_SOLL), row)
    return ('=IF({n}=0,"",IF(AND({n}=1,{s}=0),"✓ = Soll",'
            'IF({s}={c},"✓ "&{s},"≠ "&{s}&" / "&{c})))').format(n=n, s=s, c=soll)


NEW_PZN_RE = re.compile(r"\b(?:new|neue?)\s*PZN\s*:?\s*(\d{7,8})\b", re.I)


def new_pzn(name):
    """品名里的 `(New PZN 19765739)` / `neue PZN: …` → 新 PZN；没有返回 None。实物上印的是它。"""
    m = NEW_PZN_RE.search(name or "")
    return m.group(1) if m else None


def pzn_rich(code, name):
    """Produkt 格：PZN 是操作员和实物比对的号，加粗大字；前缀小灰。

    品名里有 New PZN → 新 PZN 打头 + NEU，第二行小灰「alt: 原编码」（用户 2026-10-11，ISSUES R3）。
    """
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    big = InlineFont(b=True, sz=13)
    small = InlineFont(sz=9, color="777777")
    new = new_pzn(name)
    if new:
        return CellRichText(TextBlock(big, new), TextBlock(InlineFont(b=True, sz=9, color="C00000"), "  NEU"),
                            TextBlock(small, "\nalt: " + code))
    m = re.match(r"^(.*?)(\d{7,8})$", code or "")
    if not m:
        return CellRichText(TextBlock(big, code or ""))
    parts = [TextBlock(big, m.group(2))]
    if m.group(1):
        parts.insert(0, TextBlock(small, m.group(1)))
    return CellRichText(*parts)


def ws_col(i):
    from openpyxl.utils import get_column_letter
    return get_column_letter(i)


def _entry_cells(ws, r, dv_qty):
    for col in (C_LOT, C_MHD, C_QTY):
        c = ws.cell(r, col)
        c.protection = UNLOCK
        c.fill = F_INPUT
        c.font = Font(size=12)
        c.alignment = Alignment(vertical="center")
        if col != C_QTY:
            c.number_format = "@"  # 文本：防 11.2027 变小数、1E5 变科学计数
    dv_qty.add(ws.cell(r, C_QTY))


def build_entry(pickings, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Erfassung"
    heads = ["Bezeichnung", "Produkt", "Soll", "LOT", "MHD", "Menge", "Status"]
    widths = [50, 24, 7, 20, 14, 9, 16]
    for i, (h, w) in enumerate(zip(heads, widths), start=1):
        c = ws.cell(1, i, h)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        ws.column_dimensions[c.column_letter].width = w
    for j, h in enumerate(TECH):
        ws.cell(1, C_TECH0 + j, h)
        ws.column_dimensions[ws_col(C_TECH0 + j)].hidden = True
    ws.column_dimensions[ws_col(C_STAT + 1)].hidden = True
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    dv_qty = DataValidation(type="decimal", operator="greaterThan", formula1="0",
                            allow_blank=True, error="Quantity must be a number > 0",
                            showErrorMessage=True)
    ws.add_data_validation(dv_qty)

    n_rows = 1 + sum(1 + len(p["moves"]) * ENTRY_ROWS for p in pickings.values())
    end = n_rows + 1  # END 标记行
    r = 2
    status_rows = []
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
        # 单内按品名字母序：对货时按包装上的名字找
        moves = sorted(pk["moves"].items(), key=lambda kv: split_product(kv[1]["product"])[1].casefold())
        for mid, mv in moves:
            code, name = split_product(mv["product"])
            prod_row = r
            for k in range(ENTRY_ROWS):
                for col in range(1, C_STAT + 1):
                    c = ws.cell(r, col)
                    c.border = BOX
                    if k == 0 and col < C_LOT:
                        c.fill = F_PROD
                _entry_cells(ws, r, dv_qty)
                ws.row_dimensions[r].height = ROW_H
                r += 1
            for col, v, font in ((C_NAME, name, Font(bold=True, size=11)),
                                 (C_CODE, pzn_rich(code, name), None),
                                 (C_SOLL, clean_num(mv["demand"]), Font(bold=True, size=12))):
                c = ws.cell(prod_row, col, v)
                if font:
                    c.font = font
                c.alignment = Alignment(vertical="center", wrap_text=(col in (C_NAME, C_CODE)),
                                        horizontal="center" if col == C_SOLL else None)
            if len(name) > 48 or new_pzn(name):
                ws.row_dimensions[prod_row].height = ROW_H + 12
            tech = ["P", mid, "|".join(mv["lines"]), pid, mv["product"], mv["uom"],
                    mv["loc_from"], mv["loc_to"], code]
            for j, v in enumerate(tech):
                ws.cell(prod_row, C_TECH0 + j, v)
            status_rows.append(prod_row)
    assert r == end, (r, end)
    ws.cell(end, C_TECH0, "END")
    for pr in status_rows:
        st = ws.cell(pr, C_STAT, _status_formula(pr, end))
        st.alignment = Alignment(horizontal="center", vertical="center")

    col = ws_col(C_STAT)
    rng = "%s2:%s%d" % (col, col, end)
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(%s2,1)="✓"' % col], fill=F_OK))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['LEFT(%s2,1)="≠"' % col], fill=F_WARN))

    ws.protection.sheet = True
    ws.protection.insertRows = False   # False = 允许：多批次时操作员自己插行
    ws.protection.deleteRows = False
    ws.protection.formatCells = False
    ws.protection.formatRows = False
    ws.protection.autoFilter = False
    ws.print_area = "A1:%s%d" % (ws_col(C_PRINT_LAST), end - 1)
    ws.page_setup.orientation = "portrait"
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
    """→ 产品块列表：{tech..., code, name, soll, entries:[(row, lot, mhd, qty)]}。

    行类型看隐藏列 `_type`：产品行 P 本身就是第 1 个录入行；续行与操作员插的行该列为空，
    归到上方最近的产品块。
    """
    blocks, cur = [], None
    for r in range(2, ws.max_row + 1):
        t = ws.cell(r, C_TECH0).value
        if t == "H" or t == "END":
            cur = None
            continue
        if t == "P":
            tech = {TECH[j]: ws.cell(r, C_TECH0 + j).value for j in range(1, len(TECH))}
            cur = dict(tech, row=r,
                       name=ws.cell(r, C_NAME).value, soll=num(ws.cell(r, C_SOLL).value) or 0,
                       entries=[])
            blocks.append(cur)
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
    if not any(lot for _, lot, _, _ in entries):
        # 一个批次都没写 = 包装上没有批次（用户 2026-10-11 定），跳过；写了别的就提醒一句
        if any(m not in (None, "") for _, _, m, _ in entries):
            W(b["row"], "MHD entered but no lot - not imported (Odoo stores MHD on the lot). "
                        "Lot forgotten?")
        elif entries:
            W(b["row"], "Quantity entered but no lot - not imported")
        return [], issues
    seen = set()
    name_months = name_mhd_months(b["name"])
    for row, lot, mhd, qty in entries:
        if not lot:
            E(row, "Lot is empty on this row, but other lots were entered for the product")
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
            elif d > MHD_MAX(date.today()):
                E(row, "MHD %s is more than %d years ahead - typo?" % (d.strftime("%d.%m.%Y"), MHD_MAX_YEARS))
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
    """核对清单行：(状态, 品名, 编码, 批次, MHD, 数量, 说明)。"""
    head = lambda st, msg, lot="", d="", q="": (st, b["name"], b["code"], lot, d, q, msg)
    rows = [head(level, "row %d: %s" % (row, msg)) for level, row, msg in issues]
    if not any(lot for _, lot, _, _ in b["entries"]):
        rows.append(head("NO LOT", "No lot entered - skipped. Pack without lot? If not, enter it and rerun"))
        return rows
    if any(i[0] == "ERROR" for i in issues):
        rows.append(head("SKIP", "Product NOT in the import file - fix the red rows and rerun"))
        return rows
    for lot, d, q in lots:
        rows.append(head("OK", "", lot, d.strftime("%d.%m.%Y") if d else "", clean_num(q)))
    return rows


def _write_check(check, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Check"
    ws.append(["Status", "Bezeichnung", "Produkt", "Charge", "MHD (Odoo)", "Menge", "Hinweis"])
    fills = {"ERROR": F_BAD, "SKIP": F_BAD, "WARN": F_WARN, "OK": F_OK,
             "NO LOT": PatternFill("solid", fgColor="E7E6E6")}
    for row in check:
        ws.append(list(row))
        f = fills.get(row[0])
        if f:
            ws.cell(ws.max_row, 1).fill = f
    for c in ws[1]:
        c.font = Font(bold=True)
    for col, w in zip("ABCDEFG", (9, 44, 22, 16, 12, 8, 80)):
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
        lines += ["  %-5s %s  %s" % (c[0], c[2], c[6]) for c in check if c[0] in ("ERROR", "WARN")]
        no_lot = [c[2] for c in check if c[0] == "NO LOT"]
        if no_lot:
            lines.append("No lot entered (skipped, check these are really packs without lot): "
                         + ", ".join(no_lot))
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
