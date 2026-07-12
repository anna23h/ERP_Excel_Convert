# -*- coding: utf-8 -*-
"""订货辅助表：待发货明细表 × purchase order → 一行一品的订货决策清单。

用途（独立脚本，不进 gui.py）：
    每周收到「待发货明细表」后，对每个待订货品去 ERP purchase order 里查
    最近采购价 / 供应商 / 数量 / 当前库存，据此判断向哪家 vendor 订多少。
    本脚本把这步自动化，产出一张辅助表。订货状态 / MHD 仍人工填
    （到货量是订完才录、MHD 从 purchase 端导不出且会变化）。

数据关联：
    - 主键 = PZN。purchase order 的 Internal Reference 尾部数字段即 PZN，
      两侧都去掉 PZN- 前缀、补零到 8 位再 join（6 位 EAN 类码对不上属正常）。
    - 平台裸价来自「待发货明细表」自带的 采购订单-* / 保税 分表（purchase
      order 导出本身无裸价列）。
    - 当前库存 = purchase order 的 Product/Quantity On Hand。

复用 build_excel：_short_vendor/_vendor_map（供应商简称）、_write_simple（统一版式）。

用法：
    python3 reorder_helper.py <待发货明细表.xlsx> <purchase order.xlsx> [out.xlsx]
"""

import sys, os, re
from datetime import date, datetime

import openpyxl

from build_excel import _short_vendor, _vendor_map, _write_simple, unique_path

# ---- purchase order 列名（兼容不同导出版本）----
PO_REF      = "Order Reference"
PO_VENDOR   = "Vendor"
PO_INTERNAL = "Order Lines/Product/Internal Reference"
PO_PRICE    = "Order Lines/Unit Price"
PO_QTY      = "Order Lines/Total Quantity"
PO_ONHAND   = "Product/Quantity On Hand"
PO_DATE_CANDS = ["Order Lines/Order Date", "Order Lines/Created on",
                 "Order Lines/Order Deadline"]

# 待发货明细表里带「平台裸价」的分表（按条形码取裸价）
PRICE_SHEETS = ["采购订单-健康", "采购订单-直营", "健康-保税"]

# 采购单里伪装成供应商的客户（实为我方销售平台，属噪音，整行剔除）。
# 比 build_excel 的 "Alibaba Health" 更宽：导出实见 Alibaba Health（港）与 Alibaba.com（新加坡）两个实体，均为客户。
CUSTOMER_PAT = "alibaba"

RECENT_N = 5  # 「近期采购记录」列展示最近几笔


def norm_pzn(x):
    """条形码 / Internal Reference 归一为可比较键：去 PZN- 前缀，纯数字补零到 8 位。"""
    s = re.sub(r"^PZN[-\s]*", "", str(x).strip(), flags=re.I)
    m = re.search(r"(\d+)\s*$", s)          # 取尾部数字段（Internal Reference 如 Xxx_02807988）
    if not m:
        return s
    d = m.group(1)
    return d.zfill(8) if len(d) <= 8 else d


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def _pdate(x):
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(x)[:19] if " " in str(x) else str(x)[:10], fmt)
        except (TypeError, ValueError):
            continue
    return None


def _is_pzn(v):
    """值形态判 PZN：归一后是 8 位纯数字（PZN- 前缀 / 7~8 位数字均可）。
    避开后端商品id(12 位)、采购单号(PON…长串)、EAN(13 位) 等更长的数字串。"""
    if v is None:
        return False
    k = norm_pzn(v)
    return k.isdigit() and len(k) == 8


def _find_pzn_col(rows):
    """在一张 sheet 里按『值形态』找 PZN 主键列，不认死列头。
    返回 (列号, 数据起始行号) —— 该列 PZN 命中数最多；无命中返回 None。"""
    hits, first = {}, {}
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            if _is_pzn(v):
                hits[j] = hits.get(j, 0) + 1
                first.setdefault(j, i)
    if not hits:
        return None
    col = min((c for c in hits if hits[c] == max(hits.values())))  # 命中最多；并列取最左列
    return col, first[col]


def _find_label_col(rows, data_start, exact, loose=()):
    """在数据起始行之上的表头区找一列：先精确匹配 exact，再宽松 contains loose。
    找不到返回 None（该信息缺失，不阻断）。"""
    region = rows[:data_start] if data_start else rows[:3]
    for want in (exact, None):
        for r in region:
            for j, v in enumerate(r):
                if v is None:
                    continue
                s = str(v).strip()
                if want is not None and s == want:
                    return j
                if want is None and loose and any(k in s for k in loose):
                    return j
    return None


def load_demand(path):
    """商品清单 → [{pzn, barcode, name, need}]，对输入宽容（不认死 sheet 名/列头）。
    优先用待发货明细表的『需求汇总-HK』；缺则回退——逐 sheet 按『值形态』找含 PZN 的列，
    第一张命中的即主键列。商品名称/总需求 按表头宽松匹配，取不到就留空（None），不报错，
    故『仅一列 PZN 的清单』也能跑。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    order = (["需求汇总-HK"] if "需求汇总-HK" in wb.sheetnames else []) + \
            [s for s in wb.sheetnames if s != "需求汇总-HK"]
    for name in order:
        rows = list(wb[name].iter_rows(values_only=True))
        found = _find_pzn_col(rows)
        if found:
            c_bc, data_start = found
            break
    else:
        raise ValueError("未找到含 PZN 的列（商品清单里应有一列 PZN，如 PZN-02807988）")
    c_name = _find_label_col(rows, data_start, "商品名称", loose=("名称", "品名"))
    c_need = _find_label_col(rows, data_start, "总需求", loose=("需求",))
    out = []
    for r in rows[data_start:]:
        bc = r[c_bc] if c_bc < len(r) else None
        if not _is_pzn(bc):
            continue  # 只收 PZN 行，跳过表头残留/空行/小计
        nm = r[c_name] if (c_name is not None and c_name < len(r)) else None
        out.append({"pzn": norm_pzn(bc), "barcode": str(bc).strip(),
                    "name": str(nm).strip() if nm else "",
                    "need": _num(r[c_need]) if (c_need is not None and c_need < len(r)) else None})
    return out


def load_caps(path):
    """待发货明细表自带的采购订单/保税分表 → {pzn: 平台裸价}。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    caps = {}
    for sh in PRICE_SHEETS:
        if sh not in wb.sheetnames:
            continue
        rows = list(wb[sh].iter_rows(values_only=True))
        if not rows:
            continue
        hdr = {str(h).strip(): i for i, h in enumerate(rows[0]) if h}
        cb, cp = hdr.get("条形码"), hdr.get("平台裸价")
        if cb is None or cp is None:
            continue
        for r in rows[1:]:
            if cb < len(r) and r[cb] and cp < len(r) and r[cp] not in (None, ""):
                v = _num(r[cp])
                if v is not None:
                    caps.setdefault(norm_pzn(r[cb]), v)  # 首见为准（分表内同码裸价一致）
    return caps


def load_po(path):
    """purchase order 导出（Odoo 行式：订单头字段只在每单首行）→ {pzn: [line...]}。
    line = {po, vendor(简称), price, qty, date, onhand}。"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    hdr = {str(h).strip(): i for i, h in enumerate(rows[0]) if h is not None}
    for c in (PO_REF, PO_VENDOR, PO_INTERNAL, PO_PRICE, PO_QTY):
        if c not in hdr:
            raise ValueError(f"purchase order 导出缺列: {c}")
    c_date = next((hdr[c] for c in PO_DATE_CANDS if c in hdr), None)
    c_onhand = hdr.get(PO_ONHAND)

    # 第一遍：ffill 订单头（Vendor），收集全部 vendor 全名以建简称映射
    raw = []
    cur_ref = cur_vendor = None
    for r in rows[1:]:
        if hdr[PO_REF] < len(r) and r[hdr[PO_REF]]:
            cur_ref = r[hdr[PO_REF]]
        if hdr[PO_VENDOR] < len(r) and r[hdr[PO_VENDOR]]:
            cur_vendor = r[hdr[PO_VENDOR]]
        ref = r[hdr[PO_INTERNAL]] if hdr[PO_INTERNAL] < len(r) else None
        if not ref:
            continue
        if cur_vendor and CUSTOMER_PAT in str(cur_vendor).lower():
            continue  # Alibaba 等客户单：整行剔除（噪音，非供应商）
        raw.append((cur_ref, cur_vendor, r))
    vmap = _vendor_map({v for _, v, _ in raw if v})

    out = {}
    for cur_ref, cur_vendor, r in raw:
        key = norm_pzn(r[hdr[PO_INTERNAL]])
        out.setdefault(key, []).append({
            "po": cur_ref,
            "vendor": vmap.get(cur_vendor, cur_vendor) if cur_vendor else "",
            "price": _num(r[hdr[PO_PRICE]]) if hdr[PO_PRICE] < len(r) else None,
            "qty": _num(r[hdr[PO_QTY]]) if hdr[PO_QTY] < len(r) else None,
            "date": _pdate(r[c_date]) if (c_date is not None and c_date < len(r)) else None,
            "onhand": _num(r[c_onhand]) if (c_onhand is not None and c_onhand < len(r)) else None,
        })
    return out


COLS = ["条形码", "商品名称", "总需求", "当前库存", "需补货数", "平台裸价", "最近采购单价",
        "差价", "最近采购vendor", "最近采购数量", "最近采购日期", "近期采购记录"]
# 纯数字列居中，其余（含日期/vendor/条形码/名称/记录）左对齐下沉
NUM_COLS = {"总需求", "当前库存", "需补货数", "平台裸价", "最近采购单价", "差价", "最近采购数量"}
LEFT_COLS = set(COLS) - NUM_COLS
WIDTHS = {"商品名称": 30, "最近采购vendor": 18, "近期采购记录": 56, "条形码": 15}
# 整批都无值时整列删掉（如仅 PZN 清单没有 商品名称/总需求，也没采购订单分表给平台裸价）；
# 有值/部分有值则保留、缺处留空，列序不变。其余列（条形码/库存/采购画像）恒在。
DROP_IF_EMPTY = ["商品名称", "总需求", "需补货数", "平台裸价", "差价"]


def _round(v, n):
    return round(v, n) if v is not None else ""


def build_rows(demand, caps, po):
    rows = []
    for d in demand:
        lines = po.get(d["pzn"], [])
        dated = sorted([l for l in lines if l["date"]], key=lambda x: x["date"], reverse=True)
        # 当前库存：产品级，取任一非空（优先最近记录）
        onhand = next((l["onhand"] for l in dated if l["onhand"] is not None),
                      next((l["onhand"] for l in lines if l["onhand"] is not None), None))
        cap = caps.get(d["pzn"])
        need = d["need"]
        reorder = (need - onhand) if (need is not None and onhand is not None) else need

        if dated or lines:
            last = dated[0] if dated else lines[0]
            price = last["price"]
            diff = (cap - price) if (cap is not None and price is not None) else ""
            recent = "\n".join(
                "{d}|{po}|{v}|{q}|@{p}".format(
                    d=l["date"].strftime("%m-%d") if l["date"] else "??",
                    po=l["po"] or "?",  # PO 单号，便于回 ERP 按号查明细
                    v=l["vendor"] or "?",
                    q=("" if l["qty"] is None else f"{l['qty']:g}"),
                    p=("" if l["price"] is None else f"{l['price']:g}"))
                for l in (dated or lines)[:RECENT_N])
            rows.append([
                d["barcode"], d["name"],
                "" if need is None else int(need),
                "" if onhand is None else int(onhand),
                "" if reorder is None else int(reorder),
                _round(cap, 4), _round(price, 2), _round(diff, 2) if diff != "" else "",
                last["vendor"] or "",
                "" if last["qty"] is None else f"{last['qty']:g}",
                last["date"].strftime("%Y-%m-%d") if last["date"] else "",
                recent,
            ])
        else:
            rows.append([
                d["barcode"], d["name"],
                "" if need is None else int(need),
                "" if onhand is None else int(onhand),
                "" if reorder is None else int(reorder),
                _round(cap, 4), "", "", "无采购记录", "", "", "",
            ])
    return rows


def build(demand_path, po_path, out_path=None):
    demand = load_demand(demand_path)
    caps = load_caps(demand_path)
    po = load_po(po_path)
    rows = build_rows(demand, caps, po)

    import pandas as pd
    df = pd.DataFrame(rows, columns=COLS)
    for c in DROP_IF_EMPTY:
        if df[c].map(lambda v: v in ("", None)).all():
            df = df.drop(columns=c)
    if out_path:
        outdir = os.path.dirname(out_path) or "."
        fname = os.path.basename(out_path)
    else:
        outdir = os.path.join("output", f"{date.today():%Y%m%d}")
        fname = f"订货辅助-{date.today():%Y%m%d}.xlsx"
    os.makedirs(outdir, exist_ok=True)
    path, n = _write_simple(df, outdir, fname, left_cols=LEFT_COLS, widths=WIDTHS)

    vcol = COLS.index("最近采购vendor")
    matched = sum(1 for r in rows if r[vcol] != "无采购记录")
    return path, n, matched


def main():
    args = sys.argv[1:]
    if len(args) < 2:
        print("用法: python3 reorder_helper.py <待发货明细表.xlsx> <purchase order.xlsx> [out.xlsx]")
        return
    path, n, matched = build(args[0], args[1], args[2] if len(args) > 2 else None)
    print(f"订货辅助表已生成: {path}")
    print(f"  共 {n} 个货品，其中 {matched} 个有采购记录，{n - matched} 个无匹配")


if __name__ == "__main__":
    main()
