"""B2B 可用库存 + 仓库待打包清单（每天定时；Odoo XML-RPC 只读）。

    python3 odoo_api/b2b_stock.py                    # 产出 output/YYYYMMDD/B2B可用库存-YYYYMMDD.xlsx
    python3 odoo_api/b2b_stock.py -o /共享/目录       # 指定输出目录（也可在 config.py 配 B2B_STOCK_OUTDIR）

为什么要有它（ISSUES [B2B可用] A，2026-11-15 与用户对齐）：
  Odoo 的 Available Quantity = Quantity − Reserved Quantity，而预留是**先到先得、全自动**的——
  确认 SO 时、入库确认的同一事务里、每天 OdooBot 定时任务都会预留，不反映「人决定把货给哪张单」。
  B2B 单（批发）占了全库预留的绝大部分，没付款的单也照样占着货。销售靠「客户别称」认单，
  但 SO 上的客户是开票主体，看不出是谁。

业务约定（销售维护，脚本只读）：
  - 销售确认到账后，在 SO 上打标签 `可以打包`（`--paid-tag`）→ 这张单的预留算**已付款**。
  - 客户别称也用标签：SO 上除 `可以打包` 以外的标签一律当别称显示。
    （不用 Customer Reference：它会复制到发票 ref，并打印在报价单、发票 PDF 上。）

占用分三类：
  - 硬占用 = 天猫单（SO 号 VO/GW 开头）的预留 + 已付款 B2B 单的预留 + 其他出库单（非 SO 来源）的预留
  - 软占用 = 未付款 B2B 单的预留
  - 意向   = 近 `--quote-days` 天内创建、还没确认的报价单需求（不扣，只显示；组合装按 phantom BoM 折回基础品）
  可用两列：`Odoo可用` = 在手 − 全部预留（等于 Odoo 界面上的值）；`可用(按已付款)` = 在手 − 硬占用。

在手与预留取 `WH/Stock`（及其子库位）的 quant——与用户在 Odoo 里点开在手看到的 Quantity / Available 同源。
组合装没有自己的 quant，SO 确认时已拆成组件出库，所以这里只出现基础品。

产出三个 sheet：
  产品可用   一品一行（有在手、有预留、有报价意向的都进），有 B2B 占用的排前面
  待打包     带 `可以打包` 的单逐行：需求 / 已预留 / 缺口 / 本单货齐；打了标签但还是报价单的标「需先确认 SO」
  冲突       已付款单缺货时，同一产品有空闲库存（点 Check Availability 即可）或被未付款单预留着（需人工 Unreserve）
"""
import argparse
import os
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta

import pandas as pd
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common.xlsx import style_sheet, unique_path, write_df  # noqa: E402
from odoo_api.odoo_client import Odoo, OdooError, m2o_id, m2o_name  # noqa: E402
from odoo_api.restock import pull_kit_map  # noqa: E402
from odoo_api.stock_report import pull_in_transit, resolve_excluded_vendors, say  # noqa: E402

PAID_TAG = "可以打包"
RETAIL_PREFIXES = ("VO", "GW")
STOCK_LOCATION = "WH/Stock"

CAT_RETAIL, CAT_PAID, CAT_UNPAID, CAT_OTHER = "天猫", "已付款", "未付款", "其他出库"
HARD_CATS = {CAT_RETAIL, CAT_PAID, CAT_OTHER}

SHEET_PRODUCTS, SHEET_PACK, SHEET_CONFLICT = "产品可用", "待打包", "冲突"
PRODUCT_COLS = ["SKU", "产品", "在手", "Odoo可用", "可用(按已付款)", "天猫预留", "已付款预留",
                "未付款预留", "其他出库预留", "报价意向", "在途", "B2B占用明细", "报价明细", "备注"]
PACK_COLS = ["状态", "SO", "别称", "销售员", "SKU", "产品", "需求", "已预留", "缺口", "本单货齐", "出库单"]
CONFLICT_COLS = ["类型", "SKU", "产品", "已付款单", "别称", "缺口", "空闲库存", "占着货的未付款单", "建议操作"]


# --------------------------------------------------------------------------
# 纯计算（不连 Odoo，单测覆盖）
# --------------------------------------------------------------------------
def classify(so, paid_tag=PAID_TAG):
    """出库预留的归类。so = None（出库单不来自 SO）或 {"name", "tags"}。"""
    if not so:
        return CAT_OTHER
    if so["name"].startswith(RETAIL_PREFIXES):
        return CAT_RETAIL
    return CAT_PAID if paid_tag in so["tags"] else CAT_UNPAID


def alias_of(tags, paid_tag=PAID_TAG):
    """别称 = 除状态标签外的所有标签；一个都没有返回空串。"""
    return "、".join(t for t in tags if t != paid_tag)


def fmt_qty(q):
    return f"{q:,.0f}" if abs(q - round(q)) < 1e-9 else f"{q:,.2f}"


def occupant_label(so, qty, cat, paid_tag=PAID_TAG):
    """「SO · 状态 · 数量 · 销售员[ · 别称]」。SO 号放最前、没别称就不写（2026-11-15 用户定：
    原先别称打头，还没人打别称时每行都以「（未标别称）」开头，SO 号反而被挤到中间）。"""
    parts = [so["name"], cat, fmt_qty(qty), so["salesperson"] or "—"]
    alias = alias_of(so["tags"], paid_tag)
    if alias:
        parts.append(alias)
    return " · ".join(parts)


def product_notes(quants, moves):
    """数据问题提示 → {pid: 备注}：在手为负；出库单上的预留 ≠ quant 上的预留（Odoo 预留账不一致）。"""
    mres = defaultdict(float)
    for m in moves:
        mres[m["pid"]] += m["reserved"]
    notes = {}
    for pid in set(quants) | set(mres):
        onhand, q_res = quants.get(pid, (0.0, 0.0))
        out = []
        if onhand < 0:
            out.append("在手为负：出库先于入库确认，或入库漏录")
        if abs(mres[pid] - q_res) > 0.01:
            out.append(f"Odoo 预留账不一致（出库单 {fmt_qty(mres[pid])} / 库存 {fmt_qty(q_res)}），需 ERP 订正")
        if out:
            notes[pid] = "；".join(out)
    return notes


def build_products(quants, moves, sos, quotes, incoming, products, paid_tag=PAID_TAG, notes=None):
    """→ DataFrame（产品可用）。

    quants:   {pid: (在手, 已预留)}            WH/Stock 的 quant 合计
    moves:    [{"pid", "so_id", "demand", "reserved", "picking"}]   未完成的出库 move
    sos:      {so_id: {"name", "tags", "salesperson", "state"}}
    quotes:   [{"pid", "so_id", "qty"}]          报价单需求（已折回基础品）
    incoming: {pid: 在途}
    products: {pid: {"code", "name"}}
    """
    res = defaultdict(lambda: defaultdict(float))
    occ = defaultdict(lambda: defaultdict(float))        # pid → (so_id, cat) → 预留
    for m in moves:
        if m["reserved"] <= 0:
            continue
        so = sos.get(m["so_id"])
        cat = classify(so, paid_tag)
        res[m["pid"]][cat] += m["reserved"]
        if cat in (CAT_PAID, CAT_UNPAID):
            occ[m["pid"]][(m["so_id"], cat)] += m["reserved"]
    intent = defaultdict(float)
    intent_detail = defaultdict(lambda: defaultdict(float))
    for q in quotes:
        intent[q["pid"]] += q["qty"]
        intent_detail[q["pid"]][q["so_id"]] += q["qty"]

    pids = set(quants) | set(res) | set(intent)
    rows = []
    for pid in pids:
        onhand, reserved_all = quants.get(pid, (0.0, 0.0))
        r = res[pid]
        hard = sum(v for k, v in r.items() if k in HARD_CATS)
        detail = sorted(occ[pid].items(), key=lambda kv: (kv[0][1] != CAT_PAID, -kv[1]))
        qdetail = sorted(intent_detail[pid].items(), key=lambda kv: -kv[1])
        if not (onhand or reserved_all or r or intent[pid]):
            continue
        p = products.get(pid, {})
        rows.append({
            "SKU": p.get("code") or "", "产品": p.get("name") or f"id={pid}",
            "在手": onhand, "Odoo可用": onhand - reserved_all, "可用(按已付款)": onhand - hard,
            "天猫预留": r[CAT_RETAIL], "已付款预留": r[CAT_PAID], "未付款预留": r[CAT_UNPAID],
            "其他出库预留": r[CAT_OTHER], "报价意向": intent[pid], "在途": incoming.get(pid, 0.0),
            "B2B占用明细": "\n".join(occupant_label(sos[sid], q, cat, paid_tag) for (sid, cat), q in detail),
            "报价明细": "\n".join(occupant_label(sos[sid], q, "报价中", paid_tag) for sid, q in qdetail),
            "备注": (notes or {}).get(pid, ""),
        })
    df = pd.DataFrame(rows, columns=PRODUCT_COLS)
    if df.empty:
        return df
    df["_b2b"] = (df["已付款预留"] + df["未付款预留"] + df["报价意向"]) > 0
    df["_nosku"] = df["SKU"] == ""                       # 没有 SKU 的排在同组最后
    df = df.sort_values(["_b2b", "_nosku", "SKU", "产品"], ascending=[False, True, True, True])
    df = df.drop(columns=["_b2b", "_nosku"])
    return df.reset_index(drop=True)


def build_packing(moves, sos, quotes, products, paid_tag=PAID_TAG):
    """→ (DataFrame 待打包, {(so_id, pid): 缺口})。只看带 paid_tag 的单。"""
    lines = defaultdict(lambda: {"demand": 0.0, "reserved": 0.0, "pickings": set()})
    for m in moves:
        so = sos.get(m["so_id"])
        if classify(so, paid_tag) != CAT_PAID:
            continue
        ln = lines[(m["so_id"], m["pid"])]
        ln["demand"] += m["demand"]
        ln["reserved"] += m["reserved"]
        ln["pickings"].add(m["picking"])
    complete = defaultdict(lambda: True)
    for (sid, _), ln in lines.items():
        if ln["reserved"] + 1e-9 < ln["demand"]:
            complete[sid] = False
    rows, short = [], {}
    for (sid, pid), ln in lines.items():
        so, p = sos[sid], products.get(pid, {})
        gap = max(ln["demand"] - ln["reserved"], 0.0)
        if gap > 0:
            short[(sid, pid)] = gap
        rows.append({"状态": "货齐可打包" if complete[sid] else "等货",
                     "SO": so["name"], "别称": alias_of(so["tags"], paid_tag), "销售员": so["salesperson"] or "",
                     "SKU": p.get("code") or "", "产品": p.get("name") or f"id={pid}",
                     "需求": ln["demand"], "已预留": ln["reserved"], "缺口": gap,
                     "本单货齐": "是" if complete[sid] else "否",
                     "出库单": "、".join(sorted(x for x in ln["pickings"] if x))})
    tagged_quotes = defaultdict(float)
    for q in quotes:
        so = sos.get(q["so_id"])
        if so and paid_tag in so["tags"]:
            tagged_quotes[(q["so_id"], q["pid"])] += q["qty"]
    for (sid, pid), qty in tagged_quotes.items():
        so, p = sos[sid], products.get(pid, {})
        rows.append({"状态": "需先确认 SO", "SO": so["name"], "别称": alias_of(so["tags"], paid_tag),
                     "销售员": so["salesperson"] or "", "SKU": p.get("code") or "",
                     "产品": p.get("name") or f"id={pid}", "需求": qty, "已预留": 0.0, "缺口": qty,
                     "本单货齐": "否", "出库单": "（报价单还没确认，没有出库单，预留不了）"})
    order = {"货齐可打包": 0, "等货": 1, "需先确认 SO": 2}
    df = pd.DataFrame(rows, columns=PACK_COLS)
    if not df.empty:
        df = df.sort_values(["状态", "SO", "SKU"], key=lambda s: s.map(order) if s.name == "状态" else s)
    return df.reset_index(drop=True), short


def build_conflicts(short, quants, moves, sos, products, paid_tag=PAID_TAG):
    """已付款单缺货 × 同品的空闲库存 / 未付款预留 → DataFrame（冲突）。"""
    unpaid = defaultdict(lambda: defaultdict(float))
    for m in moves:
        if m["reserved"] > 0 and classify(sos.get(m["so_id"]), paid_tag) == CAT_UNPAID:
            unpaid[m["pid"]][m["so_id"]] += m["reserved"]
    rows = []
    for (sid, pid), gap in sorted(short.items(), key=lambda kv: (sos[kv[0][0]]["name"], kv[0][1])):
        onhand, reserved_all = quants.get(pid, (0.0, 0.0))
        free = max(onhand - reserved_all, 0.0)
        holders = sorted(unpaid[pid].items(), key=lambda kv: -kv[1])
        so, p = sos[sid], products.get(pid, {})
        base = {"SKU": p.get("code") or "", "产品": p.get("name") or f"id={pid}", "已付款单": so["name"],
                "别称": alias_of(so["tags"], paid_tag), "缺口": gap, "空闲库存": free}
        if free > 0:
            rows.append({**base, "类型": "有空闲库存没预留上", "占着货的未付款单": "",
                         "建议操作": f"在 {so['name']} 的出库单上点 Check Availability"})
        if holders and free + 1e-9 < gap:
            rows.append({**base, "类型": "被未付款单占着",
                         "占着货的未付款单": "\n".join(occupant_label(sos[h], q, CAT_UNPAID, paid_tag)
                                                for h, q in holders),
                         "建议操作": "确认要先给已付款单的话：对未付款单的出库单点 Unreserve，"
                                    f"再对 {so['name']} 点 Check Availability"})
    return pd.DataFrame(rows, columns=CONFLICT_COLS)


# --------------------------------------------------------------------------
# Odoo 拉数
# --------------------------------------------------------------------------
def pull_quants(od, location=STOCK_LOCATION):
    """WH/Stock（含子库位）的 quant → {pid: (在手, 已预留)}。"""
    loc = od.execute("stock.location", "search", [[("complete_name", "=", location)]])
    if not loc:
        raise OdooError(f"找不到库位 {location}")
    rows = od.read_group_all("stock.quant", [("location_id", "child_of", loc[0])],
                             ["quantity:sum", "reserved_quantity:sum"], ["product_id"])
    out = {m2o_id(r["product_id"]): (r.get("quantity") or 0.0, r.get("reserved_quantity") or 0.0) for r in rows}
    say(f"  在手: {location} {len(out)} 个产品，合计 {sum(v[0] for v in out.values()):,.0f} 件，"
        f"已预留 {sum(v[1] for v in out.values()):,.0f}")
    return out, loc[0]


def pull_moves(od, stock_loc):
    """未完成的出库 move（源库位在 WH/Stock 下）→ [{"pid","so_id","demand","reserved","picking"}]。"""
    dom = [("picking_code", "=", "outgoing"), ("state", "not in", ["done", "cancel", "draft"]),
           ("location_id", "child_of", stock_loc)]
    rows = od.search_read_all("stock.move", dom, ["product_id", "picking_id", "product_uom_qty"], label="出库 move")
    lines = od.search_read_all("stock.move.line", [("move_id", "in", [r["id"] for r in rows])],
                               ["move_id", "product_uom_qty"], label="预留行") if rows else []
    reserved = defaultdict(float)
    for ln in lines:
        reserved[m2o_id(ln["move_id"])] += ln["product_uom_qty"] or 0.0
    pick_ids = sorted({m2o_id(r["picking_id"]) for r in rows if r["picking_id"]})
    picks = {p["id"]: p for p in od.execute("stock.picking", "read", [pick_ids, ["name", "sale_id"]])} if pick_ids else {}
    out = []
    for r in rows:
        pk = picks.get(m2o_id(r["picking_id"]), {})
        out.append({"pid": m2o_id(r["product_id"]), "so_id": m2o_id(pk.get("sale_id")) if pk.get("sale_id") else None,
                    "demand": r["product_uom_qty"] or 0.0, "reserved": reserved[r["id"]],
                    "picking": pk.get("name", "")})
    say(f"  未完成出库: {len(pick_ids)} 张出库单 / {len(out)} 条 move，预留合计 {sum(reserved.values()):,.0f}")
    return out


def pull_quotes(od, since, kit_map):
    """近期报价单（draft/sent）的库存类产品需求，组合装折回基础品 → ([{"pid","so_id","qty"}], [so_id])。"""
    sos = od.search_read_all("sale.order", [("state", "in", ["draft", "sent"]), ("create_date", ">=", since)],
                             ["name"], label="报价单")
    if not sos:
        return [], []
    lines = od.search_read_all("sale.order.line", [("order_id", "in", [s["id"] for s in sos]),
                                                   ("display_type", "=", False),
                                                   ("product_id.type", "in", ["product", "consu"])],
                               # consu 也取：组合装（phantom 套件）是 consu，要靠它折回基础品；
                               # 普通消耗品折不回去，只会多几行在手为 0 的意向，无害
                               ["order_id", "product_id", "product_uom_qty"], label="报价行")
    out = []
    for ln in lines:
        pid, qty = m2o_id(ln["product_id"]), ln["product_uom_qty"] or 0.0
        for base, k in (kit_map.get(pid) or [(pid, 1.0)]):
            out.append({"pid": base, "so_id": m2o_id(ln["order_id"]), "qty": qty * k})
    say(f"  报价意向: {since} 起创建、未确认的报价单 {len(sos)} 张 / {len(lines)} 行")
    return out, [s["id"] for s in sos]


def pull_sos(od, ids):
    """→ {so_id: {"name","tags","salesperson","state"}}。"""
    ids = sorted(i for i in ids if i)
    if not ids:
        return {}
    rows = od.execute("sale.order", "read", [ids, ["name", "tag_ids", "user_id", "state"]])
    tag_ids = sorted({t for r in rows for t in r["tag_ids"]})
    tags = {t["id"]: t["name"] for t in od.execute("crm.tag", "read", [tag_ids, ["name"]])} if tag_ids else {}
    return {r["id"]: {"name": r["name"], "tags": [tags[t] for t in r["tag_ids"] if t in tags],
                      "salesperson": m2o_name(r["user_id"]) if r["user_id"] else "", "state": r["state"]}
            for r in rows}


def pull_products(od, pids):
    ids = sorted(pids)
    out = {}
    for i in range(0, len(ids), 1000):
        for r in od.execute("product.product", "read", [ids[i:i + 1000], ["default_code", "name"]]):
            out[r["id"]] = {"code": r["default_code"] or "", "name": r["name"] or ""}
    return out


# --------------------------------------------------------------------------
# 输出
# --------------------------------------------------------------------------
def write_book(path, sheets):
    wb = Workbook()
    wb.remove(wb.active)
    left = {"SKU", "产品", "B2B占用明细", "报价明细", "备注", "别称", "销售员", "SO", "出库单", "已付款单",
            "占着货的未付款单", "建议操作", "类型", "状态"}
    for name, df, freeze in sheets:
        ws = wb.create_sheet(name)
        write_df(ws, df)
        style_sheet(ws, len(df.columns), left_cols=left)
        ws.freeze_panes = freeze
        if len(df):
            ws.auto_filter.ref = ws.dimensions
    path = unique_path(path)
    wb.save(path)
    return path


def run(outdir=None, quote_days=90, transit_days=60, exclude_vendors=None, paid_tag=PAID_TAG, timeout=180):
    today = date.today()
    od = Odoo.connect(timeout=timeout)
    say(f"B2B 可用库存 {datetime.now():%Y-%m-%d %H:%M}（已付款标签「{paid_tag}」，报价意向取近 {quote_days} 天）")
    quants, stock_loc = pull_quants(od)
    moves = pull_moves(od, stock_loc)
    kit_map = pull_kit_map(od)
    quotes, quote_so_ids = pull_quotes(od, (today - timedelta(days=quote_days)).isoformat(), kit_map)
    sos = pull_sos(od, {m["so_id"] for m in moves} | set(quote_so_ids))
    vend_ids = resolve_excluded_vendors(od, exclude_vendors)
    incoming, _ = pull_in_transit(od, vend_ids, since=(today - timedelta(days=transit_days)).isoformat())

    # 核对：出库单上的预留应等于 quant 上的预留。2026-11-15 实测 2 个产品出库单比 quant 多 374，
    # 是 Odoo 预留账本身不一致（不是脚本口径），逐品写进「备注」，不替它抹平。
    notes = product_notes(quants, moves)
    bad = [pid for pid, n in notes.items() if "预留账" in n]
    if bad:
        say(f"  ⚠ {len(bad)} 个产品 Odoo 预留账不一致（出库单预留 ≠ quant 预留），已写进「备注」")

    products = pull_products(od, set(quants) | {m["pid"] for m in moves} | {q["pid"] for q in quotes})
    dfp = build_products(quants, moves, sos, quotes, incoming, products, paid_tag, notes)
    dfk, short = build_packing(moves, sos, quotes, products, paid_tag)
    dfc = build_conflicts(short, quants, moves, sos, products, paid_tag)

    if outdir is None:
        from common import localconf
        outdir = localconf.get("B2B_STOCK_OUTDIR") or os.path.join("output", f"{today:%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)
    path = write_book(os.path.join(outdir, f"B2B可用库存-{today:%Y%m%d}.xlsx"),
                      [(SHEET_PRODUCTS, dfp, "C2"), (SHEET_PACK, dfk, "C2"), (SHEET_CONFLICT, dfc, "A2")])
    n_paid = len({sid for sid, so in sos.items() if classify(so, paid_tag) == CAT_PAID})
    say(f"\n✓ {path}")
    say(f"  产品 {len(dfp)} 个（有 B2B 占用或报价意向的 "
        f"{int(((dfp['已付款预留'] + dfp['未付款预留'] + dfp['报价意向']) > 0).sum()) if len(dfp) else 0} 个）")
    if len(dfp):
        say("  预留: " + "  ".join(f"{c} {dfp[c].sum():,.0f}" for c in ("天猫预留", "已付款预留", "未付款预留", "其他出库预留")))
    say(f"  带「{paid_tag}」的单 {n_paid} 张；待打包 {len(dfk)} 行"
        + (f"（{'、'.join(f'{k} {v}' for k, v in dfk['状态'].value_counts().items())}）" if len(dfk) else ""))
    say(f"  冲突 {len(dfc)} 条；RPC 调用 {od.call_count} 次")
    return path, dfp, dfk, dfc


def main():
    ap = argparse.ArgumentParser(description="B2B 可用库存 + 待打包清单（Odoo 只读）")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 config.py 的 B2B_STOCK_OUTDIR，否则 output/YYYYMMDD）")
    ap.add_argument("--quote-days", type=int, default=90,
                    help="报价意向只取近几天内创建的报价单（默认 90；更老的多是陈旧单）")
    ap.add_argument("--transit-days", type=int, default=60, help="在途只认近几天内创建的入库 move（默认 60）")
    ap.add_argument("--paid-tag", default=PAID_TAG, help=f"表示已付款的 SO 标签（默认「{PAID_TAG}」）")
    ap.add_argument("--exclude-vendor", default=None,
                    help="在途排除这些供应商（名字子串，逗号分隔）。默认读 config.py 的 ODOO_EXCLUDE_VENDORS")
    ap.add_argument("--timeout", type=int, default=180)
    a = ap.parse_args()
    if a.exclude_vendor is None:
        from common import localconf
        exclude_vendors = list(localconf.get("ODOO_EXCLUDE_VENDORS", []) or [])
    else:
        exclude_vendors = [x.strip() for x in a.exclude_vendor.split(",") if x.strip()]
    try:
        run(a.outdir, a.quote_days, a.transit_days, exclude_vendors, a.paid_tag, a.timeout)
    except OdooError as e:
        raise SystemExit(f"✗ {e}")


if __name__ == "__main__":
    main()
