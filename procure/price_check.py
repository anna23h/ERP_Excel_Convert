"""发票售价 vs 采购价比价表：逐张客户发票，每个产品行对上 SO 价和最近一张采购价。

用法：
    python3 -m procure.price_check INV/2026/11/0069 INV/2026/11/0070
    python3 -m procure.price_check S04084 S04094          # 取该 SO 下已过账的客户发票
    python3 -m procure.price_check S04084 -o /tmp/x

固定口径（2026-11-06 用户定，ISSUES [比价]）：
  1. **两边一律比去税后的净价。**本库销售税（19% / 7% Umsatzsteuer）是
     `price_include=True`，`price_unit` 是**含税价**；采购税（Vorsteuer）不含税。
     直接拿两个 `price_unit` 比会把毛利高估一截（2026-11-06 第一版就是这么错的）。
     所以净价一律取 Odoo 自己算好的 `price_subtotal ÷ 数量`：折扣、含不含税都已在里面，
     不在脚本里重算税。
  2. 表的第一列是 PZN，第二列是品名，方便按 PZN 查。
  3. 采购价取**发票日当天或之前**最近一张已确认的 PO（`state ∈ {purchase, done}`，
     剔除 `common/vendor.is_noise_vendor`）。不用 SO 创建日当截止日：这类单常常是
     先接单、后采购（2026-11-06 两张实测单的货全是 SO 创建之后才进的）。
     也不用「不限日期」：否则过几天再跑，会取到之后才下的 PO，结果会变。

护栏：每张发票 Σ 产品行净额必须等于 `amount_untaxed`，对不上就停——说明有不带产品的
金额行（运费、手工行）没进表，比出来的合计不可信。

⚠ 标记两种情况：开票净价 ≠ SO 净价；毛利 ≤ 0。只标出来，**不替人判断该按哪个价**。
"""
import argparse
import datetime as dt
import os
import re
import sys

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from common.vendor import is_noise_vendor
from common.xlsx import style_sheet, unique_path
from odoo_api.odoo_client import Odoo, OdooError, m2o_id, m2o_name

SO_RE = re.compile(r"^S\d+$")
PO_CODE = re.compile(r"^\s*(P\d+)")          # 同 po_price：PO 显示名后面挂着供应商发票号，只留 P 编号
PZN_TAIL = re.compile(r"(\d{7,8})$")

COLS = ["PZN", "产品", "税率", "数量", "SO 净价", "开票净价", "采购价（PO）",
        "单件毛利", "毛利率", "采购单号", "采购日期", "备注"]
PRICE_COLS = ("SO 净价", "开票净价", "采购价（PO）", "单件毛利")
WIDTHS = {"PZN": 12, "产品": 50, "备注": 30}
#: 净单价比较容差：1 分以内算一致。SO 挂 19%、发票改挂 7% 时，两边含税价各自舍入到分，
#: 折回净价会差 0.008 左右（8.94/1.19 vs 8.03/1.07）——那是换税率的舍入，不是改价。
TOL = 0.01


def say(msg):
    print(msg, flush=True)


# --------------------------------------------------------------------------
# 纯函数（单测覆盖）
# --------------------------------------------------------------------------
def pzn_of(prod):
    """产品 → PZN。优先 product 的 `pzn` 字段，空了才取 default_code 末段 7~8 位数字。"""
    p = (prod.get("pzn") or "").strip()
    if p:
        return p
    m = PZN_TAIL.search(prod.get("default_code") or "")
    return m.group(1) if m else (prod.get("default_code") or "")


def unit_net(price_unit, discount, tax_ids, taxes, subtotal=None, qty=None):
    """单行 → 去税净单价。

    按单价算（price_unit × (1 − 折扣) ÷ (1 + 含税税率之和)），**不用 price_subtotal ÷ 数量**：
    后者在行级先四舍五入到分，同一个单价在 192 件和 40 件上除出来会差 0.01，
    SO 与发票数量不同就满表假 ⚠（2026-11-06 第一版实测）。
    税里有非百分比类型（固定金额等）时公式不成立，退回 subtotal ÷ qty。
    """
    base = (price_unit or 0.0) * (1 - (discount or 0.0) / 100)
    rate = 0.0
    for t in tax_ids:
        tx = taxes.get(t)
        if tx is None or tx.get("amount_type", "percent") != "percent":
            return subtotal / qty if qty else base
        if tx["price_include"]:
            rate += tx["amount"] / 100
    return base / (1 + rate)


def tax_label(tax_ids, taxes):
    return "+".join(taxes[t]["label"] for t in tax_ids if t in taxes) or "—"


def strip_code(display):
    """`[Emser_08491730] Emser Inh.Lsg. 60St` → `Emser Inh.Lsg. 60St`，品名多行只取首行。"""
    s = re.sub(r"^\[[^\]]*\]\s*", "", display or "").strip()
    return s.splitlines()[0].strip() if s else ""


def build_row(pzn, name, tax, qty, so_net, inv_net, po, so_tax=None):
    """一个发票行 → 表上一行。po 为 None 表示发票日前没有采购记录。

    so_tax 与发票税率不同时单独标出：含税价口径下，SO 挂 19%、发票挂 7%，
    含税单价会差 ~10% 而净价一分不差（2026-11-06 S04094 实测）。只看含税价会误判成改了价。
    """
    notes = []
    if so_tax is not None and so_tax != tax:
        notes.append(f"⚠ SO 税率 {so_tax} ≠ 开票税率 {tax}")
    row = {"PZN": pzn, "产品": name, "税率": tax, "数量": qty,
           "SO 净价": None if so_net is None else round(so_net, 2),
           "开票净价": round(inv_net, 2)}
    if so_net is None:
        notes.append("没有对应的 SO 行")
    elif abs(so_net - inv_net) > TOL + 1e-9:
        notes.append(f"⚠ 开票价 ≠ SO 价（{(inv_net / so_net - 1) * 100:+.1f}%）" if so_net
                     else "⚠ 开票价 ≠ SO 价")
    if po is None:
        notes.append("⚠ 发票日前无采购记录")
    else:
        margin = inv_net - po["net"]
        row.update({"采购价（PO）": round(po["net"], 2), "单件毛利": round(margin, 2),
                    "毛利率": round(margin / inv_net, 4) if inv_net else None,
                    "采购单号": po["code"], "采购日期": po["date"]})
        if margin <= 0:
            notes.append("⚠ 毛利 ≤ 0")
    row["备注"] = "；".join(notes)
    row["_flag"] = any(n.startswith("⚠") for n in notes)
    return row


def totals(rows):
    """→ (Σ 开票净额, Σ 采购成本, 有无采购价缺失)。采购价缺的行不进成本，单独报。"""
    sales = sum(r["_inv_subtotal"] for r in rows)       # 合计用 Odoo 行净额，和发票 amount_untaxed 对得上
    cost = sum(r["_po_net"] * r["数量"] for r in rows if r.get("_po_net") is not None)
    return sales, cost, any(r.get("_po_net") is None for r in rows)


# --------------------------------------------------------------------------
# 取数
# --------------------------------------------------------------------------
def resolve_invoices(od, refs):
    """INV 号 / SO 号 → 已过账的客户发票列表（按名字去重、保序）。"""
    inv_names, so_names = [], []
    for r in refs:
        (so_names if SO_RE.match(r) else inv_names).append(r)
    fields = ["name", "invoice_origin", "partner_id", "invoice_date", "state",
              "move_type", "amount_untaxed"]
    moves = []
    if inv_names:
        got = od.search_read_all("account.move", [("name", "in", inv_names)], fields)
        missing = [n for n in inv_names if n not in {m["name"] for m in got}]
        if missing:
            raise OdooError(f"找不到发票：{missing}（发票号要写全，如 INV/2026/11/0069）")
        moves += got
    if so_names:
        sos = od.search_read_all("sale.order", [("name", "in", so_names)], ["name", "invoice_ids"])
        missing = [n for n in so_names if n not in {s["name"] for s in sos}]
        if missing:
            raise OdooError(f"找不到销售单：{missing}（单号大小写敏感）")
        ids = [i for s in sos for i in s["invoice_ids"]]
        got = od.search_read_all("account.move", [("id", "in", ids)], fields) if ids else []
        for s in sos:
            mine = [m for m in got if m["id"] in s["invoice_ids"]]
            ok = [m for m in mine if m["state"] == "posted" and m["move_type"] == "out_invoice"]
            if not ok:
                say(f"  ⚠ {s['name']} 下没有已过账的客户发票"
                    + (f"（另有 {len(mine)} 张草稿/退款单未计入）" if mine else "") + "，跳过")
            elif len(ok) < len(mine):
                say(f"  · {s['name']} 有 {len(mine) - len(ok)} 张草稿/退款单未计入")
            moves += ok
    seen, out = set(), []
    for m in moves:
        if m["name"] in seen:
            continue
        seen.add(m["name"])
        if m["move_type"] != "out_invoice" or m["state"] != "posted":
            say(f"  ⚠ {m['name']} 是 {m['move_type']} / {m['state']}，不是已过账的客户发票，照样列出")
        out.append(m)
    return out


def last_po(od, pid, date, taxes):
    """发票日当天或之前最近一张已确认的采购行 → {net, code, date}；没有返回 None。

    采购价同样走去税净价：PO 的税若是含税价，price_unit 就不是净价（本库的 Vorsteuer 不含税，
    但不靠这个假设）。
    """
    rows = od.execute("purchase.order.line", "search_read",
                      [[("product_id", "=", pid), ("state", "in", ["purchase", "done"]),
                        ("date_order", "<=", f"{date} 23:59:59")],
                       ["partner_id", "product_qty", "price_unit", "price_subtotal",
                        "taxes_id", "date_order", "order_id"]],
                      {"order": "date_order desc, id desc", "limit": 20})
    for r in rows:
        if is_noise_vendor(m2o_name(r["partner_id"])):
            continue
        name = m2o_name(r["order_id"]) or ""
        m = PO_CODE.match(name)
        ensure_taxes(od, taxes, r["taxes_id"])
        return {"net": unit_net(r["price_unit"], 0, r["taxes_id"], taxes,
                                r["price_subtotal"], r["product_qty"]),
                "code": m.group(1) if m else name, "date": r["date_order"][:10]}
    return None


def build_invoice(od, inv, taxes):
    """一张发票 → 表行。taxes 会就地补上 SO 行用到、发票上没出现的税。"""
    lines = od.search_read_all(
        "account.move.line", [("move_id", "=", inv["id"]), ("exclude_from_invoice_tab", "=", False)],
        ["product_id", "name", "quantity", "price_unit", "discount", "price_subtotal", "tax_ids",
         "sale_line_ids", "display_type", "sequence"])
    lines.sort(key=lambda l: (l.get("sequence") or 0, l["id"]))
    stray = [l for l in lines if not l["product_id"] and not l.get("display_type")
             and abs(l["price_subtotal"] or 0) > TOL]
    lines = [l for l in lines if l["product_id"] and not l.get("display_type")]

    pids = sorted({m2o_id(l["product_id"]) for l in lines})
    prods = {p["id"]: p for p in od.execute(
        "product.product", "read", [pids, ["default_code", "pzn", "name"]])} if pids else {}
    sol_ids = sorted({i for l in lines for i in l["sale_line_ids"]})
    sols = {s["id"]: s for s in od.execute(
        "sale.order.line", "read",
        [sol_ids, ["product_uom_qty", "price_unit", "discount", "price_subtotal", "tax_id"]])} if sol_ids else {}
    ensure_taxes(od, taxes, [t for l in lines for t in l["tax_ids"]]
                 + [t for s_ in sols.values() for t in s_["tax_id"]])
    pos = {pid: last_po(od, pid, inv["invoice_date"], taxes) for pid in pids}

    rows = []
    for l in lines:
        pid = m2o_id(l["product_id"])
        prod = prods.get(pid, {})
        inv_net = unit_net(l["price_unit"], l["discount"], l["tax_ids"], taxes,
                           l["price_subtotal"], l["quantity"])
        so = next((sols[i] for i in l["sale_line_ids"] if i in sols), None)
        so_net = unit_net(so["price_unit"], so["discount"], so["tax_id"], taxes,
                          so["price_subtotal"], so["product_uom_qty"]) if so else None
        r = build_row(pzn_of(prod), (prod.get("name") or strip_code(m2o_name(l["product_id"]))).strip(),
                      tax_label(l["tax_ids"], taxes), l["quantity"], so_net, inv_net, pos[pid],
                      tax_label(so["tax_id"], taxes) if so else None)
        r["_inv_subtotal"] = l["price_subtotal"] or 0.0
        r["_po_net"] = pos[pid]["net"] if pos[pid] else None
        rows.append(r)

    total = sum(l["price_subtotal"] or 0 for l in lines)
    if stray or abs(total - inv["amount_untaxed"]) > 0.01:
        detail = "；".join(f"{l['name']!r} {l['price_subtotal']}" for l in stray) or "无"
        raise OdooError(
            f"{inv['name']}：产品行净额合计 {total:.2f} ≠ 发票不含税金额 {inv['amount_untaxed']:.2f}。"
            f"不带产品的金额行：{detail}。比出来的合计不可信，先查明再跑。")
    return rows


def ensure_taxes(od, taxes, ids):
    """把 ids 里还没读过的税读进 taxes（就地）。税率与含不含税一律现查，不硬编码。"""
    need = sorted({t for t in ids if t not in taxes})
    if not need:
        return
    for t in od.execute("account.tax", "read", [need, ["name", "amount", "amount_type", "price_include"]]):
        taxes[t["id"]] = {"label": f"{t['amount']:g}%", **t}


# --------------------------------------------------------------------------
# 输出
# --------------------------------------------------------------------------
def fmt(v, pct=False):
    if v is None or v == "":
        return "—"
    if pct:
        return f"{v * 100:.1f}%"
    if isinstance(v, float):
        return f"{v:.2f}"
    return str(v)


def markdown(inv, rows):
    head = f"### {inv['name']}（{inv['invoice_origin'] or '—'}，{m2o_name(inv['partner_id'])}，" \
           f"发票日 {inv['invoice_date']}）"
    out = [head, "", "| " + " | ".join(COLS) + " |", "|" + "---|" * len(COLS)]
    for r in rows:
        cells = []
        for c in COLS:
            v = r.get(c)
            if c == "开票净价" and r["备注"].find("开票价 ≠") >= 0:
                cells.append(fmt(v) + " ⚠")
            elif c == "数量":
                cells.append(f"{v:g}")
            else:
                cells.append(fmt(v, pct=(c == "毛利率")))
        out.append("| " + " | ".join(cells) + " |")
    s, c, gap = totals(rows)
    out.append(f"| | **合计** | | | | **{s:,.2f}** | **{c:,.2f}** | **{s - c:+,.2f}** | "
               f"**{(s - c) / s * 100 if s else 0:.1f}%** | | | {'⚠ 有行缺采购价，成本未计' if gap else ''} |")
    return "\n".join(out)


def build_sheet(ws, inv, rows):
    ws.append(COLS)
    for r in rows:
        ws.append([r.get(c) for c in COLS])
    s, c, gap = totals(rows)
    ws.append(["", "合计", "", "", "", round(s, 2), round(c, 2), round(s - c, 2),
               round((s - c) / s, 4) if s else None, "", "",
               "有行缺采购价，成本未计" if gap else ""])
    for i in range(2, ws.max_row + 1):
        for col in PRICE_COLS:
            ws.cell(i, COLS.index(col) + 1).number_format = "0.00"
        ws.cell(i, COLS.index("毛利率") + 1).number_format = "0.0%"
    ws.freeze_panes = "C2"        # PZN + 品名随横向滚动留在眼前
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{max(2, ws.max_row - 1)}"
    style_sheet(ws, len(COLS), left_cols={"产品", "备注"}, small_cols={"备注"},
                widths=WIDTHS, header_row=1, auto_align=True)


def safe(s):
    return re.sub(r'[\\/:*?"<>|\[\]]', "-", s)


def run(refs, outdir=None):
    outdir = outdir or os.path.join("output", f"{dt.date.today():%Y%m%d}")
    od = Odoo.connect()
    say("· 解析发票")
    invs = resolve_invoices(od, refs)
    if not invs:
        raise OdooError("没有可比的已过账客户发票。")
    taxes = {}
    wb = Workbook()
    wb.remove(wb.active)
    blocks = []
    for inv in invs:
        rows = build_invoice(od, inv, taxes)
        blocks.append(markdown(inv, rows))
        build_sheet(wb.create_sheet(safe(inv["name"])[:31]), inv, rows)
        s, c, _ = totals(rows)
        flagged = sum(r["_flag"] for r in rows)
        say(f"  {inv['name']}  {len(rows)} 行  净额 {s:,.2f} = 发票不含税 ✓  "
            f"毛利 {s - c:+,.2f}" + (f"  ⚠ {flagged} 行要看" if flagged else ""))
    os.makedirs(outdir, exist_ok=True)
    path = unique_path(os.path.join(outdir, f"比价-{'+'.join(safe(r) for r in refs)}-{dt.date.today():%Y%m%d}.xlsx"))
    wb.save(path)
    say("")
    say("\n\n".join(blocks))
    say("")
    say(f"✓ {path}")
    return path


def main():
    ap = argparse.ArgumentParser(description="客户发票 → 售价 vs 最近采购价比价表（去税净价，首列 PZN）")
    ap.add_argument("refs", nargs="+", help="发票号（INV/2026/11/0069）或销售单号（S04084），可混用")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    a = ap.parse_args()
    try:
        run(a.refs, a.outdir)
    except OdooError as e:
        say(f"✗ {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
