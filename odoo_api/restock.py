"""备货汇总（每周三）：VO/GW 零售销量 → 建议数量 + 供应商 + 上次采购价（Odoo XML-RPC 只读）。

    python3 odoo_api/restock.py                         # 默认口径，产出 output/YYYYMMDD/备货汇总-YYYYMMDD.xlsx
    python3 odoo_api/restock.py --until 2026-11-04      # 指定截止日
    python3 odoo_api/restock.py --violations 供应商违规数据.xlsx   # 多出一列 缺货罚款EUR

取代 `stock_report.py` 成为周频定时入口（ISSUES [备货提醒] A，2026-11-08 用户定：库存周报已搁置，
合并改造后纳入使用；它原来的三张附表不再出）。拉数函数直接复用 stock_report。

口径（2026-11-08 与用户逐条对齐）：
  - 销量只算**订单号以 VO / GW 开头**的零售单（`=like 'VO%'`，前缀匹配；stock_report 的
    `--ref-contains` 是子串匹配，不适用）；SW / S0 批量单自然不在内。状态 sale / done。
  - **组合装按 phantom BoM 折回基础品**：套件卖 1 件 = 组件 × BoM 数量。2026-11-08 实测套件占
    零售件数 22%，不折算会把基础品的需求低估、并把同一实物拆成几行。有后缀却没 BoM 的品不猜，列进「待核」。
  - 窗口近 12 周（跨 13 个自然周）。**有销周数 ≥ 6 才进汇总**（常销品）；偶发长尾一律不进，缺货也不列。
  - 建议数量 = ⌈备货基数 × 备货周数 − 在手 − 在途⌉，不足 0 记 0。备货基数（ISSUES [备货提醒] B，2026-11-08 用户定）：
    · 默认 = 近 4 周周均 × `--avg-factor`（1.2）。最初用 12 周峰值周，峰值/周均中位 3.3×，需补 173/241 偏高被否。
    · 近 7 天 < 12 周周均 × `--drop-ratio`（50%）→ 「近期下降」，基数 = 近 4 周周均（不乘系数）。
      **不看在手**：在手 0 ≠ 断货——天猫链接靠虚拟库存在线、有单 48h 内采购发出，零销多因下架。
    · 近 4 周周均 < `--min-recent-avg`（3 件）→ 「低量·有单再补」，建议 0（不稳定品业务上有单再补）。
    · **扣在途**：下单若不扣路上的货会重复订（与 stock_report「派生指标不扣在途」口径不同、不冲突——
      那边判断的是此刻会不会挨罚）。在途只认近 `--transit-days` 天内创建的入库 move，老残单不算。
    · 趋势（近 4 周周均 vs 之前周均）单列一栏供判断，**不进公式**。
  - 备货周数按采购形态分档：近 12 个月采购里，单张 PO 超过 `--bulk-qty`（100）件的「大单」
    量占该品采购总量 ≥ `--bulk-share`（50%）→ 大批量（直接对厂家 / wholesaler）→ 3 周；否则 1 周。
    单张 PO 超过 `--suspect-qty`（5000）件疑录错：不参与分档、在「待核」点名。
  - 供应商：主列最近一次采购的供应商 + 去税净价；并列窗口内最低净价那家。
    近 12 个月只出现过一家供应商的 → 「询价单供应商」（该供应商走 Odoo 询价单下单，`quote_fill.py` 用它）。
    另算近 `--low-months`（3）个月的最低净价那家 →「近期最低价供应商」+ 该家窗口内最近一次净价；
    本周需补的品也进这家的询价单（ISSUES [备货提醒] D，2026-11-12 用户定，对所有供应商通用；
    多家并列最低价则留空，回 A3000 比价）。
  - 采购行剔除 `common/vendor.is_noise_vendor` 与 config `ODOO_EXCLUDE_VENDORS`（技术性映射单）。

产出（默认 output/YYYYMMDD/）：
    备货汇总-YYYYMMDD.xlsx   三个 sheet：备货汇总（常销品，需补的在前）/ 询价单供应商 / 待核
"""
import argparse
import math
import os
import sys
from collections import defaultdict
from datetime import date, timedelta

import pandas as pd
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common.po import _po_base_sku  # noqa: E402
from common.vendor import is_noise_vendor  # noqa: E402
from common.xlsx import style_sheet, unique_path, write_df  # noqa: E402
from odoo_api.odoo_client import Odoo, OdooError, m2o_id, m2o_name  # noqa: E402
from odoo_api.stock_report import (pull_in_transit, pull_qty_available,  # noqa: E402
                                   read_violations, resolve_excluded_vendors, say)

SOLD_STATES = ["sale", "done"]
PO_STATES = ["purchase", "done"]
PREFIXES = ("VO", "GW")
RECENT_DAYS = 28                 # 近 4 周（基数、趋势）
LAST_DAYS = 7                    # 近 7 天（下降判定）

MAIN_SHEET, QUOTE_SHEET, CHECK_SHEET = "备货汇总", "询价单供应商", "待核"
COL_QUOTE_VENDOR = "询价单供应商"
COL_SUGGEST = "建议数量"
COL_LAST_NET = "最近一次净价"
COL_LOW_VENDOR = "近期最低价供应商"      # 近 --low-months 个月最低净价那家（并列则空）
COL_LOW_NET = "近期最低净价"
COL_LOW_LAST = "近期该家最近净价"        # 该家窗口内最近一次净价：放宽进询价单的品用它当单价


# --------------------------------------------------------------------------
# 纯计算（不连 Odoo，单测覆盖）
# --------------------------------------------------------------------------
def explode(pid, kit_map):
    """产品 → [(基础品 id, 系数)]。非套件返回 [(pid, 1.0)]。"""
    return kit_map.get(pid) or [(pid, 1.0)]


def roll_sales(rows, kit_map):
    """sale.report 按 (product, week) 的行 [(pid, 周标签, 件数)] → {基础品: {周标签: 件数}}。"""
    out = defaultdict(lambda: defaultdict(float))
    for pid, wk, q in rows:
        for base, k in explode(pid, kit_map):
            out[base][wk] += (q or 0) * k
    return out


def sales_stats(weeks_qty, recent_qty, n_days, recent_days=RECENT_DAYS):
    """单个基础品 → dict(销量, 峰值周销量, 有销周数, 周均销量, 近4周周均, 趋势)。

    周均按窗口实际天数折周（12 周窗口跨 13 个周标签，按标签数除会低估）。
    趋势 = 近 4 周周均 ÷ 窗口内其余天的周均 − 1；之前没卖过记 None（新起量，没法比）。
    """
    total = sum(weeks_qty.values())
    peak = max(weeks_qty.values(), default=0.0)
    active = sum(1 for q in weeks_qty.values() if q > 0)
    prev_days = n_days - recent_days
    prev_avg = (total - recent_qty) / prev_days * 7 if prev_days > 0 else 0
    recent_avg = recent_qty / recent_days * 7
    trend = recent_avg / prev_avg - 1 if prev_avg > 0 else None
    return {"销量": total, "峰值周销量": peak, "有销周数": active,
            "周均销量": total / n_days * 7 if n_days else 0, "近4周周均": recent_avg, "趋势": trend}


def demand_base(recent_avg, last7, avg12, avg_factor=1.2, drop_ratio=0.5, min_recent_avg=3):
    """→ (备货基数（件/周）, 口径说明)。顺序：低量 → 近期下降 → 默认。"""
    if recent_avg < min_recent_avg:
        return 0.0, "低量·有单再补"
    if avg12 > 0 and last7 < avg12 * drop_ratio:
        return recent_avg, "近期下降·近4周周均"
    return recent_avg * avg_factor, f"近4周周均×{avg_factor:g}"


def classify(po_qty, bulk_qty=100, bulk_share=0.5, suspect_qty=5000):
    """单个基础品各 PO 的件数 {PO: 件数} → (是否大批量, 大单量占比, 疑录错 PO 列表)。

    同一 PO 里同品多行先合计（po_qty 已按 PO 汇总）。疑录错的 PO 不进分子也不进分母。
    """
    suspect = sorted(p for p, q in po_qty.items() if q > suspect_qty)
    clean = {p: q for p, q in po_qty.items() if p not in suspect}
    total = sum(clean.values())
    if total <= 0:
        return False, None, suspect
    share = sum(q for q in clean.values() if q > bulk_qty) / total
    return share >= bulk_share, share, suspect


def suggest_qty(base, cover_weeks, onhand, incoming):
    """⌈备货基数 × 备货周数 − 在手 − 在途⌉，不足 0 记 0；基数 0（低量）直接 0。"""
    if base <= 0:
        return 0
    need = base * cover_weeks - (onhand or 0) - (incoming or 0)
    return max(0, math.ceil(need - 1e-9))


def low_profile(lines, since=None):
    """近期（date ≥ since）净价 > 0 的行里最低净价那家 + 该家窗口内最近一次净价。

    按 2 位小数比价；最低价由多家并列 → 供应商留空（不替人选，回 A3000 比价）。
    单价不用最低那笔而用该家最近一次（2026-11-12 用户定：避开一次性低价）。
    """
    win = [l for l in lines if (l["net"] or 0) > 0 and (since is None or str(l["date"])[:10] >= since)]
    if not win:
        return {COL_LOW_VENDOR: "", COL_LOW_NET: None, COL_LOW_LAST: None}
    low = min(round(l["net"], 2) for l in win)
    who = {l["vendor"] for l in win if round(l["net"], 2) == low}
    if len(who) != 1:
        return {COL_LOW_VENDOR: "", COL_LOW_NET: low, COL_LOW_LAST: None}
    v = next(iter(who))
    last = max((l for l in win if l["vendor"] == v), key=lambda l: (l["date"], l["po"]))
    return {COL_LOW_VENDOR: v, COL_LOW_NET: low, COL_LOW_LAST: round(last["net"], 2)}


def vendor_profile(lines, low_since=None):
    """单个基础品的采购行 [dict(date, vendor, net, po)] → 供应商画像 dict。

    最近一次：按 (日期, PO) 取最晚；最低价：净价 > 0 的行里最低（0 元行是赠品 / 返利）。
    近期最低价（low_since 起）见 low_profile。
    """
    if not lines:
        return {"供应商数": 0}
    vendors = {l["vendor"] for l in lines}
    last = max(lines, key=lambda l: (l["date"], l["po"]))
    priced = [l for l in lines if (l["net"] or 0) > 0]
    low = min(priced, key=lambda l: (l["net"], l["date"])) if priced else None
    return {
        "供应商数": len(vendors),
        COL_QUOTE_VENDOR: next(iter(vendors)) if len(vendors) == 1 else "",
        "最近一次供应商": last["vendor"],
        COL_LAST_NET: round(last["net"], 2) if last["net"] is not None else None,
        "最近一次采购日": str(last["date"])[:10],
        "最低价供应商": low["vendor"] if low else "",
        "最低净价": round(low["net"], 2) if low else None,
        "采购单数": len({l["po"] for l in lines}),
        **low_profile(lines, low_since),
    }


def build_rows(bases, sales, recent, n_days, po_lines, onhand, incoming, products,
               min_weeks=6, bulk_qty=100, bulk_share=0.5, suspect_qty=5000,
               cover_bulk=3, cover_regular=1, fines=None, unrolled=(), last7=None,
               avg_factor=1.2, drop_ratio=0.5, min_recent_avg=3, low_since=None):
    """合并成汇总 DataFrame + 待核行。只出有销周数 ≥ min_weeks 的基础品。

    sales: {基础品: {周: 件}}；recent: {基础品: 近 4 周件数}；last7: {基础品: 近 7 天件数}；po_lines: {基础品: [采购行 dict]}；
    products: {id: dict(code, name)}；fines: {基础品: EUR} 或 None；
    unrolled: 有后缀却没 BoM 的产品 id（销量没折回基础品，备注里点名）。
    """
    rows, checks = [], []
    for b in bases:
        st = sales_stats(sales.get(b, {}), recent.get(b, 0.0), n_days)
        l7 = (last7 or {}).get(b, 0.0)
        base, basis = demand_base(st["近4周周均"], l7, st["周均销量"], avg_factor, drop_ratio, min_recent_avg)
        if st["有销周数"] < min_weeks:
            continue
        lines = po_lines.get(b, [])
        po_qty = defaultdict(float)
        for l in lines:
            po_qty[l["po"]] += l["qty"]
        bulk, share, suspect = classify(po_qty, bulk_qty, bulk_share, suspect_qty)
        cover = cover_bulk if bulk else cover_regular
        oh, inc = onhand.get(b, 0.0), incoming.get(b, 0.0)
        p = products.get(b, {})
        notes = ["有后缀无 BoM，未折回基础品"] if b in unrolled else []
        for po in suspect:
            notes.append(f"疑录错 {po} {po_qty[po]:,.0f} 件")
            checks.append({"类型": "疑录错大单", "SKU": p.get("code", ""), "产品": p.get("name", ""),
                           "说明": f"{po} 单张 {po_qty[po]:,.0f} 件，超过 {suspect_qty}，未计入分档"})
        if not lines:
            notes.append("12 个月无采购记录")
            checks.append({"类型": "常销无采购记录", "SKU": p.get("code", ""), "产品": p.get("name", ""),
                           "说明": "近 12 个月没有有效采购行，供应商与价格列为空"})
        row = {
            "档位": f"大批量·{cover:g}周" if bulk else f"常规·{cover:g}周",
            "SKU": p.get("code", ""),
            "产品": p.get("name", ""),
            COL_SUGGEST: suggest_qty(base, cover, oh, inc),
            "基数口径": basis,
            "在手": oh,
            "在途": inc,
            "备货基数/周": round(base, 1),
            "备货周数": cover,
            "可撑周数": round(oh / base, 1) if base else None,
            "近7天": round(l7),
            "近4周周均": round(st["近4周周均"], 1),
            "12周周均": round(st["周均销量"], 1),
            "峰值周销量": round(st["峰值周销量"], 1),
            "趋势": (f"{st['趋势']:+.0%}" if st["趋势"] is not None else "新起量"),
            "有销周数": st["有销周数"],
            "12周销量": round(st["销量"]),
        }
        prof = vendor_profile(lines, low_since)
        row.update({k: prof.get(k) for k in (COL_QUOTE_VENDOR, "最近一次供应商", COL_LAST_NET,
                                             "最近一次采购日", "最低价供应商", "最低净价",
                                             COL_LOW_VENDOR, COL_LOW_NET, COL_LOW_LAST,
                                             "供应商数", "采购单数")})
        row["大单量占比"] = f"{share:.0%}" if share is not None else ""
        if fines is not None:
            row["缺货罚款EUR"] = round(fines.get(b, 0.0), 2)
        row["备注"] = "；".join(notes)
        row["产品ID"] = b
        rows.append(row)
    df = pd.DataFrame(rows)
    if not df.empty:
        df["_need"] = df[COL_SUGGEST] > 0
        df = (df.sort_values(["_need", "可撑周数", "12周销量"], ascending=[False, True, False],
                             na_position="last")
              .drop(columns="_need").reset_index(drop=True))
        for c in (COL_QUOTE_VENDOR, "最近一次供应商", "最近一次采购日", "最低价供应商", COL_LOW_VENDOR):
            df[c] = df[c].fillna("")
    return df, checks


def quote_rows(df, vendor):
    """该家询价单要补的行：建议数量 > 0 且（独家 = 该家 或 近期最低价 = 该家）。

    返回 (DataFrame, 单价 Series, 来源 Series)：独家品用「最近一次净价」（单一供应商即该家上次价），
    放宽进来的品用「近期该家最近净价」（ISSUES [备货提醒] D）。
    """
    if df.empty:
        return df, pd.Series(dtype=float), pd.Series(dtype=str)
    low = df[COL_LOW_VENDOR] if COL_LOW_VENDOR in df else pd.Series("", index=df.index)
    sole = df[COL_QUOTE_VENDOR] == vendor
    take = (df[COL_SUGGEST] > 0) & (sole | (low == vendor))
    sub = df[take]
    lowlast = sub[COL_LOW_LAST] if COL_LOW_LAST in sub else pd.Series(None, index=sub.index, dtype=float)
    price = sub[COL_LAST_NET].where(sole[take], lowlast)
    src = sole[take].map({True: "独家", False: "近期最低价"})
    return sub, price, src


def rfq_vendors(df):
    """能出询价单的供应商：独家品的供应商 ∪ 本周需补品的近期最低价供应商。"""
    if df.empty:
        return []
    vs = set(df[COL_QUOTE_VENDOR])
    if COL_LOW_VENDOR in df:
        vs |= set(df.loc[df[COL_SUGGEST] > 0, COL_LOW_VENDOR])
    return sorted(v for v in vs if isinstance(v, str) and v)


def quote_vendors(df):
    """能出询价单的供应商 → 每家一行：独家品数 / 本周需补（独家 + 近期最低价）/ 建议数量合计。"""
    cols = [COL_QUOTE_VENDOR, "独家常销品数", "本周需补品数", "其中近期最低价", "建议数量合计", "需补 SKU"]
    out = []
    for v in rfq_vendors(df):
        need, _, src = quote_rows(df, v)
        out.append({COL_QUOTE_VENDOR: v, "独家常销品数": int((df[COL_QUOTE_VENDOR] == v).sum()),
                    "本周需补品数": len(need), "其中近期最低价": int((src == "近期最低价").sum()),
                    "建议数量合计": int(need[COL_SUGGEST].sum()),
                    "需补 SKU": "、".join(need["SKU"])})
    if not out:
        return pd.DataFrame(columns=cols)
    return pd.DataFrame(out, columns=cols).sort_values(["本周需补品数", "独家常销品数"], ascending=False)


# --------------------------------------------------------------------------
# Odoo 拉数
# --------------------------------------------------------------------------
def prefix_domain(prefixes=PREFIXES):
    """订单号前缀（OR）。`=like` 才是前缀匹配；`like` 会自动包成 %…% 变子串。"""
    terms = [("name", "=like", f"{p}%") for p in prefixes]
    return ["|"] * (len(terms) - 1) + terms


def pull_kit_map(od):
    """phantom BoM → {套件 product id: [(组件 product id, 系数)]}，系数 = 组件数量 ÷ BoM 产出数量。

    BoM 挂在模板上（product_id 为空）时，模板下所有变体都按它折；挂在具体变体上则只管那个变体。
    """
    boms = od.search_read_all("mrp.bom", [("type", "=", "phantom")],
                              ["product_tmpl_id", "product_id", "product_qty"])
    if not boms:
        return {}
    lines = od.search_read_all("mrp.bom.line", [("bom_id", "in", [b["id"] for b in boms])],
                               ["bom_id", "product_id", "product_qty"])
    comps = defaultdict(list)
    for l in lines:
        comps[m2o_id(l["bom_id"])].append((m2o_id(l["product_id"]), l["product_qty"] or 0))
    tmpls = sorted({m2o_id(b["product_tmpl_id"]) for b in boms if not b.get("product_id")})
    by_tmpl = defaultdict(list)
    for p in (od.search_read_all("product.product", [("product_tmpl_id", "in", tmpls)],
                                 ["product_tmpl_id"]) if tmpls else []):
        by_tmpl[m2o_id(p["product_tmpl_id"])].append(p["id"])
    out = {}
    for b in sorted(boms, key=lambda b: b["id"]):        # 同一产品多条 BoM：取 id 最小那条（Odoo 默认顺序）
        targets = [m2o_id(b["product_id"])] if b.get("product_id") else by_tmpl[m2o_id(b["product_tmpl_id"])]
        per = b["product_qty"] or 1
        for t in targets:
            if t not in out and comps[b["id"]]:
                out[t] = [(c, q / per) for c, q in comps[b["id"]]]
    say(f"  组合装: {len(boms)} 条 phantom BoM → {len(out)} 个套件产品按 BoM 折回基础品")
    return out


def pull_retail_sales(od, since, until, recent_since, last_since):
    """VO/GW 零售销量 → ([(pid, 周, 件)], {pid: 近 4 周件数}, {pid: 近 7 天件数})。"""
    dom = [("date", ">=", since), ("date", "<=", until), ("state", "in", SOLD_STATES)] + prefix_domain()
    rows = od.read_group_all("sale.report", dom, ["product_uom_qty:sum"], ["product_id", "date:week"])
    weekly = [(m2o_id(r["product_id"]), r["date:week"], r.get("product_uom_qty") or 0)
              for r in rows if m2o_id(r.get("product_id")) is not None]
    def by_product(start):
        dom = [("date", ">=", start), ("date", "<=", until), ("state", "in", SOLD_STATES)] + prefix_domain()
        return {m2o_id(r["product_id"]): r.get("product_uom_qty") or 0
                for r in od.read_group_all("sale.report", dom, ["product_uom_qty:sum"], ["product_id"])
                if m2o_id(r.get("product_id")) is not None}
    recent, last = by_product(recent_since), by_product(last_since)
    say(f"  VO/GW 零售销量: {len({w[0] for w in weekly})} 个产品有成交，"
        f"合计 {sum(w[2] for w in weekly):,.0f} 件（折算前）")
    return weekly, recent, last


def pull_purchases(od, since, until, exclude_ids, kit_map):
    """近 N 个月已确认采购行 → ({基础品: [dict(date, vendor, qty, net, po)]}, 剔除行数)。

    数量按产品计量单位（product_uom_qty，没有则 product_qty）；净价按 price_check 口径去税，
    再折到产品计量单位（× product_qty ÷ product_uom_qty），套件采购再 ÷ BoM 系数折到基础品单件。
    """
    from procure.price_check import ensure_taxes, unit_net
    have = od.fields_of("purchase.order.line")
    uq = "product_uom_qty" if "product_uom_qty" in have else "product_qty"
    fields = ["product_id", "partner_id", "product_qty", uq, "price_unit", "taxes_id", "date_order", "order_id"]
    if "discount" in have:
        fields.append("discount")
    dom = [("state", "in", PO_STATES), ("date_order", ">=", since), ("date_order", "<=", f"{until} 23:59:59")]
    if exclude_ids:
        dom.append(("partner_id", "not in", exclude_ids))
    rows = od.search_read_all("purchase.order.line", dom, fields, label="采购行")
    taxes = {}
    ensure_taxes(od, taxes, {t for r in rows for t in r["taxes_id"]})
    out, n_noise = defaultdict(list), 0
    for r in rows:
        pid, vendor = m2o_id(r.get("product_id")), m2o_name(r.get("partner_id"))
        if pid is None or not vendor:
            continue
        if is_noise_vendor(vendor):
            n_noise += 1
            continue
        qty_line, qty_u = r.get("product_qty") or 0, r.get(uq) or 0
        net = unit_net(r["price_unit"], r.get("discount", 0), r["taxes_id"], taxes)
        if qty_u and qty_line and qty_u != qty_line:
            net = net * qty_line / qty_u
        for base, k in explode(pid, kit_map):
            out[base].append({"date": r["date_order"], "vendor": vendor.strip(), "qty": qty_u * k,
                              "net": net / k if k else net, "po": m2o_name(r.get("order_id")) or ""})
    say(f"  采购行: {len(rows)} 行（剔除噪声供应商 {n_noise} 行）→ {len(out)} 个基础品有采购记录")
    return out


def pull_codes(od, pids):
    out = {}
    ids = sorted(pids)
    for i in range(0, len(ids), 1000):
        for r in od.execute("product.product", "read", [ids[i:i + 1000], ["default_code", "name"]]):
            out[r["id"]] = {"code": (r.get("default_code") or "").strip(), "name": (r.get("name") or "").strip()}
    return out


def fines_by_base(od, viol, kit_map):
    """违规表（index=vo_sku）→ {基础品: 缺货罚款EUR}。vo_sku 按 Internal Reference 对产品，再按 BoM 折。"""
    codes = [c for c in viol.index if c]
    prods = od.search_read_all("product.product", [("default_code", "in", codes)], ["default_code"]) if codes else []
    by_code = {p["default_code"]: p["id"] for p in prods}
    out, miss = defaultdict(float), 0
    for code, r in viol.iterrows():
        pid = by_code.get(code)
        if pid is None:
            miss += 1
            continue
        for base, _ in explode(pid, kit_map):
            out[base] += r.get("缺货罚款EUR", 0) or 0
    if miss:
        say(f"  ⚠ 违规表里 {miss} 个 vo_sku 在 Odoo 找不到产品，罚款未计入")
    return out


# --------------------------------------------------------------------------
# 输出
# --------------------------------------------------------------------------
def write_book(path, sheets, left_cols):
    wb = Workbook()
    wb.remove(wb.active)
    for name, df in sheets:
        ws = wb.create_sheet(name)
        write_df(ws, df)
        style_sheet(ws, len(df.columns), left_cols=left_cols)
        ws.freeze_panes = "D2" if name == MAIN_SHEET else "A2"
    path = unique_path(path)
    wb.save(path)
    return path


def run(until, weeks=12, min_weeks=6, po_months=12, bulk_qty=100, bulk_share=0.5,
        suspect_qty=5000, cover_bulk=3, cover_regular=1, transit_days=60,
        violations=None, exclude_vendors=None, outdir=None, timeout=180,
        avg_factor=1.2, drop_ratio=0.5, min_recent_avg=3, low_months=3):
    since = (date.fromisoformat(until) - timedelta(weeks=weeks)).isoformat()
    n_days = (date.fromisoformat(until) - date.fromisoformat(since)).days + 1
    recent_since = (date.fromisoformat(until) - timedelta(days=RECENT_DAYS - 1)).isoformat()
    last_since = (date.fromisoformat(until) - timedelta(days=LAST_DAYS - 1)).isoformat()
    po_since = (date.fromisoformat(until) - timedelta(days=round(30.4 * po_months))).isoformat()
    transit_since = (date.fromisoformat(until) - timedelta(days=transit_days)).isoformat()
    low_since = (date.fromisoformat(until) - timedelta(days=round(30.4 * low_months))).isoformat()
    od = Odoo.connect(timeout=timeout)
    say(f"销量窗口 {since} ~ {until}（{weeks} 周，VO/GW 前缀）  采购窗口 {po_since} 起  "
        f"在途只认 {transit_since} 起创建的入库  "
        f"近期最低价窗口 {low_since} 起")
    kit_map = pull_kit_map(od)
    weekly, recent_raw, last_raw = pull_retail_sales(od, since, until, recent_since, last_since)
    sales = roll_sales(weekly, kit_map)

    def fold(raw):
        out = defaultdict(float)
        for pid, q in raw.items():
            for base, k in explode(pid, kit_map):
                out[base] += q * k
        return out
    recent, last7 = fold(recent_raw), fold(last_raw)
    vend_ids = resolve_excluded_vendors(od, exclude_vendors)
    po_lines = pull_purchases(od, po_since, until, vend_ids, kit_map)
    n_weeks = max(len({w[1] for w in weekly}), 1)    # 窗口跨几个自然周（12 周 → 通常 13），只用于报告有销周数
    bases = set(sales)
    onhand = pull_qty_available(od, bases)
    incoming, _ = pull_in_transit(od, vend_ids, since=transit_since)
    sold_pids = {w[0] for w in weekly}
    products = pull_codes(od, bases | sold_pids)
    fines = None
    if violations:
        fines = fines_by_base(od, read_violations(violations, since), kit_map)

    # 有后缀却没 BoM：销量没法折回基础品（不按后缀猜），点名交人
    unrolled = sorted(pid for pid in sold_pids if pid not in kit_map
                      and products.get(pid, {}).get("code")
                      and _po_base_sku(products[pid]["code"]) != products[pid]["code"])
    df, checks = build_rows(bases, sales, recent, n_days, po_lines, onhand, incoming, products,
                            min_weeks, bulk_qty, bulk_share, suspect_qty, cover_bulk, cover_regular, fines,
                            set(unrolled), last7, avg_factor, drop_ratio, min_recent_avg, low_since)
    for pid in unrolled:
        q = sum(sales.get(pid, {}).values())
        checks.append({"类型": "有后缀无 BoM", "SKU": products[pid]["code"], "产品": products[pid]["name"],
                       "说明": f"12 周零售 {q:,.0f} 件，Odoo 里没有组合装 BoM，销量记在它自己名下、未折回基础品"})
    chk = pd.DataFrame(checks, columns=["类型", "SKU", "产品", "说明"])
    qv = quote_vendors(df)

    outdir = outdir or os.path.join("output", f"{date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)
    path = write_book(os.path.join(outdir, f"备货汇总-{until.replace('-', '')}.xlsx"),
                      [(MAIN_SHEET, df), (QUOTE_SHEET, qv), (CHECK_SHEET, chk)],
                      left_cols={"SKU", "产品", "备注", "说明", "需补 SKU", COL_QUOTE_VENDOR,
                                 "最近一次供应商", "最低价供应商", COL_LOW_VENDOR})
    total = sum(sum(v.values()) for v in sales.values())
    need = df[df[COL_SUGGEST] > 0] if not df.empty else df
    say(f"\n✓ {path}")
    say(f"  常销品（有销周数 ≥ {min_weeks}/{n_weeks}）{len(df)} 个，占零售件数 "
        f"{(df['12周销量'].sum() / total * 100 if total else 0):.0f}%；本周需补 {len(need)} 个")
    if not df.empty:
        say("  档位: " + "  ".join(f"{k} {v}" for k, v in df["档位"].value_counts().sort_index().items()))
        say("  基数口径: " + "  ".join(f"{k} {v}" for k, v in df["基数口径"].value_counts().items())
            + f"；建议数量合计 {int(df[COL_SUGGEST].sum()):,} 件")
    say(f"  询价单供应商 {len(qv)} 家（本周需补 {int(qv['本周需补品数'].sum())} 品，"
        f"其中近期最低价放宽进来 {int(qv['其中近期最低价'].sum())} 品）；待核 {len(chk)} 条"
        + (f"（{'、'.join(f'{k} {v}' for k, v in chk['类型'].value_counts().items())}）" if len(chk) else ""))
    say(f"  RPC 调用 {od.call_count} 次")
    return path, df, qv, chk


def main():
    ap = argparse.ArgumentParser(description="备货汇总：VO/GW 零售销量 → 建议数量 + 供应商 + 上次采购价（只读）")
    ap.add_argument("--until", help="截止日（YYYY-MM-DD，默认今天）")
    ap.add_argument("--weeks", type=int, default=12, help="销量窗口周数（默认 12）")
    ap.add_argument("--min-weeks", type=int, default=6,
                    help="有销周数达到几周才进汇总（默认 6；12 周窗口通常跨 13 个自然周）")
    ap.add_argument("--po-months", type=int, default=12, help="采购记录窗口月数（默认 12）")
    ap.add_argument("--low-months", type=int, default=3,
                    help="近期最低价窗口月数（默认 3）：需补品近期最低净价那家也进它的询价单")
    ap.add_argument("--bulk-qty", type=float, default=100, help="单张 PO 超过几件算大单（默认 100）")
    ap.add_argument("--bulk-share", type=float, default=0.5,
                    help="大单量占采购总量达到多少算大批量（默认 0.5）")
    ap.add_argument("--suspect-qty", type=float, default=5000,
                    help="单张 PO 超过几件视为疑录错、不参与分档（默认 5000）")
    ap.add_argument("--cover-bulk", type=float, default=3, help="大批量品备几个峰值周（默认 3）")
    ap.add_argument("--cover-regular", type=float, default=1, help="其余常销品备几个峰值周（默认 1）")
    ap.add_argument("--avg-factor", type=float, default=1.2,
                    help="备货基数 = 近 4 周周均 × 该系数（默认 1.2）")
    ap.add_argument("--drop-ratio", type=float, default=0.5,
                    help="近 7 天 < 12 周周均 × 该比例算近期下降，基数改用近 4 周周均（默认 0.5）")
    ap.add_argument("--min-recent-avg", type=float, default=3,
                    help="近 4 周周均低于几件算低量品，有单再补、建议 0（默认 3）")
    ap.add_argument("--transit-days", type=int, default=60,
                    help="在途只认近几天内创建的入库 move（默认 60；更老的多是残单）")
    ap.add_argument("--violations", help="天猫「供应商违规数据」导出（选填），多出一列 缺货罚款EUR")
    ap.add_argument("--exclude-vendor", default=None,
                    help="排除这些供应商的采购单（名字子串，逗号分隔）。默认读 config.py 的 ODOO_EXCLUDE_VENDORS")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    ap.add_argument("--timeout", type=int, default=180)
    a = ap.parse_args()
    if a.exclude_vendor is None:
        from common import localconf
        exclude_vendors = list(localconf.get("ODOO_EXCLUDE_VENDORS", []) or [])
    else:
        exclude_vendors = [x.strip() for x in a.exclude_vendor.split(",") if x.strip()]
    try:
        run(a.until or date.today().isoformat(), a.weeks, a.min_weeks, a.po_months, a.bulk_qty,
            a.bulk_share, a.suspect_qty, a.cover_bulk, a.cover_regular, a.transit_days,
            a.violations, exclude_vendors, a.outdir, a.timeout,
            a.avg_factor, a.drop_ratio, a.min_recent_avg, a.low_months)
    except OdooError as e:
        raise SystemExit(f"✗ {e}")


if __name__ == "__main__":
    main()
