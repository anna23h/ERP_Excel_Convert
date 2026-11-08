# -*- coding: utf-8 -*-
"""补货预判清单的「动态安全库存」+ 套件并入基础品：Odoo 只读拉数（ISSUES [安全库存] A）。

为什么：ERP 产品上的 `Safety Stock` 是静态值，除了用户本人没人看、也没人维护，
2026-11-08 实测非套件 101 个里 37% 与近期销量明显脱节，另有 59 个配在套件上
（套件没有自己的实物，等于同一批货被看两次）。用户定（2026-11-08）：
  - 不再读 ERP 字段、也不回写；清单直接显示按销量算的动态值（方案 1）。
  - 安全库存 = **12 周 VO/GW 零售周均 × 2**（覆盖两周）；**周均 < 10 件不设**（留空）。
  - 套件 / 店铺后缀品**并进基础品那一行**一起算：今日需求 = 基础品 + 套件 × BoM 数量，
    在手取基础品的 qty_available。
  - 运营的安全库存表不定期，不参与。

口径与备货汇总（odoo_api/restock.py）同一套：订单号 `=like 'VO%' / 'GW%'`、state sale/done、
phantom BoM 折算（直接复用 restock 的 pull_kit_map / roll_sales），周均按窗口天数折周。
窗口截止**昨天**（今天的单还没收完）。

拉不到（没配凭据 / 连不上）由调用方降级：套件不并、两列不出，清单照出。
"""
import os
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

WEEKS = 12
COVER_WEEKS = 2
MIN_AVG = 10          # 周均低于此不设安全库存
COL_AVG = f"周均({WEEKS}周)"
COL_SAFETY = f"安全库存({COVER_WEEKS}周)"
COL_KITS = "含套件"


def safety_of(avg, cover=COVER_WEEKS, min_avg=MIN_AVG):
    """周均 → 安全库存（整数）；周均 < min_avg 返回 None（不设）。"""
    if avg is None or avg < min_avg:
        return None
    return int(round(avg * cover))


def kit_code_map(kit_map, code_of, pid_of_sku):
    """{套件 pid: [(组件 pid, 系数)]} → {套件 SKU: (基础品 SKU, 系数)}，只管清单里出现的 SKU。

    多组件套件无法并进单一一行，不并、单列返回让调用方点名（2026-11-08 实测 129 个全是单组件）。
    → (映射, 多组件套件 SKU 列表)
    """
    out, multi = {}, []
    for sku, pid in pid_of_sku.items():
        comps = kit_map.get(pid)
        if not comps:
            continue
        if len(comps) != 1:
            multi.append(sku)
            continue
        base, k = comps[0]
        code = code_of.get(base, "")
        if code:
            out[sku] = (code, k)
    return out, sorted(multi)


def pull_safety(skus, until=None, weeks=WEEKS, timeout=60):
    """→ dict(kits={套件 SKU: (基础品 SKU, 系数)}, avg={SKU: 周均}, onhand={基础品 SKU: 在手}, info=说明)。

    avg 只给并入后清单里会出现的 SKU（非套件 SKU + 套件的基础品）。拉不到时抛异常，由调用方降级。
    """
    from odoo_api.odoo_client import Odoo, m2o_id
    from odoo_api.restock import SOLD_STATES, prefix_domain, pull_kit_map, pull_codes, roll_sales
    from odoo_api.stock_report import pull_qty_available
    import odoo_api.restock as rs

    od = Odoo.connect(timeout=timeout, verbose=False)
    until = until or (date.today() - timedelta(days=1))
    since = until - timedelta(days=weeks * 7 - 1)
    n_days = (until - since).days + 1

    skus = sorted({str(s).strip() for s in skus if str(s).strip()})
    prods = od.search_read_all("product.product", [("default_code", "in", skus)], ["default_code"])
    pid_of, dup = {}, 0
    for p in prods:                      # 按 id 升序，同码取最早建档的（同 pending_moves）
        if p["default_code"] in pid_of:
            dup += 1
            continue
        pid_of[p["default_code"]] = p["id"]

    _say, rs.say = rs.say, (lambda *a, **k: None)   # restock 的进度打印进 VOTool 日志是噪音
    try:
        kit_map = pull_kit_map(od)
        base_pids = {c for pid in pid_of.values() for c, _ in kit_map.get(pid, [])}
        code_of = {pid: v["code"] for pid, v in pull_codes(od, base_pids).items()} if base_pids else {}
    finally:
        rs.say = _say
    kits, multi = kit_code_map(kit_map, code_of, pid_of)

    dom = [("date", ">=", since.isoformat()), ("date", "<=", f"{until.isoformat()} 23:59:59"),
           ("state", "in", SOLD_STATES)] + prefix_domain()
    rows = od.read_group_all("sale.report", dom, ["product_uom_qty:sum"], ["product_id", "date:week"])
    weekly = [(m2o_id(r["product_id"]), r["date:week"], r.get("product_uom_qty") or 0)
              for r in rows if m2o_id(r.get("product_id")) is not None]
    by_base = roll_sales(weekly, kit_map)
    tot = {pid: sum(w.values()) for pid, w in by_base.items()}

    # 并入后清单里的 SKU → pid：非套件用自身，基础品用 BoM 组件 id
    final = {sku: pid for sku, pid in pid_of.items() if sku not in kits}
    base_pid_of = {code_of[c]: c for pid in pid_of.values() for c, _ in kit_map.get(pid, [])
                   if c in code_of}
    for sku, (base, _) in kits.items():
        final.setdefault(base, base_pid_of[base])
    avg = {sku: tot.get(pid, 0) / n_days * 7 for sku, pid in final.items()}

    need = {base_pid_of[b] for b, _ in kits.values()}
    oh = pull_qty_available(od, need, label="基础品在手") if need else {}
    onhand = {code_of[pid]: q for pid, q in oh.items()}

    n_set = sum(1 for v in avg.values() if safety_of(v) is not None)
    info = (f"VO/GW 零售 {since:%m-%d}~{until:%m-%d}（{weeks} 周，套件按 BoM 折算）；"
            f"{len(pid_of)}/{len(skus)} 个 SKU 在 Odoo 找到，{len(kits)} 个套件并入基础品；"
            f"周均 ≥ {MIN_AVG} 设安全库存 {n_set} 个")
    if multi:
        info += f"；⚠ 多组件套件 {len(multi)} 个未并入: {', '.join(multi[:5])}"
    if dup:
        info += f"；⚠ {dup} 个 Internal Reference 在 Odoo 有重码，取最早建档的"
    return {"kits": kits, "avg": avg, "onhand": onhand, "info": info}
