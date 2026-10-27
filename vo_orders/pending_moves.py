# -*- coding: utf-8 -*-
"""补货预判清单的「待出库 / 在途」：Odoo 只读拉未完成的出入库 move（ISSUES 补货预判 A）。

为什么要：仓库人手不足，寄出的货不能当天做出库、到的货不能当天入库，
Odoo 的 On Hand 永远滞后于实物。但开工时今天要发的单（拉单后不再变）和
今天要到的货（前晚采购完就定了）都已知，所以
    预计可用 = 在手 − 待出库 + 在途
比光看在手更接近实物。

口径（2026-10-27 用户拍板）：
  - 待出库 = 近 N 天（默认 7）创建、还没完成的出库 move，**剔除今日批次**——
    今日批次本身也是未完成出库，已算在「今日需求」里，不剔会重复计算。
    picking 的 origin = SO 名 = ERP 导出的 Order Reference，按它剔。
  - 在途 = 近 N 天创建、还没完成的入库 move，排除技术性 PO（config 的 ODOO_EXCLUDE_VENDORS）。
  - 窗口是为了压掉老的死单/残单：实测未完成出库最老到 2022 年，未完成入库 88 张里 54 张超过 8 天。
  - 组合装（phantom BoM）的 move 落在组件上，按 BoM 用量折回套件件数
    （569 个套件里 568 个是单组件；多组件取净消耗最大的那个组件）。

拉不到（没配凭据 / 连不上）由调用方降级：清单照出，只是没有这四列。
"""
import os, sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from common import localconf  # noqa: E402

WINDOW_DAYS = 7


def to_sku_pending(skus, prod_map, kit_lines, inc, out):
    """纯计算，便于单测。
    prod_map: {sku: pid}；kit_lines: {kit_pid: [(组件 pid, 每套用量)]}；
    inc/out: {pid: 数量}（组件口径）。→ {sku: (待出库, 在途)}，套件按套折算。"""
    res = {}
    for sku in skus:
        pid = prod_map.get(sku)
        if pid is None:
            continue
        o, i = out.get(pid, 0.0), inc.get(pid, 0.0)   # 套件自己理论上没有 move，保险起见也加上
        lines = kit_lines.get(pid)
        if lines:
            conv = [(out.get(c, 0.0) / q, inc.get(c, 0.0) / q) for c, q in lines if q]
            if conv:
                ko, ki = max(conv, key=lambda t: t[0] - t[1])
                o, i = o + ko, i + ki
        res[sku] = (o, i)
    return res


def pull_pending(skus, today_refs, days=WINDOW_DAYS, timeout=60):
    """→ ({sku: (待出库, 在途)}, 说明文字)。拉不到时抛异常，由调用方降级。"""
    from odoo_api.odoo_client import Odoo, m2o_id
    from odoo_api.stock_report import resolve_excluded_vendors, pull_in_transit

    od = Odoo.connect(timeout=timeout, verbose=False)
    since = (date.today() - timedelta(days=days)).isoformat()
    vend = resolve_excluded_vendors(od, list(localconf.get("ODOO_EXCLUDE_VENDORS", []) or []))
    inc, out = pull_in_transit(od, vend, since=since, exclude_origins=sorted(set(today_refs)))

    skus = sorted({str(s) for s in skus if str(s).strip()})
    prods = od.search_read_all("product.product", [("default_code", "in", skus)],
                               ["default_code", "product_tmpl_id"])
    prod_map, tmpl_of, dup = {}, {}, 0
    for p in prods:                     # 按 id 升序，同码取最早建档的那个
        if p["default_code"] in prod_map:
            dup += 1
            continue
        prod_map[p["default_code"]] = p["id"]
        tmpl_of[m2o_id(p["product_tmpl_id"])] = p["id"]

    kit_lines = {}
    if tmpl_of:
        boms = od.search_read_all("mrp.bom", [("type", "=", "phantom"),
                                              ("product_tmpl_id", "in", list(tmpl_of))],
                                  ["product_tmpl_id", "product_id", "product_qty", "bom_line_ids"])
        line_ids = [l for b in boms for l in b["bom_line_ids"]]
        lines = {l["id"]: l for l in od.search_read_all(
            "mrp.bom.line", [("id", "in", line_ids)], ["product_id", "product_qty"])} if line_ids else {}
        for b in boms:
            kit = m2o_id(b.get("product_id")) or tmpl_of.get(m2o_id(b["product_tmpl_id"]))
            per = b.get("product_qty") or 1.0
            if kit is None or kit in kit_lines:
                continue
            kit_lines[kit] = [(m2o_id(lines[l]["product_id"]), lines[l]["product_qty"] / per)
                              for l in b["bom_line_ids"] if l in lines]

    res = to_sku_pending(skus, prod_map, kit_lines, inc, out)
    info = (f"近 {days} 天（{since} 起）未完成出库/入库，剔除今日 {len(set(today_refs))} 单；"
            f"{len(prod_map)}/{len(skus)} 个 SKU 在 Odoo 找到，其中套件 {len(kit_lines)} 个")
    if dup:
        info += f"；⚠ {dup} 个 Internal Reference 在 Odoo 有重码，取最早建档的"
    return res, info
