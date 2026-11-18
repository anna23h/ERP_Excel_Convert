"""按 SO 号直接从 Odoo 读箱单数据（只读，走 odoo_api/odoo_client.py）。ISSUES [箱单] E。

产出与 packing_list.read_orders 同构的 orders 列表（kind="moves"），另给每行挂 line["pack"]
（产品箱规主数据，estimate.classify 的结果）。

数据链路（Odoo 14.0e，S04162 实测）:
  sale.order.picking_ids → 出库类 picking（排除 cancel）→ stock.move
  （product_uom_qty = 订购量，reserved_availability = 预留量）→ stock.move.line
  （lot_id；数量取 qty_done，没有就取 product_uom_qty）→ stock.production.lot.expiration_date
  （存的是 UTC，转 Europe/Berlin 再取日期，否则月底的效期会早一天）。
SO 还没有出库单（报价单）时退回 sale.order.line 的订购量，批次留空。
"""
import datetime as dt

from odoo_api.odoo_client import m2o_id

from . import estimate

PACK_FIELDS = ["carton_pieces", "carton_dimensions", "carton_weight", "pallet_pieces", "weight"]


def berlin_date(utc_str):
    """Odoo 的 'YYYY-MM-DD HH:MM:SS'（UTC）→ 柏林本地日期。"""
    if not utc_str:
        return None
    t = dt.datetime.strptime(utc_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=dt.timezone.utc)
    try:
        from zoneinfo import ZoneInfo
        return t.astimezone(ZoneInfo("Europe/Berlin")).date()
    except Exception:   # Windows 没装 tzdata 时：夏令时 = 3 月最后一个周日 ~ 10 月最后一个周日
        def last_sunday(month):
            d = dt.date(t.year, month, 31)
            return d - dt.timedelta(days=(d.weekday() + 1) % 7)
        summer = last_sunday(3) <= t.date() < last_sunday(10)
        return (t + dt.timedelta(hours=2 if summer else 1)).date()


def read_orders(odoo, so_names, new_line, add_lot):
    """→ [{order, kind, lines}]，按给定 SO 号的顺序。
    new_line / add_lot 传 packing_list 的 _new_line / _add_lot（避免循环导入，行结构只定义一处）。"""
    sos = odoo.search_read_all("sale.order", [("name", "in", list(so_names))],
                               ["name", "state", "picking_ids", "order_line"])
    by_name = {s["name"]: s for s in sos}
    missing = [n for n in so_names if n not in by_name]
    if missing:
        raise ValueError(f"Odoo 里找不到这些 SO：{', '.join(missing)}")

    orders, all_pids = [], set()
    for name in so_names:
        so = by_name[name]
        if so["state"] == "cancel":
            raise ValueError(f"{name} 已取消")
        order = {"order": name, "kind": "moves", "lines": {}, "pickings": []}
        picks = odoo.search_read_all(
            "stock.picking",
            [("id", "in", so["picking_ids"]), ("state", "!=", "cancel"),
             ("picking_type_code", "=", "outgoing")],
            ["name", "state"]) if so["picking_ids"] else []
        moves = odoo.search_read_all(
            "stock.move",
            [("picking_id", "in", [p["id"] for p in picks]), ("state", "!=", "cancel")],
            ["product_id", "product_uom_qty", "reserved_availability", "quantity_done", "state"]
        ) if picks else []
        order["pickings"] = [p["name"] for p in picks]

        if moves:
            mlines = odoo.search_read_all(
                "stock.move.line", [("move_id", "in", [m["id"] for m in moves])],
                ["move_id", "lot_id", "product_uom_qty", "qty_done"])
            lot_ids = list({m2o_id(l["lot_id"]) for l in mlines if l["lot_id"]})
            lots = {l["id"]: l for l in odoo.search_read_all(
                "stock.production.lot", [("id", "in", lot_ids)],
                ["name", "expiration_date"])} if lot_ids else {}
            rows = [(m2o_id(m["product_id"]), m["product_uom_qty"],
                     # 已完成的 move 预留归零，按已完成数算
                     m["quantity_done"] if m["state"] == "done" else m["reserved_availability"],
                     [(lots.get(m2o_id(ml["lot_id"])), ml["qty_done"] or ml["product_uom_qty"])
                      for ml in mlines if m2o_id(ml["move_id"]) == m["id"]])
                    for m in moves]
        else:
            order["kind"] = "so"
            sol = odoo.search_read_all(
                "sale.order.line", [("order_id", "=", so["id"]), ("display_type", "=", False)],
                ["product_id", "product_uom_qty", "sequence"], order="sequence,id")
            rows = [(m2o_id(l["product_id"]), l["product_uom_qty"], None, []) for l in sol
                    if l["product_id"]]
        order["_rows"] = rows
        all_pids.update(r[0] for r in rows)
        orders.append(order)

    prods = {p["id"]: p for p in odoo.search_read_all(
        "product.product", [("id", "in", list(all_pids))],
        ["default_code", "name", "barcode", "hs_code"] + PACK_FIELDS)}

    for order in orders:
        for pid, qty, reserved, lot_rows in order.pop("_rows"):
            p = prods[pid]
            sku = p["default_code"] or f"(无编码 id={pid})"
            line = order["lines"].get(sku)
            if not line:
                line = order["lines"][sku] = new_line(
                    sku, p["name"], p["barcode"] or None, p["hs_code"] or None, 0)
                line["pack"] = estimate.classify({k: p.get(k) for k in PACK_FIELDS})
                if reserved is not None:
                    line["reserved"] = 0
            line["qty"] += qty
            if reserved is not None:
                line["reserved"] += reserved
            for lot, lot_qty in lot_rows:
                add_lot(line, lot["name"] if lot else None,
                         berlin_date(lot["expiration_date"]) if lot else None, lot_qty)
    return [o for o in orders if o["lines"]]

