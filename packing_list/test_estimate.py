"""箱单 --so 预估：构造数据单测，不连 Odoo（ISSUES [箱单] E）。

    python3 -m pytest packing_list/test_estimate.py
"""
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl  # noqa: E402

from packing_list import estimate as E  # noqa: E402
from packing_list import odoo_source  # noqa: E402
from packing_list import packing_list as PL  # noqa: E402


# ---------------------------------------------------------------- 解析（样例取自 2026-11-18 全库扫描）
def test_parse_pieces():
    assert E.parse_pieces("88") == 88
    assert E.parse_pieces("1,000") == 1000
    assert E.parse_pieces("200盒/箱") == 200
    assert E.parse_pieces("30/carton") == 30
    assert E.parse_pieces("12pcs per carton，11箱/层") == 12
    for bad in ("0", "", None, False, "一箱六罐", "小12，大96", "200个   /  100个"):
        assert E.parse_pieces(bad) is None, bad


def test_parse_weight():
    assert E.parse_weight_kg("0.696KG") == 0.696
    assert E.parse_weight_kg("10kg") == 10
    assert E.parse_weight_kg("2.8kg") == 2.8
    assert E.parse_weight_kg("4400g") == 4.4
    assert E.parse_weight_kg("2.48") == 2.48        # 纯数字 ≤50 → kg
    assert E.parse_weight_kg("2664") == 2.664       # 纯数字 >50 → 克
    for bad in ("0KG", "", None, "8KG/  5KG", "0.669KG (2.126KG from Budni info)", "478.5KG"):
        assert E.parse_weight_kg(bad) is None, bad


def test_parse_dims():
    assert E.parse_dims_cm("39*39*11.5") == (39, 39, 11.5)
    assert E.parse_dims_cm("18.2x15.5x10.8") == (18.2, 15.5, 10.8)
    assert E.parse_dims_cm("35.9+21.7+26.2") == (35.9, 21.7, 26.2)
    assert E.parse_dims_cm("18,5x30x40") == (18.5, 30, 40)
    assert E.parse_dims_cm("36 + 27 + 29cm") == (36, 27, 29)
    assert E.parse_dims_cm("390\t130\t260mm") == (39, 13, 26)
    assert E.parse_dims_cm("271+258+264") == (27.1, 25.8, 26.4)   # 某边 >100 → mm
    for bad in ("", None, "18箱/层", "27x2736", "45*37*29  /  39*29*22"):
        assert E.parse_dims_cm(bad) is None, bad


def test_per_layer_and_classify():
    assert E.per_layer((39, 39, 11.5)) == 6
    assert E.per_layer((35.9, 21.7, 26.2)) == 10   # 同 S04077 预估箱单「每层 10 箱」
    a = E.classify({"carton_pieces": "100", "carton_dimensions": "39*39*11.5",
                    "carton_weight": "10kg", "weight": 0.1})
    assert a["tier"] == "A" and a["warn"] is None
    w = E.classify({"carton_pieces": "48", "carton_dimensions": "30*20*20",
                    "carton_weight": "3.984KG", "weight": 0.16})
    assert w["tier"] == "A" and "对不上" in w["warn"]
    assert E.classify({"carton_pieces": "20", "weight": 0.35})["tier"] == "B"
    assert E.classify({"carton_pieces": "20", "carton_dimensions": "39*39*11.5",
                       "carton_weight": "0KG"})["tier"] == "B"
    assert E.classify({"carton_pieces": "一箱六罐"})["tier"] == "C"


# ---------------------------------------------------------------- 排托
def _line(sku, qty, pack):
    line = PL._new_line(sku, sku, "1", None, qty)
    line["pack"] = E.classify(pack)
    return line


def _order(name, *lines):
    return {"order": name, "kind": "moves", "lines": {l["sku"]: l for l in lines}}


DROPS = {"carton_pieces": "100", "carton_dimensions": "39*39*11.5", "carton_weight": "10kg"}


def test_single_pallet():
    o = _order("S1", _line("A", 6000, DROPS))
    pallets = E.plan([o])
    assert len(pallets) == 1 and pallets[0]["height"] == 130   # 60 箱 / 6 每层 = 10 层
    assert [(s["first"], s["last"], s["pcs"], s["count"]) for s in o["lines"]["A"]["segments"]] \
        == [(1, 60, 100, 60)]


def test_height_limit_splits_pallet():
    # 145cm 可码 12 层 × 6 = 72 箱；100 箱 → 托1 72 箱 + 托2 28 箱
    o = _order("S1", _line("A", 10000, DROPS))
    pallets = E.plan([o])
    assert [p["height"] for p in pallets] == [153, 73]   # 5 层 × 11.5 + 15 = 72.5 → 向上取整
    assert all(p["height"] <= E.MAX_HEIGHT_CM for p in pallets)
    segs = o["lines"]["A"]["segments"]
    assert [(s["pallet"]["no"], s["first"], s["last"]) for s in segs] == [(1, 1, 72), (2, 73, 100)]


def test_tail_carton_and_mixed_pallet():
    b = {"carton_pieces": "48", "carton_dimensions": "35.9*21.7*26.2", "carton_weight": "4.176KG"}
    o = _order("S1", _line("A", 1050, DROPS), _line("B", 480, b))
    pallets = E.plan([o])
    # A: 11 箱（10 整 + 尾箱 50 件），2 层 = 23cm；B: 10 箱 1 层 = 26.2cm → 拼进同一托
    assert len(pallets) == 1 and pallets[0]["height"] == 65   # 64.2 向上取整
    a_segs = o["lines"]["A"]["segments"]
    assert [(s["first"], s["last"], s["pcs"], s["count"]) for s in a_segs] \
        == [(1, 10, 100, 10), (11, 11, 50, 1)]
    assert sum(s["pcs"] * s["count"] for s in a_segs) == 1050
    assert o["lines"]["B"]["segments"][0]["first"] == 12


def test_no_mixing_across_orders_and_tier_b_c():
    o1 = _order("S1", _line("A", 100, DROPS))
    o2 = _order("S2", _line("A", 100, DROPS),
                _line("B", 100, {"carton_pieces": "20", "pallet_pieces": "60", "weight": 0.35}),
                _line("C", 5, {}))
    pallets = E.plan([o1, o2])
    assert [p["no"] for p in pallets] == [1, 2, 3, 4]       # S2 的 A 不拼进 S1 的托
    b_pallets = [p for p in pallets if p["rough"]]
    assert [sum(s["count"] for _, s in p["segments"]) for p in b_pallets] == [3, 2]   # 60 件/托 = 3 箱
    assert o2["lines"]["C"]["segments"] == []


# ---------------------------------------------------------------- 端到端（假 Odoo）
class FakeOdoo:
    DATA = {
        "sale.order": [{"id": 1, "name": "S09001", "state": "sale", "picking_ids": [5],
                        "order_line": [11, 12]}],
        "stock.picking": [{"id": 5, "name": "WH/OUT/1", "state": "assigned"}],
        "stock.move": [
            {"id": 21, "product_id": [101, "x"], "product_uom_qty": 1050.0,
             "reserved_availability": 400.0, "quantity_done": 0.0, "state": "partially_available"},
            {"id": 22, "product_id": [102, "y"], "product_uom_qty": 40.0,
             "reserved_availability": 40.0, "quantity_done": 0.0, "state": "assigned"}],
        "stock.move.line": [
            {"move_id": [21, "m"], "lot_id": [7, "L1"], "product_uom_qty": 300.0, "qty_done": 0.0},
            {"move_id": [21, "m"], "lot_id": [8, "L2"], "product_uom_qty": 100.0, "qty_done": 0.0},
            {"move_id": [22, "m"], "lot_id": False, "product_uom_qty": 40.0, "qty_done": 0.0}],
        "stock.production.lot": [{"id": 7, "name": "L1", "expiration_date": "2028-12-30 23:00:00"},
                                 {"id": 8, "name": "L2", "expiration_date": "2029-06-29 22:00:00"}],
        "product.product": [
            {"id": 101, "default_code": "A", "name": "Prod A", "barcode": "111", "hs_code": False,
             **DROPS, "pallet_pieces": False, "weight": 0.1},
            {"id": 102, "default_code": "B", "name": "Prod B", "barcode": False, "hs_code": "3004",
             "carton_pieces": "20", "carton_dimensions": False, "carton_weight": False,
             "pallet_pieces": "8箱/层", "weight": 0.35}],
    }

    def search_read_all(self, model, domain, fields, **kw):
        return [dict(r) for r in self.DATA[model]]


def test_berlin_date():
    assert odoo_source.berlin_date("2028-12-30 23:00:00") == dt.date(2028, 12, 31)
    assert odoo_source.berlin_date("2029-06-29 22:00:00") == dt.date(2029, 6, 30)
    assert odoo_source.berlin_date(False) is None


def test_run_so_end_to_end(tmp_path):
    path, lines = PL.run_so(["s09001"], tmp_path, spare=0, odoo=FakeOdoo())
    assert path.name.endswith("(预估).xlsx")
    ws = openpyxl.load_workbook(path).active
    r = PL.FIRST_DATA_ROW
    # A：两个批次 + 两段箱号（10 整箱 + 尾箱）→ 2 行
    assert ws.cell(r, 6).value == "L1" and ws.cell(r + 1, 6).value == "L2"
    assert ws.cell(r, 10).value.date() == dt.date(2028, 12, 31)
    assert [ws.cell(r + i, 7).value for i in range(2)] == ["托1: 1-10", "托1: 11"]
    assert [ws.cell(r + i, 8).value for i in range(2)] == [100, 50]
    assert ws.cell(r, 12).value == 1050
    assert ws.cell(r, 13).value == "=ROUND(I7*10+I8*10+25,1)"
    assert ws.cell(r, 14).value == "120*80*38"            # 2 层 × 11.5 + 15
    # B：只有箱规 → 单独成托，粗估红字，高度未知
    rb = r + 2
    assert ws.cell(rb, 7).value == "托2: 12-13" and ws.cell(rb, 9).value == 2
    assert ws.cell(rb, 13).value == "=ROUND(I9*7+25,1)"
    assert ws.cell(rb, 14).value == "120*80*?"
    assert ws.cell(rb, 13).font.color.rgb.endswith(E.ROUGH_COLOR)
    assert ws.cell(rb + 1, 2).value == "合计 Total（预估）"
    assert ws.cell(rb + 1, 14).value == "2 托"
    assert any("预留 ≠ 订购" in ln for ln in lines)
    assert any("托1: 11 箱" in ln for ln in lines)
