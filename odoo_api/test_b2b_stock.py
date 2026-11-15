"""B2B 可用库存：构造数据单测，不连 Odoo。

    python3 -m pytest odoo_api/test_b2b_stock.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from odoo_api import b2b_stock as B  # noqa: E402

SOS = {
    1: {"name": "S09001", "tags": ["可以打包", "客户甲"], "salesperson": "销售甲", "state": "sale"},
    2: {"name": "S09002", "tags": [], "salesperson": "销售甲", "state": "sale"},
    3: {"name": "VO123", "tags": [], "salesperson": "", "state": "sale"},
    4: {"name": "S09004", "tags": ["客户乙"], "salesperson": "销售乙", "state": "draft"},
    5: {"name": "S09005", "tags": ["可以打包"], "salesperson": "销售乙", "state": "sent"},
    6: {"name": "GW9", "tags": ["可以打包"], "salesperson": "", "state": "sale"},
}
PRODUCTS = {10: {"code": "A", "name": "Prod A"}, 11: {"code": "B", "name": "Prod B"},
            12: {"code": "C", "name": "Prod C"}}


def mv(pid, so_id, demand, reserved, picking="WH/OUT/1"):
    return {"pid": pid, "so_id": so_id, "demand": demand, "reserved": reserved, "picking": picking}


def test_classify():
    assert B.classify(None) == B.CAT_OTHER
    assert B.classify(SOS[3]) == B.CAT_RETAIL
    assert B.classify(SOS[6]) == B.CAT_RETAIL          # 天猫单打了标签也还是天猫
    assert B.classify(SOS[1]) == B.CAT_PAID
    assert B.classify(SOS[2]) == B.CAT_UNPAID


def test_alias_excludes_status_tag():
    assert B.alias_of(["可以打包", "客户甲"]) == "客户甲"
    assert B.alias_of(["可以打包"]) == ""


def test_products_two_available_columns():
    quants = {10: (100.0, 90.0)}
    moves = [mv(10, 1, 50, 40), mv(10, 2, 30, 30), mv(10, 3, 20, 20)]
    df = B.build_products(quants, moves, SOS, [], {}, PRODUCTS)
    r = df.iloc[0]
    assert r["Odoo可用"] == 10                        # 100 − 全部预留 90
    assert r["可用(按已付款)"] == 40                   # 100 − 天猫 20 − 已付款 40
    assert (r["已付款预留"], r["未付款预留"], r["天猫预留"]) == (40, 30, 20)
    lines = r["B2B占用明细"].split("\n")
    assert lines[0] == "S09001 · 已付款 · 40 · 销售甲 · 客户甲"   # 已付款排前；SO 号打头
    assert lines[1] == "S09002 · 未付款 · 30 · 销售甲"            # 没别称就不写


def test_products_quote_intent_not_deducted_and_b2b_first():
    quants = {10: (5.0, 0.0), 11: (8.0, 0.0)}
    quotes = [{"pid": 11, "so_id": 4, "qty": 6.0}]
    df = B.build_products(quants, [], SOS, quotes, {11: 3.0}, PRODUCTS)
    assert list(df["SKU"]) == ["B", "A"]              # 有报价意向的排前
    b = df.iloc[0]
    assert b["报价意向"] == 6 and b["可用(按已付款)"] == 8 and b["在途"] == 3
    assert b["报价明细"] == "S09004 · 报价中 · 6 · 销售乙 · 客户乙"


def test_packing_complete_and_waiting():
    moves = [mv(10, 1, 10, 10), mv(11, 1, 5, 2, "WH/OUT/2"), mv(12, 2, 4, 4)]
    df, short = B.build_packing(moves, SOS, [], PRODUCTS)
    assert set(df["SO"]) == {"S09001"}                 # 未付款单不进待打包
    assert set(df["状态"]) == {"等货"} and set(df["本单货齐"]) == {"否"}
    assert short == {(1, 11): 3.0}
    df2, _ = B.build_packing([mv(10, 1, 10, 10)], SOS, [], PRODUCTS)
    assert list(df2["状态"]) == ["货齐可打包"]


def test_packing_tagged_quote_needs_confirmation():
    df, short = B.build_packing([], SOS, [{"pid": 10, "so_id": 5, "qty": 7.0}], PRODUCTS)
    assert list(df["状态"]) == ["需先确认 SO"] and df.iloc[0]["缺口"] == 7
    assert short == {}                                  # 报价单不进冲突：没有出库单可操作


def test_conflicts_free_stock_and_unpaid_holder():
    quants = {11: (12.0, 11.0)}                         # 空闲 1
    moves = [mv(11, 1, 5, 2), mv(11, 2, 9, 9)]
    df = B.build_conflicts({(1, 11): 3.0}, quants, moves, SOS, PRODUCTS)
    assert list(df["类型"]) == ["有空闲库存没预留上", "被未付款单占着"]
    assert "S09002" in df.iloc[1]["占着货的未付款单"]


def test_conflicts_free_stock_enough_no_unpaid_row():
    quants = {11: (20.0, 11.0)}                         # 空闲 9 ≥ 缺口 3
    moves = [mv(11, 1, 5, 2), mv(11, 2, 9, 9)]
    df = B.build_conflicts({(1, 11): 3.0}, quants, moves, SOS, PRODUCTS)
    assert list(df["类型"]) == ["有空闲库存没预留上"]


def test_notes_negative_and_reservation_mismatch():
    quants = {10: (-6.0, 21.0), 11: (5.0, 2.0), 12: (3.0, 0.0)}
    moves = [mv(10, 2, 30, 21), mv(11, 2, 9, 7)]
    n = B.product_notes(quants, moves)
    assert set(n) == {10, 11}
    assert n[10].startswith("在手为负") and "预留账" not in n[10]
    assert "出库单 7 / 库存 2" in n[11]
    df = B.build_products(quants, moves, SOS, [], {}, PRODUCTS, notes=n)
    assert df.set_index("SKU").loc["B", "备注"] == n[11]
