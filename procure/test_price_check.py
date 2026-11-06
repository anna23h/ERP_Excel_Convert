"""procure/price_check 纯函数单测（不连 Odoo）。"""
import pytest

from procure.price_check import build_row, pzn_of, strip_code, totals, unit_net

#: 本库实况：销售税含税价，采购税不含税（2026-11-06 现查）
TAXES = {
    14: {"label": "19%", "amount": 19.0, "amount_type": "percent", "price_include": True},
    15: {"label": "7%", "amount": 7.0, "amount_type": "percent", "price_include": True},
    23: {"label": "19%", "amount": 19.0, "amount_type": "percent", "price_include": False},
    99: {"label": "fix", "amount": 1.0, "amount_type": "fixed", "price_include": False},
}


def test_pzn_prefers_field_then_code_tail():
    assert pzn_of({"pzn": "08491730", "default_code": "Emser_X"}) == "08491730"
    assert pzn_of({"pzn": False, "default_code": "Temp_Biolectra_17605598"}) == "17605598"
    assert pzn_of({"pzn": "", "default_code": "NOPZN"}) == "NOPZN"


def test_strip_code():
    assert strip_code("[Emser_08491730] Emser Inh.Lsg. 60St\n4260048437640") == "Emser Inh.Lsg. 60St"


def test_unit_net_removes_included_tax_only():
    assert unit_net(12.8, 0, [14], TAXES) == pytest.approx(12.8 / 1.19)
    assert unit_net(10.0, 10, [15], TAXES) == pytest.approx(9.0 / 1.07)
    assert unit_net(10.06, 0, [23], TAXES) == pytest.approx(10.06)          # 采购税不含税：原价即净价


def test_unit_net_independent_of_qty_rounding():
    """同一单价、不同数量：按单价算不受行级舍入影响（subtotal÷qty 会差 1 分）。"""
    assert unit_net(8.94, 0, [14], TAXES, 1442.42, 192) == unit_net(8.94, 0, [14], TAXES, 300.50, 40)


def test_unit_net_falls_back_for_non_percent_tax():
    assert unit_net(10.0, 0, [99], TAXES, 50.0, 4) == pytest.approx(12.5)


def test_tax_mismatch_flagged_but_same_net_is_not_price_change():
    """S04094 实况：SO 挂 19% 含税 8.94，发票挂 7% 含税 8.03，净价只差舍入。"""
    so_net, inv_net = unit_net(8.94, 0, [14], TAXES), unit_net(8.03, 0, [15], TAXES)
    po = {"net": 7.04, "code": "P1", "date": "2026-10-20"}
    r = build_row("00427796", "X", "7%", 40, so_net, inv_net, po, so_tax="19%")
    assert "SO 税率 19% ≠ 开票税率 7%" in r["备注"]
    assert "开票价 ≠ SO 价" not in r["备注"]
    assert r["_flag"]


def test_price_change_and_negative_margin_flagged():
    po = {"net": 5.0, "code": "P1", "date": "2026-10-20"}
    r = build_row("1", "X", "7%", 1, 5.5, 4.9, po, so_tax="7%")
    assert "开票价 ≠ SO 价" in r["备注"] and "毛利 ≤ 0" in r["备注"]
    assert r["单件毛利"] == pytest.approx(-0.1)


def test_missing_po_kept_out_of_cost():
    a = build_row("1", "A", "7%", 2, 10.0, 10.0, {"net": 8.0, "code": "P", "date": "d"})
    b = build_row("2", "B", "7%", 3, 10.0, 10.0, None)
    for r, sub in ((a, 20.0), (b, 30.0)):
        r["_inv_subtotal"] = sub
    a["_po_net"], b["_po_net"] = 8.0, None
    assert "发票日前无采购记录" in b["备注"]
    assert totals([a, b]) == (50.0, 16.0, True)
