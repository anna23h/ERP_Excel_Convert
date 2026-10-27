"""补货预判 待出库/在途 测试（ISSUES: [补货预判] A）：套件折算 / 新四列算术 / 不传时版式不变。"""
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_excel as B  # noqa: E402
import step4_merge as s4  # noqa: E402
from pending_moves import to_sku_pending  # noqa: E402


def test_plain_sku():
    r = to_sku_pending(["A"], {"A": 1}, {}, inc={1: 5}, out={1: 8})
    assert r == {"A": (8, 5)}


def test_kit_converted_to_sets():
    # A_x2 是 2× A 的套件：组件 A 待出 8 / 在途 6 → 套件口径 4 / 3
    r = to_sku_pending(["A", "A_x2"], {"A": 1, "A_x2": 2}, {2: [(1, 2.0)]},
                       inc={1: 6}, out={1: 8})
    assert r["A"] == (8, 6) and r["A_x2"] == (4, 3)


def test_multi_component_kit_takes_biggest_drain():
    r = to_sku_pending(["K"], {"K": 9}, {9: [(1, 1.0), (2, 1.0)]},
                       inc={1: 0, 2: 10}, out={1: 3, 2: 4})
    assert r["K"] == (3, 0)      # 组件 1 净消耗 3 > 组件 2 的 −6


def test_unknown_sku_absent():
    assert to_sku_pending(["X"], {}, {}, {}, {}) == {}


def erp():
    return pd.DataFrame({
        s4.ERP_ORDER_REF: ["VO_1", "VO_2", "VO_3"],
        s4.ERP_INTERNAL: ["A", "A", "B"],
        s4.ERP_PICKING: ["a", "a", "b"],
        s4.ERP_BARCODE: ["1", "1", "2"],
        B.ERP_NAME: ["n", "n", "m"],
        s4.ERP_QTY: [2, 3, 1],
        s4.ERP_ONHAND: [10, 10, 0],
        B.ERP_SAFETY: [0, 0, 0],
        B.ERP_FS: ["P", "P", "G"],
        B.ERP_REMARK: ["", "", ""],
    })


def test_build_reorder_pending_columns():
    g = B.build_reorder(erp(), pending={"A": (8, 0), "B": (0, 4)}).set_index("Internal Reference")
    assert list(g.columns[4:9]) == ["缺口", "待出库", "在途", "预计可用", "预测缺口"]
    assert g.loc["A", "缺口"] == -5 and g.loc["A", "预计可用"] == 2 and g.loc["A", "预测缺口"] == 3
    assert g.loc["B", "缺口"] == 1 and g.loc["B", "预计可用"] == 4 and g.loc["B", "预测缺口"] == -3


def test_build_reorder_sku_missing_in_odoo_counts_zero():
    g = B.build_reorder(erp(), pending={}).set_index("Internal Reference")
    assert g.loc["A", "待出库"] == 0 and g.loc["A", "预测缺口"] == g.loc["A", "缺口"]


def test_build_reorder_without_pending_unchanged():
    g = B.build_reorder(erp())
    assert "预测缺口" not in g.columns and list(g.columns[5:7]) == ["缺口", "Safety Stock"]


def test_layout_pzn_barcode_tail_cols():
    e = erp()
    e[B.ERP_NAME] = ["Creme 30ml (new PZN 19231788)", "Creme 30ml (new PZN 19231788)", "Gel （NEW PZN 1786749)"]
    e[B.ERP_FS] = ["Queisser Pharma/DM/A", "Queisser Pharma/DM/A", "DMS/G"]
    g = B.build_reorder(e).set_index("Internal Reference")
    assert g.columns[0] == "PZN" and list(g.columns[-2:]) == ["Picking Name", "Name"]
    assert g.loc["A", "PZN"] == "19231788" and g.loc["B", "PZN"] == "01786749"
    assert g.loc["A", "Barcode"] == "1" and g.loc["B", "Barcode"] == ""   # DMS 不算 DM


def test_name_pzn_variants():
    for s, want in [("x (PZN18487422)  MHD2027/06", "18487422"), ("x (New PZN17550928)", "17550928"),
                    ("x （NEW PZN 17867499)", "17867499"), ("x (PZN18", ""), ("plain", ""), (None, "")]:
        assert B._name_pzn(s) == want, s


def test_barcode_vendor():
    assert B._barcode_vendor("DM") and B._barcode_vendor("B/ROSSMAN") and B._barcode_vendor("Rossmann/P")
    assert not B._barcode_vendor("DMS") and not B._barcode_vendor("") and not B._barcode_vendor(None)


def test_reorder_widths_fit_content_not_header():
    df = pd.DataFrame({"Quantity On Hand": [5, 128], "Safety Stock": [0, 60],
                       "供应商(次数)": ["P(5)、G(2)", "无采购记录"], "最低价供应商": ["P", ""]})
    w = B.reorder_widths(df)
    assert w["Quantity On Hand"] < 14 and w["Safety Stock"] < 12      # 原来被表头撑到 26 / 20
    assert w["供应商(次数)"] < 25 and w["最低价供应商"] < 15            # 原来固定 45 / 22
    assert w["Name"] == 30
    assert w["最低价供应商"] < 8                       # 中文表头可逐字折行，不再按六个字撑宽


def test_pzn_width_fits_8_digits_one_line():
    w = B.reorder_widths(pd.DataFrame({"PZN": ["19623519", ""]}))["PZN"]
    assert w >= 8 * 1.4                                 # 2026-10-27 实测 9.8 宽时 8 位数字被拆成两行
