"""补货预判 动态安全库存 + 套件并入基础品（ISSUES: [安全库存] A）。"""
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_excel as B  # noqa: E402
import safety_dynamic as SD  # noqa: E402
import step4_merge as s4  # noqa: E402


def erp():
    # A 基础品 + A_GW(1×A) + Ax2(2×A)；B 普通品；C 只有套件 Cx2 出现在今天的单里
    return pd.DataFrame({
        s4.ERP_ORDER_REF: ["VO_1", "GW_2", "VO_3", "VO_4", "GW_5"],
        s4.ERP_INTERNAL: ["Ax2", "A_GW", "A", "B", "Cx2"],
        s4.ERP_PICKING: ["ka", "kg", "a", "b", "kc"],
        s4.ERP_BARCODE: ["9", "8", "1", "2", "7"],
        B.ERP_NAME: ["A x2", "A GW", "A base", "B", "C x2"],
        s4.ERP_QTY: [1, 2, 3, 1, 2],
        s4.ERP_ONHAND: [5, 10, 10, 0, 3],
        B.ERP_FS: ["P", "P", "P", "G", "Q"],
        B.ERP_REMARK: ["", "", "", "", ""],
    })


SAFETY = {"kits": {"Ax2": ("A", 2.0), "A_GW": ("A", 1.0), "Cx2": ("C", 2.0)},
          "avg": {"A": 25.04, "B": 9.9, "C": 10.0},
          "onhand": {"A": 11.0, "C": 7.0}}


def test_safety_of_threshold():
    assert SD.safety_of(25.04) == 50 and SD.safety_of(10) == 20
    assert SD.safety_of(9.99) is None and SD.safety_of(None) is None


def test_kits_merged_into_base_row():
    g = B.build_reorder(erp(), safety=SAFETY).set_index("Internal Reference")
    assert sorted(g.index) == ["A", "B", "C"]                  # 套件行不再单列
    assert g.loc["A", "今日需求"] == 3 + 1 * 2 + 2 * 1           # 基础品 + 套件 × 系数
    assert g.loc["A", "Quantity On Hand"] == 11                # 取基础品在手(max 导出 10 / Odoo 11)
    assert g.loc["A", "Name"] == "A base" and g.loc["A", "Picking Name"] == "a"   # 文本列优先基础品
    assert g.loc["A", SD.COL_KITS] == "Ax2×2、A_GW×1" and g.loc["B", SD.COL_KITS] == ""
    assert g.loc["C", "今日需求"] == 4 and g.loc["C", "Quantity On Hand"] == 7   # 只有套件也能并出基础品行
    assert g.loc["C", "缺口"] == -3


def test_safety_columns_and_threshold():
    g = B.build_reorder(erp(), safety=SAFETY).set_index("Internal Reference")
    assert g.loc["A", SD.COL_AVG] == 25.0 and g.loc["A", SD.COL_SAFETY] == 50
    assert g.loc["B", SD.COL_SAFETY] == "" and g.loc["B", SD.COL_AVG] == 9.9   # 周均 < 10 不设
    assert g.loc["C", SD.COL_SAFETY] == 20
    cols = list(g.reset_index().columns)
    assert cols.index(SD.COL_SAFETY) == cols.index(SD.COL_AVG) + 1 and cols[-3] == SD.COL_KITS


def test_without_safety_no_merge_no_columns():
    g = B.build_reorder(erp())
    assert len(g) == 5 and SD.COL_SAFETY not in g.columns and SD.COL_KITS not in g.columns


def test_kit_code_map_single_component_only():
    kit_map = {10: [(1, 2.0)], 11: [(1, 1.0), (2, 1.0)]}
    m, multi = SD.kit_code_map(kit_map, {1: "A", 2: "Z"}, {"Ax2": 10, "AZ": 11, "A": 1})
    assert m == {"Ax2": ("A", 2.0)} and multi == ["AZ"]


def test_written_xlsx(tmp_path):
    from common.xlsx import write_simple
    g = B.build_reorder(erp(), safety=SAFETY)
    p, n = write_simple(g, tmp_path, "r.xlsx", widths=B.reorder_widths(g), wrap_rows=True)
    assert n == 3
