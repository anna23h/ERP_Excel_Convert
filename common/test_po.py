"""common/po._po_base_sku 归一规则测试。

跑：python3 common/test_po.py    （或 python3 -m pytest common/test_po.py）

背景：2026-08-08 把规则从 (x\\d+|\\*\\d+|_VO) 扩到 [xX]\\d+|\\*\\d+|_VO|_GW，
统一补货预判/FS 回写/采购频次三处口径（ISSUES [common]）。这里锁住关键行为，
防以后误改。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from common.po import _po_base_sku  # noqa: E402


def test_base_sku():
    cases = {
        # 无后缀：原样不动
        "Sinupret_00605588": "Sinupret_00605588",
        "Ferrum_02190861": "Ferrum_02190861",
        # 多件装 小写 x / 大写 X
        "DOPH_17173992x3": "DOPH_17173992",
        "Balea_049X4": "Balea_049",
        # 变体 *N
        "DOPH_06571703*5": "DOPH_06571703",
        # 渠道/门店 _VO / _GW
        "Ballistol_09060564_VO": "Ballistol_09060564",
        "Schuelke_07463832_GW": "Schuelke_07463832",
        # 组合后缀整体脱（新规则关键：旧规则脱不掉末尾 _GW）
        "Dolormin_02434139x2_GW": "Dolormin_02434139",
        "Stada_17877575x3_GW": "Stada_17877575",
        # 前导/尾随空白也归一
        "  Ferrum_02190861_GW  ": "Ferrum_02190861",
    }
    for raw, expect in cases.items():
        got = _po_base_sku(raw)
        assert got == expect, f"{raw!r} → {got!r}, 期望 {expect!r}"


def test_no_false_strip():
    # 码正文里的数字/字母不被误脱（后缀必须在末尾且匹配候选）
    assert _po_base_sku("Abtei_352000") == "Abtei_352000"
    assert _po_base_sku("1APharma_06312077") == "1APharma_06312077"


if __name__ == "__main__":
    test_base_sku()
    test_no_false_strip()
    print("all tests passed")
