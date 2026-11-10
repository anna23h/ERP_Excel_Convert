"""so_from_pdf 测试：PDF 文本解析、PZN 匹配裁决、按模板写表。

不连 Odoo、不读真 PDF：解析吃文本、匹配吃假候选表，模板用代码造（*.xlsx 不进仓库）。
文本行照 2026-10-13 实测 Bestellung 的抽取结果构造（行尾粘着右栏联系方式）。
"""
import os
import sys
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from so_import import so_from_pdf as S  # noqa: E402

PAGE1 = """PZN Artikel/Bezeichnung Inhalt Netto Anz. Nettosumme
17604759 PRODUKT A 30ST 11,784 € 192 2.262,53 € Internet:
00427796 PRODUKT B 20ST 7,504 € 192 1.440,77 € www.example.de
05561513 PRODUKT C 40ST 11,032 € 72 794,30 €
Summe 4.497,60 € Serielle DAFÜ:"""
PAGE2 = "PZN Artikel/Bezeichnung Inhalt Netto Anz. Nettosumme\nMit Annahme der vorliegenden Warenbestellung"

COLS = ["ID", "Order Reference", "Customer", S.COL_PID, S.COL_IREF, S.COL_QTY, S.COL_PRICE]


def test_parse_lines_and_summe():
    lines, summe = S.parse_text([PAGE1, PAGE2])
    assert summe == Decimal("4497.60")
    assert [l["pzn"] for l in lines] == ["17604759", "00427796", "05561513"]
    assert lines[0]["netto"] == Decimal("11.784") and lines[0]["qty"] == 192
    assert lines[0]["nsum"] == Decimal("2262.53") and lines[0]["name"] == "PRODUKT A 30ST"


def test_parse_pads_short_pzn():
    text = "427796 PRODUKT B 20ST 7,504 € 192 1.440,77 €\nSumme 1.440,77 €"
    assert S.parse_text([text])[0][0]["pzn"] == "00427796"


@pytest.mark.parametrize("text", [
    "nichts hier",                                                    # 0 行
    "00427796 PRODUKT B 20ST 7,504 € 192 1.440,77 €",                 # 无 Summe
    "00427796 PRODUKT B 20ST 7,504 € 192 1.440,77 €\nSumme 9.999,99 €",  # 合计对不上 = 漏行
])
def test_parse_refuses_unknown_layout(text):
    with pytest.raises(S.StopError):
        S.parse_text([text])


def test_parse_recovers_4dp_price():
    """S04152 形态：印 3 位，Nettosumme ÷ Anz. 整除到 4 位 → 用真实价；整除且相等的行不动。"""
    text = ("01163023 PRODUKT D 100ST 38,865 € 900 34.978,05 €\n"
            "18300412 PRODUKT E 100ST 16,245 € 600 9.747,00 €\n"
            "Summe 44.725,05 €")
    lines, _ = S.parse_text([text])
    assert lines[0]["netto"] == Decimal("38.8645") and lines[0]["printed"] == Decimal("38.865")
    assert lines[1]["netto"] == Decimal("16.245") and "printed" not in lines[1]


def test_parse_keeps_price_when_line_rounded():
    """S04151 形态：对方按行先舍入，Nettosumme ÷ Anz. 除不尽 → 照旧用印的价。"""
    lines, _ = S.parse_text(["12579956 PRODUKT F 20ML 8,224 € 504 4.144,90 €\nSumme 4.144,90 €"])
    assert lines[0]["netto"] == Decimal("8.224") and "printed" not in lines[0]


def test_parse_contradiction_stops():
    """整除到 4 位却舍入不回印的价 → PDF 自相矛盾，停。"""
    with pytest.raises(S.StopError, match="自相矛盾"):
        S.parse_text(["01163023 PRODUKT D 100ST 38,865 € 900 34.900,20 €\nSumme 34.900,20 €"])


def _lines():
    return S.parse_text([PAGE1])[0]


def _h(xid, code, kit=False, name="X"):
    """候选元组：(外部ID, 品名, sale_ok, active, 内部编码, 是否套件)。"""
    return (xid, name, True, True, code, kit)


def test_match_unique_none_multi():
    lines = _lines()
    cands = {"00427796": [("__export__.p_1", "B", False, True, "B_00427796", False)],
             "05561513": [_h("__export__.p_2", "C_05561513"), _h("__export__.p_3", "Temp_C_05561513")]}
    problems = S.match(lines, cands, {})
    assert lines[1]["xmlid"] == "__export__.p_1"
    assert [(l["pzn"], len(h)) for l, h in problems] == [("17604759", 0), ("05561513", 2)]


def test_match_pick_resolves_multi():
    lines = _lines()[2:]
    cands = {"05561513": [_h("__export__.p_2", "C_05561513"), _h("__export__.p_3", "Temp_C_05561513")]}
    assert S.match(lines, cands, {"5561513": "__export__.p_2"}) == []
    assert lines[0]["xmlid"] == "__export__.p_2"


def test_match_pick_must_be_a_hit():
    lines = _lines()[2:]
    cands = {"05561513": [_h("__export__.p_2", "C_05561513")]}
    with pytest.raises(S.StopError):
        S.match(lines, cands, {"05561513": "__export__.p_9"})


@pytest.mark.parametrize("other", [
    _h("__export__.k", "C_05561513x2_GW", kit=True),     # [2 Packs] 套件（后缀 + BoM）
    _h("__export__.k", "C_05561513_GW", kit=True),       # 同名 _GW 套件（S04151 的 Compeed 形态）
    _h("__export__.k", "C_05561513_VO"),                 # 渠道副本，没有 BoM 也排
    _h("__export__.k", "C_99999999", kit=True),          # 基础编码但挂有组件的 phantom（旧 PZN 转向套件）
])
def test_match_auto_excludes_non_base(other):
    lines = _lines()[2:]
    assert S.match(lines, {"05561513": [_h("__export__.base", "C_05561513"), other]}, {}) == []
    assert lines[0]["xmlid"] == "__export__.base"
    assert [h[0] for h, _ in lines[0]["excluded"]] == ["__export__.k"]


def test_match_only_non_base_stops():
    """只有套件 / 副本（含只命中 1 个且是套件）→ 停，不退回去用。"""
    lines = _lines()[2:]
    problems = S.match(lines, {"05561513": [_h("__export__.k", "C_05561513", kit=True)]}, {})
    assert len(problems) == 1 and "xmlid" not in lines[0]


def test_match_pick_non_base_refused():
    lines = _lines()[2:]
    cands = {"05561513": [_h("__export__.base", "C_05561513"), _h("__export__.k", "C_05561513_GW", kit=True)]}
    with pytest.raises(S.StopError, match="非基础品"):
        S.match(lines, cands, {"05561513": "__export__.k"})


def test_not_base_ignores_non_kit_bom():
    """空 BoM / normal BoM 在 odoo_candidates 里就不算套件（kit=False）；基础编码 → 基础品。"""
    assert S.not_base(_h("__export__.o", "Orthomol_10551706")) is None


def test_exclude_unknown_pzn_stops():
    with pytest.raises(S.StopError):
        S.apply_exclude(_lines(), ["99999999"])
    kept, dropped = S.apply_exclude(_lines(), ["17604759"])
    assert len(kept) == 2 and dropped[0]["pzn"] == "17604759"


def _template(path, head=("__export__.sale_order_1_x", "S00001", "Kunde GmbH")):
    wb = Workbook()
    ws = wb.active
    ws.append(COLS)
    ws.append(list(head) + [False, None, None, None])
    wb.save(path)


def test_build_follows_template(tmp_path):
    tpl, out = tmp_path / "t.xlsx", tmp_path / "o.xlsx"
    _template(tpl)
    lines = _lines()
    for i, l in enumerate(lines):
        l["xmlid"] = f"__export__.p_{i}"
        l["rate"] = Decimal("7") if i == 1 else Decimal("19.0")
    wb, cols, head = S.read_template(tpl)
    S.build(wb, cols, head, lines, out)
    rows = list(load_workbook(out).active.iter_rows(values_only=True))
    assert list(rows[0]) == COLS
    assert rows[1][:3] == ("__export__.sale_order_1_x", "S00001", "Kunde GmbH")
    assert rows[2][:3] == (None, None, None) and len(rows) == 4
    assert rows[1][4] is None                       # Internal Reference 留空
    assert rows[1][5] == 192
    assert rows[1][6] == pytest.approx(14.02296)    # 11,784 × 1.19 全精度
    assert lines[2]["price"] == Decimal("13.12808")
    assert lines[1]["price"] == Decimal("8.02928")  # 7,504 × 1.07：按产品真实税率，不写死 1.19


def _tax(id=1, name="19% Umsatzsteuer", amount=19.0, price_include=True, amount_type="percent"):
    return dict(id=id, name=name, amount=amount, price_include=price_include, amount_type=amount_type)


@pytest.mark.parametrize("taxes,fp_src,rate", [
    ([_tax()], set(), Decimal("19")),
    ([_tax(amount=7.0, name="7% Umsatzsteuer")], set(), Decimal("7")),
    ([_tax()], {99}, Decimal("19")),                          # 财政位置映射的是别的税，不碍事
])
def test_tax_resolves(taxes, fp_src, rate):
    l = {}
    assert S.tax_problem(l, taxes, fp_src, "Inland") is None
    assert l["rate"] == rate


@pytest.mark.parametrize("taxes,fp_src", [
    ([], set()),                                               # 没挂税
    ([_tax(), _tax(id=2, amount=7.0)], set()),                 # 多个税
    ([_tax(price_include=False)], set()),                      # 不含税口径
    ([_tax(amount_type="fixed")], set()),                      # 非百分比
    ([_tax()], {1}),                                           # 被财政位置替换
])
def test_tax_ambiguous_stops(taxes, fp_src):
    l = {}
    assert S.tax_problem(l, taxes, fp_src, "EU") and "rate" not in l


def test_apply_taxes_stops_whole_table():
    lines = _lines()
    for i, l in enumerate(lines):
        l["xmlid"], l["erp_name"] = f"__export__.p_{i}", "X"
    info = (None, set(), {"__export__.p_0": [_tax()], "__export__.p_1": [_tax()], "__export__.p_2": []})
    with pytest.raises(S.StopError, match="05561513"):
        S.apply_taxes(lines, info)


def test_template_missing_head_stops(tmp_path):
    tpl = tmp_path / "t.xlsx"
    _template(tpl, head=(None, "S00001", "Kunde GmbH"))
    with pytest.raises(S.StopError):
        S.read_template(tpl)
