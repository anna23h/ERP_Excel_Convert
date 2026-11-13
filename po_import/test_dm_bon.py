"""dm_bon 测试（版式 D，ISSUES [PO 导入] H）：小票 OCR 文本解析与护栏、对照表容错匹配、税率分档反验、写表。

不跑 OCR、不连 Odoo：OCR 文本照 2026-11-13 实测那张小票 rapidocr 4 倍渲染的输出构造（品名噪声、税码漏读 / 读成 52 都照搬），
产品是造的。
"""
import os
import sys
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from po_import import dm_bon as D  # noqa: E402
from po_import import po_from_pdf as P  # noqa: E402

OCR = [
    "dm-drogerie markt",
    "Marlene-Dietrich-StraBe 4",
    "13.11.202608:43D585/1324850/2 6855",
    "tetesept Femi PabylabltWeichk. 8,95 52",
    "2x 3,95 lirkulin PropolisHals 7.90 2",
    "5x 3,25 tetesept Nasen Heilsal 16,25",
    "5x 4,95 tetesept Schnupfenspra 24, 75 1",
    "12x 6,45 Doppe lherz Umega 1000 77,40 2",
    "3x 5,95 Abtei Ealdrian Forte 3 17,85 1",
    "2x 4,75Zirkulin Muskel-Gelenk 9,50 1",
    "SUMME EUR 162,60",
    "MasterCard EUR -162,60",
    "Mwst-Sat2 Brutto Net to MuSt",
    "1=19,00% 68,35 57,44 10,91",
    "2=7,00% 94,25 88,08 6,17",
    "Start: 2026-11-13 08:43:20",
    "SN-Kasse: FD307711 TA-Nummer: 841711",
]

MAP = [  # 小票原名（正确拼写）→ 外部 ID
    ("tetesept Femi Babylabl+Weichk.", "__export__.p1", 7),
    ("Zirkulin Propolis Hals", "__export__.p2", 7),
    ("tetesept Nasen Heilsal", "__export__.p3", 19),
    ("tetesept Schnupfenspra", "__export__.p4", 19),
    ("Doppelherz Omega 1000", "__export__.p5", 7),
    ("Abtei Baldrian Forte 3", "__export__.p6", 19),
    ("Zirkulin Muskel-Gelenk", "__export__.p7", 19),
]


def _map(tmp_path, rows=MAP, extra=()):
    p = tmp_path / "dm_bon_map.csv"
    p.write_text("小票品名,外部ID,备注\n" + "".join(f"{n},{x},\n" for n, x, *_ in list(rows) + list(extra)),
                 encoding="utf-8")
    return str(p)


def _prods(rates=None, code=lambda x: "Base_" + x[-2:]):
    rates = rates or {x: r for _, x, r in MAP}
    tax = {7: P.TAX_REDUCED[1], 19: P.TAX_FULL[1]}
    return lambda xids: {x: ("ERP " + x, code(x), True, (tax[rates[x]],), int(x[-1]), "0000000" + x[-1])
                         for x in xids if x in rates}


def _load(tmp_path, ocr=OCR, map_rows=MAP, **kw):
    return D.load("bon.pdf", _map(tmp_path, map_rows), ocr_fn=lambda _p, _s: ocr, products_fn=kw.pop("prods", _prods()))


def test_parse_numbers_doc_and_tax_codes():
    bon = D.parse_dm(OCR)
    assert bon["doc_no"] == "dm-20261113-841711" and bon["summe"] == Decimal("162.60")
    ls = bon["lines"]
    assert [l["qty"] for l in ls] == [1, 2, 5, 5, 12, 3, 2]
    assert ls[3]["brutto"] == Decimal("24.75") and ls[1]["unit"] == Decimal("3.95")
    # 税码只认单独的 1 / 2：`52` 当没读到，漏读的也是 None
    assert [l["code"] for l in ls] == [None, "2", None, "1", "2", "1", "1"]
    assert bon["bands"][19] == (Decimal("68.35"), Decimal("57.44"), Decimal("10.91"))


@pytest.mark.parametrize("old,new", [
    ("16,25", "16,35"),                       # 数量 × 单价 ≠ 小计
    ("SUMME EUR 162,60", "SUMME EUR 162,70"),  # 合计对不上
    ("68,35 57,44 10,91", "68,35 57,40 10,95"),  # 档内 Netto 与 Brutto ÷ 1.19 不符
    ("Start: 2026-11-13", "Start: 2026-11-12"),  # 头部日期 ≠ TSE 日期
    ("TA-Nummer: 841711", "TA-Nummer: 841/11"),  # TA 不是纯数字
    ("dm-drogerie markt", "Rossmann"),
    ("2x 4,75Zirkulin Muskel-Gelenk 9,50 1", "Rabatt -0,50"),  # 负数行
])
def test_parse_refuses(old, new):
    with pytest.raises(P.StopError):
        D.parse_dm([t.replace(old, new) for t in OCR])


def test_qty_on_its_own_line():
    ocr = [t for t in OCR]
    i = ocr.index("2x 4,75Zirkulin Muskel-Gelenk 9,50 1")
    ocr[i:i + 1] = ["2 x 4,75", "Zirkulin Muskel-Gelenk 9,50 1"]
    assert D.parse_dm(ocr)["lines"][-1]["qty"] == 2


def test_read_bon_tries_next_scale():
    calls = []

    def ocr(_p, sc):
        calls.append(sc)
        return [t.replace("16,25", "16,35") for t in OCR] if sc == 4 else OCR
    assert D.read_bon("x.pdf", ocr)["scale"] == 3 and calls == [4, 3]


def test_load_matches_noisy_names_and_prices(tmp_path):
    lines, dropped, doc, bon = _load(tmp_path)
    assert doc == "dm-20261113-841711" and not dropped
    assert [l["xmlid"] for l in lines] == [x for _, x, _ in MAP]
    assert all(l["score"] >= D.MIN_SCORE for l in lines)
    # 含税单价 ÷ 1.19 / 1.07，2 位
    assert [str(l["aep"]) for l in lines] == ["8.36", "3.69", "2.73", "4.16", "6.03", "5.00", "3.99"]
    assert [l["reduced"] for l in lines] == [True, True, False, False, True, False, False]
    assert bon["netto_diff"] == {19: Decimal("-0.01"), 7: Decimal("0.02")}


def test_unmatched_or_ambiguous_stops(tmp_path):
    with pytest.raises(P.StopError, match="没有唯一匹配"):
        _load(tmp_path, map_rows=MAP[:-1])           # 表里没有 Muskel-Gelenk
    with pytest.raises(P.StopError, match="没有唯一匹配"):
        _load(tmp_path, map_rows=MAP + [("tetesept Nasen Heilsalbe", "__export__.p9", 19)])  # 两条太接近


def test_wrong_product_caught_by_bands(tmp_path):
    rows = [r if r[1] != "__export__.p6" else (r[0], "__export__.p5", 7) for r in MAP]  # Baldrian 配成 7% 的产品
    rates = {x: r for _, x, r in rows}
    with pytest.raises(P.StopError):
        _load(tmp_path, map_rows=rows, prods=_prods(rates))


def test_ocr_code_conflicts_with_product_tax(tmp_path):
    rates = {x: r for _, x, r in MAP}
    rates["__export__.p4"] = 7                     # 小票读到 1(=19%)，产品却是 7%
    with pytest.raises(P.StopError, match="税码"):
        _load(tmp_path, prods=_prods(rates))


def test_variant_and_missing_product_stop(tmp_path):
    with pytest.raises(P.StopError, match="副本"):
        _load(tmp_path, prods=_prods(code=lambda x: "Base_" + x[-2:] + ("_VO" if x.endswith("1") else "")))
    rates = {x: r for _, x, r in MAP if not x.endswith("7")}
    with pytest.raises(P.StopError, match="查不到"):
        _load(tmp_path, prods=_prods(rates))


def test_skip_row_goes_to_dropped(tmp_path):
    rows = [r if r[1] != "__export__.p7" else (r[0], D.SKIP, 19) for r in MAP]
    lines, dropped, _doc, bon = _load(tmp_path, map_rows=rows)
    assert len(lines) == 6 and [l["bon_name"] for l in dropped] == ["Zirkulin Muskel-Gelenk"]
    assert 19 not in bon["netto_diff"]              # 该档含 SKIP 行，不比净额


def test_map_errors(tmp_path):
    with pytest.raises(P.StopError, match="不存在"):
        D.load_map(str(tmp_path / "nope.csv"))
    with pytest.raises(P.StopError, match="归一化后相同"):
        D.load_map(_map(tmp_path, MAP, [("TETESEPT nasen-heilsal", "__export__.x", 19)]))


def _template(tmp_path, vendor="DM-drogerie markt"):
    cols = ["id", "partner_ref", "name", "partner_id", "order_line/product_id/.id", "order_line/product_qty",
            "order_line/price_unit", "order_line/taxes_id"]
    wb = Workbook()
    ws = wb.active
    ws.append(cols)
    ws.append(["__export__.purchase_order_1", None, "P00001", vendor, None, None, None, None])
    p = tmp_path / "t.xlsx"
    wb.save(p)
    return str(p)


def _run(tmp_path, monkeypatch, template, **kw):
    monkeypatch.setattr(P, "read_pdf", lambda _: [""])
    monkeypatch.setattr(D, "ocr_lines", lambda _p, _s: OCR)
    monkeypatch.setattr(D, "odoo_products", _prods())
    return P.run("scandm.pdf", template, str(tmp_path / "out"), existing_fn=lambda _: [],
                 dm_map=_map(tmp_path), **kw)


def test_run_writes_dm_po(tmp_path, monkeypatch):
    out = _run(tmp_path, monkeypatch, _template(tmp_path))
    rows = list(load_workbook(out).active.iter_rows(values_only=True))
    assert rows[1][:4] == ("__export__.purchase_order_1", "dm-20261113-841711", "P00001", "DM-drogerie markt")
    assert [r[4:] for r in rows[1:3]] == [(1, 1, 8.36, P.TAX_REDUCED[1]), (2, 2, 3.69, P.TAX_REDUCED[1])]
    assert rows[3][7] == P.TAX_FULL[1] and len(rows) == 8


def test_run_refuses_other_vendor_and_exclude(tmp_path, monkeypatch):
    with pytest.raises(P.StopError, match="不是 dm"):
        _run(tmp_path, monkeypatch, _template(tmp_path, vendor="AEP GmbH"))
    assert not os.listdir(tmp_path / "out")
    with pytest.raises(P.StopError, match="SKIP"):
        _run(tmp_path, monkeypatch, _template(tmp_path), exclude=["123"])
