"""po_from_pdf 测试：PDF 文本解析与四道护栏、PZN 匹配裁决、防重、按模板写表。

不连 Odoo、不读真 PDF：解析吃文本、匹配吃假候选表，模板用代码造（*.xlsx 不进仓库）。
文本行照 2026-10-20 实测 Rechnung/Lieferschein 的 pdfplumber 抽取结果构造（€ 抽成乱码），品名与价格是造的。
"""
import os
import sys
from decimal import Decimal

import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from po_import import po_from_pdf as P  # noqa: E402

ROWS = [
    "17173992 2 120ST KAP PRODUKT A S M 10,00 0,00 20,00",
    "02758221 5 100ST HKP PRODUKT B RETARD P S 4,50 9,95 22,50",
    "20101140 10 15ML OEL PRODUKT C S 1,25 0,00 12,50",
]
FOOT = "RX-DPP:1,00 � 3/17 35,00 � 20,00 � 55,00 �\n**** Endsumme **** 55,00 �"
PAGE = "1234567 1\n" + "\n".join(ROWS) + "\n" + FOOT + "\nKUNDE GmbH\n7654321\n20.10.2026"

COLS = ["ID", "Order Reference", "Vendor", P.COL_VREF, P.COL_PID, P.COL_QTY, P.COL_PRICE, P.COL_TAX_ID]


def test_parse_lines_doc_and_tax_class():
    lines, doc = P.parse_text([PAGE])
    assert doc == "1234567"
    assert [l["pzn"] for l in lines] == ["17173992", "02758221", "20101140"]
    assert [l["reduced"] for l in lines] == [True, False, False]
    # 品名末尾的 P 不是税码字母，留在品名里
    assert lines[1]["name"] == "100ST HKP PRODUKT B RETARD P" and lines[1]["code"] == "S"
    assert lines[0]["aep"] == Decimal("10.00") and lines[2]["qty"] == 10


@pytest.mark.parametrize("bad", [
    PAGE.replace("3/17", "4/17"),                       # 行数对不上 = 漏行
    PAGE.replace("3/17", "3/18"),                       # 件数对不上
    PAGE.replace("PRODUKT A S M", "PRODUKT A S"),       # 税码误读 → 两档税基对不上
    PAGE.replace("**** Endsumme **** 55,00", "**** Endsumme **** 56,00"),
    PAGE.replace("4,50 9,95 22,50", "4,50 9,95 22,60"),  # 行内 AEP×Menge ≠ Gesamt
    PAGE.replace(FOOT, ""),                              # 没有底部合计
    PAGE.replace("1234567 1\n", ""),                     # 没有单号
    "nichts hier",
])
def test_parse_refuses(bad):
    with pytest.raises(P.StopError):
        P.parse_text([bad])


def _lines():
    return P.parse_text([PAGE])[0]


def _cand(xid, n=0, taxes=(P.TAX_FULL[1],), dbid=0):
    return (xid, "ERP " + xid, "REF_" + xid, True, n, taxes, dbid)


def test_match_unique_none_multi_and_pick():
    lines = _lines()
    cands = {"17173992": [_cand("__export__.a")],
             "02758221": [_cand("__export__.b", 75), _cand("__export__.b_vo")]}
    problems = P.match(lines, cands, {})
    assert lines[0]["xmlid"] == "__export__.a"
    assert [(l["pzn"], len(h)) for l, h in problems] == [("02758221", 2), ("20101140", 0)]

    lines = _lines()
    problems = P.match(lines, cands, {"2758221": "__export__.b"})
    assert lines[1]["xmlid"] == "__export__.b"
    assert [l["pzn"] for l, _ in problems] == ["20101140"]
    with pytest.raises(P.StopError):
        P.match(_lines(), cands, {"02758221": "__export__.nope"})


def _template(tmp_path, cols=COLS):
    wb = Workbook()
    ws = wb.active
    ws.append(cols)
    ws.append(["__export__.purchase_order_1", "P00001", "Vendor X"] + [None] * (len(cols) - 3))
    p = tmp_path / "purchase.order.xlsx"
    wb.save(p)
    return str(p)


def _run(tmp_path, monkeypatch, template, existing=(), **kw):
    monkeypatch.setattr(P, "read_pdf", lambda _: [PAGE])
    cands = {"17173992": [_cand("__export__.a", dbid=11)], "02758221": [_cand("__export__.b", dbid=12)],
             "20101140": [_cand("__export__.c", dbid=13)]}
    return P.run("L 1.pdf", template, str(tmp_path / "out"), candidates_fn=lambda _: cands,
                 existing_fn=lambda _: list(existing), **kw)


def test_run_writes_template(tmp_path, monkeypatch):
    out = _run(tmp_path, monkeypatch, _template(tmp_path))
    rows = list(load_workbook(out).active.iter_rows(values_only=True))
    assert list(rows[0]) == COLS
    assert rows[1] == ("__export__.purchase_order_1", "P00001", "Vendor X", "1234567",
                       "__export__.a", 2, 10.0, P.TAX_REDUCED[0])
    assert rows[2] == (None, None, None, None, "__export__.b", 5, 4.5, P.TAX_FULL[0])
    assert len(rows) == 4


def test_run_tax_by_name_and_exclude(tmp_path, monkeypatch):
    cols = ["ID", "Order Reference", "Vendor", P.COL_PID, P.COL_QTY, P.COL_PRICE, P.COL_TAX]
    out = _run(tmp_path, monkeypatch, _template(tmp_path, cols), exclude=["20101140"])
    rows = list(load_workbook(out).active.iter_rows(values_only=True))
    assert [r[-1] for r in rows[1:]] == [P.TAX_REDUCED[1], P.TAX_FULL[1]]


def test_run_refuses_duplicate(tmp_path, monkeypatch):
    with pytest.raises(P.StopError, match="已有"):
        _run(tmp_path, monkeypatch, _template(tmp_path), existing=[("P12345", "done")])
    assert not (tmp_path / "out").exists()


def test_template_missing_tax_column(tmp_path):
    with pytest.raises(P.StopError, match="Taxes"):
        P.read_template(_template(tmp_path, COLS[:-1]))


def test_template_total_quantity_is_explained(tmp_path):
    cols = [P.COL_TOTAL_QTY if c == P.COL_QTY else c for c in COLS]
    with pytest.raises(P.StopError, match="只读"):
        P.read_template(_template(tmp_path, cols))


def test_report_flags_default_tax_mismatch(tmp_path, monkeypatch, capsys):
    # 17173992 在 PDF 是 M（7%），候选默认税给的是 19% → 提示；其余两行一致 → 不提示
    _run(tmp_path, monkeypatch, _template(tmp_path))
    tail = capsys.readouterr().out.split("不一致")[1].split("合计核对")[0]
    assert "17173992" in tail and "02758221" not in tail and "20101140" not in tail


def test_run_technical_names_with_db_id(tmp_path, monkeypatch):
    # 2026-10-20 用户第 28 份模板：技术名表头，产品列是数据库 ID（.id），税列按名称
    cols = ["id", "partner_ref", "name", "partner_id", "order_line/product_id/.id",
            "order_line/product_qty", "order_line/price_unit", "order_line/taxes_id"]
    wb = Workbook()
    ws = wb.active
    ws.append(cols)
    ws.append(["__export__.purchase_order_1", None, "P00001", "Vendor X", False, None, None, None])
    p = tmp_path / "purchase.order.xlsx"
    wb.save(p)
    out = _run(tmp_path, monkeypatch, str(p))
    assert os.path.basename(out).startswith("P00001_import_")
    rows = list(load_workbook(out).active.iter_rows(values_only=True))
    assert rows[1] == ("__export__.purchase_order_1", "1234567", "P00001", "Vendor X", 11, 2, 10.0,
                       P.TAX_REDUCED[1])
    assert rows[2] == (None, None, None, None, 12, 5, 4.5, P.TAX_FULL[1])


def test_template_prefers_external_product_id(tmp_path, monkeypatch):
    cols = COLS + ["order_line/product_id/.id"]
    rows = list(load_workbook(_run(tmp_path, monkeypatch, _template(tmp_path, cols))).active
                .iter_rows(values_only=True))
    assert rows[1][4] == "__export__.a" and rows[1][-1] is None
