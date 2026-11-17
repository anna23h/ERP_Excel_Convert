"""procure/invoice_reprice 单测（纯文件，不连 Odoo）。样例数据是合成的，不含真实客户 / 价格。"""
import datetime as dt

import pytest
from openpyxl import Workbook, load_workbook

from procure import invoice_reprice as ir

D = dt.date


def item(code, price, start=None, end=None, minq=0, fixed=True):
    return {"code": code, "pname": "", "price": price, "minq": minq, "start": start, "end": end,
            "fixed": fixed}


def line(**kw):
    base = {"line_id": "__export__.l1", "display": "", "code": "Foo_01234567", "label": "Foo",
            "qty": 10, "price": 2.0, "disc": 0.0, "sub": 20.0, "total": 20.0}
    base.update(kw)
    return base


# -- 取价 -------------------------------------------------------------------
def test_pick_price_valid_on_day_with_succession():
    c = [item("A", 1.10, D(2026, 7, 21), D(2026, 9, 7)), item("A", 1.20, D(2026, 9, 8))]
    assert ir.pick_price(c, D(2026, 11, 17))[:2] == (1.20, "匹配")
    assert ir.pick_price(c, D(2026, 8, 1))[:2] == (1.10, "匹配")


def test_pick_price_multiple_valid_takes_latest_start_and_flags():
    c = [item("A", 7.0, D(2026, 7, 21), minq=1), item("A", 8.0, D(2026, 9, 8))]
    price, status, valid = ir.pick_price(c, D(2026, 11, 17))
    assert (price, status, len(valid)) == (8.0, "多条有效", 2)


def test_pick_price_no_start_date_and_none_valid():
    assert ir.pick_price([item("A", 3.0)], D(2026, 1, 1))[:2] == (3.0, "匹配")
    assert ir.pick_price([item("A", 3.0, D(2027, 1, 1))], D(2026, 1, 1))[:2] == (None, "无有效价")
    assert ir.pick_price([item("A", 3.0, fixed=False)], D(2026, 1, 1))[:2] == (None, "非固定价")


# -- 含税口径 ---------------------------------------------------------------
def test_tax_mode():
    assert ir.tax_mode(line())[0] == "0税"
    # 含税价 11.9 × 10 = 119 = Total；Subtotal 100
    assert ir.tax_mode(line(price=11.9, sub=100.0, total=119.0)) == ("含税", pytest.approx(0.19))
    assert ir.tax_mode(line(price=10.0, sub=100.0, total=107.0)) == ("净价", pytest.approx(0.07))
    assert ir.tax_mode(line(price=10.0, sub=50.0, total=60.0))[0] == "无法判断"


# -- 匹配 -------------------------------------------------------------------
def test_build_rows_matches_full_code_and_keeps_unmatched():
    inv = {"lines": [
        line(line_id="x.note", display="Note", code="", qty=0, price=0, sub=0, total=0),
        line(line_id="x.a", code="Foo_01234567"),
        line(line_id="x.gw", code="Foo_01234567_GW"),          # GW 是另一个产品，不能借用非 GW 的价
        line(line_id="x.tax", code="Bar_07654321", price=11.9, sub=100.0, total=119.0),
    ]}
    idx = ir.index_items([item("Foo_01234567", 1.5), item("Bar_07654321", 12.5)])
    rows = {r["line"]["line_id"]: r for r in ir.build_rows(inv, idx, D(2026, 11, 17))}
    assert "x.note" not in rows
    assert rows["x.a"]["new"] == 1.5 and rows["x.a"]["changed"] and not rows["x.a"]["warn"]
    assert rows["x.gw"]["status"] == "配不上" and rows["x.gw"]["new"] == 2.0 and not rows["x.gw"]["changed"]
    assert rows["x.gw"]["pzn"] == "01234567"
    assert rows["x.tax"]["warn"] and "含税" in rows["x.tax"]["notes"]


def test_partner_matches():
    assert ir.partner_matches("ACME Ltd", "002-ACME LTD (EUR)")
    assert not ir.partner_matches("ACME Ltd", "001-Other")


def test_price_date_fallback():
    assert ir.price_date({"inv_date": D(2026, 1, 2), "acc_date": D(2026, 1, 3)})[0] == D(2026, 1, 2)
    assert ir.price_date({"inv_date": None, "acc_date": D(2026, 1, 3)})[0] == D(2026, 1, 3)


# -- 读真实导出版式（表头重名、外部 ID 显示成「ID」） ------------------------
INV_HEADER = ["Status", "ID", "Invoice/Bill Date", "Date", "Partner/ID", "Partner/ID", "Untaxed Amount",
              "Invoice lines/ID", "Invoice lines/Display Type", "Invoice lines/Sequence",
              "Invoice lines/Quantity", "Invoice lines/Unit Price", "Invoice lines/Discount (%)",
              "Invoice lines/Taxes/ID", "Invoice lines/Taxes/ID", "Invoice lines/Account/ID",
              "Invoice lines/Account/ID", "Partner", "Total", "Invoice lines/Product/Internal Reference",
              "Invoice lines/Subtotal", "Invoice lines/Total", "Invoice lines/Label"]


def inv_rows(p1=2.0, p2=3.0, state="Draft", label2="Bar desc"):
    s1, s2 = 10 * p1, 5 * p2
    return [
        [state, "__export__.account_move_1_ab", None, dt.datetime(2026, 11, 17), "9", "__export__.p_9",
         s1 + s2, "__export__.ml_1", "Note", 1, 0, 0, 0, False, None, False, None, "ACME Ltd", s1 + s2,
         False, None, None, "a note"],
        [None, None, None, None, None, None, None, "__export__.ml_2", None, 2, 10, p1, 0, "7", "m.tax7",
         "m.acc", "11", None, None, "Foo_01234567", s1, s1, "Foo desc"],
        [None, None, None, None, None, None, None, "__export__.ml_3", None, 3, 5, p2, 0, "7", "m.tax7",
         "m.acc", "11", None, None, "Bar_07654321_GW", s2, s2, label2],
    ]


def save(path, header, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return str(path)


PL_HEADER = ["Sequence", "Pricelist Name", "ID", "Pricelist Items/Product/External ID",
             "Pricelist Items/Fixed Price", "Pricelist Items/Min. Quantity", "Pricelist Items/Start Date",
             "Pricelist Items/End Date", "Pricelist Items/Compute Price",
             "Pricelist Items/Product/Internal Reference", "Pricelist Items/Product/Name"]


def pl_rows():
    return [
        [1, "002-ACME LTD", "12", "__export__.t_1", 1.5, 0, dt.datetime(2026, 9, 8, 12, 0), None,
         "Fixed Price", "Foo_01234567", "Foo"],
        [None, None, None, "__export__.t_1", 1.8, 0, dt.datetime(2026, 7, 1), dt.datetime(2026, 9, 7),
         "Fixed Price", "Foo_01234567", "Foo"],
        [None, None, None, "__export__.t_2", 2.5, 1, dt.datetime(2026, 7, 1), None,
         "Fixed Price", "Bar_07654321_GW", "Bar GW"],
    ]


def test_parse_invoice_picks_xmlid_columns(tmp_path):
    inv = ir.parse_invoice(save(tmp_path / "inv.xlsx", INV_HEADER, inv_rows()))
    assert inv["move_id"] == "__export__.account_move_1_ab"
    assert inv["partner"] == "ACME Ltd" and inv["state"] == "Draft"
    assert [l["line_id"] for l in inv["lines"]] == ["__export__.ml_1", "__export__.ml_2", "__export__.ml_3"]
    assert inv["inv_date"] is None and inv["acc_date"] == D(2026, 11, 17)


def test_parse_invoice_missing_column_stops(tmp_path):
    hdr = [h for h in INV_HEADER if h != "Invoice lines/Subtotal"]
    rows = [r[:20] + r[21:] for r in inv_rows()]
    with pytest.raises(ir.InputError, match="Subtotal"):
        ir.parse_invoice(save(tmp_path / "inv.xlsx", hdr, rows))


def test_build_end_to_end_and_refuses_posted(tmp_path):
    invp = save(tmp_path / "inv.xlsx", INV_HEADER, inv_rows())
    plp = save(tmp_path / "pl.xlsx", PL_HEADER, pl_rows())
    out = ir.build(invp, plp, str(tmp_path / "out"))
    ws = load_workbook(out["import"]).active
    assert list(ws.iter_rows(values_only=True)) == [
        ("id", "invoice_line_ids/id", "invoice_line_ids/price_unit"),
        ("__export__.account_move_1_ab", "__export__.ml_2", 1.5),
        (None, "__export__.ml_3", 2.5),
    ]
    posted = save(tmp_path / "inv2.xlsx", INV_HEADER, inv_rows(state="Posted"))
    with pytest.raises(ir.InputError, match="不是草稿"):
        ir.build(posted, plp, str(tmp_path / "out"))


def test_verify_accepts_price_change_and_flags_other_fields(tmp_path):
    before = save(tmp_path / "b.xlsx", INV_HEADER, inv_rows())
    imp = save(tmp_path / "imp.xlsx", ["id", "invoice_line_ids/id", "invoice_line_ids/price_unit"],
               [["__export__.account_move_1_ab", "__export__.ml_2", 1.5], [None, "__export__.ml_3", 2.5]])
    ok = save(tmp_path / "ok.xlsx", INV_HEADER, inv_rows(1.5, 2.5))
    problems, changes = ir.verify(before, ok, imp)
    assert problems == [] and len(changes) == 8

    bad = save(tmp_path / "bad.xlsx", INV_HEADER, inv_rows(1.5, 2.6, label2="changed"))
    problems, _ = ir.verify(before, bad, imp)
    text = "\n".join(problems)
    assert "Label" in text and "单价应为 2.5" in text
