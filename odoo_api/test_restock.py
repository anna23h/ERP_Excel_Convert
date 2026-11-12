"""备货汇总 + 询价单回填：构造数据单测，不连 Odoo。

    python3 -m pytest odoo_api/test_restock.py
"""
import os
import sys

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from odoo_api import quote_fill as Q  # noqa: E402
from odoo_api import restock as R  # noqa: E402
from so_import.so_from_pdf import StopError  # noqa: E402


def test_prefix_domain_is_prefix_match():
    # `like` 会包成 %VO% 变子串；前缀必须 =like
    assert R.prefix_domain() == ["|", ("name", "=like", "VO%"), ("name", "=like", "GW%")]


def test_roll_sales_folds_kits_into_base():
    kit = {20: [(10, 2.0)], 21: [(10, 1.0)]}          # 20 = 基础品 10 ×2；21 = 10 的 _GW 副本
    rolled = R.roll_sales([(10, "W1", 3), (20, "W1", 4), (21, "W2", 5), (30, "W2", 1)], kit)
    assert dict(rolled[10]) == {"W1": 11.0, "W2": 5.0}
    assert dict(rolled[30]) == {"W2": 1.0}
    assert 20 not in rolled and 21 not in rolled


def test_sales_stats_peak_weeks_trend():
    st = R.sales_stats({"W1": 10, "W2": 0, "W3": 30, "W4": 20}, recent_qty=20, n_days=56, recent_days=28)
    assert st["峰值周销量"] == 30 and st["有销周数"] == 3 and st["销量"] == 60
    assert st["周均销量"] == pytest.approx(60 / 8) and st["近4周周均"] == pytest.approx(5)
    assert st["趋势"] == pytest.approx(5 / 10 - 1)                    # −50%
    assert R.sales_stats({"W1": 5}, recent_qty=5, n_days=56)["趋势"] is None   # 之前没卖过 = 新起量


def test_demand_base_low_drop_default():
    # 2026-11-08 用户定：近 4 周周均 ×1.2；近 7 天 < 12 周周均一半 → 近 4 周周均；近 4 周周均 < 3 → 有单再补
    assert R.demand_base(10, 9, 10) == (pytest.approx(12), "近4周周均×1.2")
    assert R.demand_base(10, 4, 10) == (10, "近期下降·近4周周均")
    assert R.demand_base(2.9, 0, 10) == (0.0, "低量·有单再补")        # 低量优先于下降
    assert R.demand_base(5, 0, 0) == (pytest.approx(6), "近4周周均×1.2")   # 无 12 周均值不判下降


def test_classify_share_threshold_and_suspect():
    # 大单 300 / 总 400 = 75% → 大批量
    assert R.classify({"P1": 300, "P2": 50, "P3": 50})[:2] == (True, 0.75)
    # 有过一张大单但量只占 120/420 → 不算大批量（2026-11-08 用户定按占比，不按"有过"）
    bulk, share, _ = R.classify({"P1": 120, "P2": 100, "P3": 100, "P4": 100})
    assert not bulk and share == pytest.approx(120 / 420)
    # 单张 >5000 疑录错：不进分子分母
    bulk, share, sus = R.classify({"P1": 20000, "P2": 10, "P3": 10})
    assert sus == ["P1"] and not bulk and share == 0
    assert R.classify({}) == (False, None, [])


def test_suggest_qty_subtracts_onhand_and_transit():
    assert R.suggest_qty(30, 3, 20, 10) == 60
    assert R.suggest_qty(10.2, 1, 0, 0) == 11          # 向上取整
    assert R.suggest_qty(10, 1, 8, 5) == 0              # 在途已够，不出负数
    assert R.suggest_qty(0, 3, -2, 0) == 0              # 低量基数 0：负在手也不出数


def test_vendor_profile_last_lowest_and_sole():
    lines = [{"date": "2026-01-05", "vendor": "A", "qty": 10, "net": 5.0, "po": "P1"},
             {"date": "2026-03-01", "vendor": "B", "qty": 10, "net": 6.0, "po": "P2"},
             {"date": "2026-02-01", "vendor": "A", "qty": 1, "net": 0.0, "po": "P3"}]   # 0 元赠品不算最低价
    p = R.vendor_profile(lines)
    assert (p["最近一次供应商"], p[R.COL_LAST_NET]) == ("B", 6.0)
    assert (p["最低价供应商"], p["最低净价"]) == ("A", 5.0)
    assert p["供应商数"] == 2 and p[R.COL_QUOTE_VENDOR] == ""
    assert R.vendor_profile(lines[:1])[R.COL_QUOTE_VENDOR] == "A"


def test_low_profile_window_last_price_and_tie():
    # 2026-11-12 用户定（[备货提醒] D）：近期窗口内最低净价那家；单价取该家窗口内最近一次；并列不算
    lines = [{"date": "2025-12-01 10:00:00", "vendor": "Old", "qty": 1, "net": 1.0, "po": "P0"},   # 窗口外
             {"date": "2026-10-13 09:00:00", "vendor": "B", "qty": 1, "net": 8.25, "po": "P1"},   # 一次性低价
             {"date": "2026-11-03 09:00:00", "vendor": "B", "qty": 1, "net": 14.56, "po": "P2"},
             {"date": "2026-11-01 09:00:00", "vendor": "A", "qty": 1, "net": 14.50, "po": "P3"},
             {"date": "2026-11-02 09:00:00", "vendor": "A", "qty": 1, "net": 0.0, "po": "P4"}]    # 赠品
    p = R.low_profile(lines, "2026-08-12")
    assert (p[R.COL_LOW_VENDOR], p[R.COL_LOW_NET], p[R.COL_LOW_LAST]) == ("B", 8.25, 14.56)
    assert R.vendor_profile(lines, "2026-08-12")[R.COL_LOW_VENDOR] == "B"
    assert R.low_profile(lines)[R.COL_LOW_VENDOR] == "Old"                      # 不给窗口 = 全部
    tie = lines[2:4] + [{"date": "2026-11-04", "vendor": "C", "qty": 1, "net": 14.504, "po": "P5"}]
    assert R.low_profile(tie, "2026-08-12")[R.COL_LOW_VENDOR] == ""             # A 与 C 两位小数并列
    assert R.low_profile(lines, "2026-12-01")[R.COL_LOW_VENDOR] == ""           # 窗口内无采购


def test_quote_rows_sole_plus_recent_lowest():
    df = pd.DataFrame([
        {"SKU": "S", R.COL_SUGGEST: 3, R.COL_QUOTE_VENDOR: "B", R.COL_LAST_NET: 2.0,
         R.COL_LOW_VENDOR: "B", R.COL_LOW_LAST: 2.0},
        {"SKU": "L", R.COL_SUGGEST: 5, R.COL_QUOTE_VENDOR: "", R.COL_LAST_NET: 9.0,       # 最近一次是别家价
         R.COL_LOW_VENDOR: "B", R.COL_LOW_LAST: 7.5},
        {"SKU": "Z", R.COL_SUGGEST: 0, R.COL_QUOTE_VENDOR: "", R.COL_LAST_NET: 1.0,       # 不需补
         R.COL_LOW_VENDOR: "B", R.COL_LOW_LAST: 1.0},
        {"SKU": "G", R.COL_SUGGEST: 4, R.COL_QUOTE_VENDOR: "", R.COL_LAST_NET: 1.0,
         R.COL_LOW_VENDOR: "G", R.COL_LOW_LAST: 1.0},
    ])
    sub, price, src = R.quote_rows(df, "B")
    assert sub["SKU"].tolist() == ["S", "L"]
    assert price.tolist() == [2.0, 7.5] and src.tolist() == ["独家", "近期最低价"]
    assert R.rfq_vendors(df) == ["B", "G"]
    qv = R.quote_vendors(df).set_index(R.COL_QUOTE_VENDOR)
    assert qv.loc["B", "本周需补品数"] == 2 and qv.loc["B", "其中近期最低价"] == 1
    assert qv.loc["B", "建议数量合计"] == 8 and qv.loc["G", "独家常销品数"] == 0


def _build(**kw):
    weeks = {f"W{i}": 10 for i in range(7)}                                  # 7 周 × 10 = 70 件 / 70 天 → 周均 7
    sales = {1: weeks, 2: dict(weeks, W0=50), 3: {"W1": 5, "W2": 5}}       # 3 只有 2 周有销 → 不进
    po = {1: [{"date": "2026-01-01", "vendor": "Solo", "qty": 500, "net": 2.0, "po": "P1"},
              {"date": "2026-02-01", "vendor": "Solo", "qty": 20, "net": 2.1, "po": "P2"}],
          2: [{"date": "2026-01-01", "vendor": "X", "qty": 10, "net": 3.0, "po": "P3"},
              {"date": "2026-01-02", "vendor": "Y", "qty": 10, "net": 2.5, "po": "P4"}]}
    prods = {1: {"code": "A_1", "name": "大批量独家"}, 2: {"code": "B_2", "name": "常规多家"},
             3: {"code": "C_3", "name": "偶发"}}
    kw.setdefault("last7", {1: 10, 2: 10})
    return R.build_rows({1, 2, 3}, sales, {1: 40, 2: 40}, 70, po, {1: 5, 2: 60}, {1: 0, 2: 0}, prods, **kw)


def test_build_rows_tiers_filter_and_sort():
    df, checks = _build()
    assert df["SKU"].tolist() == ["A_1", "B_2"]          # 偶发品不进；需补的在前
    a, b = df.iloc[0], df.iloc[1]
    # 近 4 周 40 件 → 周均 10，×1.2 = 12；大批量 3 周：12×3 − 5 = 31
    assert a["档位"] == "大批量·3周" and a[R.COL_SUGGEST] == 31 and a[R.COL_QUOTE_VENDOR] == "Solo"
    assert a["基数口径"] == "近4周周均×1.2"
    assert b["档位"] == "常规·1周" and b[R.COL_SUGGEST] == 0 and b["最低价供应商"] == "Y"
    assert checks == []
    qv = R.quote_vendors(df)
    assert qv.iloc[0].to_dict()["需补 SKU"] == "A_1"


def test_build_rows_drop_and_low_volume():
    df, _ = _build(last7={1: 2, 2: 10})                    # 1 近 7 天 2 件 < 周均 7 的一半 → 近期下降
    a = df.set_index("SKU").loc["A_1"]
    assert a["基数口径"] == "近期下降·近4周周均" and a[R.COL_SUGGEST] == 10 * 3 - 5
    df, _ = _build(min_recent_avg=11)                       # 近 4 周周均 10 < 11 → 低量
    assert set(df["基数口径"]) == {"低量·有单再补"} and df[R.COL_SUGGEST].sum() == 0


def test_build_rows_flags_unrolled_and_no_purchase():
    df, checks = _build(unrolled={2})
    assert "有后缀无 BoM" in df.set_index("SKU").loc["B_2", "备注"]
    df, checks = R.build_rows({9}, {9: {f"W{i}": 1 for i in range(6)}}, {}, 84, {}, {}, {},
                              {9: {"code": "Z_9", "name": "无采购"}})
    assert checks[0]["类型"] == "常销无采购记录" and df.loc[0, "供应商数"] == 0


# --- quote_fill ------------------------------------------------------------
COLS = ["ID", "Order Reference", "Vendor", "Order Lines/Product/ID", "Order Lines/Quantity",
        "Order Lines/Unit Price", "Order Lines/Taxes/ID"]


def _summary(tmp_path):
    df = pd.DataFrame([
        {"SKU": "A_1", "产品": "a", R.COL_SUGGEST: 12, R.COL_LAST_NET: 2.5, R.COL_QUOTE_VENDOR: "Solo GmbH", "产品ID": 1},
        {"SKU": "A_2", "产品": "b", R.COL_SUGGEST: 0, R.COL_LAST_NET: 3.0, R.COL_QUOTE_VENDOR: "Solo GmbH", "产品ID": 2},
        {"SKU": "B_1", "产品": "c", R.COL_SUGGEST: 5, R.COL_LAST_NET: 1.0, R.COL_QUOTE_VENDOR: "", "产品ID": 3},
    ])
    d = tmp_path / "20261111"
    d.mkdir()
    p = d / "备货汇总-20261111.xlsx"
    with pd.ExcelWriter(p) as w:
        df.to_excel(w, sheet_name=R.MAIN_SHEET, index=False)
    return str(p)


def _template(tmp_path, vendor="solo gmbh"):
    wb = Workbook()
    ws = wb.active
    ws.append(COLS)
    ws.append(["__export__.po_1", "P00001", vendor, None, None, None, None])
    p = tmp_path / "t.xlsx"
    wb.save(p)
    return str(p)


def _prods(taxes=("19% Vorsteuer",), xmlids=("__export__.product_product_1",)):
    return lambda pids: {pid: {"xmlids": list(xmlids), "code": "", "taxes": taxes} for pid in pids}


def _blank(ids):
    return {i: ("draft", 0, "") for i in ids}


def test_quote_fill_writes_only_vendor_rows_with_need(tmp_path):
    out, lines = Q.run(_template(tmp_path), _summary(tmp_path), outdir=str(tmp_path / "o"),
                       status_fn=_blank, products_fn=_prods())
    rows = list(load_workbook(out).active.iter_rows(values_only=True))
    assert len(rows) == 2                                 # 表头 + A_1；A_2 不需补、B_1 不是这家
    assert rows[1] == ("__export__.po_1", "P00001", "solo gmbh", "__export__.product_product_1", 12, 2.5,
                       "l10n_de_skr03.1_tax_vst_19_skr03")


def test_quote_fill_adds_recent_lowest_rows(tmp_path):
    df = pd.DataFrame([
        {"SKU": "A_1", "产品": "a", R.COL_SUGGEST: 12, R.COL_LAST_NET: 2.5, R.COL_QUOTE_VENDOR: "Solo GmbH",
         R.COL_LOW_VENDOR: "Solo GmbH", R.COL_LOW_LAST: 2.5, "产品ID": 1},
        {"SKU": "B_1", "产品": "c", R.COL_SUGGEST: 5, R.COL_LAST_NET: 9.0, R.COL_QUOTE_VENDOR: "",
         R.COL_LOW_VENDOR: "Solo GmbH", R.COL_LOW_LAST: 7.25, "产品ID": 3},
        {"SKU": "C_1", "产品": "d", R.COL_SUGGEST: 5, R.COL_LAST_NET: 1.0, R.COL_QUOTE_VENDOR: "",
         R.COL_LOW_VENDOR: None, R.COL_LOW_LAST: None, "产品ID": 4},
    ])
    p = tmp_path / "备货汇总-20261112.xlsx"
    with pd.ExcelWriter(p) as w:
        df.to_excel(w, sheet_name=R.MAIN_SHEET, index=False)
    out, lines = Q.run(_template(tmp_path), str(p), outdir=str(tmp_path / "o"),
                       status_fn=_blank, products_fn=_prods())
    assert [(l["sku"], l["price"], l["src"]) for l in lines] == [("A_1", 2.5, "独家"), ("B_1", 7.25, "近期最低价")]
    rows = list(load_workbook(out).active.iter_rows(values_only=True))
    assert len(rows) == 3 and rows[2][3:6] == ("__export__.product_product_1", 5, 7.25)
    # 只有近期最低价品、没有独家品的供应商也能出询价单
    out, lines = Q.run(_template(tmp_path, "Solo"), str(p), outdir=str(tmp_path / "o2"),
                       status_fn=_blank, products_fn=_prods())
    assert len(lines) == 2


def test_quote_fill_stops(tmp_path):
    s = _summary(tmp_path)
    with pytest.raises(StopError, match="不是本周汇总里的询价单供应商"):
        Q.run(_template(tmp_path, "Other AG"), s, status_fn=_blank, products_fn=_prods())
    with pytest.raises(StopError, match="不是空白询价单"):
        Q.run(_template(tmp_path), s, status_fn=lambda ids: {i: ("purchase", 3, "") for i in ids},
              products_fn=_prods())
    with pytest.raises(StopError, match="进项税"):
        Q.run(_template(tmp_path), s, status_fn=_blank, products_fn=_prods(taxes=("0% EU",)))
    with pytest.raises(StopError, match="外部 ID"):
        Q.run(_template(tmp_path), s, status_fn=_blank, products_fn=_prods(xmlids=()))


def test_latest_summary_picks_newest(tmp_path):
    for d in ("20261104", "20261111"):
        (tmp_path / d).mkdir()
        (tmp_path / d / f"备货汇总-{d}.xlsx").write_bytes(b"")
    assert Q.latest_summary(str(tmp_path)).endswith("备货汇总-20261111.xlsx")
    with pytest.raises(StopError):
        Q.latest_summary(str(tmp_path / "none"))
