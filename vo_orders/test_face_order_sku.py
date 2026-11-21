"""面单按 SKU 聚合排序 + 分隔页（ISSUES: [拣货面单] O）。

数据全是虚构的：SKU 名 / 单号 / 运单号都是造的。PDF 用 test_sort_labels 的最小 PDF 造法。
"""
import os
import sys

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_excel as B  # noqa: E402
import sort_labels as S  # noqa: E402
import step4_merge as s4  # noqa: E402
from test_sort_labels import lp, make_pdf  # noqa: E402

REF = s4.ERP_ORDER_REF


def ref(n):
    return f"VO_TOF_SCP{n:012d}"


def rows(orders):
    """orders: [(单号序号, [(SKU, 件数), …], 面单时间)] → (SKU 明细行, 面单时间 Series)。"""
    out, t = [], {}
    for n, lines, ts in orders:
        r = ref(n)
        t[r[-15:]] = pd.Timestamp(ts)
        for sku, q in lines:
            out.append({REF: r, "_key": r[-15:], s4.ERP_INTERNAL: sku, s4.ERP_QTY: q,
                        s4.ERP_TRACKING: lp(n), s4.ERP_PICKING: f"{sku} 方盒\n",
                        B.ERP_NAME: f"{sku} Langname"})
    return pd.DataFrame(out), pd.Series(t)


# A：3 单共 4 件（一单 ×2）；b：2 单共 2 件（小写，测不分大小写）；C：2 单共 2 件；
# 单 9 / 10 是多品订单。b 与 C 件数相同 → 按 Internal Reference 不分大小写升序，b 在 C 前。
SUB, T = rows([
    (1, [("C", 1)], "2026-11-01 09:00"),
    (2, [("A", 1)], "2026-11-01 08:00"),
    (3, [("A", 2)], "2026-11-01 10:00"),
    (4, [("b", 1)], "2026-11-01 07:00"),
    (5, [("A", 1)], "2026-11-01 06:00"),
    (6, [("C", 1)], "2026-11-01 05:00"),
    (9, [("A", 1), ("C", 1)], "2026-11-01 11:00"),
    (7, [("b", 1)], "2026-11-01 04:00"),
    (10, [("b", 2), ("A", 1)], "2026-11-01 03:00"),
])


def pack_seq(df):
    """文件顺序倒过来 = 打包顺序（操作员从表底往上打包）。"""
    return [int(r[-3:]) for r in pd.unique(df[REF])][::-1]


def test_packing_order_groups_by_total_qty_then_multi_last():
    out, seps = B.order_by_sku(SUB, T)
    # A 组：×2 的单 3 先，然后单件单按面单时间早→晚 5、2；b 组 7、4；C 组 6、1；多品 10、9
    assert pack_seq(out) == [3, 5, 2, 7, 4, 6, 1, 10, 9]
    assert len(out) == len(SUB)


def test_multiline_orders_stay_contiguous_and_keep_line_order():
    out, _ = B.order_by_sku(SUB, T)
    refs = out[REF].tolist()
    for r in set(refs):
        idx = [i for i, x in enumerate(refs) if x == r]
        assert idx == list(range(idx[0], idx[-1] + 1))
    assert out[out[REF] == ref(10)][s4.ERP_INTERNAL].tolist() == ["b", "A"]


def test_separators_follow_packing_order_and_anchor_on_first_packed_order():
    _, seps = B.order_by_sku(SUB, T)
    assert seps["打包顺序"].tolist() == [1, 2, 3]
    assert seps["类型"].tolist() == [B.SEP_SKU] * 3          # 多品订单不成组、不插分隔页
    assert seps["Internal Reference"].tolist() == ["A", "b", "C"]
    assert seps["张数"].tolist() == [3, 2, 2]
    assert seps["件数"].tolist() == [4, 2, 2]
    assert seps["插在运单后"].tolist() == [lp(3), lp(7), lp(6)]
    assert seps.loc[0, "拣货名"] == "A 方盒"          # Picking Name 的换行压掉
    assert seps.loc[0, "德文品名"] == "A Langname"


def test_name_columns_optional():
    sub = SUB.drop(columns=[B.ERP_NAME])
    _, seps = B.order_by_sku(sub, T)
    assert (seps["德文品名"] == "").all()


def test_only_groups_with_3_or_more_labels_get_a_separator():
    _, seps = B.order_by_sku(SUB, T)
    ins = B.separators_to_insert(seps)
    assert ins["Internal Reference"].tolist() == ["A"]       # b / C 各 2 张不插
    assert ins["打包顺序"].tolist() == [1]                      # 保留 SKU 组序号(不印出来)


def test_two_labels_with_many_pieces_still_no_separator():
    sub, t = rows([(1, [("A", 5)], "2026-11-01"), (2, [("A", 5)], "2026-11-02")])
    _, seps = B.order_by_sku(sub, t)
    assert B.separators_to_insert(seps).empty                # 门槛看张数，不看件数


def test_missing_label_time_goes_last_within_group():
    sub, t = rows([(1, [("A", 1)], "2026-11-02"), (2, [("A", 1)], "2026-11-01")])
    t = t.drop(ref(1)[-15:])
    out, _ = B.order_by_sku(sub, t)
    assert pack_seq(out) == [2, 1]


def test_limit_selection_is_unchanged_by_sku_order():
    sel, *_ = B.order_by_label_time(SUB, T, limit=4)
    out, _ = B.order_by_sku(sel, T)
    assert set(out[REF]) == set(sel[REF])            # 只重排，不增不减
    assert set(sel[REF]) == {ref(n) for n in (10, 7, 6, 5)}   # 面单时间最早的 4 单


def _seps_with_titles():
    _, seps = B.order_by_sku(SUB, T)
    seps.insert(B.SEP_COLS.index("中文品名"), "中文品名", ["甲中文名", "", "丙"])
    return seps


def test_hidden_sheet_round_trip(tmp_path):
    wb = Workbook()
    wb.active.title = "面单"
    B._write_separators(wb, _seps_with_titles())
    p = tmp_path / "x.xlsx"
    wb.save(p)
    assert load_workbook(p)[B.SEP_SHEET].sheet_state == "hidden"
    got = S.read_separators(p)
    assert [s["插在运单后"] for s in got] == [lp(3), lp(7), lp(6)]
    assert got[0]["中文品名"] == "甲中文名" and got[1]["中文品名"] == ""
    assert S.SEP_REQUIRED == tuple(B.SEP_COLS)      # 两边列名约定一致


def test_read_separators_absent_returns_empty(tmp_path):
    wb = Workbook()
    wb.active.title = "面单"
    p = tmp_path / "x.xlsx"
    wb.save(p)
    assert S.read_separators(p) == []


@pytest.fixture
def sku_batch(tmp_path):
    """按 SKU 排好的拣货表+面单（只含 面单 + 隐藏分隔页两个 sheet）+ 打乱顺序的面单 PDF。"""
    pytest.importorskip("pypdfium2")
    pytest.importorskip("reportlab")
    out, _ = B.order_by_sku(SUB, T)
    wb = Workbook()
    ws = wb.active
    ws.title = "面单"
    ws.append(["序号", "Order Reference", "VO Tracking No", "Internal Reference"])
    for i, (r, t, sku) in enumerate(zip(out[REF], out[s4.ERP_TRACKING], out[s4.ERP_INTERNAL]), 1):
        ws.append([i, r, t, sku])
    B._write_separators(wb, B.separators_to_insert(_seps_with_titles()))
    x = tmp_path / "2026年11月21日VO9单 拣货表+面单.xlsx"
    wb.save(x)
    pdf = tmp_path / "document.pdf"
    make_pdf(pdf, [lp(n) for n in (1, 2, 3, 4, 5, 6, 7, 9, 10)])
    return x, pdf, out


def test_run_inserts_separator_below_each_group(sku_batch, tmp_path):
    import pypdfium2 as pdfium
    x, pdf, out = sku_batch
    path, lines = S.run(x, [pdf], tmp_path / "o")
    file_lps = list(dict.fromkeys(out[s4.ERP_TRACKING]))
    doc = pdfium.PdfDocument(path)
    texts = [doc[i].get_textpage().get_text_bounded() for i in range(len(doc))]
    w, h = doc[0].get_size()
    sizes = {tuple(round(v) for v in doc[i].get_size()) for i in range(len(doc))}
    doc.close()
    assert len(texts) == 9 + 1                        # 只有 A 组（3 张）插
    assert sizes == {(round(w), round(h))}            # 分隔页与面单同尺寸
    seq = []
    for t in texts:
        found = S.LP_RE.findall(t)
        seq.append(found[0] if found else "SEP")
    expect = []
    anchors = {lp(3)}
    for l in file_lps:
        expect.append(l)
        if l in anchors:
            expect.append("SEP")
    assert seq == expect
    # 分隔页紧跟在每组最先打包那单下面 → 倒着翻时先看到分隔页
    assert any("插分隔页 1 页" in ln for ln in lines)
    sep_text = texts[seq.index("SEP")]
    assert "分隔页" in sep_text and "组" not in sep_text      # 不印编号


def test_run_rejects_separator_anchor_not_in_facesheet(sku_batch, tmp_path):
    x, pdf, _ = sku_batch
    wb = load_workbook(x)
    wb[B.SEP_SHEET].cell(2, B.SEP_COLS.index("插在运单后") + 1).value = lp(99)
    wb.save(x)
    with pytest.raises(S.MismatchError, match="分隔页"):
        S.run(x, [pdf], tmp_path / "o")
    assert not (tmp_path / "o").exists() or not os.listdir(tmp_path / "o")


def test_render_separators_one_page_each():
    pytest.importorskip("reportlab")
    import pypdfium2 as pdfium
    data, _ = S.render_separators(S_sep_dicts(), 288, 432)
    doc = pdfium.PdfDocument(data)
    assert len(doc) == 2
    assert "SKU-LONG" in doc[0].get_textpage().get_text_bounded()
    doc.close()


def S_sep_dicts():
    return [{"打包顺序": 1, "类型": "SKU", "Internal Reference": "SKU-LONG_12345678", "拣货名": "蓝白方盒 2350",
             "德文品名": "Sehr langer deutscher Produktname mit vielen Wörtern 120 Kapseln",
             "中文品名": "很长的中文商品标题用于测试自动换行是否正常工作一二三四五六七八九十",
             "张数": 12, "件数": 14, "插在运单后": lp(1)},
            {"打包顺序": 2, "类型": "多品", "Internal Reference": "", "拣货名": "", "德文品名": "", "中文品名": "",
             "张数": 3, "件数": 7, "插在运单后": lp(2)}]


def test_face_order_validated():
    with pytest.raises(ValueError, match="face_order"):
        B.build([], None, face_order="abc")
