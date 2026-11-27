"""return_intake 测试：假 Odoo 数据 + 代码造的退货导出 / 入库单导出，走「sheet → 填表 → po → lot」全程。

*.xlsx 一律 gitignore，真实导出（含订单号）不进仓库。
"""
import os
import sys
from datetime import datetime

import pytest
from openpyxl import Workbook, load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from return_intake import return_intake as R  # noqa: E402

V19 = "l10n_de_skr03.1_tax_vst_19_skr03"
V7 = "l10n_de_skr03.1_tax_vst_7_skr03"


class FakeData:
    """与 R.OdooData 同样的属性 / 方法，不连 Odoo。"""

    def __init__(self, taken=()):
        P = lambda code, pzn, name, tax=V19, active=True, typ="product": {
            "code": code, "pzn": pzn, "name": name, "active": active, "type": typ, "tax_xmlids": [tax]}
        self.prod = {
            1: P("Euc_18222103", "18222103", "Eucerin Augenpflege"),
            2: P("Hueb_446071", "", "Huebner Silicea"),
            3: P("Hueb_446071x3", "", "[3 Packs] Huebner Silicea"),
            4: P("Doph_15994609", "19765739", "Doppelherz (New PZN19765739)", tax=V7),
            5: P("Doph_15994609x2", "", "[2 Packs] Doppelherz", tax=V7),
            6: P("Extra_123456", "12345678", "Extra Produkt"),
            7: P("Old_1", "1111111", "Archiviert", active=False),
            8: P("Multi_1", "", "Multi Kit"),
        }
        self.kit = {3: [(2, 3.0)], 5: [(4, 2.0)], 8: [(1, 1.0), (2, 1.0)]}
        self.so_lines = {
            "SCP001": [(1, 1.0)],
            "SCP002": [(3, 2.0)],                  # 套件 ×2 → 基础品 6 件
            "SCP003": [(4, 1.0), (5, 1.0)],         # 基础品 1 + 套件 1（×2）→ 同包合并 3 件
            "SCP004": [(1, 2.0)],
        }
        self.taken = set(taken)

    def load_orders(self, scps):
        pass

    def load_products(self, pids):
        pass

    def refs_taken(self, n):
        return ["P999"] if n in self.taken else []

    def by_pzn(self, pzn):
        return [k for k, v in self.prod.items() if v["pzn"] == pzn and k not in self.kit]


HEAD = ["退货申请批次号", "退货申请人", "退货申请时间", "系统履约单号", "店铺订单号", "物流LP单号", "大包号",
        "异常拦截状态", "商品数量"]


def _export(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(HEAD)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return str(path)


def _std_exports(tmp_path):
    a = _export(tmp_path / "a.xlsx", [
        ("AN1", "x", "2026-11-03 18:54:07", "SCP002", "o2", "LP0002", "BAG1", "已出库", 2),
        ("AN1", "x", "2026-11-03 18:54:07", "SCP001", "o1", "LP0003", "BAG1", "已出库", 1),
        ("AN9", "x", "2026-11-09 10:00:00", "SCP004", "o4", "LP0009", "BAG2", "已打包", 2),
    ])
    b = _export(tmp_path / "b.xlsx", [   # 同一大包另一份导出（11-03 的 1 单）→ 并进 BAG1
        ("AN2", "x", "2026-11-03 22:29:00", "SCP003", "o3", "LP0001", "BAG1", "已出库", 2),
    ])
    return [a, b]


# ---------------------------------------------------------------- 小件
@pytest.mark.parametrize("sku,pzn,want", [
    ("Euc_18222103", "18222103", ""),
    ("Orth_17444652x2", "17444652_", ""),        # PZN 只比数字
    ("Doph_15994609x2", "19765739", "19765739"),  # 新 PZN 不同 → 保留
    ("Hueb_446071", "", ""),
])
def test_blank_pzn(sku, pzn, want):
    assert R.blank_pzn(sku, pzn) == want


def test_read_exports_merges_by_bag_and_sorts(tmp_path):
    bs = R.read_exports(_std_exports(tmp_path))
    assert [b["bag"] for b in bs] == ["BAG1", "BAG2"]
    assert len(bs[0]["orders"]) == 3 and bs[0]["batch_nos"] == ["AN1", "AN2"]
    assert R.vendor_ref(139, bs[0]) == "汇总139   3单ECMS退货包裹2026年11月03日"


def test_read_exports_duplicate_scp_stops(tmp_path):
    a = _export(tmp_path / "a.xlsx", [("AN1", "x", "2026-11-03 18:54:07", "SCP001", "o", "LP1", "B", "已出库", 1)])
    with pytest.raises(R.StopError):
        R.read_exports([a, a])


def test_read_exports_wrong_file(tmp_path):
    wb = Workbook()
    wb.active.append(["foo", "bar"])
    wb.save(tmp_path / "x.xlsx")
    with pytest.raises(R.InputError):
        R.read_exports([str(tmp_path / "x.xlsx")])


def test_assemble_kits_and_order(tmp_path):
    b = R.read_exports(_std_exports(tmp_path))[0]
    parcels = R.assemble(b, FakeData())
    assert [p["lp"] for p in parcels] == ["LP0001", "LP0002", "LP0003"]          # 按 LP 号
    assert [(i["pid"], i["soll"]) for i in parcels[0]["items"]] == [(4, 3.0)]    # 基础品 1 + 套件 1×2
    assert [(i["pid"], i["soll"]) for i in parcels[1]["items"]] == [(2, 6.0)]


@pytest.mark.parametrize("mutate,msg", [
    (lambda d: d.so_lines.pop("SCP001"), "找不到"),
    (lambda d: d.so_lines.update(SCP001=[(1, 5.0)]), "≠ 导出商品数量"),
    (lambda d: d.so_lines.update(SCP001=[(8, 1.0)]), "多组件套件"),
    (lambda d: d.so_lines.update(SCP001=[(7, 1.0)]), "已归档"),
    (lambda d: d.prod[1].update(tax_xmlids=["x.other"]), "进项税"),
])
def test_assemble_guards(tmp_path, mutate, msg):
    b = R.read_exports(_std_exports(tmp_path))[0]
    d = FakeData()
    mutate(d)
    with pytest.raises(R.StopError, match=msg):
        R.assemble(b, d)


# ---------------------------------------------------------------- sheet
def test_sheet_outputs_and_taken_number(tmp_path):
    files, lines = R.cmd_sheet(_std_exports(tmp_path), 139, str(tmp_path), FakeData())
    names = [os.path.basename(f) for f in files]
    assert names == ["汇总139.xlsx", "点货表_汇总139.xlsx", "汇总140.xlsx", "点货表_汇总140.xlsx"]
    assert any("已打包" in l for l in lines)
    s = list(load_workbook(files[0]).active.iter_rows(values_only=True))
    assert s[0] == ("SCP单号", "LP单号", "SKU", "PZN", "产品名称", "数量")
    assert s[-1] == ("合计", "3 单", None, None, None, 5)
    pzn = {r[2]: r[3] for r in s[1:-1]}
    assert pzn["Euc_18222103"] in ("", None) and pzn["Doph_15994609x2"] == "19765739"
    with pytest.raises(R.StopError, match="汇总140"):
        R.cmd_sheet(_std_exports(tmp_path), 139, str(tmp_path / "x"), FakeData(taken={140}))


def _rows(ws):
    """→ {(LP, SKU): 产品行号}。一个产品一行，LP 在每行的 A 列。"""
    return {(ws.cell(r, R.C_LP).value, ws.cell(r, R.C_SKU).value): r for r in range(2, ws.max_row + 1)
            if ws.cell(r, R.C_TECH0).value == "P"}


def _make_sheet(tmp_path):
    files, _ = R.cmd_sheet(_std_exports(tmp_path), 139, str(tmp_path), FakeData())
    return files[1]


def _fill(path, out, edit):
    wb = load_workbook(path)
    ws = wb[R.SHEET]
    edit(ws, _rows(ws))
    wb.save(out)
    return str(out)


def _all_got(ws, items):
    """每行勾「到」、不填实收（全部到齐）。"""
    for r in items.values():
        ws.cell(r, R.C_YES, 1)


def _fill_ok(ws, items):
    """LP0001：Doppelherz 3 件，两个批次 2 + 1，其中 1 件不可售；LP0002 未到（实收 0）；
    LP0003：Eucerin 1 件，批次 + MHD；多出一件 PZN 12345678。"""
    _all_got(ws, items)
    ws.cell(items[("LP0002", "Hueb_446071")], R.C_YES).value = None
    ws.cell(items[("LP0002", "Hueb_446071")], R.C_NO, True)        # Excel 复选框存 TRUE
    r = items[("LP0001", "Doph_15994609")]
    ws.cell(r, R.C_GOT, 2)
    ws.cell(r, R.C_LOT, "AB123")
    ws.cell(r, R.C_MHD, "05.29")
    ws.insert_rows(r + 1)
    ws.cell(r + 1, R.C_GOT, 1)
    ws.cell(r + 1, R.C_BAD, 1)
    ws.cell(r + 1, R.C_LOT, "CD456")
    ws.cell(r + 1, R.C_MHD, "06.2029")
    items = _rows(ws)
    r = items[("LP0003", "Euc_18222103")]
    ws.cell(r, R.C_LOT, "E1001")
    ws.cell(r, R.C_MHD, "31.12.2028")
    ws.insert_rows(r + 1)
    ws.cell(r + 1, R.C_PZN, "12345678")
    ws.cell(r + 1, R.C_GOT, 1)
    ws.cell(r + 1, R.C_NOTE, "客户多寄")


def test_sheet_layout_one_row_per_product(tmp_path):
    ws = load_workbook(_make_sheet(tmp_path))[R.SHEET]
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[R.C_TECH0 - 1] == "P"]
    assert [(r[0], r[2], r[5]) for r in rows] == [("LP0001", "Doph_15994609", 3), ("LP0002", "Hueb_446071", 6),
                                                  ("LP0003", "Euc_18222103", 1)]
    assert all(ws.cell(r, c).fill.fill_type is None for r in range(2, 5) for c in range(1, len(R.HEADS) + 1))
    assert all(ws.row_dimensions[r].height == 20 for r in range(1, 5))


def test_po_full_flow(tmp_path):
    sheet = _fill(_make_sheet(tmp_path), tmp_path / "filled.xlsx", _fill_ok)
    files, lines = R.cmd_po([sheet], str(tmp_path / "po"), FakeData())
    diff, po = files
    assert os.path.basename(po) == "汇总139_PO导入.xlsx"
    rows = list(load_workbook(po)["purchase.order"].iter_rows(values_only=True))
    assert rows[0] == tuple(R.PO_COLS)
    assert rows[1][:2] == (R.VENDOR_ID, "汇总139   3单ECMS退货包裹2026年11月03日")
    got = {r[2]: (r[3], r[5]) for r in rows[1:]}
    assert got == {1: (1, R.TAX_MAP[V19]), 4: (2, R.TAX_MAP[V7]), 6: (1, R.TAX_MAP[V19])}   # 未到 / 不可售不进
    d = list(load_workbook(diff).active.iter_rows(values_only=True))
    kinds = {(r[0], r[2]): r[9] for r in d[1:] if r[0] and r[0].startswith("LP")}
    assert kinds[("LP0002", "Hueb_446071")] == "包裹未到"
    assert kinds[("LP0001", "Doph_15994609")] == "不可售"
    assert kinds[("LP0003", "Extra_123456")] == "多出产品"
    assert any("包裹 2 / 3 到" in l and "入库（进 PO） 4" in l for l in lines)


@pytest.mark.parametrize("v,want", [(1, True), ("x", True), ("✓", True), (True, True), ("TRUE", True),
                                    (None, False), (0, False), (False, False), ("FALSE", False), ("", False)])
def test_checked(v, want):
    assert R._checked(v) is want


def test_explicit_qty_without_tick(tmp_path):
    def edit(ws, i):
        for r in i.values():
            ws.cell(r, R.C_GOT, ws.cell(r, R.C_SOLL).value)    # 不勾、直接填实收也算点过
    sheet = _fill(_make_sheet(tmp_path), tmp_path / "f.xlsx", edit)
    files, lines = R.cmd_po([sheet], str(tmp_path / "po"), FakeData())
    assert any("差异 0 行" in l for l in lines)


def test_po_all_received_no_diff(tmp_path):
    sheet = _fill(_make_sheet(tmp_path), tmp_path / "f.xlsx", _all_got)
    files, lines = R.cmd_po([sheet], str(tmp_path / "po"), FakeData())
    assert any("差异 0 行" in l for l in lines)


def _with_all(edit):
    def f(ws, i):
        _all_got(ws, i)
        edit(ws, i)
    return f


@pytest.mark.parametrize("edit,msg", [
    (lambda ws, i: None, "没勾「到」「未到」也没填实收"),
    (_with_all(lambda ws, i: ws.cell(i[("LP0001", "Doph_15994609")], R.C_NO, "x")), "都勾了"),
    (lambda ws, i: ws.cell(i[("LP0001", "Doph_15994609")], R.C_NO, 1)
     and ws.cell(i[("LP0001", "Doph_15994609")], R.C_GOT, 2), "勾了「未到」"),
    (_with_all(lambda ws, i: ws.cell(i[("LP0003", "Euc_18222103")], R.C_BAD, 3)), "不可售 3 比实收 1"),
    (_with_all(lambda ws, i: ws.cell(i[("LP0003", "Euc_18222103")], R.C_MHD, "01.2020")), "已过期"),
    (_with_all(lambda ws, i: ws.cell(i[("LP0003", "Euc_18222103")], R.C_MHD, "05.29")), "没写批次"),
    (_with_all(lambda ws, i: ws.cell(i[("LP0003", "Euc_18222103")], R.C_GOT, "x")), "不是 ≥ 0 的数字"),
])
def test_po_sheet_errors_stop(tmp_path, edit, msg):
    sheet = _fill(_make_sheet(tmp_path), tmp_path / "bad.xlsx", edit)
    with pytest.raises(R.StopError, match=msg):
        R.cmd_po([sheet], str(tmp_path / "po"), FakeData())
    assert not os.path.exists(tmp_path / "po" / "汇总139_PO导入.xlsx")


def test_po_unknown_extra_pzn_stops(tmp_path):
    def edit(ws, i):
        _fill_ok(ws, i)
        items = _rows(ws)
        ws.cell(items[("LP0003", "Euc_18222103")] + 1, R.C_PZN, "99999999")
    sheet = _fill(_make_sheet(tmp_path), tmp_path / "f.xlsx", edit)
    with pytest.raises(R.StopError, match="99999999"):
        R.cmd_po([sheet], str(tmp_path / "po"), FakeData())


def test_po_already_imported_stops(tmp_path):
    sheet = _fill(_make_sheet(tmp_path), tmp_path / "filled.xlsx", _fill_ok)
    with pytest.raises(R.StopError, match="已经建过"):
        R.cmd_po([sheet], str(tmp_path / "po"), FakeData(taken={139}))


# ---------------------------------------------------------------- lot
WH_HEAD = ["External ID", "Reference", "Contact", "Source Document", "Operations/ID", "Operations/Stock Move/ID",
           "Operations/Product", "Operations/Stock Move/Demand", "Operations/Unit of Measure", "Operations/From",
           "Operations/To"]


def _wh_export(path, moves):
    wb = Workbook()
    ws = wb.active
    ws.append(WH_HEAD)
    for k, (prod, dem) in enumerate(moves):
        head = ["__export__.pk_1", "WH/IN/9", "Vendor", "P12700"] if k == 0 else [None] * 4
        ws.append(head + [f"__export__.ml_{k}", f"__export__.mv_{k}", prod, dem, "Units",
                          "Partner Locations/Vendors", "WH/Stock"])
    wb.save(path)
    return str(path)


def test_lot_full_flow(tmp_path):
    sheet = _fill(_make_sheet(tmp_path), tmp_path / "filled.xlsx", _fill_ok)
    exp = _wh_export(tmp_path / "wh.xlsx", [("[Doph_15994609] Doppelherz", 2), ("[Euc_18222103] Eucerin", 1),
                                            ("[Extra_123456] Extra Produkt", 1)])
    files, lines = R.cmd_lot(exp, [sheet], str(tmp_path / "lot"), FakeData())
    imp = [f for f in files if "Lot_Import" in f][0]
    rows = list(load_workbook(imp).active.iter_rows(values_only=True))
    lots = {(r[3], r[4]): (r[5], r[9]) for r in rows[1:]}
    assert lots[("[Doph_15994609] Doppelherz", "AB123")] == (2, datetime(2029, 5, 31))
    assert ("[Doph_15994609] Doppelherz", "CD456") not in lots     # 这批唯一 1 件不可售 → 不入库
    assert lots[("[Euc_18222103] Eucerin", "E1001")] == (1, datetime(2028, 12, 31))
    assert any("无批次 1" in l for l in lines)          # 多出的那件没写批次 → NO LOT


def test_lot_mismatch_stops(tmp_path):
    sheet = _fill(_make_sheet(tmp_path), tmp_path / "filled.xlsx", _fill_ok)
    exp = _wh_export(tmp_path / "wh.xlsx", [("[Doph_15994609] Doppelherz", 3), ("[Euc_18222103] Eucerin", 1)])
    with pytest.raises(R.StopError, match="对不上"):
        R.cmd_lot(exp, [sheet], str(tmp_path / "lot"), FakeData())
