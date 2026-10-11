"""lot_entry 测试：解析小件 + 「导出 → 录入表 → 操作员填写 → 导入表」全程。

全程用代码造的假导出（*.xlsx 一律 gitignore，真实导出不进仓库），
列名用英文标签版，与 2026-10-11 用户实际导出 `stock.picking (9).xlsx` 同构。
"""
import os
import sys
from datetime import date, datetime

import pytest
from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lot_entry as L  # noqa: E402


# ---------------------------------------------------------------- 解析小件
@pytest.mark.parametrize("raw,want", [
    ("11.2027", date(2027, 11, 30)),
    ("2.2028", date(2028, 2, 29)),          # 闰年月底
    ("04/2029", date(2029, 4, 30)),
    ("30.01.2028", date(2028, 1, 30)),
    ("2027-10", date(2027, 10, 31)),
    ("2027-10-05", date(2027, 10, 5)),
    (" 12 . 2027 ", date(2027, 12, 31)),
    (datetime(2028, 3, 1), date(2028, 3, 1)),
])
def test_parse_mhd_ok(raw, want):
    assert L.parse_mhd(raw) == (want, None)


@pytest.mark.parametrize("raw", ["05.28", "13.2027", "31.02.2028", "2027", "11,2027x", 11.2027])
def test_parse_mhd_rejects(raw):
    d, err = L.parse_mhd(raw)
    assert d is None and err


def test_parse_mhd_empty():
    assert L.parse_mhd(None) == (None, None)
    assert L.parse_mhd("  ") == (None, None)


def test_name_mhd_months():
    assert L.name_mhd_months("Mykoderm Mund-Gel 20g(MHD 2027/10)") == {(2027, 10)}
    assert L.name_mhd_months("Bepanthen Augentropfen MHD2029/04") == {(2029, 4)}
    assert L.name_mhd_months("Otriven 10ml(MHD05.28)") == {(2028, 5)}
    assert L.name_mhd_months("Crataegutt 50StMHD28.07") == {(2028, 7)}
    assert L.name_mhd_months("Elmex Kariesschutz 400ml") == set()


def test_split_product_flattens_newlines():
    assert L.split_product("[DOPH_12454290] Doppelherz \nFlohsamenschalen \n250 G") == \
        ("DOPH_12454290", "Doppelherz Flohsamenschalen 250 G")


# ---------------------------------------------------------------- 全程
HEADER = ["Operations/ID", "Operations/Stock Move/ID", "Operations/Product",
          "Operations/Lot/Serial Number Name", "Operations/Done", "Operations/Unit of Measure",
          "Operations/From", "Operations/To", "Operations/Expiration Date", "ID",
          "External ID", "Contact", "Operations/Stock Move/Demand", "Source Document",
          "Reference"]


def _export(path):
    rows = [
        # 单 1：一个产品
        ("ml1", "mv1", "[Rosu_1] Rosuvastatin 100 St", 1000, "pk1", "Hexal", "P1", "WH/IN/1"),
        # 单 2：多个产品
        ("ml2", "mv2", "[Dopp_2] Doppelherz 480St", 360, "pk2", "Queisser", "P2", "WH/IN/2"),
        ("ml3", "mv3", "[Wolz_3] Cardio Care", 2, None, None, None, None),
        ("ml4", "mv4", "[Derm_4] Badesalz 5Kg", 2, None, None, None, None),
        ("ml5", "mv5", "[Flo_5] Flohsamen MHD 2028/05", 1, None, None, None, None),
        ("ml6", "mv6", "[Ome_6] Omega", 2, None, None, None, None),
        ("ml7", "mv7", "[Hylo_7] Hylo Dual", 9, None, None, None, None),
        ("ml8", "mv8", "[Elm_8] Elmex", 2, None, None, None, None),
        ("ml9", "mv9", "[Untouched_9] Nobody fills me", 1, None, None, None, None),
        # 单 3：没人录 → 整单不进导入表
        ("ml10", "mv10", "[Empty_10] Empty", 5, "pk3", "X", "P3", "WH/IN/3"),
    ]
    wb = Workbook()
    ws = wb.active
    ws.append(HEADER)
    for ml, mv, prod, dem, pk, partner, origin, ref in rows:
        ws.append([ml, mv, prod, None, 0, "Units", "Partner Locations/Vendors", "WH/Stock",
                   None, "123" if pk else None, pk, partner, dem, origin, ref])
    wb.save(path)


def _fill(path):
    """模拟操作员：按产品编码找到产品行，往下面的录入行填。"""
    wb = load_workbook(path)
    ws = wb["Erfassung"]
    def locate():
        return {ws.cell(r, L.C_CODE).value: r for r in range(2, ws.max_row + 1)
                if ws.cell(r, L.C_TECH0).value == "P"}
    prod = locate()

    def put(code, k, lot=None, mhd=None, qty=None):
        r = locate()[code] + 1 + k
        for col, v in ((L.C_LOT, lot), (L.C_MHD, mhd), (L.C_QTY, qty)):
            if v is not None:
                ws.cell(r, col, v)

    put("Rosu_1", 0, "25P039A", "11.2027")                      # 单批次不写数量 → 取应收
    put("Dopp_2", 0, "AB1234", "30.01.2028", 300)
    put("Dopp_2", 1, "4X3X-A", "04.2029", 60)
    put("Wolz_3", 0, "WZ0001", "12.2027", 1)                    # 合计 1 ≠ 2 → 警告
    ws.cell(prod["Derm_4"], L.C_NOLOT, "x")
    put("Derm_4", 0, mhd="06.2028")                             # 无批次只写 MHD → 警告、跳过
    put("Flo_5", 0, "FL5555", "07.2028")                        # 与品名 MHD 2028/05 不符 → 警告
    put("Ome_6", 0, "OM1111", "05.28", 1)                       # 两位年 → 错误
    put("Ome_6", 1, "OM2222", "06.2028", 1)
    # Hylo 三批：操作员在第 2 行录入行下面插一行
    ws.insert_rows(prod["Hylo_7"] + 3)
    prod = locate()
    put("Hylo_7", 0, "HY0001", "01.2028", 3)
    put("Hylo_7", 1, "HY0002", "02.2028", 3)
    put("Hylo_7", 2, "HY0003", "03.2028", 3)
    put("Elm_8", 0, "EL 12", "01.2029")                         # 批次带空格 → 错误
    wb.save(path)


def test_end_to_end(tmp_path):
    exp = tmp_path / "export.xlsx"
    _export(exp)
    kind, files, _ = L.run(str(exp), str(tmp_path))
    assert kind == "entry"
    entry = files[0]

    ws = load_workbook(entry)["Erfassung"]
    assert ws.cell(2, 1).value.startswith("WH/IN/1")
    assert ws.cell(3, L.C_CODE).value == "Rosu_1" and ws.cell(3, L.C_SOLL).value == 1000
    assert ws.cell(4, L.C_LOT).number_format == "@"
    assert ws.cell(4, L.C_LOT).protection.locked is False
    assert ws.protection.sheet is True

    _fill(entry)
    kind, files, summary = L.run(entry, str(tmp_path))
    assert kind == "import"
    imp, chk = files

    rows = list(load_workbook(imp).active.iter_rows(values_only=True))
    assert list(rows[0]) == L.IMPORT_HEADER
    body = rows[1:]
    by_lot = {r[4]: r for r in body if r[4]}
    pickings = [r[0] for r in body if r[0]]
    assert pickings == ["pk1", "pk2"]                            # 单 3 没人录 → 不出现

    assert by_lot["25P039A"][1] == "ml1" and by_lot["25P039A"][5] == 1000
    assert by_lot["25P039A"][9] == datetime(2027, 11, 30)
    assert by_lot["AB1234"][1] == "ml2" and by_lot["AB1234"][5] == 300
    assert by_lot["4X3X-A"][1] is None and by_lot["4X3X-A"][2] == "mv2"
    assert by_lot["4X3X-A"][3] == "[Dopp_2] Doppelherz 480St"
    assert by_lot["4X3X-A"][9] == datetime(2029, 4, 30)
    assert [by_lot[x][1] for x in ("HY0001", "HY0002", "HY0003")] == ["ml7", None, None]
    assert "WZ0001" in by_lot and "FL5555" in by_lot             # 只警告的照写
    assert not ({"OM1111", "OM2222", "EL 12"} & set(by_lot))    # 标红的整产品不写

    # 没写批次的已有行原样带上（防 Odoo 把未列出的行替换掉）
    lines = [r[1] for r in body if r[1]]
    for ml in ("ml4", "ml6", "ml8", "ml9"):
        assert ml in lines
    assert all(r[4] is None and r[5] == 0 for r in body if r[1] in ("ml4", "ml6", "ml8", "ml9"))
    assert "ml10" not in lines

    check = list(load_workbook(chk).active.iter_rows(values_only=True))[1:]
    msgs = {(c[0], c[1]) for c in check}
    assert ("WARN", "Wolz_3") in msgs
    assert ("WARN", "Derm_4") in msgs
    assert ("WARN", "Flo_5") in msgs
    assert ("ERROR", "Ome_6") in msgs and ("SKIP", "Ome_6") in msgs
    assert ("ERROR", "Elm_8") in msgs
    assert ("EMPTY", "Untouched_9") in msgs


def test_missing_column_is_reported(tmp_path):
    wb = Workbook()
    wb.active.append(["Operations/ID", "Operations/Product"])
    p = tmp_path / "bad.xlsx"
    wb.save(p)
    with pytest.raises(L.InputError, match="move_id"):
        L.run(str(p), str(tmp_path))
