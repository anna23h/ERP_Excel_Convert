# -*- coding: utf-8 -*-
"""本地套件对照表（ISSUES: [拣货面单] S）。全部合成数据，不连 Odoo。"""
import os
import sys
from datetime import date

import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_excel as B  # noqa: E402
import safety_dynamic as SD  # noqa: E402

TODAY = date(2026, 11, 24)


def _write(tmp_path, rows, header=SD.KITS_COLS):
    p = tmp_path / SD.KITS_FILE
    lines = [",".join(header)] + [",".join(map(str, r)) for r in rows]
    p.write_text("\n".join(lines) + "\n", encoding="utf-8-sig")
    return str(p)


def test_load_filters_to_batch_and_reports_date(tmp_path):
    p = _write(tmp_path, [("Aaa_1x2", "Aaa_1", 2, "2026-11-20"),
                          ("Bbb_2_VO", "Bbb_2", 1, "2026-11-20"),
                          ("Ccc_3x3", "Ccc_3", 3, "2026-11-20")])
    kits, info, exported, known = SD.load_kits(p, ["Aaa_1x2", "Bbb_2_VO", "Zzz_9"], today=TODAY)
    assert kits == {"Aaa_1x2": ("Aaa_1", 2), "Bbb_2_VO": ("Bbb_2", 1)}
    assert exported == date(2026, 11, 20)
    assert "Ccc_3x3" in known and "Ccc_3" in known
    assert "3 条" in info and "2026-11-20" in info and "⚠" not in info


def test_stale_or_undated_file_warns(tmp_path):
    p = _write(tmp_path, [("Aaa_1x2", "Aaa_1", 2, "2026-09-01")])
    info = SD.load_kits(p, today=TODAY)[1]
    assert "⚠" in info and "84 天" in info
    p = _write(tmp_path, [("Aaa_1x2", "Aaa_1", 2, "")])
    _, info, exported, _ = SD.load_kits(p, today=TODAY)
    assert exported is None and "没有导出日期" in info


def test_bad_file_raises(tmp_path):
    p = _write(tmp_path, [("Aaa_1x2", "Aaa_1")], header=("套件SKU", "基础品SKU"))
    with pytest.raises(ValueError, match="缺列"):
        SD.load_kits(p)
    p = _write(tmp_path, [("Aaa_1x2", "Aaa_1", "zwei", "2026-11-20")])
    with pytest.raises(ValueError, match="系数"):
        SD.load_kits(p)


def test_find_kits_file_first_existing_dir(tmp_path):
    a, b = tmp_path / "a", tmp_path / "b"
    a.mkdir(); b.mkdir()
    assert SD.find_kits_file([str(a), str(b)]) is None
    (b / SD.KITS_FILE).write_text("x", encoding="utf-8")
    assert SD.find_kits_file([str(a), str(b)]) == str(b / SD.KITS_FILE)


def test_kit_like_missing():
    # Eee_5_VO 名字像套件，但在表里是基础品 → 不报
    assert SD.kit_like_missing(["Aaa_1x2", "Bbb_2_GW", "Ccc_3", "Ddd_4x2", "Eee_5_VO"],
                               {"Aaa_1x2", "Aaa_1", "Eee_5_VO"}) == ["Bbb_2_GW", "Ddd_4x2"]


def _no_odoo(_skus):
    raise AssertionError("有本地对照表时不应连 Odoo")


def test_local_file_wins_and_odoo_is_not_called(tmp_path):
    p = _write(tmp_path, [("Aaa_1x2", "Aaa_1", 2, date.today().isoformat())])
    kits, log = B.load_kit_map(["Aaa_1x2", "Aaa_1", "Bbb_2_VO"], kits_file=p, pull=_no_odoo)
    assert kits == {"Aaa_1x2": ("Aaa_1", 2)}
    assert "本地" in log[0]
    assert any("Bbb_2_VO" in ln for ln in log)


def test_bad_file_falls_back_to_odoo(tmp_path):
    p = _write(tmp_path, [("Aaa_1x2", "Aaa_1")], header=("套件SKU", "基础品SKU"))
    kits, log = B.load_kit_map(["Aaa_1x2"], kits_file=p,
                               pull=lambda s: ({"Aaa_1x2": ("Aaa_1", 2)}, "ok"))
    assert kits == {"Aaa_1x2": ("Aaa_1", 2)}
    assert "读不了" in log[0] and "Odoo" in log[1]


def test_no_file_and_no_odoo_degrades_to_raw_sku():
    def boom(_):
        raise RuntimeError("连不上")
    kits, log = B.load_kit_map(["Aaa_1x2"], kits_file="", pull=boom)
    assert kits == {} and "⚠" in log[0] and SD.KITS_FILE in log[0]
