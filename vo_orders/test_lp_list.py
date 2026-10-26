"""阶段一 LP 号清单测试（ISSUES: [拣货面单] M）：顺序 / 每单一次 / 每 100 个一段 / 文件编码。"""
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_excel as B  # noqa: E402


def lp(n):
    return f"LP{n:014d}"


def facesheet(orders):
    """orders: [(Order Reference, Tracking, SKU 行数)] → SKU 明细行（多品订单连续多行，同运单号）。"""
    rows = []
    for i, (ref, t, k) in enumerate(orders, 1):
        rows += [{"序号": i, "Order Reference": ref, "VO Tracking No": t}] * k
    return pd.DataFrame(rows)


def test_one_per_order_in_facesheet_order():
    fs = facesheet([("VO_3", lp(3), 1), ("VO_1", lp(1), 3), ("VO_2", lp(2), 2)])
    text, n = B.build_lp_text(fs)
    assert n == 3
    assert text == f"{lp(3)}\n{lp(1)}\n{lp(2)}\n"


def test_chunks_of_100_blank_line_between():
    fs = facesheet([(f"VO_{i}", lp(i), 1 + i % 2) for i in range(1, 251)])
    text, n = B.build_lp_text(fs)
    assert n == 250
    parts = text.rstrip("\n").split("\n\n")
    assert [len(p.split("\n")) for p in parts] == [100, 100, 50]
    assert parts[0].split("\n")[0] == lp(1) and parts[2].split("\n")[-1] == lp(250)
    assert "\n\n\n" not in text and text.endswith(lp(250) + "\n")


def test_exactly_100_has_no_trailing_blank_segment():
    text, _ = B.build_lp_text(facesheet([(f"VO_{i}", lp(i), 1) for i in range(100)]))
    assert "\n\n" not in text


def test_blank_tracking_skipped_and_empty():
    fs = facesheet([("VO_1", lp(1), 1), ("VO_2", None, 1), ("VO_3", "  ", 1)])
    assert B.build_lp_text(fs) == (lp(1) + "\n", 1)


def test_write_lplist_name_and_encoding(tmp_path):
    fs = facesheet([("VO_1", lp(1), 2), ("VO_2", lp(2), 1)])
    path, n = B._write_lplist(fs, "VO", str(tmp_path))
    assert n == 2
    assert os.path.basename(path).endswith("VO2单 LP清单.txt")
    raw = open(path, "rb").read()
    assert raw == f"{lp(1)}\n{lp(2)}\n".encode()    # 无 BOM、无 \r
