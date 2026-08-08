"""po_frequency 回归测试：合成锯齿导出，验证 ffill / SKU 归一 / 频次去重 / vendor 过滤。

跑：python3 -m pytest po_frequency/test_po_frequency.py    （或 python3 po_frequency/test_po_frequency.py）
"""
import importlib.util
import os
import sys
import tempfile

import pandas as pd

# 目录不是 package（与 po_reconcile 等一致），按文件路径显式加载模块
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # 让 common/ 可导入
_spec = importlib.util.spec_from_file_location(
    "po_frequency_mod", os.path.join(os.path.dirname(os.path.abspath(__file__)), "po_frequency.py"))
_mod = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(_mod)
load, build = _mod.load, _mod.build

# 锯齿导出：订单头字段只在每单首行，其余留空（用 None 模拟 Odoo 导出）
RAW = [
    # PO,     date,                  vendor,    sku,               name,                         qty, price
    ("P001", "2025-01-10 09:00:00", "PHOENIX Pharmahandel", "Sinupret_001", "[Sinupret_001] Sinupret Saft 100ml", "100", "5.80"),
    (None,   None,                   None,      "Sinupret_001",    "[Sinupret_001] Sinupret Saft 100ml", "6",  "5.80"),   # 同单同品第二行
    ("P002", "2025-02-10 10:00:00", "PHOENIX Pharmahandel", "Sinupret_001", "[Sinupret_001] Sinupret Saft 100ml", "50",  "5.79"),
    ("P003", "2025-03-10 11:00:00", "PHOENIX Pharmahandel", "Sinupret_001x2", "[Sinupret_001x2] Sinupret Saft 100ml x2", "20", "5.79"),  # x2 归一到 Sinupret_001
    ("P004", "2025-01-15 08:00:00", "Other Vendor GmbH", "Aspirin_002", "[Aspirin_002] Aspirin 20st", "30", "3.10"),
]
COLS = ["Order Reference", "Confirmation Date", "Vendor",
        "Order Lines/Product/Internal Reference", "Order Lines/Product/Display Name",
        "Order Lines/Total Quantity", "Order Lines/Unit Price"]


def _write_tmp():
    df = pd.DataFrame(RAW, columns=COLS)
    fd, path = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
    df.to_excel(path, index=False)
    return path


def _row(summary, sku):
    return summary[summary["Internal Reference"] == sku].iloc[0]


def test_frequency_and_normalization():
    path = _write_tmp()
    try:
        po = load(path)
        summary, details = build(po)
    finally:
        os.remove(path)

    # SKU 归一：Sinupret_001x2 并入 Sinupret_001 → 只有 Sinupret_001 与 Aspirin_002 两个产品
    assert set(summary["Internal Reference"]) == {"Sinupret_001", "Aspirin_002"}

    r = _row(summary, "Sinupret_001")
    assert r["Purchase Count (Frequency)"] == 3           # P001/P002/P003 三单（同单第二行不另计）
    assert r["Total Qty"] == 176                          # 100+6+50+20
    assert r["Max per Purchase"] == 106                   # P001 两行合计
    assert r["Min per Purchase"] == 20
    assert round(r["Avg per Purchase"], 1) == 58.7        # 176/3
    assert r["First Purchase"] == "2025-01-10"
    assert r["Last Purchase"] == "2025-03-10"
    assert r["Span (days)"] == 59
    assert r["Avg Interval (days)"] == 30                 # round(59/2)

    # 频次降序：Sinupret(3) 在 Aspirin(1) 前
    assert list(summary["Internal Reference"]) == ["Sinupret_001", "Aspirin_002"]

    # 明细：ffill 生效——第二行(原空)拿到 P001；Sinupret 共 4 行
    det_sin = details[details["Internal Reference"] == "Sinupret_001"]
    assert len(det_sin) == 4
    assert (det_sin["PO Number"] == "P001").sum() == 2
    assert det_sin.iloc[0]["Order Date"] == "2025-01-10"


def test_vendor_filter():
    path = _write_tmp()
    try:
        po = load(path, vendor="phoenix")               # 忽略大小写、子串
        summary, _ = build(po)
    finally:
        os.remove(path)
    # Aspirin 属 Other Vendor，被过滤掉
    assert set(summary["Internal Reference"]) == {"Sinupret_001"}


def test_missing_column_raises():
    df = pd.DataFrame([["P1", "x"]], columns=["Order Reference", "Vendor"])
    fd, path = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
    df.to_excel(path, index=False)
    try:
        raised = False
        try:
            load(path)
        except ValueError as e:
            raised = True
            assert "缺列" in str(e)
        assert raised
    finally:
        os.remove(path)


if __name__ == "__main__":
    test_frequency_and_normalization()
    test_vendor_filter()
    test_missing_column_raises()
    print("all tests passed")
