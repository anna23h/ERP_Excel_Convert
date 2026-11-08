"""阶段一按面单生成时间排序 + 按店限定处理单数（ISSUES: [拣货面单] N）。"""
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_excel as B  # noqa: E402
import step4_merge as s4  # noqa: E402

REF = s4.ERP_ORDER_REF


def rows(orders):
    """orders: [(Order Reference, SKU 行数)] → SKU 明细行，_key = 单号后 15 位。"""
    out = []
    for ref, k in orders:
        out += [{REF: ref, "_key": ref[-15:], "sku": f"{ref}-{j}"} for j in range(k)]
    return pd.DataFrame(out)


def times(d):
    return pd.Series({k[-15:]: pd.Timestamp(v) for k, v in d.items()})


# A 最早、D 最晚；ERP 原顺序故意打乱
SUB = rows([("VO_TOF_SCP000000000002", 1), ("VO_TOF_SCP000000000004", 2),
            ("VO_TOF_SCP000000000001", 3), ("VO_TOF_SCP000000000003", 1)])
T = times({"VO_TOF_SCP000000000001": "2026-11-01 08:00",
           "VO_TOF_SCP000000000002": "2026-11-02 08:00",
           "VO_TOF_SCP000000000003": "2026-11-03 08:00",
           "VO_TOF_SCP000000000004": "2026-11-04 08:00"})


def order_seq(df):
    return [r[-1] for r in pd.unique(df[REF])]


def test_sorted_latest_first_without_limit():
    out, total, kept, no_time = B.order_by_label_time(SUB, T)
    assert order_seq(out) == ["4", "3", "2", "1"]   # 最新在上，表底是最老的
    assert (total, kept, no_time) == (4, 4, [])
    assert len(out) == len(SUB)


def test_limit_keeps_earliest_orders():
    out, total, kept, _ = B.order_by_label_time(SUB, T, limit=2)
    assert order_seq(out) == ["2", "1"]
    assert (total, kept) == (4, 2)


def test_limit_counts_orders_not_lines_and_keeps_multiline_order_whole():
    out, _, kept, _ = B.order_by_label_time(SUB, T, limit=1)
    assert kept == 1 and len(out) == 3              # 订单 1 有 3 行，整单保留
    assert set(out[REF]) == {"VO_TOF_SCP000000000001"}


def test_multiline_orders_stay_contiguous():
    out, *_ = B.order_by_label_time(SUB, T)
    refs = out[REF].tolist()
    for r in set(refs):
        idx = [i for i, x in enumerate(refs) if x == r]
        assert idx == list(range(idx[0], idx[0] + len(idx)))


def test_limit_above_total_keeps_all():
    out, total, kept, _ = B.order_by_label_time(SUB, T, limit=99)
    assert (total, kept) == (4, 4)


def test_limit_zero_keeps_nothing():
    out, total, kept, _ = B.order_by_label_time(SUB, T, limit=0)
    assert out.empty and (total, kept) == (4, 0)


def test_missing_time_goes_top_and_is_picked_last():
    t = T.drop("SCP000000000003")
    out, _, _, no_time = B.order_by_label_time(SUB, t)
    assert order_seq(out) == ["3", "4", "2", "1"]
    assert no_time == ["VO_TOF_SCP000000000003"]
    out, *_ = B.order_by_label_time(SUB, t, limit=3)
    assert "VO_TOF_SCP000000000003" not in set(out[REF])


def test_same_time_keeps_original_order():
    t = times({r: "2026-11-01 08:00" for r in SUB[REF].unique()})
    out, *_ = B.order_by_label_time(SUB, t)
    assert order_seq(out) == ["2", "4", "1", "3"]
