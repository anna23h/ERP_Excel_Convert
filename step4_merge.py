#!/usr/bin/env python3
"""步骤4 核心原型：合并 ERP 导出 + 天猫全量导出，按履约单状态分流。
第一版只做合并/分类/体检报告，不生成 Excel。用于和用户核对分流逻辑。

用法:
    python3 step4_merge.py <erp.xlsx|csv> <tmall.xlsx>
默认:
    python3 step4_merge.py   # 用 raw_data 下的默认样例
"""
import sys, warnings
import pandas as pd
warnings.filterwarnings("ignore")

# 列名常量（按列名取，不依赖列位置）
ERP_ORDER_REF   = "Order Reference"
ERP_INTERNAL    = "Order Lines/Product/Internal Reference"
ERP_PICKING     = "Order Lines/Product/Picking Name"
ERP_QTY         = "Order Lines/Quantity"
ERP_DELIVERY    = "VO Delivery Type"
ERP_TRACKING    = "VO Tracking No"
ERP_BARCODE     = "Order Lines/Product/Barcode"
ERP_ONHAND      = "Order Lines/Product/Quantity On Hand"

TM_KEY          = "系统履约单号"
TM_STATUS       = "履约单状态"
TM_LABEL_STATUS = "面单申请状态"

CANCEL_STATUSES = {"履约取消", "平台申请取消"}


# 订单级字段：多品订单的续行在 Odoo 导出里留空，需向下填充
ERP_ORDER_LEVEL = [ERP_ORDER_REF, ERP_TRACKING, ERP_DELIVERY,
                   "Order Date", "Terms and conditions"]


def load_erp(path):
    df = pd.read_csv(path) if path.endswith(".csv") else pd.read_excel(path)
    # 多品订单续行：订单级字段向下填充（续行只有商品级字段）
    df[ERP_ORDER_LEVEL] = df[ERP_ORDER_LEVEL].ffill()
    df["_key"] = df[ERP_ORDER_REF].astype(str).str[-15:]
    return df


def load_tmall(path):
    df = pd.read_excel(path, sheet_name="file",
                       usecols=[TM_KEY, TM_STATUS, TM_LABEL_STATUS])
    df[TM_KEY] = df[TM_KEY].astype(str)
    # 天猫全量含历史多日数据，同一单号可能重复，保留最后一条
    df = df.drop_duplicates(subset=TM_KEY, keep="last")
    return df


def merge(erp, tmall):
    return erp.merge(tmall, left_on="_key", right_on=TM_KEY, how="left")


def report(erp, tmall, merged):
    print("=" * 60)
    print("体检报告")
    print("=" * 60)
    print(f"ERP 行数: {len(erp)}  (订单行，可能一单多行)")
    print(f"ERP 唯一订单(_key): {erp['_key'].nunique()}")
    print(f"天猫去重后行数: {len(tmall)}")

    unmatched = merged[merged[TM_KEY].isna()]
    print(f"\n[连接键未命中] {len(unmatched)} 行 "
          f"({unmatched['_key'].nunique()} 个单号) —— 这些在天猫导出里找不到")
    for k in unmatched["_key"].unique()[:10]:
        print(f"    {k}")

    print(f"\n[履约单状态分布]")
    for s, c in merged[TM_STATUS].value_counts(dropna=False).items():
        flag = "  <- 取消" if s in CANCEL_STATUSES else ""
        print(f"    {str(s):28} {c}{flag}")

    print(f"\n[面单申请状态分布]")
    for s, c in merged[TM_LABEL_STATUS].value_counts(dropna=False).items():
        print(f"    {str(s):28} {c}")

    empty_delivery = merged[merged[ERP_DELIVERY].isna()]
    print(f"\n[VO Delivery Type 为空] {len(empty_delivery)} 行")

    # SKU 多 Barcode/Picking 异常
    g = merged.groupby(ERP_INTERNAL)
    multi = (g[ERP_BARCODE].nunique() > 1) | (g[ERP_PICKING].nunique() > 1)
    print(f"[SKU 对应多个 Barcode/Picking] {int(multi.sum())} 个")


def classify(merged):
    is_cancel = merged[TM_STATUS].isin(CANCEL_STATUSES)
    cancel = merged[is_cancel]
    facesheet = merged[~is_cancel]  # 「待」全集（出/无 待仓库反馈）
    return facesheet, cancel


def main():
    args = sys.argv[1:]
    if len(args) >= 2:
        erp_path, tm_path = args[0], args[1]
    else:
        erp_path = "raw_data/测试0611erp导出.xlsx"
        tm_path = "raw_data/天猫测试.xlsx"
        print(f"(未传参，使用默认样例)\n  ERP:   {erp_path}\n  天猫:  {tm_path}\n")

    erp = load_erp(erp_path)
    tmall = load_tmall(tm_path)
    merged = merge(erp, tmall)
    report(erp, tmall, merged)

    facesheet, cancel = classify(merged)
    print("\n" + "=" * 60)
    print("分流结果")
    print("=" * 60)
    print(f"面单(待全集): {len(facesheet)} 行 / {facesheet['_key'].nunique()} 单")
    print(f"取消单:       {len(cancel)} 行 / {cancel['_key'].nunique()} 单")


if __name__ == "__main__":
    main()
