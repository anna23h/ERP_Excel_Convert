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
# 不该出现在今日发货名单里的状态(已出库/已收货/系统取消)——若混入说明名单过期，防重复发货
SHIPPED_DONE_STATUSES = {"已发货", "履约完成(已收货)", "发货后取消(系统取消)"}
WU_TAG = "无运单"            # 无运单标记(加在 Terms and conditions 前)


def last15(series):
    return series.astype(str).str[-15:]


ID_NAMES = ("external id", "id", "外部id", "外部 id")


def find_id_col(df):
    """找出 External ID 列(兼容 External ID / ID / id / 外部ID)。找不到返回 None。"""
    for c in df.columns:
        if str(c).strip().lower() in ID_NAMES:
            return c
    return None


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


def _read_tmall_sheet(path, usecols=None):
    """天猫导出统一读法：优先 sheet 'file'，否则首个 sheet。"""
    xl = pd.ExcelFile(path)
    sheet = "file" if "file" in xl.sheet_names else 0
    return pd.read_excel(xl, sheet_name=sheet, usecols=usecols)


def load_done_keys(path):
    """面单已完成名单(天猫后台筛选导出) → 15 位单号集合 = 今天有运单可发货的订单。"""
    if not path:
        return set()
    df = _read_tmall_sheet(path)
    col = TM_KEY if TM_KEY in df.columns else df.columns[0]
    return set(last15(df[col].dropna()))


def load_cancel_keys(path):
    """完整天猫导出 → 取消单号集合(履约取消/平台申请取消)。"""
    if not path:
        return set()
    df = _read_tmall_sheet(path, usecols=[TM_KEY, TM_STATUS])
    df = df[df[TM_STATUS].isin(CANCEL_STATUSES)]
    return set(last15(df[TM_KEY]))


def load_status_map(path):
    """完整天猫导出 → {15位单号: 履约单状态} Series(去重)。无文件则返回空 Series。"""
    if not path:
        return pd.Series(dtype=object)
    df = _read_tmall_sheet(path, usecols=[TM_KEY, TM_STATUS])
    df = df.assign(_k=last15(df[TM_KEY])).drop_duplicates("_k")
    return df.set_index("_k")[TM_STATUS]


def classify4(erp, done_keys, cancel_keys):
    """在 ERP(行级，含 _key + Terms and conditions)上打 `_cat` 标签并分流。
    优先级：取消 > 发货/已补运单 > 无运单。
    - 取消        : _key ∈ cancel_keys
    - 发货        : 在面单已完成名单(done_keys)里
    - 已补运单     : 发货 且 ERP Terms 已含「无运单」(昨日无运单今日补出) — 仍发货
    - 无运单       : 其余(既不在名单也未取消) — 从拣货/面单剔除
    返回打了 `_cat` 列的 ERP 副本。`_ship` = _cat ∈ {发货, 已补运单}。"""
    df = erp.copy()
    has_wu = df["Terms and conditions"].astype(str).str.contains(WU_TAG)
    k = df["_key"]
    is_cancel = k.isin(cancel_keys)
    is_done = k.isin(done_keys) & ~is_cancel
    cat = pd.Series("无运单", index=df.index)
    cat[is_done] = "发货"
    cat[is_done & has_wu] = "已补运单"
    cat[is_cancel] = "取消"
    df["_cat"] = cat
    df["_ship"] = df["_cat"].isin(["发货", "已补运单"])
    return df


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
