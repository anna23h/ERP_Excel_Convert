#!/usr/bin/env python3
"""销售分析 + 安全库存提醒 + Safety Stock 回写 ERP。

三份输入合一：
    销售数据   Odoo「Sales Analysis」透视导出（[SKU] 名称 / Untaxed Total / # of Lines / Qty Ordered）
    产品主数据 product.product 导出（ID 唯一映射码 / Internal Reference / Name / 可选 Quantity On Hand）
    安全库存表 运营手工维护（可选，但**优先于脚本推算**）

四份产出：
    销量排名.xlsx        全 SKU 按销量排序，含累计占比（ABC 分析）
    补货提醒.xlsx        在手库存 < 安全库存 的 SKU，按缺口降序
    安全库存回写表.xlsx   ID + Safety Stock → 导入 ERP 产品主数据
    安全库存候选值.xlsx   脚本推算的值，**待运营审阅**，不进回写表

为什么推算值不自动写回：推算基于**期间均值**，会低估上升期新品、高估下滑品。
未经人审的阈值不许进 ERP 主数据（2026-08-01 用户拍板）。

用法:
    python3 sales_insight/sales_insight.py <销售数据.xlsx> \
        --products <product.product.xlsx> --safety <安全库存.xlsx> --weeks 30
"""
import argparse
import os
import re
import sys
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # 让 common/ 可导入
from common.xlsx import write_simple  # noqa: E402

# product.product 导出的列名
P_ID = "ID"                    # __export__.product_product_* —— 回写 ERP 的唯一映射码
P_SKU = "Internal Reference"
P_NAME = "Name"
P_SHOP = "VO Shop Name"
# 在手库存列名在不同导出模版下可能不同，按顺序试
P_ONHAND_CANDIDATES = ["Quantity On Hand", "Qty On Hand", "On Hand", "Quantity available"]


def read_sales(path):
    """读 Odoo「Sales Analysis」透视导出。

    结构：前 3 行多级表头、第 4 行 Total 合计行、第 5 行起明细。
    明细的商品列形如 `    [SKU] 商品名`——**前面有缩进空格**，正则必须容忍
    （`^\\[` 匹配不到，会得到 0 条，别问我怎么知道的）。
    """
    raw = pd.read_excel(path, header=None)
    if raw.shape[1] < 4:
        raise SystemExit(f"销售导出应有 4 列（商品/销售额/订单行数/销量），实得 {raw.shape[1]} 列")

    total = raw.iloc[3]
    df = raw.iloc[4:].copy()
    df.columns = ["商品", "销售额", "下单次数", "销量"]
    df["SKU"] = df["商品"].astype(str).str.extract(r"^\s*\[([^\]]+)\]")[0]
    bad = df["SKU"].isna().sum()
    if bad:
        print(f"  ⚠ {bad} 行解析不出 SKU，已跳过")
        df = df[df["SKU"].notna()]
    for c in ("销售额", "下单次数", "销量"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)

    # 与导出自带的 Total 行对拍——解析错了这里会立刻暴露
    checks = [("销售额", 1), ("下单次数", 2), ("销量", 3)]
    for name, col in checks:
        want = pd.to_numeric(total[col], errors="coerce")
        got = df[name].sum()
        if pd.notna(want) and abs(want - got) > 0.01:
            print(f"  ⚠ {name} 明细求和 {got:.2f} ≠ 导出 Total {want:.2f}")
    return df.reset_index(drop=True)


def read_products(path):
    """读 product.product 导出。Internal Reference 前面带 `\\t`（Odoo 强制文本），必须 strip。"""
    df = pd.read_excel(path)
    for c in (P_ID, P_SKU):
        if c not in df.columns:
            raise SystemExit(f"产品主数据缺列 `{c}`（现有: {list(df.columns)}）")
    df["SKU"] = df[P_SKU].astype(str).str.strip()
    df = df[df["SKU"].notna() & (df["SKU"] != "") & (df["SKU"] != "nan")]
    df = df.drop_duplicates(subset="SKU", keep="first")

    onhand = next((c for c in P_ONHAND_CANDIDATES if c in df.columns), None)
    out = pd.DataFrame({
        "SKU": df["SKU"],
        "ERP_ID": df[P_ID],
        "商品名称": df[P_NAME] if P_NAME in df else "",
        "店铺": df[P_SHOP] if P_SHOP in df else "",
        "在手库存": pd.to_numeric(df[onhand], errors="coerce") if onhand else pd.NA,
    })
    return out.reset_index(drop=True), onhand


def read_safety(path):
    """读运营维护的安全库存表。

    列名带日期前缀（`7.29在手库存`），故按「包含」匹配而非写死——
    下个月那份会叫 `8.31在手库存`。
    """
    df = pd.read_excel(path)
    sku_col = next((c for c in df.columns if "SKU" in str(c).upper()), df.columns[0])
    safe_col = next((c for c in df.columns if str(c).strip() == "安全库存"), None)
    if safe_col is None:
        raise SystemExit(f"安全库存表里找不到「安全库存」列（现有: {list(df.columns)}）")
    onhand_col = next((c for c in df.columns if "在手库存" in str(c)), None)
    note_col = next((c for c in df.columns if "备注" in str(c)), None)
    weeks = [c for c in df.columns if re.fullmatch(r"W\d+", str(c).strip())]

    out = pd.DataFrame({
        "SKU": df[sku_col].astype(str).str.strip(),
        "安全库存_人工": pd.to_numeric(df[safe_col], errors="coerce"),
        "在手库存_运营表": pd.to_numeric(df[onhand_col], errors="coerce") if onhand_col else pd.NA,
        "备注": df[note_col] if note_col else "",
    })
    if weeks:
        recent = weeks[-4:]                     # 最近 4 周，比全期均值更贴近当下
        out["近4周均销"] = df[recent].apply(pd.to_numeric, errors="coerce").mean(axis=1).round(1)
    return out[out["SKU"].notna()].reset_index(drop=True), len(weeks), onhand_col


def build(sales, prods, safety, weeks, cover_weeks):
    """三份合一 → 一张宽表。其余产出都是它的切片。"""
    df = sales.merge(prods, on="SKU", how="left")
    if safety is not None:
        df = df.merge(safety, on="SKU", how="left")
    else:
        for c in ("安全库存_人工", "在手库存_运营表", "备注", "近4周均销"):
            df[c] = pd.NA

    df = df.sort_values("销量", ascending=False).reset_index(drop=True)
    df["销量排名"] = range(1, len(df) + 1)
    tot = df["销量"].sum()
    df["累计占比"] = (df["销量"].cumsum() / tot).round(4) if tot else 0.0
    df["每单件数"] = (df["销量"] / df["下单次数"].replace(0, pd.NA)).round(2)

    # 推算：期间均值 → 周均 → ×覆盖周数。只作候选值，不自动写回。
    df["周均销量"] = (df["销量"] / weeks).round(1)
    df["安全库存_推算"] = (df["周均销量"] * cover_weeks).round().astype("Int64")

    df["安全库存"] = df["安全库存_人工"].fillna(df["安全库存_推算"]).astype("Int64")
    df["安全库存来源"] = df["安全库存_人工"].notna().map({True: "运营人工", False: "脚本推算"})

    # 在手库存：ERP 的 On Hand 优先，缺了才退回运营表那份日期快照
    df["在手库存"] = df["在手库存"].fillna(df["在手库存_运营表"])
    df["库存来源"] = pd.Series(["ERP"] * len(df)).where(
        prods.set_index("SKU")["在手库存"].reindex(df["SKU"]).notna().values, "运营表")
    df.loc[df["在手库存"].isna(), "库存来源"] = "缺"

    df["缺口"] = (df["安全库存"] - df["在手库存"]).astype("Float64")
    return df


COLS_RANK = ["销量排名", "SKU", "商品名称", "店铺", "销量", "累计占比", "下单次数", "每单件数",
             "销售额", "周均销量", "安全库存", "安全库存来源", "在手库存", "库存来源", "缺口", "备注"]
COLS_ALERT = ["SKU", "商品名称", "在手库存", "安全库存", "缺口", "周均销量", "近4周均销",
              "销量", "销量排名", "安全库存来源", "备注"]


def main():
    ap = argparse.ArgumentParser(description="销售分析 + 安全库存提醒 + Safety Stock 回写")
    ap.add_argument("sales", help="Odoo「Sales Analysis」透视导出")
    ap.add_argument("--products", required=True, help="product.product 导出（含 ID 唯一映射码）")
    ap.add_argument("--safety", help="运营维护的安全库存表（可选，但优先于推算）")
    ap.add_argument("--weeks", type=float, required=True,
                    help="销售数据覆盖的周数。**销售导出里没有任何日期，必须由你给**")
    ap.add_argument("--cover-weeks", type=float, default=2.0, help="安全库存按几周量算（默认 2）")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    args = ap.parse_args()

    print(f"销售数据覆盖 {args.weeks:g} 周，安全库存按 {args.cover_weeks:g} 周量推算")
    print("  ⚠ 期间周数是你给的，销售导出里没有日期——给错了周均和推算值全错\n")

    sales = read_sales(args.sales)
    prods, onhand_col = read_products(args.products)
    print(f"销售数据 {len(sales)} 个 SKU / 产品主数据 {len(prods)} 个 SKU")
    if onhand_col:
        print(f"  · 产品主数据带在手库存列 `{onhand_col}` —— 补货提醒可覆盖全部 SKU")
    else:
        print(f"  ⚠ 产品主数据没有在手库存列（试过 {P_ONHAND_CANDIDATES}）")
        print("    → 补货提醒只能覆盖安全库存表里那些 SKU。导出时勾上 Quantity On Hand 即可全覆盖。")

    hit = sales["SKU"].isin(set(prods["SKU"])).sum()
    print(f"  · 销售 ∩ 主数据: {hit}/{len(sales)} ({hit / len(sales) * 100:.1f}%)")
    miss = sorted(set(sales["SKU"]) - set(prods["SKU"]))
    if miss:
        print(f"    对不上主数据的 {len(miss)} 个（无 ERP ID，不能回写）: {', '.join(miss[:6])}"
              + (" …" if len(miss) > 6 else ""))

    safety = weeks_n = None
    if args.safety:
        safety, weeks_n, sc = read_safety(args.safety)
        print(f"  · 安全库存表 {len(safety)} 个 SKU（{weeks_n} 个周列"
              + (f"，在手库存列 `{sc}`" if sc else "") + "）")

    df = build(sales, prods, safety, args.weeks, args.cover_weeks)

    outdir = args.outdir or os.path.join("output", f"{date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)

    p, _ = write_simple(df[COLS_RANK], outdir, "销量排名.xlsx", left_cols={"商品名称", "备注"})
    print(f"\n✅ 销量排名: {p}")
    for q in (0.5, 0.8):
        n = int((df["累计占比"] <= q).sum()) + 1
        print(f"   前 {n} 个 SKU 贡献 {q * 100:.0f}% 销量（共 {len(df)} 个）")

    alert = df[df["缺口"] > 0].sort_values("缺口", ascending=False)
    p, _ = write_simple(alert[COLS_ALERT], outdir, "补货提醒.xlsx", left_cols={"商品名称", "备注"})
    print(f"\n🔔 补货提醒: {p}")
    known = df[df["库存来源"] != "缺"]
    print(f"   {len(alert)} 个 SKU 低于安全库存（有库存数据的共 {len(known)} 个）")
    if (df["库存来源"] == "缺").any():
        print(f"   · {(df['库存来源'] == '缺').sum()} 个 SKU 没有任何库存数据，未参与提醒")

    # ---- 回写表：只放**运营人工审过**的值 ----
    wb = df[(df["安全库存来源"] == "运营人工") & df["ERP_ID"].notna()]
    imp = pd.DataFrame({"id": wb["ERP_ID"], "Safety Stock": wb["安全库存_人工"].astype("Int64"),
                        "SKU": wb["SKU"], "商品名称": wb["商品名称"]})
    p, _ = write_simple(imp, outdir, "安全库存回写表.xlsx", left_cols={"商品名称"})
    print(f"\n📤 安全库存回写表: {p}")
    print(f"   {len(imp)} 条（仅运营人工值）→ 导入 ERP 产品主数据的 Safety Stock 字段")
    lost = df[(df["安全库存来源"] == "运营人工") & df["ERP_ID"].isna()]
    if len(lost):
        print(f"   ⚠ {len(lost)} 个有人工值但主数据里找不到 ERP ID，无法回写: {', '.join(lost['SKU'])}")

    # ---- 候选值：推算出来的，待运营审阅，**不进回写表** ----
    cand = df[(df["安全库存来源"] == "脚本推算") & (df["销量"] > 0)].copy()
    cand = cand.sort_values("销量", ascending=False)
    cc = ["SKU", "商品名称", "销量", "销量排名", "周均销量", "安全库存_推算", "在手库存", "ERP_ID"]
    p, _ = write_simple(cand[cc], outdir, "安全库存候选值.xlsx", left_cols={"商品名称"})
    print(f"\n📝 安全库存候选值: {p}")
    print(f"   {len(cand)} 个 SKU 没有人工值，已按 周均×{args.cover_weeks:g}周 推算候选值。")
    print("   这些值**不在回写表里**——推算基于期间均值，会低估上升期新品、高估下滑品，")
    print("   请运营审阅后并入安全库存表，下次跑就会走「运营人工」并进回写表。")

    # ---- 推算口径自检：拿运营已给的值反过来校准 ----
    both = df[df["安全库存来源"] == "运营人工"].dropna(subset=["安全库存_推算"])
    if len(both) >= 5:
        d = (both["安全库存_推算"].astype(float) - both["安全库存_人工"])
        within = (d.abs() <= both["安全库存_人工"] * 0.5).mean()
        print(f"\n🔍 推算口径自检（拿运营已给的 {len(both)} 个值反过来对）：")
        print(f"   推算值与人工值偏差在 ±50% 以内的占 {within * 100:.0f}%；"
              f"中位偏差 {d.median():+.0f} 件")
        print(f"   偏差大说明 --weeks {args.weeks:g} 或 --cover-weeks {args.cover_weeks:g} 可能不对，"
              f"或该品销量趋势变化大")


if __name__ == "__main__":
    main()
