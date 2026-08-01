#!/usr/bin/env python3
"""销售分析 + 安全库存提醒 + Safety Stock 回写 ERP。

三份输入合一：
    销售数据   Odoo「Sales Analysis」透视导出。两种格式都吃：
               按周分组（W28/W29/… 各 3 列 + 合计组）→ **期间周数自动数出来**；
               整期累计（4 列）→ 没有日期信息，须 --weeks 手工给。
    产品主数据 product.product 导出（External ID / Internal Reference / Name /
               建议勾上 Quantity On Hand、Safety Stock、Supply Remark）
    安全库存表 运营手工维护（可选，但**优先于脚本推算**）

四份产出：
    销量排名.xlsx        全 SKU 按销量排序，含累计占比（ABC 分析）
    补货提醒.xlsx        在手库存 < 安全库存 的 SKU，按缺口降序
    安全库存回写表.xlsx   id / SKU(勿导入) / Safety Stock / Supply Remark → 导入 ERP
    安全库存候选值.xlsx   脚本推算的值，**待运营审阅**，不进回写表

为什么推算值不自动写回：推算基于**期间均值**，会低估上升期新品、高估下滑品。
未经人审的阈值不许进 ERP 主数据（2026-08-01 用户拍板）。

回写表的两处讲究（2026-08-01）：
  · `SKU(勿导入)` 这个表头 Odoo 认不出，导入时显示为未映射，纯给人核对用。
    **不能叫 Internal Reference**——那会被自动映射，忘了取消勾选就重写了 SKU。
  · `Supply Remark` 每一行都带 ERP 现值（运营备注前置为 `YYYYMMDD:安全库存 …` 段）。
    Odoo 对「有列但留空」的处理是**清空该字段**，所以没新备注的行也必须带原值回去。

用法:
    python3 sales_insight/sales_insight.py <销售数据.xlsx> \
        --products <product.product.xlsx> --safety <安全库存.xlsx>
"""
import argparse
import os
import re
import sys
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # 让 common/ 可导入
from common.xlsx import write_simple  # noqa: E402
from common import remark as rk  # noqa: E402

# product.product 导出的列名
# ⚠ 回写映射码必须用 **External ID**（`__export__.product_product_*`），不是数据库 ID。
# 不同导出模版下这两者的列名会打架：有的模版把 External ID 直接叫 `ID`（2026-07-30 那份），
# 有的模版 `ID`=数据库整数 205448、另有一列 `External ID`=字符串（2026-08-01 那份）。
# 故按「先找 External ID，找不到再看 ID 是不是 __export__ 形式」的顺序认。
P_XID_CANDIDATES = ["External ID", "ID"]
P_SKU = "Internal Reference"
P_NAME = "Name"
P_SHOP = "VO Shop Name"
P_SAFETY = "Safety Stock"          # ERP 现值，用来看两边是否已同步
# 供应商（去谁家订），补货提醒里很有用。列名随导出模版变：有的模版走关联字段
# `Product/FS`，有的直接就是 `FS`（2026-08-01 那份完整导出），按顺序试。
P_FS_CANDIDATES = ["Product/FS", "FS"]
# ⚠ 导出里同时有 `Supply Remark` 与 `Product/Supply Remark`，认前者——
# 与 vo_orders/fs_writeback.py 写的是同一个字段，两个脚本必须对齐，否则各写各的。
P_REMARK = "Supply Remark"
# 在手库存列名在不同导出模版下可能不同，按顺序试
P_ONHAND_CANDIDATES = ["Quantity On Hand", "Qty On Hand", "On Hand", "Quantity available"]


WEEK_RE = re.compile(r"^\s*W(\d+)\b")

# 本脚本写进 Supply Remark 的段：`20260801:安全库存 <运营备注原文>`
REMARK_PREFIX = "安全库存"
REMARK_SIG = rk.signature(REMARK_PREFIX)


def merge_remark(old, new, d):
    """ERP 现值 + 本次运营备注 → 回写值。没有新备注就原样带回现值（否则 Odoo 会清空）。"""
    new = "" if new is None or new != new else str(new).strip()
    seg = f"{d}:{REMARK_PREFIX} {new}" if new else ""
    return rk.merge(old, seg, REMARK_SIG)


def read_sales(path, say=print):
    """读 Odoo「Sales Analysis」透视导出。新旧两种格式自动识别。

    两种都见过：
      旧（整期累计）  4 列：商品 / Untaxed Total / # of Lines / Qty Ordered
      新（按周分组） 16 列：W28/W29/W30/W31 各 3 列 + 末尾一组无周标签的合计

    统一靠「首列 == Total 的那一行」定位：它上面一行是三件套子表头，再上一行
    若带 `W\\d+` 就是周分组行。列从第 1 列起每 3 列切一组；带周标签的是周分组，
    不带的那组是整期合计——销售额/下单次数/销量一律取合计组，旧格式即
    「只有一组、且它就是合计组」，天然兼容。

    明细的商品列形如 `    [SKU] 商品名`——**前面有缩进空格**，正则必须容忍
    （`^\\[` 匹配不到，会得到 0 条，别问我怎么知道的）。

    say: 告警去处。默认 print(CLI)，GUI 传进来的是日志区收集器——
    这两条告警(SKU 解析失败、Total 对不上)正是解析出问题的信号，不能只在终端有。

    → (df[SKU/商品/销售额/下单次数/销量 + 每个周列], 周标签列表)
    """
    raw = pd.read_excel(path, header=None)
    if raw.shape[1] < 4 or (raw.shape[1] - 1) % 3:
        raise ValueError(f"销售导出应为「商品列 + 每 3 列一组」，实得 {raw.shape[1]} 列")

    # Total 行是锚。找不到就退回旧格式的固定位置（第 4 行）。
    first = raw[0].astype(str).str.strip()
    hits = raw.index[first == "Total"]
    t = int(hits[0]) if len(hits) else 3
    total = raw.iloc[t]

    # 组标签取自子表头(t-1)的上一行；缺行或整行空 = 没有周分组
    labels = raw.iloc[t - 2] if t >= 2 else None
    groups = []                              # [(标签 or None, 起始列)]
    for c in range(1, raw.shape[1], 3):
        lab = None if labels is None else str(labels[c])
        m = WEEK_RE.match(lab) if lab and lab != "nan" else None
        groups.append((m.group(0).strip() if m else None, c))

    weeks = [(lab, c) for lab, c in groups if lab]
    rest = [(lab, c) for lab, c in groups if not lab]
    if len(rest) != 1:
        raise ValueError(f"销售导出里认出 {len(weeks)} 个周分组、{len(rest)} 个合计组，"
                         "应恰好有一个不带 W 标签的合计组——导出模版可能变了")
    tot_c = rest[0][1]

    df = raw.iloc[t + 1:].copy()
    out = pd.DataFrame({"商品": df[0]})
    out["SKU"] = out["商品"].astype(str).str.extract(r"^\s*\[([^\]]+)\]")[0]
    bad = out["SKU"].isna().sum()
    if bad:
        say(f"  ⚠ {bad} 行解析不出 SKU，已跳过")
    for name, off in (("销售额", 0), ("下单次数", 1), ("销量", 2)):
        out[name] = pd.to_numeric(df[tot_c + off], errors="coerce").fillna(0)
    # 某周无销售 = 空格，含义是 0 而非缺失，故 fillna(0) 后才能参与均值
    for lab, c in weeks:
        out[lab] = pd.to_numeric(df[c + 2], errors="coerce").fillna(0)

    out = out[out["SKU"].notna()]

    # 与导出自带的 Total 行对拍——解析错了这里会立刻暴露
    for name, off in (("销售额", 0), ("下单次数", 1), ("销量", 2)):
        want = pd.to_numeric(total[tot_c + off], errors="coerce")
        got = out[name].sum()
        if pd.notna(want) and abs(want - got) > 0.01:
            say(f"  ⚠ {name} 明细求和 {got:.2f} ≠ 导出 Total {want:.2f}")
    return out.reset_index(drop=True), [lab for lab, _ in weeks]


def read_products(path):
    """读 product.product 导出。Internal Reference 前面带 `\\t`（Odoo 强制文本），必须 strip。"""
    df = pd.read_excel(path)
    if P_SKU not in df.columns:
        raise ValueError(f"产品主数据缺列 `{P_SKU}`（现有: {list(df.columns)}）")
    xid = next((c for c in P_XID_CANDIDATES if c in df.columns
                and df[c].astype(str).str.startswith("__export__").any()), None)
    if xid is None:
        raise ValueError("产品主数据里找不到 External ID 列（`__export__.product_product_*` 形式）。\n"
                         "导出时必须勾上 External ID —— 数据库整数 ID 不能当导入映射码。")
    df["SKU"] = df[P_SKU].astype(str).str.strip()
    df = df[df["SKU"].notna() & (df["SKU"] != "") & (df["SKU"] != "nan")]
    df = df.drop_duplicates(subset="SKU", keep="first")

    onhand = next((c for c in P_ONHAND_CANDIDATES if c in df.columns), None)
    fs = next((c for c in P_FS_CANDIDATES if c in df.columns), None)
    out = pd.DataFrame({
        "SKU": df["SKU"],
        "ERP_ID": df[xid],
        "商品名称": df[P_NAME] if P_NAME in df else "",
        "店铺": df[P_SHOP] if P_SHOP in df else "",
        "在手库存": pd.to_numeric(df[onhand], errors="coerce") if onhand else pd.NA,
        "ERP现有安全库存": pd.to_numeric(df[P_SAFETY], errors="coerce") if P_SAFETY in df else pd.NA,
        "供应商FS": df[fs] if fs else "",
        # 回写 Supply Remark 时要把原文带回去，否则 Odoo 会把这个字段清空
        "ERP现有备注": df[P_REMARK].fillna("").astype(str) if P_REMARK in df else "",
    })
    return out.reset_index(drop=True), onhand, xid, P_REMARK in df.columns


def read_safety(path):
    """读运营维护的安全库存表。

    列名带日期前缀（`7.29在手库存`），故按「包含」匹配而非写死——
    下个月那份会叫 `8.31在手库存`。
    """
    df = pd.read_excel(path)
    sku_col = next((c for c in df.columns if "SKU" in str(c).upper()), df.columns[0])
    safe_col = next((c for c in df.columns if str(c).strip() == "安全库存"), None)
    if safe_col is None:
        raise ValueError(f"安全库存表里找不到「安全库存」列（现有: {list(df.columns)}）")
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


def build(sales, prods, safety, weeks, cover_weeks, week_cols=()):
    """三份合一 → 一张宽表。其余产出都是它的切片。"""
    df = sales.merge(prods, on="SKU", how="left")
    if safety is not None:
        df = df.merge(safety, on="SKU", how="left")
    else:
        for c in ("安全库存_人工", "在手库存_运营表", "备注", "近4周均销"):
            df[c] = pd.NA

    # 近4周均销：销售导出带周分组时优先用它——覆盖全部 SKU 且是当下的数；
    # 安全库存表那份只有 62 行、且是运营手工快照，仅作退路。
    if week_cols:
        recent = list(week_cols)[-4:]
        df["近4周均销"] = df[recent].mean(axis=1).round(1)

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


COLS_RANK = ["销量排名", "SKU", "商品名称", "供应商FS", "销量", "累计占比", "下单次数", "每单件数",
             "销售额", "周均销量", "近4周均销", "安全库存", "安全库存来源", "ERP现有安全库存",
             "在手库存", "库存来源", "缺口", "备注"]
COLS_ALERT = ["SKU", "商品名称", "供应商FS", "在手库存", "安全库存", "缺口", "周均销量",
              "近4周均销", "销量", "销量排名", "安全库存来源", "备注"]


def run(sales_path, products_path, safety_path=None, weeks=None, cover_weeks=2.0,
        outdir=None, test_sku=None):
    """跑完整条流水线 → (输出目录, 摘要行列表)。CLI 与 GUI 共用一份逻辑与摘要。

    ⚠ 出错一律 `ValueError` 而非 `SystemExit`：后者是 BaseException，GUI 后台线程的
    `except Exception` 抓不到，界面会永远卡在「运行中」按钮禁用态（箱单那次踩过）。
    """
    L = []
    say = L.append

    sales, week_cols = read_sales(sales_path, say)

    # 期间周数：导出带周分组就数表头，比人工给靠谱得多（给错则周均和推算值全错）
    if week_cols:
        n_weeks = float(len(week_cols))
        say(f"销售导出按周分组：{len(week_cols)} 个周（{week_cols[0]}–{week_cols[-1]}），"
            f"期间周数自动取 {n_weeks:g}")
        if weeks and weeks != n_weeks:
            say(f"  ⚠ 你显式给了周数 {weeks:g}，与表头数出来的 {len(week_cols)} 不符，照你给的算")
            n_weeks = weeks
    elif weeks:
        n_weeks = weeks
        say(f"销售导出无周分组（整期累计），期间周数用你给的 {n_weeks:g} 周")
        say("  ⚠ 这个数是你给的，导出里没有日期——给错了周均和推算值全错")
    else:
        raise ValueError("销售导出里认不出周分组（是整期累计格式），必须给出期间周数。\n"
                         "或者在 ERP 里按周分组导出，脚本就能自己数出来。")
    say(f"安全库存按 {cover_weeks:g} 周量推算\n")

    prods, onhand_col, xid_col, has_remark = read_products(products_path)
    say(f"销售数据 {len(sales)} 个 SKU / 产品主数据 {len(prods)} 个 SKU")
    say(f"  · 回写映射码取自 `{xid_col}` 列")
    if onhand_col:
        say(f"  · 产品主数据带在手库存列 `{onhand_col}` —— 补货提醒可覆盖全部 SKU")
    else:
        say(f"  ⚠ 产品主数据没有在手库存列（试过 {P_ONHAND_CANDIDATES}）")
        say("    → 补货提醒只能覆盖安全库存表里那些 SKU。导出时勾上 Quantity On Hand 即可全覆盖。")
    if not has_remark:
        say(f"  ⚠ 产品主数据没有 `{P_REMARK}` 列 —— 回写表将不带备注列")
        say("    （带了会把 ERP 里已有的备注清空，故直接不写。导出时勾上该列即可回写运营备注。）")

    hit = sales["SKU"].isin(set(prods["SKU"])).sum()
    say(f"  · 销售 ∩ 主数据: {hit}/{len(sales)} ({hit / len(sales) * 100:.1f}%)")
    miss = sorted(set(sales["SKU"]) - set(prods["SKU"]))
    if miss:
        say(f"    对不上主数据的 {len(miss)} 个（无 ERP ID，不能回写）: {', '.join(miss[:6])}"
            + (" …" if len(miss) > 6 else ""))

    safety = None
    if safety_path:
        safety, weeks_n, sc = read_safety(safety_path)
        say(f"  · 安全库存表 {len(safety)} 个 SKU（{weeks_n} 个周列"
            + (f"，在手库存列 `{sc}`" if sc else "") + "）")

    df = build(sales, prods, safety, n_weeks, cover_weeks, week_cols)

    outdir = outdir or os.path.join("output", f"{date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)

    p, _ = write_simple(df[COLS_RANK], outdir, "销量排名.xlsx", left_cols={"商品名称", "备注"})
    say(f"\n✅ 销量排名: {p}")
    for q in (0.5, 0.8):
        k = int((df["累计占比"] <= q).sum()) + 1
        say(f"   前 {k} 个 SKU 贡献 {q * 100:.0f}% 销量（共 {len(df)} 个）")

    alert = df[df["缺口"] > 0].sort_values("缺口", ascending=False)
    p, _ = write_simple(alert[COLS_ALERT], outdir, "补货提醒.xlsx", left_cols={"商品名称", "备注"})
    say(f"\n🔔 补货提醒: {p}")
    known = df[df["库存来源"] != "缺"]
    say(f"   {len(alert)} 个 SKU 低于安全库存（有库存数据的共 {len(known)} 个）")
    if (df["库存来源"] == "缺").any():
        say(f"   · {(df['库存来源'] == '缺').sum()} 个 SKU 没有任何库存数据，未参与提醒")

    # ---- 回写表：只放**运营人工审过**的值 ----
    # Odoo 只写文件里出现的列，但「出现且为空」会把该字段清空——所以每一列都要
    # 么是真想改的、要么必须带上原值，绝不能出现「有列但留空」。
    wb_all = df[(df["安全库存来源"] == "运营人工") & df["ERP_ID"].notna()].copy()
    today = f"{date.today():%Y%m%d}"

    def _writeback(rows, fname):
        """→ (落盘路径, 导入df)。列见 COLS_IMP 的说明。"""
        t = pd.DataFrame({"id": rows["ERP_ID"],
                          "SKU(勿导入)": rows["SKU"],
                          "Safety Stock": rows["安全库存_人工"].astype("Int64")})
        if has_remark:
            t[P_REMARK] = [merge_remark(o, x, today) for o, x in
                           zip(rows["ERP现有备注"], rows["备注"])]
        return write_simple(t, outdir, fname, left_cols={P_REMARK})[0], t

    if test_sku:
        one = wb_all[wb_all["SKU"] == test_sku]
        if one.empty:
            # 逐级判定到底缺哪一环。三种原因的处理方式完全不同（换销售导出 /
            # 去安全库存表加一行 / 重导产品主数据），糊成一句话等于让人自己猜。
            row = df[df["SKU"] == test_sku]
            if row.empty:
                raise ValueError(
                    f"试水 SKU {test_sku} 不在销售数据里——这份销售导出里没有它的销量记录。\n"
                    "检查 SKU 是否写错，或换一份覆盖到它的销售导出。")
            r0 = row.iloc[0]
            if pd.isna(r0["ERP_ID"]):
                # 十有八九是产品主数据导出的筛选条件漏了一批，同类的一起列出来更快
                lost_all = sorted(df[(df["安全库存来源"] == "运营人工")
                                     & df["ERP_ID"].isna()]["SKU"])
                raise ValueError(
                    f"试水 SKU {test_sku} 在产品主数据里找不到，拿不到 ERP ID，无法回写。\n"
                    "→ 多半是**产品主数据导出的筛选条件**把它排除了，请检查导出条件后重导。\n"
                    + (f"同样情况的还有 {len(lost_all)} 个（很可能是同一个筛选造成的）：\n"
                       f"  {', '.join(lost_all)}" if len(lost_all) > 1 else ""))
            raise ValueError(
                f"试水 SKU {test_sku} 在运营的安全库存表里没有值，走的是脚本推算"
                f"（推算值 {r0['安全库存_推算']}），而推算值不进回写表。\n"
                "→ 要试水它，先把它加进运营安全库存表；或改用回写表里已有的 SKU。")
        snap = write_simple(
            prods[prods["SKU"] == test_sku].merge(
                df[df["SKU"] == test_sku][["SKU", "安全库存", "销量", "周均销量"]], on="SKU"),
            outdir, f"导入前快照-{test_sku}.xlsx", left_cols={"商品名称", "供应商FS"})[0]
        pt, _ = _writeback(one, f"安全库存回写表-试{test_sku}.xlsx")
        say(f"\n📸 导入前快照: {snap}（导入后再导一次同一商品，逐字段对比）")
        say(f"🧪 回写表·试水: {pt}")
        say(f"   只有 {test_sku} 这一条。**先导它**，回 ERP 核对无误再导下面那份全量。")
        r = one.iloc[0]
        say(f"   ERP 现值 {df.set_index('SKU').loc[test_sku, 'ERP现有安全库存']}"
            f" → 写入 {r['安全库存_人工']:.0f}")

    # 全量那份**照出不误**：试水验完直接导它，不必为了拿全量再跑一遍
    # （跑两遍之间数据可能已变，验过的和导入的就不是同一批了）
    p, imp = _writeback(wb_all, "安全库存回写表.xlsx")
    say(f"\n📤 安全库存回写表{'·全量' if test_sku else ''}: {p}")
    say(f"   {len(imp)} 条，列: {' / '.join(imp.columns)} → 导入 ERP 产品主数据")
    say("   · `SKU(勿导入)` 只是给你人工核对用，Odoo 认不出这个表头，不会被写进去")
    if has_remark:
        touched = sum(1 for o, x in zip(wb_all["ERP现有备注"], wb_all["备注"])
                      if merge_remark(o, x, today) != (o or ""))
        say(f"   · `{P_REMARK}`: {touched} 条带上了运营备注（前置 `{today}:安全库存 …`），"
            f"其余 {len(imp) - touched} 条**原样带回 ERP 现值**")
        say("     （不带原值的话，Odoo 会把这些产品已有的备注清空——含 FS 回写写进去的供应商画像）")
    lost = df[(df["安全库存来源"] == "运营人工") & df["ERP_ID"].isna()]
    if len(lost):
        say(f"   ⚠ {len(lost)} 个有人工值但主数据里找不到 ERP ID，无法回写: {', '.join(lost['SKU'])}")

    # ---- 候选值：推算出来的，待运营审阅 ----
    # 前三列与回写表**完全一致**，审完可直接导；后面的中文表头 Odoo 认不出、
    # 显示为未映射，不勾就不会被写——与 `SKU(勿导入)` 同一个路子，
    # 这样「能直接导」和「审的时候看得到依据」不必二选一。
    # 不带 Supply Remark：候选品不在运营安全库存表里、没有运营备注，
    # 带这列只会平白无故重写该字段。
    cand = df[(df["安全库存来源"] == "脚本推算") & (df["安全库存_推算"] > 0)
              & df["ERP_ID"].notna()].copy()
    cand = cand.sort_values("销量", ascending=False)
    ct = pd.DataFrame({"id": cand["ERP_ID"],
                       "SKU(勿导入)": cand["SKU"],
                       "Safety Stock": cand["安全库存_推算"].astype("Int64"),
                       "商品名称": cand["商品名称"], "销量": cand["销量"],
                       "销量排名": cand["销量排名"], "周均销量": cand["周均销量"],
                       "在手库存": cand["在手库存"]})
    p, _ = write_simple(ct, outdir, "安全库存候选值.xlsx", left_cols={"商品名称"})
    n_zero = int(((df["安全库存来源"] == "脚本推算") & (df["安全库存_推算"] <= 0)).sum())
    say(f"\n📝 安全库存候选值: {p}")
    say(f"   {len(ct)} 个 SKU 没有人工值，已按 周均×{cover_weeks:g}周 推算候选值。")
    say(f"   列与回写表同头，审完可直接导；`商品名称` 起的中文列 Odoo 认不出，不会被写。")
    if n_zero:
        say(f"   · 另有 {n_zero} 个推算值为 0（低销量长尾）已剔除——导进去只会把它们刷成 0")
    say("   ⚠ 推算基于期间均值，会低估上升期新品、高估下滑品，**务必先人审**。")

    # ---- 推算口径自检：拿运营已给的值反过来校准 ----
    both = df[df["安全库存来源"] == "运营人工"].dropna(subset=["安全库存_推算"])
    if len(both) >= 5:
        dev = (both["安全库存_推算"].astype(float) - both["安全库存_人工"])
        within = (dev.abs() <= both["安全库存_人工"] * 0.5).mean()
        say(f"\n🔍 推算口径自检（拿运营已给的 {len(both)} 个值反过来对）：")
        say(f"   推算值与人工值偏差在 ±50% 以内的占 {within * 100:.0f}%；"
            f"中位偏差 {dev.median():+.0f} 件")
        say(f"   偏差大说明 期间 {n_weeks:g} 周 或 覆盖 {cover_weeks:g} 周 可能不对，"
            f"或该品销量趋势变化大")
    return outdir, L


def main():
    ap = argparse.ArgumentParser(description="销售分析 + 安全库存提醒 + Safety Stock 回写")
    ap.add_argument("sales", help="Odoo「Sales Analysis」透视导出")
    ap.add_argument("--products", required=True, help="product.product 导出（含 ID 唯一映射码）")
    ap.add_argument("--safety", help="运营维护的安全库存表（可选，但优先于推算）")
    ap.add_argument("--weeks", type=float,
                    help="销售数据覆盖的周数。导出按周分组时自动从表头数出来，"
                         "此参数只作覆盖；旧的整期累计导出没有周信息，必须给")
    ap.add_argument("--cover-weeks", type=float, default=2.0, help="安全库存按几周量算（默认 2）")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    ap.add_argument("--test-sku", help="首次导入试水：只为这一个 SKU 产回写表，"
                                       "外加一份导入前的全字段快照供事后比对")
    args = ap.parse_args()
    try:
        _, lines = run(args.sales, args.products, args.safety, args.weeks,
                       args.cover_weeks, args.outdir, args.test_sku)
    except ValueError as e:
        raise SystemExit(str(e))
    print("\n".join(lines))


if __name__ == "__main__":
    main()
