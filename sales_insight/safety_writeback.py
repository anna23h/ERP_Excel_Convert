#!/usr/bin/env python3
"""运营安全库存表 → `Safety Stock` 回写表。**不需要销售数据。**

与同目录 `sales_insight.py` 的分工
---------------------------------
`sales_insight.py` 是「销售分析」流水线，安全库存回写只是它的副产品之一：它以**销售导出
为主表**做 left merge，SKU 全集由销售数据决定，安全库存表里有、销售导出里没有的 SKU 会被
整行丢掉。当运营只是发来一张「这些货要配多少安全库存」的表、根本不关心销量时，那条流水线
就不合用了——2026-09-22 实测：为了跑通它得伪造一份零销量的销售导出，而伪造数据一旦成为
月度动作，回写的正确性就挂在假数据上了。

本脚本反过来：**以安全库存表为主表**，只吃两份输入，只干一件事——把运营审过的值配上
ERP 的 External ID，产出可直接导入 Odoo 的回写表。

    python3 sales_insight/safety_writeback.py <安全库存表.xlsx> \
        --products <product.product.xlsx> [--col 15天安全库存] [--test-sku <SKU>]

产出
----
    安全库存回写表.xlsx   id / SKU(勿导入) / Safety Stock → 导入 ERP 产品主数据
    回写核对表.xlsx       ERP 现值 → 将写入 → 变化量，**导入前人工过目用**
    未匹配SKU.xlsx        对不上产品主数据的 SKU（有才出），多半是导出筛选条件错了
    安全库存回写表-试X.xlsx / 导入前快照-X.xlsx   给了 --test-sku 时多出这两份

**本脚本不碰 ERP**，只产导入文件，上传始终是人工动作（沿用 README「ERP 回写两条」
与 SPEC「第二档 Odoo API 暂不自动化」的既定分界）。

为什么不带 `Supply Remark` 列（2026-09-22 用户拍板）
---------------------------------------------------
Odoo 导入**只写文件里出现的列**，但「有列而该格为空」会**清空该字段**。运营这类需求表
只有数值、没有备注列，本脚本因此没有任何新内容要写进 `Supply Remark`——列整个不出现，
Odoo 就完全不碰它，`fs_writeback` 写的供应商画像与人工原文零风险。同一条理由在
`sales_insight` 的候选值表上已经用过一次。

代价是 ERP 里若残留 `sales_insight` 早先写的 `YYYYMMDD:安全库存 …` 段，日期会显得过时。
这是刻意接受的：过时的备注只是难看，被清空的供应商画像是真丢数据。

五道保护（每一条都对应文档里记过的坑）
--------------------------------------
1. 值列名**按「包含 `安全库存`」匹配**，认出 `15天安全库存` / `10天安全库存` 这类月月变的
   列名；一旦表里有多个这种列就**报错要求 --col 显式指定**，绝不自己猜口径。
2. 值为空 / 非整数 / ≤0 → **拒绝出表**并列出 Excel 行号。0 写进 ERP 是把安全库存清零，
   不能混在几十行里悄悄过去（真要清零请显式给 --allow-zero）。
3. SKU 对不上产品主数据 → **不写**，单独出表，并直接指向「产品主数据导出筛选条件」。
4. 产品主数据里同一 SKU 出现多行 → **报错**。`read_products` 是静默 `keep="first"`，
   2026-09-22 实测那份 10278 行的导出里全局就有 21 个重复 SKU；万一落到目标 SKU 上，
   会写到错误的 External ID 上去，而且事后完全看不出来。
5. `--test-sku` 沿用既有试水约定：单条表 + 导入前全字段快照，全量表照出不误。
"""
import argparse
import os
import sys
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # 让 common/ 可导入
from common.xlsx import write_simple  # noqa: E402
from sales_insight.sales_insight import read_products, P_SKU  # noqa: E402

#: 值列的识别关键字。运营的列名带口径前缀（`15天安全库存`、下次可能是 `10天安全库存`），
#: 故按包含匹配；`sales_insight.read_safety` 要求列名**恰好**是 `安全库存`，认不出这类表。
SAFETY_KEY = "安全库存"

COLS_CHECK = ["SKU", "商品名称", "供应商FS", "ERP现有安全库存", "将写入", "变化",
              "在手库存", "缺口"]


def _excel_row(idx):
    """DataFrame 行号 → Excel 里的行号（0 基 + 表头 1 行）。报错要让人能直接定位。"""
    return idx + 2


def pick_col(cols, want=None):
    """选出值列。→ 列名。

    want 给了就必须精确命中（列名当场列出来，省得为个空格来回试）；没给就按
    `SAFETY_KEY` 包含匹配，**命中多个直接报错**——「15天」还是「30天」是口径问题，
    猜错了整批数写歪，不如停下来让人指定。
    """
    cols = [str(c) for c in cols]
    if want:
        if want not in cols:
            raise ValueError(f"安全库存表里没有列 `{want}`。现有列：{cols}")
        return want
    hits = [c for c in cols if SAFETY_KEY in c]
    if not hits:
        raise ValueError(f"安全库存表里找不到含「{SAFETY_KEY}」的列。现有列：{cols}\n"
                         "→ 用 --col 显式指定要写哪一列。")
    if len(hits) > 1:
        raise ValueError(f"安全库存表里有 {len(hits)} 个含「{SAFETY_KEY}」的列：{hits}\n"
                         "→ 不同口径不能替人选，请用 --col 指定其中一个。")
    return hits[0]


def read_safety_sheet(path, sheet=None, col=None, allow_zero=False, say=print):
    """读运营的安全库存需求表 → (df[SKU/安全库存], 值列名)。

    这里**宁可报错也不跳过坏行**：跳过等于让人对着一份少了几行的回写表，事后
    很难发现某个货压根没被写。所有问题一次性收齐再抛，免得改一行跑一遍。
    """
    df = pd.read_excel(path, sheet_name=0 if sheet is None else sheet)
    if df.empty:
        raise ValueError(f"安全库存表是空的（sheet={sheet or '第一个'}）")
    sku_col = next((c for c in df.columns if "SKU" in str(c).upper()), None)
    if sku_col is None:
        sku_col = df.columns[0]
        say(f"  ⚠ 没找到含 `SKU` 的列，按第一列 `{sku_col}` 当 SKU 用")
    val_col = pick_col(df.columns, col)

    out = pd.DataFrame({
        "SKU": df[sku_col].astype(str).str.strip(),
        "安全库存": pd.to_numeric(df[val_col], errors="coerce"),
    })
    out = out[(out["SKU"] != "") & (out["SKU"].str.lower() != "nan")]

    bad = []
    blank = out[out["安全库存"].isna()]
    if len(blank):
        bad.append(f"  · {len(blank)} 行值为空或不是数字：" +
                   ", ".join(f"{s}(第{_excel_row(i)}行)" for i, s in
                             zip(blank.index[:8], blank["SKU"][:8])))
    frac = out[out["安全库存"].notna() & (out["安全库存"] % 1 != 0)]
    if len(frac):
        bad.append(f"  · {len(frac)} 行不是整数：" +
                   ", ".join(f"{s}={v}" for s, v in
                             zip(frac["SKU"][:8], frac["安全库存"][:8])))
    neg = out[out["安全库存"] < 0]
    if len(neg):
        bad.append(f"  · {len(neg)} 行是负数：" + ", ".join(neg["SKU"][:8]))
    zero = out[out["安全库存"] == 0]
    if len(zero) and not allow_zero:
        bad.append(f"  · {len(zero)} 行值为 0：" + ", ".join(zero["SKU"][:8]) +
                   "\n    （0 导进 ERP 是把安全库存**清零**。确实要清零请加 --allow-zero）")
    if bad:
        raise ValueError(f"安全库存表 `{val_col}` 列有问题，已中止（一条都没产出）：\n"
                         + "\n".join(bad))

    dup = out[out["SKU"].duplicated(keep=False)]
    if len(dup):
        pairs = dup.groupby("SKU")["安全库存"].apply(lambda s: "/".join(f"{v:g}" for v in s))
        raise ValueError("安全库存表里有重复 SKU，到底写哪个值无法替你决定：\n" +
                         "\n".join(f"  · {k} → {v}" for k, v in pairs.items()))

    out["安全库存"] = out["安全库存"].astype("Int64")
    return out.reset_index(drop=True), val_col


def check_product_dups(products_path, skus, say=print):
    """产品主数据里目标 SKU 是否出现多行 → 有就抛。

    `read_products` 做的是静默 `drop_duplicates(keep="first")`，对销售分析那条
    流水线无伤大雅，但回写是**按 External ID 改主数据**：取错那一行等于把值写到
    另一个产品上，且导入日志里看不出异常。故这里单独再扫一遍 SKU 列。
    """
    raw = pd.read_excel(products_path, usecols=[P_SKU])
    s = raw[P_SKU].astype(str).str.strip()
    hit = s[s.isin(set(skus))]
    dup = sorted(hit[hit.duplicated(keep=False)].unique())
    if dup:
        raise ValueError(
            f"产品主数据里这 {len(dup)} 个 SKU 各占多行，无法确定该写哪个 External ID：\n"
            f"  {', '.join(dup[:12])}" + (" …" if len(dup) > 12 else "") + "\n"
            "→ 多半是同一 SKU 建了多个产品档。请先在 ERP 里处理，或把重复档排除后重导。")
    say(f"  · 产品主数据里目标 SKU 无重复档（扫了 {len(raw)} 行）")


def run(safety_path, products_path, col=None, sheet=None, outdir=None,
        test_sku=None, allow_zero=False):
    """跑完整条流水线 → (输出目录, 摘要行列表)。CLI 与 GUI 共用一份逻辑与摘要。

    ⚠ 出错一律 `ValueError` 而非 `SystemExit`：后者是 BaseException，GUI 后台线程的
    `except Exception` 抓不到，界面会永远卡在「运行中」按钮禁用态。
    """
    L = []
    say = L.append

    safety, val_col = read_safety_sheet(safety_path, sheet, col, allow_zero, say)
    say(f"安全库存表 {len(safety)} 个 SKU，写的是 `{val_col}` 列"
        f"（{safety['安全库存'].min()}–{safety['安全库存'].max()}）")

    prods, onhand_col, xid_col, has_remark = read_products(products_path)
    say(f"产品主数据 {len(prods)} 个 SKU，回写映射码取自 `{xid_col}` 列")
    check_product_dups(products_path, set(safety["SKU"]), say)
    if "ERP现有安全库存" not in prods or prods["ERP现有安全库存"].isna().all():
        say("  ⚠ 产品主数据没有 `Safety Stock` 列 —— 核对表看不到 ERP 现值，"
            "无从判断哪些行其实没变化（导出时勾上该列即可）")
    say("  · 本脚本不写 `Supply Remark`：该列整个不出现，Odoo 不会碰这个字段")

    df = safety.merge(prods, on="SKU", how="left")
    outdir = outdir or os.path.join("output", f"{date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)

    # ---- 对不上主数据的：不写，单独出表 ----
    lost = df[df["ERP_ID"].isna()]
    if len(lost):
        lp, _ = write_simple(pd.DataFrame({"SKU": lost["SKU"], "本该写入": lost["安全库存"]}),
                             outdir, "未匹配SKU.xlsx")
        say(f"\n⚠ {len(lost)}/{len(df)} 个 SKU 在产品主数据里找不到，**不会被写入**: {lp}")
        say(f"   {', '.join(lost['SKU'][:8])}" + (" …" if len(lost) > 8 else ""))
        say("   → 十有八九是产品主数据的**导出筛选条件**把它们排除了。"
            "筛选只勾 `can be sold`——实测 `VO active=true` 会漏掉 x2/x3 组合装与 _VO/_GW 渠道变体。")

    ok = df[df["ERP_ID"].notna()].copy()
    if ok.empty:
        raise ValueError("没有一个 SKU 对得上产品主数据，回写表会是空的。\n"
                         "→ 先按上面的提示重导产品主数据（筛选只勾 `can be sold`）。")

    def _writeback(rows, fname):
        """→ (落盘路径, 导入df)。只有三列，**刻意不带 `Supply Remark`**，见模块 docstring。

        `SKU(勿导入)` 这个表头 Odoo 认不出，导入时显示为未映射，纯给人核对用。
        **不能叫 `Internal Reference`**——那会被自动映射，忘了取消勾选就重写了 SKU。
        """
        t = pd.DataFrame({"id": rows["ERP_ID"],
                          "SKU(勿导入)": rows["SKU"],
                          "Safety Stock": rows["安全库存"].astype("Int64")})
        return write_simple(t, outdir, fname)[0], t

    # ---- 核对表：导入前给人看的唯一依据 ----
    chk = pd.DataFrame({
        "SKU": ok["SKU"],
        "商品名称": ok["商品名称"],
        "供应商FS": ok["供应商FS"],
        "ERP现有安全库存": ok["ERP现有安全库存"],
        "将写入": ok["安全库存"],
        "变化": (ok["安全库存"].astype("Float64") - ok["ERP现有安全库存"]),
        "在手库存": ok["在手库存"],
        "缺口": (ok["安全库存"].astype("Float64") - ok["在手库存"]),
    }).sort_values("将写入", ascending=False)
    cp, _ = write_simple(chk[COLS_CHECK], outdir, "回写核对表.xlsx",
                         left_cols={"商品名称"})

    # ---- 试水：单条表 + 导入前快照 ----
    if test_sku:
        one = ok[ok["SKU"] == test_sku]
        if one.empty:
            if test_sku not in set(safety["SKU"]):
                raise ValueError(f"试水 SKU {test_sku} 不在安全库存表里。\n"
                                 f"表里有 {len(safety)} 个，如 "
                                 f"{', '.join(safety['SKU'][:5])} …")
            raise ValueError(f"试水 SKU {test_sku} 在产品主数据里找不到，拿不到 ERP ID，"
                             "无法回写（见上面未匹配清单）。")
        snap = write_simple(prods[prods["SKU"] == test_sku], outdir,
                            f"导入前快照-{test_sku}.xlsx",
                            left_cols={"商品名称", "供应商FS", "ERP现有备注"})[0]
        tp, _ = _writeback(one, f"安全库存回写表-试{test_sku}.xlsx")
        r = one.iloc[0]
        say(f"\n📸 导入前快照: {snap}（导入后再导一次同一商品，逐字段对比）")
        say(f"🧪 回写表·试水: {tp}")
        cur = r["ERP现有安全库存"]
        say(f"   只有 {test_sku} 这一条：ERP 现值 "
            f"{'未配/未导出该列' if pd.isna(cur) else f'{cur:g}'} → 写入 {r['安全库存']}")
        say("   **先导它**，回 ERP 核对无误再导下面那份全量。")

    # 全量那份**照出不误**：试水验完直接导它，不必为了拿全量再跑一遍
    # （跑两遍之间源表可能已改，验过的和导入的就不是同一批了）
    p, imp = _writeback(ok, "安全库存回写表.xlsx")
    say(f"\n📤 安全库存回写表{'·全量' if test_sku else ''}: {p}")
    say(f"   {len(imp)} 条，列: {' / '.join(imp.columns)} → 导入 ERP 产品主数据")
    say("   · `SKU(勿导入)` 只是给你人工核对用，Odoo 认不出这个表头，不会被写进去")
    say(f"\n🔍 回写核对表: {cp}（导入前对着它过一遍）")

    cur = ok["ERP现有安全库存"]
    if cur.notna().any():
        same = int((cur == ok["安全库存"].astype("Float64")).sum())
        fresh = int(cur.isna().sum())
        chg = ok[cur.notna() & (cur != ok["安全库存"].astype("Float64"))]
        say(f"   {fresh} 个 ERP 里还没配（新写入）/ {same} 个与现值相同（无变化）/ "
            f"{len(chg)} 个有变动")
        if len(chg):
            big = chg.assign(d=(chg["安全库存"].astype("Float64") - cur[chg.index]).abs()) \
                     .sort_values("d", ascending=False).head(5)
            say("   变动最大的 5 个: " + "，".join(
                f"{r['SKU']} {r['ERP现有安全库存']:.0f}→{r['安全库存']}"
                for _, r in big.iterrows()))
    if onhand_col:
        short = int((chk["缺口"] > 0).sum())
        say(f"   按这批新值算，{short} 个 SKU 当前在手低于安全库存")
    return outdir, L


def main():
    ap = argparse.ArgumentParser(
        description="运营安全库存表 → Safety Stock 回写表（不需要销售数据）")
    ap.add_argument("safety", help="运营维护/发来的安全库存表")
    ap.add_argument("--products", required=True,
                    help="product.product 导出（筛选只勾 `can be sold`，须含 External ID）")
    ap.add_argument("--col", help="要写的值列名，如 `15天安全库存`。"
                                  "不给则自动认含「安全库存」的列，命中多个会报错")
    ap.add_argument("--sheet", help="工作表名（默认第一个）")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    ap.add_argument("--test-sku", help="首次导入试水：只为这一个 SKU 产回写表，"
                                       "外加一份导入前的全字段快照供事后比对")
    ap.add_argument("--allow-zero", action="store_true",
                    help="允许值为 0 的行（= 把该 SKU 的安全库存清零）。默认拒绝")
    args = ap.parse_args()
    try:
        _, lines = run(args.safety, args.products, args.col, args.sheet,
                       args.outdir, args.test_sku, args.allow_zero)
    except ValueError as e:
        raise SystemExit(str(e))
    print("\n".join(lines))


if __name__ == "__main__":
    main()
