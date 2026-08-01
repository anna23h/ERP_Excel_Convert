#!/usr/bin/env python3
"""采购对账：采购记录 PO ↔ 财务实收 PO，算出「还有什么没到」。

背景（见 po_reconcile/README.md）：
    采购在一张 PO 里记录订购需求；供应商分批送货；财务按每批实收**另建一张 PO**
    （产品与数量与供应商 invoice 一致）。于是采购那张 PO 永远显示原始订购量，
    看不出实际还缺什么。

    注意：采购 PO 的 `Billed Qty` 在这个流程下**永远是 0**——账单挂在财务的 PO 上，
    不挂采购的 PO。所以对账必须是 PO ↔ PO，Billed Qty 只当校验旁证。

两条前提（不成立就算不出正确答案，脚本会拒绝出回写表）:
    甲 财务单里的每一件货，都是对这张采购单的交付。
    乙 采购单是订购的完整记录——行没被手工删改，也没有一部分货另开了单。

    2026-08-01 实测：P11382（七月「月度滚动 PO」实验的残骸）两条都不满足——
    部分订购行被删/被移走，且它身上同时留着两套流程的痕迹（一部分货靠挂在它
    自己身上的账单结掉、走 Billed Qty；一部分货走财务另建 PO）。这类数据无法
    自动对账，不是算法问题，是「采购到底订了什么」这个事实在 ERP 里已经残缺。

用法:
    python3 po_reconcile/po_reconcile.py <purchase.order.xlsx> \
        --buyer P11382 --finance P11416,P11665

产出（默认 output/YYYYMMDD/）:
    采购对账表-{采购PO}.xlsx      给人看：逐 SKU 的 订购/已收/未到 + 异常标记
    回写导入表-{采购PO}.xlsx      传 ERP：逐 order line 的新 Quantity（未到量，FIFO 摊回）
    results/原始订购量归档-{采购PO}.xlsx   回写前的订购量快照，防重复扣减
"""
import argparse
import os
import sys
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # 让 common/ 可导入
from common.xlsx import write_simple, unique_path  # noqa: E402

# ---- Odoo purchase.order 导出的列名 ----
F_PO = "Order Reference"
F_VENDOR = "Vendor"
F_ORIGIN = "Source Document"          # 理想关联字段；用户当前在 ERP 界面里填不了，可能整列缺失
F_BILL_STATUS = "Billing Status"
F_OL_ID = "Order Lines/ID"
F_OL_XID = "Order Lines/External ID"  # 回写必需；当前导出还没有，缺则降级
F_OL_QTY = "Order Lines/Quantity"
F_OL_BILLED = "Order Lines/Billed Qty"
F_OL_RECEIVED = "Order Lines/Received Qty"
F_OL_PRICE = "Order Lines/Unit Price"
F_OL_SKU = "Order Lines/Product/Internal Reference"
F_OL_NAME = "Order Lines/Product/Product"

SERVICE_SKU = "Service_Fee"           # 服务费行，不是商品，一律排除


def read_export(path):
    """读 Odoo 的**锯齿状** one2many 导出。

    致命陷阱：`Bills/Invoice lines/*` 与 `Order Lines/*` 是同一父记录下两条
    **互相独立**的序列，行与行之间没有任何对应关系。切块只能靠单头列前向填充，
    绝不能按行号对齐。
    """
    df = pd.read_excel(path)
    missing = [c for c in (F_PO, F_OL_ID, F_OL_QTY, F_OL_SKU) if c not in df.columns]
    if missing:
        raise SystemExit(f"导出缺少必需列: {missing}\n"
                         f"（{F_OL_SKU} 是商品身份，没有它认不出 order line 是什么货）")

    df["PO"] = df[F_PO].ffill()
    for col in (F_VENDOR, F_ORIGIN, F_BILL_STATUS):
        if col in df.columns:
            df[col] = df[col].ffill()

    ol = df[df[F_OL_ID].notna()].copy()
    # 无 product 的空壳行（P11382 里有 4 条 Quantity=0 / 单价 0 的残行）与服务费行一并剔除
    ol = ol[ol[F_OL_SKU].notna()]
    ol = ol[ol[F_OL_SKU] != SERVICE_SKU]
    ol[F_OL_ID] = ol[F_OL_ID].astype("int64")

    heads = df[df[F_PO].notna()].set_index(F_PO)
    return ol, heads


def split_roles(ol, heads, buyer_arg, finance_arg):
    """定采购单 / 财务单。手工指定优先；未指定时回退 Source Document 反查。"""
    all_po = list(dict.fromkeys(ol["PO"]))
    buyer = [p.strip() for p in buyer_arg.split(",") if p.strip()] if buyer_arg else []
    unknown = [p for p in buyer if p not in all_po]
    if unknown:
        raise SystemExit(f"导出里没有这些采购单: {unknown}\n可选: {all_po}")
    if not buyer:
        raise SystemExit(f"必须用 --buyer 指定采购记录单。导出里的 PO: {all_po}")

    if finance_arg:
        finance = [p.strip() for p in finance_arg.split(",") if p.strip()]
        unknown = [p for p in finance if p not in all_po]
        if unknown:
            raise SystemExit(f"导出里没有这些财务单: {unknown}\n可选: {all_po}")
        by_origin = {}
    else:
        # 回退：读 Source Document —— 财务建 PO 时若填了采购 PO 号，这里就能自动关联
        if F_ORIGIN not in heads.columns:
            raise SystemExit(f"未指定 --finance，且导出里没有 `{F_ORIGIN}` 列，无法自动关联。\n"
                             f"请补导出该列，或直接用 --finance 指定。")
        by_origin = {}
        for po, origin in heads[F_ORIGIN].dropna().items():
            for b in buyer:
                if b in str(origin):
                    by_origin.setdefault(b, []).append(po)
        finance = sorted({p for v in by_origin.values() for p in v})
        if not finance:
            raise SystemExit(f"`{F_ORIGIN}` 里没有任何一张单指向 {buyer}，无法自动关联。\n"
                             f"请用 --finance 手工指定。")

    overlap = set(buyer) & set(finance)
    if overlap:
        raise SystemExit(f"同一张 PO 不能既是采购单又是财务单: {sorted(overlap)}")
    return buyer, finance, by_origin


def load_archive(path):
    """读原始订购量归档 → {(PO, line_id): 原始订购量}。不存在则空。"""
    if not os.path.exists(path):
        return {}
    a = pd.read_excel(path)
    return {(r["采购PO"], int(r["OrderLineID"])): float(r["原始订购量"]) for _, r in a.iterrows()}


def reconcile(ol, buyer, finance, archive):
    """算未到量。返回 (SKU 级对账表, order line 级摊回明细)。

    订购量基准取 `max(当前 Quantity, 归档的原始订购量)`——回写过一次之后 ERP 里的
    Quantity 已经是余量，直接拿它当订购量会**重复扣减**。
    """
    buy = ol[ol["PO"].isin(buyer)].copy()
    fin = ol[ol["PO"].isin(finance)]

    buy["订购量"] = [
        max(float(q), archive.get((po, lid), 0.0))
        for po, lid, q in zip(buy["PO"], buy[F_OL_ID], buy[F_OL_QTY])
    ]
    buy["已回写过"] = [
        (po, lid) in archive and float(q) < archive[(po, lid)]
        for po, lid, q in zip(buy["PO"], buy[F_OL_ID], buy[F_OL_QTY])
    ]

    recv = fin.groupby(F_OL_SKU)[F_OL_QTY].sum()

    rows = []
    for (po, sku), g in buy.groupby(["PO", F_OL_SKU], sort=False):
        ordered = g["订购量"].sum()
        received = float(recv.get(sku, 0.0))
        rows.append({
            "采购PO": po,
            "SKU": sku,
            "商品": g[F_OL_NAME].iloc[0] if F_OL_NAME in g else "",
            "订购量": ordered,
            "财务单已收": received,
            "未到量": ordered - received,
            "行数": len(g),
            "采购单Billed": g[F_OL_BILLED].sum() if F_OL_BILLED in g else float("nan"),
            "采购单Received": g[F_OL_RECEIVED].sum() if F_OL_RECEIVED in g else float("nan"),
        })
    rec = pd.DataFrame(rows)

    # 异常标记：只报不改，交人工判断
    def flag(r):
        f = []
        if r["未到量"] < 0:
            f.append("收货多于订购")
        if pd.notna(r["采购单Billed"]) and r["采购单Billed"] > r["订购量"]:
            f.append("采购单开票>订购")
        if (pd.notna(r["采购单Billed"]) and pd.notna(r["采购单Received"])
                and r["采购单Billed"] != r["采购单Received"]):
            f.append("开票≠收货")
        return " / ".join(f)

    rec["异常"] = rec.apply(flag, axis=1) if len(rec) else pd.Series(dtype=str)
    rec = rec.sort_values(["采购PO", "未到量"], ascending=[True, False])

    # ---- FIFO 摊回 order line ----
    # 摊的是**已收量**不是未到量：先满足早下的订单（采购手工就是这么做的——
    # 「1 号订 5、5 号订 3、到货 5」应该是 1 号那笔结清、5 号那笔全欠，
    # 而不是把欠的 3 件记在 1 号头上）。
    recv_left = {(r["采购PO"], r["SKU"]): max(r["订购量"] - max(r["未到量"], 0.0), 0.0)
                 for _, r in rec.iterrows()}
    lines = []
    for (po, sku), g in buy.sort_values(F_OL_ID).groupby(["PO", F_OL_SKU], sort=False):
        left = recv_left[(po, sku)]
        for _, r in g.iterrows():
            alloc = min(left, r["订购量"])      # 这条行分到多少已收
            left -= alloc
            take = r["订购量"] - alloc          # 剩下的就是这条行的未到量
            lines.append({
                "采购PO": po,
                "OrderLineID": int(r[F_OL_ID]),
                "OrderLineXID": r[F_OL_XID] if F_OL_XID in buy.columns else "",
                "SKU": sku,
                "商品": r[F_OL_NAME] if F_OL_NAME in buy else "",
                "原始订购量": r["订购量"],
                "ERP当前数量": float(r[F_OL_QTY]),
                "新数量": take,
                "已回写过": r["已回写过"],
            })
    return rec, pd.DataFrame(lines)


def main():
    ap = argparse.ArgumentParser(description="采购对账：采购 PO ↔ 财务 PO，算未到货量")
    ap.add_argument("export", help="Odoo purchase.order 导出（含采购单与财务单）")
    ap.add_argument("--buyer", required=True, help="采购记录单号，逗号分隔，如 P11382")
    ap.add_argument("--finance", help="财务实收单号，逗号分隔；不给则读 Source Document 自动关联")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    ap.add_argument("--archive-dir", default="results", help="原始订购量归档目录（默认 results）")
    ap.add_argument("--no-writeback", action="store_true", help="只出对账表，不出回写导入表")
    args = ap.parse_args()

    ol, heads = read_export(args.export)
    buyer, finance, by_origin = split_roles(ol, heads, args.buyer, args.finance)

    outdir = args.outdir or os.path.join("output", f"{date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)
    os.makedirs(args.archive_dir, exist_ok=True)

    print(f"采购记录单: {', '.join(buyer)}")
    print(f"财务实收单: {', '.join(finance)}" + ("（读 Source Document 自动关联）" if by_origin else "（手工指定）"))
    if F_ORIGIN in heads.columns:
        no_origin = [p for p in finance if pd.isna(heads[F_ORIGIN].get(p))]
        if no_origin:
            print(f"  · {len(no_origin)} 张财务单没填 Source Document: {', '.join(no_origin)}"
                  f" —— 填上之后就不必每次手工指定")
    if F_OL_XID not in ol.columns:
        print(f"  ⚠ 导出缺 `{F_OL_XID}`，回写导入表只能带数据库 ID —— 见文末说明")

    tag = "+".join(buyer)
    apath = os.path.join(args.archive_dir, f"原始订购量归档-{tag}.xlsx")
    archive = load_archive(apath)
    if archive:
        print(f"  · 读到归档 {len(archive)} 条原始订购量（{apath}）")

    rec, lines = reconcile(ol, buyer, finance, archive)
    if rec.empty:
        raise SystemExit("采购单里没有可对账的商品行")

    p1, _ = write_simple(rec.drop(columns=["行数"]), outdir, f"采购对账表-{tag}.xlsx",
                         left_cols={"商品", "异常"})
    print(f"\n✅ 对账表: {p1}")

    pend = rec[rec["未到量"] > 0]
    print(f"   {len(rec)} 个 SKU，其中 {len(pend)} 个未到齐，共 {pend['未到量'].sum():.0f} 件未到")
    bad = rec[rec["异常"] != ""]
    if len(bad):
        print(f"   ⚠ {len(bad)} 个 SKU 有异常，脚本不自动处理，请人工判断：")
        for _, r in bad.iterrows():
            print(f"     {r['SKU']}: 订购 {r['订购量']:.0f} / 财务单已收 {r['财务单已收']:.0f}"
                  f" / 采购单Billed {r['采购单Billed']:.0f} → {r['异常']}")

    # ---- 前提校验：不成立就拒绝出回写表 ----
    # 「收的比订的多」在满足前提的数据上不可能发生。出现了只有两种可能：
    # 关联错了（把不属于这张采购单的财务单算了进来），或采购单本身不完整
    # （行被手工删过/货另开了单）。此时回写只会把错误写进 ERP，必须拦住。
    neg = rec[rec["未到量"] < 0]
    if len(neg):
        print(f"\n⛔ 拒绝生成回写导入表：{len(neg)} 个 SKU 的「财务单已收」超过订购量。")
        for _, r in neg.iterrows():
            print(f"     {r['SKU']}: 订购 {r['订购量']:.0f} < 已收 {r['财务单已收']:.0f}"
                  f"（差 {r['财务单已收'] - r['订购量']:.0f}）")
        print("   这说明本脚本的两条前提至少破了一条：")
        print("     甲 财务单里的每一件货都是对这张采购单的交付")
        print("     乙 采购单是订购的完整记录（行没被手工删改、货没另开单）")
        print("   请先核对 --finance 指定的单是否都属于这张采购单；对账表已产出，可据此排查。")
        sys.exit(2)

    if args.no_writeback:
        return

    # ---- 归档原始订购量（回写会覆盖 ERP 里的 Quantity，覆盖前必须留底）----
    new_arch = pd.DataFrame([
        {"采购PO": r["采购PO"], "OrderLineID": r["OrderLineID"], "SKU": r["SKU"],
         "原始订购量": r["原始订购量"], "首次归档日": f"{date.today():%Y-%m-%d}"}
        for _, r in lines.iterrows() if (r["采购PO"], r["OrderLineID"]) not in archive
    ])
    if len(new_arch):
        if os.path.exists(apath):
            new_arch = pd.concat([pd.read_excel(apath), new_arch], ignore_index=True)
        new_arch.to_excel(apath, index=False)
        print(f"\n📦 原始订购量已归档: {apath}（{len(new_arch)} 条）")

    idcol = "OrderLineXID" if F_OL_XID in ol.columns else "OrderLineID"
    imp = lines[[idcol, "SKU", "商品", "原始订购量", "ERP当前数量", "新数量"]].copy()
    imp = imp.rename(columns={idcol: "id" if idcol == "OrderLineXID" else ".id"})
    p2, _ = write_simple(imp, outdir, f"回写导入表-{tag}.xlsx", left_cols={"商品"})
    print(f"📤 回写导入表: {p2}")
    zeroed = (lines["新数量"] == 0).sum()
    print(f"   {len(lines)} 条 order line，其中 {zeroed} 条已到齐将被改成 0（行保留）")
    if lines["已回写过"].any():
        print(f"   · {lines['已回写过'].sum()} 条检测到已回写过，订购量已按归档还原，未重复扣减")
    print("\n⚠ 首次上传务必先拿 1 条行试，确认 ERP 更新的是既有行而不是新建行，再传全量。")


if __name__ == "__main__":
    main()
