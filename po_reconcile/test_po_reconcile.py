#!/usr/bin/env python3
"""po_reconcile 的构造数据测试。

为什么用构造数据：手上唯一的真实数据（P11382）是七月失败实验的残骸，
订购行被手工删改过、且混着两套流程的痕迹——拿它测脚本没有意义，
分不清是脚本错还是数据错（2026-08-01 用户拍板）。

这里造的样例**结构真、内容假**：完全照 Odoo purchase.order 的锯齿状
one2many 导出格式（单头一行、order line 往下铺、多张 PO 纵向堆叠），
但满足脚本的两条前提，期望值可以手算。

    python3 po_reconcile/test_po_reconcile.py
"""
import os
import subprocess
import sys
import tempfile

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "po_reconcile.py")

COLS = ["Order Reference", "Vendor", "Source Document", "Billing Status",
        "Order Lines/ID", "Order Lines/External ID", "Order Lines/Quantity",
        "Order Lines/Billed Qty", "Order Lines/Received Qty", "Order Lines/Unit Price",
        "Order Lines/Product/Internal Reference", "Order Lines/Product/Product"]


def make_export(path, pos, qty_override=None):
    """pos: [(PO号, Source Document, [(line_id, SKU, 数量), ...]), ...] → 锯齿状 xlsx。"""
    rows = []
    for po, origin, lines in pos:
        for i, (lid, sku, qty) in enumerate(lines):
            q = (qty_override or {}).get(lid, qty)
            rows.append({
                # 单头字段只在该 PO 的第一行出现，其余留空——这正是 Odoo 导出的样子
                "Order Reference": po if i == 0 else None,
                "Vendor": "样例供应商 GmbH" if i == 0 else None,
                "Source Document": origin if i == 0 else None,
                "Billing Status": "Waiting Bills" if i == 0 else None,
                "Order Lines/ID": lid,
                "Order Lines/External ID": f"__export__.purchase_order_line_{lid}",
                "Order Lines/Quantity": q,
                "Order Lines/Billed Qty": 0.0,
                "Order Lines/Received Qty": 0.0,
                "Order Lines/Unit Price": 10.0,
                "Order Lines/Product/Internal Reference": sku,
                "Order Lines/Product/Product": None if sku is None else f"[{sku}] 样例商品 {sku}",
            })
    pd.DataFrame(rows, columns=COLS).to_excel(path, index=False)


# 采购单 P90001：SKU_A 分两次订（5 件 + 3 件，正是用户举的 1号/5号 那个例子）
#                SKU_B 10 件、SKU_C 4 件，外加一条服务费行和一条无商品的空壳行
BUYER = ("P90001", None, [
    (1001, "SKU_A", 5.0),
    (1002, "SKU_A", 3.0),
    (1003, "SKU_B", 10.0),
    (1004, "SKU_C", 4.0),
    (1005, "Service_Fee", 1.0),      # 服务费，应被剔除
    (1006, None, 0.0),               # 无商品空壳行，应被剔除
])
# 两张财务单，都在 Source Document 里指回 P90001
FIN1 = ("P90002", "P90001", [(2001, "SKU_A", 3.0), (2002, "SKU_B", 10.0)])
FIN2 = ("P90003", "P90001", [(3001, "SKU_A", 2.0), (3002, "SKU_C", 1.0)])

# ---- 手算期望 ----
# 已收: SKU_A=3+2=5  SKU_B=10  SKU_C=1
# 未到: SKU_A=8-5=3  SKU_B=10-10=0  SKU_C=4-1=3
# FIFO 摊回（摊已收、先满足早下的单）:
#   SKU_A 已收 5 → 1001 分 5 → 新数量 0；1002 分 0 → 新数量 3
#   SKU_B 已收 10 → 1003 新数量 0
#   SKU_C 已收 1  → 1004 新数量 3
EXPECT_REC = {"SKU_A": (8.0, 5.0, 3.0), "SKU_B": (10.0, 10.0, 0.0), "SKU_C": (4.0, 1.0, 3.0)}
EXPECT_LINES = {1001: 0.0, 1002: 3.0, 1003: 0.0, 1004: 3.0}

fails = []


def check(name, cond, detail=""):
    print(("  ✅ " if cond else "  ❌ ") + name + (f"  {detail}" if detail else ""))
    if not cond:
        fails.append(name)


def run(export, outdir, archive, *extra, expect_rc=0):
    r = subprocess.run([sys.executable, SCRIPT, export, "--buyer", "P90001",
                        "-o", outdir, "--archive-dir", archive, *extra],
                       capture_output=True, text=True)
    if r.returncode != expect_rc:
        print(r.stdout, r.stderr)
    return r


def main():
    tmp = tempfile.mkdtemp(prefix="po_reconcile_test_")
    exp = os.path.join(tmp, "export.xlsx")
    make_export(exp, [BUYER, FIN1, FIN2])

    print("【1】手工指定财务单 —— 对账数字与 FIFO 摊回")
    r = run(exp, f"{tmp}/o1", f"{tmp}/a1", "--finance", "P90002,P90003")
    check("退出码 0", r.returncode == 0, r.stderr[-200:] if r.returncode else "")
    rec = pd.read_excel(f"{tmp}/o1/采购对账表-P90001.xlsx").set_index("SKU")
    check("服务费行与空壳行已剔除", set(rec.index) == set(EXPECT_REC), f"得到 {sorted(rec.index)}")
    for sku, (o, rcv, out) in EXPECT_REC.items():
        got = (rec.loc[sku, "订购量"], rec.loc[sku, "财务单已收"], rec.loc[sku, "未到量"])
        check(f"{sku} 订购/已收/未到 = {o:.0f}/{rcv:.0f}/{out:.0f}", got == (o, rcv, out), f"得到 {got}")
    lines = pd.read_excel(f"{tmp}/o1/回写导入表-P90001.xlsx")
    got = dict(zip(lines["id"].str.extract(r"_(\d+)$")[0].astype(int), lines["新数量"]))
    check("FIFO 摊回：早下的单先结清（1001→0, 1002→3）", got == EXPECT_LINES, f"得到 {got}")
    check("回写表用行级 External ID 列名 id", "id" in lines.columns)

    print("\n【2】不传 --finance —— 靠 Source Document 自动关联，结果应完全一致")
    r = run(exp, f"{tmp}/o2", f"{tmp}/a2")
    check("退出码 0", r.returncode == 0)
    rec2 = pd.read_excel(f"{tmp}/o2/采购对账表-P90001.xlsx").set_index("SKU")
    check("与手工指定结果一致", rec2["未到量"].to_dict() == rec["未到量"].to_dict())

    print("\n【3】防重复扣减 —— 模拟回写已上传 ERP 后再跑一次")
    exp2 = os.path.join(tmp, "export_after.xlsx")
    make_export(exp2, [BUYER, FIN1, FIN2], qty_override=EXPECT_LINES)   # ERP 里数量已变成余量
    r = run(exp2, f"{tmp}/o3", f"{tmp}/a1", "--finance", "P90002,P90003")  # 复用【1】的归档
    check("退出码 0", r.returncode == 0)
    rec3 = pd.read_excel(f"{tmp}/o3/采购对账表-P90001.xlsx").set_index("SKU")
    check("订购量按归档还原，未被余量顶替", rec3["订购量"].to_dict() == rec["订购量"].to_dict(),
          f"得到 {rec3['订购量'].to_dict()}")
    check("未到量与首次相同，没有重复扣减", rec3["未到量"].to_dict() == rec["未到量"].to_dict(),
          f"得到 {rec3['未到量'].to_dict()}")
    check("终端报出「已回写过」", "已回写过" in r.stdout or "未重复扣减" in r.stdout)

    print("\n【4】前提被破坏（收货多于订购）—— 必须拒绝出回写表")
    bad = ("P90004", "P90001", [(4001, "SKU_B", 99.0)])   # SKU_B 只订了 10，却收 10+99
    exp3 = os.path.join(tmp, "export_bad.xlsx")
    make_export(exp3, [BUYER, FIN1, FIN2, bad])
    r = run(exp3, f"{tmp}/o4", f"{tmp}/a4", "--finance", "P90002,P90003,P90004", expect_rc=2)
    check("以退出码 2 中止", r.returncode == 2, f"得到 {r.returncode}")
    check("终端说明拒绝原因", "拒绝生成回写导入表" in r.stdout)
    check("对账表仍然产出（供排查）", os.path.exists(f"{tmp}/o4/采购对账表-P90001.xlsx"))
    check("回写表未产出", not os.path.exists(f"{tmp}/o4/回写导入表-P90001.xlsx"))

    print("\n【5】护栏：采购单与财务单不能是同一张")
    r = run(exp, f"{tmp}/o5", f"{tmp}/a5", "--finance", "P90001", expect_rc=1)
    check("被拒绝", r.returncode != 0 and "不能既是" in (r.stdout + r.stderr))

    print("\n" + ("全部通过 🎉" if not fails else f"❌ {len(fails)} 项失败: {fails}"))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
