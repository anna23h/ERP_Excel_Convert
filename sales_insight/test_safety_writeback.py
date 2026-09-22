#!/usr/bin/env python3
"""safety_writeback 的构造数据测试。

样例结构照真实文件复刻，**包括那几个坑**：
  · product.product 的 Internal Reference 前面带 `\\t`
  · 运营的值列名带口径前缀（`15天安全库存`），不是干净的 `安全库存`
  · 表里同时存在第二个含「安全库存」的列（真实文件里就有 `9.18统计实库` 那种邻居列，
    换个月份很容易多出 `30天安全库存`）——必须逼出「不替人猜口径」的报错

测的重点不是「能不能出表」，而是**该拒绝的时候有没有拒绝**：回写表是要导进 ERP
主数据的，多出一行、少了一行、值歪一位都比报错难发现得多。

    python3 sales_insight/test_safety_writeback.py
"""
import os
import subprocess
import sys
import tempfile

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "safety_writeback.py")

# SKU        15天安全库存  在主数据里?  ERP 现有安全库存  ERP 在手
DATA = [
    ("AAA_111", 120, True,  100,  5),     # 有变动 100→120
    ("BBB_222",  30, True,   30,  99),    # 与现值相同 → 无变化
    ("CCC_333x2", 10, True,  None, 1),    # ERP 里还没配 → 新写入；x2 组合装
    ("DDD_444",  45, False,  None, None),  # 主数据里没有 → 不写，进未匹配表
]
VAL_COL = "15天安全库存"
IDS = {sku: f"__export__.product_product_{i}_abc" for i, (sku, *_) in enumerate(DATA)}
MATCHED = [r[0] for r in DATA if r[2]]

fails = []


def check(name, cond, detail=""):
    print(("  ✅ " if cond else "  ❌ ") + name + (f"  {detail}" if detail else ""))
    if not cond:
        fails.append(name)


def make_safety(path, rows=None, extra_col=False):
    """运营那份。列名照真实文件：`商品SKU` + 周列 + 统计实库 + `15天安全库存`。"""
    rows = DATA if rows is None else rows
    d = pd.DataFrame({
        "商品SKU": [r[0] for r in rows],
        "W36 2026": [10] * len(rows),
        "W37 2026": [12] * len(rows),
        "9.18统计实库": [1] * len(rows),
        VAL_COL: [r[1] for r in rows],
        "采购需求需求": [0] * len(rows),
    })
    if extra_col:                       # 第二个含「安全库存」的列
        d["30天安全库存"] = [r[1] * 2 for r in rows]
    d.to_excel(path, index=False)
    return path


def make_products(path, dup_sku=None, with_erp_cols=True):
    """product.product 导出。`Internal Reference` 前面带 \\t，与真实导出一致。"""
    rows = [r for r in DATA if r[2]]
    d = pd.DataFrame({
        "External ID": [IDS[r[0]] for r in rows],
        "Internal Reference": ["\t" + r[0] for r in rows],
        "Name": [f"商品 {r[0]}" for r in rows],
        "FS": ["P"] * len(rows),
    })
    if with_erp_cols:
        d["Safety Stock"] = [r[3] for r in rows]
        d["Quantity On Hand"] = [r[4] for r in rows]
        d["Supply Remark"] = ["20260715:近3月3单 最低9.9@某供应商"] * len(rows)
    if dup_sku:                         # 同一 SKU 建了两个产品档
        d = pd.concat([d, d[d["Internal Reference"] == "\t" + dup_sku]
                       .assign(**{"External ID": "__export__.product_product_999_dup"})],
                      ignore_index=True)
    d.to_excel(path, index=False)
    return path


def run(safety, prods, out, *extra):
    r = subprocess.run([sys.executable, SCRIPT, safety, "--products", prods, "-o", out,
                        *extra], capture_output=True, text=True, encoding="utf-8")
    return r


def main():
    d = tempfile.mkdtemp(prefix="safety_wb_test_")
    prods = make_products(os.path.join(d, "prods.xlsx"))

    print("【1】正常跑（目标形态）")
    safety = make_safety(os.path.join(d, "safety.xlsx"))
    r = run(safety, prods, f"{d}/o1")
    check("退出码 0", r.returncode == 0, (r.stdout + r.stderr)[-300:] if r.returncode else "")
    wb = pd.read_excel(f"{d}/o1/安全库存回写表.xlsx")
    check("认出带口径前缀的值列", f"`{VAL_COL}`" in r.stdout)
    check("回写表三列、不带 Supply Remark（Odoo 才不会碰这个字段）",
          list(wb.columns) == ["id", "SKU(勿导入)", "Safety Stock"], f"得到 {list(wb.columns)}")
    check("只含对得上主数据的 SKU（DDD 被排除）",
          set(wb["SKU(勿导入)"]) == set(MATCHED), f"得到 {sorted(wb['SKU(勿导入)'])}")
    check("Internal Reference 的 \\t 已 strip（不 strip 连接率为 0）",
          not wb["SKU(勿导入)"].astype(str).str.contains("\t").any())
    check("id 取 External ID 而非数据库整数",
          wb["id"].astype(str).str.startswith("__export__").all())
    check("值逐行等于源表", wb.set_index("SKU(勿导入)")["Safety Stock"].to_dict()
          == {r[0]: r[1] for r in DATA if r[2]})

    lost = pd.read_excel(f"{d}/o1/未匹配SKU.xlsx")
    check("未匹配 SKU 单独出表", set(lost["SKU"]) == {"DDD_444"}, f"得到 {list(lost['SKU'])}")
    check("报错文案指向产品主数据的导出筛选条件", "can be sold" in r.stdout)

    chk = pd.read_excel(f"{d}/o1/回写核对表.xlsx").set_index("SKU")
    check("核对表带 ERP 现值与变化量", chk.loc["AAA_111", "ERP现有安全库存"] == 100
          and chk.loc["AAA_111", "变化"] == 20)
    check("ERP 里还没配的行变化为空", pd.isna(chk.loc["CCC_333x2", "变化"]))
    check("缺口 = 将写入 − 在手", chk.loc["AAA_111", "缺口"] == 115)
    check("摘要区分 新写入/无变化/有变动", "1 个 ERP 里还没配" in r.stdout
          and "1 个与现值相同" in r.stdout and "1 个有变动" in r.stdout,
          [l for l in r.stdout.splitlines() if "无变化" in l])

    print("\n【2】试水：单条表 + 导入前快照，全量照出不误")
    r = run(safety, prods, f"{d}/o2", "--test-sku", "AAA_111")
    one = pd.read_excel(f"{d}/o2/安全库存回写表-试AAA_111.xlsx")
    check("试水表只有一条", len(one) == 1 and one.iloc[0]["SKU(勿导入)"] == "AAA_111")
    check("导入前快照已出", os.path.exists(f"{d}/o2/导入前快照-AAA_111.xlsx"))
    check("全量表照样出（不必为拿它再跑一遍）",
          len(pd.read_excel(f"{d}/o2/安全库存回写表.xlsx")) == len(MATCHED))
    check("报出 现值→写入", "100 → 写入 120" in r.stdout,
          [l for l in r.stdout.splitlines() if "写入" in l][:2])
    r2 = run(safety, prods, f"{d}/o2b", "--test-sku", "ZZZ_999")
    check("试水 SKU 不在表里 → 报错而非静默出空表", r2.returncode != 0
          and "不在安全库存表里" in (r2.stdout + r2.stderr))

    print("\n【3】该拒绝的都拒绝了（一条都不产出）")
    cases = [
        ("值为 0（导进去 = 清零）", [("AAA_111", 0, True, 100, 5)], "值为 0", ()),
        ("值为空", [("AAA_111", None, True, 100, 5)], "值为空", ()),
        ("值是负数", [("AAA_111", -5, True, 100, 5)], "负数", ()),
        ("值不是整数", [("AAA_111", 12.5, True, 100, 5)], "不是整数", ()),
    ]
    for name, rows, want, extra in cases:
        p = make_safety(os.path.join(d, f"bad_{want}.xlsx"), rows)
        rr = run(p, prods, f"{d}/bad", *extra)
        out = rr.stdout + rr.stderr
        check(name, rr.returncode != 0 and want in out, out.strip().splitlines()[-1:])

    p = make_safety(os.path.join(d, "zero_ok.xlsx"), [("AAA_111", 0, True, 100, 5)])
    rr = run(p, prods, f"{d}/o_zero", "--allow-zero")
    check("--allow-zero 显式放行 0", rr.returncode == 0
          and pd.read_excel(f"{d}/o_zero/安全库存回写表.xlsx").iloc[0]["Safety Stock"] == 0)

    p = make_safety(os.path.join(d, "dup.xlsx"),
                    [("AAA_111", 120, True, 100, 5), ("AAA_111", 80, True, 100, 5)])
    rr = run(p, prods, f"{d}/bad")
    check("安全库存表里重复 SKU → 报错并列出两个值", rr.returncode != 0
          and "重复 SKU" in (rr.stdout + rr.stderr) and "120/80" in (rr.stdout + rr.stderr))

    p = make_safety(os.path.join(d, "twocol.xlsx"), extra_col=True)
    rr = run(p, prods, f"{d}/bad")
    check("两个含「安全库存」的列 → 报错要求 --col（不替人猜口径）",
          rr.returncode != 0 and "--col" in (rr.stdout + rr.stderr))
    rr = run(p, prods, f"{d}/o_col", "--col", "30天安全库存")
    check("--col 指定后按那一列写",
          rr.returncode == 0
          and pd.read_excel(f"{d}/o_col/安全库存回写表.xlsx")
                .set_index("SKU(勿导入)").loc["AAA_111", "Safety Stock"] == 240)
    rr = run(safety, prods, f"{d}/bad", "--col", "不存在的列")
    check("--col 给了不存在的列 → 报错并列出现有列",
          rr.returncode != 0 and "现有列" in (rr.stdout + rr.stderr))

    dp = make_products(os.path.join(d, "prods_dup.xlsx"), dup_sku="AAA_111")
    rr = run(safety, dp, f"{d}/bad")
    check("产品主数据里同一 SKU 多行 → 报错（不静默取第一条写歪）",
          rr.returncode != 0 and "各占多行" in (rr.stdout + rr.stderr))

    np_ = make_products(os.path.join(d, "prods_noid.xlsx"))
    pd.read_excel(np_).drop(columns=["External ID"]).to_excel(np_, index=False)
    rr = run(safety, np_, f"{d}/bad")
    check("产品主数据缺 External ID → 报错", rr.returncode != 0
          and "External ID" in (rr.stdout + rr.stderr))

    print("\n【4】产品主数据没勾 Safety Stock / On Hand 时只降级、不中断")
    lean = make_products(os.path.join(d, "prods_lean.xlsx"), with_erp_cols=False)
    rr = run(safety, lean, f"{d}/o4")
    check("照样出回写表", rr.returncode == 0
          and len(pd.read_excel(f"{d}/o4/安全库存回写表.xlsx")) == len(MATCHED))
    check("明确告警看不到 ERP 现值", "没有 `Safety Stock` 列" in rr.stdout)

    print("\n" + ("❌ 失败 %d 项：%s" % (len(fails), fails) if fails else "✅ 全部通过"))
    print(f"（产出留在 {d}）")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
