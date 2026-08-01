#!/usr/bin/env python3
"""sales_insight 的构造数据测试。

样例结构照真实导出复刻，**包括那两个坑**：
  · 销售透视表的商品名前面有缩进空格（`    [SKU] 名称`）
  · product.product 的 Internal Reference 前面带 `\\t`
不复刻这两点，测试就测不到真正会出问题的地方。

    python3 sales_insight/test_sales_insight.py
"""
import os
import subprocess
import sys
import tempfile

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "sales_insight.py")

# SKU        年销量  下单次数  销售额   主数据?  ERP在手  运营安全库存  运营在手
DATA = [
    ("AAA_111", 600, 600, 6000.0, True,  5,    40,   5),      # 有人工值，低于安全库存
    ("BBB_222", 300, 250, 3000.0, True,  99,   20,   99),     # 有人工值，库存充足
    ("CCC_333", 150, 150, 1500.0, True,  1,    None, None),   # 无人工值 → 候选值，ERP 有库存
    ("DDD_444",  60,  60,  600.0, True,  None, None, None),   # 无人工值、无库存
    ("EEE_555",  30,  30,  300.0, False, None, 10,   0),      # 有人工值但主数据里没有 → 不能回写
]
PERIOD_WEEKS = 30.0
COVER_WEEKS = 2.0


def make_inputs(d, with_erp_onhand):
    sales = os.path.join(d, "sales.xlsx")
    rows = [[None, "Total", None, None], [None, "Sales", None, None],
            [None, "Untaxed Total", "# of Lines", "Qty Ordered"],
            ["Total", sum(r[3] for r in DATA), sum(r[2] for r in DATA), sum(r[1] for r in DATA)]]
    for sku, qty, lines, amt, *_ in DATA:
        rows.append([f"    [{sku}] 样例商品 {sku}", amt, lines, qty])   # ← 缩进空格，照真实导出
    pd.DataFrame(rows).to_excel(sales, index=False, header=False)

    prods = os.path.join(d, "products.xlsx")
    pr = [{"ID": f"__export__.product_product_{i}_abc",
           "Internal Reference": f"\t{r[0]}",                          # ← \t 前缀，照真实导出
           "Barcode": f"400000000{i}", "Name": f"样例商品 {r[0]}",
           "VO Shop Name": "TKOF_SHOP1_VO",
           **({"Quantity On Hand": r[5]} if with_erp_onhand else {})}
          for i, r in enumerate(DATA) if r[4]]
    pd.DataFrame(pr).to_excel(prods, index=False)

    safety = os.path.join(d, "safety.xlsx")
    sr = [{"商品SKU": r[0], "W29": 10, "W30": 12, "总销量": r[1],
           "8.01在手库存": r[7], "安全库存": r[6], "备注": None}      # ← 列名带日期前缀
          for r in DATA if r[6] is not None]
    pd.DataFrame(sr).to_excel(safety, index=False)
    return sales, prods, safety


fails = []


def check(name, cond, detail=""):
    print(("  ✅ " if cond else "  ❌ ") + name + (f"  {detail}" if detail else ""))
    if not cond:
        fails.append(name)


def run(d, sales, prods, safety, out, *extra):
    r = subprocess.run([sys.executable, SCRIPT, sales, "--products", prods, "--safety", safety,
                        "--weeks", str(PERIOD_WEEKS), "--cover-weeks", str(COVER_WEEKS),
                        "-o", out, *extra], capture_output=True, text=True)
    if r.returncode:
        print(r.stdout, r.stderr)
    return r


def main():
    d = tempfile.mkdtemp(prefix="sales_insight_test_")

    print("【1】ERP 带 Quantity On Hand（目标形态）")
    sales, prods, safety = make_inputs(d, with_erp_onhand=True)
    r = run(d, sales, prods, safety, f"{d}/o1")
    check("退出码 0", r.returncode == 0, r.stderr[-200:] if r.returncode else "")
    check("识别到 ERP 在手库存列", "带在手库存列" in r.stdout)
    rank = pd.read_excel(f"{d}/o1/销量排名.xlsx").set_index("SKU")

    check("SKU 解析成功（缩进空格没挡住）", len(rank) == len(DATA), f"得到 {len(rank)} 行")
    check("Internal Reference 的 \\t 已 strip（4 个 SKU 连上主数据）",
          rank["商品名称"].notna().sum() == 4, f"得到 {rank['商品名称'].notna().sum()}")
    check("销量降序 + 排名正确", list(rank["销量排名"]) == [1, 2, 3, 4, 5])
    check("AAA 每单件数 = 600/600 = 1.0", rank.loc["AAA_111", "每单件数"] == 1.0)
    check("BBB 每单件数 = 300/250 = 1.2", rank.loc["BBB_222", "每单件数"] == 1.2)
    check("累计占比末行 = 1.0", abs(rank["累计占比"].iloc[-1] - 1.0) < 1e-6)

    # 周均 = 年销量/30；推算 = 周均×2
    check("CCC 周均 = 150/30 = 5.0", rank.loc["CCC_333", "周均销量"] == 5.0)
    check("CCC 推算安全库存 = 5.0×2 = 10", rank.loc["CCC_333", "安全库存"] == 10)
    check("CCC 来源标为「脚本推算」", rank.loc["CCC_333", "安全库存来源"] == "脚本推算")
    check("AAA 用人工值 40 而非推算值 40/30×2", rank.loc["AAA_111", "安全库存"] == 40)
    check("AAA 来源标为「运营人工」", rank.loc["AAA_111", "安全库存来源"] == "运营人工")

    check("在手库存优先取 ERP（CCC=1，运营表里没有这个 SKU）",
          rank.loc["CCC_333", "在手库存"] == 1 and rank.loc["CCC_333", "库存来源"] == "ERP")
    check("DDD 无任何库存 → 库存来源标「缺」", rank.loc["DDD_444", "库存来源"] == "缺")

    alert = pd.read_excel(f"{d}/o1/补货提醒.xlsx")
    # AAA 5<40 缺35；BBB 99>20 不报；CCC 1<10 缺9；DDD 无库存不报；EEE 0<10 缺10
    check("补货提醒只含缺口>0 的", set(alert["SKU"]) == {"AAA_111", "CCC_333", "EEE_555"},
          f"得到 {sorted(alert['SKU'])}")
    check("按缺口降序", list(alert["缺口"]) == sorted(alert["缺口"], reverse=True))
    check("AAA 缺口 = 40-5 = 35", alert.set_index("SKU").loc["AAA_111", "缺口"] == 35)
    check("无库存的 DDD 未参与提醒", "DDD_444" not in set(alert["SKU"]))

    wb = pd.read_excel(f"{d}/o1/安全库存回写表.xlsx")
    check("回写表只含运营人工值且有 ERP ID（AAA/BBB，EEE 无主数据被排除）",
          set(wb["SKU"]) == {"AAA_111", "BBB_222"}, f"得到 {sorted(wb['SKU'])}")
    check("回写表报出「有人工值但无 ERP ID」的 EEE", "EEE_555" in r.stdout)
    check("回写表列名为 id / Safety Stock", list(wb.columns[:2]) == ["id", "Safety Stock"])
    check("回写值 = 人工值", wb.set_index("SKU").loc["AAA_111", "Safety Stock"] == 40)

    cand = pd.read_excel(f"{d}/o1/安全库存候选值.xlsx")
    check("候选值只含推算的", set(cand["SKU"]) == {"CCC_333", "DDD_444"}, f"得到 {sorted(cand['SKU'])}")
    check("候选值与回写表无交集（推算值绝不自动写回）", not (set(cand["SKU"]) & set(wb["SKU"])))

    print("\n【2】ERP 不带 On Hand —— 应降级到运营表并明确告警")
    sales2, prods2, safety2 = make_inputs(f"{d}", with_erp_onhand=False)
    r2 = run(d, sales2, prods2, safety2, f"{d}/o2")
    check("退出码 0", r2.returncode == 0)
    check("告警「没有在手库存列」", "没有在手库存列" in r2.stdout)
    rank2 = pd.read_excel(f"{d}/o2/销量排名.xlsx").set_index("SKU")
    check("AAA 退回运营表的在手库存 5", rank2.loc["AAA_111", "在手库存"] == 5
          and rank2.loc["AAA_111", "库存来源"] == "运营表")
    check("CCC 运营表里没有 → 库存来源「缺」", rank2.loc["CCC_333", "库存来源"] == "缺")

    print("\n【3】--weeks 是必需参数（销售导出里没有日期，不许瞎猜）")
    r3 = subprocess.run([sys.executable, SCRIPT, sales, "--products", prods, "-o", f"{d}/o3"],
                        capture_output=True, text=True)
    check("缺 --weeks 时报错退出", r3.returncode != 0 and "weeks" in (r3.stderr + r3.stdout))

    print("\n" + ("全部通过 🎉" if not fails else f"❌ {len(fails)} 项失败: {fails}"))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
