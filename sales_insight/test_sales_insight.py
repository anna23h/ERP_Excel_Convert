#!/usr/bin/env python3
"""sales_insight 的构造数据测试。

样例结构照真实导出复刻，**包括那几个坑**：
  · 销售透视表的商品名前面有缩进空格（`    [SKU] 名称`）
  · product.product 的 Internal Reference 前面带 `\\t`
  · 按周分组的导出里，某周无销售是**空格**而不是 0
不复刻这些，测试就测不到真正会出问题的地方。

    python3 sales_insight/test_sales_insight.py
"""
import os
import subprocess
import sys
import tempfile
from datetime import date

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "sales_insight.py")

# SKU        期间销量 下单次数 销售额  主数据? ERP在手 运营安全库存 运营在手  运营备注
DATA = [
    ("AAA_111", 600, 600, 6000.0, True,  5,    40,   5,    "已下架，清完为止"),
    ("BBB_222", 300, 250, 3000.0, True,  99,   20,   99,   None),
    ("CCC_333", 150, 150, 1500.0, True,  1,    None, None, None),
    ("DDD_444",  60,  60,  600.0, True,  None, None, None, None),
    ("EEE_555",  30,  30,  300.0, False, None, 10,   0,    "主数据里没有，回写不了"),
    # 销量低到推算值四舍五入成 0（5/30 周 ×2 = 0.4 → 0），验候选值把这类剔除
    ("FFF_666",   5,   5,   50.0, True,  None, None, None, None),
]
PERIOD_WEEKS = 30.0
COVER_WEEKS = 2.0
TODAY = f"{date.today():%Y%m%d}"

# ERP 里已有的 Supply Remark。AAA 那条是 fs_writeback 写的画像段——
# 本脚本回写时必须原样保住它，否则供应商画像就被抹了。
ERP_REMARK = {
    "AAA_111": "20260715:近3月3单 最低9.9@某供应商",
    "BBB_222": "人工写的注意事项",
}

# 按周分组导出用的周销量。CCC 前两周为空（测「空格 = 0」而非缺失），
# 且各 SKU 的「末4周均值」都刻意 ≠「期间总量/周数」，好把两个口径分开验。
WEEKS = ["W27", "W28", "W29", "W30", "W31"]
WEEKLY = {
    "AAA_111": [200, 100, 100, 100, 100],   # 600；末4均 100.0，周均 120.0
    "BBB_222": [60, 60, 60, 60, 60],        # 300
    "CCC_333": [0, 0, 60, 60, 30],          # 150；末4均 37.5，周均 30.0
    "DDD_444": [12, 12, 12, 12, 12],        # 60
    "EEE_555": [6, 6, 6, 6, 6],             # 30
    "FFF_666": [1, 1, 1, 1, 1],             # 5
}


def make_sales(path):
    """旧格式：整期累计，4 列。"""
    rows = [[None, "Total", None, None], [None, "Sales", None, None],
            [None, "Untaxed Total", "# of Lines", "Qty Ordered"],
            ["Total", sum(r[3] for r in DATA), sum(r[2] for r in DATA), sum(r[1] for r in DATA)]]
    for sku, qty, lines, amt, *_ in DATA:
        rows.append([f"    [{sku}] 样例商品 {sku}", amt, lines, qty])   # ← 缩进空格，照真实导出
    pd.DataFrame(rows).to_excel(path, index=False, header=False)


def make_weekly_sales(path):
    """新格式：每周一组 3 列 + 末尾一组无周标签的合计，表头 4 行。"""
    n = len(WEEKS) + 1                                   # 周分组 + 合计组
    hdr2 = [None] + sum([[f"{w} 2026", None, None] for w in WEEKS], []) + [None, None, None]
    hdr3 = [None] + ["Untaxed Total", "# of Lines", "Qty Ordered"] * n
    rows = [[None, "Total"] + [None] * (n * 3 - 1), [None, "Sales"] + [None] * (n * 3 - 1),
            hdr2, hdr3]

    def cells(sku, qty, lines, amt):
        w = WEEKLY[sku]
        out = []
        for q in w:
            # 某周没卖 = 三格全空，不是 0——真实导出就长这样
            out += [None, None, None] if not q else [round(amt * q / qty, 2),
                                                     max(1, round(lines * q / qty)), q]
        return out + [amt, lines, sum(w)]                # 合计组

    detail = [(f"    [{sku}] 样例商品 {sku}", *cells(sku, qty, lines, amt))
              for sku, qty, lines, amt, *_ in DATA]
    tot = ["Total"] + [sum(r[i + 1] or 0 for r in detail) for i in range(n * 3)]
    pd.DataFrame(rows + [tot] + [list(r) for r in detail]).to_excel(
        path, index=False, header=False)


def make_inputs(d, tag, with_erp_onhand=True, with_remark=True, erp_remark=None):
    """tag 让每种组合落到**不同文件名**——同名互相覆盖会让后面的用例悄悄用错输入。"""
    sales = os.path.join(d, f"sales_{tag}.xlsx")
    make_sales(sales)

    erp_remark = ERP_REMARK if erp_remark is None else erp_remark
    prods = os.path.join(d, f"products_{tag}.xlsx")
    pr = [{"ID": f"__export__.product_product_{i}_abc",
           "Internal Reference": f"\t{r[0]}",                          # ← \t 前缀，照真实导出
           "Barcode": f"400000000{i}", "Name": f"样例商品 {r[0]}",
           "VO Shop Name": "TKOF_SHOP1_VO",
           "Safety Stock": 7 if r[0] == "AAA_111" else 0,   # ERP 现值，与运营值(40)不同
           "Product/FS": "样例供应商",
           **({"Supply Remark": erp_remark.get(r[0])} if with_remark else {}),
           **({"Quantity On Hand": r[5]} if with_erp_onhand else {})}
          for i, r in enumerate(DATA) if r[4]]
    pd.DataFrame(pr).to_excel(prods, index=False)

    safety = os.path.join(d, f"safety_{tag}.xlsx")
    sr = [{"商品SKU": r[0], "W29": 10, "W30": 12, "总销量": r[1],
           "8.01在手库存": r[7], "安全库存": r[6], "备注": r[8]}      # ← 列名带日期前缀
          for r in DATA if r[6] is not None]
    pd.DataFrame(sr).to_excel(safety, index=False)
    return sales, prods, safety


fails = []


def check(name, cond, detail=""):
    print(("  ✅ " if cond else "  ❌ ") + name + (f"  {detail}" if detail else ""))
    if not cond:
        fails.append(name)


def run(d, sales, prods, safety, out, *extra, weeks=PERIOD_WEEKS):
    cmd = [sys.executable, SCRIPT, sales, "--products", prods, "--safety", safety,
           "--cover-weeks", str(COVER_WEEKS), "-o", out, *extra]
    if weeks is not None:
        cmd += ["--weeks", str(weeks)]
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode:
        print(r.stdout, r.stderr)
    return r


IDS = {sku: f"__export__.product_product_{i}_abc" for i, (sku, *_) in enumerate(DATA)}
WB_COLS = ["id", "SKU(勿导入)", "Safety Stock", "Supply Remark"]


def main():
    d = tempfile.mkdtemp(prefix="sales_insight_test_")

    print("【1】ERP 带 Quantity On Hand（目标形态）")
    sales, prods, safety = make_inputs(d, "full")
    r = run(d, sales, prods, safety, f"{d}/o1")
    check("退出码 0", r.returncode == 0, r.stderr[-200:] if r.returncode else "")
    check("识别到 ERP 在手库存列", "带在手库存列" in r.stdout)
    rank = pd.read_excel(f"{d}/o1/销量排名.xlsx").set_index("SKU")

    check("SKU 解析成功（缩进空格没挡住）", len(rank) == len(DATA), f"得到 {len(rank)} 行")
    n_master = sum(1 for r in DATA if r[4])
    check("Internal Reference 的 \\t 已 strip（主数据里的 SKU 全连上）",
          rank["商品名称"].notna().sum() == n_master, f"得到 {rank['商品名称'].notna().sum()}")
    check("销量降序 + 排名正确", list(rank["销量排名"]) == list(range(1, len(DATA) + 1)))
    check("AAA 每单件数 = 600/600 = 1.0", rank.loc["AAA_111", "每单件数"] == 1.0)
    check("BBB 每单件数 = 300/250 = 1.2", rank.loc["BBB_222", "每单件数"] == 1.2)
    check("累计占比末行 = 1.0", abs(rank["累计占比"].iloc[-1] - 1.0) < 1e-6)

    # 周均 = 期间销量/30；推算 = 周均×2
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
          set(wb["id"]) == {IDS["AAA_111"], IDS["BBB_222"]}, f"得到 {sorted(wb['id'])}")
    check("回写表报出「有人工值但无 ERP ID」的 EEE", "EEE_555" in r.stdout)
    check("回写值 = 人工值", wb.set_index("id").loc[IDS["AAA_111"], "Safety Stock"] == 40)
    cand = pd.read_excel(f"{d}/o1/安全库存候选值.xlsx")
    check("候选值只含推算的", set(cand["SKU(勿导入)"]) == {"CCC_333", "DDD_444"},
          f"得到 {sorted(cand['SKU(勿导入)'])}")
    check("候选值与回写表无交集（两张表各管一摊）",
          not ({IDS[s] for s in cand["SKU(勿导入)"]} & set(wb["id"])))
    # 推算=0 的低销量长尾必须剔除——导进去只会把它们的 Safety Stock 刷成 0
    check("推算值为 0 的 FFF 被剔除", "FFF_666" not in set(cand["SKU(勿导入)"]))
    check("候选值 Safety Stock 全部 >0", (cand["Safety Stock"] > 0).all())
    check("终端报出被剔除的条数", "推算值为 0" in r.stdout)
    # 前三列与回写表同头 → 审完可直接导；后面的中文列 Odoo 认不出，不会被写
    check("候选值前三列与回写表一致", list(cand.columns)[:3] == WB_COLS[:3],
          f"得到 {list(cand.columns)[:3]}")
    check("候选值带审阅依据列", {"销量", "周均销量", "在手库存"} <= set(cand.columns))
    check("候选值不带 Supply Remark（候选品没有运营备注，带了平白重写字段）",
          "Supply Remark" not in cand.columns)

    print("\n【2】回写表的列：人可读的 SKU + 不会清空的 Supply Remark")
    check("列恰好是 id / SKU(勿导入) / Safety Stock / Supply Remark",
          list(wb.columns) == WB_COLS, f"得到 {list(wb.columns)}")
    # 表头不能叫 Internal Reference：那会被 Odoo 自动映射，忘了取消勾选就重写了 SKU
    check("SKU 列表头 Odoo 认不出（不叫 Internal Reference）",
          "Internal Reference" not in wb.columns)
    check("SKU 列填的是人看得懂的货号", set(wb["SKU(勿导入)"]) == {"AAA_111", "BBB_222"},
          f"得到 {sorted(wb['SKU(勿导入)'])}")
    w = wb.set_index("SKU(勿导入)")["Supply Remark"]
    check("AAA 备注前置了带日期的新段",
          w["AAA_111"].startswith(f"{TODAY}:安全库存 已下架，清完为止"), f"得到 {w['AAA_111']}")
    check("AAA 原有的 FS 画像段被保住（没被抹掉）",
          ERP_REMARK["AAA_111"] in w["AAA_111"], f"得到 {w['AAA_111']}")
    # 这是本次最大的坑：留空会让 Odoo 清掉这些产品在 ERP 里已有的备注
    check("BBB 没有运营备注 → 原样带回 ERP 现值，不是空",
          w["BBB_222"] == ERP_REMARK["BBB_222"], f"得到 {w['BBB_222']!r}")

    print("\n【3】重跑不堆叠：认出自己上次写的段并替换")
    old = dict(ERP_REMARK)
    old["AAA_111"] = f"20260701:安全库存 上个月的旧话；{ERP_REMARK['AAA_111']}"
    _, prods_rr, _ = make_inputs(d, "rerun", erp_remark=old)
    r_rr = run(d, sales, prods_rr, safety, f"{d}/o_rr")
    check("退出码 0", r_rr.returncode == 0)
    w_rr = pd.read_excel(f"{d}/o_rr/安全库存回写表.xlsx").set_index("SKU(勿导入)")["Supply Remark"]
    check("上次的安全库存段被替换掉，没有堆两份",
          w_rr["AAA_111"].count(":安全库存") == 1, f"得到 {w_rr['AAA_111']}")
    check("旧段内容确实不在了", "上个月的旧话" not in w_rr["AAA_111"])
    check("别人（fs_writeback）的段仍然保留", ERP_REMARK["AAA_111"] in w_rr["AAA_111"])

    print("\n【4】产品主数据没有 Supply Remark 列 —— 不写这列，且明确告警")
    _, prods_nr, _ = make_inputs(d, "noremark", with_remark=False)
    r_nr = run(d, sales, prods_nr, safety, f"{d}/o_nr")
    check("退出码 0", r_nr.returncode == 0)
    check("告警「没有 Supply Remark 列」", "没有 `Supply Remark` 列" in r_nr.stdout)
    wb_nr = pd.read_excel(f"{d}/o_nr/安全库存回写表.xlsx")
    check("回写表不带 Supply Remark 列（带了会清空 ERP 现值）",
          list(wb_nr.columns) == ["id", "SKU(勿导入)", "Safety Stock"], f"得到 {list(wb_nr.columns)}")

    print("\n【5】按周分组的销售导出 —— 期间周数自己数出来")
    wsales = os.path.join(d, "sales_weekly.xlsx")
    make_weekly_sales(wsales)
    r_w = run(d, wsales, prods, safety, f"{d}/o_w", weeks=None)      # ← 不给 --weeks
    check("不给 --weeks 也能跑", r_w.returncode == 0, r_w.stderr[-300:] if r_w.returncode else "")
    check("回显数出了 5 个周", "5 个周（W27–W31）" in r_w.stdout, r_w.stdout.splitlines()[0])
    rank_w = pd.read_excel(f"{d}/o_w/销量排名.xlsx").set_index("SKU")
    # 销量必须取**合计组**。取成第一个周分组的话 AAA 会是 200 而不是 600
    check("销量取的是合计组而非首个周分组（AAA=600 不是 200）",
          rank_w.loc["AAA_111", "销量"] == 600, f"得到 {rank_w.loc['AAA_111', '销量']}")
    check("销售额同样取合计组", rank_w.loc["AAA_111", "销售额"] == 6000.0)
    check("周均 = 600/5 = 120.0（周数自动来自表头）",
          rank_w.loc["AAA_111", "周均销量"] == 120.0, f"得到 {rank_w.loc['AAA_111', '周均销量']}")
    check("近4周均销取末 4 周 = 100.0（≠ 周均 120，两个口径确实分开了）",
          rank_w.loc["AAA_111", "近4周均销"] == 100.0, f"得到 {rank_w.loc['AAA_111', '近4周均销']}")
    # CCC 前两周是空格。当缺失跳过的话末4周均值会算成 (60+60+30)/3=50
    check("空周按 0 算：CCC 近4周均销 = (0+60+60+30)/4 = 37.5",
          rank_w.loc["CCC_333", "近4周均销"] == 37.5, f"得到 {rank_w.loc['CCC_333', '近4周均销']}")
    check("近4周均销覆盖到没进运营表的 SKU（DDD）",
          pd.notna(rank_w.loc["DDD_444", "近4周均销"]))
    check("推算值随自动周数走：CCC 周均 30.0 × 2 = 60",
          rank_w.loc["CCC_333", "安全库存"] == 60, f"得到 {rank_w.loc['CCC_333', '安全库存']}")
    r_wo = run(d, wsales, prods, safety, f"{d}/o_wo", weeks=8)        # 显式覆盖
    check("显式 --weeks 覆盖自动值并告警", r_wo.returncode == 0 and "与表头数出来的 5 不符" in r_wo.stdout)
    check("覆盖后周均按 8 周算：AAA = 600/8 = 75.0",
          pd.read_excel(f"{d}/o_wo/销量排名.xlsx").set_index("SKU").loc["AAA_111", "周均销量"] == 75.0)

    print("\n【6】ERP 不带 On Hand —— 应降级到运营表并明确告警")
    _, prods2, _ = make_inputs(d, "noonhand", with_erp_onhand=False)
    r2 = run(d, sales, prods2, safety, f"{d}/o2")
    check("退出码 0", r2.returncode == 0)
    check("告警「没有在手库存列」", "没有在手库存列" in r2.stdout)
    rank2 = pd.read_excel(f"{d}/o2/销量排名.xlsx").set_index("SKU")
    check("AAA 退回运营表的在手库存 5", rank2.loc["AAA_111", "在手库存"] == 5
          and rank2.loc["AAA_111", "库存来源"] == "运营表")
    check("CCC 运营表里没有 → 库存来源「缺」", rank2.loc["CCC_333", "库存来源"] == "缺")

    print("\n【7】--test-sku 首次导入试水")
    r4 = run(d, sales, prods, safety, f"{d}/o4", "--test-sku", "AAA_111")
    check("退出码 0", r4.returncode == 0)
    wb4 = pd.read_excel(f"{d}/o4/安全库存回写表-试AAA_111.xlsx")
    check("试水表只剩 1 行", len(wb4) == 1 and wb4.iloc[0]["id"] == IDS["AAA_111"])
    check("列与正式回写表一致", list(wb4.columns) == WB_COLS, f"得到 {list(wb4.columns)}")
    # 试水的同时也出全量——否则「验完再跑一遍拿全量」，两遍之间数据可能已变，
    # 验过的和导入的就不是同一批了
    check("试水时**同时**产出全量回写表", os.path.exists(f"{d}/o4/安全库存回写表.xlsx"))
    full4 = pd.read_excel(f"{d}/o4/安全库存回写表.xlsx")
    check("全量表与不试水时同样是 2 条", len(full4) == 2, f"得到 {len(full4)}")
    row_t, row_f = wb4.iloc[0], full4[full4["id"] == IDS["AAA_111"]].iloc[0]
    check("试水那条与全量里同一条逐格相等（验完可直接导全量）",
          all((row_t[c] == row_f[c]) or (pd.isna(row_t[c]) and pd.isna(row_f[c]))
              for c in WB_COLS))
    check("终端说清先导哪份", "先导它" in r4.stdout and "全量" in r4.stdout)
    check("同时产出导入前快照", os.path.exists(f"{d}/o4/导入前快照-AAA_111.xlsx"))
    snap = pd.read_excel(f"{d}/o4/导入前快照-AAA_111.xlsx")
    check("快照带 ERP 现值供事后对比（7 → 将写入 40）",
          snap.iloc[0]["ERP现有安全库存"] == 7, f"得到 {snap.iloc[0].get('ERP现有安全库存')}")
    check("终端回显「现值 → 写入值」", "ERP 现值 7" in r4.stdout and "写入 40" in r4.stdout)
    # 报错必须说清缺哪一环——三种原因的处理方式完全不同（换销售导出 / 去安全库存表
    # 加一行 / 重导产品主数据）。2026-08-01 用户就因为糊成一句，误以为是脚本判断有误，
    # 而真正能解释一切的那行诊断在流程更后面、异常先抛了，他根本看不到。
    r5 = run(d, sales, prods, safety, f"{d}/o5", "--test-sku", "不存在的SKU")
    check("① 不在销售数据 → 报错并指向销售导出", r5.returncode != 0
          and "不在销售数据里" in (r5.stdout + r5.stderr))
    r5b = run(d, sales, prods, safety, f"{d}/o5b", "--test-sku", "CCC_333")
    out5b = r5b.stdout + r5b.stderr
    check("② 有销量但无人工值 → 说明走的是推算并报出推算值", r5b.returncode != 0
          and "脚本推算" in out5b and "推算值 10" in out5b, out5b.strip().split("\n")[-1][:70])
    r5c = run(d, sales, prods, safety, f"{d}/o5c", "--test-sku", "EEE_555")
    out5c = r5c.stdout + r5c.stderr
    check("③ 有人工值但主数据里没有 → 直指产品主数据的导出筛选条件", r5c.returncode != 0
          and "产品主数据里找不到" in out5c and "筛选条件" in out5c,
          out5c.strip().split("\n")[-1][:70])

    print("\n【8】External ID 优先于数据库整数 ID")
    pr = pd.read_excel(prods)
    pr = pr.rename(columns={"ID": "External ID"})
    pr["ID"] = range(1000, 1000 + len(pr))          # 数据库整数 ID，绝不能被当映射码
    p6 = os.path.join(d, "products_bothid.xlsx")
    pr.to_excel(p6, index=False)
    r6 = subprocess.run([sys.executable, SCRIPT, sales, "--products", p6, "--safety", safety,
                         "--weeks", str(PERIOD_WEEKS), "-o", f"{d}/o6"], capture_output=True, text=True)
    check("退出码 0", r6.returncode == 0, r6.stderr[-200:] if r6.returncode else "")
    check("终端回显用了 External ID 列", "`External ID`" in r6.stdout)
    wb6 = pd.read_excel(f"{d}/o6/安全库存回写表.xlsx")
    check("映射码是 __export__ 形式而非整数",
          wb6["id"].astype(str).str.startswith("__export__").all(), f"得到 {list(wb6['id'])[:2]}")

    print("\n【9】旧的整期累计格式仍要求 --weeks（那种导出里真没有日期）")
    r3 = subprocess.run([sys.executable, SCRIPT, sales, "--products", prods, "-o", f"{d}/o3"],
                        capture_output=True, text=True)
    # 断言认「说清了缺什么」，不认旗标名——报错文案要同时服务 CLI 和 GUI，不能提 --weeks
    check("缺周数时报错退出且说清原因", r3.returncode != 0
          and "期间周数" in (r3.stderr + r3.stdout), (r3.stderr + r3.stdout)[-80:])

    print("\n" + ("全部通过 🎉" if not fails else f"❌ {len(fails)} 项失败: {fails}"))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
