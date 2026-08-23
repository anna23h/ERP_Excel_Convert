"""Odoo 环境探针：把写正式报表前所有不确定项一次问清楚，只读、不产文件。

跑法：
    python3 odoo_api/discover.py                 # 全部探测项
    python3 odoo_api/discover.py --weeks 4       # 销量口径按近 N 周探

它回答这些问题（每一条都是 stock_report.py 的设计依赖）：
  1. 服务端版本、当前账号、可见公司 —— 决定字段名和是否要锁 allowed_company_ids
  2. 仓库与内部库位有几个 —— 决定在手库存要不要按仓拆
  3. 界面标签 `Safety Stock` 的**技术字段名**在哪个模型上（自定义字段不可猜名）
  4. stock.warehouse.orderpoint 到底有多少条真实规则（要按 trigger 滤掉临时建议）
  5. product.product 各筛选条件下的行数 —— 复核 `can be sold` 是否仍是正确口径
  6. sale.report 有哪些字段可用、近 N 周各 state 的分布
  7. stock.quant 汇总与 product.qty_available 抽样对拍 —— 确认两种在手口径一致
"""
import argparse
import os
import sys
from collections import Counter
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from odoo_api.odoo_client import Odoo, OdooError, m2o_id, m2o_name  # noqa: E402

SAFETY_LABELS = ["Safety Stock", "Sicherheitsbestand", "安全库存"]
QTY_FIELD_CANDIDATES = ["product_uom_qty", "qty_ordered", "product_qty"]
NBR_FIELD_CANDIDATES = ["nbr", "count", "order_count"]
AMOUNT_FIELD_CANDIDATES = ["price_total", "price_subtotal"]


def head(t):
    print("\n" + "=" * 72 + f"\n{t}\n" + "=" * 72)


def probe(od, weeks):
    # ---- 1. 环境 ---------------------------------------------------------
    head("1. 环境")
    me = od.execute("res.users", "read", [[od.uid],
                                          ["name", "login", "company_id", "company_ids", "lang"]])[0]
    print(f"服务端版本 : {od.server_version}")
    print(f"账号       : {me['name']} <{me['login']}>  uid={od.uid}  lang={me.get('lang')}")
    print(f"默认公司   : {m2o_name(me['company_id'])} (id={m2o_id(me['company_id'])})")
    comps = od.execute("res.company", "read", [me["company_ids"], ["name", "currency_id"]])
    print(f"可见公司   : {len(comps)} 个")
    for c in comps:
        print(f"             - id={c['id']:<4} {c['name']}  本位币={m2o_name(c['currency_id'])}")
    if len(comps) > 1:
        print("  ⚠ 多公司环境：拉数时必须显式传 allowed_company_ids，否则读到的是"
              "「当前用户默认公司」，换台机器/换账号跑结果会变。")

    # ---- 2. 仓库 ---------------------------------------------------------
    head("2. 仓库与内部库位")
    whs = od.search_read_all("stock.warehouse", [], ["name", "code", "lot_stock_id", "company_id"])
    print(f"仓库 {len(whs)} 个：")
    for w in whs:
        print(f"  - id={w['id']:<3} {w['code']:<6} {w['name']}  "
              f"公司={m2o_name(w['company_id'])}  主库位={m2o_name(w['lot_stock_id'])}")
    n_int = od.count("stock.location", [("usage", "=", "internal")])
    print(f"内部库位（usage=internal）：{n_int} 个")
    print("  → 在手库存按 usage=internal 汇总，"
          + ("全仓合计即可（单仓）。" if len(whs) <= 1 else "多仓，需决定合计还是按仓拆。"))

    # ---- 3. Safety Stock 字段 -------------------------------------------
    head("3. 安全库存来源 A：产品上的自定义字段")
    found = []
    for label in SAFETY_LABELS:
        rows = od.field_by_label(["product.template", "product.product"], label)
        for r in rows:
            found.append(r)
            print(f"  标签 {label!r} → {r['model']}.{r['name']}  "
                  f"类型={r['ttype']}  stored={r['store']}")
    if not found:
        print("  ✗ 没找到标签为 " + " / ".join(SAFETY_LABELS) + " 的字段。")
        print("    可能标签换了。用下面这条列出产品上所有自定义字段（x_ 开头）再人工认：")
        cust = od.search_read_all(
            "ir.model.fields",
            [("model", "in", ["product.template", "product.product"]),
             ("name", "like", "x_")],
            ["model", "name", "field_description", "ttype"])
        for r in cust:
            print(f"      {r['model']}.{r['name']:<34} {r['ttype']:<10} {r['field_description']}")
    else:
        f = found[0]
        n_set = od.count(f["model"], [(f["name"], "!=", False), (f["name"], ">", 0)])
        n_all = od.count(f["model"], [])
        print(f"  → 用 {f['model']}.{f['name']}；该模型共 {n_all} 条，"
              f"其中该字段 >0 的 {n_set} 条。")
    return_safety = found[0] if found else None

    # ---- 4. 补货规则 -----------------------------------------------------
    head("4. 安全库存来源 B：stock.warehouse.orderpoint（标准补货规则）")
    op_fields = od.fields_of("stock.warehouse.orderpoint")
    total = od.count("stock.warehouse.orderpoint", [])
    print(f"总记录数 : {total}")
    if "trigger" in op_fields:
        for t in ("auto", "manual"):
            n = od.count("stock.warehouse.orderpoint", [("trigger", "=", t)])
            note = ("← 人工配置的真规则" if t == "auto"
                    else "← Replenishment 视图现场生成的临时建议，**不是**配置好的安全库存")
            print(f"  trigger={t:<7}: {n:<6} {note}")
        real_domain = [("trigger", "=", "auto")]
    else:
        print("  该版本没有 trigger 字段（Odoo 14 及更早），全部视为真实规则。")
        real_domain = []
    n_pos = od.count("stock.warehouse.orderpoint", real_domain + [("product_min_qty", ">", 0)])
    print(f"真实规则中 product_min_qty > 0 的：{n_pos} 条")
    if n_pos:
        sample = od.search_read_all(
            "stock.warehouse.orderpoint", real_domain + [("product_min_qty", ">", 0)],
            ["product_id", "product_min_qty", "product_max_qty", "warehouse_id", "location_id"],
            batch=5)[:5]
        print("  抽样 5 条：")
        for r in sample:
            print(f"    {m2o_name(r['product_id'])[:44]:<44} "
                  f"min={r['product_min_qty']:<8g} max={r['product_max_qty']:<8g} "
                  f"仓={m2o_name(r['warehouse_id'])}")
        dup = Counter(m2o_id(r["product_id"]) for r in od.search_read_all(
            "stock.warehouse.orderpoint", real_domain, ["product_id"]))
        many = [p for p, c in dup.items() if c > 1]
        print(f"  同一产品配了多条规则（多仓/多库位）的产品数：{len(many)}")
        if many:
            print("    → 合并成一行时必须定口径：求和 / 取最大 / 只取主仓。")
    else:
        print("  → 补货规则里没有有效安全库存，报表的 B 列会整列为空；"
              "以产品自定义字段为准。")

    # ---- 5. 产品筛选口径 -------------------------------------------------
    head("5. 产品筛选口径复核")
    checks = [("全部（含停用）", [("active", "in", [True, False])]),
              ("active=True", []),
              ("can be sold（sale_ok=True）", [("sale_ok", "=", True)]),
              ("sale_ok=True 且 active=True", [("sale_ok", "=", True), ("active", "=", True)])]
    vo = od.field_by_label(["product.product", "product.template"], "VO active")
    if vo:
        checks.append((f"VO active（{vo[0]['model']}.{vo[0]['name']}）",
                       [(vo[0]["name"], "=", True)]))
    for name, dom in checks:
        print(f"  {name:<38} {od.count('product.product', dom):>7} 行")
    print("  → 对照 sales_insight/README.md：`can be sold` 实测 10266 行，"
          "`VO active=true` 4575 行会漏 9 个 SKU。数字对不上就说明主数据变了，回去复核。")

    # ---- 6. sale.report --------------------------------------------------
    head(f"6. sale.report 字段与近 {weeks} 周的 state 分布")
    sr = od.fields_of("sale.report")
    for label, cands in [("数量", QTY_FIELD_CANDIDATES),
                         ("订单行数", NBR_FIELD_CANDIDATES),
                         ("金额", AMOUNT_FIELD_CANDIDATES)]:
        hit = [c for c in cands if c in sr]
        print(f"  {label:<6} 候选 {cands} → 存在：{hit or '✗ 一个都没有'}")
    for f in ("date", "state", "product_id", "warehouse_id", "company_id", "currency_id"):
        print(f"  {f:<12} {'有' if f in sr else '✗ 没有'}"
              + (f"  ({sr[f]['string']}, {sr[f]['type']})" if f in sr else ""))
    since = (date.today() - timedelta(weeks=weeks)).isoformat()
    grp = od.read_group_all("sale.report", [("date", ">=", since)], ["__count"], ["state"])
    print(f"  近 {weeks} 周（date >= {since}）各状态记录数：")
    for g in grp:
        print(f"    state={str(g.get('state')):<12} {g.get('__count', g.get('state_count', 0)):>7}")
    print("  → 报表默认只算 state in ('sale','done')，即已确认订单；"
          "报价单（draft/sent）不该进销量。上面若有大量 draft，说明这个过滤是必要的。")

    # ---- 7. 在手库存两种口径对拍 ----------------------------------------
    head("7. 在手库存：stock.quant 汇总 vs product.qty_available 抽样对拍")
    qd = [("location_id.usage", "=", "internal")]
    print(f"  内部库位上的 quant 记录：{od.count('stock.quant', qd)} 条")
    grp = od.read_group_all("stock.quant", qd, ["quantity:sum"], ["product_id"])
    print(f"  涉及产品：{len(grp)} 个")
    grp = [g for g in grp if g.get("quantity")]
    sample = sorted(grp, key=lambda g: -abs(g["quantity"]))[:8]
    ids = [m2o_id(g["product_id"]) for g in sample]
    prods = {p["id"]: p for p in od.execute(
        "product.product", "read", [ids, ["default_code", "name", "qty_available"]])}
    bad = 0
    print(f"  {'SKU':<28}{'quant 汇总':>12}{'qty_available':>15}  一致？")
    for g in sample:
        pid = m2o_id(g["product_id"])
        p = prods[pid]
        ok = abs(g["quantity"] - p["qty_available"]) < 0.001
        bad += 0 if ok else 1
        sku = (p["default_code"] or "").strip() or f"id={pid}"
        print(f"  {sku[:28]:<28}{g['quantity']:>12g}{p['qty_available']:>15g}  {'✓' if ok else '✗'}")
    print("  → 全 ✓ 则两种口径等价，正式报表用 read_group（一次查询，且能按仓拆）。")
    print("    有 ✗ 通常是该产品有 transit/客户库位存货，或多公司下 qty_available 只算了默认公司。")

    return return_safety


def main():
    ap = argparse.ArgumentParser(description="Odoo 环境探针（只读，不产文件）")
    ap.add_argument("--weeks", type=int, default=4, help="销量口径探测窗口（默认 4 周）")
    ap.add_argument("--company-id", type=int, default=None,
                    help="多公司环境下锁定公司（allowed_company_ids）")
    ap.add_argument("--timeout", type=int, default=180)
    a = ap.parse_args()
    ctx = {"allowed_company_ids": [a.company_id]} if a.company_id else {}
    try:
        od = Odoo.connect(timeout=a.timeout, context=ctx)
        probe(od, a.weeks)
    except OdooError as e:
        raise SystemExit(f"\n✗ {e}")
    print(f"\n探测完成，共 {od.call_count} 次 RPC 调用。")


if __name__ == "__main__":
    main()
