"""构造数据测试：不连 ERP，用假 Odoo 把 build() 的合并路径整条跑通。

    python3 odoo_api/test_stock_report.py

覆盖的边角（每条都是真会发生的）：
  - 有销量但 sale_ok=False（已下架仍在卖）→ 必须进报表并标 在售=否
  - 有销量但两路安全库存都没配 → 必须且只有它进「待配清单」
  - phantom BoM 套件：quant 上是 0，在手必须取 qty_available，且标 套件=是
  - 违规数据：vo_sku 剥店铺前缀、只算缺货类、被罚过未配安全库存的要标出来
  - 峰值口径：同样销量下，尖峰型 SKU 的可撑峰值周数必须显著低于平稳型
  - 有库存无销量 / 有销量无库存
  - 安全库存只有 A / 只有 B / 两边都有且不等 / 两边相等
  - 自定义字段挂在 product.template 上（要经 product_tmpl_id 取值）
  - 同一产品配多条补货规则 → 按 sum 合并
  - orderpoint 里的 trigger='manual' 临时建议 → 必须被滤掉
  - read_group 返回的 product_id 为 False（产品已删的历史行）→ 跳过不崩
"""
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from odoo_api.stock_report import (build, mismatch, unconfigured,  # noqa: E402
                                   violation_risk, read_violations,
                                   STOCK_VIOLATION_TYPES, V_SHEET)

# --- 假数据 ---------------------------------------------------------------
# id: (sku, name, tmpl_id, active, sale_ok)
PRODUCTS = {
    1: ("A-001", "热卖品", 101, True, True),
    2: ("A-002", "次热卖", 102, True, True),
    3: ("A-003", "有库存没销量", 103, True, True),
    4: ("A-004", "已下架但仍有销量", 104, True, False),
    5: ("A-005", "两路安全库存不等", 105, True, True),
    6: ("A-006", "只配了补货规则", 106, True, True),
    7: ("A-007", "有销量但两路都没配", 107, True, True),
    8: ("A-008", "组合装套件(2x A-001)", 108, True, True),
}
# phantom BoM 套件：quant 上恒为 0，qty_available 由组件推算（A-001 在手 250 → 250/2=125）
KITS = {8}
QTY_AVAILABLE = {1: 250.0, 2: 10.0, 3: 77.0, 4: 0.0, 5: 5.0, 6: 0.0, 7: 30.0, 8: 125.0}
SALES = {1: (500, 480, 12000.0), 2: (300, 290, 7000.0), 7: (200, 195, 5000.0),
         4: (120, 118, 2400.0), 5: (60, 60, 1500.0)}
# 渠道拆分：{product_id: C 端(VO/GW)件数}，其余算 B 端。
# A-001 是典型的「全渠道看着很大、C 端其实很小」——B 端一张整箱单 480 件。
SALES_C_END = {1: 20, 2: 300, 7: 50, 4: 120, 5: 60}
# 周分布（4 周窗口）。A-002 平稳、A-007 尖峰型：同样卖 200 件，一周就吃掉 170。
WEEKLY = {1: [125, 125, 125, 125], 2: [75, 75, 75, 75],
          7: [10, 10, 10, 170], 4: [120, 0, 0, 0], 5: [15, 15, 15, 15]}
QUANTS = [  # (product_id, warehouse_id, warehouse_name, qty, reserved)
    (1, 1, "主仓", 200.0, 20.0), (1, 2, "分仓", 50.0, 0.0),
    (2, 1, "主仓", 10.0, 0.0),
    (3, 1, "主仓", 77.0, 0.0),
    (5, 1, "主仓", 5.0, 0.0),
    (7, 1, "主仓", 30.0, 0.0),
]
# 自定义字段挂 product.template 上 → 键是 tmpl_id
SAFETY_TMPL = {101: 100, 102: 50, 104: 30, 105: 40}
ORDERPOINTS = [  # (product_id, min_qty, trigger)
    (1, 60, "auto"), (1, 40, "auto"),        # 多仓两条 → sum=100，与 A 相等
    (5, 90, "auto"),                          # 与 A(40) 不等
    (6, 25, "auto"),                          # 只有 B
    (2, 999, "manual"),                       # 临时建议，必须被滤掉
]


class FakeOdoo:
    call_count = 0
    verbose = False
    server_version = "17.0-fake"

    def fields_of(self, model):
        return {
            "sale.report": {"product_uom_qty": {}, "nbr": {}, "price_total": {},
                            "date": {}, "state": {}, "product_id": {}},
            "stock.quant": {"quantity": {}, "reserved_quantity": {},
                            "product_id": {}, "warehouse_id": {}, "location_id": {}},
            "stock.warehouse.orderpoint": {"product_id": {}, "product_min_qty": {},
                                           "warehouse_id": {}, "trigger": {}},
        }[model]

    def pick_field(self, model, cands, required=True, what=""):
        have = self.fields_of(model)
        for c in cands:
            if c in have:
                return c
        if required:
            raise AssertionError(f"{model} 缺字段 {cands}")
        return None

    def execute_read_qty(self, ids):
        return [{"id": i, "qty_available": QTY_AVAILABLE.get(i, 0.0)} for i in ids]

    def field_by_label(self, models, label):
        if label != "Safety Stock":
            return []
        return [{"model": "product.template", "name": "x_studio_safety_stock",
                 "ttype": "integer", "store": True}]

    def read_group_all(self, model, domain, fields, groupby, lazy=False, **kw):
        if model == "sale.report" and "date:week" in groupby:
            out = []
            for p, ws in WEEKLY.items():
                for i, q in enumerate(ws):
                    if q:
                        out.append({"product_id": [p, PRODUCTS[p][1]],
                                    "date:week": f"W{20 + i} 2026", "product_uom_qty": q})
            return out
        if model == "sale.report":
            # 假 Odoo 只认「domain 里有没有 name 的 like 条件」，够用来验渠道分支
            filtered = any(isinstance(t, tuple) and t[0] == "name" for t in domain)
            rows = []
            for p, (q, n, amt) in SALES.items():
                qty = SALES_C_END.get(p, 0) if filtered else q
                if filtered and not qty:
                    continue
                rows.append({"product_id": [p, PRODUCTS[p][1]], "product_uom_qty": qty,
                             "nbr": n, "price_total": amt})
            rows.append({"product_id": False, "product_uom_qty": 9, "nbr": 1,
                         "price_total": 1.0})       # 产品已删的历史行
            return rows
        if model == "stock.move":
            return []      # 构造数据不测在途，只保证不崩（真实链路已在 ERP 上验过）
        if model == "stock.quant":
            by_wh = "warehouse_id" in groupby
            agg = {}
            for pid, wid, wname, q, res in QUANTS:
                k = (pid, wid) if by_wh else (pid, None)
                a = agg.setdefault(k, [0.0, 0.0, wname])
                a[0] += q
                a[1] += res
            out = []
            for (pid, wid), (q, res, wname) in agg.items():
                r = {"product_id": [pid, PRODUCTS[pid][1]], "quantity": q,
                     "reserved_quantity": res}
                if by_wh:
                    r["warehouse_id"] = [wid, wname]
                out.append(r)
            return out
        raise AssertionError(model)

    def search_read_all(self, model, domain, fields, batch=2000, order="id", label=None):
        if model == "product.template":
            return [{"id": t, "x_studio_safety_stock": v} for t, v in SAFETY_TMPL.items()]
        if model == "stock.warehouse.orderpoint":
            trig = dict(domain).get("trigger") if isinstance(domain, dict) else None
            want = [d[2] for d in domain if d[0] == "trigger"]
            rows = [(p, v, t) for p, v, t in ORDERPOINTS if not want or t == want[0]]
            return [{"product_id": [p, PRODUCTS[p][1]], "product_min_qty": v,
                     "warehouse_id": [1, "主仓"]} for p, v, t in rows]
        if model == "stock.move":
            return []          # 构造数据里不测在途，只保证不崩（真实链路已在 ERP 上验过）
        if model == "res.partner":
            return []
        if model == "mrp.bom":
            return [{"product_tmpl_id": [PRODUCTS[k][2], PRODUCTS[k][1]]} for k in KITS]
        if model == "ir.model.data":
            return [{"module": "__export__", "name": f"product_product_{p}", "res_id": p}
                    for p in PRODUCTS]
        raise AssertionError(model)

    def execute(self, model, method, args=None, kwargs=None, retries=2):
        if (model, method) == ("product.product", "search"):
            dom = args[0]
            tmpls = [t[2] for t in dom if isinstance(t, tuple) and t[0] == "product_tmpl_id"]
            if tmpls:      # pull_kits 按模板反查套件的 product id
                return [p for p, v in PRODUCTS.items() if v[2] in tmpls[0]]
            return [p for p, v in PRODUCTS.items() if v[4]]
        if (model, method) == ("product.product", "read"):
            if args[1] == ["qty_available"]:
                return self.execute_read_qty(args[0])
            return [{"id": p, "default_code": PRODUCTS[p][0], "name": PRODUCTS[p][1],
                     "product_tmpl_id": [PRODUCTS[p][2], PRODUCTS[p][1]],
                     "active": PRODUCTS[p][3], "sale_ok": PRODUCTS[p][4],
                     "uom_id": [1, "Units"]} for p in args[0]]
        raise AssertionError((model, method))


def check(cond, msg):
    print(("  ✓ " if cond else "  ✗ ") + msg)
    assert cond, msg


def main():
    od = FakeOdoo()
    df, src = build(od, "2026-07-26", "2026-08-23", 4, ["sale", "done"],
                    by_warehouse=True, op_agg="sum", want_xid=True)
    print()
    print(df.to_string(index=False))
    print()

    r = df.set_index("SKU")
    check(src == "product.template.x_studio_safety_stock", f"安全库存字段来源识别: {src}")
    check(list(df["SKU"])[:2] == ["A-001", "A-002"], "按销量降序")
    check(len(df) == 8, f"8 个产品全部进表（并集），实际 {len(df)}")
    check(r.loc["A-008", "在手库存"] == 125.0,
          f"套件在手取 qty_available=125（quant 上是 0），实际 {r.loc['A-008', '在手库存']}")
    check(r.loc["A-008", "实物库存"] == 0.0, "套件的实物库存(quant 原值)是 0")
    check(r.loc["A-008", "套件"] == "是", "套件行被标出来")
    check(r.loc["A-001", "套件"] == "", "非套件行不标")
    check(r.loc["A-001", "在手库存"] == r.loc["A-001", "实物库存"] == 250.0,
          "非套件的在手与实物一致")
    check(r.loc["A-001", "预测库存"] == 250.0, "无在途/待出时预测库存 = 在手")
    check(r.loc["A-007", "峰值周销量"] == 170.0, "A-007 峰值周 170 件")
    check(r.loc["A-002", "峰值周销量"] == 75.0, "A-002 平稳，峰值 = 周均 75")
    check(r.loc["A-002", "峰值倍数"] == 1.0, "平稳型峰值倍数 = 1.0")
    check(r.loc["A-007", "有销周数"] == 4, "A-007 四周都有销量")
    check(r.loc["A-007", "可撑峰值周数"] == 0.2,
          f"A-007 在手 30 / 峰值 170 = 0.2 周，实际 {r.loc['A-007','可撑峰值周数']}")
    check(r.loc["A-007", "峰值缺口"] == 140.0, "A-007 峰值缺口 = 170 - 30")
    check(pd.isna(r.loc["A-001", "峰值缺口"]),
          "A-001 在手 250 > 峰值 125 → 峰值缺口为空")
    # 关键对比：A-007 按周均看能撑 0.6 周，按峰值看只有 0.2 周 —— 均值口径低估 3 倍
    check(round(r.loc["A-007", "可撑周数"] / r.loc["A-007", "可撑峰值周数"], 1) == 3.0,
          "尖峰型 SKU：均值口径给出的覆盖是峰值口径的 3 倍（正是它会漏判的原因）")
    check(r.loc["A-002", "预测缺口"] == r.loc["A-002", "缺口"] == 40.0,
          "无在途时预测缺口与缺口一致")
    check(r.loc["A-004", "在售"] == "否", "已下架但有销量 → 进表且标 在售=否")
    check(r.loc["A-003", "销量"] == 0, "有库存无销量 → 销量 0")
    check(r.loc["A-001", "在手库存"] == 250, "多仓在手合计 200+50=250")
    check(r.loc["A-001", "在手·主仓"] == 200 and r.loc["A-001", "在手·分仓"] == 50,
          "按仓拆列正确")
    check(r.loc["A-001", "可用库存"] == 230, "可用 = 在手 250 - 预留 20")
    check(r.loc["A-001", "安全库存_补货规则"] == 100, "多条 auto 规则 sum=60+40=100")
    check(r.loc["A-002", "安全库存_补货规则"] != 999, "trigger=manual 的临时建议被滤掉")
    check(r.loc["A-006", "安全库存_产品字段"] is None
          or r.loc["A-006", "安全库存_产品字段"] != r.loc["A-006", "安全库存_产品字段"],
          "A-006 只有补货规则，产品字段为空")
    check(r.loc["A-006", "安全库存_取用"] == 25, "A 空时退回 B")
    check(r.loc["A-005", "安全库存_取用"] == 40, "A 有值时以 A 为准（A=40 而非 B=90）")
    check(r.loc["A-002", "缺口"] == 40, "缺口 = 安全 50 - 在手 10")
    check(r.loc["A-001", "缺口"] != r.loc["A-001", "缺口"], "在手 250 > 安全 100 → 缺口为空")
    check(abs(r.loc["A-001", "累计占比"] - 42.37) < 0.02,
          f"ABC 累计占比 500/1180=42.37%（分母不含产品已删的那 9 件），"
          f"实际 {r.loc['A-001', '累计占比']}")
    check(r.loc["A-001", "周均销量"] == 125.0, "周均 = 500/4")
    check(r.loc["A-001", "可撑周数"] == 2.0, "可撑周数 = 250/125")
    check(r.loc["A-001", "External ID"] == "__export__.product_product_1", "External ID 映射")

    mm, n_only_a = mismatch(df)
    print()
    print(mm.to_string(index=False))
    got = dict(zip(mm["SKU"], mm["差异类型"]))
    check(got.get("A-005") == "两边都有但不等", "A-005 两路不等被抓出")
    check(got.get("A-006") == "只有补货规则", "A-006 只有 B 侧被抓出")
    check("A-002" not in got, "A-002 只有 A 侧 → **不进**不一致表（B 没在用时是常态）")
    check(n_only_a == 2, f"只有 A 侧的计数单独报出（A-002/A-004 共 2 个），实际 {n_only_a}")
    check("A-001" not in got, "A-001 两边都是 100 → 不算不一致")
    check("A-003" not in got, "A-003 两边都空 → 不算不一致（那是没配，不是冲突）")

    un = unconfigured(df)
    print()
    print(un.to_string(index=False))
    check(list(un["SKU"]) == ["A-007"], f"待配清单只含 A-007，实际 {list(un['SKU'])}")
    check("A-003" not in set(un["SKU"]), "A-003 零销量 → 不进待配清单（没销量谈不上待配）")
    check("A-006" not in set(un["SKU"]), "A-006 只有 B 侧也算已配 → 不进待配清单")
    check(float(un.iloc[0]["可撑周数"]) == 0.6, "待配清单带可撑周数 30/(200/4)=0.6")
    # ---- 渠道过滤 ----
    print("\n---- 渠道过滤（--ref-contains VO,GW）----")
    df2, _ = build(od, "2026-07-26", "2026-08-23", 4, ["sale", "done"],
                   by_warehouse=False, op_agg="sum", want_xid=False,
                   contains=["VO", "GW"])
    print()
    print(df2[["排名", "SKU", "销量", "销量_其它渠道", "销量_全渠道",
               "在手库存", "可撑周数"]].to_string(index=False))
    r2 = df2.set_index("SKU")
    check(r2.loc["A-001", "销量"] == 20, "筛选后 A-001 只剩 C 端 20 件")
    check(r2.loc["A-001", "销量_其它渠道"] == 480, "其它渠道 = 500 - 20 = 480")
    check(r2.loc["A-001", "销量_全渠道"] == 500, "全渠道列保留原值")
    check(list(df2["SKU"])[0] == "A-002",
          f"排名按筛选后销量重排，A-002(300) 顶掉 A-001(20)，实际首位 {list(df2['SKU'])[0]}")
    check(r2.loc["A-001", "可撑周数"] == 2.0,
          "可撑周数仍按全渠道算 250/(500/4)=2.0，不因筛选而虚高")
    check(r2.loc["A-002", "可撑周数"] == 0.1, "A-002 全渠道=C端，可撑周数不变 10/(300/4)")
    check("销量_全渠道" not in df.columns, "不加筛选时不产出渠道拆分列")

    # ---- 违规数据 ----
    print("\n---- 违规数据接入 ----")
    import tempfile, os as _os
    vf = _os.path.join(tempfile.mkdtemp(), "viol.xlsx")
    pd.DataFrame({
        "vo_sku": ["D62-A-001", "D62-A-001", "F38-A-002", "D62-A-007",
                   "D62-A-007", "D62-A-005", "D62-", "D62-A-001"],
        "违规类型": [STOCK_VIOLATION_TYPES[0], STOCK_VIOLATION_TYPES[1],
                     STOCK_VIOLATION_TYPES[2], STOCK_VIOLATION_TYPES[3],
                     "晚于应发货时间-商家考核", STOCK_VIOLATION_TYPES[0],
                     STOCK_VIOLATION_TYPES[0], "退货退款"],
        "处罚金额(结算币)": [10.0, 20.0, 5.0, 7.0, 99.0, 3.0, 50.0, 1.0],
        "处罚时间": pd.to_datetime(["2026-08-01", "2026-08-02", "2026-08-03", "2026-08-04",
                                    "2026-08-05", "2026-01-01", "2026-08-06", "2026-08-07"]),
    }).to_excel(vf, sheet_name=V_SHEET, index=False)

    v = read_violations(vf)
    check(list(v.index) == sorted(v.index) and "A-001" in v.index,
          "vo_sku 的 D62-/F38- 店铺前缀被剥掉")
    check("" not in v.index and len(v) == 4, f"空 vo_sku 被跳过，剩 4 个 SKU，实际 {len(v)}")
    check(v.loc["A-001", "缺货违规单数"] == 2,
          "A-001 三条里只有 2 条是缺货类（退货退款不算）")
    check(v.loc["A-001", "缺货罚款EUR"] == 30.0, "A-001 缺货罚款 10+20=30，不含退货退款那 1")
    check(v.loc["A-001", "违规单数_全部"] == 3, "全部违规单数含非缺货类")
    check(v.loc["A-007", "缺货违规单数"] == 1 and v.loc["A-007", "缺货罚款EUR"] == 7.0,
          "「晚于应发货时间」不计入缺货类（那 99 EUR 不进）")

    v2 = read_violations(vf, since="2026-06-01")
    check("A-005" not in v2.index, "--violation-since 生效，2026-01 那条被排除")

    df3, _ = build(od, "2026-07-26", "2026-08-23", 4, ["sale", "done"],
                   by_warehouse=False, op_agg="sum", want_xid=False, viol=v)
    vr = violation_risk(df3)
    print()
    print(vr[["优先级", "缺货罚款EUR", "缺货违规单数", "SKU", "安全库存_取用"]].to_string(index=False))
    pr = dict(zip(vr["SKU"], vr["优先级"]))
    check(pr.get("A-001") == "被罚过·已配安全库存", "A-001 有安全库存(100) → 标已配")
    check(pr.get("A-007") == "被罚过·未配安全库存", "A-007 两路都没配 → 标未配，最该先做")
    check(list(vr["优先级"])[0] == "被罚过·未配安全库存", "未配的排在前面")
    check(set(vr["SKU"]) == {"A-001", "A-002", "A-005", "A-007"},
          f"只有有缺货类处罚的 SKU 进表，实际 {set(vr['SKU'])}")
    check(float(df3[df3.SKU == "A-003"]["缺货违规单数"].iloc[0]) == 0.0,
          "没被罚过的 SKU 在周报里补 0 而不是 NaN")

    print("\n全部通过。")


if __name__ == "__main__":
    main()
