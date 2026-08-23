"""销量 / 在手库存 / 安全库存 三合一周报（Odoo XML-RPC 只读拉数）。

    python3 odoo_api/stock_report.py                    # 近 4 周，默认口径
    python3 odoo_api/stock_report.py --weeks 12         # 换窗口
    python3 odoo_api/stock_report.py --since 2026-06-01 --until 2026-08-23
    python3 odoo_api/stock_report.py --ref-contains VO,GW   # 只看天猫 C 端，剔除 B 端整箱单
    python3 odoo_api/stock_report.py --exclude-vendor "供应商名"  # 覆盖 config 的排除表
    python3 odoo_api/stock_report.py --violations 供应商违规数据.xlsx   # 接处罚历史
    python3 odoo_api/stock_report.py --by-warehouse     # 在手库存按仓拆列
    python3 odoo_api/stock_report.py --csv-only         # 只出 CSV（cron 用）

产出（默认 output/YYYYMMDD/）：
    库存周报.xlsx / .csv        全 SKU，按销量降序，含 ABC 累计占比与两路安全库存对比
                                另有 在途入库/待出库/预测库存（已剔除测试供应商的假在途）
    安全库存不一致.csv          两路真正冲突的行（都有但不等 / 只有补货规则）
    安全库存待配清单.csv        有销量但没配任何安全库存的 SKU，按销量降序
    缺货违规风险清单.csv        给了 --violations 才产：因缺货被天猫罚过的 SKU，按罚款降序，
                                标出「被罚过但至今没配安全库存」那批

口径（2026-08-23 立项时与用户对齐，详见 ISSUES.md 与 sales_insight/README.md）：
  - 销量取 sale.report，按 date 过滤自定义窗口，**默认近 12 周**。
    天猫端需求是尖峰型的（实测峰值/周均中位数 2.42×，80% 的 SKU ≥2×），
    所以除了周均还算 `峰值周销量` / `可撑峰值周数` / `峰值缺口`——
    **罚款发生在峰值那一周，均值口径正好在最要紧的时刻失真**。
    窗口不能太短：4 周只有 4 个数据点，看不出尖峰。不用 sales_count
    （周期写死 12 个月不可改）。只算 state in ('sale','done')，报价单不进销量。
  - **渠道口径要显式给**。默认不筛，销量是全渠道。2026-08-23 实测：B 端（`S0` 开头的
    订单）只占 378 行订单却占 90% 的件数（整箱走货），不筛的话排名完全被它主导，
    而安全库存是给天猫 C 端维护的——两个口径混用会把人引向错误结论。
    `--ref-contains VO,GW` 切到 C 端口径，此时并列产出 `销量_其它渠道` / `销量_全渠道` 两列。
    **`可撑周数` 一律按全渠道算**：库存被所有渠道一起消耗，只用一侧算会高估覆盖。
  - 在手库存取 `product.product.qty_available`——与手工导出的 `Quantity On Hand` 同源。
    **不能用 stock.quant 汇总**：组合装是 phantom BoM 套件，没有自己的实物库存，
    quant 上恒为 0，只有 qty_available 会从组件推算（2026-08-23 实测 557 个套件因此被记成 0）。
    `已预留` 与按仓拆列仍取 quant（qty_available 给不了），故套件行这两列是 0。
    `实物库存` 列保留 quant 原值，与在手对照即可看出哪些是套件推算来的。
    ⚠ **在手列不可跨行求和**：套件与其基础款的在手是同一批实物。
  - 安全库存两路并列：A=产品主数据上的自定义字段（sales_insight 回写的那个，
    补货预判清单读的也是它），B=stock.warehouse.orderpoint.product_min_qty。
    缺口按 A 算，A 空则退回 B。
  - 产品范围 = sale_ok=True ∪ 窗口内有销量的 ∪ 有库存的。**并集而非 sale_ok 单条**：
    只按 sale_ok 会把「已下架但仍有销量/仍有库存」的货静默漏掉，这类恰恰最该看见。
"""
import argparse
import os
import sys
from datetime import date, timedelta

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common.xlsx import write_simple  # noqa: E402
from odoo_api.discover import (AMOUNT_FIELD_CANDIDATES, NBR_FIELD_CANDIDATES,  # noqa: E402
                               QTY_FIELD_CANDIDATES, SAFETY_LABELS)
from odoo_api.odoo_client import Odoo, OdooError, m2o_id  # noqa: E402

SOLD_STATES = ["sale", "done"]

# 天猫处罚数据里**能确定由缺货导致**的违规类型。别往里加「晚于应发货时间」——
# 它占 75% 的单数、59% 落在也有缺货违规的 SKU 上，但没有历史库存快照就分不清
# 是缺货还是仓库作业慢。硬归因等于把仓库效率问题伪装成库存问题（2026-08-23 拍板）。
STOCK_VIOLATION_TYPES = [
    "商家毁单-商家考核",
    "商家毁单-体验赔付",
    "晚于考核商家缺货时间-商家考核",
    "晚于平台判定缺货时间-体验赔付",
]
V_SHEET = "供应商违规数据"
V_SKU, V_TYPE, V_FEE, V_TIME = "vo_sku", "违规类型", "处罚金额(结算币)", "处罚时间"


def say(msg=""):
    print(msg, flush=True)


# --------------------------------------------------------------------------
# 各路数据
# --------------------------------------------------------------------------
def read_violations(path, since=None):
    """天猫「供应商违规数据」导出 → 每 SKU 的处罚统计。

    这是**非 Odoo 数据**（天猫后台导出），性质同 sales_insight 的运营安全库存表：
    选填、由人提供、决定优先级。

    `vo_sku` 形如 `D62-Eucerin_02398780` / `F38-Vagisan_00003435`，前缀是店铺代号，
    剥掉才是 ERP 的 `default_code`（2026-08-23 实测 8647/8647 全部可剥）。
    金额取**结算币**（EUR，与公司本位币一致），不取人民币列。
    """
    d = pd.read_excel(path, sheet_name=V_SHEET)
    miss = [c for c in (V_SKU, V_TYPE, V_FEE, V_TIME) if c not in d.columns]
    if miss:
        raise ValueError(f"违规数据缺列 {miss}；期望的是天猫后台导出的 `{V_SHEET}` sheet。")
    d[V_TIME] = pd.to_datetime(d[V_TIME], errors="coerce")
    n_all = len(d)
    if since:
        d = d[d[V_TIME] >= pd.Timestamp(since)]
    d["_sku"] = (d[V_SKU].astype(str).str.replace(r"^[A-Z]\d+-", "", regex=True).str.strip())
    bad = d["_sku"].isin(["", "nan", "None"])
    if bad.any():
        say(f"  ⚠ 违规数据里 {int(bad.sum())} 条没有 vo_sku，无法归到 SKU，已跳过")
    d = d[~bad]
    rng = f"{d[V_TIME].min():%Y-%m-%d} ~ {d[V_TIME].max():%Y-%m-%d}" if len(d) else "空"
    stock = d[d[V_TYPE].isin(STOCK_VIOLATION_TYPES)]
    say(f"  违规数据: {len(d)}/{n_all} 条（{rng}），涉及 {d['_sku'].nunique()} 个 SKU；"
        f"其中缺货类 {len(stock)} 条 / {stock[V_FEE].sum():,.0f} EUR"
        f"（{len(stock)/len(d)*100:.1f}% 的单数、{stock[V_FEE].sum()/max(d[V_FEE].sum(),1e-9)*100:.1f}% 的罚款）")
    g = stock.groupby("_sku").agg(缺货违规单数=(V_TYPE, "count"),
                                  缺货罚款EUR=(V_FEE, "sum"))
    g["缺货罚款EUR"] = g["缺货罚款EUR"].round(2)
    allg = d.groupby("_sku").size().rename("违规单数_全部")
    return g.join(allg, how="outer").fillna({"缺货违规单数": 0, "缺货罚款EUR": 0.0,
                                             "违规单数_全部": 0})


def ref_domain(contains, excludes):
    """按 Order Reference（`sale.report.name`）过滤渠道。

    contains 是 OR 语义（`VO` 或 `GW` 命中其一即可），excludes 是 AND 语义（逐条都要不含）。
    Odoo domain 的 `like` 会自动包成 `%…%` 即子串匹配；要前缀匹配得用 `=like` 配 `S0%`。
    """
    dom = []
    if contains:
        terms = [("name", "like", v) for v in contains]
        dom += ["|"] * (len(terms) - 1) + terms
    for v in excludes or []:
        dom.append(("name", "not like", v))
    return dom


def pull_sales(od, since, until, states, extra=None, label="销量"):
    """sale.report → {product_id: dict(销量, 下单次数, 销售额)}。

    用 read_group 按 product_id 聚合，一次查询拿全部；不要逐产品查。
    数量字段是 Odoo 折算到产品**参考计量单位**后的值，可直接跨订单行相加。
    """
    qty_f = od.pick_field("sale.report", QTY_FIELD_CANDIDATES, what="销量")
    nbr_f = od.pick_field("sale.report", NBR_FIELD_CANDIDATES, required=False)
    amt_f = od.pick_field("sale.report", AMOUNT_FIELD_CANDIDATES, required=False)
    domain = [("date", ">=", since), ("date", "<=", until), ("state", "in", states)]
    domain += list(extra or [])
    fields = [f"{qty_f}:sum"] + ([f"{amt_f}:sum"] if amt_f else []) \
        + ([f"{nbr_f}:sum"] if nbr_f else [])
    rows = od.read_group_all("sale.report", domain, fields, ["product_id"])
    out = {}
    for r in rows:
        pid = m2o_id(r.get("product_id"))
        if pid is None:
            continue      # 产品被删/为空的历史行，进不了按产品的报表
        out[pid] = {"销量": r.get(qty_f) or 0,
                    "销售额": (r.get(amt_f) or 0) if amt_f else None,
                    "下单次数": (r.get(nbr_f) or 0) if nbr_f else None}
    total = sum(v["销量"] for v in out.values())
    say(f"  {label}: {len(out)} 个产品有成交，合计 {total:,.0f} 件")
    return out, {"qty": qty_f, "nbr": nbr_f, "amt": amt_f}


def pull_kits(od):
    """phantom BoM（虚拟套件）的产品 id 集合。

    套件（`x2` = 2× 基础款、`_GW` = 1× 基础款）**没有自己的实物库存**：
    `stock.quant` 上是 0，只有 `qty_available` 会从组件推算。不认出它们，
    557 个组合装 SKU 的在手会全被记成 0（2026-08-23 实测）。
    """
    try:
        boms = od.search_read_all("mrp.bom", [("type", "=", "phantom")], ["product_tmpl_id"])
    except OdooError as e:
        # 没装 mrp 模块（或无读权限）就没有套件这回事，退回纯 quant 口径并说清楚
        say(f"  ⚠ 读不到 mrp.bom（{str(e).splitlines()[0][:60]}），"
            "无法识别套件；若存在组合装 SKU，其在手会偏低。")
        return set()
    tmpls = sorted({m2o_id(b["product_tmpl_id"]) for b in boms if b.get("product_tmpl_id")})
    if not tmpls:
        return set()
    ids = od.execute("product.product", "search", [[("product_tmpl_id", "in", tmpls)]])
    say(f"  套件: {len(boms)} 条 phantom BoM，涉及 {len(ids)} 个 product")
    return set(ids)


def pull_qty_available(od, pids, label="在手(qty_available)"):
    """按 id 批量读 `qty_available`。

    这是与手工导出 `Quantity On Hand` **同源**的口径，也是唯一对套件正确的口径。
    是计算字段，比 quant 的 read_group 贵，但实测 400 个 0.9s、全量约 25s，可接受。
    """
    out, ids = {}, sorted(pids)
    for i in range(0, len(ids), 1000):
        for r in od.execute("product.product", "read", [ids[i:i + 1000], ["qty_available"]]):
            out[r["id"]] = r["qty_available"] or 0.0
        if od.verbose:
            print(f"  {label}: {len(out)} 个…", end="\r", flush=True)
    if od.verbose:
        print(f"  {label}: {len(out)} 个  ")
    return out


def resolve_excluded_vendors(od, names):
    """供应商名子串 → partner id 列表。找不到就告警，不静默当成"没有要排除的"。"""
    if not names:
        return []
    dom = ["|"] * (len(names) - 1) + [("name", "ilike", n) for n in names]
    rows = od.search_read_all("res.partner", dom, ["name"])
    if not rows:
        say(f"  ⚠ 按名字找不到要排除的供应商 {names}，未做任何排除。")
        return []
    say("  排除供应商: " + "、".join(f"{r['name']}(id={r['id']})" for r in rows))
    return [r["id"] for r in rows]


def pull_in_transit(od, exclude_ids):
    """未完成的进/出库 move → ({pid: 在途入库}, {pid: 待出库})。

    ⚠ 必须能排除**不会真的收货的技术性采购单**（某些系统集成会用确认状态的 PO 建映射），
    否则这两列是垃圾：2026-08-23 实测全局 7460 条未完成入库里 **7350 条（98.5%）**
    来自单一供应商的这类单，真实在途只剩 105 条。不排除的话 6482 个产品都显示有在途，
    实际只有 88 个有。哪些供应商属于此类是**业务知识**，配在 config.py，不写死在代码里。

    用 `purchase_line_id.order_id.partner_id` 精确切分，**不要用 `origin` 字符串匹配**
    （PO 号可能撞）。排除条件必须带 `purchase_line_id = False` 这一支，
    否则调拨/退货等非采购来源的 move 会被一并滤掉。
    """
    pending = [("state", "not in", ["done", "cancel", "draft"])]
    inc = pending + [("location_id.usage", "in", ["supplier", "transit"]),
                     ("location_dest_id.usage", "=", "internal")]
    out = pending + [("location_id.usage", "=", "internal"),
                     ("location_dest_id.usage", "in", ["customer", "transit"])]
    skip = []
    if exclude_ids:
        n_all = od.count("stock.move", inc)
        # 光靠 purchase_line_id 会漏：实测有 move 的 purchase_line_id 是空的、
        # 只有 origin 还写着被排除 PO 的单号（2026-08-23 实测 5 条 / 5 个产品 / 2100 件，
        # 都是几年前的旧单，链大概是历史迁移中断的）。所以两条路都要堵：
        #   链得到 PO 行 → 看它的供应商；链不到 → 拿 origin 对测试 PO 的单号表。
        po_names = [o["name"] for o in od.search_read_all(
            "purchase.order", [("partner_id", "in", exclude_ids)], ["name"])]
        skip = ["|",
                "&", ("purchase_line_id", "!=", False),
                ("purchase_line_id.order_id.partner_id", "not in", exclude_ids),
                "&", ("purchase_line_id", "=", False),
                "|", ("origin", "=", False), ("origin", "not in", po_names)]
        n_keep = od.count("stock.move", inc + skip)
        say(f"  在途: 未完成入库 {n_all} 条，排除测试供应商（{len(po_names)} 张 PO）后剩 "
            f"{n_keep} 条（滤掉 {n_all - n_keep} 条）")
        n_done = od.count("stock.move", [("state", "=", "done"),
                                         ("location_dest_id.usage", "=", "internal"),
                                         ("purchase_line_id.order_id.partner_id", "in", exclude_ids)])
        if n_done:
            say(f"  ⚠ 被排除的供应商还有 {n_done} 条**已完成**入库——那部分已经真的进了在手库存，"
                "本层不做修正（要改得动 quant，属 ERP 数据订正）。")

    def agg(domain):
        return {m2o_id(r["product_id"]): r["product_uom_qty"] or 0
                for r in od.read_group_all("stock.move", domain,
                                           ["product_uom_qty:sum"], ["product_id"])
                if m2o_id(r.get("product_id")) is not None}
    return agg(inc + skip), agg(out)


def pull_weekly_peak(od, since, until, states, extra=None):
    """按「产品 × 周」聚合 → {pid: dict(峰值周销量, 有销周数)}。

    **为什么要单独看峰值**：天猫端需求是尖峰型的——流量激增 → 订单暴涨 → 缺口 → 罚款 →
    流量下跌。2026-08-23 实测（12 周、累计≥20 件的 171 个 SKU）：峰值/周均的**中位数是
    2.42×**，80% 的 SKU ≥2×、30% ≥3×。按周均算的「可撑周数」在中位数上**高估 2.4 倍**，
    而罚款恰恰发生在峰值那一周——均值口径正好在最要紧的时刻失真。

    周标签形如 `W23 2026`，排序必须按 (年, 周) 而不是字符串。
    """
    qty_f = od.pick_field("sale.report", QTY_FIELD_CANDIDATES, what="销量")
    domain = [("date", ">=", since), ("date", "<=", until), ("state", "in", states)]
    domain += list(extra or [])
    rows = od.read_group_all("sale.report", domain, [f"{qty_f}:sum"],
                             ["product_id", "date:week"])
    out = {}
    for r in rows:
        pid = m2o_id(r.get("product_id"))
        if pid is None:
            continue
        q = r.get(qty_f) or 0
        rec = out.setdefault(pid, {"峰值周销量": 0.0, "有销周数": 0})
        if q > 0:
            rec["有销周数"] += 1
            rec["峰值周销量"] = max(rec["峰值周销量"], q)
    say(f"  周维度: {len(rows)} 组 → {len(out)} 个产品有周销量分布")
    return out


def pull_stock(od, by_warehouse):
    """stock.quant → {product_id: dict(在手, 已预留[, 各仓])}。

    ⚠ 只对有实物库存的产品成立。套件（phantom BoM）在这里一律是 0，
    真实可售数量要用 `pull_qty_available`——见 pull_kits 的注释。
    """
    domain = [("location_id.usage", "=", "internal")]
    have = od.fields_of("stock.quant")
    res_f = "reserved_quantity" if "reserved_quantity" in have else None
    fields = ["quantity:sum"] + ([f"{res_f}:sum"] if res_f else [])
    groupby = ["product_id"]
    wh_ok = by_warehouse and "warehouse_id" in have
    if by_warehouse and not wh_ok:
        say("  ⚠ 该版本 stock.quant 没有 warehouse_id，--by-warehouse 忽略（按库位归仓需另查）。")
    if wh_ok:
        groupby.append("warehouse_id")
    rows = od.read_group_all("stock.quant", domain, fields, groupby)
    out, wh_names = {}, {}
    for r in rows:
        pid = m2o_id(r.get("product_id"))
        if pid is None:
            continue
        rec = out.setdefault(pid, {"在手库存": 0.0, "已预留": 0.0})
        rec["在手库存"] += r.get("quantity") or 0
        rec["已预留"] += (r.get(res_f) or 0) if res_f else 0
        if wh_ok and r.get("warehouse_id"):
            wid, wname = r["warehouse_id"][0], r["warehouse_id"][1]
            wh_names[wid] = wname
            rec[f"在手·{wname}"] = rec.get(f"在手·{wname}", 0.0) + (r.get("quantity") or 0)
    say(f"  在手: {len(out)} 个产品有内部库位存货"
        + (f"，分布在 {len(wh_names)} 个仓" if wh_ok else ""))
    return out, sorted(wh_names.values())


def pull_safety_field(od):
    """产品自定义字段 `Safety Stock`（技术名按界面标签反查，不硬编码）。

    → ({product_id: value}, "model.field" 说明串)。找不到字段返回 ({}, None)。
    """
    for label in SAFETY_LABELS:
        rows = od.field_by_label(["product.product", "product.template"], label)
        if rows:
            model, fname = rows[0]["model"], rows[0]["name"]
            break
    else:
        say("  ⚠ 产品上找不到安全库存自定义字段（试过 " + " / ".join(SAFETY_LABELS)
            + "），A 列整列为空。跑 discover.py 第 3 节列出所有 x_ 字段人工认。")
        return {}, None, "id"
    key = "id" if model == "product.product" else "product_tmpl_id"
    # 用 `> 0` 而不是 `!= False`：整数字段上 `!= False` 并不等于 `!= 0`，
    # 2026-08-23 实测它捞回 10361 行再由本地过滤剩 152 行——白翻 5 页、白传一万行。
    recs = od.search_read_all(model, [(fname, ">", 0)],
                              ["id", fname], label=f"安全库存({model}.{fname})")
    out = {r["id"]: r[fname] for r in recs if r[fname]}
    say(f"  安全库存A: {len(out)} 个 {model} 记录有值（字段 {fname}）")
    return out, f"{model}.{fname}", key


def pull_orderpoints(od, agg):
    """stock.warehouse.orderpoint.product_min_qty → {product_id: value}。

    ⚠ Odoo 15+ 的 Replenishment 视图会现场生成 trigger='manual' 的临时建议记录，
    它们不是人工配置的安全库存。只取 trigger='auto'，否则整列被垃圾数据淹没。
    同一产品配多仓多条规则时按 agg 合并（默认 sum，即全公司安全库存总量）。
    """
    have = od.fields_of("stock.warehouse.orderpoint")
    domain = [("trigger", "=", "auto")] if "trigger" in have else []
    rows = od.search_read_all("stock.warehouse.orderpoint", domain,
                              ["product_id", "product_min_qty", "warehouse_id"],
                              label="补货规则")
    out, multi = {}, {}
    for r in rows:
        pid = m2o_id(r.get("product_id"))
        if pid is None:
            continue
        v = r.get("product_min_qty") or 0
        multi[pid] = multi.get(pid, 0) + 1
        out[pid] = (out.get(pid, 0) + v) if agg == "sum" else max(out.get(pid, 0), v)
    n_multi = sum(1 for c in multi.values() if c > 1)
    say(f"  安全库存B: {len(out)} 个产品有补货规则"
        + (f"（{n_multi} 个配了多条，按 {agg} 合并）" if n_multi else "")
        + ("" if domain else "；该版本无 trigger 字段，未过滤临时建议"))
    return out


def pull_products(od, pids):
    """按 id 批量读产品明细。pids 是最终并集，不再整表拉。"""
    fields = ["default_code", "name", "product_tmpl_id", "active", "sale_ok", "uom_id"]
    out, ids = {}, sorted(pids)
    for i in range(0, len(ids), 1000):
        for r in od.execute("product.product", "read", [ids[i:i + 1000], fields]):
            out[r["id"]] = r
    return out


def pull_external_ids(od, pids):
    """External ID（`__export__.product_product_*`）—— 将来接 sales_insight 回写要用它做映射码。"""
    rows, ids = [], sorted(pids)
    for i in range(0, len(ids), 1000):
        rows += od.search_read_all("ir.model.data",
                                   [("model", "=", "product.product"),
                                    ("res_id", "in", ids[i:i + 1000])],
                                   ["module", "name", "res_id"])
    return {r["res_id"]: f"{r['module']}.{r['name']}" for r in rows}


# --------------------------------------------------------------------------
# 合并
# --------------------------------------------------------------------------
def build(od, since, until, weeks, states, by_warehouse, op_agg, want_xid,
          contains=None, excludes=None, exclude_vendors=None, viol=None,
          peak_cover=1.0):
    rd = ref_domain(contains, excludes)
    if rd:
        sales, sfields = pull_sales(od, since, until, states, rd,
                                    label=f"销量(筛选 {'/'.join(contains or [])})")
        allsales, _ = pull_sales(od, since, until, states, label="销量(全渠道)")
    else:
        sales, sfields = pull_sales(od, since, until, states, label="销量(全渠道，未筛渠道)")
        allsales = sales
        say("  ⚠ 未加渠道过滤：销量含所有渠道。若 B 端整箱单与 C 端零售混在一起，"
            "排名会被前者主导——用 --ref-contains 指定渠道。")
    weekly = pull_weekly_peak(od, since, until, states, rd)
    stock, wh_names = pull_stock(od, by_warehouse)
    kits = pull_kits(od)
    vend_ids = resolve_excluded_vendors(od, exclude_vendors)
    incoming, outgoing = pull_in_transit(od, vend_ids)
    safety_a, safety_src, safety_key = pull_safety_field(od)
    safety_b = pull_orderpoints(od, op_agg)

    sellable = set(od.execute("product.product", "search", [[("sale_ok", "=", True)]]))
    say(f"  产品: sale_ok=True {len(sellable)} 个")
    pids = (sellable | set(sales) | set(allsales) | set(stock) | set(safety_b)
            | set(incoming) | set(outgoing))
    extra = pids - sellable
    if extra:
        say(f"  ⚠ 另有 {len(extra)} 个产品不在 sale_ok 清单里但有销量/库存/补货规则，"
            "一并纳入（列 `在售` 标 否）")
    prods = pull_products(od, pids)
    # 在手改用 qty_available：与手工导出的 `Quantity On Hand` 同源，且是唯一对套件正确的口径
    onhands = pull_qty_available(od, pids)
    xids = pull_external_ids(od, pids) if want_xid else {}

    rows = []
    for pid in pids:
        p = prods.get(pid)
        if p is None:
            continue                      # 读不回来的（已删/无权限），跳过并在末尾报数
        s = sales.get(pid, {})
        st = stock.get(pid, {})
        # 自定义字段挂在 template 上时用 product_tmpl_id 取值
        akey = pid if safety_key == "id" else m2o_id(p.get("product_tmpl_id"))
        a = safety_a.get(akey)
        b = safety_b.get(pid)
        main = a if a not in (None, False) else b
        onhand = onhands.get(pid, st.get("在手库存", 0.0))
        qty = s.get("销量", 0) or 0
        qty_all = (allsales.get(pid, {}).get("销量", 0) or 0) if rd else qty
        fc = onhand + incoming.get(pid, 0.0) - outgoing.get(pid, 0.0)   # 预测库存
        wk = weekly.get(pid, {})
        peak = wk.get("峰值周销量", 0.0)
        avg = (qty / weeks) if weeks else 0
        row = {
            "SKU": (p.get("default_code") or "").strip(),
            "商品名称": p.get("name") or "",
            "销量": qty,
            "下单次数": s.get("下单次数"),
            "销售额": s.get("销售额"),
            "周均销量": round(qty / weeks, 2) if weeks else None,
            "在手库存": onhand,
            # 已预留只有 stock.quant 有，套件上必然是 0（它没有自己的实物）
            "已预留": st.get("已预留", 0.0),
            "可用库存": onhand - st.get("已预留", 0.0),
            "实物库存": st.get("在手库存", 0.0),   # quant 原值，与在手对照看套件
            "在途入库": incoming.get(pid, 0.0),
            "待出库": outgoing.get(pid, 0.0),
            "预测库存": fc,
            "安全库存_产品字段": a,
            "安全库存_补货规则": b,
            "安全库存_取用": main,
            "缺口": (main - onhand) if (main not in (None, False) and main > onhand) else None,
            # 预测缺口按 预测库存 算：在手够但已被订单吃掉、或有在途在路上，结论会不一样。
            # 两列并存不是冗余——「现在缺不缺」和「按当前进出算下来会不会缺」是两个问题。
            "预测缺口": (main - fc) if (main not in (None, False) and main > fc) else None,
            # 可撑周数按**全渠道**销量算：库存是被所有渠道一起消耗的，
            # 只用筛选后那一侧算会高估覆盖（B 端一张整箱单就能把货搬空）。
            "可撑周数": round(onhand / (qty_all / weeks), 1) if qty_all and weeks else None,
            "峰值周销量": peak or None,
            # 峰值倍数 = 这个货有多"尖"。1.0 = 需求平稳；实测中位数 2.42
            "峰值倍数": round(peak / avg, 2) if (peak and avg) else None,
            "有销周数": wk.get("有销周数") or None,
            # 按峰值算的覆盖：能接住几个峰值周。这才是"下一个高峰来了扛不扛得住"
            "可撑峰值周数": round(fc / peak, 1) if peak else None,
            "峰值缺口": (round(peak * peak_cover - fc, 1)
                         if (peak and peak * peak_cover > fc) else None),
            "套件": "是" if pid in kits else "",
            "在售": "是" if pid in sellable else "否",
            "已归档": "是" if not p.get("active") else "",
            "产品ID": pid,
        }
        if rd:
            row["销量_其它渠道"] = qty_all - qty
            row["销量_全渠道"] = qty_all
        for w in wh_names:
            row[f"在手·{w}"] = st.get(f"在手·{w}", 0.0)
        if want_xid:
            row["External ID"] = xids.get(pid, "")
        rows.append(row)

    df = pd.DataFrame(rows)
    if viol is not None:
        df = df.merge(viol.reset_index().rename(columns={"_sku": "SKU"}), on="SKU", how="left")
        for c in ("缺货违规单数", "缺货罚款EUR", "违规单数_全部"):
            df[c] = df[c].fillna(0)
        hit = int((df["缺货违规单数"] > 0).sum())
        say(f"  违规连上: {hit} 个 SKU 有缺货类处罚记录"
            f"（其中 {int(((df['缺货违规单数'] > 0) & df['安全库存_取用'].isna()).sum())} 个至今没配安全库存）")
    df = df.sort_values(["销量", "在手库存"], ascending=[False, False])
    total = df["销量"].sum()
    df.insert(0, "排名", range(1, len(df) + 1))
    df.insert(4, "累计占比", (df["销量"].cumsum() / total * 100).round(2) if total else 0)
    if sfields["nbr"] is None:
        df = df.drop(columns=["下单次数"])
    if sfields["amt"] is None:
        df = df.drop(columns=["销售额"])
    missing = len(pids) - len(rows)
    if missing:
        say(f"  ⚠ {missing} 个产品 id 读不回明细（已删或无读权限），未进报表")
    return df, safety_src


def mismatch(df):
    """两路安全库存真正对不上的行 → (DataFrame, 只有产品字段的条数)。

    **「只有产品字段」不进这张表。** 2026-08-23 实测：A 侧 152 条、B 侧 1 条——
    补货规则在这套 ERP 里基本没在用，A 有 B 无是常态而不是异常。把 151 条常态
    塞进「不一致」清单，等于让这张表自我淹没。它只作为一个计数报在终端里。

    真正该看的两类：两边都有但不等（有人改了其中一边）、只有补货规则
    （存在一条主流程读不到的规则，补货预判清单会看不见它）。
    """
    a, b = df["安全库存_产品字段"], df["安全库存_补货规则"]
    diff = a.notna() & b.notna() & (a != b)
    only_b = b.notna() & (b > 0) & a.isna()
    n_only_a = int((a.notna() & (a > 0) & b.isna()).sum())
    out = df[diff | only_b].copy()
    out["差异类型"] = ["两边都有但不等" if d else "只有补货规则"
                       for d in diff.loc[out.index]]
    cols = ["排名", "SKU", "商品名称", "差异类型",
            "安全库存_产品字段", "安全库存_补货规则", "在手库存", "销量"]
    return out[cols], n_only_a


def violation_risk(df):
    """因缺货被罚过的 SKU，按罚款降序 → 直接可以拿去做事的清单。

    最该看的是 `优先级 == 被罚过·未配安全库存` 那批：2026-08-23 首测 529 个被罚 SKU 里
    **449 个（86%）至今没配安全库存**，合计 2,294 EUR / 754 单。
    安全库存配在哪，跟罚款实际发生在哪，是脱节的——这张表就是为了把它接上。
    """
    out = df[df["缺货违规单数"] > 0].copy()
    unset = out["安全库存_取用"].isna()
    out["优先级"] = ["被罚过·未配安全库存" if u else "被罚过·已配安全库存" for u in unset]
    # 排序键用布尔量而不是「优先级」这个中文串——按字符串排的话「已」的码位在「未」之前，
    # 最该先做的那批会被排到后面（2026-08-23 构造数据测试当场抓到）。
    out["_p"] = (~unset).astype(int)
    cols = ["优先级", "缺货罚款EUR", "缺货违规单数", "违规单数_全部", "SKU", "商品名称",
            "销量", "周均销量", "峰值周销量", "峰值倍数", "有销周数",
            "在手库存", "预测库存", "安全库存_取用", "缺口", "预测缺口",
            "可撑周数", "可撑峰值周数", "峰值缺口", "套件", "在售"]
    out = out.sort_values(["_p", "缺货罚款EUR"], ascending=[True, False])
    return out[[c for c in cols if c in out.columns]]


def unconfigured(df):
    """有销量但没配任何安全库存的 SKU，按销量降序。

    这张表是三路数据合并后才看得见的东西：单看销量排名不知道哪些没设防线，
    单看安全库存清单不知道漏掉的是不是要紧货。2026-08-23 首次实测，
    销量前 20 名里 0 个配了安全库存——这正是要拉这份报表的理由。
    """
    out = df[(df["销量"] > 0) & df["安全库存_取用"].isna()].copy()
    cols = ["排名", "SKU", "商品名称", "销量", "累计占比", "周均销量", "峰值周销量",
            "峰值倍数", "在手库存", "预测库存", "可撑周数", "可撑峰值周数", "峰值缺口",
            "缺货违规单数", "缺货罚款EUR"]
    return out[[c for c in cols if c in out.columns]]


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="销量/在手/安全库存三合一周报（Odoo 只读拉数）")
    ap.add_argument("--weeks", type=int, default=12,
                    help="销量统计窗口周数（默认 12 ≈ 一个季度）。"
                         "**别为了省事调回 4**：4 周是 2026-08-01 为「用均值推算安全库存」"
                         "定的口径，那结论仍成立；但峰值识别需要更长窗口，4 个数据点看不出尖峰")
    ap.add_argument("--since", help="窗口起（YYYY-MM-DD），给了就覆盖 --weeks")
    ap.add_argument("--until", help="窗口止（YYYY-MM-DD，默认今天）")
    ap.add_argument("--states", default=",".join(SOLD_STATES),
                    help="计入销量的订单状态，逗号分隔（默认 sale,done）")
    ap.add_argument("--ref-contains", default="",
                    help="只算 Order Reference 含这些子串的订单，逗号分隔、OR 语义。"
                         "例：--ref-contains VO,GW 只看天猫 C 端，剔除 S0 开头的 B 端单")
    ap.add_argument("--ref-excludes", default="",
                    help="排除 Order Reference 含这些子串的订单，逗号分隔、AND 语义")
    ap.add_argument("--peak-cover", type=float, default=1.0,
                    help="峰值缺口按几个峰值周算（默认 1.0，即至少接得住一个高峰周）")
    ap.add_argument("--violations", help="天猫「供应商违规数据」导出（.xlsx，选填）。"
                    "给了就把处罚历史接进报表，并多产一张 缺货违规风险清单.csv"
                    "——按实际罚款决定先配哪些 SKU 的安全库存")
    ap.add_argument("--violation-since", help="只算这个日期之后的处罚（YYYY-MM-DD）")
    ap.add_argument("--exclude-vendor", default=None,
                    help="算在途/预测时排除这些供应商的采购单（名字子串，逗号分隔）。"
                         "默认读 config.py 的 ODOO_EXCLUDE_VENDORS。用于剔除"
                         "「建虚拟库存映射外部平台」那类永不收货的测试单")
    ap.add_argument("--by-warehouse", action="store_true", help="在手库存按仓拆成多列")
    ap.add_argument("--orderpoint-agg", choices=["sum", "max"], default="sum",
                    help="同一产品多条补货规则的合并方式（默认 sum）")
    ap.add_argument("--company-id", type=int, help="多公司环境锁定公司（allowed_company_ids）")
    ap.add_argument("--external-id", action="store_true",
                    help="附带 External ID 列（将来接 sales_insight 回写的映射码）")
    ap.add_argument("--csv-only", action="store_true", help="只写 CSV，不写 xlsx（cron 用）")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    ap.add_argument("--timeout", type=int, default=180)
    a = ap.parse_args()

    until = a.until or date.today().isoformat()
    if a.since:
        since = a.since
        weeks = max((date.fromisoformat(until) - date.fromisoformat(since)).days / 7, 1e-9)
    else:
        since = (date.fromisoformat(until) - timedelta(weeks=a.weeks)).isoformat()
        weeks = a.weeks
    states = [s.strip() for s in a.states.split(",") if s.strip()]
    contains = [x.strip() for x in a.ref_contains.split(",") if x.strip()]
    if a.exclude_vendor is None:
        from common import localconf
        exclude_vendors = list(localconf.get("ODOO_EXCLUDE_VENDORS", []) or [])
    else:
        exclude_vendors = [x.strip() for x in a.exclude_vendor.split(",") if x.strip()]
    excludes = [x.strip() for x in a.ref_excludes.split(",") if x.strip()]
    ctx = {"allowed_company_ids": [a.company_id]} if a.company_id else {}

    try:
        od = Odoo.connect(timeout=a.timeout, context=ctx)
        say(f"窗口 {since} ~ {until}（{weeks:g} 周）")
        if contains or excludes:
            say(f"渠道过滤 Order Reference: 含 {contains or '—'}"
                + (f"、不含 {excludes}" if excludes else ""))
        viol = None
        if a.violations:
            # 未显式给 --violation-since 就跟随销量窗口起点：两个口径别各说各话
            vs = a.violation_since or since
            viol = read_violations(a.violations, vs)
        df, safety_src = build(od, since, until, weeks, states,
                               a.by_warehouse, a.orderpoint_agg, a.external_id,
                               contains, excludes, exclude_vendors, viol, a.peak_cover)
    except OdooError as e:
        raise SystemExit(f"✗ {e}")

    outdir = a.outdir or os.path.join("output", f"{date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)
    csv_path = os.path.join(outdir, "库存周报.csv")
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")   # -sig: Windows Excel 不乱码
    say(f"\n✓ {csv_path}  {len(df)} 行")
    if not a.csv_only:
        path, n = write_simple(df, outdir, "库存周报.xlsx", left_cols={"商品名称", "SKU"})
        say(f"✓ {path}  {n} 行")

    mm, n_only_a = mismatch(df)
    mm_path = os.path.join(outdir, "安全库存不一致.csv")
    mm.to_csv(mm_path, index=False, encoding="utf-8-sig")
    say(f"✓ {mm_path}  {len(mm)} 行"
        + ("（两路无冲突）" if mm.empty else "  ← 有冲突，先看这张"))

    if viol is not None:
        vr = violation_risk(df)
        vr_path = os.path.join(outdir, "缺货违规风险清单.csv")
        vr.to_csv(vr_path, index=False, encoding="utf-8-sig")
        top = vr[vr["优先级"] == "被罚过·未配安全库存"]
        say(f"✓ {vr_path}  {len(vr)} 行  ← **最该先看这张**")
        say(f"    其中「被罚过·未配安全库存」{len(top)} 个 SKU，"
            f"历史缺货罚款 {top['缺货罚款EUR'].sum():,.0f} EUR / {int(top['缺货违规单数'].sum())} 单")

    un = unconfigured(df)
    un_path = os.path.join(outdir, "安全库存待配清单.csv")
    un.to_csv(un_path, index=False, encoding="utf-8-sig")
    say(f"✓ {un_path}  {len(un)} 行  ← 有销量但没配安全库存，按销量降序")

    say(f"\n安全库存 A 来源: {safety_src or '未找到'}"
        f"  |  B 来源: stock.warehouse.orderpoint.product_min_qty"
        + ("(trigger=auto)" if "trigger" in od.fields_of("stock.warehouse.orderpoint") else ""))
    n_a = int(df["安全库存_产品字段"].notna().sum())
    n_b = int(df["安全库存_补货规则"].notna().sum())
    say(f"配了安全库存: A {n_a} 个 / B {n_b} 个"
        f"（其中 {n_only_a} 个只有 A——B 没在用时这是常态，不计入不一致）")
    if "销量_全渠道" in df.columns:
        f, o = df["销量"].sum(), df["销量_其它渠道"].sum()
        say(f"渠道拆分: 筛选后 {f:,.0f} 件 / 其它渠道 {o:,.0f} 件 / 合计 {f + o:,.0f} 件"
            f"（排名与 ABC 按筛选后算；可撑周数按合计算）")
    sold = int((df["销量"] > 0).sum())
    top50 = int(df.head(50)["安全库存_取用"].notna().sum())
    # 峰值倍数只在**有一定量**的 SKU 上有意义：只卖过一周的长尾货，峰值倍数天然等于
    # 窗口周数（卖 1 件也是 12×），拿它们算中位数会把结论吹大。门槛：≥3 个有销周 且 ≥20 件。
    pk = df[df["峰值倍数"].notna() & (df["有销周数"].fillna(0) >= 3) & (df["销量"] >= 20)]
    if len(pk):
        say(f"需求尖峰: 样本 {len(pk)} 个 SKU（≥3 个有销周且窗口内 ≥20 件；"
            f"长尾货的峰值倍数是噪声，不计入）")
        say(f"  峰值/周均中位数 {pk['峰值倍数'].median():.2f}×；"
            f"≥2× 的 {int((pk['峰值倍数']>=2).sum())} 个（{(pk['峰值倍数']>=2).mean()*100:.0f}%）、"
            f"≥3× 的 {int((pk['峰值倍数']>=3).sum())} 个")
        say(f"  接不住一个峰值周（峰值缺口 >0）: 全表 {int(df['峰值缺口'].notna().sum())} 个 SKU，"
            f"其中样本内 {int(pk['峰值缺口'].notna().sum())} 个")
    say(f"低于安全库存: 按在手 {int(df['缺口'].notna().sum())} 个 SKU / "
        f"按预测 {int(df['预测缺口'].notna().sum())} 个 SKU"
        f"  |  有销量 {sold} 个，其中销量前 50 名配了安全库存的只有 {top50} 个")
    say(f"RPC 调用 {od.call_count} 次")


if __name__ == "__main__":
    main()
