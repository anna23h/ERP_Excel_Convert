#!/usr/bin/env python3
"""询价直通看板 · 产品字典导出（ISSUES A）

从 Odoo 一次性**离线**导出看板用的产品查找表。与「页面接入 ERP API」无关——
artifact 的 CSP 封死外部主机、`network` 能力不可用，页面永远不能自己调 ERP。
本脚本跑在本机，产出一个 JSON 文件，供粘贴导入时在本地查名用。

范围：**只做药房产品**（`pzn` 非空）。日化品没有 PZN，看板不覆盖（2026-08-25 用户拍板）。

中文名的来源不是 `product.product.name`——实测 10298 个可售产品里 zh_CN 的 name
只有 6 条真含中文，Odoo 标准产品档案上**没有**中文品名。真正的来源是
voyageone 模块的 `product.voyageone`（模型标签 "vo product"，7484 条），
它有 `product_id → product.product` 的真关联字段，关联率 100%。

三个坑（都是实测踩出来的，别去掉）：
  1. `product.voyageone.name` 常挂 VO 平台状态前缀「不可售」「不采货」，
     以及德语下架标记 `(Delisted)` / `(Außer Handel)`。不剥掉的话，
     4799 条「含中文」里有 535 条其实只是被前缀污染的德语名。
  2. `sale.order.line.vo.sku` 有隐形脏数据——实测有个 sku 开头是**零宽空格**
     (`​`)，肉眼看不出，精确匹配必然失效。故统一做 `_clean_key()` 归一。
  3. **必须排除天猫 C 端变体**（业务事实，2026-08-25 用户确认）：进这个看板的产品**必然是
     单件**；多品装与带后缀的产品（`_VO`/`_GW`/`x2`/`[N Packs]`）只存在于天猫 C 端，
     绝不会出现在询价看板里。不排除的话 PZN 会撞号——实测原始 6163 条里 **352 个 PZN 撞号**，
     其中 `[2 Packs]ANTIHYDRAL Salbe 70g` 与 `ANTIHYDRAL Salbe 70g` 同为 PZN `00052729`，
     量和价完全不同，混进来会把采购量算错。
     **排除变体后 352 → 11**，PZN 基本可作查找键。
     （组合装 `x2`/`_GW` 是 phantom BoM 套件，见根 README 库存口径那节。）
  4. **剩余 11 个撞号是同一单件产品的重复建档**，不是组合装——一条用旧内部码、一条用 PZN 码
     （`Betaisodona_00721478` vs `Betaisodona_01931491`；`Fenistil Gel 30 G (new PZN12550409)`
     vs `Fenistil Gel 30g (old)`）。挑哪条拿到的都是同一件实物，风险只是挑到已停用的那条。
     故**不静默挑一条**，写进产出的 `pznCollisions` 交导入时告警。
  5. 取 `ir.model` 时**不要**带 `modules` 字段——它会触发 `ir.module.module`
     的权限检查，而拉数账号不在 Administration/Settings 组里，直接被拒。

两份产出（ISSUES C-1）：
  · `product_dict.json`      全量，本机查名用（约 1.5MB）
  · `product_dict.min.json`  瘦身版，**粘进看板设置里的就是这份**（约 450KB）。
    形态是 {"v":1,"at":导出时间,"d":{PZN:[德语名, 中文名]}}——看板只在「新建 / 导入」
    那一刻用它补品名，不进发布文档、只存各人浏览器本地。

跑法：
    python3 dashboard/export_product_dict.py
    python3 dashboard/export_product_dict.py --out dashboard/data/product_dict.json
    python3 dashboard/export_product_dict.py --include-unsellable
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from odoo_api.odoo_client import Odoo, m2o_id  # noqa: E402
from reorder.reorder_helper import norm_pzn as _norm_pzn  # noqa: E402

CJK = re.compile(r"[一-鿿]")

#: VO 平台状态前缀 + 德语下架标记。反复剥直到不变（实测有叠加两层的）。
NOISE = re.compile(
    r"^\s*(?:不可售|不采货|不可賣|已停产|停产"
    r"|\(Delisted\)|\(Außer\s*Handel\)|\(Ausser\s*Handel\))\s*",
    re.I,
)
#: 天猫 C 端变体的形态。看板只做单件，这些一律排除（见 docstring 坑 3）。
#: `_VO` = VoyageOne（天猫），`_GW` = 另一店，`x2`/`[2 Packs]` = 多品装。
VARIANT_CODE = re.compile(r"(_VO|_GW)$|x\d+$|_\d+x\d+$", re.I)
VARIANT_NAME = re.compile(r"^\s*[\[(]\s*\d+\s*Packs?\s*[\])]", re.I)

#: 零宽字符：零宽空格 / 零宽非断空格 / BOM。sku 里实测出现过。
ZERO_WIDTH = re.compile(r"[​‌‍﻿]")

BATCH = 2000


def strip_noise(s):
    """剥掉状态前缀，反复剥到不变。剥完两端再 strip。"""
    s = (s or "").strip()
    prev = None
    while prev != s:
        prev = s
        s = NOISE.sub("", s).strip()
    return s


def is_tmall_variant(default_code, name_de):
    """判天猫 C 端变体（多品装 / 带店铺后缀）。看板只做单件，这些一律排除。"""
    return bool(VARIANT_CODE.search(_clean_key(default_code))) or \
           bool(VARIANT_NAME.search((name_de or "").strip()))


def _clean_key(s):
    """匹配键归一：去零宽字符、去两端空白。不做大小写折叠（code 是区分大小写的）。"""
    return ZERO_WIDTH.sub("", (s or "")).strip()


def norm_pzn(v):
    """PZN 归一。**复用 `reorder/reorder_helper.norm_pzn`**——那边已经处理过
    `前缀_PZN(x件装)` / 显式 `PZN-` / 整格 7~8 位数字三种形态并补零到 8 位，
    没必要在这里另写一套（两套归一规则迟早会漂）。
    差别只有一处：先去零宽字符，否则 `norm_pzn` 里的 `fullmatch` 会被隐形字符打掉。
    抓不到 PZN 的原值保留（留给人工看，不静默丢弃）。"""
    s = _clean_key(str(v or ""))
    if not s:
        return None
    return _norm_pzn(s) or s


def pull_all(od, model, fields, domain=None):
    """search_read 分页拉全量。一次拉一万多行会被截断或超时，必须 limit/offset 循环。"""
    out, off = [], 0
    while True:
        chunk = od.execute(model, "search_read", [domain or [], fields],
                           {"limit": BATCH, "offset": off, "order": "id"})
        if not chunk:
            break
        out += chunk
        off += len(chunk)
        if len(chunk) < BATCH:
            break
    return out


def build(od, include_unsellable=False, keep_variants=False):
    # ---- 1. 底表：药房产品（pzn 非空）--------------------------------------
    dom = [("pzn", "!=", False)]
    if not include_unsellable:
        dom.append(("sale_ok", "=", True))
    prods = pull_all(od, "product.product",
                     ["id", "pzn", "name", "default_code", "barcode", "sale_ok"], dom)
    print(f"  药房产品（pzn 非空{'' if include_unsellable else ' + 可售'}）: {len(prods)}")
    if not keep_variants:
        before = len(prods)
        prods = [p for p in prods
                 if not is_tmall_variant(p.get("default_code"), p.get("name"))]
        print(f"  排除天猫 C 端变体（多品装/带店铺后缀，看板只做单件）: -{before - len(prods)}"
              f" → {len(prods)}")

    by_id = {p["id"]: p for p in prods}

    # ---- 2. 中文名：product.voyageone -------------------------------------
    vo = pull_all(od, "product.voyageone", ["id", "name", "code", "product_id"])
    zh_by_pid, code_by_pid = {}, {}
    for r in vo:
        pid = m2o_id(r.get("product_id"))
        if not pid or pid not in by_id:
            continue
        name = strip_noise(r.get("name"))
        code = _clean_key(r.get("code"))
        if code:
            code_by_pid.setdefault(pid, code)
        if CJK.search(name):
            # 同一产品可能有多条 vo 记录，取最长的那条（信息最全）
            if len(name) > len(zh_by_pid.get(pid, "")):
                zh_by_pid[pid] = name
    print(f"  product.voyageone: {len(vo)} 条，命中药房产品并带中文名 {len(zh_by_pid)}")

    # ---- 3. 别名：销售订单行上的实际叫法 -----------------------------------
    #  sale.order.line.vo.sku ←→ product.voyageone.code ←→ product_id
    pid_by_code = {}
    for r in vo:
        pid = m2o_id(r.get("product_id"))
        code = _clean_key(r.get("code"))
        if pid and code:
            pid_by_code.setdefault(code, pid)
    lines = pull_all(od, "sale.order.line.vo", ["id", "sku", "title"])
    alias_by_pid = {}
    unmatched = set()
    for r in lines:
        title = strip_noise(r.get("title"))
        if not CJK.search(title):
            continue
        sku = _clean_key(r.get("sku"))
        pid = pid_by_code.get(sku)
        if not pid:
            unmatched.add(sku)
            continue
        if pid in by_id:
            alias_by_pid.setdefault(pid, set()).add(title)
    print(f"  sale.order.line.vo: {len(lines)} 行，命中药房产品的别名 {len(alias_by_pid)} 个产品"
          f"（{sum(len(v) for v in alias_by_pid.values())} 条），sku 对不上 {len(unmatched)} 个")

    # ---- 4. 组装 -----------------------------------------------------------
    items = []
    for p in prods:
        pid = p["id"]
        zh = zh_by_pid.get(pid)
        aliases = sorted(alias_by_pid.get(pid, set()) - ({zh} if zh else set()))
        items.append({
            "pzn": norm_pzn(p.get("pzn")),
            "nameDe": (p.get("name") or "").strip(),
            "nameZh": zh,
            "aliases": aliases,
            "odooId": pid,
            "defaultCode": _clean_key(p.get("default_code")) or None,
            "voCode": code_by_pid.get(pid),
            "saleOk": bool(p.get("sale_ok")),
            "source": "erp",
        })
    items.sort(key=lambda x: (x["pzn"] or "", x["nameDe"]))

    # PZN 撞号检查——同一 PZN 对应多个 Odoo 产品的话，导入时按 PZN 定位会有歧义
    seen, dup = {}, {}
    for it in items:
        if it["pzn"] in seen:
            dup.setdefault(it["pzn"], [seen[it["pzn"]]]).append(it["odooId"])
        else:
            seen[it["pzn"]] = it["odooId"]
    return items, dup


def main():
    ap = argparse.ArgumentParser(
        description="导出看板用的产品字典（只读；只做药房产品，且只做单件——天猫 C 端变体默认排除）")
    ap.add_argument("--out", default="dashboard/data/product_dict.json", help="产出路径")
    ap.add_argument("--include-unsellable", action="store_true",
                    help="连 sale_ok=False 的一起导（默认只导可售）")
    ap.add_argument("--keep-variants", action="store_true",
                    help="保留天猫 C 端变体（多品装/带店铺后缀）。默认排除——看板只做单件")
    args = ap.parse_args()

    od = Odoo.connect()
    items, dup = build(od, args.include_unsellable, args.keep_variants)

    with_zh = sum(1 for i in items if i["nameZh"])
    with_alias = sum(1 for i in items if i["aliases"])
    payload = {
        "schema": 1,
        "exportedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
        "scope": "药房产品（pzn 非空）" + ("" if args.include_unsellable else " + 可售"),
        "stats": {"total": len(items), "withNameZh": with_zh, "withAliases": with_alias,
                  "duplicatePzn": len(dup)},
        # 剩余撞号：同一单件产品的重复建档。导入时按 PZN 命中这些要告警交人工选，不可静默取一条。
        "pznCollisions": {k: v for k, v in sorted(dup.items())},
        "products": items,
    }
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)

    # 瘦身版：看板设置里粘的就是这份。紧凑数组去掉 key 开销，体积约为全量的 1/3。
    slim = {"v": 1, "at": payload["exportedAt"],
            "d": {i["pzn"]: [i["nameDe"], i["nameZh"] or ""] for i in items if i["pzn"]}}
    slim_path = args.out.replace(".json", ".min.json")
    with open(slim_path, "w", encoding="utf-8") as f:
        json.dump(slim, f, ensure_ascii=False, separators=(",", ":"))

    size = os.path.getsize(args.out) / 1024
    print(f"\n产出 {args.out}  ({size:.0f} KB)")
    print(f"     {slim_path}  ({os.path.getsize(slim_path)/1024:.0f} KB"
          f"，{len(slim['d'])} 条——粘进看板「字典」设置里的是这份)")
    print(f"  产品        : {len(items)}")
    print(f"  有中文名    : {with_zh}  ({with_zh*100//max(len(items),1)}%)")
    print(f"  有别名      : {with_alias}")
    if dup:
        print(f"  ⚠ PZN 撞号  : {len(dup)} 个——同一单件产品的重复建档（旧内部码 vs PZN 码），"
              f"已写进产出 pznCollisions，导入时须人工确认，勿静默取一条")
        for pzn, ids in list(dup.items())[:5]:
            print(f"      {pzn} → odooId {ids}")
    print(f"XML-RPC 调用次数: {od.call_count}")


if __name__ == "__main__":
    main()
