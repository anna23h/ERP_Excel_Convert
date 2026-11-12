"""询价单回填：Anna 从 Odoo 导出的空询价单（purchase.order）+ 本周备货汇总 → 可导入的询价单表。

    python3 odoo_api/quote_fill.py <空询价单.xlsx>                  # 自动用 output/ 下最新一份备货汇总
    python3 odoo_api/quote_fill.py T.xlsx --summary output/20261111/备货汇总-20261111.xlsx
    python3 odoo_api/quote_fill.py T.xlsx --po P12400               # 模板含多张 PO 时指定
    python3 odoo_api/quote_fill.py T.xlsx -o /tmp/x                 # 输出目录（默认 output/YYYYMMDD）

流程（ISSUES [备货提醒] A，2026-11-08 用户定）：Odoo 对脚本**只读**——
① Anna 在 Odoo 建询价单、选好供应商、不加行，导出 Excel → ② 本脚本按汇总补产品行、填数量与单价 →
③ Anna 核对后手动导入 Odoo。

收哪些行（`restock.py` 现算，见汇总的「询价单供应商」sheet）——建议数量 > 0，且满足其一：
  · 独家：近 12 个月采购记录里只有这一家供应商（ISSUES [备货提醒] A）；
  · 近期最低价：近 `--low-months`（3）个月最低净价那家是这一家，多家并列不算
    （ISSUES [备货提醒] D，2026-11-12 用户定，对所有供应商通用）。
其余产品在 A3000 比价下单，不走这里。

填值口径：
    Order Lines/Product/ID   ← 产品 __export__ 外部 ID（模板只有 `.id` 列则写数据库 ID）
    Order Lines/Quantity     ← 汇总「建议数量」（> 0 的行才补）
    Order Lines/Unit Price   ← 独家品：汇总「最近一次净价」（单一供应商即该供应商的上次价）；
                               近期最低价品：「近期该家最近净价」（该家窗口内最近一次，不用最低那笔——
                               2026-11-12 用户定，避开一次性低价）。都是去税净价
    Order Lines/Taxes[/ID]   ← 产品主数据的进项税（7% / 19% Vorsteuer），不是这两档就停
模板要求同 PO 导入表（`po_import/po_from_pdf.py`）：第 2 行起带 ID / Order Reference / Vendor，
勾上 Product / Quantity / Unit Price / Taxes 列；界面名、技术名表头都认。
任一行对不上（供应商不是询价单供应商、产品没外部 ID、税不是两档之一、没有上次价）就整张不出、退出码 2。
"""
import argparse
import glob
import os
import sys
from datetime import date, datetime

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from common.xlsx import unique_path  # noqa: E402
from odoo_api.restock import (COL_LOW_VENDOR, COL_QUOTE_VENDOR, COL_SUGGEST,  # noqa: E402
                              MAIN_SHEET, quote_rows, rfq_vendors)
from po_import.po_from_pdf import (TAX_FULL, TAX_REDUCED, blank_reason,  # noqa: E402
                                   odoo_po_status, read_template)
from so_import.so_from_pdf import StopError  # noqa: E402

STALE_DAYS = 7


def latest_summary(root="output"):
    """output/*/备货汇总-YYYYMMDD.xlsx 里按文件名日期取最新一份。"""
    hits = sorted(glob.glob(os.path.join(root, "*", "备货汇总-*.xlsx")),
                  key=lambda p: (os.path.basename(p), p))
    if not hits:
        raise StopError("output/ 下找不到 备货汇总-*.xlsx，先跑 python3 odoo_api/restock.py，或用 --summary 指定")
    return hits[-1]


def summary_date(path):
    stem = os.path.splitext(os.path.basename(path))[0]
    try:
        return datetime.strptime(stem.split("-")[-1][:8], "%Y%m%d").date()
    except ValueError:
        return None


def match_vendor(vendor, candidates):
    """模板里的 Vendor → 汇总里的询价单供应商。先比全名（忽略大小写、首尾空白），再看唯一的子串命中。"""
    v = str(vendor or "").strip().casefold()
    exact = [c for c in candidates if c.strip().casefold() == v]
    if exact:
        return exact[0]
    sub = [c for c in candidates if v and (v in c.casefold() or c.casefold() in v)]
    if len(sub) == 1:
        return sub[0]
    raise StopError(f"询价单的供应商「{vendor}」不是本周汇总里的询价单供应商"
                    + (f"（子串命中多家 {sub}）" if sub else "")
                    + f"。询价单供应商（有独家品或本周需补品近期最低价是它）：{sorted(candidates)}。"
                    "其余供应商的产品在 A3000 比价下单，不走询价单。")


def pick_rows(summary, vendor):
    """汇总里该供应商要补的行（独家 + 近期最低价），带 _price / _src 两列。"""
    df, price, src = quote_rows(summary, vendor)
    if df.empty:
        raise StopError(f"本周汇总里 {vendor} 没有需补的品（独家品建议数量全为 0、也没有近期最低价是它的需补品），不出表。")
    df = df.assign(_price=price, _src=src)
    noprice = df[df["_price"].isna() | (df["_price"] <= 0)]
    if not noprice.empty:
        raise StopError(f"这些行没有该家的采购净价，没法填单价：{noprice['SKU'].tolist()}")
    return df


def tax_of(names):
    """产品进项税名称元组 → (外部 ID, 名称)。只认 7% / 19% Vorsteuer 两档且必须恰好一个。"""
    known = {TAX_REDUCED[1]: TAX_REDUCED, TAX_FULL[1]: TAX_FULL}
    hit = [known[n] for n in names if n in known]
    if len(names) == 1 and len(hit) == 1:
        return hit[0]
    return None


def odoo_products(pids):
    """Odoo 只读：产品 id → dict(xmlid, code, taxes 名称元组)。"""
    from po_import.po_from_pdf import _odoo
    od = _odoo()
    prods = od.execute("product.product", "read", [sorted(pids)],
                       {"fields": ["default_code", "supplier_taxes_id"]})
    xids = od.execute("ir.model.data", "search_read",
                      [[["model", "=", "product.product"], ["module", "=", "__export__"],
                        ["res_id", "in", sorted(pids)]]], {"fields": ["res_id", "name"]})
    by_res = {}
    for x in xids:
        by_res.setdefault(x["res_id"], []).append(f"__export__.{x['name']}")
    tax_ids = sorted({t for p in prods for t in p["supplier_taxes_id"]})
    tax_names = {t["id"]: t["name"] for t in od.execute(
        "account.tax", "read", [tax_ids], {"fields": ["name"]})} if tax_ids else {}
    return {p["id"]: {"xmlids": by_res.get(p["id"], []), "code": p["default_code"] or "",
                      "taxes": tuple(sorted(tax_names[t] for t in p["supplier_taxes_id"]))}
            for p in prods}


def resolve_lines(rows, prods, need_xmlid):
    """汇总行 + Odoo 产品信息 → 导入行；有问题的行收齐了一次报。"""
    lines, bad = [], []
    for _, r in rows.iterrows():
        pid = int(r["产品ID"])
        p = prods.get(pid)
        why = []
        if p is None:
            why.append("Odoo 里读不到该产品")
        else:
            if need_xmlid and len(p["xmlids"]) != 1:
                why.append(f"__export__ 外部 ID 有 {len(p['xmlids'])} 个（先在 Odoo 导出一次该产品生成外部 ID）")
            if tax_of(p["taxes"]) is None:
                why.append(f"进项税是 {list(p['taxes']) or '空'}，不是 7% / 19% Vorsteuer 之一")
        if why:
            bad.append(f"{r['SKU']}：" + "；".join(why))
            continue
        lines.append({"sku": r["SKU"], "name": r["产品"], "dbid": pid,
                      "xmlid": p["xmlids"][0] if p["xmlids"] else None,
                      "qty": int(r[COL_SUGGEST]), "price": round(float(r["_price"]), 2),
                      "src": r["_src"], "tax": tax_of(p["taxes"])})
    if bad:
        raise StopError("这些行补不进询价单：\n  " + "\n  ".join(bad))
    return lines


def write_import(wb, cols, head, colmap, lines, out_path):
    """按模板列名列序写：第一行带头信息，其余行只填明细列（同 po_from_pdf.build，没有 Vendor Reference）。"""
    ws = wb.active
    ws.delete_rows(2, ws.max_row)
    for k, l in enumerate(lines):
        vals = {colmap["qty"]: l["qty"], colmap["price"]: l["price"]}
        if colmap["pid"]:
            vals[colmap["pid"]] = l["xmlid"]
        else:
            vals[colmap["pdbid"]] = l["dbid"]
        if colmap["tax_id"]:
            vals[colmap["tax_id"]] = l["tax"][0]
        else:
            vals[colmap["tax"]] = l["tax"][1]
        if k == 0:
            vals.update(head)
        ws.append([vals.get(c) for c in cols])
    wb.save(out_path)


def run(template, summary=None, po=None, outdir=None, status_fn=odoo_po_status, products_fn=odoo_products):
    summary = summary or latest_summary()
    d = summary_date(summary)
    print(f"备货汇总：{summary}")
    if d and (date.today() - d).days > STALE_DAYS:
        print(f"  ⚠ 这份汇总是 {d}，已超过 {STALE_DAYS} 天；建议先重跑 restock.py")
    sdf = pd.read_excel(summary, sheet_name=MAIN_SHEET)
    for c in (COL_QUOTE_VENDOR, COL_LOW_VENDOR):
        if c in sdf:
            sdf[c] = sdf[c].fillna("").astype(str)
    if COL_LOW_VENDOR not in sdf:
        print(f"  ⚠ 这份汇总没有「{COL_LOW_VENDOR}」列（旧版 restock 出的），只收独家品；重跑 restock.py 才会放宽")

    wb, cols, heads, colmap = read_template(template, po)
    if len(heads) > 1:
        raise StopError(f"模板里有 {len(heads)} 张 PO {[h[colmap['ref']] for h in heads]}，用 --po 指定写哪张。")
    head = heads[0]
    ref = head[colmap["ref"]]
    why = blank_reason(status_fn([head[colmap["id"]]]).get(head[colmap["id"]]))
    if why:
        raise StopError(f"{ref} 不是空白询价单（{why}），不往里写。请新建一张、不加行再导出。")

    vendor = match_vendor(head[colmap["vendor"]], rfq_vendors(sdf))
    rows = pick_rows(sdf, vendor)
    lines = resolve_lines(rows, products_fn(set(int(x) for x in rows["产品ID"])), need_xmlid=bool(colmap["pid"]))

    outdir = outdir or os.path.join("output", f"{date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)
    out = unique_path(os.path.join(outdir, f"{ref}_询价单回填.xlsx"))
    write_import(wb, cols, {k: v for k, v in head.items()}, colmap, lines, out)

    n_low = sum(l["src"] == "近期最低价" for l in lines)
    print(f"\n{ref} · {vendor} · {len(lines)} 行（独家 {len(lines) - n_low} · 近期最低价 {n_low}）")
    print(f"{'SKU':<28}{'数量':>6}{'净价':>10}{'金额':>11}  {'来源':<8}税")
    for l in lines:
        print(f"{l['sku']:<28}{l['qty']:>6}{l['price']:>10.2f}{l['qty'] * l['price']:>11.2f}  "
              f"{l['src']:<8}{l['tax'][1]}")
    print(f"{'合计（去税）':<26}{sum(l['qty'] for l in lines):>6}{'':>10}"
          f"{sum(l['qty'] * l['price'] for l in lines):>11.2f}")
    print(f"\n✓ {out}\n  核对数量与单价后在 Odoo 导入（Purchase → 询价单 → 导入）。")
    return out, lines


def main():
    ap = argparse.ArgumentParser(description="空询价单 + 备货汇总 → 可导入的询价单表（Odoo 只读）")
    ap.add_argument("template", help="Odoo 导出的空询价单 purchase.order xlsx（第 2 行起带 ID / Order Reference / Vendor）")
    ap.add_argument("--summary", help="备货汇总 xlsx（默认 output/ 下最新一份）")
    ap.add_argument("--po", metavar="PO号", help="模板含多张 PO 时指定（Order Reference）")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    a = ap.parse_args()
    try:
        run(a.template, a.summary, a.po, a.outdir)
    except StopError as e:
        print(f"\n✗ 停：{e}", file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
