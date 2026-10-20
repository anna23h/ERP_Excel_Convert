"""供应商 Rechnung/Lieferschein PDF → Odoo 采购单导入表（填进 Odoo 导出的 purchase.order 模板）。

用法：
    python3 -m po_import.po_from_pdf <Lieferschein.pdf> <purchase.order 模板.xlsx>
    python3 -m po_import.po_from_pdf L.pdf T.xlsx --exclude <PZN> <PZN>
    python3 -m po_import.po_from_pdf L.pdf T.xlsx --pick <PZN>=__export__.product_product_<…>
    python3 -m po_import.po_from_pdf L.pdf T.xlsx -o /tmp/x      # 输出目录（默认 output/YYYYMMDD）

同构 `so_import/so_from_pdf.py`，规矩照搬：**PZN 是唯一连接键，绝不按品名兜底**；
匹配不到 / 命中多个产品就整张表不出、退出码 2，剔除（`--exclude`）与多选一（`--pick`）由用户定。

填值口径（2026-10-20 用户定，与该供应商历史 PO 的手工口径一致）：
    Order Lines/Product/ID   ← 匹配到的外部 ID
    Order Lines/Quantity     ← PDF `Menge`
    Order Lines/Unit Price   ← PDF `AEP`（净价，不乘税）
    Order Lines/Taxes[/ID]   ← Code 含 `M`（MwSt. ermäßigt）→ 7% Vorsteuer，否则 19% Vorsteuer
    Vendor Reference         ← PDF 单号（模板有此列才填）
列名列序照抄模板；第一行带模板第 2 行的 ID / Order Reference / Vendor，其余行头信息列留空。
表头界面名与技术名（`id` / `name` / `order_line/product_qty` …）都认；产品列若是 `order_line/product_id/.id`
就写数据库 ID。

PDF 版式**只支持一家**（明细行 `PZN Menge Packung DarForm 名称 Code AEP AVP Gesamt`，
底部 `Anz. Pos.` / 两档税基 / `Warenwert gesamt` / `Endsumme`）。解析护栏，任一不符就拒绝出表：
    · 逐行 AEP × Menge = Gesamt
    · Σ Gesamt = 印的 Warenwert gesamt = Endsumme
    · 行数 / 总件数 = 印的 Anz. Pos.（如 `9/98`）
    · Code 含 M 的行 Gesamt 合计 = 印的 Warenwert ermäß. MwSt，其余 = volle MwSt
      ——税率归类有独立校验；品名末尾恰是单个字母被误读成 Code 也会在这里被拦。
防重：Odoo 里已有 Vendor Reference = 本单号的 PO 就停。
"""
import argparse
import datetime as dt
import os
import re
import sys
from decimal import Decimal

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

from openpyxl import load_workbook

from common.xlsx import unique_path
from so_import.so_from_pdf import StopError, apply_exclude, de_num, pzn8, read_pdf

HEAD_COLS = ("ID", "Order Reference", "Vendor")
COL_VREF = "Vendor Reference"
COL_PID = "Order Lines/Product/ID"
COL_QTY = "Order Lines/Quantity"
COL_PRICE = "Order Lines/Unit Price"
COL_TOTAL_QTY = "Order Lines/Total Quantity"
COL_TAX_ID = "Order Lines/Taxes/ID"
COL_TAX = "Order Lines/Taxes"

#: 角色 → 模板里认的列名：(界面名, 技术名)。Odoo 导出可选「界面名」或「技术名」两种表头，导入两种都认
#: （2026-10-20 用户第 28 份模板就是技术名）。`pdbid` 是产品的数据库 ID（`.id`），只有技术名形式。
ALIASES = {
    "id": ("ID", "id"),
    "ref": ("Order Reference", "name"),
    "vendor": ("Vendor", "partner_id"),
    "vref": (COL_VREF, "partner_ref"),
    "pid": (COL_PID, "order_line/product_id/id"),
    "pdbid": (None, "order_line/product_id/.id"),
    "qty": (COL_QTY, "order_line/product_qty"),
    "total_qty": (COL_TOTAL_QTY, "order_line/product_uom_qty"),
    "price": (COL_PRICE, "order_line/price_unit"),
    "tax_id": (COL_TAX_ID, "order_line/taxes_id/id"),
    "tax": (COL_TAX, "order_line/taxes_id"),
}

#: 税：(外部 ID, 名称)。德国 SKR03 科目表自带的两档进项税，2026-10-20 在 Odoo 查实（id 24 / 23）。
TAX_REDUCED = ("l10n_de_skr03.1_tax_vst_7_skr03", "7% Vorsteuer")
TAX_FULL = ("l10n_de_skr03.1_tax_vst_19_skr03", "19% Vorsteuer")

_NUM = r"\d{1,3}(?:\.\d{3})*,\d{2}"
#: 明细行：PZN Menge Packung DarForm 名称 Code(一个或多个 * E M S T V) AEP AVP Gesamt
ROW_RE = re.compile(rf"^(\d{{6,8}}) (\d+) (\S+) (\S+) (.+?) ((?:[*EMSTV] )*[*EMSTV]) ({_NUM}) ({_NUM}) ({_NUM})$")
#: 页眉第一行：单号 页码
DOC_RE = re.compile(r"^(\d{7}) (\d+)$")
#: 底部合计行：…Anz. Pos.(行数/件数) volle MwSt € ermäß. MwSt € gesamt €（€ 常被抽成乱码，按任意非空白吃掉）
FOOT_RE = re.compile(rf"(?<!\S)(\d+)/(\d+) ({_NUM}) \S+ ({_NUM}) \S+ ({_NUM})(?: \S+)?$")
END_RE = re.compile(rf"Endsumme \*+ ({_NUM})")


def parse_text(page_texts):
    """各页文本 → (明细行列表, 单号)。行内含 pzn/name/qty/aep/gesamt/reduced/page。"""
    lines, docs, foot, endsumme = [], set(), None, None
    for i, text in enumerate(page_texts, 1):
        for ln in (text or "").splitlines():
            ln = ln.strip()
            m = ROW_RE.match(ln)
            if m:
                lines.append(dict(page=i, pzn=pzn8(m[1]), qty=int(m[2]), name=f"{m[3]} {m[4]} {m[5]}",
                                  code=m[6], reduced="M" in m[6].split(),
                                  aep=de_num(m[7]), gesamt=de_num(m[9])))
                continue
            m = DOC_RE.match(ln)
            if m:
                docs.add(m[1])
                continue
            m = FOOT_RE.search(ln)
            if m:
                foot = dict(npos=int(m[1]), nqty=int(m[2]), full=de_num(m[3]),
                            reduced=de_num(m[4]), total=de_num(m[5]))
                continue
            m = END_RE.search(ln)
            if m:
                endsumme = de_num(m[1])
    if not lines:
        raise StopError("PDF 里没解析出任何明细行——版式不认识（本脚本只支持一家的 Rechnung/Lieferschein 版式）。")
    if len(docs) != 1:
        raise StopError(f"页眉单号认不出或不唯一：{sorted(docs)}")
    if foot is None or endsumme is None:
        raise StopError("PDF 里找不到底部合计（Anz. Pos. / Warenwert / Endsumme）——版式不认识，不敢确定明细是否抽全。")
    bad = [l for l in lines if l["aep"] * l["qty"] != l["gesamt"]]
    if bad:
        raise StopError("以下行 AEP × Menge ≠ Gesamt，解析可能错位，不出表：\n" + "\n".join(
            f"  {l['pzn']} | {l['name']} | {l['aep']} × {l['qty']} ≠ {l['gesamt']}" for l in bad))
    got = dict(npos=len(lines), nqty=sum(l["qty"] for l in lines),
               full=sum((l["gesamt"] for l in lines if not l["reduced"]), Decimal(0)),
               reduced=sum((l["gesamt"] for l in lines if l["reduced"]), Decimal(0)),
               total=sum(l["gesamt"] for l in lines))
    diff = {k: (got[k], foot[k]) for k in got if got[k] != foot[k]}
    if foot["total"] != endsumme:
        diff["endsumme"] = (foot["total"], endsumme)
    if diff:
        labels = dict(npos="行数", nqty="总件数", full="volle MwSt 税基", reduced="ermäß. MwSt 税基",
                      total="Warenwert gesamt", endsumme="Warenwert gesamt vs Endsumme")
        raise StopError("抽出的明细与 PDF 印的合计对不上（抽出 vs 印的），说明有行没抽到、抽错了或税码误读，不出表：\n"
                        + "\n".join(f"  {labels[k]}：{a} vs {b}" for k, (a, b) in diff.items()))
    return lines, docs.pop()


def match(lines, candidates, pick):
    """candidates: pzn → odoo_candidates 的元组。pick: pzn → 外部ID。

    命中恰好 1 个才填进行里；否则收进 problems 原样返回，不猜。
    """
    pick = {pzn8(k): v for k, v in pick.items()}
    unknown = set(pick) - {l["pzn"] for l in lines}
    if unknown:
        raise StopError(f"--pick 里的 PZN 不在（剔除后的）PDF 明细上：{sorted(unknown)}")
    problems = []
    for l in lines:
        hits = candidates.get(l["pzn"], [])
        if l["pzn"] in pick:
            chosen = [h for h in hits if h[0] == pick[l["pzn"]]]
            if not chosen:
                raise StopError(f"--pick {l['pzn']}={pick[l['pzn']]}：这个外部 ID 不在该 PZN 的命中里 "
                                f"{[h[0] for h in hits]}")
            hits = chosen
        if len(hits) == 1:
            l["xmlid"], l["erp_name"], _, l["active"], _, l["erp_taxes"], l["dbid"] = hits[0]
        else:
            problems.append((l, hits))
    return problems


def _odoo():
    from odoo_api.odoo_client import Odoo
    return Odoo.connect(context={"active_test": False}, verbose=False)


def odoo_candidates(pzns):
    """Odoo 只读：pzn → [(__export__ 外部 ID, 品名, Internal Reference, active, 历史采购行数,
    默认进项税名称元组, 数据库 ID)]，含已归档品。"""
    od = _odoo()
    prods = od.execute("product.product", "search_read", [[["pzn", "in", sorted(pzns)]]],
                       {"fields": ["id", "name", "pzn", "default_code", "active", "supplier_taxes_id"]})
    ids = [p["id"] for p in prods]
    xids = od.execute("ir.model.data", "search_read",
                      [[["model", "=", "product.product"], ["module", "=", "__export__"], ["res_id", "in", ids]]],
                      {"fields": ["res_id", "name"]}) if prods else []
    by_res = {}
    for x in xids:
        by_res.setdefault(x["res_id"], []).append(f"__export__.{x['name']}")
    counts = {g["product_id"][0]: g["__count"] for g in od.execute(
        "purchase.order.line", "read_group", [[["product_id", "in", ids]]],
        {"fields": ["product_id"], "groupby": ["product_id"], "lazy": False})} if prods else {}
    tax_ids = sorted({t for p in prods for t in p["supplier_taxes_id"]})
    tax_names = {t["id"]: t["name"] for t in od.execute(
        "account.tax", "read", [tax_ids], {"fields": ["name"]})} if tax_ids else {}
    out = {}
    for p in prods:
        xs = by_res.get(p["id"], [])
        if len(xs) != 1:
            raise StopError(f"产品 {p['id']}（PZN {p['pzn']}，{p['name'].strip()}）的 __export__ 外部 ID 有 "
                            f"{len(xs)} 个，没法确定写哪个。先在 Odoo 里导出一次该产品生成外部 ID。")
        out.setdefault(pzn8(p["pzn"]), []).append(
            (xs[0], (p["name"] or "").strip(), p["default_code"] or "", p["active"], counts.get(p["id"], 0),
             tuple(sorted(tax_names[t] for t in p["supplier_taxes_id"])), p["id"]))
    return out


def odoo_existing(doc_no):
    """Odoo 只读：Vendor Reference = 单号的已有 PO → [(单号, 状态)]。"""
    return [(p["name"], p["state"]) for p in _odoo().execute(
        "purchase.order", "search_read", [[["partner_ref", "=", doc_no]]], {"fields": ["name", "state"]})]


def read_template(template):
    """→ (workbook, 列名, 头信息{列名: 值}, colmap{角色: 列名})。缺列或第 2 行头信息不全就停。

    表头界面名（`Order Lines/Quantity`）与技术名（`order_line/product_qty`）都认，见 ALIASES。
    """
    wb = load_workbook(template)
    ws = wb.active
    cols = [c.value for c in ws[1]]
    colmap = {role: next((c for c in names if c and c in cols), None) for role, names in ALIASES.items()}

    def need(*roles):
        return " 或 ".join(f"{n}" for r in roles for n in ALIASES[r] if n)

    missing = [need(r) for r in ("id", "ref", "vendor", "price") if not colmap[r]]
    if not (colmap["pid"] or colmap["pdbid"]):
        missing.append(need("pid", "pdbid"))
    if not colmap["qty"]:
        # 2026-10-20 实测：Total Quantity 是 product_uom_qty，只读计算字段，导进去不改数量
        missing.append(need("qty") + (f"（模板里的 {colmap['total_qty']} 是只读字段、导入无效，导出时改勾 Quantity）"
                                      if colmap["total_qty"] else ""))
    if not (colmap["tax_id"] or colmap["tax"]):
        missing.append(need("tax_id", "tax"))
    if missing:
        raise StopError(f"模板缺列 {missing}；表头实际是 {cols}。在 Odoo 导出时把这些字段勾上重导。")
    head = {colmap[r]: ws.cell(2, cols.index(colmap[r]) + 1).value for r in ("id", "ref", "vendor")}
    if not all(head.values()):
        raise StopError(f"模板第 2 行的 ID / Order Reference / Vendor 不全：{head}")
    return wb, cols, head, colmap


def build(wb, cols, head, colmap, lines, doc_no, out_path):
    """按模板列名列序写导入表：第一行带头信息（+ Vendor Reference），其余行只填明细列。

    产品列优先写外部 ID；模板只有 `.id` 列就写数据库 ID。税列有 `/id` 写外部 ID，否则写名称
    （两档名称在 Odoo 里唯一，2026-10-20 查实）。
    """
    ws = wb.active
    ws.delete_rows(2, ws.max_row)
    for k, l in enumerate(lines):
        tax_xmlid, tax_name = TAX_REDUCED if l["reduced"] else TAX_FULL
        vals = {colmap["qty"]: l["qty"], colmap["price"]: float(l["aep"])}
        if colmap["pid"]:
            vals[colmap["pid"]] = l["xmlid"]
        else:
            vals[colmap["pdbid"]] = l["dbid"]
        if colmap["tax_id"]:
            vals[colmap["tax_id"]] = tax_xmlid
        else:
            vals[colmap["tax"]] = tax_name
        if k == 0:
            vals.update(head)
            if colmap["vref"]:
                vals[colmap["vref"]] = doc_no
        ws.append([vals.get(c) for c in cols])
    wb.save(out_path)


def report(lines, dropped):
    print("\n核对表（PZN / PDF 名称 / 数量 / AEP / Gesamt / 税 / ERP 品名）")
    for k, l in enumerate(lines, 1):
        flag = "" if l["active"] else "  ⚠ 已归档"
        tax = TAX_REDUCED[1] if l["reduced"] else TAX_FULL[1]
        print(f"{k:>3} | {l['pzn']} | {l['name']} | {l['qty']} | {l['aep']} | {l['gesamt']} | {tax} | "
              f"{l['erp_name']}{flag}")
    off = [l for l in lines if l["erp_taxes"] != ((TAX_REDUCED if l["reduced"] else TAX_FULL)[1],)]
    if off:
        # 只提示不拦：表里写的是 PDF 税码对应的税；不一致说明产品主数据的默认进项税可能该改
        print("\n⚠ 以下产品在 Odoo 的默认进项税与 PDF 税码不一致（表里按 PDF 写），请核对产品主数据：")
        for l in off:
            want = (TAX_REDUCED if l["reduced"] else TAX_FULL)[1]
            print(f"    {l['pzn']} | {l['erp_name']} | PDF → {want} | 产品默认 {', '.join(l['erp_taxes']) or '无'}")
    if dropped:
        print("\n已剔除（建档后需手补进这张 PO）：")
        for l in dropped:
            print(f"    {l['pzn']} | {l['name']} | {l['qty']} | Gesamt {l['gesamt']}")
    r = sum((l["gesamt"] for l in lines if l["reduced"]), Decimal(0))
    f = sum((l["gesamt"] for l in lines if not l["reduced"]), Decimal(0))
    print(f"\n合计核对：出表 {len(lines)} 行 {sum(l['qty'] for l in lines)} 件；净额 {r + f}"
          f"（7% 税基 {r} + 19% 税基 {f}）；含税约 {(r * Decimal('1.07') + f * Decimal('1.19')).quantize(Decimal('0.01'))}"
          + (f"；剔除行合计 {sum(l['gesamt'] for l in dropped)}" if dropped else ""))


def run(pdf, template, outdir=None, exclude=(), pick=None,
        candidates_fn=odoo_candidates, existing_fn=odoo_existing):
    wb, cols, head, colmap = read_template(template)
    lines, doc_no = parse_text(read_pdf(pdf))
    print(f"PDF 单号 {doc_no}，明细 {len(lines)} 行 / {sum(l['qty'] for l in lines)} 件"
          f"（页 {sorted({l['page'] for l in lines})}），合计 {sum(l['gesamt'] for l in lines)}")
    dup = existing_fn(doc_no)
    if dup:
        raise StopError(f"Odoo 里已有 Vendor Reference = {doc_no} 的采购单：{dup}——这张很可能已经录过，不出表。")
    lines, dropped = apply_exclude(lines, exclude)
    problems = match(lines, candidates_fn({l["pzn"] for l in lines}), pick or {})
    if problems:
        msg = ["以下行匹配异常，未出表（剔除用 --exclude PZN，多选一用 --pick PZN=外部ID）："]
        for l, hits in problems:
            msg.append(f"  {l['pzn']} | {l['name']} | 数量 {l['qty']} | 命中 {len(hits)} 个")
            msg += [f"      {h[0]}  {h[1]}  [{h[2]}]  历史采购 {h[4]} 行{'' if h[3] else '  已归档'}"
                    for h in hits]
        raise StopError("\n".join(msg))
    outdir = outdir or os.path.join("output", f"{dt.date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(pdf))[0]
    out = unique_path(os.path.join(outdir, f"{head[colmap['ref']]}_import_{stem}.xlsx"))
    build(wb, cols, head, colmap, lines, doc_no, out)
    print(f"写出 {out}（{len(lines)} 行）；头信息 {head}"
          + (f"，Vendor Reference {doc_no}" if colmap["vref"] else "（模板无 Vendor Reference 列，单号没写进去）"))
    report(lines, dropped)
    return out


def main():
    ap = argparse.ArgumentParser(description="供应商 Rechnung/Lieferschein PDF → Odoo 采购单导入表")
    ap.add_argument("pdf", help="供应商 Rechnung/Lieferschein PDF")
    ap.add_argument("template", help="Odoo 导出的 purchase.order 模板 xlsx（第 2 行带 ID / Order Reference / Vendor）")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    ap.add_argument("--exclude", nargs="+", default=[], metavar="PZN", help="剔除这些 PZN 的行（用户裁决后用）")
    ap.add_argument("--pick", nargs="+", default=[], metavar="PZN=外部ID",
                    help="一个 PZN 命中多个产品时指定用哪个（用户裁决后用）")
    a = ap.parse_args()
    pick = {}
    for kv in a.pick:
        if "=" not in kv:
            ap.error(f"--pick 要写成 PZN=外部ID：{kv}")
        k, v = kv.split("=", 1)
        pick[k.strip()] = v.strip()
    try:
        run(a.pdf, a.template, a.outdir, a.exclude, pick)
    except StopError as e:
        print(f"\n✗ {e}", file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
