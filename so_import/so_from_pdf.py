"""客户 Bestellung PDF → Odoo 销售单导入表（填进 Odoo 导出的 SO 模板）。

用法：
    python3 -m so_import.so_from_pdf <Bestellung.pdf> <sale.order 模板.xlsx>
    python3 -m so_import.so_from_pdf B.pdf T.xlsx --exclude <PZN> <PZN>
    python3 -m so_import.so_from_pdf B.pdf T.xlsx --pick <PZN>=__export__.product_product_<…>
    python3 -m so_import.so_from_pdf B.pdf T.xlsx -o /tmp/x      # 输出目录（默认 output/YYYYMMDD）

**PZN 是唯一连接键，绝不按品名兜底**（2026-10-13 用户定）。PDF 的 PZN 补足 8 位，
与 Odoo `product.product.pzn` 精确比对（含已归档品），取该产品的 `__export__` 外部 ID。
任何一行匹配不到或命中多个产品，**整张表不出**、退出码 2、把这些行列出来——
剔除哪行（`--exclude`）、多个里选哪个（`--pick`）是用户的决定，脚本不猜不跳。

**SO 只用基础品**（2026-11-10 用户定，`[SO 导入] C`）：命中里先排除非基础品——挂着**有组件的 phantom
BoM**（套件），或内部编码带渠道 / 组合装后缀（`_VO`/`_GW`/`xN`/`*N`，复用 `common/po._po_base_sku`，
同 PO 导入口径）。空 BoM / normal BoM 不算套件（实测有基础品挂空 normal BoM，按「有 BoM」会误排）。
排除后剩 1 个才自动用；剩 0 个或多个照样停。`--pick` 不能选被排除的产品。

为什么只用 Odoo API、不用产品主数据 xlsx：2026-10-13 实测 0730 的主数据快照
13 个 PZN 里查不到 10 个，API 找回其中 8 个（新建的 `Temp_` 临时品）。快照必然过期。

填值口径：
    Order Lines/Product/ID                 ← 匹配到的外部 ID
    Order Lines/Product/Internal Reference ← 留空（已有 ID，第二把匹配键冗余且易报错）
    Order Lines/Quantity                   ← PDF `Anz.`
    Order Lines/Unit Price                 ← PDF `Netto` × (1 + 产品销项税率)，全精度不四舍五入
列名列序照抄模板；第一行带模板第 2 行的 ID / Order Reference / Customer，其余行这三列留空。

税率现查、不写死（2026-11-10 `[SO 导入] B`，此前写死 × 1.19）：按模板 `Order Reference` 找到 SO
（公司、财政位置），取产品 `taxes_id` 里属于该公司的税。恰好 1 个、百分比、含税口径、且不被财政位置
映射，才按它算；其余（0 个 / 多个税、不含税口径、非百分比、被财政位置替换）整表不出、退出码 2。
非 19% 的单税照算（如 7% → × 1.07），核对表打 ⚠。

PDF 版式**只支持一家**（明细行 `PZN 名称 Netto € Anz. Nettosumme €`，`Summe` 收尾）。
解析的护栏：抽出的行 Nettosumme 合计必须等于印的 Summe，否则视为版式不认识、拒绝出表——
换了别家版式会在这里停下，而不是静默漏行。
"""
import argparse
import datetime as dt
import os
import re
import sys
from decimal import Decimal

# Windows 控制台默认 cp936/cp1252，打印 × ✓ 这类符号会 UnicodeEncodeError 直接崩，强制 UTF-8。
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

from openpyxl import load_workbook

from common.po import _po_base_sku
from common.xlsx import unique_path

STD_RATE = Decimal("19")   # 只用于核对表提示「非标准税率」，不参与计价
HEAD_COLS = ("ID", "Order Reference", "Customer")
COL_PID = "Order Lines/Product/ID"
COL_IREF = "Order Lines/Product/Internal Reference"
COL_QTY = "Order Lines/Quantity"
COL_PRICE = "Order Lines/Unit Price"

#: 明细行：PZN 名称 净价(三位小数) € 数量 行净额 €；行尾常粘着页面右栏的联系方式，不锚行尾。
ROW_RE = re.compile(r"^(\d{6,8}) (.+?) (\d{1,3}(?:\.\d{3})*,\d{2,4}) € (\d+) (\d{1,3}(?:\.\d{3})*,\d{2}) €")
SUM_RE = re.compile(r"^Summe (\d{1,3}(?:\.\d{3})*,\d{2}) €")


class StopError(Exception):
    """要停下来交给用户裁决的情况：不出表，退出码 2。"""


def de_num(s):
    """德式数字 `2.262,53` → Decimal('2262.53')。"""
    return Decimal(s.replace(".", "").replace(",", "."))


def pzn8(v):
    s = str(v).strip()
    return s.zfill(8) if s.isdigit() else s


def parse_text(page_texts):
    """各页文本 → (明细行列表, Summe)。行内含 pzn/name/netto/qty/nsum/page。"""
    lines, summe = [], None
    for i, text in enumerate(page_texts, 1):
        for ln in (text or "").splitlines():
            m = ROW_RE.match(ln)
            if m:
                lines.append(dict(page=i, pzn=pzn8(m[1]), name=m[2], netto=de_num(m[3]),
                                  qty=int(m[4]), nsum=de_num(m[5])))
                continue
            m = SUM_RE.match(ln)
            if m:
                summe = de_num(m[1])
    if not lines:
        raise StopError("PDF 里没解析出任何明细行——版式不认识（本脚本只支持一家的 Bestellung 版式）。")
    if summe is None:
        raise StopError("PDF 里找不到 `Summe` 行——版式不认识，不敢确定明细是否抽全。")
    printed = sum(l["nsum"] for l in lines)
    if printed != summe:
        raise StopError(f"抽出的 {len(lines)} 行 Nettosumme 合计 {printed} ≠ 印的 Summe {summe}，"
                        "说明有行没抽到或抽错了，不出表。")
    return lines, summe


def read_pdf(path):
    try:
        import pdfplumber
    except ImportError:
        sys.exit("缺 pdfplumber：pip install -r requirements-pdf.txt")
    with pdfplumber.open(path) as pdf:
        return [pg.extract_text() for pg in pdf.pages]


def apply_exclude(lines, exclude):
    """剔除用户指定的 PZN；指定了 PDF 里没有的 PZN 视为笔误，停下。"""
    exclude = {pzn8(p) for p in exclude}
    unknown = exclude - {l["pzn"] for l in lines}
    if unknown:
        raise StopError(f"--exclude 里的 PZN 不在 PDF 上：{sorted(unknown)}")
    return ([l for l in lines if l["pzn"] not in exclude],
            [l for l in lines if l["pzn"] in exclude])


def not_base(hit):
    """非基础品的原因（None = 基础品）。hit = (外部ID, 品名, sale_ok, active, 内部编码, 是否套件)。"""
    code = (hit[4] or "").strip()
    if hit[5]:
        return "套件（有组件的 phantom BoM）"
    if _po_base_sku(code) != code:
        return f"渠道 / 组合装副本（编码 {code}）"
    return None


def match(lines, candidates, pick):
    """candidates: pzn → [(外部ID, 品名, sale_ok, active, 内部编码, 是否套件)]。pick: pzn → 外部ID。

    先排除非基础品（`not_base`），剩恰好 1 个基础品才填进行里，被排除的记在 `l["excluded"]` 供核对表留痕；
    剩 0 个或多个收进 problems（附全部命中）原样返回，不猜、也不退回去用套件 / 副本。
    """
    pick = {pzn8(k): v for k, v in pick.items()}
    unknown = set(pick) - {l["pzn"] for l in lines}
    if unknown:
        raise StopError(f"--pick 里的 PZN 不在（剔除后的）PDF 明细上：{sorted(unknown)}")
    problems = []
    for l in lines:
        hits = candidates.get(l["pzn"], [])
        base = [h for h in hits if not not_base(h)]
        l["excluded"] = [(h, not_base(h)) for h in hits if not_base(h)]
        if l["pzn"] in pick:
            if pick[l["pzn"]] in {h[0] for h, _ in l["excluded"]}:
                raise StopError(f"--pick {l['pzn']}={pick[l['pzn']]}：这是非基础品（套件 / 副本），SO 只用基础品。")
            chosen = [h for h in base if h[0] == pick[l["pzn"]]]
            if not chosen:
                raise StopError(f"--pick {l['pzn']}={pick[l['pzn']]}：这个外部 ID 不在该 PZN 的基础品命中里 "
                                f"{[h[0] for h in base]}")
            base = chosen
        if len(base) == 1:
            l["xmlid"], l["erp_name"], l["sale_ok"], l["active"] = base[0][:4]
        else:
            problems.append((l, hits))
    return problems


def odoo_candidates(pzns):
    """Odoo 只读：pzn → [(__export__ 外部 ID, 品名, sale_ok, active, 内部编码, 是否套件)]，含已归档品。

    套件 = 产品（或其模板上不限变体）挂着生效中的、至少一个组件的 phantom BoM。
    """
    from odoo_api.odoo_client import Odoo
    od = Odoo.connect(context={"active_test": False}, verbose=False)
    prods = od.execute("product.product", "search_read", [[["pzn", "in", sorted(pzns)]]],
                       {"fields": ["id", "name", "pzn", "sale_ok", "active", "default_code", "product_tmpl_id"]})
    boms = od.execute("mrp.bom", "search_read",
                      [[["active", "=", True], ["type", "=", "phantom"], ["bom_line_ids", "!=", False],
                        ["product_tmpl_id", "in", sorted({p["product_tmpl_id"][0] for p in prods})]]],
                      {"fields": ["product_tmpl_id", "product_id"]}) if prods else []
    kit_tmpl = {b["product_tmpl_id"][0] for b in boms if not b["product_id"]}
    kit_var = {b["product_id"][0] for b in boms if b["product_id"]}
    xids = od.execute("ir.model.data", "search_read",
                      [[["model", "=", "product.product"], ["module", "=", "__export__"],
                        ["res_id", "in", [p["id"] for p in prods]]]],
                      {"fields": ["res_id", "name"]}) if prods else []
    by_res = {}
    for x in xids:
        by_res.setdefault(x["res_id"], []).append(f"__export__.{x['name']}")
    out = {}
    for p in prods:
        ids = by_res.get(p["id"], [])
        if len(ids) != 1:
            # 没有或有多个导出 ID：模板要的是外部 ID，这种产品没法安全写进表
            raise StopError(f"产品 {p['id']}（PZN {p['pzn']}，{p['name'].strip()}）的 __export__ 外部 ID 有 "
                            f"{len(ids)} 个，没法确定写哪个。先在 Odoo 里导出一次该产品生成外部 ID。")
        kit = p["id"] in kit_var or p["product_tmpl_id"][0] in kit_tmpl
        out.setdefault(pzn8(p["pzn"]), []).append((ids[0], (p["name"] or "").strip(), p["sale_ok"], p["active"],
                                                   p["default_code"] or "", kit))
    return out


def tax_problem(l, taxes, fp_src, fp_name):
    """taxes: 该产品在 SO 公司下的销项税 [dict(id,name,amount,amount_type,price_include)]。

    能确定税率就写进 l["rate"]（Decimal 百分数）返回 None；否则返回原因（不猜、不默认 19%）。
    """
    if len(taxes) != 1:
        return f"产品挂了 {len(taxes)} 个销项税 {[t['name'] for t in taxes]}，不知按哪个算"
    t = taxes[0]
    if t["id"] in fp_src:
        return f"税「{t['name']}」会被 SO 财政位置「{fp_name}」替换，脚本不推算替换后的税"
    if t.get("amount_type", "percent") != "percent":
        return f"税「{t['name']}」不是百分比税（{t.get('amount_type')}）"
    if not t["price_include"]:
        return f"税「{t['name']}」是不含税口径，Unit Price 该填净价还是含税价待定（ISSUES [SO 导入] B）"
    l["rate"] = Decimal(str(t["amount"]))
    return None


def odoo_tax_info(xmlids, order_ref):
    """Odoo 只读：→ (财政位置名或 None, 财政位置映射的源税 id 集合, {外部ID: [该 SO 公司的销项税]})。"""
    from odoo_api.odoo_client import Odoo
    od = Odoo.connect(context={"active_test": False}, verbose=False)
    so = od.execute("sale.order", "search_read", [[["name", "=", order_ref]]],
                    {"fields": ["company_id", "fiscal_position_id"]})
    if len(so) != 1:
        raise StopError(f"Odoo 里按 Order Reference「{order_ref}」找到 {len(so)} 张 SO，没法确定税的公司与财政位置。")
    company, fp = so[0]["company_id"][0], so[0]["fiscal_position_id"]
    fp_src = set()
    if fp:
        for m in od.execute("account.fiscal.position.tax", "search_read", [[["position_id", "=", fp[0]]]],
                            {"fields": ["tax_src_id"]}):
            if m["tax_src_id"]:
                fp_src.add(m["tax_src_id"][0])
    names = {x.split(".", 1)[1]: x for x in xmlids}
    xd = od.execute("ir.model.data", "search_read",
                    [[["model", "=", "product.product"], ["module", "=", "__export__"], ["name", "in", list(names)]]],
                    {"fields": ["name", "res_id"]})
    by_res = {x["res_id"]: names[x["name"]] for x in xd}
    prods = od.execute("product.product", "read", [list(by_res), ["taxes_id"]])
    tids = sorted({t for p in prods for t in p["taxes_id"]})
    taxes = {t["id"]: t for t in od.execute(
        "account.tax", "read", [tids, ["name", "amount", "amount_type", "price_include", "company_id"]])} if tids else {}
    out = {by_res[p["id"]]: [taxes[t] for t in p["taxes_id"] if taxes[t]["company_id"][0] == company]
           for p in prods}
    return (fp[1] if fp else None), fp_src, out


def apply_taxes(lines, tax_info):
    """逐行定税率；有拿不准的行就整表停。"""
    fp_name, fp_src, by_xmlid = tax_info
    problems = []
    for l in lines:
        why = tax_problem(l, by_xmlid.get(l["xmlid"], []), fp_src, fp_name)
        if why:
            problems.append(f"  {l['pzn']} | {l['name']} | {l['erp_name']}：{why}")
    if problems:
        raise StopError("以下行的销项税拿不准，未出表（请在 Odoo 里核对产品税 / SO 财政位置后重跑）：\n"
                        + "\n".join(problems))


def read_template(template):
    """→ (workbook, 列名, 头信息)。缺列或第 2 行头信息不全就停。"""
    wb = load_workbook(template)
    ws = wb.active
    cols = [c.value for c in ws[1]]
    missing = [c for c in HEAD_COLS + (COL_PID, COL_IREF, COL_QTY, COL_PRICE) if c not in cols]
    if missing:
        raise StopError(f"模板缺列 {missing}；表头实际是 {cols}")
    head = {c: ws.cell(2, cols.index(c) + 1).value for c in HEAD_COLS}
    if not all(head.values()):
        raise StopError(f"模板第 2 行的 ID / Order Reference / Customer 不全：{head}")
    return wb, cols, head


def build(wb, cols, head, lines, out_path):
    """按模板列名列序写导入表：第一行带头信息，其余行只填明细列。"""
    ws = wb.active
    ws.delete_rows(2, ws.max_row)
    for k, l in enumerate(lines):
        l["price"] = l["netto"] * (1 + l["rate"] / 100)
        vals = {COL_PID: l["xmlid"], COL_IREF: None, COL_QTY: l["qty"], COL_PRICE: float(l["price"])}
        if k == 0:
            vals.update(head)
        ws.append([vals.get(c) for c in cols])
    wb.save(out_path)


def report(lines, dropped, summe):
    print("\n核对表（PZN / PDF 名称 / 数量 / 净价 / 税率 / 含税价 / ERP 品名 / 可销售）")
    for k, l in enumerate(lines, 1):
        flag = "" if l["active"] else "  ⚠ 已归档"
        rate = f"{l['rate'].normalize()}%" + ("" if l["rate"] == STD_RATE else " ⚠非19%")
        print(f"{k:>3} | {l['pzn']} | {l['name']} | {l['qty']} | {l['netto']} | {rate} | {l['price'].normalize()} | "
              f"{l['erp_name']} | {'✓' if l['sale_ok'] else '✗'}{flag}")
    auto = [l for l in lines if l.get("excluded")]
    if auto:
        print("\n自动排除的非基础品（SO 只用基础品）：")
        for l in auto:
            for h, why in l["excluded"]:
                print(f"    {l['pzn']} | {h[0]}  {h[1]} | {why}")
    if dropped:
        print("\n已剔除：")
        for l in dropped:
            print(f"    {l['pzn']} | {l['name']} | {l['qty']} | Nettosumme {l['nsum']}")
    full = sum(l["netto"] * l["qty"] for l in lines)
    dsum = sum(l["nsum"] for l in dropped)
    target = summe - dsum
    print(f"\n合计核对：出表行 三位净价×数量 全精度 {full}；PDF Summe {summe}"
          + (f" − 剔除行 {dsum} = {target}" if dropped else "") + f"；差 {target - full}")
    for l in lines:
        diff = l["nsum"] - l["netto"] * l["qty"]
        if diff:
            print(f"    行级四舍五入 {l['pzn']}：{l['netto'] * l['qty']} 印成 {l['nsum']}（{diff:+}）")
    no_sale = [l for l in lines if not l["sale_ok"]]
    if no_sale:
        # 只提示不拦：2026-10-13 实测按外部 ID 导入 SO，8 个 sale_ok=False 的 Temp_ 品照常导入成功
        print(f"\nℹ {len(no_sale)} 个产品未勾可销售（sale_ok=False），不影响按外部 ID 导入（2026-10-13 实测）；"
              f"在 SO 界面手动选产品时搜不到它们：{', '.join(l['pzn'] for l in no_sale)}")


def run(pdf, template, outdir=None, exclude=(), pick=None, candidates_fn=odoo_candidates,
        tax_fn=odoo_tax_info):
    wb, cols, head = read_template(template)
    lines, summe = parse_text(read_pdf(pdf))
    print(f"PDF 明细 {len(lines)} 行（页 {sorted({l['page'] for l in lines})}），Summe {summe}")
    lines, dropped = apply_exclude(lines, exclude)
    problems = match(lines, candidates_fn({l["pzn"] for l in lines}), pick or {})
    if problems:
        msg = ["以下行匹配异常，未出表（剔除用 --exclude PZN，基础品多选一用 --pick PZN=外部ID）："]
        for l, hits in problems:
            nb = sum(1 for h in hits if not not_base(h))
            msg.append(f"  {l['pzn']} | {l['name']} | 数量 {l['qty']} | 命中 {len(hits)} 个，其中基础品 {nb} 个")
            for h in hits:
                why = not_base(h)
                msg.append(f"      {h[0]}  {h[1]}  [{h[4]}]  sale_ok={h[2]}{'' if h[3] else '  已归档'}"
                           + (f"  ✗ 已排除：{why}" if why else ""))
        raise StopError("\n".join(msg))
    apply_taxes(lines, tax_fn([l["xmlid"] for l in lines], head["Order Reference"]))
    outdir = outdir or os.path.join("output", f"{dt.date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(pdf))[0]
    out = unique_path(os.path.join(outdir, f"{head['Order Reference']}_import_{stem}.xlsx"))
    build(wb, cols, head, lines, out)
    print(f"写出 {out}（{len(lines)} 行）；头信息 {head}")
    report(lines, dropped, summe)
    return out


def main():
    ap = argparse.ArgumentParser(description="客户 Bestellung PDF → Odoo 销售单导入表")
    ap.add_argument("pdf", help="客户 Bestellung PDF")
    ap.add_argument("template", help="Odoo 导出的 sale.order 模板 xlsx（第 2 行带 ID / Order Reference / Customer）")
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
