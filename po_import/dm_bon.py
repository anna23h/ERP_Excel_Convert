"""版式 D：dm 门店收银小票（扫描 PDF，无文字层、无 PZN）→ PO 明细行。见 ISSUES [PO 导入] H。

入口由 `po_from_pdf.run` 在 PDF 抽不出任何文字时调 `load`，返回的明细行结构与 A/B/C 相同，写表照旧。

用户定（2026-11-13）：
    · 文字：本地 OCR（pypdfium2 渲染 + rapidocr_onnxruntime），渲染倍数 4 → 3 → 5 依次试，取第一个护栏全过的。
    · 品名 → 产品：用户维护的对照表 CSV（`小票品名,外部ID,备注`，默认在私有库 `dm_bon_map.csv`），
      外部 ID 写 `SKIP` 表示这次不进 PO。OCR 品名**只在对照表内容错**：归一化后相似度唯一最高、≥ MIN_SCORE、
      与次高差 ≥ MIN_GAP 才认，否则停。不碰 Odoo 品名。
    · 税率取 Odoo 产品默认进项税，按档求和须等于小票印的两档 Brutto；OCR 读到的税码只做参考，矛盾就停。
    · 单价 = 含税单价 ÷ (1 + 税率)，2 位 HALF_UP；每档 Σ 净额与印的 Netto 差 ≤ 0.005 × 件数才放行。
    · Vendor Reference = `dm-YYYYMMDD-<TSE TA-Nummer>`。

OCR 实测（2026-11-13 那张）：数字三种倍数全对，品名每次不同（Zirkulin → Tirkulin / lirkulin），税码 `2` 常漏读。
"""
import csv
import difflib
import os
import re
import unicodedata
from decimal import ROUND_HALF_UP, Decimal

from po_import.po_from_pdf import TAX_FULL, TAX_REDUCED, StopError, is_variant

SCALES = (4, 3, 5)
MIN_SCORE = 0.8
MIN_GAP = 0.1
SKIP = "SKIP"
MAP_NAME = "dm_bon_map.csv"
MAP_COLS = ("小票品名", "外部ID", "备注")
RATES = {19: (Decimal("1.19"), TAX_FULL[1]), 7: (Decimal("1.07"), TAX_REDUCED[1])}

#: 金额：OCR 常把 `24,75` 读成 `24, 75`、`7,90` 读成 `7.90`
_M = r"\d{1,4}[,.] ?\d{2}"
SHOP_RE = re.compile(r"dm-?drogerie", re.I)
HEAD_RE = re.compile(r"^(\d{2})\.(\d{2})\.(\d{4}) ?\d{2}:\d{2}")
ITEM_RE = re.compile(rf"^(?:(\d{{1,3}}) ?[xX×] ?({_M}) ?)?(.*?[A-Za-zÄÖÜäöüß].*?) (-?{_M})(?: (\S{{1,2}}))?$")
QTY_ONLY_RE = re.compile(rf"^(\d{{1,3}}) ?[xX×] ?({_M})$")
SUM_RE = re.compile(rf"^SUMME ?EUR (-?{_M})$", re.I)
BAND_RE = re.compile(rf"^([12]) ?[=\-] ?(19|7)[,.]00 ?% ({_M}) ({_M}) ({_M})$")
START_RE = re.compile(r"Start:? ?(\d{4})-(\d{2})-(\d{2})")
TA_RE = re.compile(r"TA-Nummer:? ?(\S+)")


def money(s):
    s = s.replace(" ", "")
    return Decimal(s[:-3].replace(".", "").replace(",", "") + "." + s[-2:])


# ---------------------------------------------------------------- OCR

_ENGINE = None


def ocr_lines(path, scale):
    """PDF 各页渲染后 OCR → 文本行列表（同一行的文字块按 x 排序、空格拼接）。"""
    global _ENGINE
    try:
        import numpy as np
        import pypdfium2 as pdfium
        from rapidocr_onnxruntime import RapidOCR
    except ImportError as e:
        raise StopError(f"这张 PDF 没有文字层，要 OCR，但缺包（{e.name}）：pip install -r requirements-pdf.txt")
    if _ENGINE is None:
        _ENGINE = RapidOCR()
    out = []
    for page in pdfium.PdfDocument(path):
        img = page.render(scale=scale).to_pil().convert("RGB")
        res, _ = _ENGINE(np.array(img), box_thresh=0.3, unclip_ratio=1.8, text_score=0.3)
        boxes = sorted(((b[0][1] + b[2][1]) / 2, b[0][0], t.strip()) for b, t, _c in (res or []))
        rows = []
        for y, x, t in boxes:
            if rows and abs(rows[-1][0] - y) < 8 * scale:
                rows[-1][1].append((x, t))
            else:
                rows.append([y, [(x, t)]])
        out += [" ".join(t for _, t in sorted(parts)) for _, parts in rows]
    return out


# ---------------------------------------------------------------- 解析

def parse_dm(texts):
    """小票文本行 → dict(lines, summe, bands{税率: (brutto, netto, mwst)}, doc_no)。数字护栏任一不符就停。"""
    texts = [" ".join(t.split()) for t in texts]
    if not any(SHOP_RE.search(t) for t in texts):
        raise StopError("OCR 文本里没有 dm-drogerie markt 字样——不是 dm 小票，版式不认识。")
    head = next(((i, m) for i, t in enumerate(texts) if (m := HEAD_RE.match(t))), None)
    end = next((i for i, t in enumerate(texts) if SUM_RE.match(t)), None)
    if head is None or end is None or end <= head[0]:
        raise StopError("找不到小票头（日期 时间）或 SUMME EUR 行——版式不认识。")
    summe = money(SUM_RE.match(texts[end])[1])
    if summe <= 0:
        raise StopError(f"SUMME {summe} 不是正数——退货 / 冲销小票不做 PO 导入。")
    lines, odd, carry = [], [], None
    for t in texts[head[0] + 1:end]:
        m = QTY_ONLY_RE.match(t)
        if m:
            carry = m
            continue
        m = ITEM_RE.match(t)
        if not m:
            odd.append(t)
            continue
        q = m if m[1] else carry
        carry = None
        total = money(m[4])
        qty, unit = (int(q[1]), money(q[2])) if q else (1, total)
        code = m[5] if m[5] in ("1", "2") else None
        lines.append(dict(ocr=m[3].strip(), qty=qty, unit=unit, brutto=total, code=code, raw=t))
    if odd:
        raise StopError("明细区有认不出的行（不是「[Nx 单价] 品名 小计」），不出表：\n"
                        + "\n".join(f"  {t}" for t in odd))
    if not lines:
        raise StopError("小票里没解析出任何明细行——版式不认识。")
    neg = [l for l in lines if l["brutto"] <= 0 or l["unit"] <= 0]
    if neg:
        raise StopError("明细里有负数 / 0 行（折扣、退货、押金之类），脚本不会处理，不出表：\n"
                        + "\n".join(f"  {l['raw']}" for l in neg))
    bad = [l for l in lines if l["unit"] * l["qty"] != l["brutto"]]
    if bad:
        raise StopError("以下行 数量 × 单价 ≠ 小计（OCR 读错数字或漏读数量），不出表：\n"
                        + "\n".join(f"  {l['raw']}" for l in bad))
    got = sum(l["brutto"] for l in lines)
    if got != summe:
        raise StopError(f"明细小计合计 {got} ≠ SUMME {summe}——有行没读到或读错，不出表。")
    bands, codes = {}, {}
    for t in texts[end:]:
        m = BAND_RE.match(t)
        if m:
            rate = int(m[2])
            codes[m[1]] = rate
            bands[rate] = (money(m[3]), money(m[4]), money(m[5]))
    if not bands:
        raise StopError("找不到 MwSt-Satz 两档（Brutto / Netto / MwSt）行——版式不认识。")
    for rate, (b, n, w) in bands.items():
        if b - n != w or (b / RATES[rate][0]).quantize(Decimal("0.01"), ROUND_HALF_UP) != n:
            raise StopError(f"{rate}% 档 Brutto {b} / Netto {n} / MwSt {w} 自相矛盾——OCR 读错了，不出表。")
    if sum(b for b, _, _ in bands.values()) != summe:
        raise StopError(f"两档 Brutto 合计 {sum(b for b, _, _ in bands.values())} ≠ SUMME {summe}，不出表。")
    for l in lines:
        l["code_rate"] = codes.get(l["code"])
    d, mo, y = head[1].groups()
    start = next((m for t in texts if (m := START_RE.search(t))), None)
    if not start or (start[1], start[2], start[3]) != (y, mo, d):
        raise StopError(f"TSE Start 日期 {start and '-'.join(start.groups())} 与小票头日期 {y}-{mo}-{d} 对不上"
                        "（或没读到）——不敢确定单号，不出表。")
    ta = next((m[1] for t in texts if (m := TA_RE.search(t))), None)
    if not ta or not ta.isdigit():
        raise StopError(f"TSE TA-Nummer 没读到或不是纯数字（{ta}）——单号 / 防重靠它，不出表。")
    return dict(lines=lines, summe=summe, bands=bands, doc_no=f"dm-{y}{mo}{d}-{ta}")


def read_bon(path, ocr_fn=ocr_lines):
    """OCR + 解析，倍数依次试；全失败就报第一个倍数的原因并附 OCR 文本。"""
    errors = []
    for sc in SCALES:
        texts = ocr_fn(path, sc)
        try:
            bon = parse_dm(texts)
            bon["scale"] = sc
            return bon
        except StopError as e:
            errors.append((sc, e, texts))
    sc, e, texts = errors[0]
    raise StopError(f"{e}\n（渲染倍数 {', '.join(str(s) for s, _, _ in errors)} 都没过护栏；以下是 {sc} 倍的 OCR 文本）\n"
                    + "\n".join(f"  | {t}" for t in texts))


# ---------------------------------------------------------------- 对照表

def norm(s):
    s = unicodedata.normalize("NFKD", s.lower().replace("ß", "ss"))
    return re.sub(r"[^0-9a-z]", "", s)


def default_map_path():
    """私有库 dm_bon_map.csv：候选同开工前拉取脚本（$ERP_PRIVATE_DIR / docs 的上一级 / 原始开发文档 / ../erp-private）。"""
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    cands = [os.environ.get("ERP_PRIVATE_DIR"),
             os.path.join(os.path.realpath(os.path.join(root, "docs")), os.pardir),
             os.path.join(root, "原始开发文档"),
             os.path.join(root, os.pardir, "erp-private")]
    for c in cands:
        if c and os.path.isfile(os.path.join(c, "ISSUES.md")):
            return os.path.normpath(os.path.join(c, MAP_NAME))
    raise StopError("找不到私有库（erp-private），没法定位 dm 对照表；用 --dm-map 指定路径。")


def load_map(path):
    """→ [(小票品名, 外部ID, 备注)]。文件不存在 / 缺列 / 同名重复就停。"""
    if not os.path.isfile(path):
        raise StopError(f"dm 对照表不存在：{path}\n按表头 {','.join(MAP_COLS)} 新建（UTF-8），每行一种小票品名。")
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if rows and not set(MAP_COLS[:2]) <= set(rows[0]):
        raise StopError(f"dm 对照表 {path} 表头应是 {','.join(MAP_COLS)}，实际 {list(rows[0])}")
    out, seen = [], {}
    for r in rows:
        name, xid = (r.get(MAP_COLS[0]) or "").strip(), (r.get(MAP_COLS[1]) or "").strip()
        if not name:
            continue
        if not xid:
            raise StopError(f"dm 对照表里「{name}」没填外部ID（不进 PO 写 {SKIP}）。")
        if norm(name) in seen:
            raise StopError(f"dm 对照表里「{name}」与「{seen[norm(name)]}」归一化后相同，删掉一条。")
        seen[norm(name)] = name
        out.append((name, xid, (r.get(MAP_COLS[2]) or "").strip()))
    return out


def score(a, b):
    return difflib.SequenceMatcher(None, norm(a), norm(b)).ratio()


def match_names(lines, mapping):
    """每行 OCR 品名在对照表里找：唯一最高、≥ MIN_SCORE、领先次高 ≥ MIN_GAP 才认；否则收进 problems。"""
    problems = []
    for l in lines:
        ranked = sorted(((score(l["ocr"], n), n, x) for n, x, _ in mapping), reverse=True)
        best = ranked[0] if ranked else (0, None, None)
        second = ranked[1][0] if len(ranked) > 1 else 0
        if best[0] >= MIN_SCORE and best[0] - second >= MIN_GAP:
            l["score"], l["bon_name"], l["xmlid"] = best
        else:
            problems.append((l, ranked[:2]))
    return problems


# ---------------------------------------------------------------- Odoo

def odoo_products(xmlids):
    """Odoo 只读：外部 ID → (品名, Internal Reference, active, 默认进项税名称元组, 数据库 ID, PZN)。查不到的不在结果里。"""
    from po_import.po_from_pdf import _odoo
    od = _odoo()
    pairs = [x.split(".", 1) for x in xmlids if "." in x]
    recs = od.execute("ir.model.data", "search_read",
                      [[["model", "=", "product.product"], ["module", "in", sorted({m for m, _ in pairs})],
                        ["name", "in", [n for _, n in pairs]]]], {"fields": ["module", "name", "res_id"]}) if pairs else []
    by_res = {r["res_id"]: f"{r['module']}.{r['name']}" for r in recs}
    prods = od.execute("product.product", "read", [sorted(by_res)],
                       {"fields": ["name", "default_code", "active", "supplier_taxes_id", "pzn"]}) if by_res else []
    tax_ids = sorted({t for p in prods for t in p["supplier_taxes_id"]})
    tax_names = {t["id"]: t["name"] for t in od.execute(
        "account.tax", "read", [tax_ids], {"fields": ["name"]})} if tax_ids else {}
    return {by_res[p["id"]]: ((p["name"] or "").strip(), p["default_code"] or "", p["active"],
                              tuple(sorted(tax_names[t] for t in p["supplier_taxes_id"])), p["id"], p["pzn"] or "")
            for p in prods}


# ---------------------------------------------------------------- 组装

def load(path, map_path=None, ocr_fn=None, products_fn=None):
    """→ (明细行, 剔除行, 单号, bon)。明细行带 PO 写表要的 xmlid/dbid/qty/aep/reduced 及核对表字段。"""
    products_fn = products_fn or odoo_products
    bon = read_bon(path, ocr_fn or ocr_lines)
    map_path = map_path or default_map_path()
    problems = match_names(bon["lines"], load_map(map_path))
    if problems:
        msg = [f"以下小票行在 dm 对照表里没有唯一匹配（相似度 ≥ {MIN_SCORE} 且领先次高 ≥ {MIN_GAP}），未出表。",
               f"请在 {map_path} 补一行「小票品名,外部ID,备注」（品名照小票原样写；不进 PO 写 {SKIP}）后重跑："]
        for l, near in problems:
            msg.append(f"  OCR「{l['ocr']}」 数量 {l['qty']} 小计 {l['brutto']}"
                       + ("".join(f"\n      最接近：{n} → {x}（{s:.2f}）" for s, n, x in near) or "\n      对照表是空的"))
        raise StopError("\n".join(msg))
    kept = [l for l in bon["lines"] if l["xmlid"] != SKIP]
    dropped = [l for l in bon["lines"] if l["xmlid"] == SKIP]
    if not kept:
        raise StopError("对照表把所有行都标成 SKIP，没有可进 PO 的行。")
    prods = products_fn(sorted({l["xmlid"] for l in kept}))
    bad = []
    for l in kept:
        p = prods.get(l["xmlid"])
        if p is None:
            bad.append(f"  {l['bon_name']} → {l['xmlid']}：Odoo 里查不到这个外部 ID")
            continue
        l["erp_name"], code, l["active"], l["erp_taxes"], l["dbid"], l["pzn"] = p
        rate = {RATES[r][1]: r for r in RATES}.get(l["erp_taxes"][0]) if len(l["erp_taxes"]) == 1 else None
        if is_variant(code):
            bad.append(f"  {l['bon_name']} → {l['xmlid']} [{code}]：是渠道 / 组合装副本，采购只用基础品，改对照表")
        elif rate is None:
            bad.append(f"  {l['bon_name']} → {l['xmlid']}：产品默认进项税是 {l['erp_taxes'] or '无'}，"
                       "不是恰好 7% / 19% Vorsteuer 之一，没法定税率")
        elif l["code_rate"] and l["code_rate"] != rate:
            bad.append(f"  {l['bon_name']} → {l['xmlid']}：小票税码是 {l['code_rate']}%，产品默认进项税是 {rate}%"
                       "——对照表可能配错了产品")
        l["rate"] = rate
    if bad:
        raise StopError("对照表匹配到的产品有问题，未出表：\n" + "\n".join(bad))
    check_bands(kept, dropped, bon["bands"])
    for l in kept:
        l["aep"] = (l["unit"] / RATES[l["rate"]][0]).quantize(Decimal("0.01"), ROUND_HALF_UP)
        l["gesamt"] = l["aep"] * l["qty"]
        l["reduced"] = l["rate"] == 7
        l["name"] = l["bon_name"]
        l["variants"] = []
    bon["netto_diff"] = check_netto(kept, bon["bands"])
    return kept, dropped, bon["doc_no"], bon


def check_bands(kept, dropped, bands):
    """按产品税率分档的 Brutto 合计必须等于小票印的；有 SKIP 行时，各档差额须 ≥ 0 且合计 = SKIP 行小计合计。"""
    got = {r: sum((l["brutto"] for l in kept if l["rate"] == r), Decimal(0)) for r in RATES}
    left = {r: bands.get(r, (Decimal(0),))[0] - got[r] for r in RATES}
    skipped = sum((l["brutto"] for l in dropped), Decimal(0))
    if any(v < 0 for v in left.values()) or sum(left.values()) != skipped:
        raise StopError("按 Odoo 产品税率分档的 Brutto 合计与小票印的对不上（多半是对照表配错了产品）：\n" + "\n".join(
            f"  {r}%：出表 {got[r]} vs 小票 {bands.get(r, (Decimal(0),))[0]}" for r in RATES)
            + (f"\n  SKIP 行合计 {skipped}" if dropped else ""))


def check_netto(kept, bands):
    """每档 Σ(净价 × 数量) 与印的 Netto 比；有 SKIP 行落在该档（Brutto 不等）就不比。差额超逐件舍入上限就停。"""
    diff = {}
    for r, (b, n, _w) in bands.items():
        rows = [l for l in kept if l["rate"] == r]
        if sum((l["brutto"] for l in rows), Decimal(0)) != b:
            continue
        d = sum((l["gesamt"] for l in rows), Decimal(0)) - n
        if abs(d) > Decimal("0.005") * sum(l["qty"] for l in rows):
            raise StopError(f"{r}% 档 Σ 净价 × 数量 与小票 Netto {n} 差 {d}，超过逐件舍入上限，不出表。")
        diff[r] = d
    return diff


def report(lines, dropped, bon):
    print(f"\n核对表（OCR 渲染 {bon['scale']} 倍；PZN / OCR 原文 / 小票名(相似度) / 数量 / 含税单价 / 净价 / 税 / ERP 品名）")
    for k, l in enumerate(lines, 1):
        flag = "" if l["active"] else "  ⚠ 已归档"
        print(f"{k:>3} | {l['pzn']} | {l['ocr']} | {l['bon_name']}({l['score']:.2f}) | {l['qty']} | {l['unit']} | "
              f"{l['aep']} | {l['rate']}% | {l['erp_name']}{flag}")
    if dropped:
        print("\n已剔除（对照表标 SKIP）：")
        for l in dropped:
            print(f"    {l['ocr']} → {l['bon_name']} | {l['qty']} | 小计 {l['brutto']}")
    print(f"\n合计核对：出表 {len(lines)} 行 {sum(l['qty'] for l in lines)} 件；小票 SUMME {bon['summe']}")
    for r, (b, n, _w) in sorted(bon["bands"].items(), reverse=True):
        rows = [l for l in lines if l["rate"] == r]
        net = sum((l["gesamt"] for l in rows), Decimal(0))
        tail = (f"，舍入差 {bon['netto_diff'][r]:+}" if r in bon["netto_diff"] else "（含 SKIP 行，不比净额）")
        print(f"    {r}%：Brutto 出表 {sum((l['brutto'] for l in rows), Decimal(0))} / 小票 {b}；"
              f"Netto 出表 {net} / 小票 {n}{tail}")
