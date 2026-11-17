"""草稿客户发票按价格表批量改单价：出 Odoo 导入表 + 核对表；导入后比对改前 / 改后导出。

**纯文件进出，不连 Odoo、不读 config.py**（2026-11-17 用户定，ISSUES [发票改价] A）：
账号的 API key 有写权限，本地只读白名单拦不住临时脚本，所以数据一律由人从 Odoo 界面导出。

用法：
    # 1) 出导入表 + 核对表
    python3 -m procure.invoice_reprice build <发票导出.xlsx> <价格表导出.xlsx> [-o 输出目录]
    # 2) 在 Odoo 发票列表里导入「导入表」，再用同一个导出模板导出一次，然后比对
    python3 -m procure.invoice_reprice verify <改前导出.xlsx> <改后导出.xlsx> [--expect 导入表.xlsx]

两份输入都用 Odoo「普通导出」（不勾导入兼容），字段见 README 对应章节：
  发票：Status / ID(外部 ID) / Invoice/Bill Date / Date / Partner / Untaxed Amount / Total /
        Invoice lines/ID(外部 ID) / Display Type / Quantity / Unit Price / Discount (%) /
        Product/Internal Reference / Subtotal / Total / Label（另带 Taxes、Account 供比对）
  价格表：Pricelist Name / Pricelist Items/ Product/Internal Reference / Product/Name /
        Fixed Price / Min. Quantity / Start Date / End Date / Compute Price

固定口径（用户定）：
  1. 按**完整产品编码**精确匹配（`_GW` 与非 GW 是不同产品，分开配）。
  2. 取价日 = 发票日期；草稿没有发票日期时用记账日期 `Date`（核对表标出）。
     有效 = Start ≤ 取价日 ≤ End（空则不限）。多条有效取 Start 最新一条并标 ⚠ 列出候选；
     没有有效价 / 配不上 → 保持原价并标出。Min. Quantity 只显示，不参与取价。
  3. 含税口径不写死：逐行用 Subtotal / Total 与 单价×数量×(1−折扣) 反推。价格表价直接填
     price_unit（与 Odoo 在 SO 上套价格表一致）；税率不为 0 的行标 ⚠ 请人确认口径。
  4. SO 单价不动；只改价格有变化的行；备注 / 分节行不进导入表。

导入表格式（2026-11-17 在 47306 上实测导入成功）：
    id | invoice_line_ids/id | invoice_line_ids/price_unit
发票外部 ID 只写在第一行。走 account.move.write 的发票自动补全，小计 / 应收行 / 合计由 Odoo 重算。
"""
import argparse
import datetime as dt
import os
import re
import shutil
import sys

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

from openpyxl import Workbook, load_workbook

from common.xlsx import style_sheet, unique_path

TOL = 0.01                    # 金额比较容差（行级四舍五入到分）
DRAFT = {"draft", "草稿", "entwurf"}
SKIP_DISPLAY = {"note", "section", "line_note", "line_section", "备注", "部分", "notiz", "abschnitt"}
XMLID_RE = re.compile(r"^[A-Za-z_][\w-]*\.[\w.-]+$")
PZN_RE = re.compile(r"_(\d{7,8})(?:_[A-Za-z]+)?$")


class InputError(RuntimeError):
    """输入文件不符合约定——停下让人看，不猜。"""


def say(msg=""):
    print(msg, flush=True)


# --------------------------------------------------------------------------
# 读表：按表头定位，再用列内容区分同名列（导出模板把外部 ID 也显示成「ID」）
# --------------------------------------------------------------------------
def read_rows(path):
    ws = load_workbook(path, data_only=True).active
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    if not rows:
        raise InputError(f"{path}：空表")
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    body = [r + [None] * (len(header) - len(r)) for r in rows[1:]
            if any(v not in (None, "") for v in r)]
    return header, body


def is_xmlid(v):
    return isinstance(v, str) and bool(XMLID_RE.match(v.strip())) and not v.strip().isdigit()


def find_col(header, body, names, want_xmlid=None, required=True, what=""):
    """names 里任一表头名命中的列。want_xmlid=True 只要装外部 ID 的列，False 只要不是外部 ID 的列。"""
    cands = [i for i, h in enumerate(header) if h in names]
    if want_xmlid is not None:
        def kind(i):
            vals = [r[i] for r in body if r[i] not in (None, "", False)]
            return bool(vals) and all(is_xmlid(v) for v in vals)
        cands = [i for i in cands if kind(i) == want_xmlid]
    if not cands:
        if required:
            raise InputError(f"导出里找不到「{what or names[0]}」列（认的表头：{' / '.join(names)}）。"
                             "请用固定导出模板重新导出。")
        return None
    return cands[0]


def num(v):
    if v in (None, "", False):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return float(str(v).replace(",", "."))


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if v in (None, "", False):
        return None
    s = str(v).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y", "%d.%m.%Y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise InputError(f"认不出的日期：{v!r}")


def text(v):
    return "" if v in (None, False) else str(v).strip()


def pzn_of(code):
    m = PZN_RE.search(code or "")
    return m.group(1) if m else ""


# --------------------------------------------------------------------------
# 发票导出
# --------------------------------------------------------------------------
def parse_invoice(path):
    header, body = read_rows(path)
    c = {
        "state": find_col(header, body, ["Status", "状态"], what="Status"),
        "move_id": find_col(header, body, ["ID", "External ID", "外部 ID"], want_xmlid=True,
                            what="发票外部 ID（ID / External ID）"),
        "inv_date": find_col(header, body, ["Invoice/Bill Date", "Invoice Date"], what="Invoice/Bill Date"),
        "date": find_col(header, body, ["Date"], required=False),
        "partner": find_col(header, body, ["Partner", "Customer"], what="Partner"),
        "untaxed": find_col(header, body, ["Untaxed Amount"], what="Untaxed Amount"),
        "total": find_col(header, body, ["Total"], what="Total"),
        "line_id": find_col(header, body, ["Invoice lines/ID", "Invoice lines/External ID"],
                            want_xmlid=True, what="Invoice lines/ID（外部 ID）"),
        "display": find_col(header, body, ["Invoice lines/Display Type"], what="Invoice lines/Display Type"),
        "qty": find_col(header, body, ["Invoice lines/Quantity"], what="Invoice lines/Quantity"),
        "price": find_col(header, body, ["Invoice lines/Unit Price"], what="Invoice lines/Unit Price"),
        "disc": find_col(header, body, ["Invoice lines/Discount (%)"], what="Invoice lines/Discount (%)"),
        "code": find_col(header, body, ["Invoice lines/Product/Internal Reference"],
                         what="Invoice lines/Product/Internal Reference"),
        "sub": find_col(header, body, ["Invoice lines/Subtotal"], what="Invoice lines/Subtotal"),
        "ltotal": find_col(header, body, ["Invoice lines/Total"], what="Invoice lines/Total"),
        "label": find_col(header, body, ["Invoice lines/Label"], required=False),
    }
    move_ids = {text(r[c["move_id"]]) for r in body if text(r[c["move_id"]])}
    if len(move_ids) != 1:
        raise InputError(f"{path}：应当正好一张发票，实际 {len(move_ids)} 张。一次只处理一张。")
    first = next(r for r in body if text(r[c["move_id"]]))
    lines = []
    for r in body:
        lid = text(r[c["line_id"]])
        if not lid:                       # 多税码等子行：不是新的发票行
            continue
        lines.append({
            "line_id": lid,
            "display": text(r[c["display"]]),
            "code": text(r[c["code"]]),
            "label": text(r[c["label"]]) if c["label"] is not None else "",
            "qty": num(r[c["qty"]]),
            "price": num(r[c["price"]]),
            "disc": num(r[c["disc"]]),
            "sub": num(r[c["sub"]]),
            "total": num(r[c["ltotal"]]),
        })
    inv_date = as_date(first[c["inv_date"]])
    acc_date = as_date(first[c["date"]]) if c["date"] is not None else None
    return {
        "move_id": move_ids.pop(),
        "state": text(first[c["state"]]),
        "inv_date": inv_date,
        "acc_date": acc_date,
        "partner": text(first[c["partner"]]),
        "untaxed": num(first[c["untaxed"]]),
        "total": num(first[c["total"]]),
        "lines": lines,
    }


def is_product_line(line):
    return line["display"].lower() not in SKIP_DISPLAY


def price_date(inv, today=None):
    """→ (取价日, 说明)。草稿没有发票日期时退到记账日期。"""
    if inv["inv_date"]:
        return inv["inv_date"], "发票日期"
    if inv["acc_date"]:
        return inv["acc_date"], "发票日期为空，用记账日期 Date"
    return (today or dt.date.today()), "发票日期与记账日期都为空，用今天"


# --------------------------------------------------------------------------
# 价格表导出
# --------------------------------------------------------------------------
def _pl_names(suffix):
    return [f"Pricelist Items/{suffix}", suffix]


def parse_pricelist(path):
    header, body = read_rows(path)
    header = [re.sub(r"/\s+", "/", h) for h in header]       # 「Pricelist Items/ Product/…」容错
    c = {
        "name": find_col(header, body, ["Pricelist Name", "Pricelist"], required=False),
        "code": find_col(header, body, _pl_names("Product/Internal Reference"),
                         what="Pricelist Items/Product/Internal Reference"),
        "pname": find_col(header, body, _pl_names("Product/Name"), required=False),
        "price": find_col(header, body, _pl_names("Fixed Price"), what="Fixed Price"),
        "minq": find_col(header, body, _pl_names("Min. Quantity"), required=False),
        "start": find_col(header, body, _pl_names("Start Date"), what="Start Date"),
        "end": find_col(header, body, _pl_names("End Date"), what="End Date"),
        "compute": find_col(header, body, _pl_names("Compute Price"), required=False),
    }
    names = {text(r[c["name"]]) for r in body if text(r[c["name"]])} if c["name"] is not None else set()
    if len(names) > 1:
        raise InputError(f"{path}：含 {len(names)} 张价格表（{' / '.join(sorted(names))}），一次只给一张。")
    items, no_code = [], 0
    for r in body:
        code = text(r[c["code"]])
        if not code:
            no_code += 1
            continue
        compute = text(r[c["compute"]]) if c["compute"] is not None else "Fixed Price"
        items.append({
            "code": code,
            "pname": text(r[c["pname"]]) if c["pname"] is not None else "",
            "price": num(r[c["price"]]),
            "minq": num(r[c["minq"]]) if c["minq"] is not None else 0.0,
            "start": as_date(r[c["start"]]),
            "end": as_date(r[c["end"]]),
            "fixed": compute.lower() in ("fixed price", "fixed", "固定价格"),
        })
    return {"name": names.pop() if names else "", "items": items, "no_code": no_code}


def index_items(items):
    idx = {}
    for it in items:
        idx.setdefault(it["code"], []).append(it)
    return idx


# --------------------------------------------------------------------------
# 取价与含税口径（纯函数，单测覆盖）
# --------------------------------------------------------------------------
def pick_price(cands, day):
    """→ (价格 or None, 状态, 有效候选列表)。状态：匹配 / 多条有效 / 无有效价 / 非固定价。"""
    fixed = [it for it in cands if it["fixed"]]
    if not fixed:
        return None, "非固定价", []
    valid = [it for it in fixed
             if (it["start"] is None or it["start"] <= day) and (it["end"] is None or it["end"] >= day)]
    if not valid:
        return None, "无有效价", []
    valid.sort(key=lambda it: it["start"] or dt.date.min, reverse=True)
    return valid[0]["price"], ("多条有效" if len(valid) > 1 else "匹配"), valid


def tax_mode(line):
    """→ (口径, 税率)。口径：0税 / 含税 / 净价 / 无法判断。"""
    base = line["qty"] * line["price"] * (1 - line["disc"] / 100)
    sub, tot = line["sub"], line["total"]
    rate = (tot / sub - 1) if abs(sub) > TOL else 0.0
    if abs(tot - sub) <= TOL:
        return "0税", 0.0
    if abs(base - tot) <= TOL:
        return "含税", rate
    if abs(base - sub) <= TOL:
        return "净价", rate
    return "无法判断", rate


def partner_matches(partner, pricelist_name):
    p, n = partner.strip().lower(), pricelist_name.strip().lower()
    return bool(p) and bool(n) and p in n


def build_rows(inv, pl_index, day):
    rows = []
    for ln in inv["lines"]:
        if not is_product_line(ln):
            continue
        mode, rate = tax_mode(ln)
        notes, warn = [], False
        if not ln["code"]:
            new, status, valid = None, "无产品编码", []
        elif ln["code"] not in pl_index:
            new, status, valid = None, "配不上", []
        else:
            new, status, valid = pick_price(pl_index[ln["code"]], day)
        if status != "匹配":
            warn = True
        if status == "多条有效":
            notes.append("取 Start 最新一条")
        if mode == "无法判断":
            warn = True
            notes.append("Subtotal/Total 与 单价×数量 对不上，口径无法判断")
        elif mode != "0税":
            warn = True
            notes.append(f"税率 {rate:.0%}，请确认价格表价与本行{mode}口径一致")
        price_new = new if new is not None else ln["price"]
        changed = new is not None and abs(new - ln["price"]) > 1e-9
        rows.append({
            "line": ln, "pzn": pzn_of(ln["code"]), "new": price_new, "status": status,
            "changed": changed, "mode": mode, "rate": rate, "warn": warn,
            "cands": "; ".join(f"{it['price']} (起 {it['start'] or '-'} 止 {it['end'] or '-'}"
                               f"{'，最小量 %g' % it['minq'] if it['minq'] else ''})" for it in valid)
                     if len(valid) > 1 else "",
            "minq": valid[0]["minq"] if valid else None,
            "notes": "；".join(notes),
        })
    return rows


# --------------------------------------------------------------------------
# 输出
# --------------------------------------------------------------------------
def move_tag(move_xmlid):
    m = re.search(r"account_move_(\d+)", move_xmlid)
    return m.group(1) if m else re.sub(r"[^\w-]+", "_", move_xmlid)


def write_import(path, move_xmlid, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(["id", "invoice_line_ids/id", "invoice_line_ids/price_unit"])
    first = True
    for r in rows:
        if not r["changed"]:
            continue
        ws.append([move_xmlid if first else None, r["line"]["line_id"], r["new"]])
        first = False
    wb.save(path)


CHECK_COLS = ["PZN", "产品编码", "品名", "数量", "旧价", "新价", "价差", "新小计", "结果",
              "改价", "含税口径", "有效候选", "最小量", "备注"]


def write_check(path, inv, pl, day, day_note, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "核对表"
    ws.append([f"发票 {move_tag(inv['move_id'])} · {inv['partner']} · 价格表 {pl['name'] or '（未注明）'}"])
    ws.append([f"取价日 {day}（{day_note}）；状态 {inv['state']}；原合计 {inv['total']:.2f}"])
    ws.append(CHECK_COLS)
    for r in rows:
        ln = r["line"]
        base = ln["qty"] * r["new"] * (1 - ln["disc"] / 100)
        ws.append([r["pzn"], ln["code"], ln["label"], ln["qty"], ln["price"], r["new"],
                   round(r["new"] - ln["price"], 4), round(base, 2),
                   ("⚠ " if r["warn"] else "") + r["status"], "是" if r["changed"] else "",
                   r["mode"], r["cands"], r["minq"], r["notes"]])
    style_sheet(ws, len(CHECK_COLS), header_row=3, auto_align=True,
                widths={"品名": 50, "有效候选": 40, "备注": 40})
    ws.freeze_panes = "C4"
    wb.save(path)


def build(inv_path, pl_path, outdir=None):
    inv = parse_invoice(inv_path)
    if inv["state"].lower() not in DRAFT:
        raise InputError(f"发票状态是「{inv['state']}」，不是草稿。只有草稿能导入改价，"
                         "请先在 Odoo 里重置为草稿再导出。")
    pl = parse_pricelist(pl_path)
    day, day_note = price_date(inv)
    rows = build_rows(inv, index_items(pl["items"]), day)

    outdir = outdir or os.path.join("output", f"{dt.date.today():%Y%m%d}")
    os.makedirs(outdir, exist_ok=True)
    tag = move_tag(inv["move_id"])
    check = unique_path(os.path.join(outdir, f"发票改价_核对表_{tag}.xlsx"))
    write_check(check, inv, pl, day, day_note, rows)
    snap = unique_path(os.path.join(outdir, f"发票改价_改前快照_{tag}.xlsx"))
    shutil.copy2(inv_path, snap)
    n_change = sum(r["changed"] for r in rows)
    imp = None
    if n_change:
        imp = unique_path(os.path.join(outdir, f"发票改价_导入表_{tag}.xlsx"))
        write_import(imp, inv["move_id"], rows)

    say(f"发票 {tag} · {inv['partner']} · 状态 {inv['state']}")
    say(f"价格表 {pl['name'] or '（未注明）'}：{len(pl['items'])} 条"
        + (f"，另有 {pl['no_code']} 条没有产品编码已跳过" if pl["no_code"] else ""))
    say(f"取价日 {day}（{day_note}）")
    if pl["name"] and not partner_matches(inv["partner"], pl["name"]):
        say(f"⚠ 价格表名称里没有发票客户「{inv['partner']}」——确认是不是给错了价格表")
    say()
    say("| PZN | 产品编码 | 数量 | 旧价 | 新价 | 结果 | 含税口径 |")
    say("|---|---|---:|---:|---:|---|---|")
    for r in rows:
        ln = r["line"]
        say(f"| {r['pzn']} | {ln['code']} | {ln['qty']:g} | {ln['price']:g} | {r['new']:g} | "
            f"{'⚠ ' if r['warn'] else ''}{r['status']} | {r['mode']} |")
    new_total = sum(round(r["line"]["qty"] * r["new"] * (1 - r["line"]["disc"] / 100), 2) for r in rows)
    old_total = sum(round(r["line"]["qty"] * r["line"]["price"] * (1 - r["line"]["disc"] / 100), 2)
                    for r in rows)
    say()
    say(f"产品行 {len(rows)} 行，改价 {n_change} 行；按单价×数量估算 {old_total:.2f} → {new_total:.2f}"
        "（实际合计以 Odoo 重算为准）")
    say(f"核对表：{check}")
    say(f"改前快照：{snap}")
    say(f"导入表：{imp}" if imp else "没有需要改价的行，不出导入表。")
    return {"rows": rows, "check": check, "import": imp, "snapshot": snap}


# --------------------------------------------------------------------------
# verify：改前 / 改后两份导出逐字段比对
# --------------------------------------------------------------------------
#: 允许变化的列（其余任何列变了都报出来）
ALLOWED_LINE = {"Invoice lines/Unit Price", "Invoice lines/Subtotal", "Invoice lines/Total"}
ALLOWED_HEAD = {"Untaxed Amount", "Total", "Tax amount by group", "Display Name"}


def _records(header, body, line_col, move_col):
    """→ (发票头 {列序号: 值}, {行外部 ID: {列序号: 值}})。子行（多税码）拼到上一行。"""
    head, lines, cur = {}, {}, None
    line_cols = [i for i, h in enumerate(header) if h.startswith("Invoice lines/")]
    head_cols = [i for i, h in enumerate(header) if not h.startswith("Invoice lines/")]
    for k, r in enumerate(body):
        if k == 0:
            head = {i: r[i] for i in head_cols}
        lid = text(r[line_col])
        if lid:
            cur = lid
            lines[lid] = {i: r[i] for i in line_cols}
        elif cur:
            for i in line_cols:
                if r[i] not in (None, "", False):
                    lines[cur][i] = f"{lines[cur][i]} | {r[i]}"
    return head, lines


def _same(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) <= 1e-9
    return text(a) == text(b)


def read_expect(path):
    header, body = read_rows(path)
    lc = header.index("invoice_line_ids/id")
    pc = header.index("invoice_line_ids/price_unit")
    return {text(r[lc]): num(r[pc]) for r in body if text(r[lc])}


def verify(before_path, after_path, expect_path=None):
    hb, bb = read_rows(before_path)
    ha, ba = read_rows(after_path)
    if hb != ha:
        raise InputError("改前、改后两份导出的列不一样——必须用同一个导出模板。")
    line_col = find_col(hb, bb, ["Invoice lines/ID", "Invoice lines/External ID"], want_xmlid=True,
                        what="Invoice lines/ID（外部 ID）")
    move_col = find_col(hb, bb, ["ID", "External ID"], want_xmlid=True, what="发票外部 ID")
    head_b, lines_b = _records(hb, bb, line_col, move_col)
    head_a, lines_a = _records(ha, ba, line_col, move_col)
    expect = read_expect(expect_path) if expect_path else None

    problems, changes = [], []
    if text(head_b.get(move_col)) != text(head_a.get(move_col)):
        problems.append(f"不是同一张发票：{head_b.get(move_col)} vs {head_a.get(move_col)}")
    for i in head_b:
        if not _same(head_b[i], head_a.get(i)):
            (changes if hb[i] in ALLOWED_HEAD else problems).append(
                f"发票头 {hb[i]}：{head_b[i]} → {head_a.get(i)}")
    for lid in sorted(set(lines_b) - set(lines_a)):
        problems.append(f"行 {lid} 改后不见了")
    for lid in sorted(set(lines_a) - set(lines_b)):
        problems.append(f"多出了行 {lid}")

    pi = hb.index("Invoice lines/Unit Price")
    qi, di = hb.index("Invoice lines/Quantity"), hb.index("Invoice lines/Discount (%)")
    si, ti = hb.index("Invoice lines/Subtotal"), hb.index("Invoice lines/Total")
    for lid in [l for l in lines_b if l in lines_a]:
        b, a = lines_b[lid], lines_a[lid]
        for i in b:
            if not _same(b[i], a[i]):
                (changes if hb[i] in ALLOWED_LINE else problems).append(
                    f"行 {lid} {hb[i]}：{b[i]} → {a[i]}")
        new_p = num(a[pi])
        if expect is not None:
            want = expect.get(lid)
            if want is not None and abs(want - new_p) > 1e-9:
                problems.append(f"行 {lid} 单价应为 {want}，实际 {new_p}")
            if want is None and not _same(b[pi], a[pi]):
                problems.append(f"行 {lid} 不在导入表里，单价却变了：{b[pi]} → {a[pi]}")
        if not _same(b[pi], a[pi]):           # 改了价的行：小计 / 含税小计要和新价自洽
            mode, _ = tax_mode({"qty": num(b[qi]), "price": num(b[pi]), "disc": num(b[di]),
                                "sub": num(b[si]), "total": num(b[ti])})
            base = round(num(a[qi]) * new_p * (1 - num(a[di]) / 100), 2)
            key = {"0税": (si, ti), "净价": (si,), "含税": (ti,)}.get(mode, ())
            for k in key:
                if abs(num(a[k]) - base) > TOL:
                    problems.append(f"行 {lid} {hb[k]} = {a[k]}，按新单价应为 {base:.2f}")
    if expect is not None:
        for lid in sorted(set(expect) - set(lines_a)):
            problems.append(f"导入表里的行 {lid} 在改后导出里找不到")

    sub_sum = sum(num(lines_a[l][si]) for l in lines_a)
    tot_sum = sum(num(lines_a[l][ti]) for l in lines_a)
    ui, gi = find_col(ha, ba, ["Untaxed Amount"]), find_col(ha, ba, ["Total"])
    if abs(num(head_a[ui]) - sub_sum) > TOL:
        problems.append(f"Untaxed Amount {head_a[ui]} ≠ 各行 Subtotal 之和 {sub_sum:.2f}")
    if abs(num(head_a[gi]) - tot_sum) > TOL:
        problems.append(f"Total {head_a[gi]} ≠ 各行 Total 之和 {tot_sum:.2f}")

    say(f"允许范围内的变化 {len(changes)} 处：")
    for s in changes:
        say(f"  · {s}")
    if problems:
        say(f"\n⚠ 需要人看的 {len(problems)} 处：")
        for s in problems:
            say(f"  ⚠ {s}")
    else:
        say("\n✓ 除单价和由它重算的小计 / 合计外，没有别的字段变化；合计与各行之和一致。")
    return problems, changes


# --------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description="草稿客户发票按价格表批量改单价（纯文件进出，不连 Odoo）")
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build", help="发票导出 + 价格表导出 → 导入表 + 核对表 + 改前快照")
    b.add_argument("invoice", help="Odoo 发票导出 xlsx（固定导出模板，单张草稿发票）")
    b.add_argument("pricelist", help="Odoo 价格表导出 xlsx（固定导出模板，单张价格表）")
    b.add_argument("-o", "--outdir", help="输出目录（默认 output/YYYYMMDD）")
    v = sub.add_parser("verify", help="改前 / 改后两份发票导出逐字段比对")
    v.add_argument("before", help="改前导出（build 产出的改前快照即可）")
    v.add_argument("after", help="导入后用同一模板再导出的文件")
    v.add_argument("--expect", help="build 产出的导入表：核对单价是否逐行等于预期")
    args = ap.parse_args(argv)
    try:
        if args.cmd == "build":
            build(args.invoice, args.pricelist, args.outdir)
            return 0
        problems, _ = verify(args.before, args.after, args.expect)
        return 2 if problems else 0
    except InputError as e:
        say(f"✗ {e}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
