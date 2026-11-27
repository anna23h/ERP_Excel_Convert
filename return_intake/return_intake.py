#!/usr/bin/env python3
"""天猫 ECMS 退货入库：一次点货完成核对 + 批次 / 效期。（ISSUES: [退货入库] A）

三步，一个入口三个子命令（Odoo 一律只读）：

  sheet  退货批次导出 VoAbnormalBatchExport*.xlsx（可多份）--start N
         → 同一「大包号」并为一批，批间按退货申请时间升序编号 汇总N、N+1 …
         → 每批：汇总N.xlsx（SCP / LP / SKU / PZN / 品名 / 数量）+ 点货表_汇总N.xlsx（按 LP 一包一块）
  po     填好的 点货表_汇总N.xlsx（可多份）
         → 差异清单_汇总N.xlsx（包裹没到 / 少收 / 多收 / 多出产品 / 不可售）
         → 汇总N_PO导入.xlsx（按实收可售件数新建退货 PO，照以往退货 PO 的写法）
  lot    Odoo 入库单导出（lot_entry 同一导出模板）+ 填好的 点货表（可多份）
         → Lot_Import_*.xlsx + Lot_Check_*.xlsx（复用 lot_entry 的校验与导入格式，不用再打印 Lot_Entry）

用户定（2026-11-27）：PO 按实收建；不可售不入库（不进 PO，进差异清单）；点货表在仓库电脑上直接填；
包裹按 LP 号排；汇总表照旧单独出；汇总号由用户给起始号，Odoo 已有同号 PO 就停。

退出码：0 成功；1 输入 / 用法错误；2 护栏停下（不出表，原因打印在终端）。
"""
import argparse
import os
import re
import sys
from collections import OrderedDict, defaultdict
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

from openpyxl import Workbook, load_workbook  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402

from common.vendor import is_noise_vendor  # noqa: E402
from common.xlsx import unique_path  # noqa: E402


def _load_lot_entry():
    """按文件路径载入 lot_entry/lot_entry.py。

    `lot_entry/` 不是包，它自己的测试把该目录塞进 sys.path 后 `import lot_entry` 拿到的是模块本身，
    同一 pytest 进程里 `from lot_entry import lot_entry` 就会失败，所以不走包导入。
    """
    import importlib.util
    name = "_return_intake_lot_entry"
    if name in sys.modules:
        return sys.modules[name]
    spec = importlib.util.spec_from_file_location(name, os.path.join(ROOT, "lot_entry", "lot_entry.py"))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


L = _load_lot_entry()

# ---------------------------------------------------------------- 常量
EXPORT_COLS = ("退货申请批次号", "退货申请时间", "系统履约单号", "店铺订单号", "物流LP单号",
               "大包号", "异常拦截状态", "商品数量")
STATUS_OK = "已出库"
SO_PREFIXES = ("VO_TOF_", "GW_TOF_")

#: 以往退货 PO（P12694 等）的对手方：`common/vendor.NOISE_VENDOR_PATS` 里那家，库里 res.partner 2092。
VENDOR_ID = 2092
#: 产品进项税 → 退货 PO 用的进口增值税（2026-11-27 在 Odoo 查实：以往退货 PO 145/146 行照此映射）。
TAX_MAP = {
    "l10n_de_skr03.1_tax_vst_19_skr03": "l10n_de_skr03.1_tax_import_19_and_payable_skr03",
    "l10n_de_skr03.1_tax_vst_7_skr03": "l10n_de_skr03.1_tax_import_7_and_payable_skr03",
}
PO_COLS = ["partner_id/.id", "partner_ref", "order_line/product_id/.id", "order_line/product_qty",
           "order_line/price_unit", "order_line/taxes_id/id"]

META_SHEET = "_meta"
META_KIND = "return_intake_v1"
SHEET = "点货"
ARRIVED, MISSING = "到", "未到"      # 由实收推出：包裹所有行实收 0 = 未到

# 点货表列（1 起）
(C_LP, C_SCP, C_SKU, C_PZN, C_NAME, C_SOLL, C_YES, C_NO, C_GOT, C_BAD, C_LOT, C_MHD, C_NOTE) = range(1, 14)
HEADS = ["LP单号", "SCP单号", "SKU", "PZN", "品名", "应收", "到", "未到", "实收", "不可售", "批次", "MHD", "备注"]
WIDTHS = [20, 18, 26, 11, 52, 6, 5, 5, 7, 7, 16, 11, 24]
#: 「到」「未到」两列是勾选列（用户 2026-11-27：下拉步骤多）：填 1 显示 ✓；Excel 365 也可把这两列
#: 「插入 → 复选框」变成可点的框（存 TRUE/FALSE），两种都认。
CHECK_FMT = '"✓";"✓";""'

TECH = ["_type", "pid", "lp", "scp"]   # 隐藏列：P = 应收产品行，END = 表尾；操作员插的行为空
C_TECH0 = 15                      # O 起（N 留空作隔）

F_HEAD = PatternFill("solid", fgColor="1F3864")
F_BAD = PatternFill("solid", fgColor="F8CBAD")
F_WARN = PatternFill("solid", fgColor="FFE699")
THIN = Side(style="thin", color="A6A6A6")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


class InputError(Exception):
    """输入文件 / 参数不对 → 退出码 1。"""


class StopError(Exception):
    """护栏：数据对不上，不出表 → 退出码 2。"""


def say(*a):
    print(*a, flush=True)


def out_dir(dest=None):
    d = dest or os.path.join(ROOT, "output", date.today().strftime("%Y%m%d"))
    os.makedirs(d, exist_ok=True)
    return d


def digits(s):
    return re.sub(r"\D", "", str(s or ""))


def blank_pzn(sku, pzn):
    """汇总表口径：PZN 与 SKU 里某段数字相同就留空（PZN 只比数字，`17444652_` 也算相同）。"""
    d = digits(pzn)
    return "" if d and d in re.findall(r"\d+", sku or "") else (pzn or "")


def fmt_q(q):
    return int(q) if isinstance(q, float) and q.is_integer() else q


# ---------------------------------------------------------------- 1. 读退货导出
def _to_dt(v):
    if isinstance(v, datetime):
        return v
    return datetime.strptime(str(v).strip()[:19], "%Y-%m-%d %H:%M:%S")


def read_exports(paths):
    """→ [batch]，batch = {bag, apply_time, batch_nos, orders:[order]}，按退货申请时间升序。

    order = {scp, lp, shop_order, qty, status, batch_no, apply_time}；同一大包号的多份导出并为一批。
    """
    bags = OrderedDict()
    seen = {}
    for p in paths:
        wb = load_workbook(p, read_only=False)
        ws = wb.worksheets[0]
        rows = [r for r in ws.iter_rows(values_only=True) if any(c not in (None, "") for c in r)]
        if not rows:
            raise InputError(f"{os.path.basename(p)} 是空表")
        head = [str(h).strip() if h is not None else "" for h in rows[0]]
        miss = [c for c in EXPORT_COLS if c not in head]
        if miss:
            raise InputError(f"{os.path.basename(p)} 不像退货批次导出，缺列：{', '.join(miss)}")
        ix = {c: head.index(c) for c in EXPORT_COLS}
        for r in rows[1:]:
            g = lambda c: r[ix[c]]
            scp = str(g("系统履约单号") or "").strip()
            if not re.fullmatch(r"SCP\d+", scp):
                raise InputError(f"{os.path.basename(p)}：系统履约单号 {scp!r} 不是 SCP 单号")
            if scp in seen:
                raise StopError(f"SCP {scp} 在导出里出现两次（{seen[scp]} / {os.path.basename(p)}）——"
                                f"同一份导出给了两遍？")
            seen[scp] = os.path.basename(p)
            o = {"scp": scp, "lp": str(g("物流LP单号") or "").strip(),
                 "shop_order": str(g("店铺订单号") or "").strip(),
                 "qty": float(g("商品数量") or 0), "status": str(g("异常拦截状态") or "").strip(),
                 "batch_no": str(g("退货申请批次号") or "").strip(), "apply_time": _to_dt(g("退货申请时间"))}
            bags.setdefault(str(g("大包号") or "").strip(), []).append(o)
    batches = []
    for bag, orders in bags.items():
        batches.append({"bag": bag, "orders": orders,
                        "apply_time": min(o["apply_time"] for o in orders),
                        "batch_nos": sorted({o["batch_no"] for o in orders})})
    batches.sort(key=lambda b: b["apply_time"])
    return batches


# ---------------------------------------------------------------- 2. Odoo 只读
class OdooData:
    """从 Odoo 只读拉齐需要的数据；测试里用假对象替代（同样的属性）。

    so_lines: {scp: [(pid, qty)]}；prod: {pid: {code, pzn, name, active, type, tax_xmlids}}；
    kit: {套件 pid: [(组件 pid, 系数)]}；refs_taken(n) → 已占用 汇总n 的 PO 名列表；
    by_pzn(pzn) → [pid]（多出的产品按 PZN 找）。
    """

    def __init__(self, od=None):
        from odoo_api.odoo_client import Odoo
        import odoo_api.restock as rs
        self.od = od or Odoo.connect(verbose=False)
        _say, rs.say = rs.say, (lambda *a, **k: None)
        try:
            self.kit = rs.pull_kit_map(self.od)
        finally:
            rs.say = _say
        self.prod = {}
        self.so_lines = {}
        v = self.od.execute("res.partner", "read", [[VENDOR_ID], ["name"]])
        if not v or not is_noise_vendor(v[0]["name"]):
            raise StopError(f"Odoo res.partner {VENDOR_ID} 不是以往退货 PO 的供应商了，先查清再出表")

    def load_orders(self, scps):
        names = [p + s for s in scps for p in SO_PREFIXES]
        sos = self.od.search_read_all("sale.order", [("name", "in", names)], ["name", "order_line"])
        by_scp = defaultdict(list)
        for s in sos:
            by_scp[re.search(r"SCP\d+", s["name"]).group(0)].append(s)
        dup = {k: [s["name"] for s in v] for k, v in by_scp.items() if len(v) > 1}
        if dup:
            raise StopError(f"同一 SCP 在 Odoo 有多张销售单：{dup}")
        ids = [l for v in by_scp.values() for l in v[0]["order_line"]]
        lines = self.od.execute("sale.order.line", "read",
                                [ids, ["order_id", "product_id", "product_uom_qty", "display_type"]]) if ids else []
        sid2scp = {v[0]["id"]: k for k, v in by_scp.items()}
        for l in lines:
            if l.get("display_type") or not l["product_id"]:
                continue
            self.so_lines.setdefault(sid2scp[l["order_id"][0]], []).append(
                (l["product_id"][0], l["product_uom_qty"]))
        pids = {pid for v in self.so_lines.values() for pid, _ in v}
        self.load_products(pids | {c for p in pids for c, _ in self.kit.get(p, [])})

    def load_products(self, pids):
        pids = sorted(set(pids) - set(self.prod))
        if not pids:
            return
        rows = self.od.execute("product.product", "read",
                               [pids, ["default_code", "pzn", "name", "active", "type", "supplier_taxes_id"]],
                               {"context": {"active_test": False}})
        tax_ids = sorted({t for r in rows for t in r["supplier_taxes_id"]})
        xml = {}
        if tax_ids:
            for d in self.od.execute("ir.model.data", "search_read",
                                     [[("model", "=", "account.tax"), ("res_id", "in", tax_ids)]],
                                     {"fields": ["module", "name", "res_id"]}):
                xml[d["res_id"]] = f"{d['module']}.{d['name']}"
        for r in rows:
            self.prod[r["id"]] = {"code": (r["default_code"] or "").strip(), "pzn": (r["pzn"] or "").strip(),
                                  "name": (r["name"] or "").strip(), "active": r["active"], "type": r["type"],
                                  "tax_xmlids": [xml.get(t, str(t)) for t in r["supplier_taxes_id"]]}

    def refs_taken(self, n):
        pos = self.od.execute("purchase.order", "search_read",
                              [[("partner_ref", "=like", f"汇总{n} %")]], {"fields": ["name", "partner_ref"]})
        pos += self.od.execute("purchase.order", "search_read",
                               [[("partner_ref", "=", f"汇总{n}")]], {"fields": ["name", "partner_ref"]})
        return [p["name"] for p in pos]

    def by_pzn(self, pzn):
        rows = self.od.search_read_all("product.product", [("pzn", "=", pzn)], ["id"])
        pids = [r["id"] for r in rows if r["id"] not in self.kit]
        self.load_products(pids)
        return pids


def base_of(data, pid, qty):
    """订单行 → (基础品 pid, 实物件数, 套件说明)。单组件套件折回基础品；多组件套件停下。"""
    comps = data.kit.get(pid, [])
    if not comps:
        return pid, qty, ""
    if len(comps) > 1:
        raise StopError(f"{data.prod[pid]['code']} 是多组件套件，退货入库不知道怎么拆，请人工处理")
    c, f = comps[0]
    return c, qty * f, f"订单为套件 {data.prod[pid]['code']} ×{fmt_q(qty)}"


def check_product(data, pid):
    p = data.prod[pid]
    if not p["active"]:
        raise StopError(f"{p['code']} 在 Odoo 已归档")
    if p["type"] != "product":
        raise StopError(f"{p['code']} 不是库存产品（type={p['type']}）")
    if len(p["tax_xmlids"]) != 1 or p["tax_xmlids"][0] not in TAX_MAP:
        raise StopError(f"{p['code']} 的进项税不是单一的 19% / 7% Vorsteuer：{p['tax_xmlids']}")


def assemble(batch, data):
    """一批 → parcels：[{lp, scp, shop_order, status, sum_rows:[(sku, pid, qty)], items:[item]}]，按 LP 排。

    sum_rows 是汇总表口径（订单上的原始 SKU）；items 是点货口径（基础品，同包同品合并）。
    """
    parcels = []
    for o in batch["orders"]:
        lines = data.so_lines.get(o["scp"])
        if not lines:
            raise StopError(f"{o['scp']}（LP {o['lp']}）在 Odoo 找不到 VO_TOF_ / GW_TOF_ 销售单")
        q = sum(x for _, x in lines)
        if abs(q - o["qty"]) > 1e-9:
            raise StopError(f"{o['scp']} Odoo 订单行合计 {fmt_q(q)} 件 ≠ 导出商品数量 {fmt_q(o['qty'])}")
        items = OrderedDict()
        for pid, qty in lines:
            b, n, note = base_of(data, pid, qty)
            check_product(data, b)
            it = items.setdefault(b, {"pid": b, "soll": 0, "notes": []})
            it["soll"] += n
            if note:
                it["notes"].append(note)
        parcels.append(dict(o, sum_rows=[(data.prod[p]["code"], p, x) for p, x in lines],
                            items=list(items.values())))
    parcels.sort(key=lambda p: p["lp"])
    return parcels


def vendor_ref(n, batch):
    t = batch["apply_time"]
    return f"汇总{n}   {len(batch['orders'])}单ECMS退货包裹{t.year}年{t.month:02d}月{t.day:02d}日"


# ---------------------------------------------------------------- 3. 写汇总表 / 点货表
def write_summary(path, n, batch, data):
    wb = Workbook()
    ws = wb.active
    ws.title = f"汇总{n}"
    ws.append(["SCP单号", "LP单号", "SKU", "PZN", "产品名称", "数量"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
        c.alignment = Alignment(horizontal="center")
    tot = 0
    for o in sorted(batch["orders"], key=lambda o: o["apply_time"]):
        for pid, qty in data.so_lines[o["scp"]]:
            p = data.prod[pid]
            pzn = p["pzn"]
            if not pzn and len(data.kit.get(pid, [])) == 1:
                pzn = data.prod[data.kit[pid][0][0]]["pzn"]
            ws.append([o["scp"], o["lp"], p["code"], blank_pzn(p["code"], pzn), p["name"], fmt_q(qty)])
            tot += qty
    ws.append(["合计", f"{len(batch['orders'])} 单", None, None, None, fmt_q(tot)])
    for c in ws[ws.max_row]:
        c.font = Font(bold=True)
    for col, w in zip("ABCDEF", (18, 20, 26, 12, 60, 7)):
        ws.column_dimensions[col].width = w
    for (c,) in ws.iter_rows(min_row=2, min_col=4, max_col=4):
        c.number_format = "@"
    ws.freeze_panes = "A2"
    wb.save(path)
    return tot


def write_sheet(path, n, batch, parcels, data):
    """一个产品一行（LP / SCP 每行都写），不加包裹头行、不涂底色（用户 2026-11-27：少噪音）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    for i, (h, w) in enumerate(zip(HEADS, WIDTHS), start=1):
        c = ws.cell(1, i, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = F_HEAD
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[c.column_letter].width = w
    for j, h in enumerate(TECH):
        c = ws.cell(1, C_TECH0 + j, h)
        ws.column_dimensions[c.column_letter].hidden = True
    ws.freeze_panes = "A2"
    dv_num = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                            error="数量要填 ≥ 0 的数字", showErrorMessage=True)
    ws.add_data_validation(dv_num)
    r = 2
    for p in parcels:
        for it in p["items"]:
            pr = data.prod[it["pid"]]
            name = pr["name"] + (f"（{'；'.join(it['notes'])}）" if it["notes"] else "")
            row = {C_LP: p["lp"], C_SCP: p["scp"], C_SKU: pr["code"], C_PZN: pr["pzn"], C_NAME: name,
                   C_SOLL: fmt_q(it["soll"])}
            for col in range(1, len(HEADS) + 1):
                ws.cell(r, col, row.get(col)).border = BOX
            for k, v in enumerate(("P", it["pid"], p["lp"], p["scp"])):
                ws.cell(r, C_TECH0 + k, v)
            r += 1
    end = r
    ws.cell(end, C_TECH0, "END")
    dv_num.add(f"{ws.cell(2, C_GOT).coordinate}:{ws.cell(end + 200, C_BAD).coordinate}")
    for rr in range(2, end + 200):
        for col in (C_YES, C_NO):
            c = ws.cell(rr, col)
            c.number_format = CHECK_FMT
            c.alignment = Alignment(horizontal="center")
    # 操作员插的行（第二个批次 / 多出的产品）也要是文本格式：整列预设到表尾后 200 行
    for rr in range(2, end + 200):
        for col in (C_PZN, C_LOT, C_MHD):
            ws.cell(rr, col).number_format = "@"

    meta = wb.create_sheet(META_SHEET)
    for k, v in (("kind", META_KIND), ("n", n), ("vendor_ref", vendor_ref(n, batch)), ("bag", batch["bag"]),
                 ("batch_nos", ",".join(batch["batch_nos"])), ("orders", len(batch["orders"])),
                 ("made", date.today().isoformat())):
        meta.append([k, v])
    meta.sheet_state = "hidden"

    help_ = wb.create_sheet("说明", 1)
    for line in SHEET_HELP:
        help_.append([line])
    help_.column_dimensions["A"].width = 110
    wb.active = 0
    wb.save(path)


SHEET_HELP = [
    "点货表怎么填（在电脑上直接填）",
    "",
    "1. 一个产品一行，按 LP 号排序；拿起包裹，Ctrl+F 搜面单上的 LP 号，找到它的产品行（一个包裹有几个产品就有几行）。",
    "2. 每一行勾「到」或「未到」（填 1 就显示 ✓；Excel 365 可选中这两列 → 插入 → 复选框，变成点一下就勾）：",
    "   - 勾「到」、件数和「应收」一样：「实收」不用填；件数不一样才填「实收」。",
    "   - 勾「未到」：这个产品没收到（整包没到就把这包每一行都勾「未到」），「实收」空着。",
    "   - 两个都没勾、也没填「实收」会报错（= 还没点）；两个都勾也会报错。",
    "3. 「不可售」填破损、拆封、过期等不能再卖的件数，这部分不入库。",
    "4. 「批次」「MHD」照包装抄。MHD 写法同入库批次录入：05.28 / 05.2028 / 31.05.2028。包装上没有批次就空着。",
    "5. 同一产品有两个批次：在它下面插一行，只填「实收」「不可售」「批次」「MHD」。",
    "6. 包裹里多出订单上没有的产品：在这个包裹的最后一行下面插一行，填「PZN」「实收」，有批次也填上；品名可以不填。",
    "7. 其他情况写在「备注」，会抄到差异清单。",
    "",
    "填完交回，跑 po 生成差异清单和 PO 导入表；PO 在 Odoo 确认后导出入库单，再跑 lot 生成批次导入表。",
]


def cmd_sheet(paths, start, dest=None, data=None):
    batches = read_exports(paths)
    data = data or OdooData()
    data.load_orders([o["scp"] for b in batches for o in b["orders"]])
    dest = out_dir(dest)
    plan = []
    for k, b in enumerate(batches):
        n = start + k
        taken = data.refs_taken(n)
        if taken:
            raise StopError(f"Odoo 里已有 Vendor Reference 为 汇总{n} 的采购单：{', '.join(taken)}——起始号给错了？")
        plan.append((n, b, assemble(b, data)))
    files, lines = [], []
    for n, b, parcels in plan:
        s = unique_path(os.path.join(dest, f"汇总{n}.xlsx"))
        tot = write_summary(s, n, b, data)
        c = unique_path(os.path.join(dest, f"点货表_汇总{n}.xlsx"))
        write_sheet(c, n, b, parcels, data)
        files += [s, c]
        n_items = sum(len(p["items"]) for p in parcels)
        soll = sum(it["soll"] for p in parcels for it in p["items"])
        lines.append(f"汇总{n}: 大包 {b['bag']}（{', '.join(b['batch_nos'])}）{len(b['orders'])} 单 / "
                     f"{fmt_q(tot)} 件；点货表 {len(parcels)} 包 {n_items} 行，实物应收 {fmt_q(soll)} 件")
        bad = sorted({o["status"] for o in b["orders"] if o["status"] != STATUS_OK})
        if bad:
            lines.append(f"  ⚠ 异常拦截状态是 {'/'.join(bad)}（不是{STATUS_OK}）：包裹可能还没到仓，点货等实物到了再做")
    return files, lines


# ---------------------------------------------------------------- 4. 读回点货表
def _txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()


def _qty(v, row, what, errs):
    if v in (None, ""):
        return None
    q = L.num(v)
    if q is None or q < 0:
        errs.append((row, f"{what} {v!r} 不是 ≥ 0 的数字"))
        return None
    return q


def read_sheet(path):
    """→ (meta, parcels, errs)。parcels = [{lp, scp, row, arrived, note, items:[item]}]，按表里出现顺序；

    item = {row, pid|None, pzn, soll(多出为 0), extra, lots:[{row, got, bad, lot, mhd}]}。
    行归属：P 行（隐藏列，带 pid / LP / SCP）是应收产品；操作员插的行填了 PZN = 多出的产品，
    否则 = 上一个产品的续批次；都归上一 P 行的包裹。包裹「未到」= 它所有行实收都是 0、不可售也没填。
    errs = [(row, 说明)]：只做「表填得对不对」的检查，不连 Odoo。
    """
    wb = load_workbook(path)
    if META_SHEET not in wb.sheetnames or SHEET not in wb.sheetnames:
        raise InputError(f"{os.path.basename(path)} 不是本工具生成的点货表")
    meta = {r[0]: r[1] for r in wb[META_SHEET].iter_rows(values_only=True) if r[0]}
    if meta.get("kind") != META_KIND:
        raise InputError(f"{os.path.basename(path)} 不是本工具生成的点货表")
    ws = wb[SHEET]
    parcels, errs = OrderedDict(), []
    cur_p = cur_i = None
    for r in range(2, ws.max_row + 1):
        t = ws.cell(r, C_TECH0).value
        g = lambda c: ws.cell(r, c).value
        if t == "END":
            break
        filled = any(g(c) not in (None, "") for c in (C_PZN, C_GOT, C_BAD, C_LOT, C_MHD, C_NOTE)) or \
            _checked(g(C_YES)) or _checked(g(C_NO))
        if t != "P" and not filled:
            continue
        lot = {"row": r, "got": _qty(g(C_GOT), r, "实收", errs), "bad": _qty(g(C_BAD), r, "不可售", errs),
               "lot": _txt(g(C_LOT)), "mhd": g(C_MHD), "note": _txt(g(C_NOTE)),
               "yes": _checked(g(C_YES)), "no": _checked(g(C_NO))}
        if t == "P":
            lp, scp = (_txt(ws.cell(r, C_TECH0 + k).value) for k in (2, 3))
            cur_p = parcels.setdefault(lp, {"lp": lp, "scp": scp, "row": r, "arrived": ARRIVED, "note": "",
                                            "items": []})
            cur_i = {"row": r, "pid": int(ws.cell(r, C_TECH0 + 1).value), "sku": _txt(g(C_SKU)),
                     "pzn": _txt(g(C_PZN)), "name": _txt(g(C_NAME)), "soll": L.num(g(C_SOLL)) or 0,
                     "extra": False, "lots": [lot]}
            cur_p["items"].append(cur_i)
        elif cur_p is None:
            errs.append((r, "这一行在第一个产品行上面，不知道属于哪个包裹"))
        elif _txt(g(C_PZN)):                       # 操作员插的行填了 PZN = 多出的产品
            cur_i = {"row": r, "pid": None, "sku": "", "pzn": _txt(g(C_PZN)), "name": _txt(g(C_NAME)),
                     "soll": 0, "extra": True, "lots": [lot]}
            cur_p["items"].append(cur_i)
        else:                                      # 续行 = 上面那个产品的第 2、3 … 个批次
            cur_i["lots"].append(lot)
    parcels = list(parcels.values())
    for p in parcels:
        for it in p["items"]:
            _check_item(p, it, errs)
        lots = [l for it in p["items"] for l in it["lots"]]
        if all(l["got"] == 0 and not l["bad"] for l in lots):
            p["arrived"] = MISSING
    return meta, parcels, errs


def _checked(v):
    """勾选格：TRUE（Excel 复选框）/ 1 / x / ✓ … 都算勾；空、0、FALSE 算没勾。"""
    if v is None or v is False:
        return False
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return v != 0
    return str(v).strip().upper() not in ("", "0", "FALSE")


def _check_item(p, it, errs):
    lots = it["lots"]
    label = it["sku"] or f"PZN {it['pzn']}"
    first = lots[0]
    if first["yes"] and first["no"]:
        errs.append((first["row"], f"{label}（LP {p['lp']}）「到」「未到」都勾了"))
        first["got"] = 0 if first["got"] is None else first["got"]      # 只报这一条
    elif first["no"]:
        if first["got"] or first["bad"] or len(lots) > 1:
            errs.append((first["row"], f"{label}（LP {p['lp']}）勾了「未到」，却填了实收 / 不可售 / 批次行"))
        first["got"] = 0
    elif first["yes"] and first["got"] is None and len(lots) == 1 and not it["extra"]:
        first["got"] = it["soll"]                  # 勾「到」不填实收 = 和应收一样
    for l in lots:
        if l["got"] is None:
            errs.append((l["row"], f"{label}（LP {p['lp']}）没勾「到」「未到」也没填实收——没点？"
                         + ("（同一产品有多个批次时每行都要填实收）" if len(lots) > 1 else "")))
    if len(lots) > 1 and not all(l["lot"] for l in lots):
        errs.append((it["row"], f"{label} 有 {len(lots)} 行：每行都要填批次"))
    seen = set()
    for l in lots:
        if l["got"] is not None and l["bad"] is not None and l["bad"] > l["got"]:
            errs.append((l["row"], f"不可售 {fmt_q(l['bad'])} 比实收 {fmt_q(l['got'])} 还多"))
        if l["lot"]:
            if not L.LOT_RE.match(l["lot"]):
                errs.append((l["row"], f"批次 {l['lot']!r}：只能有字母、数字和 - / .，3–20 位"))
            if l["lot"].upper() in seen:
                errs.append((l["row"], f"批次 {l['lot']!r} 在同一产品下写了两次"))
            seen.add(l["lot"].upper())
        d, err = L.parse_mhd(l["mhd"])
        if err:
            errs.append((l["row"], err))
        elif d is not None:
            if d <= date.today():
                errs.append((l["row"], f"MHD {d:%d.%m.%Y} 已过期——过期货算不可售"))
            elif d > L.MHD_MAX(date.today()):
                errs.append((l["row"], f"MHD {d:%d.%m.%Y} 超过 {L.MHD_MAX_YEARS} 年，笔误？"))
            if not l["lot"]:
                errs.append((l["row"], "写了 MHD 没写批次：Odoo 的 MHD 挂在批次上，没有批次存不进去"))
        l["mhd_date"] = d


# ---------------------------------------------------------------- 5. 差异清单 + PO
def resolve_extras(parcels, data, errs):
    for p in parcels:
        for it in p["items"]:
            if not it["extra"]:
                continue
            pids = data.by_pzn(it["pzn"])
            if len(pids) != 1:
                errs.append((it["row"], f"PZN {it['pzn']} 在 Odoo 找到 {len(pids)} 个（非套件）产品，"
                                        f"{'没有这个 PZN' if not pids else '不唯一'}——请改成正确的 PZN 或删掉这行写进备注"))
                continue
            it["pid"] = pids[0]
            pr = data.prod[it["pid"]]
            it["sku"], it["name"] = pr["code"], it["name"] or pr["name"]
            try:
                check_product(data, it["pid"])
            except StopError as e:
                errs.append((it["row"], str(e)))


def diff_rows(parcels):
    """→ [(LP, SCP, SKU, PZN, 品名, 应收, 实收, 不可售, 入库, 类型, 备注)]，只列有差异的。"""
    out = []
    for p in parcels:
        if p["arrived"] == MISSING:
            for it in p["items"]:
                out.append((p["lp"], p["scp"], it["sku"], it["pzn"], it["name"], fmt_q(it["soll"]), 0, 0, 0,
                            "包裹未到", p["note"]))
            continue
        for it in p["items"]:
            got = sum(l["got"] or 0 for l in it["lots"])
            bad = sum(l["bad"] or 0 for l in it["lots"])
            kinds = []
            if it["extra"]:
                kinds.append("多出产品")
            elif got < it["soll"]:
                kinds.append("少收")
            elif got > it["soll"]:
                kinds.append("多收")
            if bad:
                kinds.append("不可售")
            notes = "；".join(l["note"] for l in it["lots"] if l["note"])
            if kinds or notes:
                out.append((p["lp"], p["scp"], it["sku"], it["pzn"], it["name"],
                            None if it["extra"] else fmt_q(it["soll"]), fmt_q(got), fmt_q(bad),
                            fmt_q(got - bad), "、".join(kinds) or "备注", notes))
    return out


def po_qty(parcels):
    """→ OrderedDict pid → 可售件数（实收 − 不可售），跨包合并，0 件的不进。"""
    q = defaultdict(float)
    for p in parcels:
        if p["arrived"] != ARRIVED:
            continue
        for it in p["items"]:
            for l in it["lots"]:
                q[it["pid"]] += (l["got"] or 0) - (l["bad"] or 0)
    return OrderedDict((k, v) for k, v in q.items() if v > 0)


def write_diff(path, n, rows, totals):
    wb = Workbook()
    ws = wb.active
    ws.title = f"差异_汇总{n}"
    ws.append(["LP单号", "SCP单号", "SKU", "PZN", "品名", "应收", "实收", "不可售", "入库", "类型", "备注"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    for r in rows:
        ws.append(list(r))
        if any(k in (r[9] or "") for k in ("未到", "少收", "不可售")):
            ws.cell(ws.max_row, 10).fill = F_WARN
    if not rows:
        ws.append(["（没有差异：全部包裹都到了，件数与订单一致，没有不可售）"])
    ws.append([])
    for k, v in totals:
        ws.append([k, v])
        ws.cell(ws.max_row, 1).font = Font(bold=True)
    for col, w in zip("ABCDEFGHIJK", (20, 18, 26, 11, 50, 6, 6, 7, 6, 14, 40)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(path)


def write_po(path, ref, qty, data):
    wb = Workbook()
    ws = wb.active
    ws.title = "purchase.order"
    ws.append(PO_COLS)
    for c in ws[1]:
        c.font = Font(bold=True)
    chk = wb.create_sheet("核对")
    chk.append(["SKU", "PZN", "产品名称", "数量", "税"])
    for c in chk[1]:
        c.font = Font(bold=True)
    items = sorted(qty.items(), key=lambda x: data.prod[x[0]]["code"])
    for k, (pid, q) in enumerate(items):
        p = data.prod[pid]
        tax = TAX_MAP[p["tax_xmlids"][0]]
        ws.append([VENDOR_ID if k == 0 else None, ref if k == 0 else None, pid, fmt_q(q), 0, tax])
        chk.append([p["code"], p["pzn"], p["name"], fmt_q(q), "7%" if "_7_" in tax else "19%"])
    chk.append(["合计", None, f"{len(items)} 个产品", fmt_q(sum(qty.values())), None])
    for col, w in zip("ABCDEF", (16, 26, 26, 22, 14, 52)):
        ws.column_dimensions[col].width = w
    for col, w in zip("ABCDE", (26, 12, 60, 8, 6)):
        chk.column_dimensions[col].width = w
    wb.save(path)


def cmd_po(paths, dest=None, data=None):
    sheets = [(p,) + read_sheet(p) for p in paths]
    data = data or OdooData()
    dest = out_dir(dest)
    problems = []
    for p, meta, parcels, errs in sheets:
        resolve_extras(parcels, data, errs)
        data.load_products({it["pid"] for pc in parcels for it in pc["items"] if it["pid"]})
        for pc in parcels:
            for it in pc["items"]:
                if it["pid"] and not it["extra"]:
                    try:
                        check_product(data, it["pid"])
                    except StopError as e:
                        errs.append((it["row"], str(e)))
        taken = data.refs_taken(meta["n"])
        if taken:
            errs.append((1, f"Odoo 里已有 汇总{meta['n']} 的采购单：{', '.join(taken)}——这批已经建过 PO？"))
        problems += [f"{os.path.basename(p)} 第 {r} 行：{m}" for r, m in sorted(errs)]
    if problems:
        raise StopError("点货表有下面这些问题，改好再跑（没有出任何文件）：\n  " + "\n  ".join(problems))
    files, lines = [], []
    for p, meta, parcels, _ in sheets:
        n = meta["n"]
        qty = po_qty(parcels)
        rows = diff_rows(parcels)
        arrived = [pc for pc in parcels if pc["arrived"] == ARRIVED]
        got = sum(l["got"] or 0 for pc in arrived for it in pc["items"] for l in it["lots"])
        bad = sum(l["bad"] or 0 for pc in arrived for it in pc["items"] for l in it["lots"])
        soll = sum(it["soll"] for pc in parcels for it in pc["items"])
        totals = [("包裹", f"{len(arrived)} / {len(parcels)} 到"), ("应收", fmt_q(soll)), ("实收", fmt_q(got)),
                  ("不可售", fmt_q(bad)), ("入库（进 PO）", fmt_q(sum(qty.values())))]
        d = unique_path(os.path.join(dest, f"差异清单_汇总{n}.xlsx"))
        write_diff(d, n, rows, totals)
        files.append(d)
        lines.append(f"汇总{n}: " + "，".join(f"{k} {v}" for k, v in totals) + f"；差异 {len(rows)} 行")
        if qty:
            po = unique_path(os.path.join(dest, f"汇总{n}_PO导入.xlsx"))
            write_po(po, meta["vendor_ref"], qty, data)
            files.append(po)
            lines.append(f"  PO：{meta['vendor_ref']}，{len(qty)} 行 {fmt_q(sum(qty.values()))} 件")
        else:
            lines.append("  没有可入库的货，不出 PO 导入表")
    return files, lines


# ---------------------------------------------------------------- 6. 入库单导出 + 点货表 → Lot_Import
def lots_by_code(sheets):
    """多份点货表 → {SKU: [(row 标签, lot, mhd 原值, 可售件数)]}，同 SKU 同批次跨包合并。"""
    acc = OrderedDict()
    for path, meta, parcels in sheets:
        for p in parcels:
            if p["arrived"] != ARRIVED:
                continue
            for it in p["items"]:
                for l in it["lots"]:
                    q = (l["got"] or 0) - (l["bad"] or 0)
                    if q <= 0:
                        continue
                    by_lot = acc.setdefault(it["sku"], OrderedDict())
                    key = l["lot"].upper()
                    if key in by_lot:
                        e = by_lot[key]
                        if e["mhd_date"] != l["mhd_date"]:
                            raise StopError(f"{it['sku']} 批次 {l['lot']} 两处 MHD 不一样（第 {e['row']} 行 / "
                                            f"第 {l['row']} 行）")
                        e["qty"] += q
                    else:
                        by_lot[key] = {"row": l["row"], "lot": l["lot"], "mhd": l["mhd"],
                                       "mhd_date": l["mhd_date"], "qty": q}
    return {sku: [(e["row"], e["lot"], e["mhd"], e["qty"]) for e in v.values()] for sku, v in acc.items()}


def cmd_lot(export_path, sheet_paths, dest=None, data=None):
    sheets = []
    problems = []
    for p in sheet_paths:
        meta, parcels, errs = read_sheet(p)
        if any(it["extra"] for pc in parcels for it in pc["items"]):
            data = data or OdooData()          # 多出的产品只有 PZN，SKU 要按 PZN 去 Odoo 找（同 po 那步）
            resolve_extras(parcels, data, errs)
        problems += [f"{os.path.basename(p)} 第 {r} 行：{m}" for r, m in errs]
        sheets.append((p, meta, parcels))
    if problems:
        raise StopError("点货表有问题，先改好：\n  " + "\n  ".join(problems))
    wb = load_workbook(export_path)
    if L.is_entry_sheet(wb):
        raise InputError("这是 Lot_Entry 录入表，要的是 Odoo 入库单导出（Inventory → Transfers → Export）")
    try:
        pickings = L.read_export(wb.active)
    except L.InputError as e:
        raise InputError(f"入库单导出读不了：{e}")

    codes_in_export = {}
    for pid_, pk in pickings.items():
        for mid, mv in pk["moves"].items():
            code, name = L.split_product(mv["product"])
            if code in codes_in_export:
                raise StopError(f"{code} 在入库单导出里出现在两个 move 上，请分开导出")
            codes_in_export[code] = (pid_, mid, mv, name)
    want = lots_by_code(sheets)

    miss = sorted(set(want) - set(codes_in_export))
    extra = sorted(set(codes_in_export) - set(want))
    diff = []
    for code, lots in want.items():
        if code in codes_in_export:
            dem = codes_in_export[code][2]["demand"]
            tot = sum(q for *_, q in lots)
            if abs(dem - tot) > 1e-9:
                diff.append(f"{code}：入库单应收 {fmt_q(dem)}，点货表可售 {fmt_q(tot)}")
    if miss or extra or diff:
        msg = []
        if miss:
            msg.append("点货表有、入库单没有：" + ", ".join(miss))
        if extra:
            msg.append("入库单有、点货表没有：" + ", ".join(extra))
        msg += diff
        raise StopError("入库单和点货表对不上（导出的是这批 PO 的入库单吗？）：\n  " + "\n  ".join(msg))

    blocks = []
    for code, (pid_, mid, mv, name) in codes_in_export.items():
        lots = want[code]
        blocks.append({"picking_id": pid_, "move_id": mid, "line_ids": "|".join(mv["lines"]),
                       "product": mv["product"], "uom": mv["uom"], "loc_from": mv["loc_from"],
                       "loc_to": mv["loc_to"], "code": code, "name": name, "soll": mv["demand"],
                       "row": lots[0][0],
                       "entries": [(row, lot, mhd, q) for row, lot, mhd, q in lots]})
    dest = out_dir(dest)
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    imp = os.path.join(dest, f"Lot_Import_{stamp}.xlsx")
    chk = os.path.join(dest, f"Lot_Check_{stamp}.xlsx")
    n, check = L.build_import(blocks, imp, chk)
    cnt = {k: sum(1 for c in check if c[0] == k) for k in ("ERROR", "WARN", "NO LOT")}
    lines = [f"入库单 {len(pickings)} 张，产品 {len(blocks)} 个；写进导入表 {n} 个",
             f"ERROR {cnt['ERROR']}  WARN {cnt['WARN']}  无批次 {cnt['NO LOT']}"]
    lines += [f"  {c[0]:<6} {c[2]}  {c[6]}" for c in check if c[0] in ("ERROR", "WARN")]
    no_lot = [c[2] for c in check if c[0] == "NO LOT"]
    if no_lot:
        lines.append("  没有批次（包装上没有？这些只按数量收货，不进导入表）：" + ", ".join(no_lot))
    return ([imp] if n else []) + [chk], lines


# ---------------------------------------------------------------- 入口
def main(argv=None):
    ap = argparse.ArgumentParser(description="天猫 ECMS 退货入库：点货表 → 差异清单 + PO 导入表 → 批次导入表")
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("sheet", help="退货批次导出 → 汇总表 + 点货表")
    a.add_argument("exports", nargs="+", help="VoAbnormalBatchExport*.xlsx，可多份")
    a.add_argument("--start", type=int, required=True, help="第一批的汇总号（如 139）")
    b = sub.add_parser("po", help="填好的点货表 → 差异清单 + PO 导入表")
    b.add_argument("sheets", nargs="+", help="点货表_汇总N.xlsx，可多份")
    c = sub.add_parser("lot", help="Odoo 入库单导出 + 点货表 → Lot_Import + Lot_Check")
    c.add_argument("export", help="Odoo 入库单导出（lot_entry 同一导出模板）")
    c.add_argument("sheets", nargs="+", help="点货表_汇总N.xlsx，可多份")
    for p in (a, b, c):
        p.add_argument("-o", "--out", help="产出目录（默认 output/YYYYMMDD/）")
    args = ap.parse_args(argv)
    try:
        if args.cmd == "sheet":
            files, lines = cmd_sheet(args.exports, args.start, args.out)
        elif args.cmd == "po":
            files, lines = cmd_po(args.sheets, args.out)
        else:
            files, lines = cmd_lot(args.export, args.sheets, args.out)
    except InputError as e:
        say(f"✗ {e}")
        return 1
    except StopError as e:
        say(f"✗ 停下，不出表：{e}")
        return 2
    for l in lines:
        say(l)
    for f in files:
        say(f"→ {f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
