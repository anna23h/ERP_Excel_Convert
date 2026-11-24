# -*- coding: utf-8 -*-
"""补货预判清单的「动态安全库存」+ 套件并入基础品：Odoo 只读拉数（ISSUES [安全库存] A）。

为什么：ERP 产品上的 `Safety Stock` 是静态值，除了用户本人没人看、也没人维护，
2026-11-08 实测非套件 101 个里 37% 与近期销量明显脱节，另有 59 个配在套件上
（套件没有自己的实物，等于同一批货被看两次）。用户定（2026-11-08）：
  - 不再读 ERP 字段、也不回写；清单直接显示按销量算的动态值（方案 1）。
  - 安全库存 = **12 周 VO/GW 零售周均 × 2**（覆盖两周）；**周均 < 10 件不设**（留空）。
  - 套件 / 店铺后缀品**并进基础品那一行**一起算：今日需求 = 基础品 + 套件 × BoM 数量，
    在手取基础品的 qty_available。
  - 运营的安全库存表不定期，不参与。

口径与备货汇总（odoo_api/restock.py）同一套：订单号 `=like 'VO%' / 'GW%'`、state sale/done、
phantom BoM 折算（直接复用 restock 的 pull_kit_map / roll_sales），周均按窗口天数折周。
窗口截止**昨天**（今天的单还没收完）。

拉不到（没配凭据 / 连不上）由调用方降级：套件不并、两列不出，清单照出。
"""
import csv
import os
import re
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

WEEKS = 12
COVER_WEEKS = 2
MIN_AVG = 10          # 周均低于此不设安全库存
COL_AVG = f"周均({WEEKS}周)"
COL_SAFETY = f"安全库存({COVER_WEEKS}周)"
COL_KITS = "含套件"


def safety_of(avg, cover=COVER_WEEKS, min_avg=MIN_AVG):
    """周均 → 安全库存（整数）；周均 < min_avg 返回 None（不设）。"""
    if avg is None or avg < min_avg:
        return None
    return int(round(avg * cover))


def kit_code_map(kit_map, code_of, pid_of_sku):
    """{套件 pid: [(组件 pid, 系数)]} → {套件 SKU: (基础品 SKU, 系数)}，只管清单里出现的 SKU。

    多组件套件无法并进单一一行，不并、单列返回让调用方点名（2026-11-08 实测 129 个全是单组件）。
    → (映射, 多组件套件 SKU 列表)
    """
    out, multi = {}, []
    for sku, pid in pid_of_sku.items():
        comps = kit_map.get(pid)
        if not comps:
            continue
        if len(comps) != 1:
            multi.append(sku)
            continue
        base, k = comps[0]
        code = code_of.get(base, "")
        if code:
            out[sku] = (code, k)
    return out, sorted(multi)


def _kit_lookup(od, skus):
    """SKU → Odoo pid + 单组件套件映射。→ (pid_of, 重码数, kit_map, code_of, kits, multi)。"""
    from odoo_api.restock import pull_kit_map, pull_codes
    import odoo_api.restock as rs

    prods = od.search_read_all("product.product", [("default_code", "in", skus)], ["default_code"])
    pid_of, dup = {}, 0
    for p in prods:                      # 按 id 升序，同码取最早建档的（同 pending_moves）
        if p["default_code"] in pid_of:
            dup += 1
            continue
        pid_of[p["default_code"]] = p["id"]

    _say, rs.say = rs.say, (lambda *a, **k: None)   # restock 的进度打印进 VOTool 日志是噪音
    try:
        kit_map = pull_kit_map(od)
        base_pids = {c for pid in pid_of.values() for c, _ in kit_map.get(pid, [])}
        code_of = {pid: v["code"] for pid, v in pull_codes(od, base_pids).items()} if base_pids else {}
    finally:
        rs.say = _say
    kits, multi = kit_code_map(kit_map, code_of, pid_of)
    return pid_of, dup, kit_map, code_of, kits, multi


def pull_kits(skus, timeout=60):
    """只拉套件映射（ISSUES [拣货面单] Q：面单按基础品聚合）。→ (kits={套件 SKU: (基础品 SKU, 系数)}, 说明)。
    拉不到时抛异常，由调用方降级。"""
    from odoo_api.odoo_client import Odoo

    od = Odoo.connect(timeout=timeout, verbose=False)
    skus = sorted({str(s).strip() for s in skus if str(s).strip()})
    pid_of, dup, _, _, kits, multi = _kit_lookup(od, skus)
    info = f"{len(pid_of)}/{len(skus)} 个 SKU 在 Odoo 找到，{len(kits)} 个套件归到基础品"
    if multi:
        info += f"；⚠ 多组件套件 {len(multi)} 个不归并: {', '.join(multi[:5])}"
    if dup:
        info += f"；⚠ {dup} 个 Internal Reference 在 Odoo 有重码，取最早建档的"
    return kits, info


# ---------- 本地套件对照表（ISSUES [拣货面单] S）----------
# 同事机器不连 Odoo 也能把套件归到基础品：开发机用 export_kits.py 导出，同事把文件放到程序旁。
# 用户定（2026-11-24）：**有本地对照表就只用它、不连 Odoo**；固定文件名、自动找，不加界面栏位。
KITS_FILE = "套件对照表.csv"
KITS_COLS = ("套件SKU", "基础品SKU", "系数", "导出日期")
KITS_STALE_DAYS = 30
_KIT_LIKE = re.compile(r"(x\d+|_(VO|GW))$", re.I)   # 名字像套件：…x2 / …_VO / …_GW


def kits_file_dirs():
    """找对照表的目录，按顺序：打包 exe 所在目录；仓库根（源码版）。"""
    dirs = []
    if getattr(sys, "frozen", False):
        dirs.append(os.path.dirname(sys.executable))
    dirs.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return dirs


def find_kits_file(dirs=None):
    """→ 第一个存在的对照表路径；都没有返回 None。"""
    for d in dirs if dirs is not None else kits_file_dirs():
        p = os.path.join(d, KITS_FILE)
        if os.path.isfile(p):
            return p
    return None


def load_kits(path, skus=None, today=None):
    """读本地对照表 → (kits={套件 SKU: (基础品 SKU, 系数)}, 说明, 导出日期 date 或 None, known=表里全部套件+基础品 SKU)。

    skus 给了就只留这批出现的套件；文件格式不对抛 ValueError（调用方降级）。
    说明里带导出日期，超过 KITS_STALE_DAYS 天加 ⚠ 提醒更新。
    """
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if rows and not set(KITS_COLS[:3]) <= set(rows[0]):
        raise ValueError(f"{KITS_FILE} 缺列，应有：{' / '.join(KITS_COLS)}")
    want = None if skus is None else {str(s).strip() for s in skus}
    kits, dates, known = {}, set(), set()
    for i, r in enumerate(rows, start=2):
        kit, base = (r.get("套件SKU") or "").strip(), (r.get("基础品SKU") or "").strip()
        if not kit or not base:
            continue
        try:
            k = float(r.get("系数") or "")
        except ValueError:
            raise ValueError(f"{KITS_FILE} 第 {i} 行系数不是数字：{r.get('系数')!r}") from None
        if k <= 0:
            raise ValueError(f"{KITS_FILE} 第 {i} 行系数必须大于 0：{k}")
        known.update((kit, base))
        if (r.get("导出日期") or "").strip():
            dates.add(r["导出日期"].strip())
        if want is None or kit in want:
            kits[kit] = (base, int(k) if k == int(k) else k)
    exported = None
    if dates:
        try:
            exported = max(date.fromisoformat(d) for d in dates)
        except ValueError:
            exported = None
    info = f"本地 {KITS_FILE}（{len(rows)} 条，本批用到 {len(kits)} 个套件"
    info += f"，导出于 {exported.isoformat()}）" if exported else "，导出日期不明）"
    age = ((today or date.today()) - exported).days if exported else None
    if age is None:
        info += "；⚠ 对照表没有导出日期，请找 Anna 要新的"
    elif age > KITS_STALE_DAYS:
        info += f"；⚠ 对照表已导出 {age} 天（超过 {KITS_STALE_DAYS} 天），请找 Anna 要新的"
    return kits, info, exported, known


def kit_like_missing(skus, known):
    """本批 SKU 里名字像套件、对照表里（套件和基础品两列）都没有的（对照表可能过期）→ 排序列表。
    11-24 全量 568 个套件里 514 个名字命中；也有 12 个基础品命中，所以基础品列也要算「已知」。"""
    return sorted({str(s).strip() for s in skus
                   if _KIT_LIKE.search(str(s).strip()) and str(s).strip() not in known})


def export_kits(out_path, timeout=60, today=None):
    """Odoo 只读 → 写对照表 CSV（全部单组件 phantom BoM）。→ (条数, 多组件套件 SKU 列表, 重码数)。"""
    from odoo_api.odoo_client import Odoo
    from odoo_api.restock import pull_kit_map, pull_codes

    od = Odoo.connect(timeout=timeout, verbose=False)
    kit_map = pull_kit_map(od)
    pids = set(kit_map) | {c for comps in kit_map.values() for c, _ in comps}
    code_of = {pid: v["code"] for pid, v in pull_codes(od, pids).items()} if pids else {}
    stamp = (today or date.today()).isoformat()
    rows, multi, seen, dup = [], [], set(), 0
    for pid in sorted(kit_map):                      # 同码取最早建档的（同 _kit_lookup）
        code = code_of.get(pid, "")
        if not code:
            continue
        if code in seen:
            dup += 1
            continue
        seen.add(code)
        comps = kit_map[pid]
        if len(comps) != 1:
            multi.append(code)
            continue
        base, k = comps[0]
        if code_of.get(base) and k > 0:
            rows.append((code, code_of[base], int(k) if k == int(k) else k, stamp))
    rows.sort(key=lambda r: r[0].lower())
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:   # utf-8-sig：Excel 直接打开不乱码
        w = csv.writer(f)
        w.writerow(KITS_COLS)
        w.writerows(rows)
    return len(rows), sorted(multi), dup


def pull_safety(skus, until=None, weeks=WEEKS, timeout=60):
    """→ dict(kits={套件 SKU: (基础品 SKU, 系数)}, avg={SKU: 周均}, onhand={基础品 SKU: 在手}, info=说明)。

    avg 只给并入后清单里会出现的 SKU（非套件 SKU + 套件的基础品）。拉不到时抛异常，由调用方降级。
    """
    from odoo_api.odoo_client import Odoo, m2o_id
    from odoo_api.restock import SOLD_STATES, prefix_domain, roll_sales
    from odoo_api.stock_report import pull_qty_available

    od = Odoo.connect(timeout=timeout, verbose=False)
    until = until or (date.today() - timedelta(days=1))
    since = until - timedelta(days=weeks * 7 - 1)
    n_days = (until - since).days + 1

    skus = sorted({str(s).strip() for s in skus if str(s).strip()})
    pid_of, dup, kit_map, code_of, kits, multi = _kit_lookup(od, skus)

    dom = [("date", ">=", since.isoformat()), ("date", "<=", f"{until.isoformat()} 23:59:59"),
           ("state", "in", SOLD_STATES)] + prefix_domain()
    rows = od.read_group_all("sale.report", dom, ["product_uom_qty:sum"], ["product_id", "date:week"])
    weekly = [(m2o_id(r["product_id"]), r["date:week"], r.get("product_uom_qty") or 0)
              for r in rows if m2o_id(r.get("product_id")) is not None]
    by_base = roll_sales(weekly, kit_map)
    tot = {pid: sum(w.values()) for pid, w in by_base.items()}

    # 并入后清单里的 SKU → pid：非套件用自身，基础品用 BoM 组件 id
    final = {sku: pid for sku, pid in pid_of.items() if sku not in kits}
    base_pid_of = {code_of[c]: c for pid in pid_of.values() for c, _ in kit_map.get(pid, [])
                   if c in code_of}
    for sku, (base, _) in kits.items():
        final.setdefault(base, base_pid_of[base])
    avg = {sku: tot.get(pid, 0) / n_days * 7 for sku, pid in final.items()}

    need = {base_pid_of[b] for b, _ in kits.values()}
    oh = pull_qty_available(od, need, label="基础品在手") if need else {}
    onhand = {code_of[pid]: q for pid, q in oh.items()}

    n_set = sum(1 for v in avg.values() if safety_of(v) is not None)
    info = (f"VO/GW 零售 {since:%m-%d}~{until:%m-%d}（{weeks} 周，套件按 BoM 折算）；"
            f"{len(pid_of)}/{len(skus)} 个 SKU 在 Odoo 找到，{len(kits)} 个套件并入基础品；"
            f"周均 ≥ {MIN_AVG} 设安全库存 {n_set} 个")
    if multi:
        info += f"；⚠ 多组件套件 {len(multi)} 个未并入: {', '.join(multi[:5])}"
    if dup:
        info += f"；⚠ {dup} 个 Internal Reference 在 Odoo 有重码，取最早建档的"
    return {"kits": kits, "avg": avg, "onhand": onhand, "info": info}
