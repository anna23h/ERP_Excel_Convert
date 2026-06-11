#!/usr/bin/env python3
"""第二阶段：仓库反馈无货后，从"实际发货订单"生成 B/C/D。

实际发货 = 步骤4 的「待全集」 − 无货清单。

用法:
    python3 stage2.py --mmdd 0610 \
        [--erp raw_data/xxx.xlsx] [--tmall raw_data/天猫测试.xlsx] \
        [--nogoods raw_data/无货清单.xlsx] \
        [--billing raw_data/账单模板导出.xlsx] \
        [--outdir output]

- 无货清单：含 Order Reference 或 系统履约单号(15位) 任一列即可；不传则视为全部发货。
- 账单模板导出(D 用)：ERP 导出，需含 External ID(ID) + Order Reference + Terms and conditions(原值=渠道+运单)。
  不传则跳过 D。
"""
import argparse, os, re
from datetime import date
import pandas as pd
from openpyxl import Workbook

import step4_merge as s4
import build_excel as be  # 复用样式


def last15(series):
    return series.astype(str).str[-15:]


SCP_RE = re.compile(r"SCP\d+")   # 连接键: 系统履约单号(SCP…) / Order Reference(VO_TOF_SCP…)

MARK_PREFIXES = ("无货", "缺货", "勾选")
TRUTHY_STR = {"1", "x", "✓", "√", "是", "y", "yes", "true", "无货", "缺货"}


def _truthy(v):
    if isinstance(v, bool):
        return v
    if pd.isna(v):
        return False
    s = str(v).strip().lower()
    if s in TRUTHY_STR:
        return True
    try:                       # 数字: 非 0 即真(兼容 1/0)
        return float(s) != 0
    except ValueError:
        return False


NOGOODS_SHEET = "无货勾选"


def read_marked(paths):
    """读无货返回文件中被标记的行。多 sheet 时优先读『无货勾选』页；有标记列则只取标记为真的行。
    paths 支持多份做冗余(同店分多次导出)，合并。"""
    if isinstance(paths, str):
        paths = [paths]
    frames = []
    for path in paths:
        xl = pd.ExcelFile(path)
        sheet = NOGOODS_SHEET if NOGOODS_SHEET in xl.sheet_names else 0
        df = pd.read_excel(path, sheet_name=sheet)
        mark = next((c for c in df.columns
                     if str(c).strip().startswith(MARK_PREFIXES)), None)
        if mark is not None:
            df = df[df[mark].apply(_truthy)].copy()
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def load_nogoods(paths):
    """无货返回文件 → 无货连接键集合(SCP)。有『无货/缺货/勾选』标记列则只取标记为真的行，
    否则视全部行为无货。按值模式 `SCP\\d+` 识别单号(跳过标记列)，不依赖列名/sheet。"""
    if not paths:
        return set()
    df = read_marked(paths)
    keys = set()
    for c in df.columns:
        if str(c).strip().startswith(MARK_PREFIXES):
            continue
        for v in df[c].dropna().astype(str):
            m = SCP_RE.search(v)
            if m:
                keys.add(m.group(0)[-15:])
    return keys


def load_shipped_keys(paths):
    """读「有货(真实发货)订单清单」→ 连接键集合(SCP 单号)。

    入口高度模糊：不要求特定结构/sheet/列名。只要文件里含连接键(履约单号或 Order Reference,
    形如 `SCP…`)即可——按**值模式** `SCP\\d+` 扫描每个文件的**所有 sheet、所有列**，
    其余列(数量/名称/日期等)自动忽略。多份文件作冗余合并(同店分多次导出)。"""
    if not paths:
        return set()
    if isinstance(paths, str):
        paths = [paths]
    keys = set()
    for path in paths:
        for df in pd.read_excel(path, sheet_name=None).values():   # 所有 sheet
            for c in df.columns:
                for v in df[c].dropna().astype(str):
                    m = SCP_RE.search(v)
                    if m:
                        keys.add(m.group(0)[-15:])   # SCP 单号(后15位)，与 ERP _key 对齐
    return keys


def channel_of(order_ref):
    """VO_TOF_SCP... -> VO ; GW_TOF_SCP... -> GW"""
    return str(order_ref).split("_", 1)[0]


def get_shipped_orders(erp, shipped_keys):
    """有货(真实发货)单号 ∩ ERP → 订单级 DataFrame(Order Reference, VO Tracking No, _key, channel)。
    入口直接是真实发货订单号，结合 ERP 取明细，不再由无货倒推。"""
    orders = (erp[[s4.ERP_ORDER_REF, s4.ERP_TRACKING, "_key"]]
              .drop_duplicates(subset="_key"))
    shipped = orders[orders["_key"].isin(shipped_keys)].copy()
    shipped["channel"] = shipped[s4.ERP_ORDER_REF].map(channel_of)
    return shipped


# ---------- B: 系统履约单号 ----------
def build_B(shipped, outdir):
    out = pd.DataFrame({"系统履约单号": shipped["_key"].tolist()})
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    be.write_df(ws, out)
    be.style_sheet(ws, 1)
    path = os.path.join(outdir, "系统履约单号.xlsx")
    wb.save(path)
    return path, len(out)


# ---------- C: 发货表 (GW/VO 分 sheet) ----------
def build_C(shipped, outdir):
    wb = Workbook(); first = True
    counts = {}
    for ch in ["GW", "VO"]:
        sub = shipped[shipped["channel"] == ch]
        counts[ch] = len(sub)
        df = sub[[s4.ERP_ORDER_REF, s4.ERP_TRACKING]].rename(
            columns={s4.ERP_ORDER_REF: "Order Reference", s4.ERP_TRACKING: "VO Tracking No"})
        ws = wb.active if first else wb.create_sheet()
        ws.title = ch
        first = False
        be.write_df(ws, df)
        be.style_sheet(ws, 2)
    path = os.path.join(outdir, "发货表.xlsx")
    wb.save(path)
    return path, counts


# ---------- E: 出库单 (从 stock picking 全量导出过滤，按店铺拆 VO/GW) ----------
SCP_RE = re.compile(r"(SCP\d+)")
PICK_SRC_NAMES = ("Source Document", "Referenzbeleg")
PICK_TRK_NAMES = ("Tracking Reference", "Tracking-Referenz")


def _scp(v):
    m = SCP_RE.search(str(v))
    return m.group(1) if m else None


def build_E(picking_paths, shipped, shipdate, outdir):
    """从 stock picking 全量导出(出库原始数据)生成出库单。

    过滤: Source Document 的 SCP ∈ 实际发货订单。沿用 pool 的英文表头与原 Status，
    仅把 Tracking Reference 统一覆盖成发货日期。按 VO/GW 拆成两个文件。
    返回 (results{ch:(path,n)}, missing{ch:[scp...]})。"""
    results, missing = {}, {}
    if not picking_paths:
        return results, missing
    frames = [pd.read_excel(p) for p in picking_paths]
    pool = pd.concat(frames, ignore_index=True)
    srccol = next((c for c in pool.columns
                   if str(c).strip() in PICK_SRC_NAMES), None)
    if srccol is None:
        raise ValueError(f"出库原始数据缺少来源单据列(任一: {PICK_SRC_NAMES})")
    trkcol = next((c for c in pool.columns
                   if str(c).strip() in PICK_TRK_NAMES), None)

    pool["_scp"] = pool[srccol].map(_scp)
    pool["_ch"] = pool[srccol].astype(str).str.split("_", n=1).str[0]
    shipped_scp = {s for s in (_scp(k) for k in shipped["_key"]) if s}

    kept = pool[pool["_scp"].isin(shipped_scp)].copy()
    if trkcol is not None and shipdate:
        kept[trkcol] = shipdate

    pool_channels = set(pool["_ch"].unique())
    for ch in ["VO", "GW"]:
        sub = kept[kept["_ch"] == ch].drop(columns=["_scp", "_ch"])
        sub = sub.where(pd.notna(sub), "")   # 空单元格写空串，避免出现字面 nan
        if not sub.empty:
            wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
            be.write_df(ws, sub)
            be.style_sheet(ws, len(sub.columns))
            path = os.path.join(outdir, f"出库单{ch}.xlsx")
            wb.save(path)
            results[ch] = (path, len(sub))
        # 仅对 pool 覆盖到的店铺报缺(发货订单在 pool 里找不到对应 picking)
        if ch in pool_channels:
            want = {s for s in (_scp(k) for k in
                                shipped[shipped["channel"] == ch]["_key"]) if s}
            miss = sorted(want - set(pool.loc[pool["_ch"] == ch, "_scp"]))
            if miss:
                missing[ch] = miss
    return results, missing


# ---------- 缺货记录 (明细 + SKU 汇总) ----------
def _sku_lut(erp):
    """SKU -> 条码/货位/系统在售库存 查找表(SKU 级，从 ERP 导出取)。"""
    g = erp.groupby(s4.ERP_INTERNAL)
    lut = g.agg(
        Barcode=(s4.ERP_BARCODE, "first"),
        PickingName=(s4.ERP_PICKING, "first"),
        OnHand=(s4.ERP_ONHAND, "max"),
    )
    return lut


def _merge_same(ws, df, cols, start=2):
    """对已按某列排序的 df，把 cols 中各列相同连续值纵向合并(直观分组)。"""
    names = list(df.columns)
    idx = {c: names.index(c) + 1 for c in cols}
    keys = df[cols[0]].tolist()  # 以第一列(SKU)为分组依据
    i = 0
    while i < len(keys):
        j = i
        while j + 1 < len(keys) and keys[j + 1] == keys[i]:
            j += 1
        if j > i:
            for c in idx.values():
                ws.merge_cells(start_row=start + i, start_column=c,
                               end_row=start + j, end_column=c)
        i = j + 1


def build_shortage(marked, erp, mmdd, outdir):
    """从标记无货的行生成缺货记录：明细(按SKU合并) + SKU汇总。"""
    if marked.empty or "SKU" not in marked.columns:
        return None, 0
    lut = _sku_lut(erp)

    def info(sku, field, default=""):
        return lut.loc[sku, field] if sku in lut.index else default

    rows = []
    for _, r in marked.iterrows():
        sku = r["SKU"]
        rows.append({
            "系统履约单号":    r.get("系统履约单号", ""),
            "SKU":           sku,
            "商品名":         r.get("商品名", ""),
            "Barcode":       info(sku, "Barcode"),
            "Picking Name":  info(sku, "PickingName"),
            "缺货数量":       int(pd.to_numeric(r.get("数量", 1), errors="coerce") or 1),
            "系统在售库存":    info(sku, "OnHand", ""),
            "VO Delivery Type": r.get("VO Delivery Type", ""),
            "备注":           "",
        })
    detail = pd.DataFrame(rows).sort_values(
        ["SKU", "系统履约单号"]).reset_index(drop=True)
    detail.insert(0, "序号", range(1, len(detail) + 1))

    summary = (detail.groupby("SKU", sort=False)
               .agg(商品名=("商品名", "first"),
                    Barcode=("Barcode", "first"),
                    **{"Picking Name": ("Picking Name", "first")},
                    缺货订单数=("系统履约单号", "nunique"),
                    缺货总数量=("缺货数量", "sum"),
                    系统在售库存=("系统在售库存", "first"))
               .reset_index()
               .sort_values("缺货总数量", ascending=False))

    wb = Workbook()
    ws_d = wb.active
    ws_d.title = "明细"
    be.write_df(ws_d, detail)
    be.style_sheet(ws_d, len(detail.columns))
    # 合并相同 SKU 的 SKU 级列(直观看出哪个商品缺、缺几单)
    _merge_same(ws_d, detail, ["SKU", "商品名", "Barcode", "Picking Name", "系统在售库存"])

    ws_s = wb.create_sheet("SKU汇总")
    be.write_df(ws_s, summary)
    be.style_sheet(ws_s, len(summary.columns))

    path = os.path.join(outdir, "缺货记录.xlsx")
    wb.save(path)
    return path, len(detail), len(summary)


# ---------- D: 开账单上传表 ----------
TAG_RE = re.compile(r"^(账单|开发票)\d{4}")


find_id_col = s4.find_id_col   # 复用 step4_merge 的实现(External ID/ID/外部ID)


def has_id_col(df):
    return find_id_col(df) is not None


def build_billing(src, shipped_keys, mmdd, outdir):
    """从含 External ID 的来源(订单导出 或 账单模板导出)生成账单上传表。
    去重到订单级；Terms = 账单MMDD + 原值(剥离旧标签)；只留实际发货订单。
    保留来源的 ID 列原名(External ID / ID)，回传 Odoo 时按其匹配。"""
    df = src.copy()
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    idc = find_id_col(df)
    # 去重到订单级(订单导出是逐行的；first 跳过空值取订单头行的 ID/Date/Terms)
    agg = df.groupby("Order Reference", sort=False).agg(**{
        "Order Date":           ("Order Date", "first"),
        idc:                    (idc, "first"),
        "Terms and conditions": ("Terms and conditions", "first"),
    }).reset_index()
    agg["_key"] = last15(agg["Order Reference"])
    if shipped_keys:
        agg = agg[agg["_key"].isin(shipped_keys)]
    raw = (agg["Terms and conditions"].astype(str)
           .str.replace(TAG_RE, "", regex=True)
           .str.replace(s4.WU_TAG, "", regex=False))   # 已补运单：剥掉「无运单」恢复原值
    agg["Terms and conditions"] = f"账单{mmdd}" + raw
    out = agg[["Order Date", idc, "Order Reference", "Terms and conditions"]]
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    be.write_df(ws, out)
    be.style_sheet(ws, len(out.columns))
    path = os.path.join(outdir, "账单上传.xlsx")
    wb.save(path)
    return path, len(out)


def run(mmdd, erp_paths, shipped_files=None, nogoods_files=None, done_path=None,
        full_tmall_path=None, billing=None, outdir="output", picking=None, shipdate=None):
    """第二阶段核心：结合 ERP 生成 B/C/D/E(+缺货记录)。两种入口二选一(都给则优先有货)：

    - **有货清单**(shipped_files，白名单)：真实发货的履约单号/Order Reference，直接 ∩ ERP 取明细。
    - **无货勾选**(nogoods_files，黑名单)：仓库返回的无货单，发货 = 拣货候选(需 done) − 无货。
    done_path/full_tmall_path：无货入口算候选必需；有货入口仅用于缺货记录(可选)。
    返回结果文字行列表。供 CLI 与 GUI 共用。"""
    os.makedirs(outdir, exist_ok=True)
    log = []
    erp_df = be._load_erps(erp_paths)         # 只读一次，发货/账单/缺货复用
    erp_keys = set(erp_df["_key"])
    # 拣货候选(stage1 发货集合)：无货倒推 & 缺货记录用
    cand_keys = None
    if done_path:
        ann = s4.classify4(erp_df, s4.load_done_keys(done_path),
                           s4.load_cancel_keys(full_tmall_path))
        cand_keys = set(ann[ann["_ship"]]["_key"]) & erp_keys

    if shipped_files:                          # —— 有货入口(白名单) ——
        ship_in = load_shipped_keys(shipped_files)
        shipped_keys = ship_in & erp_keys
        log.append(f"[有货入口] 输入 {len(ship_in)} 单号，命中 ERP {len(shipped_keys)} 单")
        miss_erp = ship_in - erp_keys
        if miss_erp:
            log.append(f"⚠ {len(miss_erp)} 个有货单号在 ERP 里找不到(未结合): "
                       f"{', '.join(list(miss_erp)[:10])}{' …' if len(miss_erp) > 10 else ''}")
    elif nogoods_files:                        # —— 无货入口(黑名单) ——
        if cand_keys is None:
            log.append("✗ 无货入口需要『面单已完成名单』(--done)算拣货候选，未提供，已中止")
            return log
        ng = load_nogoods(nogoods_files)
        shipped_keys = cand_keys - ng
        log.append(f"[无货入口] 拣货候选 {len(cand_keys)} − 无货 {len(ng)} = 发货 {len(shipped_keys)} 单")
    else:
        log.append("✗ 未提供有货清单(shipped)或无货返回文件(nogoods)，无法确定发货集合")
        return log

    shipped = get_shipped_orders(erp_df, shipped_keys)
    shipped_keys = set(shipped["_key"])        # 规整(去重到订单级后的实际键)
    log.append(f"实际发货订单: {len(shipped)} 单 "
               f"(GW {sum(shipped['channel']=='GW')} / VO {sum(shipped['channel']=='VO')})")

    pB, nB = build_B(shipped, outdir)
    log.append(f"系统履约单号 已生成: {pB}  ({nB} 个)")

    pC, cC = build_C(shipped, outdir)
    log.append(f"发货表 已生成: {pC}  (GW {cC['GW']} / VO {cC['VO']})")

    # E 出库单：从 stock picking 全量导出过滤出实际发货订单
    if picking:
        sd = shipdate or date.today().strftime("%Y%m%d")
        resE, missE = build_E(picking, shipped, sd, outdir)
        if resE:
            for ch, (p, n) in resE.items():
                log.append(f"出库单{ch} 已生成: {p}  ({n} 行，发货日期 {sd})")
        else:
            log.append("出库单 跳过 (出库原始数据无匹配的发货订单)")
        for ch, miss in missE.items():
            log.append(f"⚠ 出库单{ch}: {len(miss)} 个发货订单在出库原始数据里找不到 picking: "
                       f"{', '.join(miss[:10])}{' …' if len(miss) > 10 else ''}")
    else:
        log.append("出库单 跳过 (未传出库原始数据 --picking)")

    # 账单上传：优先用含 ID 的订单导出直接生成；否则用单独的账单模板导出
    if has_id_col(erp_df):
        pD, nD = build_billing(erp_df, shipped_keys, mmdd, outdir)
        log.append(f"账单上传 已生成(来自订单导出含ID): {pD}  ({nD} 行，标签 账单{mmdd})")
    elif billing:
        pD, nD = build_billing(pd.read_excel(billing), shipped_keys, mmdd, outdir)
        log.append(f"账单上传 已生成(来自账单模板导出): {pD}  ({nD} 行，标签 账单{mmdd})")
    else:
        log.append("账单上传 跳过 (订单导出无ID列且未传账单模板；在 Odoo 订单导出模板勾上 External ID 列即可自动生成)")

    # 缺货记录
    if nogoods_files and not shipped_files:
        # 无货入口：用返回文件的 SKU 明细(SKU 级，多品订单可缺单件)
        res = build_shortage(read_marked(nogoods_files), erp_df, mmdd, outdir)
        if res[0]:
            log.append(f"缺货记录 已生成(SKU级,来自无货勾选): {res[0]}  (明细 {res[1]} 行 / SKU {res[2]} 种)")
        else:
            log.append("缺货记录 跳过 (无货返回文件无 SKU 明细)")
    elif cand_keys is not None:
        # 有货入口：整单缺货 = 拣货候选 − 有货，SKU 明细回查 ERP(订单级)
        missing = cand_keys - shipped_keys
        if missing:
            mr = erp_df[erp_df["_key"].isin(missing)]
            marked = pd.DataFrame({
                "系统履约单号":      mr["_key"].values,
                "SKU":              mr[s4.ERP_INTERNAL].values,
                "商品名":            mr[be.ERP_NAME].values,
                "数量":             mr[s4.ERP_QTY].values,
                "VO Delivery Type": mr[s4.ERP_DELIVERY].values,
            })
            res = build_shortage(marked, erp_df, mmdd, outdir)
            if res[0]:
                log.append(f"缺货记录 已生成(整单级,拣货候选−有货): {res[0]}  "
                           f"(明细 {res[1]} 行 / SKU {res[2]} 种)")
            else:
                log.append("缺货记录 跳过")
        else:
            log.append("缺货记录: 0 单(有货已覆盖全部拣货候选)")
    else:
        log.append("缺货记录 跳过 (有货入口未传 --done，无法算拣货候选)")
    return log


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mmdd", required=True, help="标签日期 MMDD，如 0610")
    ap.add_argument("--erp", nargs="+", required=True,
                    help="ERP 导出，可多份(VO/GW 各一份)")
    ap.add_argument("--shipped", nargs="*", default=None,
                    help="有货入口：真实发货订单清单(含履约单号/Order Reference 即可)，可多份冗余")
    ap.add_argument("--nogoods", nargs="*", default=None,
                    help="无货入口：仓库返回的无货勾选文件，可多份冗余(需配 --done 算拣货候选)")
    ap.add_argument("--done", default=None,
                    help="面单已完成名单：无货入口算拣货候选必需；有货入口仅用于缺货记录")
    ap.add_argument("--tmall-full", dest="full", default=None,
                    help="完整天猫导出(含取消状态)，配 --done 精确算拣货候选")
    ap.add_argument("--billing", default=None)
    ap.add_argument("--picking", nargs="*", default=None,
                    help="出库原始数据(stock picking 全量导出)，可传多个(VO/GW 各一份)")
    ap.add_argument("--shipdate", default=None, help="发货日期 YYYYMMDD，默认今天")
    ap.add_argument("--outdir", default="output")
    a = ap.parse_args()
    for line in run(a.mmdd, a.erp, a.shipped, a.nogoods, a.done, a.full, a.billing,
                    a.outdir, a.picking, a.shipdate):
        print(line)


if __name__ == "__main__":
    main()
