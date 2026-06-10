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
import pandas as pd
from openpyxl import Workbook

import step4_merge as s4
import build_excel as be  # 复用样式


def last15(series):
    return series.astype(str).str[-15:]


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


def read_marked(path):
    """读返回文件中被标记无货的行。多 sheet 时优先读『无货勾选』页。"""
    xl = pd.ExcelFile(path)
    sheet = NOGOODS_SHEET if NOGOODS_SHEET in xl.sheet_names else 0
    df = pd.read_excel(path, sheet_name=sheet)
    mark = next((c for c in df.columns
                 if str(c).strip().startswith(MARK_PREFIXES)), None)
    if mark is not None:
        df = df[df[mark].apply(_truthy)].copy()
    return df


def load_nogoods(path):
    """返回无货的 15 位单号集合。
    - 取标记为真(1/✓/x...)的行；取号优先用 系统履约单号 / Order Reference 列。"""
    if not path:
        return set()
    df = read_marked(path)
    preferred = [c for c in df.columns
                 if str(c).strip() in ("系统履约单号", "Order Reference")]
    cols = preferred if preferred else [c for c in df.columns
                                        if not str(c).strip().startswith(MARK_PREFIXES)]
    keys = set()
    for col in cols:
        vals = df[col].dropna().astype(str)
        keys |= set(vals.str[-15:])  # 后15位兼容 Order Reference 和纯单号
    return {k for k in keys if k and k != "nan"}


def channel_of(order_ref):
    """VO_TOF_SCP... -> VO ; GW_TOF_SCP... -> GW"""
    return str(order_ref).split("_", 1)[0]


def get_shipped_orders(erp_path, tmall_path, nogoods_keys):
    """返回实际发货订单(去重到单)的 DataFrame: Order Reference, VO Tracking No, _key, channel"""
    erp = s4.load_erp(erp_path)
    tmall = s4.load_tmall(tmall_path)
    merged = s4.merge(erp, tmall)
    facesheet, _ = s4.classify(merged)  # 待全集
    # 去重到订单级
    orders = (facesheet[[s4.ERP_ORDER_REF, s4.ERP_TRACKING, "_key"]]
              .drop_duplicates(subset="_key"))
    shipped = orders[~orders["_key"].isin(nogoods_keys)].copy()
    shipped["channel"] = shipped[s4.ERP_ORDER_REF].map(channel_of)
    return shipped


# ---------- B: 系统履约单号 ----------
def build_B(shipped, outdir):
    out = pd.DataFrame({"系统履约单号": shipped["_key"].tolist()})
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    be.write_df(ws, out)
    be.style_sheet(ws, 1)
    path = os.path.join(outdir, "B_系统履约单号.xlsx")
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
    path = os.path.join(outdir, "C_发货表.xlsx")
    wb.save(path)
    return path, counts


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


def build_D(billing_path, shipped_keys, mmdd, outdir):
    df = pd.read_excel(billing_path)
    # 去掉空列(Unnamed / 全空)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df = df.dropna(axis=1, how="all")
    # 只留实际发货订单
    df["_key"] = last15(df["Order Reference"])
    if shipped_keys:
        df = df[df["_key"].isin(shipped_keys)]
    # Terms = 账单MMDD + 原值；若已带旧标签先剥离
    raw = df["Terms and conditions"].astype(str).str.replace(TAG_RE, "", regex=True)
    df["Terms and conditions"] = f"账单{mmdd}" + raw
    cols = [c for c in ["Order Date", "ID", "Order Reference", "Terms and conditions"]
            if c in df.columns]
    out = df[cols]
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    be.write_df(ws, out)
    be.style_sheet(ws, len(cols))
    path = os.path.join(outdir, "D_账单上传.xlsx")
    wb.save(path)
    return path, len(out)


def run(mmdd, erp_path, tmall_path, nogoods=None, billing=None, outdir="output"):
    """第二阶段核心：生成 B/C/D + 缺货记录。返回结果文字行列表。供 CLI 与 GUI 共用。"""
    os.makedirs(outdir, exist_ok=True)
    log = []
    ng = load_nogoods(nogoods)
    shipped = get_shipped_orders(erp_path, tmall_path, ng)
    shipped_keys = set(shipped["_key"])
    log.append(f"无货清单: {len(ng)} 单")
    log.append(f"实际发货订单: {len(shipped)} 单 "
               f"(GW {sum(shipped['channel']=='GW')} / VO {sum(shipped['channel']=='VO')})")

    pB, nB = build_B(shipped, outdir)
    log.append(f"B 已生成: {pB}  ({nB} 个系统履约单号)")

    pC, cC = build_C(shipped, outdir)
    log.append(f"C 已生成: {pC}  (GW {cC['GW']} / VO {cC['VO']})")

    if billing:
        pD, nD = build_D(billing, shipped_keys, mmdd, outdir)
        log.append(f"D 已生成: {pD}  ({nD} 行，标签 账单{mmdd})")
    else:
        log.append("D 跳过 (未传账单模板导出)")

    if nogoods:
        erp = s4.load_erp(erp_path)
        res = build_shortage(read_marked(nogoods), erp, mmdd, outdir)
        if res[0]:
            log.append(f"缺货记录 已生成: {res[0]}  (明细 {res[1]} 行 / SKU {res[2]} 种)")
        else:
            log.append("缺货记录 跳过 (返回文件无 SKU 明细)")
    return log


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mmdd", required=True, help="标签日期 MMDD，如 0610")
    ap.add_argument("--erp", default="raw_data/测试0611erp导出.xlsx")
    ap.add_argument("--tmall", default="raw_data/天猫测试.xlsx")
    ap.add_argument("--nogoods", default=None)
    ap.add_argument("--billing", default=None)
    ap.add_argument("--outdir", default="output")
    a = ap.parse_args()
    for line in run(a.mmdd, a.erp, a.tmall, a.nogoods, a.billing, a.outdir):
        print(line)


if __name__ == "__main__":
    main()
