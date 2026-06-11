#!/usr/bin/env python3
"""在 step4_merge 分流基础上生成「拣货表+面单」workbook (输出表 A)。

用法:
    python3 build_excel.py [erp] [tmall] [out.xlsx]
默认输出: output/拣货表+面单.xlsx
"""
import sys, os, re
from datetime import date
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

import step4_merge as s4

YELLOW = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
FONT = Font(size=15)
HEAD_FONT = Font(size=15, bold=True)
ROW_H = 40

# 面单列（顺序即 A..F），末尾加空白「仓库备注」
FACE_COLS = [
    (s4.ERP_ORDER_REF, "Order Reference"),
    (s4.ERP_TRACKING,  "VO Tracking No"),
    (s4.ERP_INTERNAL,  "Internal Reference"),
    (s4.ERP_PICKING,   "Picking Name"),
    (s4.ERP_QTY,       "Quantity"),
    (s4.ERP_DELIVERY,  "VO Delivery Type"),
]
WAREHOUSE_NOTE = "仓库备注"


def build_picking(facesheet):
    """扁平单层：一行一个 SKU。"""
    g = facesheet.groupby(s4.ERP_INTERNAL, sort=False)
    out = g.agg(**{
        "Picking Name":     (s4.ERP_PICKING,  "first"),
        "Barcode":          (s4.ERP_BARCODE,  "first"),
        "Quantity":         (s4.ERP_QTY,      "sum"),
        "Quantity On Hand": (s4.ERP_ONHAND,   "max"),
    }).reset_index().rename(columns={s4.ERP_INTERNAL: "Internal Reference"})
    # 数量统一整数，消除 csv(1.0) 与 xlsx(1) 的显示差异
    for c in ("Quantity", "Quantity On Hand"):
        out[c] = pd.to_numeric(out[c], errors="coerce").round().astype("Int64")
    return out


ERP_NAME = "Order Lines/Product/Name"


def build_nogoods_helper(facesheet):
    """逐商品行的无货勾选表(方案B)：一行一个 SKU，无合并、无空值，可直接筛选。
    首列『无货』为布尔复选框(默认 False)。单品订单一行一勾；多品订单缺某件只勾那一行，
    被勾行的 SKU 天然即缺货记录(零手抄)。stage2 任一行勾→整单扣下。"""
    df = facesheet.copy()
    out = pd.DataFrame({
        "序号":             df["序号"].values,  # 与面单同一编号，逐行对应便于定位
        "无货(1=缺货)":      0,  # 数字 0/1，全 Excel 版本通用；标无货填 1
        "Order Reference":  df[s4.ERP_ORDER_REF].values,
        "系统履约单号":      df["_key"].values,
        "SKU":              df[s4.ERP_INTERNAL].values,
        "商品名":            df[ERP_NAME].values,
        "数量":             pd.to_numeric(df[s4.ERP_QTY], errors="coerce")
                                .round().astype("Int64").values,
        "VO Delivery Type": df[s4.ERP_DELIVERY].values,
    })
    return out


def build_facesheet(facesheet):
    df = facesheet[[c for c, _ in FACE_COLS]].copy()
    df.columns = [h for _, h in FACE_COLS]
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").round().astype("Int64")
    df.insert(0, "序号", facesheet["序号"].values)  # 与无货勾选同一编号
    df[WAREHOUSE_NOTE] = ""
    return df


def is_x2(v):
    return bool(re.search(r"x2$", str(v), re.I))


def style_sheet(ws, n_cols, header_font=HEAD_FONT):
    for row in ws.iter_rows():
        for cell in row:
            cell.border = BORDER
            cell.alignment = CENTER
            cell.font = FONT
    for cell in ws[1]:
        cell.font = header_font
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].height = ROW_H
    # 按内容自动算列宽（字号15 比默认大，需放大系数，否则日期显示为 ######）
    for c in range(1, n_cols + 1):
        maxlen = 0
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            s = v.strftime("%Y-%m-%d %H:%M:%S") if hasattr(v, "strftime") else str(v)
            maxlen = max(maxlen, len(s))
        ws.column_dimensions[get_column_letter(c)].width = max(12, maxlen * 1.5 + 2)


def write_df(ws, df):
    ws.append(list(df.columns))
    for _, row in df.iterrows():
        ws.append(list(row))


def highlight_facesheet(ws, df):
    """三条规则，标黄触发的单元格（让仓库知道原因）。"""
    col = {name: i + 1 for i, name in enumerate(df.columns)}
    for ridx, (_, row) in enumerate(df.iterrows(), start=2):
        if is_x2(row["Internal Reference"]):
            ws.cell(ridx, col["Internal Reference"]).fill = YELLOW
        if pd.notna(row["Quantity"]) and float(row["Quantity"]) > 1:
            ws.cell(ridx, col["Quantity"]).fill = YELLOW
        if str(row["VO Delivery Type"]) == "CC":
            ws.cell(ridx, col["VO Delivery Type"]).fill = YELLOW


def merge_multiproduct(ws, df):
    """多品订单：订单级列(Order Reference / VO Tracking No / VO Delivery Type)纵向合并。
    按列名定位，避免加序号列后错位。"""
    cols = list(df.columns)
    merge_cols = {n: cols.index(n) + 1
                  for n in ("Order Reference", "VO Tracking No", "VO Delivery Type")}
    start = 2
    refs = df["Order Reference"].tolist()
    i = 0
    while i < len(refs):
        j = i
        while j + 1 < len(refs) and refs[j + 1] == refs[i]:
            j += 1
        if j > i:  # 多行同单
            for c in merge_cols.values():
                ws.merge_cells(start_row=start + i, start_column=c,
                               end_row=start + j, end_column=c)
        i = j + 1


def make_output_name(facesheet, outdir):
    """步骤9 命名：YYYY年MM月DD日{渠道}{n}单 拣货表+面单.xlsx"""
    n = facesheet[s4.ERP_ORDER_REF].nunique()
    chans = sorted({str(r).split("_", 1)[0] for r in facesheet[s4.ERP_ORDER_REF].dropna()})
    ch = "+".join(chans)
    d = date.today()
    fname = f"{d.year}年{d.month:02d}月{d.day:02d}日{ch}{n}单 拣货表+面单.xlsx"
    return os.path.join(outdir, fname)


def apply_print(ws, landscape=False, fit_width=False, footer="第 &P 页，共 &N 页",
                top=0.9, bottom=0.9, left=0.8, right=0.8):
    """步骤9 打印设置。footer 用 Excel 字段码：&P=当前页码，&N=总页数。
    默认『第 &P 页，共 &N 页』(第1页，共3页...)。页边距单位为英寸。"""
    if landscape:
        ws.page_setup.orientation = "landscape"
    if fit_width:  # 所有列压到一页宽
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins = PageMargins(top=top, bottom=bottom, left=left, right=right,
                                  header=0.3, footer=0.3)
    ws.oddFooter.center.text = footer


def _load_erps(erp_paths):
    """接受单个路径或路径列表(VO/GW 各一份)，concat 成一张 ERP 行级表。"""
    if isinstance(erp_paths, str):
        erp_paths = [erp_paths]
    return pd.concat([s4.load_erp(p) for p in erp_paths], ignore_index=True)


def _order_level(df, idcol):
    """ERP 行级 → 订单级(多品合一行，取订单头行)。"""
    return df.groupby(s4.ERP_ORDER_REF, sort=False).agg(**{
        "Order Date":           ("Order Date", "first"),
        idcol:                  (idcol, "first"),
        "Terms and conditions": ("Terms and conditions", "first"),
    }).reset_index()


def build_upload(orders, idcol, terms, outdir, fname):
    """通用 ERP 上传表：Order Date | <idcol> | Order Reference | Terms。orders 为订单级。"""
    out = pd.DataFrame({
        "Order Date":           orders["Order Date"].values,
        idcol:                  orders[idcol].values,
        "Order Reference":      orders[s4.ERP_ORDER_REF].values,
        "Terms and conditions": terms,
    })
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    write_df(ws, out)
    style_sheet(ws, len(out.columns))
    path = os.path.join(outdir, fname)
    wb.save(path)
    return path, len(out)


def _write_pickface(facesheet, outdir, out_arg=None):
    """把一个店铺的发货集合写成一个「拣货表+面单+无货勾选」workbook。
    facesheet 需已含 `序号`。返回 (路径, SKU数, 订单数)。"""
    pick_df = build_picking(facesheet)
    face_df = build_facesheet(facesheet)
    wb = Workbook()
    ws_pick = wb.active; ws_pick.title = "拣货表"
    write_df(ws_pick, pick_df); style_sheet(ws_pick, len(pick_df.columns))
    apply_print(ws_pick, fit_width=True)
    ws_face = wb.create_sheet("面单")
    write_df(ws_face, face_df); style_sheet(ws_face, len(face_df.columns))
    highlight_facesheet(ws_face, face_df)
    merge_multiproduct(ws_face, face_df)
    apply_print(ws_face, landscape=True)
    chk_df = build_nogoods_helper(facesheet)
    ws_chk = wb.create_sheet("无货勾选")
    write_df(ws_chk, chk_df); style_sheet(ws_chk, len(chk_df.columns))
    path = out_arg or make_output_name(facesheet, outdir)
    wb.save(path)
    return path, len(pick_df), int(face_df["Order Reference"].nunique())


def build(erp_paths, done_path, full_tmall_path=None, out_arg=None, outdir="output"):
    """阶段一核心(步骤4+7/8/9)：分流 + 生成交付。返回 (log行列表, stats)。

    - 拣货表+面单：只含「发货」订单(已剔除无运单/取消)，**按店铺(VO/GW)各出一份**。
    - 取消单 / 无运单清单：ERP 上传表(External ID 匹配键)，**两店合并一份**。
    - 已补运单清单：系统履约单号(去天猫后台打面单)，**两店合并一份**。
    供 CLI(main) 与 GUI 共用。"""
    os.makedirs(outdir, exist_ok=True)
    log = []
    erp = _load_erps(erp_paths)
    done = s4.load_done_keys(done_path)
    status_map = s4.load_status_map(full_tmall_path)   # 完整天猫: 单号→履约单状态(选填)
    cancel_keys = set(status_map[status_map.isin(s4.CANCEL_STATUSES)].index)
    ann = s4.classify4(erp, done, cancel_keys)
    idcol = s4.find_id_col(ann)

    o = ann.drop_duplicates("_key")
    log.append("分流(订单级): " + " / ".join(
        f"{k} {v}" for k, v in o["_cat"].value_counts().items()))

    # ---- 拣货表 + 面单 (发货集合，按店铺 VO/GW 各出一份) ----
    facesheet = ann[ann["_ship"]].copy()
    main_paths = []
    if facesheet.empty:
        log.append("⚠ 无发货订单(全部无运单/取消)，未生成拣货表+面单")
    else:
        facesheet["_ch"] = (facesheet[s4.ERP_ORDER_REF].astype(str)
                            .str.split("_", n=1).str[0])
        chans = sorted(facesheet["_ch"].unique())
        for ch in chans:
            sub = (facesheet[facesheet["_ch"] == ch]
                   .drop(columns="_ch").reset_index(drop=True))
            sub.insert(0, "序号", range(1, len(sub) + 1))   # 序号按店内独立编号
            p, nsku, nord = _write_pickface(
                sub, outdir, out_arg if len(chans) == 1 else None)
            main_paths.append(p)
            log.append(f"拣货表+面单[{ch}] 已生成: {p}  ({nsku} SKU / {nord} 单)")

    # ---- 取消单 (Terms=平台订单取消, 上传 ERP 作废) ----
    d = date.today()
    cancel_df = ann[ann["_cat"] == "取消"]
    if not cancel_df.empty and idcol:
        orders = _order_level(cancel_df, idcol)
        tag = f"{d.year}年{d.month:02d}月{d.day:02d}日平台订单取消"
        p, n = build_upload(orders, idcol, [tag] * len(orders), outdir, "取消单.xlsx")
        log.append(f"取消单 已生成: {p}  ({n} 单)")
    elif not cancel_df.empty:
        log.append("⚠ 有取消单但 ERP 无 External ID 列，无法生成上传表(请在订单导出勾上 External ID)")
    else:
        log.append("取消单: 0 单" + ("" if full_tmall_path else " (未传完整天猫导出，无法识别取消)"))

    # ---- 无运单清单 (Terms='无运单'+原值, 上传 ERP) ----
    wu_df = ann[ann["_cat"] == "无运单"]
    if not wu_df.empty and idcol:
        orders = _order_level(wu_df, idcol)
        raw = orders["Terms and conditions"].astype(str)
        terms = raw.where(raw.str.contains(s4.WU_TAG), s4.WU_TAG + raw)  # 已带则不重复加
        p, n = build_upload(orders, idcol, terms.values, outdir, "无运单清单.xlsx")
        log.append(f"无运单清单 已生成: {p}  ({n} 单)")
    elif not wu_df.empty:
        log.append("⚠ 有无运单订单但 ERP 无 External ID 列，无法生成上传表")
    else:
        log.append("无运单清单: 0 单")

    # ---- 已补运单清单 (系统履约单号, 去天猫后台打面单；按店各一份: 分批次下载运单防混淆) ----
    refill = ann[ann["_cat"] == "已补运单"].drop_duplicates("_key").copy()
    if not refill.empty:
        refill["_ch"] = (refill[s4.ERP_ORDER_REF].astype(str)
                         .str.split("_", n=1).str[0])
        for ch in sorted(refill["_ch"].unique()):
            keys = refill[refill["_ch"] == ch]["_key"].tolist()
            out = pd.DataFrame({"系统履约单号": keys})
            wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
            write_df(ws, out); style_sheet(ws, 1)
            p = os.path.join(outdir, f"已补运单清单{ch}.xlsx"); wb.save(p)
            log.append(f"已补运单清单{ch} 已生成: {p}  ({len(out)} 单)")
    else:
        log.append("已补运单清单: 0 单")

    # ---- 异常上报(不静默) ----
    dup = erp.drop_duplicates(s4.ERP_ORDER_REF)["_key"].duplicated().sum()
    if dup:
        log.append(f"⚠ 连接键冲突: {dup} 个订单的 Order Reference 后15位与他单相同(可能误判状态)")
    # 护栏：发货名单反查完整天猫——若其实是 已发货/已收货/发货后取消，说明名单过期(防重复发货)
    if len(status_map) and not facesheet.empty:
        ship_keys = set(facesheet["_key"])
        bad = status_map[status_map.index.isin(ship_keys)
                         & status_map.isin(s4.SHIPPED_DONE_STATUSES)]
        if len(bad):
            log.append(f"⚠ 发货名单异常: {len(bad)} 单在完整天猫里其实是 "
                       f"已发货/已收货/发货后取消(名单可能过期，当心重复发货): "
                       f"{', '.join(list(bad.index)[:8])}{' …' if len(bad) > 8 else ''}")
    if not facesheet.empty:
        empty_dt = facesheet.drop_duplicates("_key")[s4.ERP_DELIVERY].isna().sum()
        if empty_dt:
            log.append(f"⚠ 发货订单中 VO Delivery Type 为空: {empty_dt} 单")

    stats = {
        "ship": int(o["_ship"].sum()),
        "cancel": int((o["_cat"] == "取消").sum()),
        "nowaybill": int((o["_cat"] == "无运单").sum()),
        "refill": int((o["_cat"] == "已补运单").sum()),
        "main_paths": main_paths,
    }
    return log, stats


def main():
    args = sys.argv[1:]
    if len(args) < 2:
        print("用法: python3 build_excel.py <erp[,erp2...]> <面单已完成名单> [完整天猫导出] [out.xlsx]")
        return
    erp_paths = args[0].split(",")
    done_path = args[1]
    full_tmall = args[2] if len(args) > 2 else None
    out_arg = args[3] if len(args) > 3 else None
    outdir = os.path.dirname(out_arg) if out_arg else "output"
    log, st = build(erp_paths, done_path, full_tmall, out_arg, outdir)
    for ln in log:
        print(ln)


if __name__ == "__main__":
    main()
