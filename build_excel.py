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


def build(erp_path, tm_path, out_arg=None, outdir="output"):
    """阶段一核心：生成「拣货表+面单」workbook。返回 (输出路径, 统计字典)。
    供 CLI(main) 与 GUI 共用。"""
    os.makedirs(outdir, exist_ok=True)
    erp = s4.load_erp(erp_path)
    tmall = s4.load_tmall(tm_path)
    merged = s4.merge(erp, tmall)
    facesheet, cancel = s4.classify(merged)
    # 统一序号：面单与无货勾选共用同一编号，保证两表逐行严格对应(纸质面单↔表格定位)
    facesheet = facesheet.reset_index(drop=True)
    facesheet.insert(0, "序号", range(1, len(facesheet) + 1))

    pick_df = build_picking(facesheet)
    face_df = build_facesheet(facesheet)

    wb = Workbook()
    ws_pick = wb.active
    ws_pick.title = "拣货表"
    write_df(ws_pick, pick_df)
    style_sheet(ws_pick, len(pick_df.columns))
    apply_print(ws_pick, fit_width=True)       # 拣货单：所有列压一页宽

    ws_face = wb.create_sheet("面单")
    write_df(ws_face, face_df)
    style_sheet(ws_face, len(face_df.columns))
    highlight_facesheet(ws_face, face_df)
    merge_multiproduct(ws_face, face_df)
    apply_print(ws_face, landscape=True)        # 面单：横向

    chk_df = build_nogoods_helper(facesheet)
    ws_chk = wb.create_sheet("无货勾选")
    write_df(ws_chk, chk_df)
    style_sheet(ws_chk, len(chk_df.columns))  # 无合并，可直接筛选(数字工作页，不打印)

    out_path = out_arg or make_output_name(facesheet, outdir)
    wb.save(out_path)
    stats = {
        "sku": len(pick_df),
        "lines": len(face_df),
        "orders": int(face_df["Order Reference"].nunique()),
        "multi": int(face_df["Order Reference"].duplicated(keep=False).sum()),
        "highlight": sum(is_x2(v) for v in face_df["Internal Reference"])
        + int((pd.to_numeric(face_df["Quantity"], errors="coerce") > 1).sum())
        + int((face_df["VO Delivery Type"] == "CC").sum()),
    }
    return out_path, stats


def main():
    args = sys.argv[1:]
    erp_path = args[0] if len(args) > 0 else "raw_data/测试0611erp导出.xlsx"
    tm_path  = args[1] if len(args) > 1 else "raw_data/天猫测试.xlsx"
    out_arg  = args[2] if len(args) > 2 else None
    outdir   = os.path.dirname(out_arg) if out_arg else "output"
    out_path, st = build(erp_path, tm_path, out_arg, outdir)
    print(f"已生成: {out_path}")
    print(f"  拣货表: {st['sku']} 个 SKU")
    print(f"  面单:   {st['lines']} 行 / {st['orders']} 单")
    print(f"  多品订单行数(已合并 A/B/F): {st['multi']}")
    print(f"  标黄单元格数: {st['highlight']}")


if __name__ == "__main__":
    main()
