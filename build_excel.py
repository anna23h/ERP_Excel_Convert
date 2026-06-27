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
LEFT_BOTTOM = Alignment(horizontal="left", vertical="bottom", wrap_text=True)  # 左对齐+下沉
LEFT_CENTER = Alignment(horizontal="left", vertical="center", wrap_text=True)  # 合并单元格用
FONT = Font(size=15)
SMALL_FONT = Font(size=13)  # 比正文小 2 号
HEAD_FONT = Font(size=15, bold=True)
ROW_H = 35

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
    # 按 Internal Reference 升序，便于捡货员按货号顺序拣货(扁平表行间独立，排序不影响数据)
    out = out.sort_values("Internal Reference", key=lambda s: s.astype(str).str.lower(),
                          kind="stable").reset_index(drop=True)
    return out


ERP_NAME = "Order Lines/Product/Name"
# 订货预判用字段(员工在 ERP 订单导出时勾上这 3 个产品级列；缺则跳过订货清单)
ERP_FS     = "Order Lines/Product/FS"            # 供应商(去谁家订)
ERP_SAFETY = "Order Lines/Product/Safety Stock"  # 安全库存
ERP_REMARK = "Order Lines/Product/Supply Remark" # 备注(可能过期，仅背景)
REORDER_FLAG_PAT = re.compile(r"暂不采购|MHD|效期")  # 备注红旗：别自动下单，核实/问老板


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


def _first_nonempty(s):
    """组内第一个非空值(跳过 NaN/空串)，全空返回空串。"""
    for x in s:
        if pd.notna(x) and str(x).strip():
            return x
    return ""


def build_reorder(erp):
    """补货预判清单(Solo 作战清单·模式一 step 0 盘前预判)：
    把整份 ERP 订单导出按 Internal Reference 聚合，Quantity 求和=今日需求，与 On Hand 比。
    全部 SKU、按缺口(需求−在售)降序。需 ERP 导出含 FS/Safety Stock/Supply Remark 三列，缺则返回 None。"""
    if not all(c in erp.columns for c in (ERP_FS, ERP_SAFETY, ERP_REMARK)):
        return None
    g = erp.groupby(s4.ERP_INTERNAL, sort=False).agg(**{
        "产品名":        (ERP_NAME,       _first_nonempty),
        "今日需求":      (s4.ERP_QTY,     "sum"),
        "在售(On Hand)": (s4.ERP_ONHAND,  "max"),
        "安全库存":      (ERP_SAFETY,     "max"),
        "供应商(FS)":    (ERP_FS,         _first_nonempty),
        "备注":          (ERP_REMARK,     _first_nonempty),
    }).reset_index().rename(columns={s4.ERP_INTERNAL: "SKU"})
    for c in ("今日需求", "在售(On Hand)", "安全库存"):
        g[c] = pd.to_numeric(g[c], errors="coerce").round().astype("Int64")
    g["缺口"] = g["今日需求"] - g["在售(On Hand)"]
    # 红旗：备注含 暂不采购/MHD/效期 → 别自动下单，核实或问老板
    g["红旗"] = g["备注"].astype(str).apply(
        lambda v: "⚠核实" if REORDER_FLAG_PAT.search(v) else "")
    g = g.sort_values("缺口", ascending=False, kind="stable").reset_index(drop=True)
    return g[["产品名", "SKU", "今日需求", "在售(On Hand)", "缺口",
              "安全库存", "供应商(FS)", "红旗", "备注"]]


def build_facesheet(facesheet):
    df = facesheet[[c for c, _ in FACE_COLS]].copy()
    df.columns = [h for _, h in FACE_COLS]
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce").round().astype("Int64")
    df.insert(0, "序号", facesheet["序号"].values)  # 与无货勾选同一编号
    df[WAREHOUSE_NOTE] = ""
    return df


def is_multipack(v):
    # x 后跟数字结尾 = 多件装（x2/x3/x4/x10…），标色提醒打包员
    return bool(re.search(r"x\d+$", str(v), re.I))


def style_sheet(ws, n_cols, header_font=HEAD_FONT, left_cols=(), small_cols=()):
    """left_cols: 内容左对齐+下沉的列名集合；small_cols: 字号小2号的列名集合。表头始终居中。"""
    headers = [c.value for c in ws[1]]
    left_idx = {i + 1 for i, h in enumerate(headers) if h in left_cols}
    small_idx = {i + 1 for i, h in enumerate(headers) if h in small_cols}
    for row in ws.iter_rows():
        for cell in row:
            cell.border = BORDER
            if cell.row == 1:
                cell.alignment = CENTER
                cell.font = header_font
            else:
                cell.alignment = LEFT_BOTTOM if cell.column in left_idx else CENTER
                cell.font = SMALL_FONT if cell.column in small_idx else FONT
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
        if is_multipack(row["Internal Reference"]):
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


def fix_merged_alignment(ws, left_cols):
    """合并后调用：左对齐列若为合并单元格，首格改为 左对齐+垂直居中（非下沉）。"""
    headers = [c.value for c in ws[1]]
    left_idx = {i + 1 for i, h in enumerate(headers) if h in left_cols}
    for rng in ws.merged_cells.ranges:
        if rng.min_col in left_idx:
            ws.cell(rng.min_row, rng.min_col).alignment = LEFT_CENTER


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


def _wu_add(series):
    """无运单：Terms 前缀『无运单』(已带则不重复加)。"""
    raw = series.astype(str)
    return raw.where(raw.str.contains(s4.WU_TAG), s4.WU_TAG + raw).values


def _wu_strip(series):
    """已补运单：剥掉『无运单』恢复原值。"""
    return series.astype(str).str.replace(s4.WU_TAG, "", regex=False).str.strip().values


def unique_path(path):
    """目标已存在则在扩展名前加序号 (1)/(2)…，避免覆盖既有产出。"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base}({i}){ext}"):
        i += 1
    return f"{base}({i}){ext}"


def _write_simple(out, outdir, fname, n_cols=None):
    """把一张 DataFrame 写成单 sheet workbook(统一样式)。返回 (路径, 行数)。"""
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    write_df(ws, out)
    style_sheet(ws, n_cols or len(out.columns))
    path = unique_path(os.path.join(outdir, fname))
    wb.save(path)
    return path, len(out)


def _upload_terms(cat_df, idcol, terms):
    """某一类(取消/无运单/已补运单)的 ERP 上传行：订单级 4 列，Terms 取传入值/series。"""
    orders = _order_level(cat_df, idcol)
    out = pd.DataFrame({
        "Order Date":           orders["Order Date"].values,
        idcol:                  orders[idcol].values,
        "Order Reference":      orders[s4.ERP_ORDER_REF].values,
        "Terms and conditions": terms(orders) if callable(terms) else terms,
    })
    return out


def _write_pickface(facesheet, outdir, out_arg=None):
    """把一个店铺的发货集合写成一个「拣货表+面单+无货勾选」workbook。
    facesheet 需已含 `序号`。返回 (路径, SKU数, 订单数)。"""
    pick_df = build_picking(facesheet)
    face_df = build_facesheet(facesheet)
    wb = Workbook()
    ws_pick = wb.active; ws_pick.title = "拣货表"
    write_df(ws_pick, pick_df)
    style_sheet(ws_pick, len(pick_df.columns),
                left_cols={"Internal Reference", "Picking Name", "Barcode"},
                small_cols={"Picking Name"})
    apply_print(ws_pick, fit_width=True)
    ws_face = wb.create_sheet("面单")
    write_df(ws_face, face_df)
    face_left = {"Order Reference", "VO Tracking No", "Internal Reference", "Picking Name"}
    style_sheet(ws_face, len(face_df.columns),
                left_cols=face_left,
                small_cols={"Internal Reference", "Picking Name"})
    highlight_facesheet(ws_face, face_df)
    merge_multiproduct(ws_face, face_df)
    fix_merged_alignment(ws_face, face_left)
    apply_print(ws_face, landscape=True)
    chk_df = build_nogoods_helper(facesheet)
    ws_chk = wb.create_sheet("无货勾选")
    write_df(ws_chk, chk_df); style_sheet(ws_chk, len(chk_df.columns))
    path = unique_path(out_arg or make_output_name(facesheet, outdir))
    wb.save(path)
    return path, len(pick_df), int(face_df["Order Reference"].nunique())


def build(erp_paths, full_tmall_path, out_arg=None, outdir="output"):
    """阶段一核心(步骤4+7/8/9)：分流 + 生成交付。返回 (log行列表, stats)。

    输入单店 ERP（天猫两店混合，经 ∩ERP 收敛到单店）+ 一份完整天猫导出，产出全部按店带后缀：
    - 新订单获单清单：履约单状态=新订单 ∩ ERP 的系统履约单号(去天猫批量获单)。
    - 拣货表+面单：只含「发货」订单(已剔除无运单/取消)。
    - 回传ERP销售上传表：取消/无运单/已补运单三类 Terms 写回**合并一张**(External ID 匹配键)。
    - 已补运单清单：系统履约单号(去天猫后台打面单)。
    完整天猫导出是唯一天猫输入：发货范围由二段式(履约∈{新订单,商家已接单}∧面单已完成)推出，
    负集再按履约状态拆 取消/无运单。供 CLI(main) 与 GUI 共用。"""
    os.makedirs(outdir, exist_ok=True)
    log = []
    erp = _load_erps(erp_paths)
    full = s4.load_full_tmall(full_tmall_path)         # 唯一天猫输入
    done = s4.done_keys_from_full(full)                # 发货范围(二段式)
    status_map = full[s4.TM_STATUS] if not full.empty else pd.Series(dtype=object)
    cancel_keys = s4.cancel_keys_from_full(full)
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

    # ---- 今日预计发货总获单清单 (发货集的 系统履约单号 单列；含新订单+商家已接单∧面单已完成) ----
    # 拿去天猫后台批量获单：覆盖今日要发的全部单(新订单+已接单)，不止新订单那一部分。
    if not facesheet.empty:
        for ch in sorted(facesheet["_ch"].unique()):
            keys = facesheet[facesheet["_ch"] == ch].drop_duplicates("_key")["_key"].tolist()
            p, n = _write_simple(pd.DataFrame({"系统履约单号": keys}),
                                 outdir, f"今日预计发货总获单清单{ch}.xlsx", n_cols=1)
            log.append(f"今日预计发货总获单清单{ch} 已生成: {p}  ({n} 单)")

    # ---- 新订单获单清单 (履约单状态=新订单 ∩ ERP；复制履约单号去天猫批量获单；按店各一份) ----
    d = date.today()
    if len(status_map):
        new_keys = set(status_map[status_map == s4.NEW_ORDER_STATUS].index)
        no = ann.drop_duplicates("_key")
        no = no[no["_key"].isin(new_keys)].copy()
        if not no.empty:
            no["_ch"] = no[s4.ERP_ORDER_REF].astype(str).str.split("_", n=1).str[0]
            for ch in sorted(no["_ch"].unique()):
                keys = no[no["_ch"] == ch]["_key"].tolist()
                p, n = _write_simple(pd.DataFrame({"系统履约单号": keys}),
                                     outdir, f"新订单获单清单{ch}.xlsx", n_cols=1)
                log.append(f"新订单获单清单{ch} 已生成: {p}  ({n} 单)")
        else:
            log.append("新订单获单清单: 0 单")
    else:
        log.append("新订单获单清单: 跳过 (未传完整天猫导出，无法识别新订单)")

    # ---- 回传ERP销售上传表 (取消/无运单/已补运单 三类 Terms 写回一张，按店各一份) ----
    tag = f"{d.year}年{d.month:02d}月{d.day:02d}日平台订单取消"
    cats = {
        "取消":     (ann[ann["_cat"] == "取消"],     tag),
        "无运单":    (ann[ann["_cat"] == "无运单"],    lambda o: _wu_add(o["Terms and conditions"])),
        "已补运单":  (ann[ann["_cat"] == "已补运单"],  lambda o: _wu_strip(o["Terms and conditions"])),
    }
    present = {k: df for k, (df, _) in cats.items() if not df.empty}
    if present and idcol:
        parts = [_upload_terms(df, idcol, terms)
                 for k, (df, terms) in cats.items() if not df.empty]
        allup = pd.concat(parts, ignore_index=True)
        allup["_ch"] = allup["Order Reference"].astype(str).str.split("_", n=1).str[0]
        for ch in sorted(allup["_ch"].unique()):
            sub = allup[allup["_ch"] == ch].drop(columns="_ch")
            p, n = _write_simple(sub, outdir, f"回传ERP销售上传表{ch}.xlsx")
            log.append(f"回传ERP销售上传表{ch} 已生成: {p}  ({n} 单)")
        log.append("  └ 含 " + " / ".join(
            f"{k} {df['_key'].nunique()}" for k, df in present.items()))
    elif present:
        log.append("⚠ 有需回传订单(取消/无运单/已补运单)但 ERP 无 External ID 列，"
                   "无法生成上传表(请在订单导出勾上 External ID)")
    else:
        log.append("回传ERP销售上传表: 0 单 (无取消/无运单/已补运单)"
                   + ("" if full_tmall_path else " (未传完整天猫导出，取消无法识别)"))

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
            p = unique_path(os.path.join(outdir, f"已补运单清单{ch}.xlsx")); wb.save(p)
            log.append(f"已补运单清单{ch} 已生成: {p}  ({len(out)} 单)")
    else:
        log.append("已补运单清单: 0 单")

    # ---- 补货预判清单 (Solo 作战清单·模式一 step 0；需 ERP 含 FS/Safety/Remark) ----
    reorder = build_reorder(erp)
    if reorder is None:
        log.append("补货预判清单: 跳过 (ERP 订单导出未含 FS/Safety Stock/Supply Remark 列；"
                   "在 Odoo 订单导出模板勾上这 3 列即可生成)")
    elif reorder.empty:
        log.append("补货预判清单: 0 SKU")
    else:
        reorder["_ch"] = (erp.drop_duplicates(s4.ERP_INTERNAL)
                          .set_index(s4.ERP_INTERNAL)[s4.ERP_ORDER_REF]
                          .reindex(reorder["SKU"]).astype(str)
                          .str.split("_", n=1).str[0].values)
        for ch in sorted(reorder["_ch"].dropna().unique()):
            sub = reorder[reorder["_ch"] == ch].drop(columns="_ch")
            short = int((sub["缺口"] > 0).sum())
            p, n = _write_simple(sub, outdir, f"{ch}补货预判清单.xlsx")
            log.append(f"{ch}补货预判清单 已生成: {p}  ({n} SKU，其中缺口>0 {short} 个)")

    # ---- 异常上报(不静默) ----
    dup = erp.drop_duplicates(s4.ERP_ORDER_REF)["_key"].duplicated().sum()
    if dup:
        log.append(f"⚠ 连接键冲突: {dup} 个订单的 Order Reference 后15位与他单相同(可能误判状态)")
    # 注：发货范围由二段式从活单状态(新订单/商家已接单)推出，已结构性排除已发货/已收货/
    # 取消的历史单，故无需再做「名单过期」反查护栏。
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
        print("用法: python3 build_excel.py <erp[,erp2...]> <完整天猫导出> [out.xlsx]")
        return
    erp_paths = args[0].split(",")
    full_tmall = args[1]
    out_arg = args[2] if len(args) > 2 else None
    outdir = os.path.dirname(out_arg) if out_arg else "output"
    log, st = build(erp_paths, full_tmall, out_arg, outdir)
    for ln in log:
        print(ln)


if __name__ == "__main__":
    main()
