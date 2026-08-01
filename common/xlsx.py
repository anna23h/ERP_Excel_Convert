"""Excel 排版 / 写盘的公共部件（原先长在 vo_orders/build_excel.py 里）。

搬出来的理由：reorder_helper 一直在掏 build_excel 的下划线私有函数用，
packing_list 干脆把样式代码复制了一份。两条新流水线会把这事再重复两遍。

**刻意不 import pandas**：write_df / write_simple 只用到 .columns 和 .iterrows()，
没碰任何 pandas API。这样不产 DataFrame 的流水线（packing_list 等）不必背 pandas
依赖，将来单独打包也不会把 pandas 塞进 exe。
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

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


def style_sheet(ws, n_cols, header_font=HEAD_FONT, left_cols=(), small_cols=(), widths=None):
    """left_cols: 内容左对齐+下沉的列名集合；small_cols: 字号小2号的列名集合。表头始终居中。
    widths: {表头名: 列宽} 固定列宽表，命中的列用固定宽度（操作员手工调好、免二次拖列），
    未命中的列仍按内容自动算宽。"""
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
        # 含单元格内换行(\n)的行按行数放大，否则固定高度会裁掉第二行起的内容；
        # 无换行的行恒为 ROW_H(其他表无 \n 内容，行为不变)
        lines = max((str(c.value).count("\n") + 1 for c in ws[r] if c.value is not None),
                    default=1)
        ws.row_dimensions[r].height = ROW_H if lines == 1 else lines * 22
    # 列宽：固定表命中的列用固定宽度（操作员调好的成品宽），其余按内容自动算宽
    # （字号15 比默认大，需放大系数，否则日期显示为 ######）
    widths = widths or {}
    for c in range(1, n_cols + 1):
        hdr = headers[c - 1] if c - 1 < len(headers) else None
        if hdr in widths:
            ws.column_dimensions[get_column_letter(c)].width = widths[hdr]
            continue
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


def unique_path(path):
    """目标已存在则在扩展名前加序号 (1)/(2)…，避免覆盖既有产出。"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 1
    while os.path.exists(f"{base}({i}){ext}"):
        i += 1
    return f"{base}({i}){ext}"


def write_simple(out, outdir, fname, n_cols=None, left_cols=(), small_cols=(), widths=None):
    """把一张 DataFrame 写成单 sheet workbook(统一样式)。返回 (路径, 行数)。
    left_cols/small_cols/widths 透传给 style_sheet（左对齐下沉/小字号/固定列宽），默认空=全居中自动宽。"""
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    write_df(ws, out)
    style_sheet(ws, n_cols or len(out.columns), left_cols=left_cols, small_cols=small_cols,
                widths=widths)
    path = unique_path(os.path.join(outdir, fname))
    wb.save(path)
    return path, len(out)


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
