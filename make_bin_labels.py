#!/usr/bin/env python3
"""
储位编码 -> 条码标签 PDF (得力 DL-720C, 203 dpi 热敏)

每一页 = 一张标签，页面尺寸 = 实际标签物理尺寸，打印时必须选「实际大小 / 100%」。
条码模块宽度严格取打印头点距(0.125mm)的整数倍，避免栅格化时条宽不均。

用法:
    python3 make_bin_labels.py 储位基本信息列表.xlsx -o out/
"""
import argparse
import os
import sys

# Windows 控制台默认 cp1252，打印中文文件名会 UnicodeEncodeError，强制 UTF-8
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

import pandas as pd
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.graphics.barcode import code128

DPI = 203
DOT = 25.4 / DPI * mm          # 0.1251 mm，得力 DL-720C 打印头点距
FONT = "Helvetica-Bold"


def fit_font_size(text, max_width, max_size, font=FONT):
    """二分不必要，直接按比例缩放到 max_width 以内。"""
    w = stringWidth(text, font, max_size)
    return max_size if w <= max_width else max_size * max_width / w


def draw_code128_label(c, code, w, h):
    """在 (0,0)-(w,h) 的坐标系内画一张 Code128 标签。w/h 单位为 point。"""
    # X 尺寸 = 1 点 = 0.125mm(5 mil)。11 位字符 = 156 模块 = 19.52mm
    bar_width = 1 * DOT
    bar_height = 9.5 * mm
    bc = code128.Code128(code, barWidth=bar_width, barHeight=bar_height,
                         humanReadable=False, quiet=False)
    bx = (w - bc.width) / 2
    by = h - 1.5 * mm - bar_height
    bc.drawOn(c, bx, by)

    size = fit_font_size(code, w - 3 * mm, 15)
    c.setFont(FONT, size)
    c.drawCentredString(w / 2, 3.2 * mm, code)


def draw_qr_label(c, code, w, h):
    """QR 变体：左侧 QR，右侧大号人眼可读文字。"""
    import segno
    qr = segno.make_qr(code, error="m")          # 强制标准 QR，非 Micro QR
    mods = qr.symbol_size(border=0)[0]
    cell = 3 * DOT                                # 0.375mm/模块
    side = mods * cell
    qx, qy = 2 * mm, (h - side) / 2
    c.setFillColorRGB(0, 0, 0)
    for r, row in enumerate(qr.matrix):
        for col, val in enumerate(row):
            if val:
                c.rect(qx + col * cell, qy + side - (r + 1) * cell,
                       cell, cell, stroke=0, fill=1)

    tx0 = qx + side + 2 * mm
    tw = w - tx0 - 1.5 * mm
    head, tail = code[:7], code[7:]               # JDH-01- / 30-5
    s1 = fit_font_size(head, tw, 13)
    s2 = fit_font_size(tail, tw, 17)
    c.setFont(FONT, s1)
    c.drawString(tx0, h / 2 + 0.8 * mm, head)
    c.setFont(FONT, s2)
    c.drawString(tx0, h / 2 - 5.2 * mm, tail)


def build(codes, path, label_w, label_h, rotate, drawer):
    c = canvas.Canvas(path, pagesize=(label_w, label_h))
    c.setTitle(os.path.basename(path))
    # 内容始终按 40x20 横向绘制；rotate=True 时整体旋转 90° 贴到 20x40 页面
    cw, ch = (label_h, label_w) if rotate else (label_w, label_h)
    for code in codes:
        c.saveState()
        if rotate:
            c.translate(label_w, 0)
            c.rotate(90)
        drawer(c, code, cw, ch)
        c.restoreState()
        c.showPage()
    c.save()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("-c", "--column", default="储位编码")
    ap.add_argument("-o", "--outdir", default=".")
    args = ap.parse_args()

    if args.xlsx.lower().endswith(".txt"):
        with open(args.xlsx, encoding="utf-8-sig") as f:
            codes = [ln.strip() for ln in f if ln.strip()]
    else:
        df = pd.read_excel(args.xlsx)
        if args.column not in df.columns:
            sys.exit(f"找不到列 {args.column}；可用列: {list(df.columns)}")
        codes = [str(v).strip() for v in df[args.column].dropna()]
    codes = list(dict.fromkeys(codes))            # 去重，保持原顺序
    os.makedirs(args.outdir, exist_ok=True)

    jobs = [
        ("储位标签_40x20_Code128.pdf", 40 * mm, 20 * mm, False, draw_code128_label),
        ("储位标签_20x40_Code128.pdf", 20 * mm, 40 * mm, True, draw_code128_label),
        ("储位标签_40x20_QR.pdf",      40 * mm, 20 * mm, False, draw_qr_label),
    ]
    for name, w, h, rot, fn in jobs:
        p = os.path.join(args.outdir, name)
        build(codes, p, w, h, rot, fn)
        print(f"{p}  ({len(codes)} 页)")

    csv = os.path.join(args.outdir, "储位编码.csv")
    pd.DataFrame({"code": codes}).to_csv(csv, index=False, encoding="utf-8-sig")
    print(csv)


if __name__ == "__main__":
    main()
