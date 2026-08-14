#!/usr/bin/env python3
"""
储位标签生成器 — 尺寸驱动、点阵对齐

核心原则:
  1. 所有几何尺寸先换算成打印机 dot，向下取整对齐，再换回 mm。
     这样模块边界永远落在打印头点阵上,栅格化零舍入误差。
  2. QR version 显式锁定。编码变长时直接报错,而不是静默升版缩小模块。
  3. 模块边长有硬下限校验,低于扫描枪规格直接拒绝生成。

用法:
  python3 make_labels.py codes.txt              # 生成 PDF
  python3 make_labels.py codes.txt --png        # 额外生成 1-bit PNG(最高保真)
  python3 make_labels.py codes.txt --verify     # 打印前退化回测
"""
from __future__ import annotations
import sys, math, argparse, os
from dataclasses import dataclass, field

# Windows 控制台默认 cp1252，打印中文文件名会 UnicodeEncodeError，强制 UTF-8
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8")

import segno
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# ────────────────────────────── 配置 ──────────────────────────────

@dataclass
class Config:
    # —— 载体(贴纸)——
    label_w_mm: float = 40.0
    label_h_mm: float = 20.0
    edge_mm: float = 0.75           # QR 距标签上下边的最小留白(静区已含在 QR 块内)
    margin_mm: float = 1.0          # 文字侧安全边距

    # —— 打印机 ——
    dpi: int = 203                  # Deli DL720C = 203dpi (1 dot = 0.125mm)

    # —— 二维码 ——
    qr_version: int = 1             # 锁定; None = 自动(不推荐)
    qr_ecc: str = "M"               # L/M/Q/H
    qr_border: int = 4              # 静区模块数,QR 标准强制 4
    min_module_mm: float = 0.30     # 硬下限。DS9308 规格约 0.17,留足余量

    # —— 文字 ——
    text_split: int = 2             # 在第几个 '-' 处折行
    drop_trailing_hyphen: bool = True
    text_gap_mm: float = 1.2        # QR 与文字之间的间隙
    line_spacing: float = 1.06      # 行距系数
    font_name: str = "Helvetica-Bold"

    # 派生量
    dot_mm: float = field(init=False)

    def __post_init__(self):
        self.dot_mm = 25.4 / self.dpi

    def snap(self, mm_val: float) -> float:
        """把 mm 吸附到打印机点阵整数倍"""
        return round(mm_val / self.dot_mm) * self.dot_mm

    def dots(self, mm_val: float) -> int:
        return round(mm_val / self.dot_mm)


# ────────────────────────── 布局计算 ──────────────────────────

@dataclass
class Layout:
    cell_dots: int
    cell_mm: float
    qr_block_mm: float      # 含静区的总边长
    qr_x_mm: float          # 含静区块的左下角(PDF 坐标系)
    qr_y_mm: float
    text_x_mm: float
    text_w_mm: float
    font_size_pt: float


def compute_layout(cfg: Config, sample: str) -> Layout:
    qr = segno.make(sample, error=cfg.qr_ecc, version=cfg.qr_version)
    n_mod = qr.symbol_size(border=0)[0]              # V1 -> 21
    total_mod = n_mod + 2 * cfg.qr_border            # -> 29

    # 垂直方向决定模块大小: 用满可用高度,向下取整到整数 dot
    avail_h = cfg.label_h_mm - 2 * cfg.edge_mm
    cell_dots = math.floor(avail_h / total_mod / cfg.dot_mm)
    cell_mm = cell_dots * cfg.dot_mm

    if cell_mm < cfg.min_module_mm:
        raise ValueError(
            f"模块边长 {cell_mm:.3f}mm 低于下限 {cfg.min_module_mm}mm。\n"
            f"  标签高 {cfg.label_h_mm}mm 装不下 {total_mod} 模块。\n"
            f"  对策: 加大标签 / 降 ECC / 改用 DataMatrix(静区仅 1 模块)。"
        )

    qr_block = total_mod * cell_mm
    # 整块吸附到点阵整数位置 —— 这是消除模糊的关键
    qr_x = cfg.snap(cfg.edge_mm)
    qr_y = cfg.snap((cfg.label_h_mm - qr_block) / 2)

    text_x = cfg.snap(qr_x + qr_block + cfg.text_gap_mm)
    text_w = cfg.label_w_mm - cfg.margin_mm - text_x
    if text_w <= 0:
        raise ValueError("文字区宽度为负,标签太窄或 QR 太大")

    return Layout(cell_dots, cell_mm, qr_block, qr_x, qr_y,
                  text_x, text_w, 0.0)


def split_code(code: str, cfg: Config) -> tuple[str, str]:
    """JDH-01-01-1 -> ('JDH-01', '01-1')"""
    parts = code.split("-")
    if len(parts) <= cfg.text_split:
        return code, ""
    l1 = "-".join(parts[:cfg.text_split])
    l2 = "-".join(parts[cfg.text_split:])
    if not cfg.drop_trailing_hyphen:
        l1 += "-"
    return l1, l2


def fit_font(lines: list[str], cfg: Config, lay: Layout) -> float:
    """在文字区内自适应最大字号(pt)"""
    from reportlab.pdfbase.pdfmetrics import stringWidth
    avail_w_pt = lay.text_w_mm * mm
    avail_h_pt = (cfg.label_h_mm - 2 * cfg.margin_mm) * mm
    size = 4.0
    while size < 60:
        nxt = size + 0.25
        w = max(stringWidth(t, cfg.font_name, nxt) for t in lines if t)
        h = nxt * cfg.line_spacing * len([t for t in lines if t])
        if w > avail_w_pt or h > avail_h_pt:
            break
        size = nxt
    return size


# ────────────────────────── PDF 输出 ──────────────────────────

def draw_qr_vector(c, qr, x_mm, y_mm, cell_mm, border):
    """手动逐模块绘制,完全掌控坐标,不依赖 segno 的 PDF 输出"""
    matrix = list(qr.matrix)
    n = len(matrix)
    c.setFillColorRGB(0, 0, 0)
    for r, row in enumerate(matrix):
        # 同行连续黑模块合并成一个矩形,减少路径数、避免相邻边缝隙
        run = 0
        for col in range(n + 1):
            bit = row[col] if col < n else 0
            if bit:
                run += 1
            elif run:
                px = x_mm + (border + col - run) * cell_mm
                py = y_mm + (border + n - 1 - r) * cell_mm
                c.rect(px * mm, py * mm, run * cell_mm * mm, cell_mm * mm,
                       stroke=0, fill=1)
                run = 0


def build_pdf(codes, cfg: Config, out_path: str) -> Layout:
    lay = compute_layout(cfg, codes[0])
    l1, l2 = split_code(codes[0], cfg)
    lay.font_size_pt = fit_font([l1, l2], cfg, lay)

    c = canvas.Canvas(out_path,
                      pagesize=(cfg.label_w_mm * mm, cfg.label_h_mm * mm))
    for code in codes:
        qr = segno.make(code, error=cfg.qr_ecc, version=cfg.qr_version)
        draw_qr_vector(c, qr, lay.qr_x_mm, lay.qr_y_mm,
                       lay.cell_mm, cfg.qr_border)

        a, b = split_code(code, cfg)
        c.setFont(cfg.font_name, lay.font_size_pt)
        c.setFillColorRGB(0, 0, 0)
        lh = lay.font_size_pt * cfg.line_spacing
        block_h = lh * (2 if b else 1)
        base = (cfg.label_h_mm * mm + block_h) / 2 - lay.font_size_pt
        tx = lay.text_x_mm * mm
        c.drawString(tx, base, a)
        if b:
            c.drawString(tx, base - lh, b)
        c.showPage()
    c.save()
    return lay


# ────────────────────────── PNG 输出(最高保真) ──────────────────────────

def build_png(codes, cfg: Config, lay: Layout, out_dir: str):
    """严格 dpi 的 1-bit 位图,模块边长 = 整数像素,驱动必须 1:1 送图"""
    import os
    os.makedirs(out_dir, exist_ok=True)
    W, H = cfg.dots(cfg.label_w_mm), cfg.dots(cfg.label_h_mm)
    try:
        font = ImageFont.truetype("DejaVuSans-Bold.ttf",
                                  int(lay.font_size_pt / 72 * cfg.dpi))
    except Exception:
        font = ImageFont.load_default()

    for code in codes:
        qr = segno.make(code, error=cfg.qr_ecc, version=cfg.qr_version)
        img = Image.new("1", (W, H), 1)
        d = ImageDraw.Draw(img)
        ox, oy = cfg.dots(lay.qr_x_mm), cfg.dots(lay.qr_y_mm)
        cd, b = lay.cell_dots, cfg.qr_border
        for r, row in enumerate(qr.matrix):
            for col, bit in enumerate(row):
                if bit:
                    x0 = ox + (b + col) * cd
                    y0 = oy + (b + r) * cd
                    d.rectangle([x0, y0, x0 + cd - 1, y0 + cd - 1], fill=0)
        a, bb = split_code(code, cfg)
        tx = cfg.dots(lay.text_x_mm)
        lh = int(lay.font_size_pt / 72 * cfg.dpi * cfg.line_spacing)
        ty = (H - lh * (2 if bb else 1)) // 2
        d.text((tx, ty), a, font=font, fill=0)
        if bb:
            d.text((tx, ty + lh), bb, font=font, fill=0)
        img.save(f"{out_dir}/{code}.png", dpi=(cfg.dpi, cfg.dpi))


# ────────────────────────── 退化回测 ──────────────────────────

def _degrade(a, cell, gain_ratio, blur, noise, rng):
    """模拟热敏打印退化: 黑色横向扩散(dot gain) + 打印头模糊 + 介质噪声"""
    import numpy as np, cv2
    g = max(1, int(round(cell * gain_ratio)))
    k = 2 * g + 1
    d = cv2.erode(a, np.ones((k, k), np.uint8))          # 白收缩 = 黑扩散 g 像素
    if blur:
        d = cv2.GaussianBlur(d, (blur, blur), 0)
    if noise:
        d = np.clip(d.astype(int) + rng.normal(0, noise, d.shape), 0, 255)
    return d.astype(np.uint8)


def _decode(img) -> str:
    """优先用 zbar(行为最接近真实扫描枪),回退到 OpenCV"""
    try:
        from pyzbar.pyzbar import decode as _zd
        from PIL import Image as _I
        r = _zd(_I.fromarray(img))
        return r[0].data.decode() if r else ""
    except ImportError:
        import cv2
        try:
            t, *_ = cv2.QRCodeDetector().detectAndDecode(img)
            return t
        except cv2.error:
            return ""


def verify(codes, cfg: Config, lay: Layout, gain_levels=(0.10, 0.20),
           trials=3, seed=0):
    """打印前退化回测。gain_ratio = 黑模块单边扩散量 / 模块边长。
    热敏实测通常 0.05~0.15;0.20 已属浓度偏高的坏情况。"""
    import numpy as np, cv2
    rng = np.random.default_rng(seed)
    cd, b = lay.cell_dots, cfg.qr_border
    fails = []
    for code in codes:
        qr = segno.make(code, error=cfg.qr_ecc, version=cfg.qr_version)
        n = qr.symbol_size(border=0)[0]
        side = (n + 2 * b) * cd
        a = np.full((side, side), 255, np.uint8)
        for r, row in enumerate(qr.matrix):
            for col, bit in enumerate(row):
                if bit:
                    a[(b+r)*cd:(b+r+1)*cd, (b+col)*cd:(b+col+1)*cd] = 0
        bad = 0
        for g in gain_levels:
            for _ in range(trials):
                d = _degrade(a, cd, g, 3, 15, rng)
                d = cv2.resize(d, None, fx=4, fy=4,
                               interpolation=cv2.INTER_LINEAR)
                txt = _decode(d)
                if txt != code:
                    bad += 1
        if bad:
            fails.append((code, bad, len(gain_levels) * trials))
    return fails


# ────────────────────────── 入口 ──────────────────────────

def read_codes(path: str, column: str) -> list[str]:
    """从 .txt(每行一个编码) 或 Excel(取指定列) 读出编码，去重保序。"""
    if path.lower().endswith((".xlsx", ".xls", ".xlsm")):
        import pandas as pd
        df = pd.read_excel(path)
        if column not in df.columns:
            sys.exit(f"找不到列 {column}；可用列: {list(df.columns)}")
        codes = [str(v).strip() for v in df[column].dropna() if str(v).strip()]
    else:
        with open(path, encoding="utf-8-sig") as f:
            codes = [ln.strip() for ln in f if ln.strip()]
    return list(dict.fromkeys(codes))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("codes", help=".txt(每行一个) 或 Excel 表格")
    ap.add_argument("-c", "--column", default="储位编码", help="Excel 里取哪一列")
    ap.add_argument("-o", "--outdir", default="output/labels", help="输出目录")
    ap.add_argument("--png", action="store_true")
    ap.add_argument("--verify", action="store_true")
    ap.add_argument("--dpi", type=int, default=203)
    ap.add_argument("--width", type=float, default=40.0)
    ap.add_argument("--height", type=float, default=20.0)
    ap.add_argument("--ecc", default="M")
    args = ap.parse_args()

    cfg = Config(label_w_mm=args.width, label_h_mm=args.height,
                 dpi=args.dpi, qr_ecc=args.ecc)
    codes = read_codes(args.codes, args.column)
    if not codes:
        sys.exit("没有读到任何编码")

    os.makedirs(args.outdir, exist_ok=True)
    pdf_path = os.path.join(args.outdir, "储位标签_QR.pdf")

    lay = build_pdf(codes, cfg, pdf_path)
    print(f"标签 {cfg.label_w_mm}x{cfg.label_h_mm}mm @ {cfg.dpi}dpi")
    print(f"模块边长 {lay.cell_mm:.3f}mm = {lay.cell_dots} dots")
    print(f"QR 含静区 {lay.qr_block_mm:.2f}mm, 起点 ({lay.qr_x_mm:.3f}, {lay.qr_y_mm:.3f})mm"
          f" = ({cfg.dots(lay.qr_x_mm)}, {cfg.dots(lay.qr_y_mm)}) dots")
    print(f"文字起点 {lay.text_x_mm:.2f}mm, 字号 {lay.font_size_pt:.1f}pt")
    print(f"-> {pdf_path}  ({len(codes)} 张)")

    csv_path = os.path.join(args.outdir, "储位编码.csv")
    with open(csv_path, "w", encoding="utf-8-sig") as f:
        f.write("code\n")
        f.writelines(f"{c}\n" for c in codes)
    print(f"-> {csv_path}")

    if args.verify:
        f = verify(codes, cfg, lay)
        print(f"退化回测: {len(codes)-len(f)}/{len(codes)} 全通过"
              + (f"\n  存在失败样本: {[x[0] for x in f[:8]]}" if f else ""))
    if args.png:
        png_dir = os.path.join(args.outdir, "png")
        build_png(codes, cfg, lay, png_dir)
        print(f"-> {png_dir}/")


if __name__ == "__main__":
    main()
