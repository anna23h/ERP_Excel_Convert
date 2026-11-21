#!/usr/bin/env python3
"""天猫面单 PDF 按「拣货表+面单」顺序重排、合并成一份（ISSUES: [拣货面单] L）。

天猫批量打印下载的面单 PDF 每份最多 100 页，页序与阶段一产出的「面单」sheet 序号对不上。
本脚本按每页上的 LP 运单号（= 面单 sheet 的 VO Tracking No）重排，合并成一份 PDF。

用法:
    python vo_orders/sort_labels.py <拣货表+面单.xlsx> <面单.pdf> [<面单.pdf> ...] [-o 输出目录]
默认输出到 xlsx 同目录：…{店}{n}单 面单（按拣货表排序）.pdf；原 PDF / xlsx 只读。

两边对不上就全停、不出 PDF（退出码 2）：某页没有 LP 或有多个；同一运单两页；
Excel 有 PDF 缺；PDF 有 Excel 无。多品订单的续行同运单号只算一次，取首次出现的位置。

xlsx 带隐藏 sheet「分隔页」（阶段一 `--face-order sku` 产出，ISSUES: [拣货面单] O）时，
在 sheet 列出的每组（运单 ≥3 张的 SKU 组，[拣货面单] P）面单下面插一页程序生成的分隔页
（与面单页同尺寸，印 SKU / 拣货名 / 中文名 / 德文名 / 张数 / 件数，不印编号）。
操作员倒着翻（从最后一页往前），先看到分隔页再看到这一组。没有这个 sheet 时行为不变。
分隔页要 reportlab（requirements-labels.txt），中文字体按 Windows / Mac 常见路径找。
"""
import argparse
import os
import re
import sys

from openpyxl import load_workbook

FACE_SHEET = "面单"
TRACKING_COL = "VO Tracking No"
LP_RE = re.compile(r"LP\d{14}")
OUT_SUFFIX = "面单（按拣货表排序）"


class MismatchError(ValueError):
    """面单 PDF 与拣货表对不上（CLI 退出码 2）。"""


def _pdfium():
    try:
        import pypdfium2
    except ImportError:
        raise ValueError("缺 pypdfium2：pip install -r requirements.txt") from None
    return pypdfium2


def read_order(xlsx):
    """「面单」sheet 的运单号顺序（去重，取首次出现）。按列名取，不依赖列位置。"""
    wb = load_workbook(xlsx, read_only=True, data_only=True)
    if FACE_SHEET not in wb.sheetnames:
        raise ValueError(f"{os.path.basename(xlsx)} 里没有「{FACE_SHEET}」sheet，不是拣货表+面单？")
    rows = wb[FACE_SHEET].iter_rows(values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(rows, [])]
    if TRACKING_COL not in header:
        raise ValueError(f"「{FACE_SHEET}」sheet 缺「{TRACKING_COL}」列")
    col = header.index(TRACKING_COL)
    order, seen = [], set()
    for r in rows:
        lp = str(r[col]).strip() if col < len(r) and r[col] is not None else ""
        if lp and lp not in seen:
            seen.add(lp)
            order.append(lp)
    wb.close()
    if not order:
        raise ValueError(f"「{FACE_SHEET}」sheet 没有运单号")
    return order


SEP_SHEET = "分隔页"          # 与 build_excel.SEP_SHEET 同名（这里不 import build_excel，免得拖进 pandas）
SEP_REQUIRED = ("打包顺序", "类型", "Internal Reference", "拣货名", "德文品名", "中文品名", "张数", "件数",
                "插在运单后")


def read_separators(xlsx):
    """隐藏 sheet「分隔页」→ [dict]（按打包顺序）；没有这个 sheet 返回 []。"""
    wb = load_workbook(xlsx, read_only=True, data_only=True)
    try:
        if SEP_SHEET not in wb.sheetnames:
            return []
        rows = wb[SEP_SHEET].iter_rows(values_only=True)
        header = [str(c).strip() if c is not None else "" for c in next(rows, [])]
        miss = [c for c in SEP_REQUIRED if c not in header]
        if miss:
            raise ValueError(f"「{SEP_SHEET}」sheet 缺列：{', '.join(miss)}")
        out = []
        for r in rows:
            d = {h: ("" if v is None else v) for h, v in zip(header, r)}
            if str(d["插在运单后"]).strip():
                d["插在运单后"] = str(d["插在运单后"]).strip()
                out.append(d)
        return out
    finally:
        wb.close()


# 中文字体候选（TTF / TrueType 轮廓的 TTC）：Windows 雅黑 / 黑体 / 宋体，Mac Arial Unicode / 黑体。
# 都找不到就退回 reportlab 内置 CID 字体 STSong-Light（不嵌入，靠阅读器自带字体，打印前要看一眼）。
CJK_FONTS = [
    (r"C:\Windows\Fonts\msyh.ttc", 0), (r"C:\Windows\Fonts\simhei.ttf", 0),
    (r"C:\Windows\Fonts\simsun.ttc", 0),
    ("/Library/Fonts/Arial Unicode.ttf", 0), ("/System/Library/Fonts/STHeiti Medium.ttc", 0),
]


def _cjk_font():
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError:
        raise ValueError("插分隔页要 reportlab：pip install -r requirements-labels.txt") from None
    for path, idx in CJK_FONTS:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("SepCJK", path, subfontIndex=idx))
                return "SepCJK", path
            except Exception:     # TTC 里是 CFF 轮廓等 reportlab 不支持的，换下一个
                continue
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    return "STSong-Light", "STSong-Light（未嵌入）"


def _wrap(text, font, size, width, stringWidth):
    """按宽度折行：拉丁文按空格断词，中文逐字；单词本身超宽就硬断。"""
    lines, cur = [], ""
    for tok in re.findall(r"[A-Za-z0-9ÄÖÜäöüß.,%+\-/()&']+\s*|\s+|.", str(text)):
        if stringWidth(cur + tok, font, size) <= width:
            cur += tok
            continue
        if cur.strip():
            lines.append(cur.rstrip())
        cur = tok.lstrip()
        while stringWidth(cur, font, size) > width and len(cur) > 1:
            k = len(cur)
            while k > 1 and stringWidth(cur[:k], font, size) > width:
                k -= 1
            lines.append(cur[:k])
            cur = cur[k:]
    if cur.strip():
        lines.append(cur.rstrip())
    return lines


def _fit(text, font, max_size, width, stringWidth, min_size=8):
    size = max_size
    while size > min_size and stringWidth(str(text), font, size) > width:
        size -= 0.5
    return size


def render_separators(seps, width, height):
    """分隔页 PDF（bytes），每条一页，页面 width×height pt（与面单页同尺寸）。返回 (bytes, 字体说明)。"""
    import io
    font, font_desc = _cjk_font()
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    m = width * 0.06
    inner = width - 2 * m
    for s in seps:
        band = height * 0.10                       # 顶部黑条：纸堆侧面也能看出是分隔页
        c.setFillColorRGB(0, 0, 0)
        c.rect(0, height - band, width, band, stroke=0, fill=1)
        c.setFillColorRGB(1, 1, 1)
        head = "分隔页"                             # 不印编号（ISSUES: [拣货面单] P）
        c.setFont(font, _fit(head, font, band * 0.45, inner, stringWidth))
        c.drawString(m, height - band * 0.68, head)
        c.setFillColorRGB(0, 0, 0)
        y = height - band - height * 0.04
        big = str(s["Internal Reference"])
        sz = _fit(big, font, height * 0.075, inner, stringWidth)
        y -= sz
        c.setFont(font, sz)
        c.drawString(m, y, big)
        y -= height * 0.03
        # 拣货名（仓库自己的识货提示）最显眼，其次中文标题，德文全名最小
        shown = set()
        for txt, mx in ((s["拣货名"], height * 0.05), (s["中文品名"], height * 0.045),
                        (s["德文品名"], height * 0.032)):
            k = re.sub(r"\s+", "", str(txt)).lower()
            if not k or k in shown:        # 天猫标题有时就是德文全名，不重复印
                continue
            shown.add(k)
            c.setFont(font, mx)
            for ln in _wrap(txt, font, mx, inner, stringWidth)[:2]:
                y -= mx * 1.2
                c.drawString(m, y, ln)
            y -= height * 0.02
        cnt = f"{s['张数']} 张  ·  {s['件数']} 件"
        sz = _fit(cnt, font, height * 0.09, inner, stringWidth)
        c.setFont(font, sz)
        c.drawString(m, height * 0.08, cnt)
        c.setLineWidth(3)
        c.rect(m * 0.5, m * 0.5, width - m, height - band - m * 0.5, stroke=1, fill=0)
        c.showPage()
    c.save()
    return buf.getvalue(), font_desc


def scan_pdf(path):
    """每页抽到的 LP 号列表（去重保序）。"""
    pdfium = _pdfium()
    doc = pdfium.PdfDocument(path)
    try:
        pages = []
        for page in doc:
            text = page.get_textpage().get_text_bounded()
            pages.append(list(dict.fromkeys(LP_RE.findall(text))))
        return pages
    finally:
        doc.close()


def plan(order, scans):
    """order: Excel 运单顺序；scans: [(pdf 路径, [每页 LP 列表])]（列表而非 dict：同一文件选两次也要报出来）。
    返回 [(pdf 路径, 页下标)]，按 order 排。对不上抛 MismatchError，列全所有问题。"""
    name = os.path.basename
    problems, loc = [], {}
    for path, pages in scans:
        for i, lps in enumerate(pages):
            where = f"{name(path)} 第 {i + 1} 页"
            if len(lps) != 1:
                problems.append(f"{where}：{'没有 LP 运单号' if not lps else '有多个 LP：' + ', '.join(lps)}")
                continue
            lp = lps[0]
            if lp in loc:
                p0, i0 = loc[lp]
                problems.append(f"{lp} 重复：{name(p0)} 第 {i0 + 1} 页 与 {where}")
            else:
                loc[lp] = (path, i)
    want = set(order)
    missing = [lp for lp in order if lp not in loc]
    extra = [lp for lp in loc if lp not in want]
    if missing:
        problems.append(f"拣货表有、PDF 缺 {len(missing)} 单：" + ", ".join(missing))
    for lp in extra:
        p, i = loc[lp]
        problems.append(f"PDF 有、拣货表无：{lp}（{name(p)} 第 {i + 1} 页）")
    if problems:
        raise MismatchError("面单 PDF 与拣货表对不上，未生成 PDF：\n  " + "\n  ".join(problems))
    return [loc[lp] for lp in order]


def out_path(xlsx, outdir=None):
    stem = os.path.splitext(os.path.basename(xlsx))[0]
    stem = stem.replace("拣货表+面单", OUT_SUFFIX) if "拣货表+面单" in stem else f"{stem} {OUT_SUFFIX}"
    return os.path.join(outdir or os.path.dirname(os.path.abspath(xlsx)), stem + ".pdf")


def run(xlsx, pdfs, outdir=None):
    """返回 (输出路径, 日志行)。"""
    pdfium = _pdfium()
    if not pdfs:
        raise ValueError("没有选面单 PDF")
    order = read_order(xlsx)
    seps = read_separators(xlsx)
    after = {}
    for s in seps:
        lp = s["插在运单后"]
        if lp not in set(order) or lp in after:
            raise MismatchError(f"「{SEP_SHEET}」sheet 的运单 {lp} 不在面单里或重复，未生成 PDF"
                                "（xlsx 被改过？请重新跑阶段一）")
        after[lp] = s
    seen = set()
    for p in pdfs:
        k = os.path.normcase(os.path.abspath(p))
        if k in seen:
            raise MismatchError(f"同一份 PDF 选了两次，未生成 PDF：{p}")
        seen.add(k)
    scans = [(p, scan_pdf(p)) for p in pdfs]
    picks = plan(order, scans)
    out = out_path(xlsx, outdir)
    if os.path.normcase(os.path.abspath(out)) in seen:
        raise ValueError(f"输出文件就是输入之一，请别把上次合并出的 PDF 再选进来：{out}")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    docs = {p: pdfium.PdfDocument(p) for p in pdfs}
    new = pdfium.PdfDocument.new()
    sep_doc, font_desc = None, ""
    try:
        if seps:
            p0, i0 = picks[0]
            w, h = docs[p0][i0].get_size()
            data, font_desc = render_separators(seps, w, h)
            sep_doc = pdfium.PdfDocument(data)
            idx = {s["插在运单后"]: k for k, s in enumerate(seps)}
        for lp, (p, i) in zip(order, picks):
            new.import_pages(docs[p], [i])
            if lp in after:
                new.import_pages(sep_doc, [idx[lp]])
        new.save(out)
        npages = len(new)
    finally:
        new.close()
        if sep_doc is not None:
            sep_doc.close()
        for d in docs.values():
            d.close()
    lines = [f"面单 PDF {len(pdfs)} 份 / {sum(len(pg) for _, pg in scans)} 页，"
             f"拣货表 {len(order)} 单，一一对应"]
    if seps:
        lines.append(f"插分隔页 {len(seps)} 页（插在对应 SKU 组的运单下面；字体 {font_desc}）")
    lines.append(f"按拣货表排序的面单 已生成: {out}  ({npages} 页)")
    return out, lines


def main(argv=None):
    ap = argparse.ArgumentParser(description="天猫面单 PDF 按拣货表+面单顺序重排、合并成一份")
    ap.add_argument("xlsx", help="阶段一产出的 …拣货表+面单.xlsx")
    ap.add_argument("pdfs", nargs="+", help="天猫下载的面单 PDF（可多份）")
    ap.add_argument("-o", "--outdir", help="输出目录（默认 xlsx 同目录）")
    a = ap.parse_args(argv)
    try:
        _, lines = run(a.xlsx, a.pdfs, a.outdir)
    except MismatchError as e:
        print(e, file=sys.stderr)
        return 2
    except (ValueError, OSError) as e:
        print(f"出错：{e}", file=sys.stderr)
        return 1
    print("\n".join(lines))
    return 0


if __name__ == "__main__":
    sys.exit(main())
