"""箱单预估：按 Odoo 产品主数据（箱规 / 整箱尺寸 / 整箱重）排托，填箱号、毛重、尺寸、体积重。

只在 `packing_list.py --so` 路径用（ISSUES [箱单] E）。纯函数，不连 Odoo，便于单测。

口径（2026-11-18 用户定）:
  - 箱数按 SO 订购量算，不按预留量；尾数单列一行「尾箱」，箱规 = 余数。
  - 120×80 托，托盘本身 15cm、空托 25kg，**托高（含托盘）上限 160cm**。
  - 同一张 SO 内可以拼托：按 SO 行顺序一层一层往上码，放不下就开新托。
    不同 SO 不拼（不同 SO 的货是分开打包的，同 read_orders 的口径）。
  - 可靠性分三档:
      A 箱规 + 整箱尺寸 + 整箱重都有 → 排托，毛重 / 托高 / 体积重都估（浅蓝斜体）。
      B 只有箱规 → 填箱规 / 箱数；毛重按单件 weight × 箱规粗估（红字），单独成托，
        托数按「件/托」(pallet_pieces) 算，托高 / 体积重留空。
      C 没有箱规 → 不估，留给仓库。
  - 体积重 = L×W×H / 6000。

主数据字段都是自由文本，解析规则见 ISSUES [箱单] E 的全库扫描结果。
"""
import math
import re

from openpyxl.styles import Alignment, Font, PatternFill

from common.xlsx import BORDER

PALLET_L, PALLET_W = 120, 80
PALLET_BASE_CM = 15
PALLET_KG = 25
MAX_HEIGHT_CM = 160
VOL_DIVISOR = 6000

EST_FILL = PatternFill("solid", fgColor="DDEBF7")
ROUGH_COLOR = "C00000"

COL_CARTON_NO, COL_BOXSPEC, COL_BOXCOUNT = 7, 8, 9
COL_WEIGHT, COL_DIMS, COL_VOLW = 13, 14, 15

_NUM = r"\d+(?:[.,]\d+)?"


def _num(s):
    return float(s.replace(",", "."))


# ---------------------------------------------------------------- 主数据解析
def parse_pieces(value):
    """carton_pieces → 正整数或 None。

    `88` / `1,000`（千分位）/ `200盒/箱` / `30/carton` / `12pcs per carton，11箱/层` 取开头整数；
    `一箱六罐`、`小12，大96`、`200个 / 100个` 这类有两个数或没有数字开头的算歧义 → None。
    """
    s = str(value or "").strip()
    if not s:
        return None
    if re.fullmatch(r"\d{1,3}(,\d{3})+", s):
        s = s.replace(",", "")
    if re.fullmatch(r"\d+", s):
        n = int(s)
        return n or None
    m = re.match(r"(\d+)\s*(盒/箱|/carton|pcs per carton)", s, re.I)
    if m:
        return int(m.group(1)) or None
    return None


def parse_weight_kg(value):
    """carton_weight → kg 或 None。`2.07KG` / `2kg` / `4400g` / 纯数字（>50 视为克）；0 或有歧义 → None。"""
    s = str(value or "").strip()
    m = re.fullmatch(rf"({_NUM})\s*(kg|g)?", s, re.I)
    if not m:
        return None
    v = _num(m.group(1))
    unit = (m.group(2) or "").lower()
    if unit == "g" or (not unit and v > 50):
        v /= 1000
    if not 0 < v <= 60:   # 一箱超过 60kg 或为 0 → 当录错，不用
        return None
    return v


def parse_dims_cm(value):
    """carton_dimensions → (L, W, H) cm 或 None。

    分隔符 x / * / + / Tab / 空格都认，小数点 . 或 ,；明写 mm 或任一边 >100 → 按 mm ÷10。
    不是恰好三个数（例如 `45*37*29  /  39*29*22` 两种箱）→ None。
    """
    s = str(value or "").strip()
    if not s:
        return None
    nums = re.findall(_NUM, s)
    if len(nums) != 3:
        return None
    rest = re.sub(_NUM, "", s).lower().replace("cm", "").replace("mm", "")
    if not re.fullmatch(r"[\sx*+×]*", rest):
        return None
    dims = [_num(n) for n in nums]
    if "mm" in s.lower() or max(dims) > 100:
        dims = [d / 10 for d in dims]
    if min(dims) <= 0 or max(dims) > 100:
        return None
    return tuple(dims)


def parse_pallet_pieces(value):
    """pallet_pieces 纯整数时是「件/托」；`8箱/层` 之类的说明文字 → None。"""
    s = str(value or "").strip()
    return int(s) if re.fullmatch(r"\d+", s) and int(s) > 0 else None


def per_layer(dims):
    """120×80 托一层能放几箱（箱高朝上，两种摆向取多的）。"""
    L, W, _ = dims
    return int(max((PALLET_L // L) * (PALLET_W // W), (PALLET_L // W) * (PALLET_W // L)))


def classify(pack):
    """pack = 产品主数据原值 dict → 补上解析结果与档位 'A' / 'B' / 'C'。"""
    pcs = parse_pieces(pack.get("carton_pieces"))
    kg = parse_weight_kg(pack.get("carton_weight"))
    dims = parse_dims_cm(pack.get("carton_dimensions"))
    unit_kg = pack.get("weight") or 0
    out = dict(pack, pcs=pcs, carton_kg=kg, dims=dims, unit_kg=unit_kg,
               pallet_pcs=parse_pallet_pieces(pack.get("pallet_pieces")), warn=None)
    if pcs and kg and dims and per_layer(dims) and dims[2] <= MAX_HEIGHT_CM - PALLET_BASE_CM:
        out["tier"] = "A"
        if unit_kg:
            ratio = unit_kg * pcs / kg
            if not 0.6 <= ratio <= 1.1:
                out["warn"] = (f"单件 weight {unit_kg:g}kg × {pcs} = {unit_kg * pcs:g}kg，"
                               f"与整箱重 {kg:g}kg 对不上，按整箱重算")
    elif pcs:
        out["tier"] = "B"
    else:
        out["tier"] = "C"
    return out


# ---------------------------------------------------------------- 排托
def plan(orders, first_pallet=1, first_carton=1):
    """给每个 line 挂 line['segments']，返回托盘列表。

    segment = {pallet, first, last, pcs, count}：同一托上、同一箱规的一段连续箱号。
    pallet = {no, height (cm, 含托盘; None=估不出), rough, segments:[(line, seg)]}
    """
    pallets, pno, cno = [], first_pallet, first_carton
    for order in orders:
        cur, used = None, 0.0   # 当前可拼的 A 档托、已码高度（不含托盘）
        room = MAX_HEIGHT_CM - PALLET_BASE_CM
        for line in order["lines"].values():
            pack = line.get("pack")
            line["segments"] = []
            if not pack or pack["tier"] == "C" or not line["qty"]:
                continue
            pcs = pack["pcs"]
            qty = int(math.ceil(line["qty"]))
            n_cartons = -(-qty // pcs)
            tail = qty - (n_cartons - 1) * pcs   # 最后一箱的件数
            chunks = []   # [(pallet, n_cartons)]
            if pack["tier"] == "A":
                per, h = per_layer(pack["dims"]), pack["dims"][2]
                left = n_cartons
                while left:
                    if cur is None or used + h > room:
                        cur = {"no": pno, "height": None, "rough": False, "segments": []}
                        pallets.append(cur)
                        pno, used = pno + 1, 0.0
                    take = min(left, int((room - used) // h) * per)
                    used += math.ceil(take / per) * h
                    cur["height"] = math.ceil(used + PALLET_BASE_CM - 1e-9)   # 宁高勿低
                    chunks.append((cur, take))
                    left -= take
            else:   # B 档：单独成托，托数按件/托，高度未知
                per_pallet = (pack["pallet_pcs"] // pcs if pack["pallet_pcs"]
                              and pack["pallet_pcs"] >= pcs else n_cartons)
                left = n_cartons
                while left:
                    p = {"no": pno, "height": None, "rough": True, "segments": []}
                    pallets.append(p)
                    pno += 1
                    take = min(left, per_pallet)
                    chunks.append((p, take))
                    left -= take
            done = 0
            for pallet, take in chunks:
                start = cno
                cno += take
                done += take
                is_last = done == n_cartons
                full = take - 1 if is_last and tail != pcs else take
                if full:
                    seg = {"pallet": pallet, "first": start, "last": start + full - 1,
                           "pcs": pcs, "count": full}
                    line["segments"].append(seg)
                    pallet["segments"].append((line, seg))
                if full != take:
                    seg = {"pallet": pallet, "first": start + full, "last": start + full,
                           "pcs": tail, "count": 1}
                    line["segments"].append(seg)
                    pallet["segments"].append((line, seg))
    return pallets


def carton_label(seg):
    no = seg["pallet"]["no"]
    if seg["first"] == seg["last"]:
        return f"托{no}: {seg['first']}"
    return f"托{no}: {seg['first']}-{seg['last']}"


def _seg_kg(line):
    """每箱按多少 kg 算；B 档没有单件 weight 时 → None（毛重写 ?）。"""
    pack = line["pack"]
    if pack["tier"] == "A":
        return pack["carton_kg"]
    return pack["unit_kg"] * pack["pcs"] or None


# ---------------------------------------------------------------- 写表
def write(ws, orders, pallets, first_row):
    """把预估写进已排好版的表（line['_row'] 由 packing_list._write_line 记下）。
    返回 (合计行号, 说明文字)。"""
    rough_font = Font(name="宋体", size=11, italic=True, color=ROUGH_COLOR)
    est_font = Font(name="宋体", size=11, italic=True)
    for order in orders:
        for line in order["lines"].values():
            for i, seg in enumerate(line.get("segments", [])):
                r = line["_row"] + i
                seg["row"] = r
                ws.cell(r, COL_CARTON_NO, carton_label(seg))
                ws.cell(r, COL_BOXSPEC, seg["pcs"])
                ws.cell(r, COL_BOXCOUNT, seg["count"])
                for c in (COL_CARTON_NO, COL_BOXSPEC, COL_BOXCOUNT, COL_WEIGHT, COL_DIMS, COL_VOLW):
                    ws.cell(r, c).fill = EST_FILL
                    ws.cell(r, c).font = est_font

    for p in pallets:
        rows = [seg["row"] for _, seg in p["segments"]]
        r0 = rows[0]
        kgs = [(seg["row"], _seg_kg(line)) for line, seg in p["segments"]]
        if all(kg for _, kg in kgs):
            terms = "+".join(f"I{r}*{kg:g}" for r, kg in kgs)
            ws.cell(r0, COL_WEIGHT, f"=ROUND({terms}+{PALLET_KG},1)")
        else:
            ws.cell(r0, COL_WEIGHT, "?")
        if p["height"]:
            ws.cell(r0, COL_DIMS, f"{PALLET_L}*{PALLET_W}*{p['height']}")
            ws.cell(r0, COL_VOLW, round(PALLET_L * PALLET_W * p["height"] / VOL_DIVISOR, 1))
        else:
            ws.cell(r0, COL_DIMS, f"{PALLET_L}*{PALLET_W}*?")
        if p["rough"]:
            ws.cell(r0, COL_WEIGHT).font = rough_font
            ws.cell(r0, COL_DIMS).font = rough_font
        if len(rows) > 1:
            if rows == list(range(r0, r0 + len(rows))):   # 行连续 → 合并，同 S04077 预估样式
                for c in (COL_WEIGHT, COL_DIMS, COL_VOLW):
                    ws.merge_cells(start_row=r0, start_column=c, end_row=rows[-1], end_column=c)
            else:   # 中间隔着别的 SKU 的空行 / 批次行 → 写「同托N」，不跨行合并
                for r in rows[1:]:
                    for c in (COL_WEIGHT, COL_DIMS, COL_VOLW):
                        ws.cell(r, c, f"同托{p['no']}")
                        ws.cell(r, c).font = est_font

    last = ws.max_row
    tr = last + 1
    ws.cell(tr, 2, "合计 Total（预估）")
    ws.cell(tr, COL_BOXCOUNT, f"=SUM(I{first_row}:I{last})")
    ws.cell(tr, 12, f"=SUM(L{first_row}:L{last})")
    ws.cell(tr, COL_WEIGHT, f"=SUM(M{first_row}:M{last})")
    ws.cell(tr, COL_DIMS, f"{len(pallets)} 托")
    ws.cell(tr, COL_VOLW, f"=SUM(O{first_row}:O{last})")
    for c in range(1, 17):
        cell = ws.cell(tr, c)
        cell.border = BORDER
        cell.font = Font(name="宋体", size=11, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[tr].height = 30
    return tr


def note_text(orders, made_on):
    """表末说明：口径 + 每个产品用的主数据 + 估不了 / 粗估 / 对不上的产品。"""
    basis, warns = [], []
    for order in orders:
        for line in order["lines"].values():
            pack = line.get("pack")
            if not pack:
                continue
            sku = line["sku"]
            if pack["tier"] == "A":
                L, W, H = pack["dims"]
                basis.append(f"{sku} {pack['pcs']} 件/箱，箱 {L:g}×{W:g}×{H:g}cm，整箱 {pack['carton_kg']:g}kg，"
                             f"每层 {per_layer(pack['dims'])} 箱")
                if pack["warn"]:
                    warns.append(f"{sku}：{pack['warn']}")
            elif pack["tier"] == "B":
                warns.append(f"{sku}：Odoo 只有箱规 {pack['pcs']} 件/箱，没有可用的整箱尺寸 / 重量"
                             f"（原值 尺寸「{pack.get('carton_dimensions') or ''}」重量「{pack.get('carton_weight') or ''}」），"
                             f"单独成托；红字毛重 = 单件 {pack['unit_kg']:g}kg × {pack['pcs']}，"
                             f"不含纸箱，单件 weight 全库不可靠，托高 / 体积重待实量")
            else:
                warns.append(f"{sku}：Odoo 没有可用的箱规（原值「{pack.get('carton_pieces') or ''}」），"
                             f"箱号 / 箱规 / 箱数 / 毛重留给仓库")
    text = (f"※ 浅蓝斜体格为 {made_on:%d.%m.%Y} 按 Odoo 产品主数据预估，数量按 SO 订购量："
            + "；".join(basis or ["（无）"])
            + f"。{PALLET_L}×{PALLET_W} 托，托高含托盘 ≤{MAX_HEIGHT_CM}cm（托盘 {PALLET_BASE_CM}cm、"
              f"空托 {PALLET_KG}kg），同一 SO 内按行顺序逐层拼托，毛重含托盘，尾箱按整箱重算，"
              f"体积重 = L×W×H/{VOL_DIVISOR}。批次号为出库单当前预留，不对应箱号。到货打包后以实称实量为准。")
    if warns:
        text += "\n⚠ " + "\n⚠ ".join(warns)
    return text, warns


def write_note(ws, row, text):
    ws.cell(row, 1, text)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=16)
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).font = Font(name="宋体", size=10)
    ws.row_dimensions[row].height = 45 + 15 * text.count("\n") + 15 * (len(text) // 160)
