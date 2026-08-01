"""purchase order 导出 → 按基础 SKU 聚合的采购画像。

原先长在 `vo_orders/build_excel.py` 里，2026-08-01 搬来 common/：
它被**两条流水线**用着——VOTool 的补货预判清单，和 FS 回写——正好符合 common/ 的
收纳标准。搬之前 `fs_writeback` 得裸 `import build_excel`，逼得调用方(如
`erp_writeback_gui.py`)把 `vo_orders/` 塞进 sys.path 才跑得起来。
"""
import re

import pandas as pd

from common.vendor import vendor_map


# 采购画像追加列(来自 purchase order 导出，见 load_po_stats)
PO_COLS = ["供应商(次数)", "最低价", "最低价供应商", "最近一次采购", "采购总量"]

# 采购单里伪装成供应商的客户(实为我方客户，属噪音，整行剔除)
PO_CUSTOMER_PAT = "Alibaba Health"
# 采购单里不是真实进货的行，整行剔除——不然会污染采购画像。
#   Alibaba Health: 伪装成供应商的客户(实为我方客户)
#   VO Test Order : 测试单。2026-08-01 实测不滤的话有 397 个商品的 FS 会被写成 "VO"
PO_NOISE_PATS = [PO_CUSTOMER_PAT, "VO Test Order"]

def _po_base_sku(s):
    """SKU 归一：去掉多件装 x2 / 变体 *2 / 渠道 _VO 等尾缀，对齐采购单里的基础 SKU。"""
    return re.sub(r"(x\d+|\*\d+|_VO)+$", "", str(s).strip())


def load_po_stats(path):
    """purchase order 导出(Odoo 行式：订单头只在每单首行) → 按基础 SKU 聚合采购画像。
    最低价只统计单价>0(价 0/负数是赠品/返利)；窗口=导出里有多少算多少，不写死3月。
    返回 (stats_df[_sku + PO_COLS], 窗口描述str)。缺必需列抛 ValueError。"""
    po = pd.read_excel(path, dtype=str)
    need = ["Order Reference", "Vendor", "Order Lines/Product/Internal Reference",
            "Order Lines/Unit Price", "Order Lines/Total Quantity", "Order Lines/Created on"]
    missing = [c for c in need if c not in po.columns]
    if missing:
        raise ValueError("采购单导出缺列: " + ", ".join(missing))
    po[["Order Reference", "Vendor"]] = po[["Order Reference", "Vendor"]].ffill()
    po = po.dropna(subset=["Order Lines/Product/Internal Reference", "Vendor"]).copy()
    is_noise = po["Vendor"].str.contains("|".join(PO_NOISE_PATS), case=False, na=False)
    n_cust = int(is_noise.sum())
    po = po[~is_noise].copy()
    po["Vendor"] = po["Vendor"].map(vendor_map(po["Vendor"].unique()))
    po["_sku"] = po["Order Lines/Product/Internal Reference"].map(_po_base_sku)
    po["_price"] = pd.to_numeric(po["Order Lines/Unit Price"], errors="coerce")
    po["_qty"] = pd.to_numeric(po["Order Lines/Total Quantity"], errors="coerce")
    po["_dt"] = pd.to_datetime(po["Order Lines/Created on"], errors="coerce")
    rows = []
    for sku, g in po.groupby("_sku"):
        vc = g.groupby("Vendor")["Order Reference"].nunique().sort_values(ascending=False)
        vendors = "\n".join(f"{v}×{n}" for v, n in vc.items())  # 多家纵向排开(单元格内换行)
        priced = g[g["_price"] > 0]
        if len(priced):
            low_row = priced.loc[priced["_price"].idxmin()]
            low, low_v = float(low_row["_price"]), low_row["Vendor"]
        else:
            low, low_v = None, ""
        last = ""
        if g["_dt"].notna().any():
            lr = g.loc[g["_dt"].idxmax()]
            price_s = f" @{lr['_price']:g}" if pd.notna(lr["_price"]) else ""
            # 主次排布：供应商+价格一行，日期换行
            last = f"{lr['Vendor']}{price_s}\n{lr['_dt']:%Y-%m-%d}"
        rows.append((sku, vendors, low, low_v, last, g["_qty"].sum()))
    stats = pd.DataFrame(rows, columns=["_sku"] + PO_COLS)
    stats["采购总量"] = pd.to_numeric(stats["采购总量"], errors="coerce").round().astype("Int64")
    info = (f"{po['_dt'].min():%Y-%m-%d}~{po['_dt'].max():%Y-%m-%d} "
            f"{po['Order Reference'].nunique()} 单 / {stats.shape[0]} SKU")
    if n_cust:
        info += f" (已剔除非进货行 {n_cust}: {'/'.join(PO_NOISE_PATS)})"
    return stats, info
