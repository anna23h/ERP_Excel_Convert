#!/usr/bin/env python3
"""FS/Supply Remark 回写：从 purchase order 提炼供应商画像，生成 Odoo 导入文件。

用法:
    python3 fs_writeback.py <purchase.order.xlsx> <product.product.xlsx> [outdir]

范围: product 导出里的商品(按 Safety Stock>0 筛的囤货品)；无采购记录的商品整行跳过不动。
策略: FS = 近期供应商短名列表(覆盖；"/"分隔，对齐公司现有 "DM/ROSSMANN" 记法)；
      Supply Remark = 带日期画像段前置追加(原文保留；本脚本旧画像段按签名替换，重跑不堆叠)。
产出: 单文件两 sheet——「导入」(id/FS/Supply Remark，直接上 Odoo 导入界面) +
      「对照」(新旧对比，上传前人工复核)。
上传保持人工。product 导出本身即现值备份，请保留原文件。
个人月频维护工具，不进 GUI(同事无入口即不会误触 ERP 回写)。
"""
import sys, os, re
from datetime import date

import pandas as pd

import build_excel as be

PROD_NEED = ["Internal Reference", "FS", "Supply Remark", "External ID"]
# 本脚本写入的画像段签名：重跑时先剥掉旧段再前置新段，人工备注不受影响
SIG = re.compile(r"^\d{8}:近3月[^；]*(；\s*|$)")


def _picture(st, d):
    """一条 stats 行 → (FS 新值, Supply Remark 画像段)。"""
    vend_lines = str(st["供应商(次数)"]).split("\n")
    names = [v.rsplit("×", 1)[0] for v in vend_lines]
    n_orders = sum(int(v.rsplit("×", 1)[1]) for v in vend_lines)
    fs_new = "/".join(names)
    seg = f"{d}:近3月{n_orders}单"
    if pd.notna(st["最低价"]) and st["最低价"] != "":
        seg += f" 最低{st['最低价']:g}@{st['最低价供应商']}"
    if st["最近一次采购"]:
        vend_price, buy_date = str(st["最近一次采购"]).split("\n")
        seg += f" 最近{buy_date[5:]} {vend_price}"
    return fs_new, seg


def make_writeback(po_path, prod_path):
    """→ (导入df[id/FS/Supply Remark], 对照df, 跳过数, 采购窗口描述)。"""
    stats, info = be.load_po_stats(po_path)
    s = stats.set_index("_sku")
    prod = pd.read_excel(prod_path, dtype=str)
    missing = [c for c in PROD_NEED if c not in prod.columns]
    if missing:
        raise ValueError("product 导出缺列: " + ", ".join(missing))
    if not prod["External ID"].is_unique:
        raise ValueError("product 导出 External ID 有重复，无法作导入键")
    d = date.today().strftime("%Y%m%d")
    imp, chk, skipped = [], [], 0
    for _, p in prod.iterrows():
        base = be._po_base_sku(p["Internal Reference"])
        if base not in s.index:
            skipped += 1
            continue
        fs_new, seg = _picture(s.loc[base], d)
        old_sr = str(p["Supply Remark"]) if pd.notna(p["Supply Remark"]) else ""
        old_kept = SIG.sub("", old_sr).strip("； ").strip()
        sr_new = seg + (f"；{old_kept}" if old_kept else "")
        imp.append({"id": p["External ID"], "FS": fs_new, "Supply Remark": sr_new})
        chk.append({"Internal Reference": p["Internal Reference"],
                    "FS 旧": p["FS"] if pd.notna(p["FS"]) else "",
                    "FS 新": fs_new,
                    "Supply Remark 旧": old_sr,
                    "Supply Remark 新": sr_new})
    return pd.DataFrame(imp), pd.DataFrame(chk), skipped, info


def main():
    args = sys.argv[1:]
    if len(args) < 2:
        print(__doc__)
        return
    po_path, prod_path = args[0], args[1]
    outdir = args[2] if len(args) > 2 else "output"
    os.makedirs(outdir, exist_ok=True)
    imp, chk, skipped, info = make_writeback(po_path, prod_path)
    d = date.today().strftime("%Y%m%d")
    path = be.unique_path(os.path.join(outdir, f"FS回写导入 {d}.xlsx"))
    with pd.ExcelWriter(path) as xw:
        imp.to_excel(xw, sheet_name="导入", index=False)
        chk.to_excel(xw, sheet_name="对照", index=False)
    print(f"采购参考: {info}")
    print(f"已生成: {path}")
    print(f"回写 {len(imp)} 个商品；跳过 {skipped} 个(近期无采购记录，FS/Remark 原样不动)")
    print("上传前请先看「对照」sheet 复核；Odoo 导入界面选「导入」sheet。")


if __name__ == "__main__":
    main()
