#!/usr/bin/env python3
"""FS 回写：从 purchase order 提炼供应商画像，生成 Odoo 导入文件。

用法:
    python3 fs_writeback.py <purchase.order.xlsx> <product.product.xlsx> [outdir]

范围: product 导出里的商品(按 Safety Stock>0 筛的囤货品)；无采购记录的商品整行跳过不动。
策略: FS = 近期供应商短名列表(覆盖；"/"分隔，对齐公司现有 "DM/ROSSMANN" 记法)。
产出: 单文件两 sheet——「导入」(id/FS，直接上 Odoo 导入界面) +
      「对照」(FS 新旧对比 + Supply Remark 现值作只读上下文，上传前人工复核)。
上传保持人工。product 导出本身即现值备份，请保留原文件。
个人月频维护工具，不进 GUI(同事无入口即不会误触 ERP 回写)。

⚠ **本脚本不写 `Supply Remark`，那个字段属于运营同事**——留给他们自己修改/添加，
机器不占位。2026-07-08 建这个脚本时曾往里前置带日期的采购画像段，做得很克制
(原文保留、按签名替换、重跑不堆叠)，但 2026-08-01 用户复盘时点明：追加式不堆叠
解决的是"怎么写才不弄坏别人的东西"，而正确答案是**不写**。已整段删除。
「导入」sheet 里没有这一列，Odoo 就绝不会碰这个字段——别再"顺手写一下"。
(那些画像段从未真正上传过 ERP：2026-08-01 查 10331 行产品导出，含画像段的 0 条。)
"""
import sys, os
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # 让 common/ 可导入
from common.xlsx import unique_path  # noqa: E402
import build_excel as be  # noqa: E402

# Supply Remark 只用于「对照」shet 的只读上下文(复核 FS 该不该覆盖时，
# 看一眼这货是不是"停产"很有用)，不是回写目标。
PROD_NEED = ["Internal Reference", "FS", "Supply Remark", "External ID"]


def _fs(st):
    """一条 stats 行 → FS 新值(近期供应商短名列表)。"""
    vend_lines = str(st["供应商(次数)"]).split("\n")
    names = [v.rsplit("×", 1)[0] for v in vend_lines]
    return "/".join(names)


def make_writeback(po_path, prod_path):
    """→ (导入df[id/FS], 对照df, 跳过数, 采购窗口描述)。"""
    stats, info = be.load_po_stats(po_path)
    s = stats.set_index("_sku")
    prod = pd.read_excel(prod_path, dtype=str)
    missing = [c for c in PROD_NEED if c not in prod.columns]
    if missing:
        raise ValueError("product 导出缺列: " + ", ".join(missing))
    if not prod["External ID"].is_unique:
        raise ValueError("product 导出 External ID 有重复，无法作导入键")
    imp, chk, skipped = [], [], 0
    for _, p in prod.iterrows():
        base = be._po_base_sku(p["Internal Reference"])
        if base not in s.index:
            skipped += 1
            continue
        fs_new = _fs(s.loc[base])
        imp.append({"id": p["External ID"], "FS": fs_new})
        chk.append({"Internal Reference": p["Internal Reference"],
                    "FS 旧": p["FS"] if pd.notna(p["FS"]) else "",
                    "FS 新": fs_new,
                    # 只读上下文，不回写——列名点明，免得看表的人以为它会被写进去
                    "Supply Remark 现值(不回写)":
                        p["Supply Remark"] if pd.notna(p["Supply Remark"]) else ""})
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
    path = unique_path(os.path.join(outdir, f"FS回写导入 {d}.xlsx"))
    with pd.ExcelWriter(path) as xw:
        imp.to_excel(xw, sheet_name="导入", index=False)
        chk.to_excel(xw, sheet_name="对照", index=False)
    print(f"采购参考: {info}")
    print(f"已生成: {path}")
    print(f"回写 {len(imp)} 个商品的 FS；跳过 {skipped} 个(近期无采购记录，FS 原样不动)")
    print("· 本脚本**只写 FS**，不碰 Supply Remark——那个字段留给运营同事自己维护。")
    print("  「对照」sheet 里的 Supply Remark 是现值，只作复核上下文，不会被写回。")
    print("上传前请先看「对照」sheet 复核；Odoo 导入界面选「导入」sheet。")


if __name__ == "__main__":
    main()
