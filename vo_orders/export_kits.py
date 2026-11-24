#!/usr/bin/env python3
"""导出套件对照表（BoM）给同事用（ISSUES: [拣货面单] S）。

阶段一面单按货聚合要把套件 SKU 归到基础品。同事机器不连 Odoo 时，把这里导出的
`套件对照表.csv` 放到程序旁（源码版放仓库根，VOTool.exe 放 exe 同目录），程序就只读它、不连 Odoo。

Odoo 只读：全部 phantom BoM 里**单组件**的套件 → 列 套件SKU / 基础品SKU / 系数 / 导出日期；
多组件套件不归并，只在终端点名。CSV 已被 .gitignore 兜底，不会进公开仓库。

用法：python vo_orders/export_kits.py [-o 输出路径]   （默认 output/套件对照表.csv）
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import safety_dynamic as SD  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description="导出套件对照表（Odoo 只读）给同事离线用")
    ap.add_argument("-o", "--out", default=os.path.join("output", SD.KITS_FILE),
                    help=f"输出路径，默认 output/{SD.KITS_FILE}（文件名要保持 {SD.KITS_FILE} 程序才找得到）")
    a = ap.parse_args()
    n, multi, dup = SD.export_kits(a.out)
    print(f"已导出 {n} 个单组件套件 → {a.out}")
    if multi:
        print(f"⚠ 多组件套件 {len(multi)} 个不归并：{', '.join(multi)}")
    if dup:
        print(f"⚠ {dup} 个 Internal Reference 在 Odoo 有重码，取最早建档的")
    if os.path.basename(a.out) != SD.KITS_FILE:
        print(f"⚠ 文件名不是 {SD.KITS_FILE}，发给同事前要改回这个名字")
    print("发给同事：源码版放仓库根，VOTool.exe 放 exe 同目录。有这个文件时阶段一不再连 Odoo 拉套件。")


if __name__ == "__main__":
    main()
