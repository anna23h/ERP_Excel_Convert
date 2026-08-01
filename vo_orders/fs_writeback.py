#!/usr/bin/env python3
"""FS 回写：从 purchase order 提炼供应商画像，生成 Odoo 导入文件。

用法:
    python3 fs_writeback.py <purchase.order.xlsx> <product.product.xlsx> [outdir]

范围: product 导出里有近期采购记录的商品；无采购记录的整行跳过不动。
策略: FS = 近期供应商代号列表("/"分隔，对齐公司现有 "DM/ROSSMANN" 记法)。
      代号来自 config.py 的 VENDOR_ALIAS(不进公开库)——按真实供应商区分到家，
      但写进 ERP 的是代号，不点名(2026-08-01 用户要求，防信息泄露)。
      覆盖既有值，**但 FS 现值看着像人写的采购判断时整行跳过**——
      "首选AEP 不在Phoenix订"、"2026年7月2日MHD原因暂时停止订货" 这类是画像给不了的
      业务经验，机器不该拿聚合结果盖掉(见 _looks_human)。
产出: 单文件两 sheet——「导入」(id/FS，直接上 Odoo 导入界面) +
      「对照」(全部有采购记录的商品 + 处理/FS新旧 + Supply Remark 现值作只读上下文)。
上传保持人工。product 导出本身即现值备份，请保留原文件。
个人月频维护工具，不进 GUI(同事无入口即不会误触 ERP 回写)。

⚠ **本脚本不写 `Supply Remark`，那个字段属于运营同事**——留给他们自己修改/添加，
机器不占位。2026-07-08 建这个脚本时曾往里前置带日期的采购画像段，做得很克制
(原文保留、按签名替换、重跑不堆叠)，但 2026-08-01 用户复盘时点明：追加式不堆叠
解决的是"怎么写才不弄坏别人的东西"，而正确答案是**不写**。已整段删除。
「导入」sheet 里没有这一列，Odoo 就绝不会碰这个字段——别再"顺手写一下"。
(那些画像段从未真正上传过 ERP：2026-08-01 查 10331 行产品导出，含画像段的 0 条。)
"""
import sys, os, re
from datetime import date

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # 让 common/ 可导入
from common.xlsx import unique_path  # noqa: E402
from common import vendor as vd  # noqa: E402
import build_excel as be  # noqa: E402

# Supply Remark 只用于「对照」sheet 的只读上下文(复核 FS 该不该覆盖时，
# 看一眼这货是不是"停产"很有用)，不是回写目标。
PROD_NEED = ["Internal Reference", "FS", "Supply Remark", "External ID"]

# FS 现值里"看着像人写的采购判断"的特征。命中则整行跳过，不覆盖。
# 实测(2026-08-01，2334 个有 FS 的商品)恰好命中 7 种取值 / 15 个商品，48 种编码值零误伤：
#   首选AEP 不在Phoenix订 / 不在AEP订 / 2026年7月2日MHD原因暂时停止订货 /
#   2025年3月24日停产 / 大药房，处方药，慎重采购库存 / 熊猫家MHD27 / Y2024 B2B Season Price
_HUMAN_KW = re.compile(r"[，。；]|\d{4}年|\d+月\d+日|MHD"
                       r"|不在|首选|慎重|暂时|停产|停止|订货|采购|原因|Price")

# 费用/耗材类 SKU，不是进货商品，FS 对它们没意义——整行不回写。
# 不滤的话 Service_Fee 的 FS 会被写成一串 50 个供应商(DHL/保险/餐厅/海关/加油站…)，
# Service_Insurance 会被写成保险公司名(2026-08-01 实测 31 个这类 SKU)。
_NON_GOODS = re.compile(r"^(Service|Consumable|Fee|Shipping)_", re.I)


def _looks_human(fs):
    """FS 现值是不是人写的判断（而非渠道代号）。

    **必须按 `/` 拆段逐段判**：FS 是斜杠分隔的多值字段，`DM/ROSSMANN` 这类纯编码整格
    有 11 字符，拿整格长度当阈值会把它当成人写的——实测那样误伤 463 个商品(2026-08-01)。
    """
    for t in str(fs).split("/"):
        t = t.strip()
        if not t:
            continue
        if _HUMAN_KW.search(t) or (" " in t and len(t) >= 8) or len(t) >= 12:
            return True
    return False


def _fs(st):
    """一条 stats 行 → FS 新值(近期供应商代号列表)。

    代号来自 common.vendor 的 VENDOR_ALIAS(落 config.py，不进公开库)——
    按真实供应商区分到家，但写进 ERP 的是代号，不点名。
    """
    vend_lines = str(st["供应商(次数)"]).split("\n")
    names = [v.rsplit("×", 1)[0] for v in vend_lines]
    return "/".join(names)


def make_writeback(po_path, prod_path):
    """→ (导入df[id/FS], 对照df, 计数dict, 采购窗口描述)。

    对照 df 含**全部有采购记录的商品**（跳过的也在里面，`处理` 列标明原因），
    这样"哪些没被回写、为什么"一眼可查，不用去翻终端。
    """
    stats, info = be.load_po_stats(po_path)
    s = stats.set_index("_sku")
    prod = pd.read_excel(prod_path, dtype=str)
    missing = [c for c in PROD_NEED if c not in prod.columns]
    if missing:
        raise ValueError("product 导出缺列: " + ", ".join(missing))
    if not prod["External ID"].is_unique:
        raise ValueError("product 导出 External ID 有重复，无法作导入键")
    imp, chk = [], []
    n = {"无采购记录": 0, "新增": 0, "覆盖": 0, "跳过(FS像人写的)": 0, "跳过(费用类SKU)": 0}
    for _, p in prod.iterrows():
        base = be._po_base_sku(p["Internal Reference"])
        if base not in s.index:
            n["无采购记录"] += 1
            continue
        if _NON_GOODS.match(str(p["Internal Reference"]).strip()):
            n["跳过(费用类SKU)"] += 1
            continue
        fs_old = p["FS"] if pd.notna(p["FS"]) else ""
        fs_new = _fs(s.loc[base])
        # 覆盖，但人写的采购判断("首选AEP 不在Phoenix订"这类)整行跳过——
        # 那是画像给不了的业务经验，机器不该拿聚合结果盖掉它(2026-08-01 用户拍板)
        if fs_old and _looks_human(fs_old):
            act = "跳过(FS像人写的)"
        else:
            act = "新增" if not fs_old else "覆盖"
            imp.append({"id": p["External ID"], "FS": fs_new})
        n[act] += 1
        chk.append({"Internal Reference": p["Internal Reference"],
                    "处理": act,
                    "FS 旧": fs_old,
                    "FS 新": fs_new if act != "跳过(FS像人写的)" else "",
                    # 只读上下文，不回写——列名点明，免得看表的人以为它会被写进去
                    "Supply Remark 现值(不回写)":
                        p["Supply Remark"] if pd.notna(p["Supply Remark"]) else ""})
    return pd.DataFrame(imp), pd.DataFrame(chk), n, info


def main():
    args = sys.argv[1:]
    if len(args) < 2:
        print(__doc__)
        return
    po_path, prod_path = args[0], args[1]
    outdir = args[2] if len(args) > 2 else "output"
    os.makedirs(outdir, exist_ok=True)
    imp, chk, n, info = make_writeback(po_path, prod_path)
    d = date.today().strftime("%Y%m%d")
    path = unique_path(os.path.join(outdir, f"FS回写导入 {d}.xlsx"))
    with pd.ExcelWriter(path) as xw:
        imp.to_excel(xw, sheet_name="导入", index=False)
        chk.to_excel(xw, sheet_name="对照", index=False)
    print(f"采购参考: {info}")
    print(f"已生成: {path}")
    print(f"回写 {len(imp)} 个商品的 FS = 新增 {n['新增']} + 覆盖 {n['覆盖']}")
    print(f"  · 跳过 {n['跳过(FS像人写的)']} 个：FS 现值像人写的采购判断，不拿聚合结果盖掉")
    print(f"  · 跳过 {n['跳过(费用类SKU)']} 个：费用/耗材类 SKU，不是进货商品")
    print(f"  · 跳过 {n['无采购记录']} 个：近期无采购记录，FS 原样不动")
    if not vd.VENDOR_ALIAS:
        print("  ⚠ 没读到供应商代号对照(config.py 的 VENDOR_ALIAS)，FS 会写成供应商真名")
    print("· 本脚本**只写 FS**，不碰 Supply Remark——那个字段留给运营同事自己维护。")
    print("  「对照」sheet 里的 Supply Remark 是现值，只作复核上下文，不会被写回。")
    print("上传前请先看「对照」sheet 复核(含被跳过的行及原因)；Odoo 导入界面选「导入」sheet。")


if __name__ == "__main__":
    main()
