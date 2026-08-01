#!/usr/bin/env python3
"""FS 回写：从 purchase order 提炼供应商画像，生成 Odoo 导入文件。

用法:
    python3 fs_writeback.py <purchase.order.xlsx> <product.product.xlsx> [outdir]
    python3 fs_writeback.py ... --sample 24     # 首次导入试水：按覆盖面挑 24 行

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
个人月频维护工具。**不进 VOTool**(同事无入口即不会误触 ERP 回写)——
2026-08-01 起有界面了，但是单独的 `erp_writeback_gui.py`，那条隔离仍然成立。

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
# 2026-08-01 起不再 import build_excel——采购画像已搬去 common/po，
# 于是本模块只依赖 common/，调用方(erp_writeback_gui)不必再把 vo_orders/ 塞进 sys.path。
from common.xlsx import unique_path  # noqa: E402
from common import vendor as vd  # noqa: E402
from common.po import load_po_stats, _po_base_sku  # noqa: E402

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
    stats, info = load_po_stats(po_path)
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
        base = _po_base_sku(p["Internal Reference"])
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
        chk.append({"id": p["External ID"],
                    "Internal Reference": p["Internal Reference"],
                    "处理": act,
                    "FS 旧": fs_old,
                    "FS 新": fs_new if act != "跳过(FS像人写的)" else "",
                    # 只读上下文，不回写——列名点明，免得看表的人以为它会被写进去
                    "Supply Remark 现值(不回写)":
                        p["Supply Remark"] if pd.notna(p["Supply Remark"]) else ""})
    return pd.DataFrame(imp), pd.DataFrame(chk), n, info


def _shape(fs):
    """FS 新值的形态：几段、是否含真实供应商名。取样按它铺开覆盖面。"""
    ts = [t.strip() for t in str(fs).split("/") if t.strip()]
    kind = "含真名" if any(" " in t or len(t) > 6 for t in ts) else "纯代号"
    return f"{kind}({len(ts)}段)"


def sample(imp, chk, k):
    """从全量里挑 k 行做首次导入试水 → (导入df, 对照df)。

    **不是切前 k 行**——前 k 行大概率全是同一种情形(单代号新增)，验不到覆盖、
    多代号、含真名。改为按覆盖面取样：
      · 批内每行 FS 新值互不相同(用户 2026-08-01 要求"FS 值各不同")
      · 轮流从每个「形态 × 处理」格子里取，稀有格子先取，保证都露面
      · 同等条件下优先挑 FS 旧值没出现过的，让覆盖情形也铺开
    """
    pool = chk[chk["处理"].isin(["新增", "覆盖"])].copy()
    pool["_形态"] = pool["FS 新"].map(_shape)
    # 稀有格子排前面：轮流取时它们先被满足，不会被大格子挤掉
    cells = sorted(pool.groupby(["_形态", "处理"]).groups.items(), key=lambda kv: len(kv[1]))
    buckets = [list(pool.loc[idx].to_dict("records")) for _, idx in cells]

    picked, seen_new, seen_old = [], set(), set()
    while len(picked) < k and any(buckets):
        for b in buckets:
            if len(picked) >= k:
                break
            # 先要「新旧都没见过」的，退而求其次只要新值没见过
            cand = next((r for r in b if r["FS 新"] not in seen_new
                         and r["FS 旧"] not in seen_old), None) \
                or next((r for r in b if r["FS 新"] not in seen_new), None)
            if cand is None:
                b.clear()
                continue
            b.remove(cand)
            picked.append(cand)
            seen_new.add(cand["FS 新"])
            seen_old.add(cand["FS 旧"])

    s_chk = pd.DataFrame(picked).drop(columns=["_形态"])
    s_imp = imp[imp["id"].isin(set(s_chk["id"]))]
    return s_imp, s_chk


def run(po_path, prod_path, outdir="output", k=0):
    """跑一次回写 → (落盘路径, 摘要行列表)。CLI 与 GUI 共用一份逻辑与摘要。

    k>0 时只出试水样本（见 sample()）。
    出错抛 ValueError 而非 SystemExit——后者 GUI 后台线程的 except Exception 抓不到。
    """
    L = []
    say = L.append
    os.makedirs(outdir, exist_ok=True)
    imp, chk, n, info = make_writeback(po_path, prod_path)
    d = date.today().strftime("%Y%m%d")
    full = len(imp)
    if k:
        imp, chk = sample(imp, chk, k)
    tag = f"-试{len(imp)}" if k else ""
    path = unique_path(os.path.join(outdir, f"FS回写导入 {d}{tag}.xlsx"))
    with pd.ExcelWriter(path) as xw:
        imp.to_excel(xw, sheet_name="导入", index=False)
        chk.to_excel(xw, sheet_name="对照", index=False)
    say(f"采购参考: {info}")
    say(f"已生成: {path}")
    if k:
        c = chk["处理"].value_counts()
        say(f"🧪 首次导入试水样本 {len(imp)} 行（全量 {full} 行，本次**只导这批**）")
        say(f"   新增 {c.get('新增', 0)} / 覆盖 {c.get('覆盖', 0)}；"
            f"FS 新值 {chk['FS 新'].nunique()} 种互不相同、FS 旧值 {chk['FS 旧'].nunique()} 种")
        say(f"   形态覆盖: {' / '.join(sorted(set(chk['FS 新'].map(_shape))))}")
        if len(imp) < k:
            say(f"   （你要 {k} 行，只给出 {len(imp)} —— 「FS 值各不同」是硬约束，"
                f"全量里不同的 FS 值就这么多）")
        say("   导完回 ERP 抽查几行，确认代号与覆盖行为都对，再跑全量（试水行数填 0）。\n")
    say(f"{'全量口径: ' if k else ''}回写 {full} 个商品的 FS "
        f"= 新增 {n['新增']} + 覆盖 {n['覆盖']}")
    say(f"  · 跳过 {n['跳过(FS像人写的)']} 个：FS 现值像人写的采购判断，不拿聚合结果盖掉")
    say(f"  · 跳过 {n['跳过(费用类SKU)']} 个：费用/耗材类 SKU，不是进货商品")
    say(f"  · 跳过 {n['无采购记录']} 个：近期无采购记录，FS 原样不动")
    if not vd.VENDOR_ALIAS:
        say("  ⚠ 没读到供应商代号对照(config.py 的 VENDOR_ALIAS)，FS 会写成供应商真名")
    say("· 本脚本**只写 FS**，不碰 Supply Remark——那个字段留给运营同事自己维护。")
    say("  「对照」sheet 里的 Supply Remark 是现值，只作复核上下文，不会被写回。")
    say("上传前请先看「对照」sheet 复核(含被跳过的行及原因)；Odoo 导入界面选「导入」sheet。")
    return path, L


def main():
    args = sys.argv[1:]
    k = 0
    if "--sample" in args:                      # 位置参数保持 po/prod/[outdir] 向后兼容
        i = args.index("--sample")
        k = int(args[i + 1])
        args = args[:i] + args[i + 2:]
    if len(args) < 2:
        print(__doc__)
        return
    outdir = args[2] if len(args) > 2 else "output"
    try:
        _, lines = run(args[0], args[1], outdir, k)
    except ValueError as e:
        raise SystemExit(str(e))
    print("\n".join(lines))


if __name__ == "__main__":
    main()
