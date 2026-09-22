"""供应商全名 → 简称（原先长在 vo_orders/build_excel.py 里，reorder 一直在跨模块掏它）。"""
import re

VENDOR_LEGAL = {"gmbh", "gmbh,", "kg", "kgaa", "ag", "ohg", "mbh", "mbb", "co", "co.",
                "&", "e.k.", "ek", "e.u", "ltd", "ltd.", "limited", "inc", "inc.",
                "s.a.r.l.,", "s.a.r.l.", "sarl", "sas", "bv", "se",
                "niederlassung", "deutschland", "holding"}
# 首词全大写但属行业通用词，单独指代会误导(PHARMA LUPUS ≠ "PHARMA")
VENDOR_GENERIC = {"PHARMA", "APOTHEKE", "MED"}

# 采购单里不是真实进货的对手方，整行剔除——不然会污染采购画像与「最低价」。
# **全仓唯一出处**（2026-09-22 收敛，见 ISSUES [采购缺口] J）：common/po.load_po_stats、
# procure/po_price.py、procure/gap_report.py 共用这一份，不要再各自抄一份字面量。
# 此前是三处各写各的、注释还互相声称「同 X 的口径」，却没有任何机制保证——已实际漂开过
# （reorder 那份放宽了，这三份没跟上），代价见下。改这里之前先想清楚它影响的是全部三条线。
#   Alibaba —— 伪装成供应商的**我方客户**，名下是退货包裹单，单价恒为 0。
#     **只写 "Alibaba" 不写全名**：2026-09-22 实测 Odoo 里名字含 alibaba 的 partner 有
#     **四个**（港、新加坡各自还不止一条记录，另有一条内部对照单），写全名注定要漏。
#     四个的具体名字不写进公开库，见 ISSUES [采购缺口] J。
#     当天实测：旧口径漏掉的那家在近 12 个月窗口内有 54 行 purchase.order.line
#     （采购单 P10504 / P10505，共 90 件）单价全为 0，把 **39 个 SKU 的最低价拉成 0**，
#     另有 7 个 SKU（如 Avene_13883685x2）窗口内只有这条噪声行、剔后再无采购行。
#     放宽成子串后，同族再冒出第五个法人实体也不必改代码。
#     （reorder/reorder_helper.py 的 CUSTOMER_PAT 早就放宽成 "alibaba" 了，却没扩散到
#      另外三处——那正是本常量要收敛成一份的理由。它口径不同，暂不并入，见 ISSUES J。）
#   VO Test Order —— 建虚拟库存映射的测试单。
#     2026-08-01 实测不滤的话有 397 个商品的 FS 会被写成 "VO"。
NOISE_VENDOR_PATS = ["Alibaba", "VO Test Order"]

# 上面那份名单的正则形态：**子串、大小写不敏感**。这是实现细节，
# **按行过滤请用下面的 is_noise_vendor()**，不要 import 这个正则、更不要自己再 compile 一份
# —— 「各自 compile 一份」正是当年几份名单能漂开的入口。
NOISE_VENDOR_RE = re.compile("|".join(re.escape(p) for p in NOISE_VENDOR_PATS), re.I)


def is_noise_vendor(name):
    """供应商名是否属于「不是真实进货的对手方」(见 NOISE_VENDOR_PATS)。None/空 → False。"""
    return bool(name) and bool(NOISE_VENDOR_RE.search(str(name)))


# 个别简称覆盖：规则产物 → 最终代号。
# ⚠ **对照表本身不进公开库**，放 config.py(已 gitignore，走 Syncthing 两台 Mac 同步)。
# 用户要代号(P/G/A/B…)正是为了不点名供应商；把 {"某批发商": "P"} 写在这里等于把
# 「P 是谁」发到 GitHub，既抵消目的又违反本项目 CLAUDE.md「公开仓库勿写真实供应商」。
# 没有 config.py 或没配这项时为空 dict——规则产出原样使用，不报错(2026-08-01)。
VENDOR_ALIAS = {}


def _load_alias():
    """从仓库根的 config.py 读 VENDOR_ALIAS（加载细节见 common/localconf.py）。"""
    from common import localconf
    return dict(localconf.get("VENDOR_ALIAS", {}) or {})


VENDOR_ALIAS = _load_alias()


def short_vendor(name):
    """供应商全名 → 简称(2026-07-07 全量 65 家实测零碰撞)：
    去括号注记 → 滤法律/地名后缀 → 首词全大写(≥2字符、连字符取头段、非通用词)
    则单词指代(PHOENIX/AEP/GEHE/DM/UPS)，否则取前两词；结果过短再多取一词。"""
    s = re.sub(r"[（(].*?[)）]", "", str(name)).strip()
    words = [w for w in s.split() if w.lower() not in VENDOR_LEGAL]
    if not words:
        return str(name).strip()
    head = words[0].split("-")[0]
    if head.isupper() and len(head) >= 2 and head not in VENDOR_GENERIC:
        return VENDOR_ALIAS.get(head, head)
    n = 2 if len(" ".join(words[:2])) >= 4 else 3
    res = " ".join(words[:n])
    return VENDOR_ALIAS.get(res, res)


def vendor_map(vendors):
    """全名→简称映射；不同全名缩成同一简称(前瞻防护，当前数据零碰撞)则保留全名。"""
    m = {v: short_vendor(v) for v in vendors}
    dup = {s for s in m.values() if list(m.values()).count(s) > 1}
    return {v: (v if s in dup else s) for v, s in m.items()}
