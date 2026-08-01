"""供应商全名 → 简称（原先长在 vo_orders/build_excel.py 里，reorder 一直在跨模块掏它）。"""
import re

VENDOR_LEGAL = {"gmbh", "gmbh,", "kg", "kgaa", "ag", "ohg", "mbh", "mbb", "co", "co.",
                "&", "e.k.", "ek", "e.u", "ltd", "ltd.", "limited", "inc", "inc.",
                "s.a.r.l.,", "s.a.r.l.", "sarl", "sas", "bv", "se",
                "niederlassung", "deutschland", "holding"}
# 首词全大写但属行业通用词，单独指代会误导(PHARMA LUPUS ≠ "PHARMA")
VENDOR_GENERIC = {"PHARMA", "APOTHEKE", "MED"}
# 个别简称覆盖(用户指定)：规则产物 → 最终简称
VENDOR_ALIAS = {"Dirk Rossmann": "Rossmann"}


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
