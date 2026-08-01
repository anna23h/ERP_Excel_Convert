"""ERP `Supply Remark` 字段的分段维护。

这个字段是多方共用的自由文本：运营手写的注意事项、`fs_writeback` 写的供应商画像、
`sales_insight` 写的安全库存备注，都往同一格里塞。约定用 `；` 分段，每个脚本写的段
以 `YYYYMMDD:<自己的前缀>` 开头，重跑时**按签名认出自己的旧段替换掉**，别人的段和
人工原文一律保留。

⚠ 判定必须**逐段**做，不能锚在整串开头（`^\\d{8}:…`）：一旦另一个脚本往前面插了段，
锚定写法就再也匹配不到自己的旧段，于是每跑一次堆一段。2026-08-01 加 sales_insight
的备注回写时发现 fs_writeback 原来就是锚定写法，两边一起改成了逐段判定。

已知残留风险：人工备注里若出现 `；` 会被误切成两段。实测样本用的都是 `，`，暂接受。
"""
import re

SEP = "；"


def merge(old, new_seg, sig):
    """把 new_seg 前置进 old，并剔除 old 里由本脚本写的旧段。

    old:     ERP 现值（可能为空/NaN）
    new_seg: 本次要写的段，形如 `20260801:安全库存 已下架…`；给空表示本次没有内容
    sig:     认自己旧段的正则（compiled），逐段 match

    没有新段、且旧值里也没有自己的段时**原样返回**——不重排、不改标点，
    保证「本次不涉及的行」写回 ERP 是真正的无变化。
    """
    old = "" if old is None or old != old else str(old)   # NaN 判定：不 import pandas
    segs = [s.strip() for s in old.split(SEP) if s.strip()]
    mine = [s for s in segs if sig.match(s)]
    if not new_seg and not mine:
        return old
    kept = [s for s in segs if not sig.match(s)]
    return SEP.join(([new_seg] if new_seg else []) + kept)


def signature(prefix):
    """`20260801:安全库存 …` 这类段的识别正则。prefix 会被转义。"""
    return re.compile(r"^\d{8}:" + re.escape(prefix))
