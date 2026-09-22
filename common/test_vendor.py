"""common/vendor 的「非进货对手方」名单测试。

跑：python3 common/test_vendor.py    （或 python3 -m pytest common/test_vendor.py）

背景：2026-09-22 之前这份名单在仓库里有四份各自维护的拷贝，`reorder` 那份早就放宽成
`"alibaba"`，另外三份仍写死 `"Alibaba Health"`，于是漏掉了同一客户的新加坡法人实体
（17 行 0 元退货行把 12 个 SKU 的最低价拉成 0）。收敛成一份之后，用这个用例把
「按子串、大小写不敏感、覆盖同族所有法人实体」钉住，防第三次漂移
（ISSUES [采购缺口] J）。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from common.po import PO_NOISE_PATS  # noqa: E402
from common.vendor import NOISE_VENDOR_PATS, is_noise_vendor  # noqa: E402


def test_noise_vendor():
    noise = [
        # 我方客户伪装成供应商：同族至少两个法人实体，名字完全对不上，只能按子串认
        "Alibaba Health",
        "Alibaba.com Singapore E-Commerce Private Limited",
        "ALIBABA.COM SINGAPORE E-COMMERCE PRIVATE LIMITED",   # 大小写不敏感
        "alibaba health technology",
        # 建虚拟库存映射的测试单
        "VO Test Order",
        "vo test order",
    ]
    keep = [
        # 真实供应商：名字里带 Health / Ali 片段，不能被误杀
        "GEHE Alliance Healthcare Deutschland GmbH",
        "PHOENIX Pharmahandel GmbH & Co KG",
        "Ali Import GmbH",
        "Baba Pharma GmbH",
        "Alliance Healthcare",
    ]
    bad = [n for n in noise if not is_noise_vendor(n)]
    assert not bad, f"该剔除却放过了: {bad}"
    bad = [n for n in keep if is_noise_vendor(n)]
    assert not bad, f"真实供应商被误杀: {bad}"
    # 空值不该炸，也不该算噪音
    assert not is_noise_vendor(None)
    assert not is_noise_vendor("")


def test_single_source():
    """名单只有一份：common/po 的 PO_NOISE_PATS 是别名，不是第二份拷贝。"""
    assert PO_NOISE_PATS is NOISE_VENDOR_PATS


if __name__ == "__main__":
    test_noise_vendor()
    test_single_source()
    print("all tests passed")
