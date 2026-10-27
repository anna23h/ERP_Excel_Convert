---
name: PO导入表
description: 把批发商的 Rechnung/Lieferschein PDF 填进 Odoo 导出的空 purchase.order 模板、出可直接导入的 PO 表时用。用户给出一张 Rechnung / Lieferschein / 供货商账单 PDF 加一张 purchase.order*.xlsx，说「填进 PO」「做 PO 导入表」「账单转采购单」「导入成 PO」时触发。流程：跑 po_import/po_from_pdf.py → 停下就把原因/异常行交用户裁决 → 带 --exclude / --pick 重跑 → 交核对表与合计核对。PZN 是唯一键，绝不按品名兜底、不替用户选。
---

# PO 导入表（Rechnung/Lieferschein PDF → Odoo 采购单导入）

规则已于 2026-10-20 与用户定好，**写在脚本里**，不必再复述需求或逐条征求意见。
来龙去脉见私有库 `ISSUES.md` 的 `[PO 导入] A`；与 `SO导入表` skill 同构。

## 第一步：直接跑

```bash
python3 -m po_import.po_from_pdf <Lieferschein.pdf> <purchase.order 模板.xlsx>
```

（Windows 用 `py -3`。）依赖 `pdfplumber`：缺的话 `pip install -r requirements-pdf.txt`。
匹配和防重走 Odoo API（只读，凭据同 `config.py` 的 `ODOO`）。产出默认落
`output/YYYYMMDD/{PO号}_import_{PDF名}.xlsx`，同名自动加序号。

用户只给了 PDF、还没有模板时：请他在 Odoo 建好空 PO（选好供应商）后导出一张，
字段勾 `Order Reference` / `Vendor` / `Vendor Reference` / `Order Lines/Product/ID` / `Quantity` / `Unit Price` / `Taxes`。
数量要 `Quantity`，**不是** `Total Quantity`（只读计算字段，导进去无效；脚本遇到会停并说明）。
要勾「兼容导入的导出」，这样才带 PO 的 `ID` 列（没有它导入会新建 PO）。表头是界面名还是技术名（`id` / `order_line/product_qty` …）
都行，产品列是 `.id`（数据库 ID）也行，脚本自己认。不要拿「产品ID总表」之类的产品快照来补 ID——要的是 PO 的 ID。

模板里有多张空白 PO 时**不用问配对**：不给 `--po`，每张 PDF 各跑一次，脚本自动写进空白 PO 并打印「自动分配：P…」，交付时报给用户（2026-10-22 用户定）。

核对表后若有「默认进项税与 PDF 税码不一致」提示：表里已按 PDF 写，转告用户去核那几个产品的主数据即可，不拦出表。

## 第二步：退出码 2 = 停下交给用户

| 情况 | 交给用户的问题 |
|---|---|
| 某 PZN 在 Odoo 查不到 | 先去建档后重跑，还是这次剔掉这行？（0 元附件行如 `BEIPACK` 常见，2026-10-21 用户选剔掉） |
| 某 PZN 只有渠道 / 组合装副本、没有基础品 | 先在 Odoo 建基础品后重跑，还是这次剔掉这行？（**不能**改用副本） |
| 排除副本后仍有多个基础品 | 用哪个外部 ID？（清单里带 Internal Reference 和历史采购行数） |
| PDF 是 Gutschrift（贷项 / 退货） | 这张是退货，请在 Odoo 走退货 / 供应商贷项，不做 PO 导入表 |
| Odoo 已有同单号（Vendor Reference）的 PO | 这张是不是已经录过？ |
| 解析 0 行 / 无单号 / 无合计 / 任一合计或税基对不上 | PDF 版式不认识——**如实说**，不要自己改正则硬凑 |
| 模板缺列（尤其税列）、第 2 行头信息不全 | 让用户重新导出模板 |
| 模板里有多张 PO、没有可用的空白 PO | 列出的那几张为何不空白（已有行 / 已确认 / 已填 Vendor Reference）——要不要重新导出模板，或用 `--po` 指定？ |

**采购只用基础品**（2026-10-20 用户定）：带 `_VO` / `_GW` / `xN` / `*N` 尾缀的渠道 / 组合装副本脚本已自动排除，
核对表后列出排除了哪些，交付时一并报给用户。剩下的多基础品（常见新旧编码并存，如 `DOPH_001327` vs `DOPH_06445453`）
**附上你的看法可以**（「看历史采购应是 X」），**但决定必须等用户说**。拿到答复后：

```bash
python3 -m po_import.po_from_pdf L.pdf T.xlsx --exclude <PZN…> --pick <PZN=外部ID…>
```

## 第三步：交付

1. 产出路径
2. **核对表**整张贴给用户（PZN / PDF 名称 / 数量 / AEP / Gesamt / 税 / ERP 品名）
3. 剔除了哪些行（若有），提醒建档后手补进这张 PO
4. 自动排除的副本清单（若有）
5. **合计核对**：净额与两档税基（应等于 PDF 印的 volle / ermäß. MwSt），含税估算
6. 若模板没有 `Vendor Reference` 列，提醒单号没写进去、导入后手填（防重靠它）

## 禁止项

- 不用品名做模糊匹配兜底；不替用户选多命中、不自行剔除行。
- 不绕过护栏另写脚本出表。
- 单价就是 AEP（C 家 Ek.Preis）净价，不乘 1.19 / 1.07；税按各版式自己的标记分 7% / 19%（见下）。
- 认两家版式（A：`Code` 含 `M` = 7%；B：表头 `Menge Einheit Artikelbezeichnung PZN Symbol`，`Symbol` 含 `*` = 7%，
  B 的 `M` 是 phonetischer Artikel、**不是**减税），脚本按表头自动分流；C：门户 Lieferschein 网页存的 `.html`（按后缀分流，
  单价 = Ek.Preis，税看 MwSt 列 7% / 19%，`Kz` 与税无关；只给截图时请用户在页面 Ctrl+S 存 HTML，不靠识图转录）。
  再有新版式先告诉用户，按项目规则立项后再扩展。
