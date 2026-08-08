# 采购数量与频次（po_frequency）

把某个供应商的 `purchase.order` 导出，整理成一张能看出**「采购了哪些产品、各买了几次、每次多少」**的表。**纯数据整理，不下任何结论**——卖穿率、价格趋势、客户维度都不在这里（那些是 2026-08-08 会话延伸出去的分析，见 `docs/journal/2026-08-08.md`；本工具刻意只回答同事最初的原始需求）。

与 `po_reconcile/` 的区别：那个是**对账**（采购 PO ↔ 财务 PO，算还有多少没到货）；本工具只做**频次/数量统计**，不碰到货、不回写。

## 用法

```bash
python3 po_frequency/po_frequency.py <purchase.order.xlsx> [--vendor NAME] [--out PATH]
```

- `--vendor NAME`：按供应商名过滤（**子串、忽略大小写**）。导出通常已在 ERP 里筛到单一供应商，此时可省略；给了就再过滤一次，也决定默认输出文件名里的供应商段。
- `--out PATH`：自定义输出路径。默认 `output/<Vendor>_Purchase_Quantity_and_Frequency.xlsx`（英文名，落 `output/`，已 gitignore）。同名自动加序号，不覆盖。

## 输入

`purchase.order` 的**行式导出**（Order Lines 粒度，一行一个明细）。Odoo 这种导出里订单头字段只出现在每单**第一行**、下面留空——**脚本自动向下填充，导出时不用手动补**。

必需列（缺任一会报 `采购导出缺列` 并中止）：

| 列 | 层级 | 用途 |
|---|---|---|
| `Order Reference` | 订单头 | 频次 = 不同单号去重计数 + 「每次采购」分组 |
| `Confirmation Date` | 订单头 | 首/末采购、跨度、间隔、明细日期 |
| `Order Lines/Product/Internal Reference` | 明细 | 产品关联键 |
| `Order Lines/Product/Display Name` | 明细 | 产品可读名 |
| `Order Lines/Total Quantity` | 明细 | 数量 |
| `Order Lines/Unit Price` | 明细 | 单价（仅 Details 用） |
| `Vendor` | 订单头 | **仅 `--vendor` 时必需**（用于过滤/命名） |

**筛选建议**（在 ERP 导出前做）：Vendor = 目标供应商；只要已确认的采购单（用 Purchase Orders 而非 RFQ 询价，排除草稿/已取消）；时间范围随意，工具「导出里有多少算多少」，不写死窗口。

## 产出

一个 xlsx，两 sheet，**英文表头**（同事不懂中文）：

- **Summary**：每产品一行——`Purchase Count (Frequency)` / `Total Qty` / `Avg·Min·Max per Purchase` / `First·Last Purchase` / `Span (days)` / `Avg Interval (days)`，按频次降序。
- **Details**：每明细行——`Order Date` / `PO Number` / `Product` / `Internal Reference` / `Qty` / `Unit Price €`，按产品+日期排序。

## 口径（三条固定规则，保证可复现）

1. **订单头字段向下 ffill**（补齐导出里空着的头字段）。
2. **SKU 归一**：复用 `common/po.py::_po_base_sku`，去掉 `x2 / *2 / _VO` 尾缀，令同一产品的多件装/渠道变体并到一行。
3. **频次 = 不同 `Order Reference` 去重计数**：同一张采购单里同一产品拆成多行，只算 **1 次**采购（该单内多行数量先合并，再计入 Avg/Min/Max per Purchase）。因此 Details 的行数可能略多于 Summary 的频次之和——这是 ERP 拆行造成的，不是错。

## 测试

```bash
python3 po_frequency/test_po_frequency.py     # 或 python3 -m pytest po_frequency/
```

合成锯齿数据，覆盖 ffill、SKU 归一、频次去重、`--vendor` 过滤、缺列报错。
