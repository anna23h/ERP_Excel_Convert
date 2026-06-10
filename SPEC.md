# SPEC — VO 拉单 Excel 自动化（开发契约）

来源流程：Obsidian `VO拉单流程逻辑梳理.md`。本文件锁定脚本的输入/输出/规则，开发以此为准。

## 输入

| 名称 | 文件示例 | 关键列 |
|---|---|---|
| ERP 导出 | `测试0611erp导出.xlsx` / `sale.order.csv`（16列） | `Order Reference`, `VO Tracking No`, `Order Lines/Product/Internal Reference`, `Order Lines/Product/Picking Name`, `Order Lines/Quantity`, `VO Delivery Type`, `Order Lines/Product/Barcode`, `Terms and conditions`, `Order Lines/Product/Quantity On Hand` |
| 天猫全量导出 | `天猫测试.xlsx`（sheet `file`，66列） | `系统履约单号`, `履约单状态`, `面单申请状态`（**按列名取，不靠列位置**） |
| 昨日发货表 | `1006发货表.xlsx`（sheet `GW`/`VO`） | `Order Reference`, `VO Tracking No` |

## 连接键
ERP `Order Reference` 取**后15位**（`right(,15)`，形如 `SCP...`）== 天猫 `系统履约单号`。实测 255/255 命中。

## 状态取值（天猫）
- `履约单状态`：`履约完成(已收货)` / `履约取消` / `平台申请取消` / `已发货` / `新订单` / `商家已接单` / `发货后取消(系统取消)`
- `面单申请状态`：`已完成` / `待处理` / `生成中`

## 输出表（A/B/C/D；E 出库表暂缓）

### A. 「拣货表+面单」workbook
- **拣货表** sheet（扁平单层，英文列名，一行一个 SKU）：
  `Internal Reference | Picking Name | Barcode | Quantity(求和) | Quantity On Hand(最大值)`
  实测每个 SKU 唯一对应一个 Picking Name/Barcode，扁平不丢数据；异常需报出。
- **面单** sheet（订单明细，6列 + 末尾空白「仓库备注」列）：
  `Order Reference | VO Tracking No | Internal Reference | Picking Name | Quantity | VO Delivery Type`
  脚本只产出**「待」全集**（出/无 拆分等仓库捡货后人工反馈）。
  标黄规则（满足任一）：
  1. `Internal Reference` 以 `x2` 结尾（套装，忽略大小写，结尾锚定）
  2. `Quantity` > 1
  3. `VO Delivery Type` == `CC`
  多品订单：A/B/F 列合并单元格。

### B. 系统履约单号表（上传天猫）
- 单列 `系统履约单号`（= Order Reference 后15位），来源 = 实际发货订单。

### C. 发货表
- 两个 sheet `GW` / `VO`，各2列：`Order Reference | VO Tracking No`。

### D. ERP 开账单上传表
- `Order Date | ID | Order Reference | Terms and conditions`
- `Terms and conditions` = `账单MMDD` + 原值（渠道+运单）。标签统一英文，去空列、去重复列。

## 全局规则
- 标签统一英文优先（`账单MMDD`，废弃 `开发票` 旧叫法）。
- 去冗余：统一拣货表格式、删空列、删重复列（如 `履约单号(文本)`）。
- 跨多个文件选格式时**先问用户**再定。
- 异常（如 delivery type 为空、SKU 多 Barcode、连接键未命中）一律单独列出给用户，不静默丢弃。

## 范围外
- 步骤1(VO API)、2(Odoo导出/确认)、5(回传ERP取消/开账单写操作)、3/4.4/4.8(天猫后台)、10(运单下载)：保持人工。
- E 出库表：导出/上传机制未确认，暂缓。
