# SPEC — VO 拉单 Excel 自动化（开发契约）

来源流程：Obsidian `VO拉单流程逻辑梳理.md`。本文件锁定脚本的输入/输出/规则，开发以此为准。

## 输入

| 名称 | 文件示例 | 关键列 |
|---|---|---|
| ERP 导出（可多份 VO/GW，concat） | `VO0612ERP原始文档233单.xlsx`（17列含 `External ID`） | `Order Reference`, `VO Tracking No`, `Order Lines/Product/Internal Reference`, `Order Lines/Product/Picking Name`, `Order Lines/Quantity`, `VO Delivery Type`, `Order Lines/Product/Barcode`, `Terms and conditions`, `Order Lines/Product/Quantity On Hand`, `External ID` |
| 面单已完成名单（天猫筛选导出） | `0611天猫新订单+商家已接单+面单已完成.xlsx`（sheet `file`） | `系统履约单号`（= 今天有运单、可发货的订单集合） |
| 完整天猫导出（识别取消，选填） | `天猫测试.xlsx`（sheet `file`，66列） | `系统履约单号`, `履约单状态`（**按列名取**） |
| 昨日发货表 | `1006发货表.xlsx`（sheet `GW`/`VO`） | `Order Reference`, `VO Tracking No` |

## 连接键
ERP `Order Reference` 取**后15位**（`right(,15)`，形如 `SCP...`）== 天猫 `系统履约单号`。
注意：极少数情况两个不同 `Order Reference` 后15位相同（连接键冲突），脚本需报出（可能误判状态）。

## 分流（步骤4 核心，`step4_merge.classify4`）
按 ERP 订单 `_key` 打 `_cat`，优先级 取消 > 发货/已补运单 > 无运单：
- **取消**：`_key` ∈ 完整天猫导出中 `履约单状态 ∈ {履约取消, 平台申请取消}`。→ 取消单。
- **发货**：在「面单已完成名单」里。→ 进拣货/面单。
- **已补运单**：发货 且 ERP `Terms and conditions` 已含 `无运单`（昨日无运单今日补出）。→ 仍发货，并列入已补运单清单；账单 D 里剥掉 `无运单` 恢复原值。
- **无运单**：既不在名单也未取消。→ **从拣货/面单剔除**，列入无运单清单（Terms 前缀 `无运单`，已带则不重复加）。

## 状态取值（天猫）
- `履约单状态`：`履约完成(已收货)` / `履约取消` / `平台申请取消` / `已发货` / `新订单` / `商家已接单` / `发货后取消(系统取消)`
- `面单申请状态`：`已完成` / `待处理` / `生成中`（**判定发货改用「面单已完成名单」，而非此列**：名单已在天猫后台按 面单已完成 + 新订单/商家已接单 筛过，自动排除已发货老单与待处理/生成中）

## 输出表（A/B/C/D/E）

### A. 「拣货表+面单」workbook
- **拣货表** sheet（扁平单层，英文列名，一行一个 SKU）：
  `Internal Reference | Picking Name | Barcode | Quantity(求和) | Quantity On Hand(最大值)`
  实测每个 SKU 唯一对应一个 Picking Name/Barcode，扁平不丢数据；异常需报出。
- **面单** sheet（订单明细，6列 + 末尾空白「仓库备注」列）：
  `Order Reference | VO Tracking No | Internal Reference | Picking Name | Quantity | VO Delivery Type`
  只含**发货集合**（= classify4 的 发货+已补运单，已剔除无运单/取消）；出/无 拆分等仓库捡货后人工反馈。
  **按店铺(VO/GW)各出一份**（`YYYY年MM月DD日{ch}{n}单 拣货表+面单.xlsx`，序号店内独立 1..N）：仓库/打印按店进行。两店 ERP 可一次传入，自动分别出文件。
  标黄规则（满足任一）：
  1. `Internal Reference` 以 `x2` 结尾（套装，忽略大小写，结尾锚定）
  2. `Quantity` > 1
  3. `VO Delivery Type` == `CC`
  多品订单：A/B/F 列合并单元格。

### B. 系统履约单号表（上传天猫）
- 单列 `系统履约单号`（= Order Reference 后15位），来源 = 实际发货订单。

### C. 发货表
- 两个 sheet `GW` / `VO`，各2列：`Order Reference | VO Tracking No`。

### D. ERP 开账单上传表（账单上传.xlsx）
- `Order Date | ID | Order Reference | Terms and conditions`
- `Terms and conditions` = `账单MMDD` + 原值（渠道+运单）；去空列。
- `ID` = Odoo External ID，是导入更新订单的匹配键，无法凭空生成。
- **来源优先级**：订单导出(`--erp`)含 `External ID` 列时直接从它生成（推荐，一份文件搞定）；否则回落到单独的账单模板导出(`--billing`)。
- 去重到订单级（订单导出逐行，多品订单合一行，ID 取订单头行）。

### E. 出库单（回传 Odoo 标记出库）
- 输入：stock picking 全量导出「出库原始数据」（英文表头 7 列：`ID | Source Document | Reference | Creation Date | Carrier/ID | Tracking Reference | Status`）。`Source Document` 形如 `VO_TOF_SCP…`，可多份（VO/GW 各一份）。
- 过滤：`Source Document` 的 SCP ∈ **实际发货订单**（= 待全集 − 无货，与 B/C 同源 `get_shipped_orders`）。
- 改写：仅把 `Tracking Reference` 整列覆盖成**统一发货日期** YYYYMMDD（默认运行当天）；其余列沿用 pool 原值，**英文表头、保留原 Status**（不改 Bereit/Ready）。
- 拆分：按 `Source Document` 前缀拆成 `出库单VO.xlsx` / `出库单GW.xlsx`。
- 异常：对 pool 覆盖到的店铺，发货订单在 pool 里找不到对应 picking 的，单独列出 SCP 给用户（不静默）。

### F. 取消单 / 无运单清单（ERP 上传，步骤4.6/4.9，`build_excel`）
- 同为 ERP 上传表：`Order Date | <External ID> | Order Reference | Terms and conditions`，去重到订单级（`External ID` 是导入匹配键）。**两店合并一份**（`External ID` 唯一；ERP 上传按店人工分开时可按 `Order Reference` 前缀筛）。
- **取消单.xlsx**：`Terms` = `YYYY年MM月DD日平台订单取消`（覆盖）。无完整天猫导出则 0 单。
- **无运单清单.xlsx**：`Terms` = `无运单` + 原值（已带则不重复加）。

### G. 已补运单清单（步骤4.8，`build_excel`）
- 单列 `系统履约单号`（= 已补运单订单的 `_key`），拿去天猫后台批量打面单。两店合并一份（天猫后台不分店）。

## 阶段二运行粒度
阶段二(B/C/D/E + 缺货记录)**按店铺各跑一次**：单店 ERP + 该店仓库返回文件。
因回传 ERP 取消/无运单/账单需按店区分、人工分别上传，脚本不做**跨店**合并以避免复杂度与不稳定。
返回文件支持**多选作冗余**：同一店铺被分多次导出时合并为该店无货集合（不用于 VO/GW 跨店合并）。

## 为什么必须在脚本里筛/剔，而非只在天猫导出端筛（实测结论）
天猫导出端只能回答「谁有面单」，但拉单需要**发货/取消/无运单**三个桶，分桶依赖：① 与单店 ERP 批次取交集 ② ERP 的取消状态 ③ ERP 历史 Terms（已补运单）。这些信息分散在 ERP + 天猫全量，单次平台筛选表达不了。
实测（0611）：59 个「在 ERP、不在面单已完成名单」的单，真实状态是 6 取消 + 53 无运单**混在一起**——天猫的面单筛选只会一律排除，分不出该作废还是该暂扣。故必须用完整天猫导出在外部做差集分桶。
完整性保证：classify4 四桶穷尽互斥，每单必落且仅落一桶（实测 310 = 246+5+53+6，无遗漏）。

## 全局规则
- 标签统一英文优先（`账单MMDD`，废弃 `开发票` 旧叫法）。
- 取消两状态合并：`履约取消`/`平台申请取消` 同归「取消」，统一打 `平台订单取消`。`发货后取消(系统取消)` 暂不处理（用户定）。
- 去冗余：统一拣货表格式、删空列、删重复列（如 `履约单号(文本)`）。
- 跨多个文件选格式时**先问用户**再定。
- 异常（delivery type 为空、SKU 多 Barcode、连接键后15位冲突）一律单独列出，不静默丢弃。
- **护栏**：传了完整天猫导出时，对发货集合反查其真实状态——若是 `已发货/履约完成(已收货)/发货后取消`（不该出现在今日名单），报警提示名单可能过期、当心重复发货。

## 范围外
- 步骤1(VO API)、2(Odoo导出/确认)、5(回传ERP取消/开账单写操作)、3/4.4/4.8(天猫后台)、10(运单下载)：保持人工。
