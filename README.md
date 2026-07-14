# erp_orders_convert

VO（VoyageOne）拉单流程中 Excel 处理环节的自动化脚本。

## 目标

把 [[VO拉单流程逻辑梳理]] 里**第一档**（本地确定性数据处理）的人工 Excel 操作替换为脚本：

- 步骤4：提取15位履约单号、VLOOKUP 合并 ERP 导出 + 天猫导出、筛选「履约取消/平台申请取消」、状态字段改写、无运单处理、已补运单回填
- 步骤6：与昨天发货表 VO Tracking No 去重比对
- 步骤7：捡货单数据透视 + 格式美化
- 步骤8：面单筛选高亮（x2 两件装 / 剔除单件 / VO Delivery=CC）
- 步骤9：命名 + 打印格式

第二档（VO API / Odoo / 回传ERP取消）、第三档（天猫后台、运单下载）暂不自动化。

## 数据与合规

- `raw_data/` 存真实订单数据，**含个人信息（收件人姓名/电话/地址），已在 .gitignore 中排除，绝不提交、绝不放进 Obsidian vault（iCloud 同步）**。
- 开发/分享用脱敏样例，放 `samples/`（结构真、内容假）。

## 输入文件类型（来自 raw_data）

| 来源 | 文件举例 | 对应流程 |
|---|---|---|
| ERP 导出 | `测试0611erp导出.xlsx` / `sale.order.csv` | 步骤2/4 主输入 |
| 天猫后台导出 | `0609天猫履约单48单.xlsx` / `06102026天猫系统履约单号.xlsx` | 步骤3，做 VLOOKUP 比对 |
| VO 账单/出库 | `0610VO开发票.xlsx` / `VO出库198单.xlsx` | 步骤2 导出 |
| 昨日发货表 | `1006发货表.xlsx` | 步骤6 去重 |
| 成品参考 | `…面单+拣货单.xlsx` | 步骤7/8 目标格式 |

## 产出文件

工具按店各跑一次（ERP 单店输入；天猫两店混合经连接键 ∩ERP 收敛），产出全部带 `VO`/`GW` 后缀。

**阶段一（分流后，给员工分头手工跟进）：**

| 产出 | 内容 | 下一步手工动作 |
|---|---|---|
| `新订单获单清单{店}.xlsx` | 系统履约单号（履约单状态=新订单 ∩ ERP）| 复制 → 天猫后台批量获单 |
| `YYYY年MM月DD日{店}{n}单 拣货表+面单.xlsx` | 发货+已补运单（含无货勾选页）| 打印交仓库 |
| `回传ERP销售上传表{店}.xlsx` | 取消/无运单/已补运单三类 Terms 写回**一张** | 上传 ERP，按关键词分别 Cancel/标记/恢复 |
| `已补运单清单{店}.xlsx` | 系统履约单号 | 天猫后台批量打面单、发群 |
| `取消订单清单.xlsx` | 取消订单的 系统履约单号 + Order Reference（种子表，仅有取消单时产出）| 回传天猫后把后到的取消单手工补录 → 阶段二生成取消出库单 |

**阶段二（仓库反馈缺货后）：**

| 产出 | 内容 |
|---|---|
| `系统履约单号.xlsx` (B) | 实际发货履约单号 → 上传天猫 |
| `发货表.xlsx` (C) | Order Reference + Tracking（GW/VO 分 sheet）|
| `账单上传.xlsx` (D) | External ID + 账单标签 → ERP 开账单 |
| `出库单{店}.xlsx` (E) | stock picking 过滤+统一发货日期 → ERP 标记出库 |
| `取消出库单.xlsx` | stock picking 过滤取消订单，Tracking Reference 统一写 `订单取消`、不写 Carrier/ID、**合并一张不分店** → ERP 按标记筛出批量取消 |
| `缺货记录.xlsx` | 明细 + SKU 汇总，回连 ERP 库存/条码/货位 |

**当天收尾（跨店）：**

| 产出 | 内容 |
|---|---|
| `IHTCTGMBH+IH{YYYYMMDD}+{单数}.xlsx` | N 份发货表合并去重 → 上传货代核对（唯一跨店产出）|

阶段二无货入口采用**「直接取有货(0)」**：仓库返回表按 0/1 标记，多品订单全 0 才整单发货，任一无货整单不发、任一留空报警——漏返回不会默认全发。

### 取消出库单用法（清理取消订单遗留的 dangling picking）

订单在天猫取消后需人工进 ERP 取消订单，但取消**不连带取消其 picking（出库单）**，ERP 出库端因此堆积大量未处理 picking。此功能产出一张可导入 Odoo 的 picking 回写文件，把取消订单对应 picking 的 `Tracking Reference` 统一写成 `订单取消`，同事导入后在 ERP 按此标记一次性筛出、全选、批量取消。

因取消是**滚动产生**的（打包寄出后、回传天猫前买家仍可取消，回传后还会冒 1~5 单），取消集在阶段一时非最终态，故采用「播种 + 补录 + 生成」两步：

1. **阶段一** 自动产出 `取消订单清单.xlsx` 种子表（当批取消单）。
2. 回传天猫后，把后到的取消单（填**系统履约号 SCP**）手工 append 进该表。
3. **阶段二** 传「取消订单清单」+「出库原始数据」→ 生成 `取消出库单.xlsx`。

CLI：

```
python3 stage2.py --erp <ERP导出> --cancel-list <取消订单清单> --picking <出库原始数据>
```

- 只带 `--cancel-list` + `--picking`（不给有货/无货清单）也能单独跑出取消出库单，便于收尾时补跑；`--erp` 仍需带（作「别漏 ERP」护栏，与其余产出共用入口）。
- GUI：阶段二「取消订单清单」为选填；只填它 + 出库原始数据即进入「仅取消模式」，可不填有货/无货清单。
- 与实际发货/打包/寄出互不影响；未传取消清单则该产出跳过，其余四张不变。

## 订货辅助工具（`reorder_helper`，独立脚本 + 全英文 GUI）

给**订货同事**（含不会中文的）用的独立工具，与上面拉单主流程互不相关：把「要订的货品」逐个去 ERP 采购记录里查最近采购价/供应商/数量/库存，产出一张**一行一品的订货决策表**。产出列名全英文。

> 📖 **单页速查**：[启动说明/订货辅助输入说明.md](启动说明/订货辅助输入说明.md)——输入顺序、每个入口必需/选填的字段名清单（含 ERP 导出勾选建议）、匹配逻辑与用法，末附英文 Quick Reference。

**输入**（前两个必填，第三个选填）：

| 输入 | 说明 |
|---|---|
| 商品/需求清单 | 待发货明细表、ERP 导出（sale.order / 销售分析，`[前缀_PZN] 名称` 嵌入式）、或纯 PZN 清单皆可。PZN 从引用/名称里按 `前缀_PZN` / `PZN-####` / 整格 7~8 位数字 抽取（金额/12位id/13位EAN 不会误判）；若带 `…Product/ID` / `…Product/External ID` 列则自动识别产品 ID |
| purchase order 导出 | Odoo `purchase.order` 行式导出，提供最近采购 vendor/价/量/日期 + 库存。带 `Order Lines/Product/ID` 列时建 ID 索引 |
| product.product 主数据（选填） | Odoo 产品主数据。传了就用**干净**的官方 PZN / Name / Barcode(EAN) / Internal Reference / 库存；不传则从名称/引用回退（身份信息较稀） |

**连接键（逐行决定）**：Product ID 优先——两侧都有产品 ID 时直连（数字 `200392` 与 External ID `__export__.product_product_200392_…` 两形态自动归一互通），绕开官方 PZN 字段空白（实测 ERP 导出约 1/4 行空白）、脏值（`17444652_`/`17173992x3`）与 IntRef 嵌旧 PZN 三坑；外部 PZN 输入配 master 时走 **PZN → master → ID → PO** 三段桥。无 ID 可用的行回退 PZN 匹配（待发货明细表 = 旧行为，零变化）。

**产出**（15 列）：`Product ID` · `PZN` · `Name` · `Barcode` · `Internal Reference` · `总需求` · `Quantity On Hand` · `Reorder Qty`(需求−库存) · `平台裸价` · `Last Unit Price` · `Price Diff`(裸价−采购价) · `Last Vendor` · `Last Qty` · `Last Order Date` · `Recent Purchases`（最近 5 笔）。整批无值的列自动省略（如仅 PZN 清单没 Product ID/总需求/裸价）。

> 主数据里 PZN 会更新、但名称/Internal Reference 仍嵌旧 PZN——`load_master` 按「IntRef 嵌入 PZN」和「官方 PZN 字段」双键索引桥接该错位；带 ID 的输入则直接绕开此坑。

**运行**：
- GUI（推荐，全英文）：双击 `Reorder-Windows.bat` / `Reorder-Mac.command`（首次自动建环境），或 `python3 reorder_gui.py`。打包成 Windows exe：`build_reorder_exe.bat`（产出 `dist/ReorderHelper.exe`）。
- CLI：
  ```
  python3 reorder_helper.py <需求清单.xlsx> <purchase order.xlsx> [out.xlsx] [--master product.product.xlsx]
  ```

## 环境

Python + pandas + openpyxl。

## 进度

- [x] git init + .gitignore（保护真实数据）
- [x] 确认各输入文件字段结构
- [x] 步骤4 合并/筛选（`build_excel.py`：拣货表 + 面单 + 无货勾选0/1 + 序号）
- [x] **步骤4 分流**：发货范围由完整天猫导出二段式(`履约∈{新订单,商家已接单} ∧ 面单=已完成`)推出，`classify4` 再分流 取消/无运单/已补运单/发货；拣货面单**剔除无运单+取消**。stage1/stage2 同源。GUI 输入 ERP(多选)+完整天猫导出(必选，唯一天猫输入)。
- [x] **新订单获单清单**（缺口补全）：`履约单状态=新订单` ∩ ERP → 系统履约单号，桥接「确认 order」与「天猫批量获单」两个手工步。
- [x] **回传ERP销售上传表**（三合一）：取消/无运单/已补运单三类 Terms 写回合并**一张**（替代原 取消单/无运单清单 两文件）。
- [x] 第二阶段 B/C/D（`stage2.py`）；无货入口改**直接取有货(0)**，多品全0才发、未确认报警，消除「漏返回默认全发」。
- [x] 缺货记录（明细按SKU合并 + SKU汇总，回连ERP增强库存/条码/货位）
- [x] 步骤9 文件命名 + 打印格式
- [x] E 出库单（`stage2.build_E`）：stock picking 过滤+统一发货日期，拆 VO/GW，回传 Odoo 标记出库。
- [x] **取消出库单**（`stage2.build_cancel`，与 build_E 共享 `build_picking_writeback` 原语）：过滤取消订单 picking，Tracking Reference 写 `订单取消`、不写 Carrier/ID、合并一张 → ERP 批量取消。阶段一播种取消清单 + 人工补后到的 + 阶段二生成；可仅取消模式单独补跑。
- [x] **货代合并发货表**（`stage2.build_forwarder`）：N 份发货表去重 → `IHTCTGMBH+IH{日期}+{单数}.xlsx`，唯一跨店产出。
- [x] GUI(`gui.py`) + Windows exe 打包：办公室员工双击使用；含「④ 货代合并」入口。
- [x] **先核对再发货**：采用护栏（发货集合反查完整天猫真实状态报警），替代原「昨日发货 VO Tracking 去重」方案——覆盖面更大。
- [x] **订货辅助工具**（`reorder_helper.py` + 全英文 `reorder_gui.py`）：需求清单 × purchase order → 一行一品订货决策表；PZN 按模式抽取（支持销售分析 `[前缀_PZN]` 嵌入 + 金额列不误判 + 无 PZN 报错护栏）；选填 product.product 主数据富化干净身份字段（PZN/Name/Barcode/Internal Reference/库存），双键索引桥接 PZN 更新错位；连接键 Product ID 优先（数字/External ID 归一互通）+ 逐行回退 PZN，绕开官方 PZN 空白/脏值。启动器 `Reorder-Windows.bat`/`Reorder-Mac.command` + 打包 `build_reorder_exe.bat`。
