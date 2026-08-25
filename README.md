# erp_excel_convert

围绕同一套 ERP 导出的本地 Excel 自动化脚本。**多条互不相干的流水线共处一个仓库**——
业务上各管一摊，技术上共用 `common/`（Excel 排版）、双击运行分发、多机同步与 journal/ISSUES 工作流。

## 仓库结构

| 目录 | 流水线 | 入口 | 详细文档 | 状态 |
|---|---|---|---|---|
| `vo_orders/` | **VO 拉单**（天猫 B2C 履约，本文档主体） | `Mac双击运行.command` / `Windows双击运行.bat` → `vo_orders/gui.py` | 本文档 + [SPEC.md](SPEC.md) | 在用 |
| `reorder/` | **订货辅助**（需求清单 × purchase order → 订货决策表） | `Reorder-Mac.command` / `Reorder-Windows.bat` → `reorder/reorder_gui.py` | [订货辅助输入说明](启动说明/订货辅助输入说明.md) | 在用 |
| `packing_list/` | **出口箱单**（B2B，SO 导出 → Packing List 半成品） | VO 拉单 GUI 的「箱单」标签页；或 `python3 packing_list/packing_list.py <sale.order.xlsx>` | [下方章节](#出口箱单b2b) | 在用 |
| `扫码/` | **运单扫码回流**——**非独立流水线**：VO 拉单阶段一↔阶段二之间的可选替换段（单文件 HTML，零依赖，跑在仓库机） | 浏览器打开 `扫码/扫码回流.html` | [扫码/README.md](扫码/README.md) + [下方章节](#运单扫码回流替代纸质勾选--手工转录) | 在用 |
| `sales_insight/` | **销售分析 + 安全库存提醒 + Safety Stock 回写 ERP** | `ERP回写-Mac.command` → `erp_writeback_gui.py`「销售分析」页；或 `python3 sales_insight/sales_insight.py <销售数据.xlsx> --products <product.product.xlsx>` | [sales_insight/README.md](sales_insight/README.md) | 在用 |
| `vo_orders/fs_writeback.py` | **FS 回写**（采购单 → 供应商代号写回产品主数据 `FS`） | `ERP回写-Mac.command` →「FS 回写」页；或 `python3 vo_orders/fs_writeback.py <purchase.order.xlsx> <product.product.xlsx>` | [下方章节](#erp-回写两条销售分析--fs-回写) + 脚本 docstring | 在用 |
| `po_reconcile/` | **采购对账**（采购 PO ↔ 财务 PO，算未到货量） | `python3 po_reconcile/po_reconcile.py <purchase.order.xlsx> --buyer P… --finance P…` | [po_reconcile/README.md](po_reconcile/README.md) | 算法就绪，待真实干净数据验证 |
| `po_frequency/` | **采购数量与频次**（指定供应商采购导出 → 每产品次数/数量 + 逐笔明细，纯整理不下结论） | `python3 po_frequency/po_frequency.py <purchase.order.xlsx> [--vendor …]` | [po_frequency/README.md](po_frequency/README.md) | 在用 |
| `odoo_api/` | **库存周报**（Odoo XML-RPC **只读**拉数：销量 × 在手库存 × 安全库存 三合一，周频 launchd 自动跑） | `python3 odoo_api/discover.py`（先跑，环境探针）→ `python3 odoo_api/stock_report.py` | [odoo_api/README.md](odoo_api/README.md) | 代码就绪，**待真实 ERP 连通验证** |
| `make_labels.py` | **储位标签生成**（储位编码 → 每码一页 QR + 人眼可读文字标签 PDF；尺寸驱动、热敏点阵对齐） | `python3 make_labels.py <储位.xlsx / codes.txt>` | 脚本 docstring | 在用 |
| `vo_orders/jd_export.py` | **京东选列导出**（京东后台导出 → 按预设选列） | 无（原 GUI 标签页已移除） | 脚本 docstring | **已下架，代码保留** |
| `dashboard/` | **询价直通看板**（销售 ↔ 采购，询价阶段不进 ERP）：看板页面源码 `board.html` + 从 Odoo **只读**离线导出的产品查找表 | 看板跑在 claude.ai artifact；字典 `python3 dashboard/export_product_dict.py` | [下方章节](#询价直通看板销售--采购) + 脚本 docstring | 建设中 |
| `common/` | 跨流水线共享层（Excel 排版 / 供应商简称 / 采购画像 / Supply Remark 分段） | 不单独运行 | — | — |

双击运行脚本一律留在**仓库根目录**（同事的使用习惯），内部指向各流水线入口。

**依赖按受众分三份**（2026-08-15）：`requirements.txt` = `pandas`+`openpyxl`，**启动器给同事装的最小集**（上表各条流水线的三方依赖完全相同，仅此两项）；`requirements-dev.txt` = 最小集 + `pyinstaller`（打 exe，`build_exe.bat`/`build_reorder_exe.bat` 用）；`requirements-labels.txt` = `segno`/`reportlab`/`Pillow`（储位标签专用，按需单独装）。**别往 `requirements.txt` 加只有开发/个别工具才用的包**——启动器会替同事装，装不上就直接挡住他跑拉单。启动器另有两条兜底：`.venv` 每次**探活**（目录在但已坏则自动重建）、装依赖失败**只警告不阻断**（真缺包时 Python 报清楚的 ImportError，好过打不开）。
`odoo_api/` **渠道口径必须显式给，这是最容易错且错了不报错的地方**。销量默认是全渠道；2026-08-23 实测
B 端（`S0` 开头的订单）只占 378 行订单却占 **90% 的件数**（整箱走货），而安全库存是给天猫
C 端维护的。同一窗口下「销量前 50 名里配了安全库存的」，全渠道口径算出来 3 个、
C 端口径（`--ref-contains VO,GW`）算出来 **45 个**——口径混用会把人引向完全相反的结论。
筛选时并列多出 `销量_其它渠道` / `销量_全渠道` 两列；`可撑周数` 一律按全渠道算
（库存被所有渠道一起消耗）。launchd plist 默认带 `--ref-contains VO,GW`。

**与手工导出对拍已通过**（2026-08-23，同窗口同渠道 vs `sales_insight` 2026-08-01 产出）：
安全库存 62 个运营人工值 **59 个与 ERP 现值完全相等**（另 3 个手工为 0、ERP 为空，等价）；
销量 567 个共有 SKU 中 324 个完全一致、差异中位数 0；排名 Spearman **0.913**，
前 20 名重叠 19/20。**注意 `sale.report` 是实时视图不是快照**——同一历史窗口过些天再拉数字会变
（窗口内 4.9% 的已确认单在导出后被改动过），**周报不可事后复现，要留证据就留产出文件**。

**在手库存用 `qty_available` 不用 `stock.quant`**：组合装（`x2`/`_GW`）是 phantom BoM 套件，
没有自己的实物库存、quant 上恒为 0，只有 `qty_available` 会从组件推算。实测 560 个套件里
557 个 quant 为 0，改口径后「低于安全库存」从 97 个降到 **59 个——38 个是套件造成的假警报**，
且假警报恰好落在运营配了安全库存的重点 SKU 上。非套件两种口径完全等价。
报表另出 `实物库存`（quant 原值）与 `套件` 两列；**在手列不可跨行求和**，
套件与其基础款是同一批实物。

**在途/预测必须先剔掉不会真收货的技术性采购单**：并非所有确认状态的采购单都会真的收货——某些系统集成会用 PO 建技术性映射，这类单永远不到货，却在 Odoo 里是确认状态、会生成未完成入库 move。2026-08-23 实测
**全局 7460 条未完成入库里 7350 条（98.5%）来自单一供应商的这类单**，真实在途只有 105 条 / 88 个产品。
哪些供应商属于此类是**业务知识**，配在 `config.py` 的 `ODOO_EXCLUDE_VENDORS`（不进公开库），
也可 `--exclude-vendor` 覆盖。识别要两条路都堵——链得到 PO 行的看
`purchase_line_id.order_id.partner_id`，链不到的拿 `origin` 对测试 PO 单号表（实测有 5 条只剩 origin）。
排除后 `在途入库` / `待出库` / `预测库存` / `预测缺口` 四列才有意义：
**按在手算 59 个 SKU 低于安全库存，按预测算是 75 个**——两个口径都保留，
「现在缺不缺」和「按当前进出算下来会不会缺」是两个问题。

多产一张 `安全库存待配清单.csv`（有销量但两路都没配安全库存，按销量降序）：
单看销量排名不知道哪些没设防线，单看安全库存清单不知道漏掉的是不是要紧货，
这是三路数据合并后才看得见的东西。

**不引入新依赖**：`xmlrpc.client` 是标准库，pandas/openpyxl 已在最小集里。

**所有入口与 CLI 参数的一站式清单**：[启动说明/运行指令.md](启动说明/运行指令.md)。

**两个 GUI 是刻意分开的，不是忘了合并**：`vo_orders/gui.py`（VOTool）给办公室同事用，
里面**没有任何 ERP 回写入口**；`erp_writeback_gui.py` 是个人月频维护工具，产 ERP 导入文件。
同事没有入口就不会误触回写（2026-07-08 决定，2026-08-01 加 GUI 时复核维持）。
后者只在 Mac 上跑源码，不打包 exe。

## 目标（VO 拉单）

把 [[VO拉单流程逻辑梳理]] 里**第一档**（本地确定性数据处理）的人工 Excel 操作替换为脚本：

- 步骤4：提取15位履约单号、VLOOKUP 合并 ERP 导出 + 天猫导出、筛选「履约取消/平台申请取消」、状态字段改写、无运单处理、已补运单回填
- ~~步骤6：与昨天发货表 VO Tracking No 去重比对~~ → **方案已否决**，改用「发货集合反查完整天猫真实状态」的护栏（覆盖面更大，见进度节）
- 步骤7：捡货单数据透视 + 格式美化
- 步骤8：面单筛选高亮（x2 两件装 / 剔除单件 / VO Delivery=CC）
- 步骤9：命名 + 打印格式

第二档（VO API / Odoo / 回传ERP取消）、第三档（天猫后台、运单下载）暂不自动化。

## 数据与合规

- `raw_data/` 存真实订单数据，**含个人信息（收件人姓名/电话/地址），已在 .gitignore 中排除，绝不提交、绝不放进 Obsidian vault（iCloud 同步）**。
- 测试/参考数据放 `test-data/`、产出放 `output/`，均已 gitignore，两台 Mac 间走 Syncthing 同步，不走 git。
- **任何位置的表格文件（`*.xlsx` / `*.ods` / `*.csv`）一律不提交**，这是 .gitignore 里的 PII 兜底规则，没有例外目录。

## 输入文件类型（来自 raw_data）

| 来源 | 文件举例 | 对应流程 |
|---|---|---|
| ERP 导出 | `测试0611erp导出.xlsx` / `sale.order.csv` | 步骤2/4 主输入 |
| 天猫后台导出 | `0609天猫履约单48单.xlsx` / `06102026天猫系统履约单号.xlsx` | 步骤3，做 VLOOKUP 比对 |
| VO 账单/出库 | `0610VO开发票.xlsx` / `VO出库198单.xlsx` | 步骤2 导出 |
| 昨日发货表 | `1006发货表.xlsx` | 步骤6 去重 |
| 成品参考 | `…面单+拣货单.xlsx` | 步骤7/8 目标格式 |

## 询价直通看板（销售 ↔ 采购）

询价阶段的东西不构成真实订单、不进 ERP，单独用这个看板承载。它**跑在 claude.ai artifact 上**，
本仓放两样东西：页面源码 `dashboard/board.html`，和从 Odoo 离线导出的**产品查找表**。

### ⚠ 两个「记录」分开，改动前必读

- `dashboard/board.html` 是**代码之记录**。
- 线上 artifact 是**数据之记录**。

看板会**自我发布**——有人在页面里改一条分配，它就把自己整份重新发布一遍，state 内嵌在
发布出去的 HTML 里。所以两者必然会漂。**改代码前必须先把线上版本拉下来、把 state 抠出来、
注入新代码再发布**，不能拿本地文件直接覆盖，否则抹掉的是真实业务数据。

### 数据模型（schema 2，两层）

```
采购任务 task   一个产品一轮采购（主键 = PZN + 这一轮）
  └ 分配条目 allocation   同一产品可拆给多家，同一家也可以给多批
产品字典 products         旁挂引用：pzn / nameDe / nameZh / aliases[]
```

**只做药房产品，且只做单件**——多品装与带店铺后缀的产品只存在于天猫 C 端。

### 产品字典导出

**与「页面接入 ERP API」无关**：artifact 的 CSP 封死一切外部主机，`network` 能力也不可用，
页面永远不能自己调 ERP。本脚本跑在本机、产出一个 JSON，供粘贴导入时在本地查名。
（因此查找表**不塞进看板**——1.5MB，而看板每次保存都要把自己整份重发一遍。）

```bash
python3 dashboard/export_product_dict.py                    # → dashboard/data/product_dict.json
python3 dashboard/export_product_dict.py --include-unsellable
```

- **范围只做药房产品**（`pzn` 非空）**且只做单件**。日化品没有 PZN，看板不覆盖；
  多品装与带店铺后缀的产品（`_VO`=VoyageOne 天猫 / `_GW` / `x2` / `[N Packs]`）
  **只存在于天猫 C 端，绝不进询价看板**，故默认排除（`--keep-variants` 可保留）。
  实测可售药房品 6163 → 排除变体后 **5692** 个。
- **中文名不在 `product.product` 上**——实测 10298 个可售产品里 `zh_CN` 的 name 只有 6 条真含中文。
  真正来源是 voyageone 模块的 **`product.voyageone`**（7484 条，`product_id → product.product`
  关联率 100%）。导出后**有中文名 2915 个（51%）**。
- **中文名要剥状态前缀**：`不可售` / `不采货` / `(Delisted)` / `(Außer Handel)`。
  不剥的话 4799 条「含中文」里有 **535 条其实只是被前缀污染的德语名**。
- **别名来自 `sale.order.line.vo.title`**（销售写在订单上的实际叫法，91% 含中文，
  如 `双心儿童鱼油60粒` / `Remifemin 莉芙敏片 经典版 60粒`），按 `sku ↔ vo.code` 关联，
  命中药房产品 59 个。**`sku` 里实测有零宽空格**，必须先归一再匹配。
- **⚠ 排除变体是 PZN 可用的前提**：不排除的话实测 **352 个 PZN 撞号**——
  `[2 Packs]ANTIHYDRAL Salbe 70g` 与 `ANTIHYDRAL Salbe 70g` 同为 PZN `00052729`，
  量价不同，混进来会把采购量算错。**排除后 352 → 11**。
  剩余 11 个是**同一单件产品的重复建档**（一条旧内部码、一条 PZN 码，如
  `Betaisodona_00721478` vs `Betaisodona_01931491`），挑哪条都是同一件实物，
  风险只是挑到已停用的那条 → 写进产出 `pznCollisions`，导入时告警交人选，**不静默取一条**。
  PZN 归一直接复用 `reorder/reorder_helper.norm_pzn`（那边的「ID 优先，PZN 回退」
  是另一场景的结论，本表因已排除变体而可用 PZN），不另写一套。
- **产出 `dashboard/data/product_dict.json` 已 gitignore**（1.5MB，可随时重跑）。
  每条：`pzn` / `nameDe` / `nameZh` / `aliases[]` / `odooId` / `defaultCode` / `voCode` /
  `saleOk` / `source`；顶层另有 `stats` 与 `pznCollisions`。


## 产出文件

工具按店各跑一次（ERP 单店输入；天猫两店混合经连接键 ∩ERP 收敛），产出全部带 `VO`/`GW` 后缀。

**阶段一（分流后，给员工分头手工跟进）：**

| 产出 | 内容 | 下一步手工动作 |
|---|---|---|
| `新订单获单清单{店}.xlsx` | 系统履约单号（履约单状态=新订单 ∩ ERP）| 复制 → 天猫后台批量获单 |
| `YYYY年MM月DD日{店}{n}单 拣货表+面单.xlsx` | 发货+已补运单（含无货勾选页）| 打印交仓库 |
| `YYYY年MM月DD日{店}{n}单 扫码清单.csv` | 订单级白名单（序号 / Order Reference / VO Tracking No / 店），与拣货表+面单同源同批 | 载入扫码 HTML，走[运单扫码回流](#运单扫码回流替代纸质勾选--手工转录)（选用） |
| `回传ERP销售上传表{店}.xlsx` | 取消/无运单/已补运单三类 Terms 写回**一张** | 上传 ERP，按关键词分别 Cancel/标记/恢复 |
| `取消订单清单.xlsx` | 取消订单的 系统履约单号 + Order Reference（种子表，仅有取消单时产出）| 回传天猫后把后到的取消单手工补录 → 阶段二生成取消出库单 |
| `{店}补货预判清单.xlsx` | 今日需求 / 在手 / 缺口 + `FS` + `Safety Stock` + 采购画像（**仅 GUI 有入口**：需在阶段一选填「采购单导出」；ERP 导出还须勾上 `FS`/`Safety Stock`/`Supply Remark` 三列）| 补货决策；读的 `Safety Stock` 正是 `sales_insight` 写回 ERP 的那个字段 |

**阶段二（仓库反馈缺货后）：**

| 产出 | 内容 |
|---|---|
| `系统履约单号.xlsx` (B) | 实际发货履约单号 → 上传天猫 |
| `发货表.xlsx` (C) | Order Reference + Tracking（GW/VO 分 sheet）|
| `账单上传.xlsx` (D) | External ID + 账单标签 → ERP 开账单 |
| `出库单.xlsx` (E) | stock picking 过滤+统一发货日期，**合并一张不分店** → ERP 标记出库 |
| `取消出库单.xlsx` | stock picking 过滤取消订单，Tracking Reference 统一写 `订单取消`、不写 Carrier/ID、**合并一张不分店** → ERP 按标记筛出批量取消 |

> ⚠ **`缺货记录.xlsx` 已从阶段二移除**，现在不产出。未来单独成一个阶段（无货清单 × 库存 ERP 筛查），
> `build_shortage()` / `read_marked()` 作休眠代码保留（`stage2.py:356` 起）。

**当天收尾（跨店）：**

| 产出 | 内容 |
|---|---|
| `IHTCTGMBH+IH{YYYYMMDD}+{单数}.xlsx` | N 份发货表合并去重 → 上传货代核对（唯一跨店产出）|
| `YYYY年MM月DD日{n}单 天猫回执.xlsx` | 同一次货代合并的第二份产出：所有发货 Order Reference 后15位、各渠道合并去重 → 上传天猫做回执 |

阶段二无货入口采用**「直接取有货(0)」**：仓库返回表按 0/1 标记，多品订单全 0 才整单发货，任一无货整单不发、任一留空报警——漏返回不会默认全发。

### 取消出库单用法（清理取消订单遗留的 dangling picking）

订单在天猫取消后需人工进 ERP 取消订单，但取消**不连带取消其 picking（出库单）**，ERP 出库端因此堆积大量未处理 picking。此功能产出一张可导入 Odoo 的 picking 回写文件，把取消订单对应 picking 的 `Tracking Reference` 统一写成 `订单取消`，同事导入后在 ERP 按此标记一次性筛出、全选、批量取消。

因取消是**滚动产生**的（打包寄出后、回传天猫前买家仍可取消，回传后还会冒 1~5 单），取消集在阶段一时非最终态，故采用「播种 + 补录 + 生成」两步：

1. **阶段一** 自动产出 `取消订单清单.xlsx` 种子表（当批取消单）。
2. 回传天猫后，把后到的取消单（填**系统履约号 SCP**）手工 append 进该表。
3. **阶段二** 传「取消订单清单」+「出库原始数据」→ 生成 `取消出库单.xlsx`。

CLI：

```
python3 vo_orders/stage2.py --erp <ERP导出> --cancel-list <取消订单清单> --picking <出库原始数据>
```

- 只带 `--cancel-list` + `--picking`（不给有货/无货清单）也能单独跑出取消出库单，便于收尾时补跑；`--erp` 仍需带（作「别漏 ERP」护栏，与其余产出共用入口）。
- GUI：阶段二「取消订单清单」为选填；只填它 + 出库原始数据即进入「仅取消模式」，可不填有货/无货清单。
- 与实际发货/打包/寄出互不影响；未传取消清单则该产出跳过，其余四张不变。

## 运单扫码回流（替代纸质勾选 + 手工转录）

> 📖 **完整文档（含给仓库/新同事的使用说明）**：[扫码/README.md](扫码/README.md)。
> 那里讲清了它的层级位置（**不是独立流水线**，是 VO 拉单阶段一↔阶段二之间的可选替换段）、
> 与其它流水线的关系（无）、四态反馈、导出文件去向、存储降级与恢复三前提。本节是概览。

**选用**功能，替代仓库在打印件上「纸笔勾选完成打包的订单 → 员工手工把勾选结果录进 Excel」这两步。核心心智：**扫到 = 有货，清单内未扫到 = 无货**。扫码集合直接对齐阶段二「有货订单清单」入口，无信息量损失（SKU 级缺货仍由拣货员标在拣货单上，无货勾选本就是逐订单 0/1）。

```
阶段一 vo_orders/build_excel.py  →  …{店}{n}单 扫码清单.csv （与拣货表+面单同源同批）
                            │ 仓库机浏览器载入
扫码端 扫码/扫码回流.html  →  有货清单{店}{n}单.csv   （扫一单记一单；按店各一份）
                            │ 投阶段二「有货订单清单」入口
                            └→  未知来源运单{n}单.csv   （有名单外扫描时才产出，交人工核对）
阶段二 stage2 load_shipped_map  →  实际发货集合      （吃 .csv，见下）
```

**扫码端 `扫码/扫码回流.html`**：单个 HTML 文件、零安装、仓库机浏览器直接打开（USB / 蓝牙 HID 扫码枪对浏览器行为一致，换硬件零改动）。载入当日 `…{店}{n}单 扫码清单.csv`（文件名带当天日期）作白名单（**可多选/多次追加，两店可一起载入**，追加不清空已扫），扫面单顶部 LP 一维码（= VO Tracking No）校验四态：

| 扫码结果 | 反馈 |
|---|---|
| 清单内 + 首次 | 绿底满屏大字 + 确认音 + 记入有货 |
| 清单内 + 重复 | 红底 + 报警音 + 拒绝（勿重复装箱）|
| LP 结构但不在任何已载入清单 | **黄底 + 上行双音 + 记入「未知来源」**（隔日残留/串批本就会一起寄走，硬拒与流程冲突）；单独导出交人工 |
| 非 LP 格式（SF 号/订单号/二维码）| 红底 + 报警音 + 拒绝 |

扫码本身无需点击，**声音为主异常通道**（400 单节奏下操作员手在贴单放箱、不盯屏）；红色状态挂到下一次成功扫描。**唯一的例外是分批清点**：每扫满 N 件（默认 10，底栏可改可关）弹全屏浮层，要操作员当场数一遍刚才那一摞、点「对上了/对不上」再继续——把「收工清点」拆进扫描过程，将误差定位范围从全天 400 件缩到 N（详见 [扫码/README.md](扫码/README.md) 第二节）。界面常驻 已扫/总 + 未知来源计数 + **本批 x/N** + 时钟 + 最近记录。持久化按日期存档、误刷不丢，写失败自动降级 IndexedDB（顶栏常驻黄「⚠备用存储」）、两者皆挂则仅内存（红「⚠仅内存·勿刷新」）——**任何降级都不阻断扫码与导出**（数据源是内存）；白名单也随存档持久化，**误刷新自动恢复**（"已恢复今日"，无需重载清单）；启动自动清理 30 天前旧存档。收工点「导出有货清单」得**单个 zip**（内含各店有货清单 + 未知来源；一次下载规避 Chrome 多文件拦截，解压后投阶段二）。

**阶段二入口**：`stage2.load_shipped_map` 已兼容 `.csv`（`_read_tables`：`.csv` 走 `pd.read_csv`，否则 `pd.read_excel`），把导出文件填进 GUI 阶段二「有货订单清单」即可（连接键由每行 `SCP\d+` 提取，不认列名/sheet）。走**有货清单（单号集合）**而非无货勾选 0/1 网格，天然绕开多品订单多行触发留空报警的隐患。

> ⚠ **护栏方向反转（上线必读）**：现行阶段二「任一留空报警」防的是漏返回导致误发；扫码后未扫即默认无货、**不会留空、该护栏自然失效**。已寄出但漏扫的包裹会被判缺货（天猫不回传、客户查不到物流、向缺货记录注入假数据）。**必须以收工计数对账替代**：扫码计数 vs 货代实际箱数 / 取件回单件数，对不上当场查。上线前几日与纸质勾选**并行**双向对账，数据稳定后再撤纸质。**分批清点**（每 N 件停一次核对）是把收工对账拆进过程、当场缩小误差范围的执行机制，与收工计数对账**两者都要、互不替代**。
>
> 硬件：扫码枪须配「回车(Enter)后缀」+ **实时传输**模式（禁批处理/存储转发，否则错误要等扫完回底座才暴露）；面单有多个条码（LP / SF / 二维码），建议硬纸板开窗物理只露顶部 LP 条码，从源头消灭误扫；操作员离屏较远建议外接音箱。

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
- GUI（推荐，全英文）：双击 `Reorder-Windows.bat` / `Reorder-Mac.command`（首次自动建环境），或 `python3 reorder/reorder_gui.py`。打包成 Windows exe：`build_reorder_exe.bat`（产出 `dist/ReorderHelper.exe`）。
- CLI：
  ```
  python3 reorder/reorder_helper.py <需求清单.xlsx> <purchase order.xlsx> [out.xlsx] [--master product.product.xlsx]
  ```

## 出口箱单（B2B）

与拉单主流程无关的独立流水线：ERP 导出的 `sale.order` **行明细** → 可直接交仓库填的 Packing List 半成品。
入口在 VO 拉单 GUI 的「箱单」标签页（同事日常用），CLI 见 [运行指令](启动说明/运行指令.md#二箱单b2b-出口)。

**分工原则：机器知道的填好，现场才知道的留空。** 品名 / SKU / 条码 / HS 编码 / 原产国由脚本填；
托盘号 / 批次号 / 箱号 / 箱规 / 箱数 / 保质期 / 毛重 / 尺寸 / 体积重留空给仓库手填；
`Quantity total` 写成公式 `=箱规×箱数`，仓库填完自动出数。每个 SKU 底下预留空行（默认 2）供拆批次。
要几张 SO 合并成一张箱单，就在 ERP 里一次性导到同一份文件里，脚本按 `Order Reference` 自动分单。
输出同名自动加序号，不覆盖。

## ERP 回写两条（销售分析 / FS 回写）

这两条共用入口 `ERP回写-Mac.command` → `erp_writeback_gui.py`，是**个人月频维护工具**，
产出的是给 Odoo 导入界面用的文件——**上传一律保持人工**，脚本不碰 ERP。

| | 销售分析 · 安全库存 | FS 回写 |
|---|---|---|
| 写回哪个字段 | `Safety Stock`（+ `Supply Remark` 的自己那一段） | `FS`（**不碰 `Supply Remark`**） |
| 输入 | 销售数据（按周分组导出）+ `product.product` + 运营的安全库存表（选填但优先） | `purchase.order` + `product.product` |
| 首次导入试水 | `--test-sku <SKU>`：出单条表核对，**同时照常出全量**，验完直接导全量 | `--sample N`：按覆盖面挑 N 行（每行 FS 值互不相同） |
| 详细文档 | [sales_insight/README.md](sales_insight/README.md) | `vo_orders/fs_writeback.py` 顶部 docstring |

**两条吃同一份产品主数据导出，筛选条件写死：只勾 `can be sold`。**
加别的条件会实打实漏货——2026-08-02 实测 `VO active=true` 只有 4575 行（无筛选 10331 行），
漏掉 9 个运营在管的 SKU（多为 `x2`/`x3` 组合装与渠道变体，但也有普通 SKU）。

两道保护（2026-08-02 首次真正导入 ERP 后已在真实数据上验证）：
- **FS 现值看着像人写的采购判断时整行跳过**（如「首选AEP 不在Phoenix订」「MHD原因暂时停止订货」）——
  那是画像给不了的业务经验，机器不该拿聚合结果盖掉。
- **`Supply Remark` 重跑不堆叠**：按签名认出自己写的旧段**替换**，别人的段与人工原文原样保留。

供应商代号来自 `config.py` 的 `VENDOR_ALIAS`（不进公开库）：写进 ERP 的是代号不是供应商真名。

## 采购对账（算法就绪，待验证）

采购在一张 PO 里记订购需求，财务按每批实收另建 PO，于是采购单永远显示原始订购量、看不出还缺什么。
`po_reconcile` 比对两边算出未到货量，并可回写采购 PO。详见 [po_reconcile/README.md](po_reconcile/README.md)。

它依赖两条前提（财务单每件都是对这张采购单的交付；采购单是订购的完整记录）。
一旦出现「财务单已收 > 订购量」说明前提被破，脚本**拒绝生成回写导入表并以退出码 2 中止**——
负的未到量在满足前提的数据上不可能发生，此时回写只会把错误写进 ERP。

## 库存周报（`odoo_api/`，Odoo 只读拉数 + 天猫处罚数据驱动优先级）

**这条流水线的目标（2026-08-23 明确）：让库存合理化以降低天猫端违规订单，
至少不能因为库存问题挨罚。** 此前报表按销量排名、按安全库存算缺口，从未与实际罚款挂钩
——等于在猜哪里要紧。接入天猫「供应商违规数据」导出（`--violations`）后，优先级由真实罚款决定。

**需求是尖峰型的，均值口径正好在最要紧的时刻失真**（2026-08-23 实测，6 月至今 12 周，
样本 = 有销周数 ≥3 且 ≥20 件的 170 个 SKU）：峰值周/周均**中位数 2.38×**，76% 的 SKU ≥2×。
更要命的是——**86 个 SKU 按周均看「能撑 ≥2 周」，其中 30 个按峰值看连一个高峰周都接不住**，
而罚款恰恰发生在峰值那一周。故报表除周均外还算 `峰值周销量`/`峰值倍数`/`可撑峰值周数`/`峰值缺口`。
窗口默认 **12 周**：4 周是 2026-08-01 为「均值推算安全库存」定的口径（仍成立），
但峰值识别需要更长窗口，4 个数据点看不出尖峰。

处罚之后销量会塌：6 月以来被罚 SKU 罚后 4 周均销量**中位 −56.9%**，对照组 −7.5%。
**⚠ 相关非因果**——缺货本身就压销量，分不清是流量惩罚还是货没补上，只用于说明代价量级。

**归因口径最终定为 `narrow`**：「晚于应发货时间」**不算**缺货类。曾试过放宽（实际操作中它大部分
确由缺货导致），实测后否决——归因面会从 15% 扩到 **95.5%** 的单数，「是不是库存问题」这个筛子
几乎不再筛任何东西，分类退化成同义反复。**分类的价值在于能排除。** `--stock-violation-mode wide` 可对比。

**派生指标一律以在手库存为分子**（`预测库存`/`预测缺口` 已移除）：预测会掺进虚拟库存，
且天猫的处罚触发于"此刻能不能发货"。但这是双向取舍——预测 = 在手 + 在途 − 待出，
在途虚高、待出压低，实测 170 个样本里 **94 个在手 > 预测**，换口径后接不住峰值的 SKU
从 111 减到 **86**。被放过的 25 个里有真风险（`Femibion_15199958` 在手 128 但已预留 108，
实际可接新单仅 20，峰值周要 124）。故给了 `--stock-basis available`（在手−已预留）可切换。

**补货优先级看近期是否还在卖，但有个会反噬的陷阱**：近两周零销的 511 个 SKU 里 **201 个在手为 0**
——零销可能正是缺货造成的。天真地「零销即降级」等于让缺货自我强化（卖光→没销量→判定不重要
→不补货→继续缺货）。故分五档，`3-待观察`（零销且无货）**不降级**。新品用 `sale.report` 的
`date:min` 识别，**不能用 `product.product.create_date`**（实测有销量的 969 个 SKU 全部建档 >90 天）。

首次实测（8647 条处罚，7557 条能归到 SKU，2026-01 ~ 08）：**缺货类违规占 14.3% 的单数
但占 54.7% 的罚款**，单均是其它类型的 5 倍；单均最贵的两个类型都是缺货类。
**缺货类涉及 529 个 SKU，其中 449 个（86%）至今没配安全库存**，合计 2,294 EUR / 754 单
——安全库存配在哪，跟罚款实际发生在哪，是脱节的。新产出 `缺货违规风险清单.csv` 就是接上它。

上面两条回写吃的是**手工导出**的三份 Excel。周频跑一次，每次都要人进 ERP 点导出、
还得记住筛选条件。`odoo_api/` 把这一步换成 XML-RPC 拉数：销量（`sale.report`，
默认近 4 周）× 在手库存（`stock.quant`，内部库位汇总）× 安全库存（**两路并列**）
合成一张周报，`launchd` 每周一自动跑。详见 [odoo_api/README.md](odoo_api/README.md)。

**只读。** 本层不写 ERP——`Safety Stock` 回写继续走人工导入表，保留人工过目、出错能中断
（沿用 SPEC「第二档 Odoo API 暂不自动化」的判断，本次只放开读）。写方法在客户端就被拦下，
不指望服务端权限兜底。

**安全库存两路并列是刻意的**：A = 产品主数据上的自定义 `Safety Stock` 字段
（`sales_insight` 回写、补货预判清单读的就是它），B = `stock.warehouse.orderpoint.product_min_qty`
（Odoo 标准补货规则，须滤掉 `trigger=manual` 的临时建议——该字段 Odoo 14 就有，2026-08-23 实测 5 条里 4 条正是 manual）。**2026-08-23 首测已给出答案：A 侧 152 个、B 侧 1 个，补货规则基本没在用**，所以「不一致」只报真冲突（都有但不等 / 只有 B），「只有 A」是常态、只报计数——所以并列成两列 + 单出一张
`安全库存不一致.csv`，先看清楚再谈以谁为准。缺口暂按 A 算，A 空退回 B。

**不引入新依赖**：`xmlrpc.client` 是标准库，pandas/openpyxl 已在最小集里。
凭据走 `config.py` 的 `ODOO = {...}` 或环境变量，用 **API Key** 不用登录密码。
脚本账号只需 Sales 与 Inventory 的读权限，**不需要 Administration/Settings，不需要 Studio**。

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
- [x] ~~缺货记录（明细按SKU合并 + SKU汇总，回连ERP增强库存/条码/货位）~~ → **已从阶段二移除**，未来单独成一个阶段（无货清单 × 库存 ERP 筛查），代码休眠保留
- [x] 步骤9 文件命名 + 打印格式
- [x] E 出库单（`stage2.build_E`）：stock picking 过滤+统一发货日期，**合并一张不分店**（2026-08-15 由拆 VO/GW 改），回传 Odoo 标记出库。
- [x] **取消出库单**（`stage2.build_cancel`，与 build_E 共享 `build_picking_writeback` 原语）：过滤取消订单 picking，Tracking Reference 写 `订单取消`、不写 Carrier/ID、合并一张 → ERP 批量取消。阶段一播种取消清单 + 人工补后到的 + 阶段二生成；可仅取消模式单独补跑。
- [x] **货代合并发货表**（`stage2.build_forwarder`）：N 份发货表去重 → `IHTCTGMBH+IH{日期}+{单数}.xlsx`，唯一跨店产出；**同时出第二份「天猫回执」**（发货单号后15位合并去重，上传天猫）。
- [x] GUI(`gui.py`) + Windows exe 打包：办公室员工双击使用；含「④ 货代合并」入口。
- [x] **先核对再发货**：采用护栏（发货集合反查完整天猫真实状态报警），替代原「昨日发货 VO Tracking 去重」方案——覆盖面更大。
- [x] **运单扫码回流**（选用，替代纸质勾选+手工转录）：阶段一 `build_excel.py` 产 `YYYY年MM月DD日{店}{n}单 扫码清单.csv`（订单级白名单，与拣货表+面单同源）；单文件 `扫码/扫码回流.html` 零安装扫 LP 校验四态（首次绿+确认音 / 重复红拒 / 名单外黄屏记录 / 非 LP 红拒），多店清单可一起载入，声音为主、红态挂到下次成功、持久化三层降级（localStorage→IndexedDB→仅内存，顶栏常驻标识）+ 白名单入库**刷新自动恢复**（红色「清空」键彻底清空白名单+存档，作载错清单的换清单逃生口）+ 30 天自动清理，导出**单个 zip**（内含 `有货清单{店}{n}单.csv` + `未知来源运单{n}单.csv`，规避 Chrome 多文件拦截，解压后投）；`stage2.load_shipped_map` 加 `_read_tables` 兼容 `.csv` 投「有货订单清单」入口。走单号集合绕开留空报警；上线须以收工计数对账替代失效护栏。
- [x] **订货辅助工具**（`reorder_helper.py` + 全英文 `reorder_gui.py`）：需求清单 × purchase order → 一行一品订货决策表；PZN 按模式抽取（支持销售分析 `[前缀_PZN]` 嵌入 + 金额列不误判 + 无 PZN 报错护栏）；选填 product.product 主数据富化干净身份字段（PZN/Name/Barcode/Internal Reference/库存），双键索引桥接 PZN 更新错位；连接键 Product ID 优先（数字/External ID 归一互通）+ 逐行回退 PZN，绕开官方 PZN 空白/脏值。启动器 `Reorder-Windows.bat`/`Reorder-Mac.command` + 打包 `build_reorder_exe.bat`。
- [x] **重构：按流水线拆目录**（`vo_orders/` / `reorder/` / `packing_list/`）+ 抽最小 `common/`（`xlsx` 排版 / `vendor` 供应商简称 / `po` 采购画像 / `remark` Supply Remark 分段）。
- [x] **出口箱单**（`packing_list/packing_list.py`）：SO 行明细 → Packing List 半成品，机器可知的列填好、现场才知道的留空，`Quantity total` 用公式；接进 VO 拉单 GUI「箱单」标签页（同批下架了京东标签页，`jd_export.py` 代码保留）。
- [x] **销售分析 + 安全库存**（`sales_insight/`）：销量排名 + 安全库存提醒 + `Safety Stock` 回写 ERP，吃按周分组导出（周数自动）；候选值表可直接导入；`--test-sku` 试水时**同时出全量**；试水报错逐级判定说清缺哪一环。回写的 `Safety Stock` 被阶段一「补货预判清单」读走——两条流水线在此接上。
- [x] **FS 回写**（`vo_orders/fs_writeback.py`）：采购画像 → 供应商**代号**写回产品主数据 `FS`；`--sample N` 按覆盖面挑试水样本；人写的采购判断整行跳过；滤掉测试单与费用类 SKU；**只写 FS，不碰 `Supply Remark`**（那字段属于运营）。
- [x] **ERP 回写 GUI**（`erp_writeback_gui.py`）：销售分析 / FS 回写两页，个人月频维护用。**与 VOTool 刻意分开**——同事界面里没有任何 ERP 回写入口就不会误触（2026-07-08 决定，2026-08-01 复核维持）；只在 Mac 跑源码，不打包 exe。
- [x] **两条回写首次真正导入 ERP 并反向复核**（2026-08-02）：FS 1569/1569 一致；人写值与费用类 SKU 两道保护实证生效；`Supply Remark` 重跑替换而非堆叠、人写原文保住。产品主数据导出条件由此写死为**只勾 `can be sold`**。
- [x] **采购数量+频次**（`po_frequency/po_frequency.py`）：单一供应商 purchase.order 行式导出 → `Summary`（每产品：采购频次/总量/均·最·大 per purchase/首末采购/跨度/平均间隔）+ `Details` 两 sheet 英文表头；`--vendor` 子串过滤、`--out` 可覆盖默认落位。复用 `common/po` 归一 + `common/xlsx` 排版（Details 上万行走轻量表头样式，27s→3.3s）。纯数量+频次、不掺结论/分析列。
- [x] **SKU 归一统一**（`common/po._po_base_sku`）：口径统一为 `([xX]\d+|\*\d+|_VO|_GW)+$`，全仓一套；修掉 `x2_GW` 组合后缀旧规则脱不掉的漏匹配；三调用方（po_frequency / 补货预判 / FS 回写）对称受益。
- [x] **储位标签生成**（`make_labels.py`）：储位编码 → 每码一页「QR + 人眼可读文字」标签 PDF，尺寸驱动、几何全部吸附打印头点阵（得力 DL720C 40×20mm@203dpi），QR 版本锁定 + 模块边长硬下限校验（低于扫描枪规格直接拒绝），可选退化回测 `--verify` 与 1-bit PNG `--png`。输入吃 Excel（默认 `储位编码` 列，`-c` 可改）或 txt（每行一个），去重保序；产出落 `output/labels/`（`储位标签_QR.pdf` + `储位编码.csv`）。依赖不在主 requirements 里，按需 `pip install -r requirements-labels.txt`。UTF-8 控制台兜底免去 Windows cp1252 崩溃。打印须选「实际大小 / 100%」，否则模块宽被缩放。早期 Code128 三变体版 `make_bin_labels.py` **已删除**（2026-08-15）：其 QR 变体是本脚本的劣化版（模块 0.375mm、位置不吸附点阵、无版本锁定/下限校验、文字按 7 位硬切），独有的 Code128 与 20×40 竖版现场未采用；将来若要，做成本脚本的开关复用同一套点阵吸附与校验，不再另起脚本。
- [ ] **采购对账**（`po_reconcile/`）：算法与前提校验已实现、构造数据测试通过；**等真实干净数据验证后再上线**。
- [ ] **库存周报**（`odoo_api/`）：XML-RPC **只读**拉数替掉三份手工导出——`sale.report` 自定义窗口销量（默认近 4 周，不用周期写死 12 个月的 `sales_count`）× `stock.quant` 内部库位在手 × 安全库存两路并列（产品自定义字段 vs `orderpoint.product_min_qty`，后者已滤 `trigger=manual` 临时建议）→ `库存周报.xlsx/.csv` + `安全库存不一致.csv`；产品范围取并集（`sale_ok` ∪ 有销量 ∪ 有库存 ∪ 有规则），避免漏掉已下架但仍在卖的货；`discover.py` 环境探针先跑、`test_stock_report.py` 构造数据 21 项断言通过；`launchd` 周一 08:30。**代码就绪，待用真实凭据跑通并与手工导出逐列对拍**。
