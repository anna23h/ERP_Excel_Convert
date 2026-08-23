# odoo_api — Odoo XML-RPC 只读拉数层

**只读。本层不写 ERP。** Safety Stock 回写继续走 `sales_insight` 产的人工导入表，
保留人工过目、出错能中断（沿用 SPEC「第二档 Odoo API 暂不自动化」的判断，本次只放开读）。
写操作在 `odoo_client.Odoo.ALLOWED_METHODS` 里被客户端直接拦下，不指望服务端权限兜底。

## 它解决什么

`sales_insight` 的「销量排名 + 补货提醒」口径已经调好并在用，但输入是**三份手工导出**
（Sales Analysis 透视表 / `product.product` / 运营安全库存表）。周频跑一次，每次都要
人进 ERP 点导出、还得记住筛选条件。这一层把导出换成 API 拉数，周频 `launchd` 自动跑。

## 两个入口

### 1. `discover.py` —— 先跑这个

环境探针，只读、不产文件。把写报表前所有不确定项一次问清楚：服务端版本、可见公司、
仓库数、`Safety Stock` 自定义字段的技术名、补货规则真实条数、产品各筛选条件的行数、
`sale.report` 可用字段与 state 分布、两种在手库存口径的抽样对拍。

```bash
python3 odoo_api/discover.py
python3 odoo_api/discover.py --weeks 12 --company-id 1
```

| 参数 | 说明 |
|---|---|
| `--weeks N` | 销量口径探测窗口，默认 4 |
| `--company-id N` | 多公司环境锁定公司（`allowed_company_ids`） |
| `--timeout N` | RPC 超时秒数，默认 180 |

### 2. `stock_report.py` —— 正式周报

```bash
python3 odoo_api/stock_report.py --ref-contains VO,GW        # 推荐：只看天猫 C 端
python3 odoo_api/stock_report.py                            # 全渠道（含 B 端整箱单）
python3 odoo_api/stock_report.py --weeks 12 --external-id
python3 odoo_api/stock_report.py --since 2026-06-01 --until 2026-08-23
python3 odoo_api/stock_report.py --by-warehouse
python3 odoo_api/stock_report.py --csv-only                 # 定时任务用
```

| 参数 | 默认 | 说明 |
|---|---|---|
| `--weeks N` | 4 | 销量统计窗口周数 |
| `--since` / `--until` | — | 精确窗口；给了 `--since` 就覆盖 `--weeks` |
| `--states` | `sale,done` | 计入销量的订单状态，逗号分隔 |
| `--ref-contains` | 空（全渠道） | 只算 Order Reference 含这些子串的订单，逗号分隔、**OR** 语义。`VO,GW` = 只看天猫 C 端 |
| `--ref-excludes` | 空 | 排除 Order Reference 含这些子串的订单，逗号分隔、**AND** 语义 |
| `--by-warehouse` | 关 | 在手库存按仓拆成 `在手·<仓名>` 多列 |
| `--orderpoint-agg` | `sum` | 同一产品多条补货规则怎么合（`sum` / `max`） |
| `--company-id` | — | 多公司环境锁定公司 |
| `--external-id` | 关 | 附带 `External ID` 列（将来接回写的映射码） |
| `--csv-only` | 关 | 只写 CSV，不写 xlsx |
| `-o / --outdir` | `output/YYYYMMDD` | 输出目录 |
| `--timeout` | 180 | RPC 超时秒数 |

### 产出

| 文件 | 内容 |
|---|---|
| `库存周报.xlsx` / `.csv` | 全 SKU 按销量降序：排名 / 累计占比(ABC) / 销量 / 周均销量 / 下单次数 / 销售额 / 在手 / 已预留 / 可用 / **两路安全库存并列** / 取用值 / 缺口 / 可撑周数 / 在售 / 产品ID |
| `安全库存不一致.csv` | 两路**真正冲突**的行：都有值但不等 / 只有补货规则。「只有产品字段」不进这张表——B 侧没在用时那是常态，只在终端报个计数 |
| `安全库存待配清单.csv` | **有销量但两路都没配安全库存**的 SKU，按销量降序。三路数据合并后才看得见的东西 |

### 3. `test_stock_report.py` —— 不连 ERP 的构造数据测试

```bash
python3 odoo_api/test_stock_report.py
```

## 凭据

按优先级：环境变量 `ODOO_URL` / `ODOO_DB` / `ODOO_USER` / `ODOO_API_KEY`，
其次仓库根 `config.py` 里的 `ODOO = {...}`（模板见 `config.example.py`）。
`config.py` 已 gitignore。**用 API Key 不要用登录密码**——密钥能单独吊销、
不受双因素影响、泄露了不等于账号被接管。

⚠ `http://` 明文传输时 API 密钥在链路上是明文。内网/Tailscale 尚可，公网必须换 HTTPS。
连上时脚本会就此告警一次。

## 脚本账号需要的读权限

申请时按这个清单提，比笼统说「只读权限」好办：

| 模型 | 用途 | 一般随哪个组来 |
|---|---|---|
| `sale.report` | 销量、下单次数、销售额 | Sales / User: All Documents |
| `product.product`、`product.template` | SKU、品名、自定义 Safety Stock 字段 | 上面任一即可 |
| `stock.quant`、`stock.location`、`stock.warehouse` | 在手库存、仓库与库位 | Inventory / User |
| `stock.warehouse.orderpoint` | 补货规则最小值 | Inventory / User |
| `ir.model.fields`、`ir.model.data` | 反查自定义字段技术名、External ID | 登录用户默认可读 |
| `res.users`、`res.company` | 探针里报当前账号与公司 | 登录用户默认可读 |

**不需要** Administration / Settings 权限，**不需要** Studio。本层不装模块、不改视图。
标准 Odoo 的组权限是「读写一起给」的，要真正只读得再配 record rule；
本层已在客户端侧限死只读方法，够用。

## 部署（Mac + launchd）

用 `launchd` 不用 `cron`：到点时机器睡着了，launchd 会在唤醒后补跑，cron 直接跳过这一次。

```bash
sed -i '' "s#/Users/CHANGEME/projects/erp_excel_convert#$PWD#g" odoo_api/com.ihtct.stockreport.plist
cp odoo_api/com.ihtct.stockreport.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.ihtct.stockreport.plist
```

立刻试跑一次、看日志、卸载：

```bash
launchctl start com.ihtct.stockreport
tail -40 output/odoo_api.log
launchctl unload ~/Library/LaunchAgents/com.ihtct.stockreport.plist
```

`run_weekly.sh` 固定 cwd 到仓库根、优先用 `.venv/bin/python`、把每次运行追加进
`output/odoo_api.log`（`output/` 已 gitignore）。plist 里 `--weeks 4` 之后的参数原样传给脚本。

放本机 Mac 还是自托管服务器：本机够用，且产出直接落进已经在走 Syncthing 的 `output/`。
放服务器要多解决凭据分发和产出回传，本次没有收益。

## 坑（都已在代码里处理，改动时别退回去）

1. **产品筛选口径**。必须 `sale_ok = True`（导出界面的 `can be sold`，实测 10266 行）。
   `VO active = true` 只有 4575 行，会漏 9 个 SKU，其中 `Pollival_13748591` 是个普通 SKU
   ——这条件不是「更宽」，是实打实漏货。2026-08-01/02 踩过两次，详见 `sales_insight/README.md`。
2. **在手库存必须用 `qty_available`，不能用 `stock.quant` 汇总**。组合装（`x2` / `_GW`）
   是 **phantom BoM 套件**，没有自己的实物库存：quant 上恒为 0，只有 `qty_available`
   会从组件推算（基础款 106 → `x2` 显示 53）。2026-08-23 实测 **560 个套件里 557 个
   quant 为 0**，改口径后「低于安全库存」从 97 个降到 **59 个——38 个是套件造成的假警报**，
   而且假警报恰好落在运营配了安全库存的重点 SKU 上。
   `已预留` 与按仓拆列仍取 quant（`qty_available` 给不了这两个），故套件行这两列是 0；
   `实物库存` 列保留 quant 原值，与在手对照即可认出套件。
   ⚠ **在手列不可跨行求和**：套件与其基础款的在手是同一批实物。
   非套件产品两种口径**完全等价**（实测 314 个变化全是套件，非套件零变化）。
3. **产品范围取并集不取 `sale_ok` 单条**。`sale_ok=True ∪ 有销量 ∪ 有库存 ∪ 有补货规则`。
   只按 `sale_ok` 会把「已下架但仍有销量/仍有库存」的货静默漏掉，这类恰恰最该看见。
   报表里用 `在售` 列标出来。
4. **`stock.warehouse.orderpoint` 混着临时建议**。Replenishment 视图会现场生成
   `trigger='manual'` 的记录，它们不是人工配置的安全库存，必须滤掉，只取 `trigger='auto'`。
   **`trigger` 字段在 Odoo 14 就有了**（2026-08-23 实测：5 条 orderpoint 里 4 条是 manual，
   不滤的话 B 列 80% 是垃圾）。更早的版本没有该字段，代码按 `fields_get` 判断它在不在、
   自动走对应分支，**不要改成按版本号硬编码**；探针第 4 节会把两类各有多少条数出来。
5. **自定义字段技术名不许猜**。`Safety Stock` 的技术名（`x_studio_*` 之类）按**界面标签**
   经 `ir.model.fields` 反查。字段可能挂在 `product.template` 上，此时要经 `product_tmpl_id`
   取值而不是用 `product.product` 的 id。
6. **分页**。`search_read` 必须 limit/offset 循环，并且 `order='id'`——翻页期间有人改数据时，
   非稳定排序会漏行或重行。`read_group` 服务端一次返回全部分组，不分页；但按多字段分组
   必须 `lazy=False`。
7. **`sale.report` 字段跨版本改名**。数量/行数/金额都按候选名列表认（`fields_get` 查有没有），
   沿用 `packing_list` 德文列名那次的「别名元组」做法。
8. **多币种不做二次换算**。`sale.report.price_total` 已是 Odoo 按下单时汇率折算到公司本位币
   的值。排名按**数量**不按金额，币种对结论没有影响；金额列只作参考。
9. **多公司**。不传 `allowed_company_ids` 时读到的是「当前用户默认公司」，换台机器/换账号
   跑结果会变。探针会数出可见公司数并在 >1 时告警。
10. **计量单位**。`sale.report` 的数量是 Odoo 折算到产品参考 UoM 后的值，可跨订单行相加。
   若将来有产品按箱卖按瓶存，这里要重新核。
11. **超时**。`xmlrpc.client.ServerProxy` 默认**没有超时**，Odoo 一卡 launchd 任务就永远挂着。
    客户端自带了带超时的 Transport，默认 180 秒，网络类错误退避重试 2 次。

## 渠道口径：**必须显式指定，否则排名被 B 端主导**

这是本流水线最容易踩错、且错了不报错的地方。

2026-08-23 实测（窗口 2026-07-26 ~ 08-23，state in `sale,done`）：

| 口径 | 订单行数 | 销量(件) |
|---|---|---|
| 全渠道 | 5496 | **59,258** |
| 仅 VO/GW（天猫 C 端） | 5118 | **5,742** |
| 非 VO/GW（全是 `S0` 开头的 B 端单） | 378 | **53,516** |

B 端只占 378 行订单，却占 **90% 的件数**——整箱走货。订单号形如 `VO_TOF_SCP…` /
`GW_TOF_SCP…` / `S04007`，两类互补无交集。**而安全库存是给 C 端维护的**
（`sales_insight` 吃的手工导出就只保留 VO/GW）。两个口径混用会把人引向错误结论：
同一窗口下，「销量前 50 名里配了安全库存的」全渠道口径算出来是 3 个，
C 端口径算出来是 **45 个**。

- `--ref-contains VO,GW` 切到 C 端口径，此时并列多出 `销量_其它渠道` / `销量_全渠道` 两列
  ——「这个 SKU 卖 9100 件其实只有 200 件走天猫」一眼可见，那正是决定备货的关键信息。
- **`可撑周数` 一律按全渠道销量算**，不受筛选影响：库存被所有渠道一起消耗，
  只用 C 端销量算会高估覆盖（B 端一张整箱单就能把货搬空）。
- 不加筛选时终端会打提示。launchd plist 里默认带 `--ref-contains VO,GW`。

## 首次实测（2026-08-23，Odoo 14.0+e，窗口 2026-07-26 ~ 08-23）

环境比预想的简单：**单公司**（IHTCT GMBH，EUR）、**单仓库**（WH，7 个内部库位）。
`--company-id` 和 `--by-warehouse` 在这套环境里都用不上。

| 探测项 | 结果 | 对设计的影响 |
|---|---|---|
| 安全库存字段 | `product.product.safety_stock`（integer, stored） | **不是** `x_studio_*`，且挂在 `product.product` 上不是 template。按标签反查这个决定还本了 |
| orderpoint | 5 条，`auto` 1 条 / `manual` 4 条 | Odoo 14 就有 `trigger`；不滤的话 B 列 80% 是垃圾 |
| 两路安全库存 | A 152 个 / B **1** 个 | **B 基本没在用**。「不一致」的定义因此改了口径（见产出表） |
| `can be sold` | 10297 行 | 与 2026-08-02 的 10266 行一致（+31），口径没漂 |
| `sale.report` | `product_uom_qty` / `nbr` / `price_total` 都在；**无 `currency_id`** | 单公司 EUR，金额无需换算 |
| 近 4 周 state | done 5423 / sale 73 / draft 297 / cancel 183 | `state in ('sale','done')` 的过滤是必要的 |
| quant vs `qty_available` | 抽样 8 个全一致，**但抽样有偏**——按 quant 数量取前 8，套件永远抽不中 | 后来发现套件全错（见坑 2），探针的抽样已修：现在强制带上 quant=0 的套件 |

产出：10305 行产品、859 个有成交、97 个低于安全库存、**727 个有销量但没配安全库存**、
两路真冲突 1 条。全程 31 次 RPC。

**最值钱的一条**：销量前 20 名里配了安全库存的是 **0 个**，前 50 名里 3 个；
反过来 152 个配了安全库存的 SKU 销量排名中位数是 268。
**安全库存配置和实际销量几乎不重叠**——这正是把三路数据拉到一起才看得见的东西。

## 与手工导出的逐列对拍（2026-08-23 已做）

拿 `sales_insight` 2026-08-01 的产出 `销量排名.xlsx`（W28–W31，590 个 SKU、4391 件）
与 API 同窗口同渠道（`--since 2026-07-06 --until 2026-08-02 --ref-contains VO,GW`）对拍：

| 列 | 结果 |
|---|---|
| **安全库存** | 运营人工的 62 个 SKU，**59 个与 ERP 现值完全相等**；另外 3 个手工表里是 0、ERP 里为空（等价）。**坐实 API 读的就是 `sales_insight` 写的那个字段** |
| **销量** | 567 个共有 SKU 中 **324 个完全一致**，差异绝对值中位数 **0**；但 API 系统性偏多（213 个多、30 个少） |
| **排名** | Spearman **0.913**；前 20 名重叠 **19/20**、前 50 名 47/50、前 100 名 94/100。**决策层面是同一份报表** |
| **在手库存** | 不可比（手工那份是 08-01 快照，相隔三周） |

**销量偏多的原因：`sale.report` 是实时视图，不是快照。** 同一历史窗口过些天再拉，
数字会变——窗口内 4626 张已确认单里有 **228 张（4.9%）在导出之后被改动过**
（全部在导出前就已存在），另有 261 张后来被取消。把这批排除后差距从 +796 件收窄到 +504 件
（补回 37%）。剩下的 11.5% 无法归因——手工导出当时的确切筛选条件已不可复原。

排除过的**错误**假设，别再重查：① 窗口尾巴（截到 07-31 / 08-01 / 08-02 都是 5187 件，
8 月 1–2 日是周末无单）；② `--until` 日期被截断成 00:00:00（Odoo 会把只给日期的比较
自动展开成整天，实测两种写法行数一致）；③ 时区（按 Europe/Berlin 还原边界，行数不变）。

**结论：口径与字段映射都验证通过，可以用。** 但要记住这条性质——
**周报不可事后复现**，同一窗口重跑会得到不同数字。要留证据就留产出文件，别指望重跑复现。

## 与 sales_insight 的关系（当前）

**还没接。** 对拍已通过，接不接是下一个决定。在那之前 `sales_insight` 一个字都不动
——它每月在用，改它得单独立项。
