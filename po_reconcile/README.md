# po_reconcile — 采购对账：采购 PO ↔ 财务 PO

**状态：算法已实现并通过构造数据测试；等真实干净数据验证后再上线**

## 要解决的问题

采购在一张 PO 里记录订购需求；供应商分批送货；财务按每批实收**另建一张 PO**
（产品与数量和供应商 invoice 一致）。于是采购那张 PO 永远显示原始订购量，
采购看不出实际还缺什么。

脚本比对两边，算出「还有什么没到」，并可回写采购 PO，让采购随时打开单子就看到当下缺什么。

## 两条前提（不成立就算不出正确答案）

> **甲** 财务单里的每一件货，都是对这张采购单的交付。
> **乙** 采购单是订购的完整记录 —— 行没被手工删改，也没有一部分货另开了单。

脚本会校验：一旦出现「财务单已收 > 订购量」，说明前提至少破了一条，
**拒绝生成回写导入表并以退出码 2 中止**（对账表仍产出，供排查）。
理由：负的未到量在满足前提的数据上不可能发生，此时回写只会把错误写进 ERP。

### 已知不满足前提的数据

`P11382`（2026 年 7 月「月度滚动 PO」实验的残骸）**两条都不满足**：部分订购行
被删/被移走；且它身上同时留着两套流程的痕迹 —— 一部分货靠挂在它自己身上的账单
结掉（走 Billed Qty），一部分货走财务另建 PO。这类数据无法自动对账，不是算法
问题，是「采购到底订了什么」这个事实在 ERP 里已经残缺。

## 关键设计

- **对账是 PO ↔ PO，不是 Quantity − Billed Qty**。新流程下账单挂在财务的 PO 上，
  **采购 PO 的 Billed Qty 永远停在 0**，用不了。Billed Qty 降为校验旁证。
- **按 SKU 汇总，不做行级分配**。采购只需知道「A 还差 5 件」，不需要知道 1 号那笔
  差 2、5 号那笔差 3 —— 行级分配产生的是记账噪音。汇总还顺带消噪：真实数据里
  逐行看「开票>订购」异常 2/31 条，汇总到 SKU 后只剩 1/19 个。
- **FIFO 摊回时摊的是「已收量」不是「未到量」**：先满足早下的订单。
  「1 号订 5、5 号订 3、到货 5」应该是 1 号那笔结清、5 号那笔全欠。
- **回写前自动归档原始订购量**到 `results/原始订购量归档-{PO}.xlsx`。回写用余量
  覆盖 Quantity 会造成两个后果：原始订购量在 ERP 里永久丢失；下次跑读到的
  「订购量」已是上次的余量，**再减一次会重复扣减**。归档同时解决这两件事 ——
  订购量基准取 `max(当前 Quantity, 归档值)`。
- **已到齐的行 Quantity 改 0、行保留**，不删行（删行不可逆，且已关联账单的行
  Odoo 可能拒绝删除）。

## 用法

```bash
# 手工指定财务单
python3 po_reconcile/po_reconcile.py <purchase.order.xlsx> \
    --buyer P11382 --finance P11416,P11665

# 财务在 Source Document 里填了采购 PO 号时，可省略 --finance 自动关联
python3 po_reconcile/po_reconcile.py <purchase.order.xlsx> --buyer P11382

python3 po_reconcile/test_po_reconcile.py        # 构造数据测试
```

产出（默认 `output/YYYYMMDD/`）：

| 文件 | 内容 |
|---|---|
| `采购对账表-{PO}.xlsx` | 逐 SKU 的 订购/财务单已收/未到量 + 采购单 Billed/Received 旁证 + 异常标记 |
| `回写导入表-{PO}.xlsx` | 逐 order line 的新 Quantity（未到量，FIFO 摊回）→ 上传 ERP |
| `results/原始订购量归档-{PO}.xlsx` | 回写前的订购量快照，防重复扣减 |

**首次上传务必先拿 1 条行试**，确认 ERP 更新的是既有行而不是新建行，再传全量。

## 导出字段要求

必需：`Order Reference` / `Order Lines/ID` / `Order Lines/Quantity` /
`Order Lines/Product/Internal Reference`

强烈建议加上：

- **`Order Lines/External ID`** —— 回写必需。Odoo 导入更新已有行靠行级 External ID；
  缺失时脚本降级用数据库 ID（列名 `.id`），能不能导入取决于 ERP 配置。
- **`Source Document`** —— 财务建 PO 时填采购的 PO 号，脚本即可自动关联，免去每次
  手工指定。字段在 Odoo 里叫 `origin`，默认只在「由其它单据生成」时才显示在表单上，
  需管理员在开发者模式下把它加进表单视图。

可以不导：`Activities`、`Product/Products/Bill of Materials`、`Product/Product/ID`（无信息量）。

## 解析要点（Odoo 锯齿状导出）

`Bills/Invoice lines/*` 与 `Order Lines/*` 是同一父记录下**两条互相独立的序列**，
行与行之间没有任何对应关系。切块只能靠单头列前向填充，**绝不能按行号对齐**。

另外：`Bills/Invoice lines/Quantity` 里**冲销发票的数量是正数**，方向只能靠
`Reference` 是否以 `Reversal of:` 开头判断（实测直接求和会多算 8）。Odoo 自己算
Billed Qty 时处理对了，导出的数量列没有。本脚本目前不依赖发票行，此坑留档备查。

## 尚未实现：财务 ↔ 供应商 Invoice 对账

供应商的 Invoice 通常覆盖多批发货、数量大于单张财务 PO。这块尚未见过任何样本，
另行立项。
