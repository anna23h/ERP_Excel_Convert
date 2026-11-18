# invoice-reprice 使用说明（草稿发票按价格表批量改价）

**作用**：上传一张 Odoo **草稿**客户发票的导出 + 一张价格表的导出，Claude 生成一份可以直接导入 Odoo 的改价表，
外加一份核对表（每行旧价 / 新价 / 需要注意的地方）。Claude 不连 Odoo，也不会改 Odoo 里的数据，导入由你自己做。

## 1. 启用（只需一次）

**方式 A：claude.ai / Claude 桌面版（推荐）**

1. 打开 Anna 发来的共享链接，或在 claude.ai 的 **Customize → Skills** 里找到 `invoice-reprice`，把开关打开。
2. 在 **Settings → Capabilities** 里确认已打开 **Code execution and file creation**（代码执行与文件创建），没开的话 skill 跑不起来。

**方式 B：本地 Claude Code**（只在方式 A 用不了时用）

1. 把 `invoice-reprice.zip` 解压到 `C:\Users\<你的用户名>\.claude\skills\`，解压后应该是
   `...\.claude\skills\invoice-reprice\SKILL.md`。
2. 本机需要装好 Python 和 openpyxl：`pip install openpyxl`。
3. 重开 Claude Code。

## 2. 从 Odoo 导出两个文件

导出时选普通导出，**不要**勾「导入兼容」。建议各存一个导出模板，以后每次直接选。

- **发票**：会计 → 客户发票，只勾**一张**草稿发票 → 导出。字段：Status、ID、Invoice/Bill Date、Date、Partner、
  Untaxed Amount、Total、Invoice lines/ID、Invoice lines/Display Type、Invoice lines/Quantity、
  Invoice lines/Unit Price、Invoice lines/Discount (%)、Invoice lines/Product/Internal Reference、
  Invoice lines/Subtotal、Invoice lines/Total、Invoice lines/Label。
- **价格表**：导出**一张**价格表，价格条目作为子行导出。字段：Pricelist Name、Pricelist Items/Product/Internal Reference、
  Product/Name、Fixed Price、Min. Quantity、Start Date、End Date、Compute Price。

## 3. 调用

新开一个对话，上传这两个文件，然后说：

> 按这张价格表批量改这张草稿发票的价格

Claude 会给出一张对照表，以及三个文件：**导入表**、**核对表**、**改前快照**（三个都下载保存）。
标了 ⚠ 的行 Claude 会逐条说明，常见的有：价格表里没有这个产品（保持原价）、同一天有好几条有效价（取开始日期最新的一条，请你确认）。

## 4. 导入，然后核对

1. Odoo → 客户发票列表 → 导入「导入表」→ **先点「测试」**，没有报错再点导入。
2. 用**同一个导出模板**再导出这张发票。
3. 在同一个对话里上传这次的导出，说「核对一下改价结果」。如果换了新对话，要把「改前快照」和「导入表」一起上传。
   Claude 会确认除了单价和由它算出的金额以外，没有别的字段被改动。

## 先试一下

`samples/invoice-reprice/` 里有两份虚构的样例文件（`demo_invoice_export.xlsx`、`demo_pricelist_export.xlsx`），
可以先用它们试一遍。正确结果是：两行改价（3.50、5.40），一行「配不上」、保持原价。

有问题请告诉 Anna。
