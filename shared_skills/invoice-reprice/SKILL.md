---
name: invoice-reprice
description: Reprice a draft Odoo customer invoice from an Odoo pricelist. Use when the user uploads two Odoo exports - one draft invoice (account.move*.xlsx) and one pricelist (product.pricelist*.xlsx) - and asks to "reprice the invoice", "update invoice prices from the pricelist", "Rechnungspreise nach Preisliste ändern", "批量改账单价格" or similar. Produces an Odoo import file (3 columns) that sets the new unit prices, a check sheet (old vs new price per line, warnings) and a before-snapshot. After the user imports and re-exports the invoice, the same skill verifies that only prices and the totals derived from them changed. Works on files only; never connects to Odoo.
---

# Invoice reprice

Change the unit prices of **one draft** customer invoice to the prices of **one** Odoo pricelist.
All logic is in `scripts/invoice_reprice.py` (inside this skill's folder). Run that script; do not
write your own calculation, do not edit the input files, do not hand-edit the import file.

Reply in the language the user writes in. The script prints in Chinese: translate its messages and
table headers for the user (glossary at the end), keep codes and numbers exactly as printed.

## Requirements

Python 3 with `openpyxl` (`pip install openpyxl` if `import openpyxl` fails).
Run with `python3`; on Windows use `python` if `python3` is not found.

## Step 1 - check the two input files, then run build

The user exports both files from Odoo with the normal export (**not** "import-compatible"):

- **Invoice**: Accounting > Customer Invoices, tick **one** draft invoice > Export. Needed columns:
  Status, ID (external ID), Invoice/Bill Date, Date, Partner, Untaxed Amount, Total, Invoice lines/ID,
  Invoice lines/Display Type, Invoice lines/Quantity, Invoice lines/Unit Price, Invoice lines/Discount (%),
  Invoice lines/Product/Internal Reference, Invoice lines/Subtotal, Invoice lines/Total, Invoice lines/Label.
- **Pricelist**: export **one** pricelist with its items as sub-rows. Needed columns:
  Pricelist Items/Product/Internal Reference, Fixed Price, Start Date, End Date
  (Pricelist Name, Product/Name, Min. Quantity, Compute Price are used if present).

When both files are there, run it right away:

```bash
python3 <skill-folder>/scripts/invoice_reprice.py build "<invoice.xlsx>" "<pricelist.xlsx>" -o "<output-dir>"
```

`<output-dir>`: where the user can get the files back. In claude.ai use `/mnt/user-data/outputs`;
in Claude Code use `output` in the current working directory. Three files are written:
`发票改价_导入表_<id>.xlsx` (import file), `发票改价_核对表_<id>.xlsx` (check sheet),
`发票改价_改前快照_<id>.xlsx` (before-snapshot). If no line changes price, there is no import file.

**Exit code 2 (line starts with `✗`) = stop and tell the user the reason.** Common cases:

| Message says | Tell the user |
|---|---|
| status is not draft (不是草稿) | Reset the invoice to draft in Odoo, export again |
| should be exactly one invoice / one pricelist (应当正好一张 / 一次只给一张) | Export only one, try again |
| column not found (找不到「…」列) | Re-export with the fields listed above |

Never work around an error by changing the input files or the script.

## Step 2 - hand over the result

1. Show the Markdown table the script printed (translated headers), then the file paths / downloads.
2. Pass on every ⚠ message (e.g. price date fell back to the accounting date; pricelist name does not
   contain the customer name - maybe the wrong pricelist).
3. Explain **only the ⚠ lines**. No ⚠ at all: "all lines matched, repriced per the pricelist".
4. Next step for the user: in Odoo **Customer Invoices** list > Import the import file >
   click **Test** first, then Import. Afterwards export the same invoice again **with the same export
   template** and give that file back for checking.

## Warnings (check sheet columns 结果 / 备注)

| Marker | Meaning | What to say |
|---|---|---|
| `配不上` not in pricelist | Pricelist has no item with this **exact** product code (`_GW` and non-GW are different products) | Price kept; ask whether the pricelist is missing this product |
| `无有效价` no valid price | Code exists, but the price date is outside every item's start/end | Price kept; list it for the user to confirm |
| `多条有效` several valid | Several items valid on that date; the one with the latest Start was taken | Read out the candidates (有效候选) and ask the user to confirm |
| `非固定价` not fixed price | Item is not a Fixed Price rule | Price kept; formula prices are not calculated |
| `无产品编码` no product | Invoice line has no product | Price kept |
| tax rate not 0 (税率 …) | Line has tax; the pricelist price goes **directly** into the unit price | Ask the user to confirm the pricelist price uses the same basis (tax-incl./net) as shown in 含税口径 |
| `口径无法判断` basis unclear | Subtotal/Total do not fit unit price × quantity | Report the line, ask the user to look at it |

## Step 3 - after the import: verify

When the user gives the re-exported invoice:

```bash
python3 <skill-folder>/scripts/invoice_reprice.py verify "<before-snapshot.xlsx>" "<after-export.xlsx>" --expect "<import-file.xlsx>"
```

- Exit 0: "Apart from the unit prices and the subtotals/totals recalculated from them, nothing changed."
- Exit 2: show the ⚠ list as printed (translated). Do not guess causes, do not suggest changing things
  back in Odoo - the user checks it.
- "before" must be the before-snapshot from build (or the original export); both exports must use the
  same export template.

## Fixed rules (do not change)

- Match on the **full product code**, never on PZN or product name.
- Price date = invoice date; if empty (draft), the accounting date `Date`. Valid = Start <= date <= End
  (empty = open). Several valid: latest Start, flagged. Not found / no valid price: keep the old price, flagged.
  Min. Quantity is shown only, not used.
- Pricelist price goes straight into `price_unit`; tax basis is derived per line from Subtotal/Total.
- Only lines whose price changes go into the import file; note and section lines never do.
- Import file has exactly 3 columns `id` / `invoice_line_ids/id` / `invoice_line_ids/price_unit`,
  invoice ID only on the first row.

## Never

- Never connect to Odoo or ask for API keys - if data is missing, ask the user for an export.
- Never build or edit an import file by hand (re-importing a modified full export fails; editing
  journal items breaks the invoice balance). Only use the file the script writes.
- Never pick between several candidate prices for the user, never match "not in pricelist" lines by
  PZN or name.

## Glossary (script output is Chinese)

| Chinese | English |
|---|---|
| 发票 / 价格表 / 取价日 | invoice / pricelist / price date |
| 发票日期 / 记账日期 | invoice date / accounting date |
| 产品编码 / 品名 / 数量 | product code / product name / quantity |
| 旧价 / 新价 / 价差 / 新小计 | old price / new price / difference / new subtotal |
| 结果 / 改价 / 备注 | result / price changed / notes |
| 匹配 | matched |
| 含税口径: 0税 / 含税 / 净价 / 无法判断 | tax basis: 0 % tax / tax-incl. / net / unclear |
| 有效候选 / 最小量 | valid candidates / min. quantity |
| 导入表 / 核对表 / 改前快照 | import file / check sheet / before-snapshot |
| 产品行 N 行，改价 M 行 | N product lines, M repriced |
| 允许范围内的变化 / 需要人看的 | expected changes / needs a human look |
