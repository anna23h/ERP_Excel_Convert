# invoice-reprice - how to use (reprice a draft invoice from a pricelist)

**What it does**: you upload an Odoo export of one **draft** customer invoice plus an export of one pricelist.
Claude produces an import file that sets the new unit prices in Odoo, and a check sheet (old / new price per
line, anything that needs attention). Claude never connects to Odoo; you do the import yourself.

## 1. Enable it (once)

**Option A - claude.ai / Claude desktop app (preferred)**

1. Open the share link from Anna, or find `invoice-reprice` under **Customize > Skills** on claude.ai and
   switch it on.
2. Under **Settings > Capabilities** make sure **Code execution and file creation** is on - the skill needs it.

**Option B - local Claude Code** (only if option A does not work for you)

1. Unzip `invoice-reprice.zip` into `~/.claude/skills/` (Windows: `C:\Users\<you>\.claude\skills\`), so you
   end up with `.../.claude/skills/invoice-reprice/SKILL.md`.
2. Python with openpyxl is needed: `pip install openpyxl`.
3. Restart Claude Code.

## 2. Export two files from Odoo

Use the normal export (do **not** tick "I want to update data / import-compatible"). Saving an export template
for each makes this a one-click step next time.

- **Invoice**: Accounting > Customer Invoices, tick **one** draft invoice > Export. Fields: Status, ID,
  Invoice/Bill Date, Date, Partner, Untaxed Amount, Total, Invoice lines/ID, Invoice lines/Display Type,
  Invoice lines/Quantity, Invoice lines/Unit Price, Invoice lines/Discount (%),
  Invoice lines/Product/Internal Reference, Invoice lines/Subtotal, Invoice lines/Total, Invoice lines/Label.
- **Pricelist**: export **one** pricelist with its items as sub-rows. Fields: Pricelist Name,
  Pricelist Items/Product/Internal Reference, Product/Name, Fixed Price, Min. Quantity, Start Date, End Date,
  Compute Price.

## 3. Use it

Start a new chat, upload both files and write for example:

> Reprice this draft invoice with this pricelist

You get a summary table and three files: **import file**, **check sheet**, **before-snapshot** - download all three.
(File names and sheet headers are in Chinese; Claude explains them in English.) Lines marked ⚠ are explained one
by one, typically: product not in the pricelist (price kept), or several prices valid on the invoice date
(the one with the latest start date was taken - please confirm).

## 4. Import, then verify

1. Odoo > Customer Invoices list > Import the import file > click **Test** first, then **Import**.
2. Export the same invoice again **with the same export template**.
3. Upload that export in the same chat and ask "verify the repricing". In a new chat also upload the
   before-snapshot and the import file. Claude confirms that nothing but the unit prices and the totals
   derived from them changed.

## Try it first

`samples/invoice-reprice/` has two made-up sample files (`demo_invoice_export.xlsx`,
`demo_pricelist_export.xlsx`). Expected result: two lines repriced (3.50 and 5.40), one line "not in
pricelist" keeping its price.

Questions: ask Anna.
