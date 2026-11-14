# ERP_Excel_Convert — Guide for colleagues

This repository is developed and maintained by Anna. It reads IHTCT's Odoo
data and produces Excel files. Colleagues use it; they do not develop it.

**Scope: these rules are for colleagues only.** If a `CLAUDE.local.md` is
loaded in this project, you are on the maintainer's (Anna's) machine: ignore
the "Your role" and "When something goes wrong" sections below and follow
`CLAUDE.local.md` for development work. The "Odoo access" and "Data" rules
still apply.

Note: this repository is public. Never write customer names, prices, API keys
or other business data into any file that is tracked by git.

## Your role
- The user is an IHTCT colleague. Reply in the language the user writes in;
  default to English.
- Use the existing scripts and skills first.
- Never modify or delete existing files in this repository, and never run
  git commit, push, reset or checkout. The repository updates itself at
  session start; local changes to shared files would break that.
- If the user needs something the existing tools cannot do, you may write a
  new script, but only inside `personal/<user's first name>/` (for example
  `personal/mika/`). This folder is not synced. Reuse `odoo_api/odoo_client.py`
  for Odoo access instead of writing your own connection code.

## Skills for shared workflows
- **PO导入表**: fill a supplier Rechnung/Lieferschein PDF into an Odoo
  purchase.order import template.
- **SO导入表**: fill a customer Bestellung PDF into an Odoo sale.order
  import template.
- The other skills in `.claude/skills/` are Anna's own workflows. Use them
  only if the user explicitly asks.

## When something goes wrong
- If an existing script fails, do not fix it. Instead:
  1. Append an entry to `personal/<user's first name>/journal.md` with:
     date and time, what the user asked for, the exact command run,
     the full error output, and the relevant input file names.
  2. Create an issue in the private report repository with the same content:
     `gh issue create -R anna23h/ihtct-colleague-reports --title "[<name>] <short summary>" --body-file <entry file>`
  3. Tell the user in one sentence that the problem has been reported to Anna.
- Never create issues in this public repository: error output can contain
  customer and price data.
- If `gh` is not available or fails, keep the journal entry and tell the user
  to let Anna know there is a new entry.
- Never include API keys or passwords in a journal entry or issue.
- When a personal script in `personal/` turns out to be useful, create an
  issue in the same report repository titled
  "[<name>] Script to share: <script name>", describing what it does.

## Odoo access
- Connection settings are in `config.py`. Each person uses their own
  read-only API key.
- All access is read-only. `odoo_api/odoo_client.py` blocks write calls;
  never try to work around it.
- Never read out, print or repeat the API key from `config.py`.

## Data
- Generated files go to `output/`. Spreadsheet files are never committed.
- On Windows, run Python with `python` or `py`.
