# Claude Code setup for colleagues

> **Status: outline (2026-11-14).** Steps marked `TODO` are filled in after the colleague permission settings
> have been tested. Maintained by Anna; this file in the repository is the source version.

This guide sets up the Claude desktop app's **Code** tab on a colleague's computer so they can use this
repository's skills (PO导入表, SO导入表) to work with Odoo data. Who does each step is shown in brackets.

| | Daniel | Mika |
| --- | --- | --- |
| Computer | Mac | Windows |
| Project folder | `~/ERP_Excel_Convert` | `C:\Users\<name>\ERP_Excel_Convert` |
| Python in the project | `.venv/bin/python` | `.venv\Scripts\python` |
| Claude Code settings file | `~/.claude/settings.json` | `%USERPROFILE%\.claude\settings.json` |

## 0. Before installation (Anna)

All of these must be true before starting:

- [ ] The read-only Odoo API user exists, and this colleague's personal key has been checked with
      `personal/anna/probe_rights.py` (result: read-only).
- [ ] Odoo is reachable over HTTPS, or Daniel has decided to stay on HTTP for now.
- [ ] The colleague is a member of the company's Claude Team.
- [ ] GitHub invitations to `ERP_Excel_Convert` and `ihtct-colleague-reports` have been sent.

## 1. Accounts (colleague)

1. Create a GitHub account and accept both repository invitations.
2. Accept the Claude Team invitation and sign in with the company email address.

## 2. Install software (colleague, Anna assists)

**Mac (Daniel)** — Anna runs the terminal steps.

1. Git: `xcode-select --install`
2. Python 3: TODO (python.org installer or Homebrew)
3. Claude desktop app
4. GitHub CLI: TODO, then `gh auth login`

**Windows (Mika)**

1. Git for Windows (includes Git Bash, which the session-start pull script needs)
2. Python 3 — tick "Add Python to PATH" during installation
3. Claude desktop app
4. GitHub CLI, then `gh auth login`

## 3. Get the project (colleague)

1. Clone the repository into the project folder from the table above.
2. Always open the project in the Code tab **through the folder picker**, never by typing the path:
   Claude Code stores session history by the exact path spelling.
3. Create the Python environment and install the requirements: TODO (reuse the logic of
   `Mac双击运行.command` / `Windows双击运行.bat`).

## 4. Odoo connection (Anna and colleague together)

1. Copy `config.example.py` to `config.py` and fill in the address, database, the read-only user's login
   and the colleague's personal key.
2. Type the key in person. Never send it by email or chat.
3. Run the check script; the result must be read-only.

## 5. Claude Code permissions (Anna)

Create the settings file from the table above with the colleague permission settings: TODO (insert the
tested version). They keep shared code read-only, allow new scripts only in `personal/<name>/`, block
reading `config.py`, and block git write commands, file deletion and network commands.

## 6. First start (colleague, Anna checks)

1. Claude desktop app → Code tab → open the project folder → sign in with the company account.
2. Check: `/memory` lists only `./CLAUDE.md` (no `CLAUDE.local.md`).
3. Check: the session start shows "Pulled, up to date".

## 7. Acceptance test (Anna and colleague)

- [ ] Run PO导入表 once with a sample PDF.
- [ ] Ask Claude to change a shared file → it is refused.
- [ ] Ask Claude to create a test script in `personal/<name>/` → it works.
- [ ] Trigger a test error report → the issue appears in `ihtct-colleague-reports`.

## 8. Daily use (for the colleague)

TODO: how to start a session, example questions, where result files are saved, what happens when a script
fails (journal entry + issue to Anna).

## 9. Maintenance and offboarding (Anna)

TODO: replacing or revoking a key, removing repository access, updating the permission settings.
