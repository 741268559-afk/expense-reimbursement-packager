# Agent Instructions

Read `SKILL.md` before handling a reimbursement request. Treat `SKILL.md` as the workflow contract and read `references/workflow.md` only when detailed field, invoice, or verification rules are needed.

Run scripts with the active Python interpreter. If required Python modules are unavailable, run `python scripts/bootstrap.py` and use the Python path returned in its JSON output for later commands.

Never invent amounts, approval fields, invoice data, exception resolutions, or reimbursement-form mappings. Stop at the corresponding review gate and present the generated report to the user.

Windows and macOS behavior is documented in `references/platforms.md`.
