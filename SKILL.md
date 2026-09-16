---
name: expense-reimbursement-packager
description: Build finance-ready reimbursement packages from expense screenshots, payment receipts, invoices, and expense folders. Use when the user asks to整理报销, make reimbursement forms, classify expenses, calculate an invoice gap, add compliant replacement invoices, output the yellow reimbursement detail table, organize screenshots/invoices into a finance submission folder, or create print-ready files where three expense screenshots are combined per page and invoices are printed one per page.
license: MIT
---

# Expense Reimbursement Packager

This Skill is host-neutral. Run its Python scripts from the Skill directory with the active host's Python interpreter. For Windows, macOS, WorkBuddy, and other Agent installation/runtime details, read `references/platforms.md` (WorkBuddy may expose it as `@references/platforms.md`).

## Core Workflow

1. Collect all payment screenshots, invoices, ride itineraries, and other invoice attachments. Keep ride itineraries printable but exclude them from invoice value.
2. Extract and review one expense row per payment screenshot, normalize purposes/categories using `references/workflow.md`, assign `FY001`-style voucher IDs, write `报销金额确认单.md` and `报销异常清单.md`, then report the exact reimbursement total before final packaging.
3. Wait for the user's DingTalk approval screenshot. Transcribe only visible fields into an approval JSON and require the DingTalk approved amount to equal the confirmed screenshot total exactly. Never guess fields that are not visible.
4. Analyze invoice coverage against the confirmed reimbursement total. If coverage is short, stop with `needs_invoices` and report the exact gap. Accept finance-approved replacement invoices at the beginning or later, while keeping them labeled as `replacement`.
5. Build only after the expense total, unresolved exceptions, DingTalk form fields, and invoice coverage are ready. Put the reimbursement-form PDF first in the combined print PDF, followed by expense screenshots and then original invoices, replacement invoices, and ride itineraries. Add continuous page numbers, write `报销审计摘要.md`, and verify every amount, formula cache, page count, invoice total, and final zip before reporting the pack ready.

## Use The Script

For one-command folder packaging, run:

```bash
python scripts/package_from_folder.py --input /path/to/source-folder --project-name 项目名 --reimburser 报销人 --out /path/to/output-folder
```

This command writes `preflight_report.json`, `报销金额确认单.md`, `报销异常清单.md`, and `钉钉信息缺口.md`, then analyzes invoice coverage. It builds only when expense rows and exceptions are reviewed, DingTalk metadata is complete, and recognized invoices cover the reimbursement total. Otherwise `package_result.json` returns the exact pending state and report paths. Add `--stop-on-preflight-warnings` when the team wants to stop and inspect duplicate files, PDFs mixed into screenshot folders, or unsupported files before packaging.

After the user submits DingTalk, transcribe the visible fields into JSON and rerun:

```bash
python scripts/package_from_folder.py --input /path/to/source-folder --project-name 项目名 --reimburser 报销人 --approval-metadata /path/to/approval_metadata.json --out /path/to/output-folder
```

The approval JSON normally contains `dingding_number`, title-derived `reimbursement_type`, `reimburser`, `payee`, `invoice_entity`, and `approved_amount`. `approved_amount` is a hard equality check against the screenshot total.

For repeat users, pass `--profile /path/to/reimbursement_profile.json` or keep `03_profile/reimbursement_profile.json` beside the team starter. A profile may store only stable defaults such as reimburser, payee, invoice entity, and company. Never store an approval number or project amount in the profile.

For messy inputs, first run:

```bash
python scripts/preflight_inputs.py --input /path/to/source-folder-or.zip --invoice-input /path/to/invoices-or.zip --out /path/to/preflight_report.json
```

Review `preflight_report.json` to confirm which files will be treated as expense screenshots, invoices, or ignored files. Direct `.zip` inputs are scanned in place and shown as `archive.zip::member/path.png` in the report.

If OCR leaves blanks or flags `NEEDS_REVIEW`, the one-command script stops and writes both `draft_manifest.json` and an editable `draft_manifest_review.xlsx`.

Prefer the Excel review workbook for team review:

```bash
python scripts/package_reviewed_workbook.py --manifest /path/to/output-folder/draft_manifest.json --workbook /path/to/output-folder/draft_manifest_review.xlsx --out /path/to/output-folder
```

This applies the edited workbook to `reviewed_manifest.json`, builds the pack, runs verification, and refreshes `package_result.json`.

The one-command flow uses the validated built-in `assets/费用报销单模板.xlsx` and `assets/form_cells.json` automatically. To override it with another `费用报销单` Excel template, include it in the command:

```bash
python scripts/package_from_folder.py --input /path/to/source-folder --project-name 项目名 --reimburser 报销人 --out /path/to/output-folder --form-template /path/to/费用报销单.xlsx
```

The one-command script inspects an override template, adds recommended `reimbursement_form` mappings to the manifest, then builds the filled form when the mappings are usable. For another fixed team template, pass a reviewed mapping with `--form-cells /path/to/form_cells.json`; if that JSON contains `template`, `--form-template` may be omitted.

If `--form-cells` is omitted, `package_from_folder.py` first checks explicit arguments, environment variables, the current folder, platform configuration folders, and the legacy Codex configuration folder. If none is present, it falls back to the built-in `assets/form_cells.json`. Set `EXPENSE_REIMBURSEMENT_HOME` for a host-neutral shared location. The selected path and origin are recorded in `package_result.json`.

The bundled form has five expense rows. It dynamically writes nonzero categories in manifest order and refuses to generate when all six default categories are nonzero, instead of silently merging categories.

`--input` can be a folder or a `.zip` archive. If invoices are in a separate folder, file, or `.zip`, add one or more `--invoice-input` arguments:

```bash
python scripts/package_from_folder.py --input /path/to/screenshots --invoice-input /path/to/invoices --project-name 项目名 --reimburser 报销人 --out /path/to/output-folder
```

If the first run returns `needs_invoices`, read `发票缺口清单.md`, collect approved replacement invoices for at least the reported gap, and rerun into a fresh output folder:

```bash
python scripts/package_from_folder.py --input /path/to/screenshots --invoice-input /path/to/original-invoices --replacement-invoice-input /path/to/approved-replacement-invoices --project-name 项目名 --reimburser 报销人 --out /path/to/final-output-folder
```

Replacement invoices remain labeled `替代发票` in `03_发票`, are included in `发票_一张一页.pdf` and `全部打印_一键打印.pdf`, and are counted separately in the coverage report. Use `--invoice-policy none` only when the finance policy for this reimbursement explicitly requires no invoice coverage.

Ride itineraries and similar invoice attachments detected in the source folder are stored as `supporting_documents`. They are copied to `03_发票` with a `发票附件_行程单` label and included in both print PDFs, but they contribute `0.00` to invoice coverage. Use `--supporting-document-input` when such files are supplied separately.

Image inputs can include `.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.tif`, `.tiff`, `.heic`, and `.heif`. HEIC/HEIF files are converted to PNG for Excel and PDF output.

For a manual two-step flow, first draft a manifest:

```bash
python scripts/draft_manifest_from_folder.py --input /path/to/source-folder --project-name 项目名 --reimburser 报销人 --out /path/to/manifest.json
```

Review every row in the draft manifest, especially rows marked `NEEDS_REVIEW`, then create the final pack:

```bash
python scripts/analyze_invoice_coverage.py --manifest /path/to/manifest.json --out /path/to/invoice_coverage.json --update-manifest
python scripts/build_reimbursement_pack.py --manifest /path/to/manifest.json --out /path/to/output-folder
```

The build command refuses a `full_amount` manifest unless `invoice_coverage.status` is `covered`.

For manual builds, verify the final pack:

```bash
python scripts/verify_reimbursement_pack.py --pack /path/to/output-folder --manifest /path/to/manifest.json
```

The script creates:

- `报销表/<项目>_<报销人>_报销明细表.xlsx`
- `报销表/<项目>_<报销人>_费用报销单.xlsx`, using the built-in form unless overridden
- `打印/<项目>_费用报销单_打印.pdf`, using the built-in form unless overridden
- `打印/<项目>_费用截图_三张一页.pdf`
- `打印/<项目>_打印排版.xlsx`
- `打印/<项目>_发票_一张一页.pdf` when invoices exist
- `打印/<项目>_全部打印_一键打印.pdf`, with the reimbursement form first, expense screenshots second, and invoices/ride itineraries last
- `财务提交文件夹/` with copied/renamed reimbursement forms, screenshots, invoices, print files, and `manifest.json`
- `<项目>_<报销人>_财务提交文件夹.zip`, a clean zip of the finance folder for direct submission
- `交付清单.md`, a human-readable handoff checklist listing the finance zip, reimbursement tables, one-click print PDF, screenshot/invoice counts, page counts, and category totals
- `报销金额确认单.md`, the pre-DingTalk row-by-row amount confirmation with voucher IDs and category subtotals
- `报销异常清单.md`, the refund, split/personal, foreign-currency, multi-amount, and discount review surface
- `钉钉信息缺口.md`, the required DingTalk fields and amount-match gate
- `报销审计摘要.md`, the concise final expense, invoice, exception, print-page, and verification summary
- `package_result.json`, a final one-command result index with status, verification status, preflight warnings, finance zip, delivery checklist, reimbursement workbook, expense form, and one-click print PDF paths
- `preflight_report.json`, an input audit showing which source files were treated as expense screenshots, invoices, or ignored files
- `invoice_coverage.json`, the required, recognized, missing, excess, original-invoice, and replacement-invoice reconciliation
- `发票缺口清单.md`, the human-readable first-stage invoice gap and invoice-file status list
- `draft_manifest_review.xlsx` when rows need review, so the team can correct date, time, purpose, category, amount, and notes in Excel before applying the edits back to JSON
- `verification.json` when `package_from_folder.py` runs verification after building

Use the bundled script for deterministic output once the manifest is ready. Host vision tools may help review screenshots, but the manifest and generated files must still pass the deterministic checks.

To verify a local installation before processing real reimbursements, run:

```bash
python scripts/self_test.py --out /path/to/self-test-output
```

The self-test creates synthetic screenshots, a mock invoice, a test `费用报销单` template, a portable `form_cells.json`, a manifest review workbook, a first-stage missing-invoice result, a replacement-invoice final pack, a tamper-rejection test, and `verification.json`.

For a faster team-readiness check, or after editing an override `form_cells.json`, run:

```bash
python scripts/check_readiness.py --out /path/to/readiness_report.json
```

Without `--form-cells`, readiness validates the built-in form and should return `ready`. Pass `--form-cells` only to validate an override.

## Manifest Rules

Required top-level fields:

- `project_name`: project or shoot name, such as `新品发布会拍摄`
- `reimburser`: person being reimbursed
- `entries`: expense rows

Optional top-level field:

- `approval_metadata`: fields transcribed from a DingTalk approval screenshot. Only use values visibly present in the screenshot or stable defaults explicitly authorized in the profile. Supported keys are `dingding_number`, `reimbursement_type`, `reimburser`, `payee`, `invoice_entity`, `approved_amount`, `company`, `reimbursement_project`, `contract_number`, `reimbursement_reason`, and `business_line`.
- `approval_requirement`: defaults to `policy: required` with required fields `dingding_number`, `reimbursement_type`, `payee`, `invoice_entity`, and `approved_amount`.
- `expense_confirmation`: records the amount confirmed by the matching DingTalk approval.
- `expense_exceptions`: generated exception records; blocking items require an entry-level `exception_resolution`.

Required entry fields:

- `date`: `YYYY-MM-DD`
- `purpose`: finance-facing purpose, not merchant name
- `expense_type`: one of the default categories: `交通费`, `餐饮费`, `设备费`, `场地费`, `演员费`, or `其他费用`
- `amount`: numeric amount in RMB
- `screenshot`: path to the payment/expense screenshot

Optional entry fields: `voucher_id`, `time`, `merchant`, `notes`, `exception_resolution`, `invoices`. New folder drafts assign voucher IDs automatically.

Top-level invoice fields:

- `invoices`: invoice files not yet matched to a specific expense row. These are copied to `03_发票` and included in the one-invoice-per-page print PDF.
- `invoice_requirement.policy`: `full_amount` by default, or `none` only when finance explicitly waives invoice coverage.
- `invoice_items`: one object per invoice with `path` and `role` (`original` or `replacement`). Analysis adds recognized amount, invoice number, date, buyer-name/entity check, linked voucher IDs, SHA-256, and status. Tax-platform authenticity remains an external finance check.
- `invoice_coverage`: written by `analyze_invoice_coverage.py`; final builds require `status: covered` under `full_amount` policy.
- `supporting_documents`: invoice attachments such as ride itineraries. They are preserved and printed but excluded from recognized invoice amounts.

Optional form field:

- `reimbursement_form.template`: path to the selected `费用报销单` template; the one-command flow supplies the built-in template by default.
- `reimbursement_form.sheet`: sheet name to fill; defaults to the active sheet.
- `reimbursement_form.cells`: mapping from cell references to literal values or tokens such as `{{project_name}}`, `{{reimburser}}`, `{{dingding_number}}`, `{{reimbursement_type}}`, `{{payee}}`, `{{invoice_entity}}`, `{{total_amount}}`, `{{date_range}}`, `{{total_print_pages}}`, `{{交通费_amount}}`, and `{{交通费_count}}`.
- `reimbursement_form.category_rows`: optional reusable list of `{label_cell, amount_cell}` slots. The builder fills only nonzero categories in manifest order and stops if the template has too few rows, rather than silently combining categories.

## Reimbursement Form Templates

If the user provides an additional `费用报销单` template, inspect it first, identify the target cells, add a `reimbursement_form` mapping to the manifest, then run the script. Keep this as a separate output under `报销表/` and copy it into `财务提交文件夹/01_报销表/`. When no override is provided, use the built-in template and mapping.

To inspect a new Excel form template and draft cell mappings, run:

```bash
python scripts/inspect_reimbursement_form.py --template /path/to/费用报销单.xlsx --out /path/to/form_inspection.json --cells-out /path/to/form_cells.json --relative-template
```

Or use the onboarding wrapper to copy the template into a portable config folder and write an onboarding report:

```bash
python scripts/onboard_reimbursement_form.py --template /path/to/费用报销单.xlsx --out /path/to/form-template-config --copy-template
```

The onboarding wrapper also writes `form_config_validation.json` and `readiness_report.json`. Review `recommended_manifest_reimbursement_form.cells` or `form_cells.json`, adjust it against the visible template, then validate the reviewed config:

```bash
python scripts/validate_form_config.py --form-cells /path/to/form-template-config/form_cells.json --out /path/to/form-template-config/form_config_validation.json
```

Keep the reviewed and validated `form_cells.json` beside the template and reuse it:

```bash
python scripts/package_from_folder.py --input /path/to/source-folder --project-name 项目名 --reimburser 报销人 --out /path/to/output-folder --form-cells /path/to/form_cells.json
```

When running from a team starter folder that contains `02_form_template_config/form_cells.json`, `--form-cells` can be omitted:

```bash
python scripts/package_from_folder.py --input /path/to/source-folder --project-name 项目名 --reimburser 报销人 --out /path/to/output-folder
```

If the template is moved separately from `form_cells.json`, update the `template` value in that JSON or pass the new path with `--form-template`.

## Verification

Before finalizing:

- Run `scripts/verify_reimbursement_pack.py --pack <output-folder> --manifest <manifest.json>` after all files are generated and require `status: passed`. Treat this as a delivery gate: do not present a failed pack as finance-ready.
- Require every manifest amount to be positive and precise to cents, reject duplicated screenshot references, reject rows still marked `NEEDS_REVIEW`, and confirm an OCR-backed amount appears in that row's OCR text when OCR text is available.
- Compare every workbook detail row with the sorted manifest, including sequence, date, purpose, reimburser, category, amount, and invoice filename.
- Independently reconcile the manifest total, pack summary total, workbook detail sum, workbook category summaries and counts, workbook summary total, workbook footer total, finance-folder manifest, and delivery-checklist category totals. Require exact equality at `0.01` CNY precision.
- Scan the reimbursement workbook for invalid numeric cells and Excel error values. Record the reconciled amounts in `verification.json.financial_reconciliation`.
- Require recognized, non-duplicate invoice totals to cover the required amount. Record required, recognized, missing, and replacement amounts in `verification.json.invoice_reconciliation`.
- Confirm replacement invoices remain labeled in `03_发票`, their count matches the manifest and coverage report, and both invoice reports are included in the finance zip.
- Confirm screenshot count in the workbook equals the number of entries with screenshots.
- Confirm print screenshot PDF page count equals `ceil(expense screenshot count / 3)`.
- Confirm invoice PDF page count is at least the number of invoice files when invoices are present.
- Confirm combined one-click print PDF page count equals expense-form, screenshot, and invoice/attachment pages; confirm every page has a continuous `current / total` footer.
- Confirm the DingTalk approved amount and stored confirmation amount equal the payment-screenshot total exactly.
- Confirm voucher IDs are unique and appear in the detail workbook and confirmation report.
- Confirm `报销金额确认单.md`, `报销异常清单.md`, `钉钉信息缺口.md`, and `报销审计摘要.md` exist and reconcile to the manifest.
- Confirm `交付清单.md` exists and points to the finance zip and one-click print PDF.
- Confirm `package_result.json` has status `passed` after one-command packaging.
- Confirm reusable `form_cells.json` passes `scripts/validate_form_config.py` before team reuse.
- Prefer `scripts/check_readiness.py --form-cells <form_cells.json>` on each teammate machine before first real use.
- When `package_from_folder.py` stops with `needs_review`, confirm `draft_manifest_review.xlsx` exists and use `scripts/package_reviewed_workbook.py` after Excel edits to apply, build, verify, and update `package_result.json`.
- When a `费用报销单` template is configured, confirm mapped form cells were populated with manifest values.
- When the template contains total or Chinese-uppercase amount formulas, confirm their cached display values match the calculated reimbursement total so previews and direct printing do not show stale values.
- Confirm the generated finance zip contains `01_报销表`, `02_费用截图`, `03_发票`, `04_打印文件`, and no `._`, `.DS_Store`, or `__MACOSX` files.
- Open or preview the first generated page/workbook when possible to ensure images are visible.
- If using `draft_manifest_from_folder.py`, treat OCR as a draft only and visually verify dates, purposes, categories, and amounts before final packing.
