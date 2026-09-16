# Reimbursement Workflow Reference

## Output Contract

Create a finance-ready pack with:

1. A yellow reimbursement detail workbook matching the existing template:
   - Sheet name: `报销明细`
   - Title row: `<项目><报销人> 报销明细表`
   - Default summary rows: `交通费`, `餐饮费`, `设备费`, `场地费`, `演员费`, `其他费用`, `总  计`
   - Detail columns: `凭证编号`, `日期`, `用途`, `报销人`, `费用类型`, `金额(元)`, `凭证截图`, `发票`
   - Category totals and counts must match the manifest and be visible without requiring Excel recalculation.
2. A filled `费用报销单` when the user provides its template:
   - Inspect the template visually and by cell values.
   - Add `reimbursement_form.template`, optional `sheet`, and `cells` mapping to the manifest.
   - Use tokens in cells when helpful: `{{project_name}}`, `{{reimburser}}`, `{{total_amount}}`, `{{expense_count}}`, `{{date_range}}`, `{{date_start}}`, `{{date_end}}`, `{{交通费_amount}}`, `{{餐饮费_amount}}`, `{{设备费_amount}}`, `{{场地费_amount}}`, `{{演员费_amount}}`, and `{{其他费用_amount}}`.
   - Do not guess a cell mapping from labels alone; inspect the workbook and verify the filled result.
3. A finance submission folder containing:
   - `01_报销表`
   - `02_费用截图`
   - `03_发票`
   - `04_打印文件`
   - `manifest.json`
   - `invoice_coverage.json`
   - `发票缺口清单.md`
   - `报销金额确认单.md`
   - `报销异常清单.md`
   - `钉钉信息缺口.md`
   - `报销审计摘要.md`
   - `交付清单.md`
   - A clean zip copy named `<项目>_<报销人>_财务提交文件夹.zip`
4. Print-ready files:
   - Three expense screenshots per A4 landscape page, sorted by date/time.
   - Invoices as one invoice per A4 page. Keep source invoice PDFs intact when possible.
   - A combined `<项目>_全部打印_一键打印.pdf`, with the reimbursement-form PDF first, expense screenshots second, and original invoices, replacement invoices, and ride itineraries last, for a single print job.
   - Stamp every combined-PDF page with a continuous `current / total` page number and require the actual total to match the form's `单据及附件共 N 页` value.
5. Automated verification:
   - Run `scripts/verify_reimbursement_pack.py --pack <output-folder> --manifest <manifest.json>`.
   - Write `verification.json` in the output root, beside `pack_summary.json`.
   - Treat verification as a hard delivery gate. Reject non-positive or over-precision amounts, duplicated screenshot references, rows still marked `NEEDS_REVIEW`, and OCR-backed amounts that are absent from available OCR text.
   - Reconcile every detail row and all monetary controls: manifest, `pack_summary.json`, workbook detail sum, category amounts and counts, workbook summary total, workbook footer total, finance-folder manifest, and checklist category totals.
   - Store the CNY reconciliation at cent precision in `verification.json.financial_reconciliation` and fail if any value differs.
   - Require `status: passed` before final delivery.
6. Input audit:
   - One-command packaging writes `preflight_report.json` in the output root before OCR/building.
   - Use the report to confirm which source files became expense screenshots, invoices, or ignored files.
7. Handoff checklist:
   - Write `交付清单.md` in the output root and finance submission folder.
   - Include the finance zip path, reimbursement table path, expense-form status, one-click print PDF, screenshot/invoice counts, page counts, total amount, and category totals.
8. Final one-command result:
   - Write `package_result.json` in the output root after `package_from_folder.py` stops for review or finishes building.
   - For successful runs, require `status: passed`, `verification_status: passed`, and non-empty paths for `delivery_checklist`, `finance_zip`, `reimbursement_workbook`, and `one_click_print_pdf`.
9. Excel manifest review:
   - When a folder draft stops with `needs_review`, write `draft_manifest_review.xlsx` beside `draft_manifest.json`.
   - Reviewers should edit date, time, purpose, expense type, amount, and notes in Excel. Keep `用途` finance-facing and do not add a merchant column.
   - After review, prefer `scripts/package_reviewed_workbook.py --manifest <draft_manifest.json> --workbook <draft_manifest_review.xlsx> --out <output-folder>` to apply edits, build, verify, and refresh `package_result.json` in one command.
10. Team readiness:
   - Run `scripts/check_readiness.py --form-cells <form_cells.json> --out <readiness_report.json>` on each teammate machine before first real use.
   - `scripts/onboard_reimbursement_form.py` also writes `readiness_report.json` beside `form_cells.json` after inspecting and validating a new real template.
   - Treat `status: ready` as ready for full workflow.
   - Treat `status: needs_form_template` as a clean install without the real `费用报销单` mapping yet: detail tables, folders, and print packs can still run, but final objective is not complete.
11. Invoice coverage gate:
   - Default to `invoice_requirement.policy: full_amount`, so recognized invoices must cover the reimbursement total.
   - After expense review, write `invoice_coverage.json` and `发票缺口清单.md` with required, recognized, missing, excess, and replacement-invoice amounts.
   - Return `needs_invoices` when the recognized amount is short and do not create the final finance zip or one-click print PDF.
   - Return `needs_invoice_review` when an invoice total cannot be read or duplicate file content/invoice numbers are found.
   - Classify ride itineraries and similar files as `supporting_documents`: preserve and print them, but never add their displayed trip totals to invoice coverage.
   - Accept replacement invoices only when they are real, verifiable, unused in this pack, and permitted by the user's finance policy. Preserve `role: replacement`; do not attach them to an expense as if they were its original invoice.
   - Build only when coverage is `covered` and the DingTalk screenshot fields needed by the reimbursement form have been received. Include every original and replacement invoice in `03_发票`, the invoice print PDF, and the combined print PDF. Count the reimbursement form itself in the combined print page total.
12. Pre-DingTalk amount confirmation:
   - Assign stable voucher IDs `FY001`, `FY002`, and so on after chronological sorting.
   - Write `报销金额确认单.md` before final packaging, showing voucher ID, date, purpose, category, amount, payment-evidence status, invoice status, category subtotals, and exact total.
   - The DingTalk approval metadata must include `dingding_number`, title-derived `reimbursement_type`, `payee`, `invoice_entity`, and `approved_amount`. Require `approved_amount` to equal the payment-screenshot total exactly.
   - Return `needs_approval`, `needs_approval_and_invoices`, or `needs_approval_and_invoice_review` when the relevant gates are not complete. Do not build final materials in those states.
13. Expense exceptions:
   - Write `报销异常清单.md` and flag refunds/reversals, personal or split expenses, foreign currency, multiple displayed amounts, and discounts/subsidies.
   - Blocking exceptions require an `exception_resolution` in the Excel review workbook before final packaging. Warnings remain visible for review.
14. Reusable profile:
   - A profile may provide only stable defaults such as `reimburser`, `payee`, `invoice_entity`, and `company`.
   - Never reuse project-specific `dingding_number` or `approved_amount` from a profile.
   - Discover `reimbursement_profile.json` from `--profile`, `EXPENSE_REIMBURSEMENT_PROFILE`, `./03_profile/`, or `${CODEX_HOME:-~/.codex}/expense-reimbursement/`.
15. Final audit summary:
   - Write `报销审计摘要.md` with expense count, voucher range, reimbursement total, DingTalk amount, category totals, invoice coverage, replacement amount, exception count, print pages, and verification status.
   - State explicitly that file/OCR checks do not replace tax-platform invoice-authenticity verification.

## Classification Rules

Use finance-facing purposes, not merchant names.

- Eating, restaurant, takeout, coffee/tea, meals: purpose `餐费`, type `餐饮费`.
- Didi, taxi, ride-hailing, metro, bus, train, parking, tolls: type `交通费`.
  - Didi/taxi purpose: `打车费`.
  - Metro/bus purpose: `地铁费` or `公交费`.
  - Parking purpose: `停车费`.
- Huolala / delivery of equipment: purpose `设备配送费`, type `设备费`.
- Camera, phone, lens, rental, accessories, adapters, and equipment-specific props: use `设备租赁费` or `设备配件费`, type `设备费`.
- Studio, venue, and location rental: purpose `场地租赁费`, type `场地费`.
- Actors, models, performers, and on-camera talent: purpose `演员费`, type `演员费`.
- SF Express/courier, software/app/subscription purchases, and uncategorized project costs: use a specific purpose such as `快递费` or `软件服务费`, type `其他费用`.

If a screenshot is ambiguous, inspect the image and OCR text. If still ambiguous, ask the user for the exact purpose or category before generating final files.

## Sorting And Naming

Sort rows by `date`, then `time` when available, then original filename. Preserve the original screenshot/invoice files by copying them into the finance folder with a numeric prefix:

`FY001_YYYY-MM-DD_用途_金额_原文件名.ext`

Keep the generated workbook and PDFs in both their main output folders and the finance submission folder.

## Expense Form Template Handling

For Excel templates (`.xlsx` or `.xlsm`), fill only the cells mapped in `reimbursement_form.cells` and preserve the rest of the workbook. For other template types, copy the template into the output pack unchanged and tell the user that manual or custom filling is still needed.

Use `scripts/inspect_reimbursement_form.py --template <费用报销单.xlsx> --out <inspection.json> --cells-out <form_cells.json> --relative-template` to list sheet names, merged ranges, non-empty cells, candidate fill cells, and a reusable mapping file. Treat its recommendations as a draft; visually confirm the target cells before reusing `form_cells.json` with `package_from_folder.py --form-cells`. The reusable JSON stores the template path; `--relative-template` makes that path portable when the JSON and template move together.

For a real team template, prefer `scripts/onboard_reimbursement_form.py --template <费用报销单.xlsx> --out <form-template-config> --copy-template`. It writes `form_inspection.json`, portable `form_cells.json`, and `onboarding_report.json`. If sample screenshots are available, add `--sample-input <folder>` and optional `--sample-invoice-input <folder-or-file>` so the onboarding run also builds and verifies a sample pack.

After editing `form_cells.json`, run `scripts/validate_form_config.py --form-cells <form_cells.json> --out <form_config_validation.json>`. Require no errors before team reuse. Warnings about merged cells or missing optional fields should be reviewed against the visible template.

Recommended default mappings after inspecting a common form:

- Applicant / 报销人: `{{reimburser}}`
- Project / 事由: `{{project_name}}`
- Period / 日期范围: `{{date_range}}`
- Amount / 报销金额: `{{total_amount}}`
- Count / 单据张数: `{{expense_count}}`
- Repeated category slots: use `reimbursement_form.category_rows` with six `label_cell` and `amount_cell` slots. Fill nonzero categories only. The approved team template has six rows so all default categories fit without merging.

## OCR Guidance

For WeChat/Alipay screenshots, extract:

- Amount near the top, usually negative, such as `-57.90`; store as positive `57.90`.
- Payment time / 支付时间.
- Merchant or product text only for audit notes; do not put merchant names in the final `用途` column unless the user asks.
- Invoice availability when visible, but do not mark `发票` as present unless an actual invoice file is included.

Use OCR as a draft. Visually verify every amount before finalizing.

When OCR is unavailable, the folder draft script can also use filename fallback for deterministic drafts. Filename text such as `2026-01-01` and `12.30` can fill date and amount, while Chinese purpose words such as `滴滴`, `餐费`, or `货拉拉` can drive category classification. Still verify the resulting manifest before final submission.

## Folder-To-Pack Flow

When the user provides a folder rather than a prepared manifest:

1. Prefer `scripts/package_from_folder.py` for a one-command run. It automatically writes `<output-folder>/preflight_report.json` before OCR/building.
2. For especially messy inputs, run `scripts/preflight_inputs.py --input <folder-or-zip> --invoice-input <invoice-folder-or-zip> --out <preflight_report.json>` first. Direct `.zip` archives are scanned in place and reported as `archive.zip::member/path.png`. Confirm the report's screenshot/invoice counts and warnings.
3. `--input` may be a folder or `.zip`; zip inputs are extracted into the output folder before OCR and packing.
4. If original invoices are in a separate file, folder, or `.zip`, pass them with one or more `--invoice-input <path>` arguments. These become top-level `invoices` unless manually matched to specific entries later.
5. If a `费用报销单` Excel template is available, pass it with `--form-template <template.xlsx>`. The script writes `form_inspection.json`, adds recommended mappings, and builds the filled form when usable.
6. For a fixed team form template, pass the reviewed `form_cells.json` through `--form-cells` so the one-command run uses the approved template, sheet, and cell mapping instead of only the recommendation.
7. When `--form-cells` is omitted, `package_from_folder.py` auto-detects `form_cells.json` from `--form-config-dir`, `EXPENSE_REIMBURSEMENT_FORM_CELLS`, `EXPENSE_REIMBURSEMENT_FORM_CONFIG_DIR`, the current folder, `./02_form_template_config/`, or `${CODEX_HOME:-~/.codex}/expense-reimbursement/`. Check `package_result.json.form_cells` to confirm which config was used.
8. If it stops with `needs_review`, open `draft_manifest_review.xlsx`, correct blank fields and rows marked `NEEDS_REVIEW`, then finish the pack:

```bash
python scripts/package_reviewed_workbook.py --manifest <output-folder>/draft_manifest.json --workbook <output-folder>/draft_manifest_review.xlsx --out <output-folder>
```

Also inspect any `form_inspection.json` when a `费用报销单` template is involved.
9. Open `报销金额确认单.md`, confirm the exact total, submit DingTalk, then transcribe the visible screenshot fields into an approval JSON and pass it with `--approval-metadata <approval.json>`.
10. After expense review, inspect `invoice_coverage.json` and `发票缺口清单.md`. If status is `needs_invoices`, provide approved replacement invoices with `--replacement-invoice-input <path>` and rerun into a fresh final output folder.
11. Keep replacement invoices at top level with `role: replacement`. Original invoices may be attached to entries when there is an obvious one-to-one match; otherwise leave them top level.
12. Use `scripts/draft_manifest_from_folder.py` directly only when you need a draft-only pass. Run `scripts/prepare_expense_confirmation.py`, then `scripts/analyze_invoice_coverage.py`, before a manual build.

The folder draft script can use Apple Vision OCR on macOS through `scripts/vision_ocr.swift`. OCR is best-effort. Do not trust it blindly for money, dates, or categories.

`package_from_folder.py` runs preflight automatically before building unless `--skip-preflight` is passed. Add `--stop-on-preflight-warnings` when warnings should stop the run. It runs verification automatically after building unless `--skip-verify` is passed.

After a successful one-command build, check `package_result.json` for machine-readable status and open `交付清单.md` for the human handoff surface.

HEIC/HEIF screenshots or invoice images are valid inputs. `build_reimbursement_pack.py` converts them to PNG in `_converted_images` before embedding them in Excel/PDF outputs.

Run `scripts/check_readiness.py --form-cells <form_cells.json> --out <readiness_report.json>` before first real use to check scripts, Python dependencies, optional OCR availability, and the reusable form config. When onboarding a real template, `scripts/onboard_reimbursement_form.py` writes this readiness report automatically.

Run `scripts/self_test.py --out <folder>` on a new machine to verify local dependencies, portable form config resolution, readiness reports, Excel manifest review export/apply, missing-invoice interruption, replacement-invoice packaging and printing, reimbursement workbook generation, form filling, finance zip cleanup, strict amount tamper detection, and pack verification.

## Invoice Handling

Use entry-level `invoices` when an original invoice clearly belongs to a specific expense. Use top-level `invoices` for standalone invoice files. Record each file in `invoice_items` with `role: original` or `role: replacement`.

The analyzer extracts invoice totals, invoice numbers, invoice dates, and buyer names when available. It hashes files, records linked voucher IDs, checks the expected invoice entity, flags unreasonable date windows, and rejects unresolved or duplicate invoices from coverage. Recognized amounts are summed independently from expense totals. With `full_amount` policy, `recognized_amount` must be at least `required_amount`; otherwise the final build is blocked.

Approved replacement invoices stay explicitly labeled as replacements in the manifest, finance folder, gap report, and verification. The pack builder copies all recognized invoice files to `财务提交文件夹/03_发票` and combines them into `打印/<项目>_发票_一张一页.pdf` and the final one-click print PDF.

Ride itineraries, travel statements, and similar evidence belong in top-level `supporting_documents`. Copy them to `03_发票` with a `发票附件_行程单` label and include them after invoices in the print PDF. Their amounts are contextual only and must not contribute to `recognized_amount`.
