#!/usr/bin/env python3
import argparse
import json
import os
import subprocess
import sys
import zipfile
from decimal import Decimal, InvalidOperation
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser(description="Draft a manifest from a folder and build a reimbursement pack in one command.")
    parser.add_argument("--input", required=True, help="Folder or .zip containing screenshots and invoices.")
    parser.add_argument("--invoice-input", action="append", default=[], help="Additional invoice file, folder, or .zip. Can be passed multiple times.")
    parser.add_argument("--replacement-invoice-input", action="append", default=[], help="Finance-approved replacement invoice file, folder, or .zip. Can be passed multiple times.")
    parser.add_argument("--supporting-document-input", action="append", default=[], help="Invoice supporting document file, folder, or .zip. Not counted toward invoice coverage.")
    parser.add_argument("--invoice-policy", choices=["full_amount", "none"], default="full_amount", help="full_amount requires recognized invoices to cover the reimbursement total; none disables the coverage gate.")
    parser.add_argument("--project-name", required=True, help="Project/shoot name.")
    parser.add_argument("--reimburser", help="Person being reimbursed. May be supplied by --profile.")
    parser.add_argument("--profile", help="Optional reimbursement profile JSON path or JSON object for stable reimburser/payee/invoice-entity defaults.")
    parser.add_argument("--approval-metadata", help="DingTalk fields transcribed from the approval screenshot, as a JSON path or JSON object.")
    parser.add_argument("--approval-policy", choices=["required", "none"], default="required", help="required blocks final packaging until DingTalk fields and approved amount are complete; none explicitly waives the gate.")
    parser.add_argument("--out", required=True, help="Output folder for the generated pack.")
    parser.add_argument("--manifest-out", help="Optional path for the draft manifest. Defaults inside --out.")
    parser.add_argument("--preflight-out", help="Optional path for the input preflight report. Defaults inside --out.")
    parser.add_argument("--skip-preflight", action="store_true", help="Skip writing the automatic input preflight report.")
    parser.add_argument("--stop-on-preflight-warnings", action="store_true", help="Stop before OCR/building when preflight_report.json has warnings.")
    parser.add_argument("--recursive", action="store_true", help="Scan folders recursively.")
    parser.add_argument("--ocr", choices=["auto", "apple-vision", "tesseract", "none"], default="auto", help="OCR mode for screenshots and image invoices.")
    parser.add_argument("--categories", default="交通费,餐饮费,设备费,场地费,演员费,其他费用", help="Comma-separated category list.")
    parser.add_argument("--review-policy", choices=["auto", "always", "never"], default="auto", help="auto stops when rows need review; always drafts only; never stops for review.")
    parser.add_argument("--form-template", help="Optional 费用报销单 Excel template to fill. May be omitted when --form-cells JSON contains template.")
    parser.add_argument("--form-sheet", help="Optional sheet name in the 费用报销单 template.")
    parser.add_argument("--form-output-name", help="Optional generated filename for the filled 费用报销单.")
    parser.add_argument("--form-cells", help="Optional JSON object or JSON file path for reusable reimbursement_form template/sheet/cells.")
    parser.add_argument("--form-config-dir", help="Optional folder containing form_cells.json. Used when --form-cells is omitted.")
    parser.add_argument("--form-inspection-out", help="Optional JSON output path for form inspection. Defaults inside --out.")
    parser.add_argument("--skip-verify", action="store_true", help="Skip automatic pack verification after build.")
    return parser.parse_args()


def run(cmd):
    result = subprocess.run(
        cmd,
        check=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    if result.stderr:
        print(result.stderr, file=sys.stderr, end="")
    if result.stdout:
        print(result.stdout, end="")
    return result


def run_no_check(cmd):
    result = subprocess.run(
        cmd,
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    if result.stderr:
        print(result.stderr, file=sys.stderr, end="")
    if result.stdout:
        print(result.stdout, end="")
    return result


def read_json_if_exists(path):
    path = Path(path)
    if not path.exists():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))


def money(value):
    return Decimal(str(value)).quantize(Decimal("0.01"))


def load_json_arg(value):
    if not value:
        return {}
    text = str(value).strip()
    if text.startswith("{"):
        data = json.loads(text)
    else:
        path = Path(text).expanduser().resolve()
        if not path.exists():
            raise SystemExit(f"JSON input not found: {path}")
        data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise SystemExit("JSON input must contain an object.")
    return data


def resolve_profile_source(args):
    if args.profile:
        if str(args.profile).strip().startswith("{"):
            return None, "argument-json"
        return Path(args.profile).expanduser().resolve(), "argument"
    env_profile = os.environ.get("EXPENSE_REIMBURSEMENT_PROFILE")
    if env_profile:
        path = Path(env_profile).expanduser().resolve()
        if not path.exists():
            raise SystemExit(f"EXPENSE_REIMBURSEMENT_PROFILE does not exist: {path}")
        return path, "EXPENSE_REIMBURSEMENT_PROFILE"
    candidates = [
        Path.cwd().resolve() / "03_profile" / "reimbursement_profile.json",
        *[folder / "expense-reimbursement" / "reimbursement_profile.json" for folder in default_config_dirs()],
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate.resolve(), "auto"
    return None, ""


def apply_profile_and_approval(manifest_path, args, profile, profile_source=None):
    manifest_path = Path(manifest_path)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    reimburser = str(args.reimburser or profile.get("reimburser") or "").strip()
    if not reimburser:
        raise SystemExit("--reimburser is required unless the reimbursement profile provides reimburser.")
    manifest["reimburser"] = reimburser
    manifest["approval_requirement"] = {
        "policy": args.approval_policy,
        "required_fields": ["dingding_number", "reimbursement_type", "payee", "invoice_entity", "approved_amount"],
    }

    approval = dict(manifest.get("approval_metadata") or {})
    stable_defaults = profile.get("approval_defaults") or profile.get("stable_defaults") or {}
    for key in ["payee", "invoice_entity", "company"]:
        if stable_defaults.get(key) not in (None, ""):
            approval[key] = stable_defaults[key]
    explicit = load_json_arg(args.approval_metadata)
    approval.update(explicit)
    if approval.get("reimburser") and str(approval["reimburser"]).strip() != reimburser:
        raise SystemExit(
            f"DingTalk reimburser {approval['reimburser']!r} does not match reimbursement profile/argument {reimburser!r}."
        )
    manifest["approval_metadata"] = approval
    manifest["profile_source"] = str(profile_source) if profile_source else ("argument-json" if args.profile else "")

    total = sum((money(entry.get("amount", 0)) for entry in manifest.get("entries", []) or []), Decimal("0.00"))
    try:
        approved_amount = money(approval.get("approved_amount"))
    except (InvalidOperation, ValueError, TypeError):
        approved_amount = None
    if approval.get("dingding_number") and approved_amount == total:
        manifest["expense_confirmation"] = {
            "status": "confirmed_by_dingtalk",
            "confirmed_amount": f"{total:.2f}",
            "dingding_number": str(approval.get("dingding_number")),
        }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return manifest


def approval_issues(manifest):
    requirement = manifest.get("approval_requirement") or {}
    if requirement.get("policy", "none") != "required":
        return []
    approval = manifest.get("approval_metadata") or {}
    issues = []
    for field in requirement.get("required_fields", []) or []:
        if approval.get(field) in (None, ""):
            issues.append({"field": field, "type": "missing", "message": f"DingTalk field is missing: {field}"})
    total = sum((money(entry.get("amount", 0)) for entry in manifest.get("entries", []) or []), Decimal("0.00"))
    if approval.get("approved_amount") not in (None, ""):
        try:
            approved_amount = money(approval["approved_amount"])
            if approved_amount != total:
                issues.append({
                    "field": "approved_amount",
                    "type": "mismatch",
                    "message": f"DingTalk approved amount {approved_amount:.2f} does not match expense total {total:.2f}",
                })
        except (InvalidOperation, ValueError, TypeError):
            issues.append({"field": "approved_amount", "type": "invalid", "message": "DingTalk approved amount is invalid."})
    confirmation = manifest.get("expense_confirmation") or {}
    if not issues and confirmation.get("status") != "confirmed_by_dingtalk":
        issues.append({"field": "expense_confirmation", "type": "unconfirmed", "message": "Expense total has not been confirmed by matching DingTalk metadata."})
    return issues


def write_approval_gap_report(manifest_path, out_dir):
    manifest_path = Path(manifest_path)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    issues = approval_issues(manifest)
    labels = {
        "dingding_number": "钉钉审批编号",
        "reimbursement_type": "报销类型",
        "payee": "领款人",
        "invoice_entity": "开票对象",
        "approved_amount": "钉钉审批金额",
        "expense_confirmation": "金额确认状态",
    }
    rows = "\n".join(
        f"- {labels.get(item['field'], item['field'])}: {item['message']}"
        for item in issues
    ) or "- 无"
    output = Path(out_dir) / "钉钉信息缺口.md"
    output.write_text(
        f"# 钉钉信息缺口\n\n项目：{manifest.get('project_name', '')}\n"
        f"当前状态：{'需要补充或更正' if issues else '完整'}\n\n## 待处理\n\n{rows}\n",
        encoding="utf-8",
    )
    reports = dict(manifest.get("workflow_reports") or {})
    reports["approval_gap"] = str(output)
    manifest["workflow_reports"] = reports
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output, issues


def run_workflow_reports(script_dir, manifest_path, out_dir):
    result = run([
        sys.executable,
        str(script_dir / "prepare_expense_confirmation.py"),
        "--manifest", str(manifest_path),
        "--out-dir", str(out_dir),
        "--update-manifest",
    ])
    return json.loads(result.stdout)


def prerequisite_status(coverage_status, current_approval_issues):
    if current_approval_issues and coverage_status == "needs_invoices":
        return "needs_approval_and_invoices"
    if current_approval_issues and coverage_status == "needs_invoice_review":
        return "needs_approval_and_invoice_review"
    if current_approval_issues:
        return "needs_approval"
    if coverage_status in {"needs_invoices", "needs_invoice_review"}:
        return coverage_status
    return ""


def refresh_finance_zip(finance_folder, finance_zip):
    finance_folder = Path(finance_folder)
    finance_zip = Path(finance_zip)
    with zipfile.ZipFile(finance_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for folder in sorted(path for path in finance_folder.rglob("*") if path.is_dir()):
            arcname = Path(finance_folder.name) / folder.relative_to(finance_folder)
            zf.writestr(str(arcname).rstrip("/") + "/", "")
        for path in sorted(path for path in finance_folder.rglob("*") if path.is_file()):
            if path.name == ".DS_Store" or path.name.startswith("._") or "__MACOSX" in path.parts or path.suffix == ".pyc":
                continue
            arcname = Path(finance_folder.name) / path.relative_to(finance_folder)
            zf.write(path, arcname)


def finalize_audit_summary(out_dir, verification_path, passed):
    out_dir = Path(out_dir)
    summary_path = out_dir / "pack_summary.json"
    summary = read_json_if_exists(summary_path)
    audit_path = Path(summary.get("audit_summary", out_dir / "报销审计摘要.md"))
    verification = read_json_if_exists(verification_path)
    status_text = f"全部通过（{len(verification.get('checks', []))} 项）" if passed else "未通过，请查看 verification.json"
    if audit_path.exists():
        text = audit_path.read_text(encoding="utf-8")
        text = text.replace("自动校验：待执行", f"自动校验：{status_text}")
        audit_path.write_text(text, encoding="utf-8")
        finance_audit = Path(summary.get("finance_folder", out_dir / "财务提交文件夹")) / audit_path.name
        if finance_audit.parent.exists():
            finance_audit.write_text(text, encoding="utf-8")
    summary["verification_status"] = "passed" if passed else "failed"
    summary["verification_check_count"] = len(verification.get("checks", []))
    if summary:
        summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    if passed and summary.get("finance_folder") and summary.get("finance_zip"):
        refresh_finance_zip(summary["finance_folder"], summary["finance_zip"])


def find_generated(summary, needle):
    for value in summary.get("generated_files", []) or []:
        path = Path(value)
        if needle in path.name:
            return str(path)
    return ""


def write_package_result(out_dir, status, manifest_path, preflight_path=None, form_inspection=None, review_rows=None, verification_path=None, review_workbook=None, form_cells_source=None, form_cells_origin=""):
    summary_path = out_dir / "pack_summary.json"
    summary = read_json_if_exists(summary_path)
    preflight = read_json_if_exists(preflight_path) if preflight_path else {}
    verification = read_json_if_exists(verification_path) if verification_path else read_json_if_exists(out_dir / "verification.json")
    invoice_coverage_path = out_dir / "invoice_coverage.json"
    invoice_gap_path = out_dir / "发票缺口清单.md"
    invoice_coverage = read_json_if_exists(invoice_coverage_path)
    manifest = read_json_if_exists(manifest_path)
    workflow_reports = manifest.get("workflow_reports") or {}
    current_approval_issues = approval_issues(manifest) if manifest else []
    expense_form = find_generated(summary, "费用报销单")
    result = {
        "status": status,
        "output_folder": str(out_dir),
        "package_result": str(out_dir / "package_result.json"),
        "manifest": str(manifest_path),
        "manifest_review_workbook": str(review_workbook) if review_workbook else "",
        "form_cells": str(form_cells_source) if form_cells_source else "",
        "form_cells_origin": form_cells_origin,
        "preflight_report": str(preflight_path) if preflight_path else "",
        "preflight_status": preflight.get("status", ""),
        "preflight_warnings": preflight.get("warnings", []),
        "form_inspection": str(form_inspection) if form_inspection else "",
        "review_rows": review_rows or [],
        "profile_source": manifest.get("profile_source", ""),
        "confirmation_report": workflow_reports.get("expense_confirmation", ""),
        "exception_report": workflow_reports.get("expense_exceptions", ""),
        "approval_gap_report": workflow_reports.get("approval_gap", ""),
        "approval_issues": current_approval_issues,
        "voucher_count": len(manifest.get("entries", []) or []),
        "verification": str(verification_path or (out_dir / "verification.json")),
        "verification_status": verification.get("status", "skipped" if status == "built_unverified" else ""),
        "delivery_checklist": summary.get("delivery_checklist", ""),
        "finance_folder": summary.get("finance_folder", ""),
        "finance_zip": summary.get("finance_zip", ""),
        "reimbursement_workbook": find_generated(summary, "报销明细表"),
        "expense_form": expense_form,
        "expense_form_print_pdf": find_generated(summary, "费用报销单_打印"),
        "expense_form_status": "generated" if expense_form else "not_generated",
        "one_click_print_pdf": find_generated(summary, "全部打印_一键打印"),
        "screenshot_print_pdf": find_generated(summary, "费用截图_三张一页"),
        "invoice_print_pdf": find_generated(summary, "发票_一张一页"),
        "supporting_document_count": summary.get(
            "supporting_document_count",
            (preflight.get("counts") or {}).get("supporting_documents", 0),
        ),
        "invoice_coverage_report": str(invoice_coverage_path) if invoice_coverage_path.exists() else "",
        "invoice_gap_report": str(invoice_gap_path) if invoice_gap_path.exists() else "",
        "invoice_coverage_status": invoice_coverage.get("status", ""),
        "invoice_required_amount": invoice_coverage.get("required_amount", ""),
        "invoice_recognized_amount": invoice_coverage.get("recognized_amount", ""),
        "invoice_missing_amount": invoice_coverage.get("missing_amount", ""),
        "replacement_invoice_count": invoice_coverage.get("replacement_invoice_count", 0),
        "replacement_invoice_amount": invoice_coverage.get("replacement_invoice_amount", "0.00"),
        "invoice_validation_warnings": invoice_coverage.get("validation_warnings", []),
        "print_layout_workbook": find_generated(summary, "打印排版"),
        "audit_summary": summary.get("audit_summary", ""),
        "next_step": "",
    }
    if status == "needs_review":
        if review_workbook:
            result["next_step"] = "Edit draft_manifest_review.xlsx, then run package_reviewed_workbook.py --manifest draft_manifest.json --workbook draft_manifest_review.xlsx --out <output-folder> to apply, build, verify, and refresh package_result.json."
        else:
            result["next_step"] = "Edit the manifest, then run build_reimbursement_pack.py --manifest ... --out ..."
    elif status == "passed":
        result["next_step"] = "Open 交付清单.md, submit the finance zip, and print the one-click PDF."
    elif status == "built_unverified":
        result["next_step"] = "Run verify_reimbursement_pack.py before submitting to finance."
    elif status == "failed_verification":
        result["next_step"] = "Open verification.json, fix failed checks, then rebuild or verify again."
    elif status == "preflight_needs_review":
        result["next_step"] = "Review inputs or rerun without --stop-on-preflight-warnings when the warnings are acceptable."
    elif status == "needs_invoices":
        result["next_step"] = "Provide finance-approved replacement invoices for the reported missing amount, then rerun with --replacement-invoice-input."
    elif status == "needs_invoice_review":
        result["next_step"] = "Review invoice_coverage.json, confirm unresolved invoice totals, then rerun invoice coverage before building the final pack."
    elif status == "needs_approval":
        result["next_step"] = "Provide DingTalk approval metadata, including the approval number, title-derived reimbursement type, payee, invoice entity, and approved amount, then rerun."
    elif status == "needs_approval_and_invoices":
        result["next_step"] = "Provide the DingTalk approval metadata and finance-approved invoices for the reported missing amount, then rerun."
    elif status == "needs_approval_and_invoice_review":
        result["next_step"] = "Resolve the DingTalk metadata gaps and invoice review items, then rerun."
    output = out_dir / "package_result.json"
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output, result


def run_preflight(args, script_dir, out_dir):
    if args.skip_preflight:
        return None, None
    preflight_path = Path(args.preflight_out).expanduser().resolve() if args.preflight_out else out_dir / "preflight_report.json"
    preflight_cmd = [
        sys.executable,
        str(script_dir / "preflight_inputs.py"),
        "--input", args.input,
        "--out", str(preflight_path),
    ]
    for invoice_input in args.invoice_input:
        preflight_cmd.extend(["--invoice-input", invoice_input])
    for invoice_input in args.replacement_invoice_input:
        preflight_cmd.extend(["--replacement-invoice-input", invoice_input])
    for supporting_input in args.supporting_document_input:
        preflight_cmd.extend(["--supporting-document-input", supporting_input])
    if args.recursive:
        preflight_cmd.append("--recursive")
    run(preflight_cmd)
    report = json.loads(preflight_path.read_text(encoding="utf-8"))
    return preflight_path, report


def should_skip_archive_member(name):
    parts = Path(name).parts
    return (
        not name
        or name.endswith("/")
        or "__MACOSX" in parts
        or "__pycache__" in parts
        or Path(name).name == ".DS_Store"
        or Path(name).name.startswith("._")
        or Path(name).suffix == ".pyc"
    )


def safe_extract_zip(zip_path, destination):
    destination.mkdir(parents=True, exist_ok=True)
    root = destination.resolve()
    with zipfile.ZipFile(zip_path) as zf:
        for info in zf.infolist():
            if should_skip_archive_member(info.filename):
                continue
            target = (root / info.filename).resolve()
            if root not in target.parents and target != root:
                raise SystemExit(f"Unsafe zip member path: {info.filename}")
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as source, target.open("wb") as dest:
                dest.write(source.read())
    return destination


def prepare_input_path(value, out_dir, label, allow_file=False):
    path = Path(value).expanduser().resolve()
    if path.is_dir():
        return path, False
    if path.is_file() and path.suffix.lower() != ".zip" and allow_file:
        return path, False
    if path.is_file() and path.suffix.lower() == ".zip":
        return safe_extract_zip(path, out_dir / "_extracted_inputs" / label), True
    if path.is_file():
        raise SystemExit(f"Expected folder or .zip for {label}: {path}")
    raise SystemExit(f"Input not found: {path}")


def needs_review(manifest_path):
    data = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    rows = []
    for idx, entry in enumerate(data.get("entries", []), start=1):
        required_missing = [key for key in ["date", "purpose", "expense_type", "amount", "screenshot"] if entry.get(key) in ("", None)]
        if entry.get("notes") or required_missing:
            rows.append({
                "row": idx,
                "voucher_id": entry.get("voucher_id", ""),
                "file": Path(entry.get("screenshot", "")).name,
                "missing": required_missing,
                "notes": entry.get("notes", ""),
            })
    for item in data.get("expense_exceptions", []) or []:
        if item.get("severity") != "blocking" or item.get("resolved"):
            continue
        voucher_id = item.get("voucher_id", "")
        matching_index = next(
            (
                index
                for index, entry in enumerate(data.get("entries", []) or [], start=1)
                if entry.get("voucher_id") == voucher_id
            ),
            "exception",
        )
        rows.append({
            "row": matching_index,
            "voucher_id": voucher_id,
            "file": "",
            "missing": ["exception_resolution"],
            "notes": item.get("message", ""),
        })
    return rows


def resolve_config_path(value, base_dir=None):
    if not value:
        return ""
    path = Path(value).expanduser()
    if not path.is_absolute() and base_dir:
        path = base_dir / path
    return str(path.resolve())


def default_config_dirs():
    candidates = []
    explicit = os.environ.get("EXPENSE_REIMBURSEMENT_HOME")
    if explicit:
        candidates.append(Path(explicit).expanduser())
    if os.name == "nt":
        appdata = os.environ.get("APPDATA")
        if appdata:
            candidates.append(Path(appdata) / "expense-reimbursement-packager")
    elif sys.platform == "darwin":
        candidates.append(Path.home() / "Library" / "Application Support" / "expense-reimbursement-packager")
    else:
        xdg = os.environ.get("XDG_CONFIG_HOME")
        candidates.append(Path(xdg).expanduser() / "expense-reimbursement-packager" if xdg else Path.home() / ".config" / "expense-reimbursement-packager")

    codex_home = os.environ.get("CODEX_HOME")
    candidates.append(Path(codex_home).expanduser() if codex_home else Path.home() / ".codex")
    unique = []
    seen = set()
    for candidate in candidates:
        resolved = candidate.resolve()
        key = str(resolved)
        if key not in seen:
            seen.add(key)
            unique.append(resolved)
    return unique


def form_cells_from_dir(folder):
    path = Path(folder).expanduser().resolve() / "form_cells.json"
    return path if path.exists() else None


def resolve_form_cells_source(args):
    if args.form_cells:
        if args.form_cells.strip().startswith("{"):
            return None, "argument-json"
        return Path(args.form_cells).expanduser().resolve(), "argument"
    if args.form_config_dir:
        path = form_cells_from_dir(args.form_config_dir)
        if not path:
            raise SystemExit(f"form_cells.json not found in --form-config-dir: {Path(args.form_config_dir).expanduser().resolve()}")
        return path, "--form-config-dir"
    env_cells = os.environ.get("EXPENSE_REIMBURSEMENT_FORM_CELLS")
    if env_cells:
        path = Path(env_cells).expanduser().resolve()
        if not path.exists():
            raise SystemExit(f"EXPENSE_REIMBURSEMENT_FORM_CELLS does not exist: {path}")
        return path, "EXPENSE_REIMBURSEMENT_FORM_CELLS"
    env_dir = os.environ.get("EXPENSE_REIMBURSEMENT_FORM_CONFIG_DIR")
    if env_dir:
        path = form_cells_from_dir(env_dir)
        if not path:
            raise SystemExit(f"form_cells.json not found in EXPENSE_REIMBURSEMENT_FORM_CONFIG_DIR: {Path(env_dir).expanduser().resolve()}")
        return path, "EXPENSE_REIMBURSEMENT_FORM_CONFIG_DIR"

    cwd = Path.cwd().resolve()
    candidates = [
        cwd / "form_cells.json",
        cwd / "02_form_template_config" / "form_cells.json",
        *[folder / "expense-reimbursement" / "form_cells.json" for folder in default_config_dirs()],
    ]
    seen = set()
    for candidate in candidates:
        key = str(candidate)
        if key in seen:
            continue
        seen.add(key)
        if candidate.exists():
            return candidate.resolve(), "auto"
    built_in = Path(__file__).resolve().parent.parent / "assets" / "form_cells.json"
    if built_in.exists():
        return built_in.resolve(), "built-in"
    return None, ""


def load_form_config_arg(value):
    if not value:
        return {}
    text = value.strip()
    base_dir = None
    if text.startswith("{"):
        data = json.loads(text)
    else:
        path = Path(value).expanduser().resolve()
        base_dir = path.parent
        data = json.loads(path.read_text(encoding="utf-8"))
    if "recommended_manifest_reimbursement_form" in data:
        form = data.get("recommended_manifest_reimbursement_form", {}) or {}
        return {
            "template": resolve_config_path(form.get("template", ""), base_dir),
            "sheet": form.get("sheet", ""),
            "output_name": form.get("output_name", ""),
            "cells": form.get("cells", {}),
            "category_rows": form.get("category_rows", []),
        }
    if "reimbursement_form" in data:
        form = data.get("reimbursement_form", {}) or {}
        return {
            "template": resolve_config_path(form.get("template", ""), base_dir),
            "sheet": form.get("sheet", ""),
            "output_name": form.get("output_name", ""),
            "cells": form.get("cells", {}),
            "category_rows": form.get("category_rows", []),
        }
    if "cells" in data:
        return {
            "template": resolve_config_path(data.get("template", ""), base_dir),
            "sheet": data.get("sheet", ""),
            "output_name": data.get("output_name", ""),
            "cells": data.get("cells", {}),
            "category_rows": data.get("category_rows", []),
        }
    return {"cells": data}


def apply_form_template(manifest_path, args, script_dir, out_dir):
    form_config = load_form_config_arg(args.form_cells) if args.form_cells else {}
    template_value = args.form_template or form_config.get("template")
    if not template_value:
        if args.form_cells:
            raise SystemExit("--form-cells requires --form-template or a template field in the JSON.")
        return None
    template = Path(template_value).expanduser().resolve()
    if not template.exists():
        raise SystemExit(f"Form template not found: {template}")
    inspection_path = Path(args.form_inspection_out).expanduser().resolve() if args.form_inspection_out else out_dir / "form_inspection.json"
    inspect_cmd = [
        sys.executable,
        str(script_dir / "inspect_reimbursement_form.py"),
        "--template", str(template),
        "--out", str(inspection_path),
    ]
    if args.form_sheet:
        inspect_cmd.extend(["--sheet", args.form_sheet])
    run(inspect_cmd)

    inspection = json.loads(inspection_path.read_text(encoding="utf-8"))
    form = dict(inspection.get("recommended_manifest_reimbursement_form") or {})
    form["template"] = str(template)
    if args.form_sheet:
        form["sheet"] = args.form_sheet
    if args.form_output_name:
        form["output_name"] = args.form_output_name
    if form_config:
        if form_config.get("sheet") and not args.form_sheet:
            form["sheet"] = form_config["sheet"]
        if form_config.get("output_name") and not args.form_output_name:
            form["output_name"] = form_config["output_name"]
        form["cells"] = form_config.get("cells", {})
        form["category_rows"] = form_config.get("category_rows", [])

    data = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    data["reimbursement_form"] = form
    Path(manifest_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return inspection_path


def form_review_rows(manifest_path):
    data = json.loads(Path(manifest_path).read_text(encoding="utf-8"))
    form = data.get("reimbursement_form") or {}
    if form.get("template") and not form.get("cells"):
        return [{
            "row": "reimbursement_form",
            "file": Path(form.get("template", "")).name,
            "missing": ["cells"],
            "notes": "No fillable cells were detected; inspect the template and add reimbursement_form.cells.",
        }]
    return []


def export_manifest_review_workbook(script_dir, manifest_path, out_dir):
    review_workbook = out_dir / "draft_manifest_review.xlsx"
    run([
        sys.executable,
        str(script_dir / "manifest_review_workbook.py"),
        "export",
        "--manifest",
        str(manifest_path),
        "--workbook",
        str(review_workbook),
    ])
    return review_workbook


def run_invoice_coverage(script_dir, manifest_path, out_dir, ocr="auto"):
    coverage_path = out_dir / "invoice_coverage.json"
    markdown_path = out_dir / "发票缺口清单.md"
    run([
        sys.executable,
        str(script_dir / "analyze_invoice_coverage.py"),
        "--manifest", str(manifest_path),
        "--out", str(coverage_path),
        "--markdown-out", str(markdown_path),
        "--update-manifest",
        "--ocr", ocr,
    ])
    return coverage_path, json.loads(coverage_path.read_text(encoding="utf-8"))


def main():
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    out_dir = Path(args.out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    profile_source, profile_origin = resolve_profile_source(args)
    if args.profile:
        profile = load_json_arg(args.profile)
    elif profile_source:
        profile = load_json_arg(str(profile_source))
    else:
        profile = {}
    args.reimburser = str(args.reimburser or profile.get("reimburser") or "").strip()
    if not args.reimburser:
        raise SystemExit("--reimburser is required unless a reimbursement profile provides it.")
    if profile_source or args.profile:
        print(json.dumps({
            "profile": str(profile_source) if profile_source else "argument-json",
            "profile_origin": profile_origin,
        }, ensure_ascii=False))
    form_cells_source, form_cells_origin = resolve_form_cells_source(args)
    if form_cells_source:
        args.form_cells = str(form_cells_source)
        print(json.dumps({
            "form_cells": str(form_cells_source),
            "form_cells_origin": form_cells_origin,
        }, ensure_ascii=False))
    manifest_path = Path(args.manifest_out).expanduser().resolve() if args.manifest_out else out_dir / "draft_manifest.json"
    preflight_path, preflight_report = run_preflight(args, script_dir, out_dir)
    if args.stop_on_preflight_warnings and preflight_report and preflight_report.get("status") != "ready":
        package_result_path, _ = write_package_result(out_dir, "preflight_needs_review", manifest_path, preflight_path, form_cells_source=form_cells_source, form_cells_origin=form_cells_origin)
        print(json.dumps({
            "status": "preflight_needs_review",
            "preflight_report": str(preflight_path),
            "package_result": str(package_result_path),
            "warnings": preflight_report.get("warnings", []),
            "next_step": "Review inputs or rerun without --stop-on-preflight-warnings when the warnings are acceptable.",
        }, ensure_ascii=False, indent=2))
        raise SystemExit(1)
    source_input, source_is_zip = prepare_input_path(args.input, out_dir, "source")
    prepared_invoice_inputs = [
        prepare_input_path(invoice_input, out_dir, f"invoice_{index}", allow_file=True)
        for index, invoice_input in enumerate(args.invoice_input, start=1)
    ]
    prepared_replacement_invoice_inputs = [
        prepare_input_path(invoice_input, out_dir, f"replacement_invoice_{index}", allow_file=True)
        for index, invoice_input in enumerate(args.replacement_invoice_input, start=1)
    ]
    prepared_supporting_document_inputs = [
        prepare_input_path(supporting_input, out_dir, f"supporting_document_{index}", allow_file=True)
        for index, supporting_input in enumerate(args.supporting_document_input, start=1)
    ]
    invoice_inputs = [path for path, _ in prepared_invoice_inputs]
    replacement_invoice_inputs = [path for path, _ in prepared_replacement_invoice_inputs]
    supporting_document_inputs = [path for path, _ in prepared_supporting_document_inputs]
    scan_recursive = (
        args.recursive
        or source_is_zip
        or any(is_zip for _, is_zip in prepared_invoice_inputs)
        or any(is_zip for _, is_zip in prepared_replacement_invoice_inputs)
        or any(is_zip for _, is_zip in prepared_supporting_document_inputs)
    )

    draft_cmd = [
        sys.executable,
        str(script_dir / "draft_manifest_from_folder.py"),
        "--input", str(source_input),
        "--project-name", args.project_name,
        "--reimburser", args.reimburser,
        "--out", str(manifest_path),
        "--ocr", args.ocr,
        "--categories", args.categories,
        "--invoice-policy", args.invoice_policy,
        "--approval-policy", args.approval_policy,
    ]
    for invoice_input in invoice_inputs:
        draft_cmd.extend(["--invoice-input", str(invoice_input)])
    for invoice_input in replacement_invoice_inputs:
        draft_cmd.extend(["--replacement-invoice-input", str(invoice_input)])
    for supporting_input in supporting_document_inputs:
        draft_cmd.extend(["--supporting-document-input", str(supporting_input)])
    if scan_recursive:
        draft_cmd.append("--recursive")
    run(draft_cmd)
    apply_profile_and_approval(manifest_path, args, profile, profile_source)
    inspection_path = apply_form_template(manifest_path, args, script_dir, out_dir)
    run_workflow_reports(script_dir, manifest_path, out_dir)
    write_approval_gap_report(manifest_path, out_dir)

    review_rows = needs_review(manifest_path) + form_review_rows(manifest_path)
    if args.review_policy == "always" or (args.review_policy == "auto" and review_rows):
        review_workbook = export_manifest_review_workbook(script_dir, manifest_path, out_dir)
        package_result_path, _ = write_package_result(out_dir, "needs_review", manifest_path, preflight_path, inspection_path, review_rows, review_workbook=review_workbook, form_cells_source=form_cells_source, form_cells_origin=form_cells_origin)
        print(json.dumps({
            "status": "needs_review",
            "manifest": str(manifest_path),
            "manifest_review_workbook": str(review_workbook),
            "preflight_report": str(preflight_path) if preflight_path else "",
            "form_inspection": str(inspection_path) if inspection_path else "",
            "package_result": str(package_result_path),
            "review_rows": review_rows,
            "next_step": "Edit draft_manifest_review.xlsx, then run package_reviewed_workbook.py to apply, build, verify, and refresh package_result.json.",
        }, ensure_ascii=False, indent=2))
        return

    coverage_path, coverage = run_invoice_coverage(script_dir, manifest_path, out_dir, args.ocr)
    coverage_status = coverage.get("status")
    run_workflow_reports(script_dir, manifest_path, out_dir)
    _, current_approval_issues = write_approval_gap_report(manifest_path, out_dir)
    blocking_status = prerequisite_status(coverage_status, current_approval_issues)
    if blocking_status:
        package_result_path, package_result = write_package_result(
            out_dir,
            blocking_status,
            manifest_path,
            preflight_path,
            inspection_path,
            form_cells_source=form_cells_source,
            form_cells_origin=form_cells_origin,
        )
        print(json.dumps({
            "status": blocking_status,
            "manifest": str(manifest_path),
            "confirmation_report": package_result.get("confirmation_report", ""),
            "exception_report": package_result.get("exception_report", ""),
            "approval_gap_report": package_result.get("approval_gap_report", ""),
            "approval_issues": current_approval_issues,
            "invoice_coverage": str(coverage_path),
            "invoice_gap_report": package_result.get("invoice_gap_report", ""),
            "invoice_required_amount": coverage.get("required_amount", ""),
            "invoice_recognized_amount": coverage.get("recognized_amount", ""),
            "invoice_missing_amount": coverage.get("missing_amount", ""),
            "package_result": str(package_result_path),
            "next_step": package_result.get("next_step", ""),
        }, ensure_ascii=False, indent=2))
        return

    pack_cmd = [
        sys.executable,
        str(script_dir / "build_reimbursement_pack.py"),
        "--manifest", str(manifest_path),
        "--out", str(out_dir),
    ]
    run(pack_cmd)
    verification_path = out_dir / "verification.json"
    status = "built_unverified"
    if not args.skip_verify:
        verify_cmd = [
            sys.executable,
            str(script_dir / "verify_reimbursement_pack.py"),
            "--pack", str(out_dir),
            "--manifest", str(manifest_path),
        ]
        verify_result = run_no_check(verify_cmd)
        status = "passed" if verify_result.returncode == 0 else "failed_verification"
        finalize_audit_summary(out_dir, verification_path, verify_result.returncode == 0)
        package_result_path, package_result = write_package_result(out_dir, status, manifest_path, preflight_path, inspection_path, verification_path=verification_path, form_cells_source=form_cells_source, form_cells_origin=form_cells_origin)
        print(json.dumps({
            "status": status,
            "package_result": str(package_result_path),
            "delivery_checklist": package_result.get("delivery_checklist", ""),
            "finance_zip": package_result.get("finance_zip", ""),
            "one_click_print_pdf": package_result.get("one_click_print_pdf", ""),
        }, ensure_ascii=False, indent=2))
        if verify_result.returncode != 0:
            raise SystemExit(verify_result.returncode)
        return
    package_result_path, package_result = write_package_result(out_dir, status, manifest_path, preflight_path, inspection_path, verification_path=verification_path, form_cells_source=form_cells_source, form_cells_origin=form_cells_origin)
    print(json.dumps({
        "status": status,
        "package_result": str(package_result_path),
        "delivery_checklist": package_result.get("delivery_checklist", ""),
        "finance_zip": package_result.get("finance_zip", ""),
        "one_click_print_pdf": package_result.get("one_click_print_pdf", ""),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    from runtime_utils import configure_utf8_stdio

    configure_utf8_stdio()
    main()
