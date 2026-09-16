#!/usr/bin/env python3
import argparse
import json
import sys
from pathlib import Path

from package_from_folder import (
    approval_is_pre_submission,
    apply_profile_and_approval,
    build_dingtalk_submission,
    finalize_audit_summary,
    load_json_arg,
    needs_review,
    prerequisite_status,
    resolve_profile_source,
    run,
    run_invoice_coverage,
    run_no_check,
    run_workflow_reports,
    submission_info_issues,
    write_approval_gap_report,
    write_package_result,
)


def parse_args():
    parser = argparse.ArgumentParser(description="Apply an edited manifest review workbook, build the reimbursement pack, and verify it.")
    parser.add_argument("--manifest", required=True, help="Original draft_manifest.json.")
    parser.add_argument("--workbook", required=True, help="Edited draft_manifest_review.xlsx.")
    parser.add_argument("--out", required=True, help="Output folder for the generated pack.")
    parser.add_argument("--reviewed-manifest-out", help="Output reviewed manifest JSON. Defaults to <out>/reviewed_manifest.json.")
    parser.add_argument("--preflight-report", help="Optional preflight_report.json. Defaults to <out>/preflight_report.json when present.")
    parser.add_argument("--form-inspection", help="Optional form_inspection.json. Defaults to <out>/form_inspection.json when present.")
    parser.add_argument("--allow-new-categories", action="store_true", help="Allow new expense_type values outside manifest categories while applying workbook edits.")
    parser.add_argument("--reimburser", help="Optional reimburser override; normally retained from the draft manifest or profile.")
    parser.add_argument("--profile", help="Optional reimbursement profile JSON path or JSON object.")
    parser.add_argument("--approval-metadata", help="Pre-submission form fields and later DingTalk screenshot fields, as a JSON path or JSON object.")
    parser.add_argument("--approval-policy", choices=["required", "none"], help="Optional approval gate override; defaults to the draft manifest policy.")
    parser.add_argument("--ocr", choices=["auto", "apple-vision", "tesseract", "none"], default="auto", help="OCR mode for image invoices.")
    parser.add_argument("--skip-verify", action="store_true", help="Build the pack but skip automatic verification.")
    return parser.parse_args()


def existing_default(path):
    return path if path.exists() else None


def main():
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    out_dir = Path(args.out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = Path(args.manifest).expanduser().resolve()
    workbook_path = Path(args.workbook).expanduser().resolve()
    reviewed_manifest = Path(args.reviewed_manifest_out).expanduser().resolve() if args.reviewed_manifest_out else out_dir / "reviewed_manifest.json"
    preflight_path = Path(args.preflight_report).expanduser().resolve() if args.preflight_report else existing_default(out_dir / "preflight_report.json")
    form_inspection = Path(args.form_inspection).expanduser().resolve() if args.form_inspection else existing_default(out_dir / "form_inspection.json")

    apply_cmd = [
        sys.executable,
        str(script_dir / "manifest_review_workbook.py"),
        "apply",
        "--manifest",
        str(manifest_path),
        "--workbook",
        str(workbook_path),
        "--out",
        str(reviewed_manifest),
    ]
    if args.allow_new_categories:
        apply_cmd.append("--allow-new-categories")
    run(apply_cmd)

    reviewed_data = json.loads(reviewed_manifest.read_text(encoding="utf-8"))
    profile_source, _ = resolve_profile_source(args)
    if args.profile:
        profile = load_json_arg(args.profile)
    elif profile_source:
        profile = load_json_arg(str(profile_source))
    else:
        profile = {}
    args.reimburser = str(args.reimburser or reviewed_data.get("reimburser") or profile.get("reimburser") or "").strip()
    args.approval_policy = args.approval_policy or (reviewed_data.get("approval_requirement") or {}).get("policy", "required")
    apply_profile_and_approval(reviewed_manifest, args, profile, profile_source)
    run_workflow_reports(script_dir, reviewed_manifest, out_dir)
    write_approval_gap_report(reviewed_manifest, out_dir)
    unresolved_review_rows = needs_review(reviewed_manifest)
    if unresolved_review_rows:
        package_result_path, package_result = write_package_result(
            out_dir,
            "needs_review",
            reviewed_manifest,
            preflight_path,
            form_inspection,
            unresolved_review_rows,
            review_workbook=workbook_path,
        )
        print(json.dumps({
            "status": "needs_review",
            "package_result": str(package_result_path),
            "manifest": str(reviewed_manifest),
            "review_rows": unresolved_review_rows,
            "next_step": package_result.get("next_step", ""),
        }, ensure_ascii=False, indent=2))
        return

    coverage_path, coverage = run_invoice_coverage(script_dir, reviewed_manifest, out_dir, args.ocr)
    coverage_status = coverage.get("status")
    run_workflow_reports(script_dir, reviewed_manifest, out_dir)
    _, current_approval_issues = write_approval_gap_report(reviewed_manifest, out_dir)
    blocking_status = prerequisite_status(coverage_status, current_approval_issues)
    current_manifest = json.loads(reviewed_manifest.read_text(encoding="utf-8"))
    current_submission_info_issues = submission_info_issues(current_manifest)
    approval_needs_correction = any(
        item.get("type") in {"invalid", "mismatch"}
        for item in current_approval_issues
    )
    if blocking_status == "needs_approval" and current_submission_info_issues and not approval_needs_correction:
        package_result_path, package_result = write_package_result(
            out_dir,
            "needs_submission_info",
            reviewed_manifest,
            preflight_path,
            form_inspection,
            review_workbook=workbook_path,
        )
        print(json.dumps({
            "status": "needs_submission_info",
            "package_result": str(package_result_path),
            "manifest": str(reviewed_manifest),
            "manifest_review_workbook": str(workbook_path),
            "submission_info_issues": current_submission_info_issues,
            "next_step": package_result.get("next_step", ""),
        }, ensure_ascii=False, indent=2))
        return
    if blocking_status == "needs_approval" and approval_is_pre_submission(current_approval_issues):
        submission = build_dingtalk_submission(script_dir, reviewed_manifest, out_dir)
        package_result_path, package_result = write_package_result(
            out_dir,
            "ready_for_dingtalk",
            reviewed_manifest,
            preflight_path,
            form_inspection,
            review_workbook=workbook_path,
        )
        print(json.dumps({
            "status": "ready_for_dingtalk",
            "package_result": str(package_result_path),
            "manifest": str(reviewed_manifest),
            "manifest_review_workbook": str(workbook_path),
            "dingtalk_submission_folder": submission.get("folder", ""),
            "dingtalk_submission_zip": submission.get("zip", ""),
            "dingtalk_submission_expense_form": submission.get("expense_form", ""),
            "dingtalk_submission_expense_form_pdf": submission.get("expense_form_pdf", ""),
            "dingtalk_submission_invoice_pdf": submission.get("invoice_pdf", ""),
            "next_step": package_result.get("next_step", ""),
        }, ensure_ascii=False, indent=2))
        return
    if blocking_status:
        package_result_path, package_result = write_package_result(
            out_dir,
            blocking_status,
            reviewed_manifest,
            preflight_path,
            form_inspection,
            review_workbook=workbook_path,
        )
        print(json.dumps({
            "status": blocking_status,
            "package_result": str(package_result_path),
            "manifest": str(reviewed_manifest),
            "manifest_review_workbook": str(workbook_path),
            "invoice_coverage": str(coverage_path),
            "invoice_gap_report": package_result.get("invoice_gap_report", ""),
            "invoice_missing_amount": coverage.get("missing_amount", ""),
            "approval_issues": current_approval_issues,
            "next_step": package_result.get("next_step", ""),
        }, ensure_ascii=False, indent=2))
        return

    run([
        sys.executable,
        str(script_dir / "build_reimbursement_pack.py"),
        "--manifest",
        str(reviewed_manifest),
        "--out",
        str(out_dir),
    ])

    verification_path = out_dir / "verification.json"
    status = "built_unverified"
    if not args.skip_verify:
        verify_result = run_no_check([
            sys.executable,
            str(script_dir / "verify_reimbursement_pack.py"),
            "--pack",
            str(out_dir),
            "--manifest",
            str(reviewed_manifest),
        ])
        status = "passed" if verify_result.returncode == 0 else "failed_verification"
        finalize_audit_summary(out_dir, verification_path, verify_result.returncode == 0)
    package_result_path, package_result = write_package_result(
        out_dir,
        status,
        reviewed_manifest,
        preflight_path,
        form_inspection,
        verification_path=verification_path,
        review_workbook=workbook_path,
    )
    output = {
        "status": status,
        "package_result": str(package_result_path),
        "manifest": str(reviewed_manifest),
        "manifest_review_workbook": str(workbook_path),
        "delivery_checklist": package_result.get("delivery_checklist", ""),
        "finance_zip": package_result.get("finance_zip", ""),
        "one_click_print_pdf": package_result.get("one_click_print_pdf", ""),
    }
    print(json.dumps(output, ensure_ascii=False, indent=2))
    if status == "failed_verification":
        raise SystemExit(1)


if __name__ == "__main__":
    from runtime_utils import configure_utf8_stdio

    configure_utf8_stdio()
    main()
