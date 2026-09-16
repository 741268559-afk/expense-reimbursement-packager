#!/usr/bin/env python3
import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser(description="Onboard a real 费用报销单 template into a reusable portable form config.")
    parser.add_argument("--template", required=True, help="Path to the real 费用报销单 Excel template.")
    parser.add_argument("--out", required=True, help="Output folder for the portable form config.")
    parser.add_argument("--sheet", help="Optional sheet name to inspect/fill.")
    parser.add_argument("--copy-template", action="store_true", help="Copy the template into --out and make form_cells.json reference the copied file.")
    parser.add_argument("--sample-input", help="Optional screenshot folder or .zip to run a sample package validation.")
    parser.add_argument("--sample-invoice-input", action="append", default=[], help="Optional sample invoice file/folder/.zip. Can be passed multiple times.")
    parser.add_argument("--project-name", default="模板接入测试", help="Project name for optional sample validation.")
    parser.add_argument("--reimburser", default="测试报销人", help="Reimburser for optional sample validation.")
    parser.add_argument("--ocr", choices=["auto", "apple-vision", "none"], default="auto", help="OCR mode for optional sample validation.")
    parser.add_argument("--review-policy", choices=["auto", "always", "never"], default="auto", help="Review policy for optional sample validation.")
    return parser.parse_args()


def run(cmd):
    result = subprocess.run(cmd, text=True, capture_output=True)
    if result.returncode != 0:
        if result.stdout:
            print(result.stdout, file=sys.stderr, end="")
        if result.stderr:
            print(result.stderr, file=sys.stderr, end="")
        raise SystemExit(result.returncode)
    return result.stdout


def main():
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    template = Path(args.template).expanduser().resolve()
    if not template.exists():
        raise SystemExit(f"Template not found: {template}")
    out_dir = Path(args.out).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    active_template = template
    if args.copy_template:
        active_template = out_dir / template.name
        if active_template.resolve() != template:
            shutil.copy2(template, active_template)

    inspection = out_dir / "form_inspection.json"
    cells = out_dir / "form_cells.json"
    inspect_cmd = [
        sys.executable,
        str(script_dir / "inspect_reimbursement_form.py"),
        "--template",
        str(active_template),
        "--out",
        str(inspection),
        "--cells-out",
        str(cells),
        "--relative-template",
    ]
    if args.sheet:
        inspect_cmd.extend(["--sheet", args.sheet])
    run(inspect_cmd)
    validation = out_dir / "form_config_validation.json"
    validation_cmd = [
        sys.executable,
        str(script_dir / "validate_form_config.py"),
        "--form-cells",
        str(cells),
        "--out",
        str(validation),
    ]
    validation_result = subprocess.run(validation_cmd, text=True, capture_output=True)
    if validation.exists():
        validation_report = json.loads(validation.read_text(encoding="utf-8"))
    else:
        validation_report = {
            "status": "failed",
            "errors": ["validate_form_config.py did not write a report."],
            "stdout": validation_result.stdout,
            "stderr": validation_result.stderr,
        }
    readiness = out_dir / "readiness_report.json"
    readiness_cmd = [
        sys.executable,
        str(script_dir / "check_readiness.py"),
        "--form-cells",
        str(cells),
        "--out",
        str(readiness),
    ]
    readiness_result = subprocess.run(readiness_cmd, text=True, capture_output=True)
    if readiness.exists():
        readiness_report = json.loads(readiness.read_text(encoding="utf-8"))
    else:
        readiness_report = {
            "status": "failed",
            "errors": ["check_readiness.py did not write a report."],
            "stdout": readiness_result.stdout,
            "stderr": readiness_result.stderr,
        }

    form_cells = json.loads(cells.read_text(encoding="utf-8"))
    report = {
        "status": "needs_review" if not form_cells.get("cells") or validation_report.get("status") == "failed" else "ready_for_review",
        "template": str(active_template),
        "form_inspection": str(inspection),
        "form_cells": str(cells),
        "form_config_validation": str(validation),
        "form_config_validation_status": validation_report.get("status", ""),
        "form_config_validation_returncode": validation_result.returncode,
        "readiness_report": str(readiness),
        "readiness_status": readiness_report.get("status", ""),
        "readiness_returncode": readiness_result.returncode,
        "cell_count": len(form_cells.get("cells", {})),
        "sheet": form_cells.get("sheet", ""),
        "next_step": "Open form_inspection.json and the template, then confirm or edit form_cells.json before real reimbursement packaging.",
    }

    if args.sample_input:
        sample_out = out_dir / "sample_validation_pack"
        package_cmd = [
            sys.executable,
            str(script_dir / "package_from_folder.py"),
            "--input",
            str(Path(args.sample_input).expanduser().resolve()),
            "--project-name",
            args.project_name,
            "--reimburser",
            args.reimburser,
            "--out",
            str(sample_out),
            "--form-cells",
            str(cells),
            "--ocr",
            args.ocr,
            "--review-policy",
            args.review_policy,
        ]
        for invoice_input in args.sample_invoice_input:
            package_cmd.extend(["--invoice-input", str(Path(invoice_input).expanduser().resolve())])
        sample_result = subprocess.run(package_cmd, text=True, capture_output=True)
        report["sample_validation"] = {
            "output_folder": str(sample_out),
            "returncode": sample_result.returncode,
            "stdout": sample_result.stdout,
            "stderr": sample_result.stderr,
            "verification": str(sample_out / "verification.json"),
        }
        verification_path = sample_out / "verification.json"
        if verification_path.exists():
            verification = json.loads(verification_path.read_text(encoding="utf-8"))
            report["sample_validation"]["status"] = verification.get("status", "")
        elif sample_result.returncode == 0 and "\"status\": \"needs_review\"" in sample_result.stdout:
            report["sample_validation"]["status"] = "needs_review"
        elif sample_result.returncode == 0:
            report["sample_validation"]["status"] = "completed_without_verification"
        else:
            report["sample_validation"]["status"] = "failed"
            report["sample_validation"]["next_step"] = "Inspect stdout/stderr, usually OCR or review rows need correction before sample validation can build a final pack."

    report_path = out_dir / "onboarding_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
