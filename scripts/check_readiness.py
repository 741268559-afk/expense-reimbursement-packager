#!/usr/bin/env python3
import argparse
import importlib
import json
import platform
import subprocess
import sys
from pathlib import Path

from ocr_utils import backend_status


REQUIRED_MODULES = ["openpyxl", "PIL", "pypdf", "reportlab"]
OPTIONAL_MODULES = ["pillow_heif"]
REQUIRED_SCRIPTS = [
    "analyze_invoice_coverage.py",
    "bootstrap.py",
    "build_release_packages.py",
    "build_reimbursement_pack.py",
    "draft_manifest_from_folder.py",
    "inspect_reimbursement_form.py",
    "manifest_review_workbook.py",
    "ocr_utils.py",
    "onboard_reimbursement_form.py",
    "package_from_folder.py",
    "package_reviewed_workbook.py",
    "platform_smoke_test.py",
    "preflight_inputs.py",
    "prepare_expense_confirmation.py",
    "self_test.py",
    "validate_form_config.py",
    "verify_reimbursement_pack.py",
    "vision_ocr.swift",
]


def parse_args():
    parser = argparse.ArgumentParser(description="Check whether the reimbursement packager skill is ready for team use.")
    parser.add_argument("--form-cells", help="Optional reviewed form_cells.json to validate.")
    parser.add_argument("--template", help="Optional template override for --form-cells.")
    parser.add_argument("--run-self-test", action="store_true", help="Run the full self-test after lightweight readiness checks.")
    parser.add_argument("--self-test-out", help="Output folder for --run-self-test.")
    parser.add_argument("--out", help="Optional JSON report path.")
    return parser.parse_args()


def add_check(checks, name, status, details=""):
    checks.append({
        "name": name,
        "status": status,
        "details": details,
    })


def run_json_command(cmd):
    result = subprocess.run(cmd, text=True, capture_output=True)
    output = result.stdout.strip()
    data = {}
    if output:
        try:
            data = json.loads(output)
        except json.JSONDecodeError:
            data = {"raw_stdout": output}
    return result, data


def main():
    args = parse_args()
    script_dir = Path(__file__).resolve().parent
    checks = []
    errors = []
    warnings = []

    missing_scripts = []
    for script in REQUIRED_SCRIPTS:
        path = script_dir / script
        if path.exists():
            add_check(checks, f"script:{script}", "passed", str(path))
        else:
            missing_scripts.append(script)
            add_check(checks, f"script:{script}", "failed", "Missing required script.")
    if missing_scripts:
        errors.append(f"Missing required script(s): {', '.join(missing_scripts)}")

    missing_modules = []
    for module in REQUIRED_MODULES:
        try:
            importlib.import_module(module)
            add_check(checks, f"python_module:{module}", "passed")
        except Exception as exc:
            missing_modules.append(module)
            add_check(checks, f"python_module:{module}", "failed", str(exc))
    if missing_modules:
        errors.append(f"Missing required Python module(s): {', '.join(missing_modules)}")

    for module in OPTIONAL_MODULES:
        try:
            importlib.import_module(module)
            add_check(checks, f"python_module:{module}", "passed")
        except Exception as exc:
            warnings.append(f"Optional module {module} is unavailable; HEIC/HEIF support may be limited: {exc}")
            add_check(checks, f"python_module:{module}", "warning", str(exc))

    system = platform.system()
    add_check(checks, "platform", "passed", f"{system} / Python {platform.python_version()}")
    ocr = backend_status(script_dir)
    if system == "Darwin":
        if ocr["apple_vision"]:
            add_check(checks, "apple_vision_ocr", "passed", "Swift and Apple Vision OCR are available.")
        else:
            add_check(checks, "apple_vision_ocr", "warning", "Apple Vision OCR is unavailable.")
    else:
        add_check(checks, "apple_vision_ocr", "skipped", "Apple Vision is macOS-only.")
    if ocr["tesseract"]:
        language_text = ", ".join(ocr["tesseract_languages"]) or "languages not reported"
        if "chi_sim" in ocr["tesseract_languages"]:
            add_check(checks, "tesseract_ocr", "passed", f"{ocr['tesseract']} ({language_text})")
        else:
            warnings.append("Tesseract is installed but the chi_sim language pack is missing; Chinese OCR will be limited.")
            add_check(checks, "tesseract_ocr", "warning", f"{ocr['tesseract']} ({language_text})")
    else:
        add_check(checks, "tesseract_ocr", "warning", "Tesseract was not found.")
    if ocr["auto_backend"] == "none":
        warnings.append(
            "No automatic OCR backend is available; filename fallback and manual Excel review still work. "
            "Install Tesseract on Windows/macOS or Swift on macOS for local OCR."
        )
        add_check(checks, "ocr_auto_backend", "warning", "none")
    else:
        add_check(checks, "ocr_auto_backend", "passed", ocr["auto_backend"])

    form_validation = {}
    if args.form_cells:
        form_cmd = [
            sys.executable,
            str(script_dir / "validate_form_config.py"),
            "--form-cells",
            args.form_cells,
        ]
        if args.template:
            form_cmd.extend(["--template", args.template])
        form_result, form_validation = run_json_command(form_cmd)
        form_status = form_validation.get("status", "failed" if form_result.returncode else "passed")
        check_status = "passed" if form_result.returncode == 0 else "failed"
        add_check(checks, "form_config", check_status, form_status)
        if form_result.returncode != 0:
            errors.append("form_cells.json failed validation.")
        elif form_status == "passed_with_warnings":
            warnings.extend(form_validation.get("warnings", []))
    else:
        warnings.append("No form_cells.json provided; detail table and print pack can run, but the real 费用报销单 is not ready.")
        add_check(checks, "form_config", "warning", "No form_cells.json provided.")

    self_test = {}
    if args.run_self_test:
        self_test_out = Path(args.self_test_out).expanduser().resolve() if args.self_test_out else Path.cwd().resolve() / "expense_reimbursement_readiness_self_test"
        self_result, self_test = run_json_command([
            sys.executable,
            str(script_dir / "self_test.py"),
            "--out",
            str(self_test_out),
        ])
        if self_result.returncode == 0 and self_test.get("status") == "passed":
            add_check(checks, "self_test", "passed", str(self_test_out))
        else:
            add_check(checks, "self_test", "failed", self_result.stderr.strip() or self_result.stdout.strip())
            errors.append("self_test.py failed.")

    if errors:
        status = "failed"
        next_step = "Fix failed checks before processing reimbursements."
    elif not args.form_cells:
        status = "needs_form_template"
        next_step = "Add the real 费用报销单 template, run onboard_reimbursement_form.py, validate form_cells.json, then rerun this readiness check."
    else:
        status = "ready"
        next_step = "Run package_from_folder.py with --form-cells for real reimbursement folders."

    result = {
        "status": status,
        "skill_dir": str(script_dir.parent),
        "checks": checks,
        "warnings": warnings,
        "errors": errors,
        "form_validation": form_validation,
        "self_test": self_test,
        "next_step": next_step,
    }
    text = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        output = Path(args.out).expanduser().resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(text + "\n", encoding="utf-8")
    print(text)
    if status == "failed":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
