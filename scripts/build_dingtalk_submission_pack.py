#!/usr/bin/env python3
import argparse
import hashlib
import json
import shutil
import sys
import zipfile
from copy import deepcopy
from decimal import Decimal
from pathlib import Path

from build_reimbursement_pack import (
    build_expense_form,
    build_expense_form_pdf,
    build_invoice_pdf,
    collect_invoices,
    invoice_items_by_path,
    load_manifest,
    manifest_context,
    money,
    pdf_page_count,
    prepare_image_for_output,
    safe_name,
    validate_manifest,
)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Build the reimbursement form and invoice attachments needed before submitting DingTalk."
    )
    parser.add_argument("--manifest", required=True, help="Reviewed manifest with covered invoice amount.")
    parser.add_argument("--out", required=True, help="Output folder for the DingTalk submission materials.")
    return parser.parse_args()


def sha256(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def copy_submission_invoices(manifest, destination):
    destination.mkdir(parents=True, exist_ok=True)
    invoice_items = invoice_items_by_path(manifest)
    copied = []
    for index, source in enumerate(collect_invoices(manifest), start=1):
        item = invoice_items.get(str(source), {})
        role = "替代发票" if item.get("role") == "replacement" else "原始发票"
        target = destination / f"{index:03d}_{role}_{safe_name(source.stem)}{source.suffix.lower()}"
        shutil.copy2(source, target)
        if sha256(source) != sha256(target):
            raise SystemExit(f"Copied invoice hash mismatch: {source}")
        copied.append(target)
    for index, source in enumerate(manifest.get("supporting_document_paths", []), start=1):
        target = destination / f"{index:03d}_发票附件_行程单_{safe_name(source.stem)}{source.suffix.lower()}"
        shutil.copy2(source, target)
        if sha256(source) != sha256(target):
            raise SystemExit(f"Copied supporting-document hash mismatch: {source}")
        copied.append(target)
    return copied


def write_zip(folder, output):
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(
            item
            for item in folder.rglob("*")
            if item.is_file() and not item.name.startswith(".")
        ):
            archive.write(path, Path(folder.name) / path.relative_to(folder))
    return output


def main():
    args = parse_args()
    manifest = load_manifest(args.manifest)
    validation_manifest = deepcopy(manifest)
    validation_manifest["approval_requirement"] = {"policy": "none"}
    validate_manifest(validation_manifest)
    approval = manifest.get("approval_metadata") or {}
    missing_submission_fields = [
        field
        for field in ["reimbursement_type", "payee", "invoice_entity"]
        if approval.get(field) in (None, "")
    ]
    if missing_submission_fields:
        raise SystemExit(
            "DingTalk submission form requires pre-submission fields: "
            + ", ".join(missing_submission_fields)
        )
    source_manifest = deepcopy(manifest)

    root = Path(args.out).expanduser().resolve()
    marker = root / ".generated-by-expense-reimbursement-packager"
    if root.exists():
        if marker.exists():
            shutil.rmtree(root)
        elif any(root.iterdir()):
            raise SystemExit(f"Refusing to replace non-generated output folder: {root}")
    form_dir = root / "01_费用报销单"
    invoice_dir = root / "02_发票及行程单"
    print_dir = root / "03_便捷打印"
    converted_dir = root / "_converted_images"
    for folder in [form_dir, invoice_dir, print_dir, converted_dir]:
        folder.mkdir(parents=True, exist_ok=True)
    marker.write_text("dingtalk_submission\n", encoding="ascii")
    dirs = {
        "root": root,
        "tables": form_dir,
        "print": print_dir,
        "converted": converted_dir,
    }

    for entry in manifest["entries"]:
        entry["invoice_paths"] = [
            prepare_image_for_output(invoice, dirs) for invoice in entry.get("invoice_paths", [])
        ]
    manifest["invoice_paths"] = [
        prepare_image_for_output(invoice, dirs) for invoice in manifest.get("invoice_paths", [])
    ]
    manifest["supporting_document_paths"] = [
        prepare_image_for_output(document, dirs)
        for document in manifest.get("supporting_document_paths", [])
    ]
    form = dict(manifest.get("reimbursement_form") or {})
    template_path = form.get("template_path")
    if template_path:
        form["output_name"] = (
            f"{safe_name(manifest['project_name'])}_{safe_name(manifest['reimburser'])}_费用报销单_钉钉提交版"
            f"{template_path.suffix}"
        )
    manifest["reimbursement_form"] = form
    expense_form = build_expense_form(manifest, dirs)
    expense_form_pdf = build_expense_form_pdf(manifest, dirs, expense_form)
    if expense_form_pdf:
        actual_form_pages = pdf_page_count(expense_form_pdf)
        expected_form_pages = manifest_context(manifest)["expense_form_print_pages"]
        if actual_form_pages != expected_form_pages:
            manifest["_expense_form_print_pages"] = actual_form_pages
            expense_form = build_expense_form(manifest, dirs)
            expense_form_pdf = build_expense_form_pdf(manifest, dirs, expense_form)
            if pdf_page_count(expense_form_pdf) != actual_form_pages:
                raise SystemExit("DingTalk submission form page count changed after refreshing the attachment total.")
    if expense_form_pdf:
        renamed_pdf = expense_form_pdf.with_name(
            f"{safe_name(manifest['project_name'])}_费用报销单_钉钉提交版.pdf"
        )
        expense_form_pdf.replace(renamed_pdf)
        expense_form_pdf = renamed_pdf
    invoice_pdf = build_invoice_pdf(manifest, dirs)
    copied_documents = copy_submission_invoices(source_manifest, invoice_dir)
    shutil.rmtree(converted_dir, ignore_errors=True)

    required_fields = (manifest.get("approval_requirement") or {}).get("required_fields", []) or []
    pending_fields = [field for field in required_fields if approval.get(field) in (None, "")]
    total = sum((money(entry["amount"]) for entry in manifest.get("entries", [])), Decimal("0.00"))
    checklist = root / "钉钉提交清单.md"
    checklist.write_text(
        "# 钉钉提交清单\n\n"
        f"项目：{manifest.get('project_name', '')}\n\n"
        f"报销人：{manifest.get('reimburser', '')}\n\n"
        f"报销金额：{total:.2f} 元\n\n"
        f"发票覆盖：{(manifest.get('invoice_coverage') or {}).get('recognized_amount', '0.00')} 元\n\n"
        "## 上传材料\n\n"
        f"- 费用报销单：{expense_form.name if expense_form else '未生成'}\n"
        f"- 费用报销单 PDF：{expense_form_pdf.name if expense_form_pdf else '未生成'}\n"
        f"- 发票及行程单原件：{len(copied_documents)} 个文件\n"
        f"- 发票便捷打印 PDF：{invoice_pdf.name if invoice_pdf else '无'}\n\n"
        "## 说明\n\n"
        "这是钉钉提交前使用的报销单。钉钉审批编号等尚未产生的字段允许留空；"
        "提交后请提供审批截图，系统会核对审批金额并重新生成最终财务版。\n\n"
        f"待审批后补齐字段：{', '.join(pending_fields) if pending_fields else '无'}\n",
        encoding="utf-8",
    )

    if not expense_form or not expense_form.exists() or not expense_form_pdf or pdf_page_count(expense_form_pdf) < 1:
        raise SystemExit("DingTalk submission form or its PDF was not generated correctly.")
    if (manifest.get("invoice_requirement") or {}).get("policy", "full_amount") == "full_amount":
        if not invoice_pdf or pdf_page_count(invoice_pdf) < 1 or not copied_documents:
            raise SystemExit("Covered invoices were not included in the DingTalk submission materials.")

    zip_path = root.parent / f"{safe_name(manifest['project_name'])}_{safe_name(manifest['reimburser'])}_钉钉提交材料.zip"
    write_zip(root, zip_path)
    result = {
        "status": "passed",
        "stage": "dingtalk_submission",
        "folder": str(root),
        "zip": str(zip_path),
        "checklist": str(checklist),
        "expense_form": str(expense_form),
        "expense_form_pdf": str(expense_form_pdf),
        "invoice_pdf": str(invoice_pdf) if invoice_pdf else "",
        "document_count": len(copied_documents),
        "total_amount": f"{total:.2f}",
        "invoice_coverage_status": (manifest.get("invoice_coverage") or {}).get("status", ""),
        "pending_approval_fields": pending_fields,
    }
    result_path = root.parent / "dingtalk_submission_result.json"
    result["result"] = str(result_path)
    result_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    from runtime_utils import configure_utf8_stdio

    configure_utf8_stdio()
    main()
