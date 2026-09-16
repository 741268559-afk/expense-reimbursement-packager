#!/usr/bin/env python3
import argparse
import json
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path, PurePosixPath


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".heic", ".heif"}
PDF_EXTS = {".pdf"}
SUPPORTED_EXTS = IMAGE_EXTS | PDF_EXTS
INVOICE_HINTS = ("发票", "invoice", "fapiao")
SUPPORTING_HINTS = ("行程单", "行程报销单", "trip table")


@dataclass(frozen=True)
class FileItem:
    display: str
    name: str
    suffix: str
    origin: str


def parse_args():
    parser = argparse.ArgumentParser(description="Preflight reimbursement input folders, files, or zip archives before packaging.")
    parser.add_argument("--input", required=True, help="Folder, file, or .zip containing expense screenshots and possibly invoices.")
    parser.add_argument("--invoice-input", action="append", default=[], help="Additional invoice file, folder, or .zip. Can be passed multiple times.")
    parser.add_argument("--replacement-invoice-input", action="append", default=[], help="Finance-approved replacement invoice file, folder, or .zip. Can be passed multiple times.")
    parser.add_argument("--supporting-document-input", action="append", default=[], help="Invoice supporting document file, folder, or .zip. Not counted toward invoice coverage.")
    parser.add_argument("--recursive", action="store_true", help="Scan folders recursively.")
    parser.add_argument("--out", help="Optional JSON report path. Defaults to stdout only.")
    return parser.parse_args()


def to_file_item(path):
    return FileItem(
        display=str(path),
        name=path.name,
        suffix=path.suffix.lower(),
        origin="filesystem",
    )


def should_skip_archive_member(name):
    member_path = PurePosixPath(name)
    parts = member_path.parts
    return (
        not name
        or name.endswith("/")
        or "__MACOSX" in parts
        or "__pycache__" in parts
        or member_path.name == ".DS_Store"
        or member_path.name.startswith("._")
        or member_path.suffix == ".pyc"
    )


def assert_safe_archive_member(name):
    member_path = PurePosixPath(name)
    if member_path.is_absolute() or any(part == ".." for part in member_path.parts):
        raise SystemExit(f"Unsafe zip member path: {name}")


def walk_zip(path):
    items = []
    with zipfile.ZipFile(path) as zf:
        for info in sorted(zf.infolist(), key=lambda item: item.filename):
            if should_skip_archive_member(info.filename):
                continue
            assert_safe_archive_member(info.filename)
            member_path = PurePosixPath(info.filename)
            items.append(FileItem(
                display=f"{path}::{info.filename}",
                name=member_path.name,
                suffix=member_path.suffix.lower(),
                origin=f"zip:{path}",
            ))
    return items


def walk_path(value, recursive=False):
    path = Path(value).expanduser().resolve()
    if path.is_file():
        if path.suffix.lower() == ".zip":
            return walk_zip(path)
        return [to_file_item(path)]
    if path.is_dir():
        pattern = "**/*" if recursive else "*"
        return [to_file_item(item.resolve()) for item in sorted(path.glob(pattern)) if item.is_file()]
    raise SystemExit(f"Input not found: {path}")


def is_invoice_hint(item):
    lowered = item.name.lower()
    return item.suffix in PDF_EXTS or any(hint in lowered for hint in INVOICE_HINTS)


def is_supporting_hint(item):
    lowered = item.name.lower()
    return any(hint in lowered for hint in SUPPORTING_HINTS)


def classify_files(input_files, invoice_files, supporting_files):
    invoice_set = {item.display for item in invoice_files if item.suffix in SUPPORTED_EXTS}
    supporting_set = {item.display for item in supporting_files if item.suffix in SUPPORTED_EXTS}
    supported_input = [item for item in input_files if item.suffix in SUPPORTED_EXTS]
    ignored = [item for item in input_files + invoice_files + supporting_files if item.suffix not in SUPPORTED_EXTS]

    inferred_invoices = []
    inferred_supporting = []
    expense_screenshots = []
    for item in supported_input:
        if item.display in supporting_set or is_supporting_hint(item):
            inferred_supporting.append(item)
        elif item.display in invoice_set or is_invoice_hint(item):
            inferred_invoices.append(item)
        elif item.suffix in IMAGE_EXTS:
            expense_screenshots.append(item)

    explicit_invoices = [item for item in invoice_files if item.suffix in SUPPORTED_EXTS]
    all_invoices = []
    seen = set()
    for item in inferred_invoices + explicit_invoices:
        key = item.display
        if key in seen:
            continue
        seen.add(key)
        all_invoices.append(item)
    all_supporting = []
    seen = set()
    for item in inferred_supporting + [item for item in supporting_files if item.suffix in SUPPORTED_EXTS]:
        key = item.display
        if key in seen:
            continue
        seen.add(key)
        all_supporting.append(item)
    return expense_screenshots, all_invoices, all_supporting, ignored


def duplicate_basenames(items):
    by_name = defaultdict(list)
    for item in items:
        by_name[item.name].append(item)
    return {name: [item.display for item in values] for name, values in by_name.items() if len(values) > 1}


def main():
    args = parse_args()
    input_files = walk_path(args.input, args.recursive)
    invoice_files = []
    for value in args.invoice_input:
        invoice_files.extend(walk_path(value, args.recursive))
    replacement_invoice_files = []
    for value in args.replacement_invoice_input:
        replacement_invoice_files.extend(walk_path(value, args.recursive))
    supporting_document_files = []
    for value in args.supporting_document_input:
        supporting_document_files.extend(walk_path(value, args.recursive))

    screenshots, invoices, supporting_documents, ignored = classify_files(
        input_files,
        invoice_files + replacement_invoice_files,
        supporting_document_files,
    )
    warnings = []
    if not screenshots:
        warnings.append("No expense screenshots were detected.")
    if not invoices:
        warnings.append("No invoice files were detected. Under full_amount policy, packaging will stop after expense review and report the full invoice gap.")
    if ignored:
        warnings.append("Some files use unsupported extensions and will be ignored.")
    if any(item.suffix in PDF_EXTS for item in input_files):
        warnings.append("PDF files in --input are classified as invoices or supporting documents from their filenames; review the lists below.")
    if any(item.suffix in IMAGE_EXTS for item in invoice_files + replacement_invoice_files):
        warnings.append("Image files in --invoice-input are treated as invoices.")

    duplicates = duplicate_basenames(input_files + invoice_files + replacement_invoice_files + supporting_document_files)
    if duplicates:
        warnings.append("Duplicate basenames detected; output files will still be numbered, but review source grouping.")

    report = {
        "input": str(Path(args.input).expanduser().resolve()),
        "invoice_inputs": [str(Path(value).expanduser().resolve()) for value in args.invoice_input],
        "replacement_invoice_inputs": [str(Path(value).expanduser().resolve()) for value in args.replacement_invoice_input],
        "supporting_document_inputs": [str(Path(value).expanduser().resolve()) for value in args.supporting_document_input],
        "recursive": args.recursive,
        "counts": {
            "expense_screenshots": len(screenshots),
            "invoices": len(invoices),
            "replacement_invoice_files": len(replacement_invoice_files),
            "supporting_documents": len(supporting_documents),
            "ignored": len(ignored),
            "total_seen": len(input_files) + len(invoice_files) + len(replacement_invoice_files) + len(supporting_document_files),
        },
        "expense_screenshots": [item.display for item in screenshots],
        "invoices": [item.display for item in invoices],
        "supporting_documents": [item.display for item in supporting_documents],
        "ignored": [item.display for item in ignored],
        "duplicate_basenames": duplicates,
        "warnings": warnings,
        "status": "needs_review" if warnings else "ready",
    }
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out:
        output = Path(args.out).expanduser().resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(text + "\n", encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
