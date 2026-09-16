#!/usr/bin/env python3
import argparse
import json
import re
import shutil
import subprocess
import sys
from decimal import Decimal
from pathlib import Path


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".heic", ".heif"}
PDF_EXTS = {".pdf"}
INVOICE_HINTS = ("发票", "invoice", "fapiao")
SUPPORTING_HINTS = ("行程单", "行程报销单", "trip table")


def parse_args():
    parser = argparse.ArgumentParser(description="Draft a reimbursement manifest from a folder of screenshots and invoices.")
    parser.add_argument("--input", required=True, help="Folder containing expense screenshots and invoices.")
    parser.add_argument("--invoice-input", action="append", default=[], help="Additional invoice file or folder. Can be passed multiple times.")
    parser.add_argument("--replacement-invoice-input", action="append", default=[], help="Finance-approved replacement invoice file or folder. Can be passed multiple times.")
    parser.add_argument("--supporting-document-input", action="append", default=[], help="Invoice supporting document file or folder, such as ride itineraries. Not counted toward invoice coverage.")
    parser.add_argument("--project-name", required=True, help="Project/shoot name.")
    parser.add_argument("--reimburser", required=True, help="Person being reimbursed.")
    parser.add_argument("--out", required=True, help="Output manifest JSON path.")
    parser.add_argument("--recursive", action="store_true", help="Scan folders recursively.")
    parser.add_argument("--ocr", choices=["auto", "apple-vision", "none"], default="auto", help="OCR mode for image screenshots.")
    parser.add_argument("--categories", default="交通费,餐饮费,设备费,场地费,演员费,其他费用", help="Comma-separated category list.")
    parser.add_argument("--invoice-policy", choices=["full_amount", "none"], default="full_amount", help="Invoice coverage policy. full_amount requires invoices to cover the reimbursement total.")
    parser.add_argument("--approval-policy", choices=["required", "none"], default="required", help="Whether DingTalk metadata and approved amount are required before final packaging.")
    return parser.parse_args()


def money(value):
    return Decimal(str(value)).quantize(Decimal("0.01"))


def is_invoice_path(path):
    lowered = path.name.lower()
    return any(hint in lowered for hint in INVOICE_HINTS) or path.suffix.lower() in PDF_EXTS


def is_supporting_document_path(path):
    lowered = path.name.lower()
    return any(hint in lowered for hint in SUPPORTING_HINTS)


def walk_files(folder, recursive=False):
    pattern = "**/*" if recursive else "*"
    return sorted(p for p in folder.glob(pattern) if p.is_file())


def collect_invoice_inputs(values, recursive=False):
    invoices = []
    for value in values:
        path = Path(value).expanduser().resolve()
        if path.is_file():
            invoices.append(path)
        elif path.is_dir():
            invoices.extend(p for p in walk_files(path, recursive) if p.suffix.lower() in (IMAGE_EXTS | PDF_EXTS))
        else:
            raise SystemExit(f"Invoice input not found: {path}")
    return invoices


def unique_paths(paths):
    seen = set()
    unique = []
    for path in paths:
        key = str(path.resolve())
        if key in seen:
            continue
        seen.add(key)
        unique.append(path)
    return unique


def run_apple_vision_ocr(images, script_dir):
    swift = shutil.which("swift")
    if not swift or not images:
        return {}
    script = script_dir / "vision_ocr.swift"
    cmd = [swift, str(script), *[str(p) for p in images]]
    result = subprocess.run(cmd, check=True, text=True, capture_output=True)
    data = json.loads(result.stdout)
    return {Path(item["path"]).resolve(): item for item in data}


def extract_amount(text):
    candidates = []
    for match in re.finditer(r"[-−—]\s*¥?\s*(\d{1,6}(?:\.\d{1,2})?)", text):
        candidates.append(match.group(1))
    if not candidates:
        for match in re.finditer(r"¥\s*(\d{1,6}(?:\.\d{1,2})?)", text):
            candidates.append(match.group(1))
    if not candidates:
        return ""
    try:
        return float(money(candidates[0]))
    except Exception:
        return ""


def extract_amount_from_filename(filename):
    for match in re.finditer(r"(?<!\d)(\d{1,5}\.\d{1,2})(?!\d)", filename):
        try:
            return float(money(match.group(1)))
        except Exception:
            continue
    return ""


def normalize_date(year, month, day):
    return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"


def extract_datetime(text):
    match = re.search(r"(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日\s*(\d{1,2}:\d{2}(?::\d{2})?)?", text)
    if match:
        return normalize_date(match.group(1), match.group(2), match.group(3)), match.group(4) or ""
    match = re.search(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\s*(\d{1,2}:\d{2}(?::\d{2})?)?", text)
    if match:
        return normalize_date(match.group(1), match.group(2), match.group(3)), match.group(4) or ""
    return "", ""


def classify(text, filename):
    combined = f"{text} {filename}"
    rules = [
        (("货拉拉",), "设备配送费", "设备费"),
        (("滴滴", "快车", "特惠快车", "阳光出行", "打车", "出租车"), "打车费", "交通费"),
        (("地铁", "公共交通", "公交"), "地铁费", "交通费"),
        (("停车",), "停车费", "交通费"),
        (("机票", "火车", "高铁", "动车"), "交通费", "交通费"),
        (("摄影棚", "影棚", "场地", "场租", "棚租"), "场地租赁费", "场地费"),
        (("演员", "模特", "出镜", "艺人"), "演员费", "演员费"),
        (("设备租赁", "器材租赁", "镜头租赁", "相机租赁", "灯光租赁", "vivo", "X300", "增距"), "设备租赁费", "设备费"),
        (("设备", "器材", "配件", "闲鱼"), "设备配件费", "设备费"),
        (("外卖", "餐", "拉面", "老碗", "人和馆", "美团", "大众点评", "点餐", "闪购", "奈雪", "foodplace"), "餐费", "餐饮费"),
        (("顺丰", "快递", "先寄后付"), "快递费", "其他费用"),
        (("App Store", "Apple Music", "Apple"), "软件服务费", "其他费用"),
    ]
    for needles, purpose, expense_type in rules:
        if any(needle in combined for needle in needles):
            return purpose, expense_type
    return "", ""


def main():
    args = parse_args()
    folder = Path(args.input).expanduser().resolve()
    output = Path(args.out).expanduser().resolve()
    files = walk_files(folder, args.recursive)
    image_files = [p for p in files if p.suffix.lower() in IMAGE_EXTS]
    inferred_supporting_files = [p for p in files if is_supporting_document_path(p)]
    original_invoice_files = unique_paths(
        [p for p in files if is_invoice_path(p) and not is_supporting_document_path(p)]
        + collect_invoice_inputs(args.invoice_input, args.recursive)
    )
    replacement_invoice_files = unique_paths(collect_invoice_inputs(args.replacement_invoice_input, args.recursive))
    supporting_document_files = unique_paths(
        inferred_supporting_files + collect_invoice_inputs(args.supporting_document_input, args.recursive)
    )
    replacement_keys = {str(path.resolve()) for path in replacement_invoice_files}
    invoice_files = unique_paths(original_invoice_files + replacement_invoice_files)
    expense_images = [p for p in image_files if p not in invoice_files and p not in supporting_document_files]

    ocr_by_path = {}
    if args.ocr in {"auto", "apple-vision"}:
        try:
            ocr_by_path = run_apple_vision_ocr(expense_images, Path(__file__).resolve().parent)
        except Exception as exc:
            if args.ocr == "apple-vision":
                raise
            print(f"OCR unavailable, drafting from filenames only: {exc}", file=sys.stderr)

    entries = []
    for image in expense_images:
        ocr = ocr_by_path.get(image.resolve(), {})
        lines = ocr.get("lines", [])
        text = "\n".join(line.get("text", "") for line in lines)
        date, time = extract_datetime(text)
        if not date:
            date, time = extract_datetime(image.name)
        purpose, expense_type = classify(text, image.name)
        amount = extract_amount(text)
        if not amount:
            amount = extract_amount_from_filename(image.stem)
        entries.append({
            "date": date,
            "time": time,
            "purpose": purpose,
            "expense_type": expense_type,
            "amount": amount,
            "screenshot": str(image),
            "invoices": [],
            "merchant": "",
            "exception_resolution": "",
            "notes": "NEEDS_REVIEW" if not (date and purpose and expense_type and amount) else "",
            "ocr_text": text[:2000],
        })

    entries.sort(key=lambda e: (e.get("date") or "9999-12-31", e.get("time") or "23:59:59", e["screenshot"]))
    manifest = {
        "workflow_version": 2,
        "project_name": args.project_name,
        "reimburser": args.reimburser,
        "categories": [item.strip() for item in args.categories.split(",") if item.strip()],
        "invoices": [str(p) for p in invoice_files],
        "invoice_requirement": {
            "policy": args.invoice_policy,
        },
        "approval_requirement": {
            "policy": args.approval_policy,
            "required_fields": ["dingding_number", "reimbursement_type", "payee", "invoice_entity", "approved_amount"],
        },
        "invoice_items": [
            {
                "path": str(path),
                "role": "replacement" if str(path.resolve()) in replacement_keys else "original",
            }
            for path in invoice_files
        ],
        "supporting_documents": [str(path) for path in supporting_document_files],
        "reimbursement_form": {
            "template": "",
            "sheet": "",
            "output_name": "",
            "cells": {}
        },
        "entries": entries,
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({
        "manifest": str(output),
        "expense_screenshots": len(expense_images),
        "invoice_files": len(invoice_files),
        "replacement_invoice_files": len(replacement_invoice_files),
        "supporting_document_files": len(supporting_document_files),
        "needs_review": sum(1 for e in entries if e.get("notes")),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
