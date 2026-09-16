#!/usr/bin/env python3
import argparse
import hashlib
import json
import re
import shutil
import subprocess
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from pypdf import PdfReader


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".heic", ".heif"}
PDF_EXTS = {".pdf"}
VALID_ROLES = {"original", "replacement"}


def parse_args():
    parser = argparse.ArgumentParser(description="Analyze invoice amounts and reimbursement invoice coverage.")
    parser.add_argument("--manifest", required=True, help="Reimbursement manifest JSON.")
    parser.add_argument("--out", required=True, help="Invoice coverage JSON output path.")
    parser.add_argument("--markdown-out", help="Human-readable gap report. Defaults beside --out as 发票缺口清单.md.")
    parser.add_argument("--update-manifest", action="store_true", help="Write invoice_items and invoice_coverage back to the manifest.")
    return parser.parse_args()


def money(value):
    return Decimal(str(value)).quantize(Decimal("0.01"))


def resolve_path(value, base_dir):
    path = Path(str(value)).expanduser()
    if not path.is_absolute():
        path = base_dir / path
    return path.resolve()


def file_hash(path):
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_pdf_text(path):
    return "\n".join((page.extract_text() or "") for page in PdfReader(str(path)).pages)


def run_apple_vision_ocr(images, script_dir):
    swift = shutil.which("swift")
    if not swift or not images:
        return {}
    result = subprocess.run(
        [swift, str(script_dir / "vision_ocr.swift"), *[str(path) for path in images]],
        check=True,
        text=True,
        capture_output=True,
    )
    data = json.loads(result.stdout)
    return {
        str(Path(item["path"]).resolve()): "\n".join(line.get("text", "") for line in item.get("lines", []))
        for item in data
    }


def parse_candidate(value):
    try:
        parsed = money(str(value).replace(",", ""))
    except (InvalidOperation, ValueError, TypeError):
        return None
    return parsed if parsed > 0 else None


def unique_amounts(values):
    amounts = []
    seen = set()
    for value in values:
        amount = parse_candidate(value)
        if amount is None or amount in seen:
            continue
        seen.add(amount)
        amounts.append(amount)
    return amounts


def extract_invoice_amount(text):
    normalized = str(text or "").replace("，", ",")
    compact = re.sub(r"\s+", "", normalized)
    strong = []
    label_pattern = re.compile(r"价税合计|合计金额|发票金额|小写|Amount", re.IGNORECASE)
    amount_pattern = re.compile(r"(?:CNY|RMB|[¥￥])?([0-9]{1,9}(?:,[0-9]{3})*\.\d{2})(?:[¥￥])?", re.IGNORECASE)
    for label in label_pattern.finditer(compact):
        segment = compact[label.end():label.end() + 160]
        match = amount_pattern.search(segment)
        if match:
            strong.append(match.group(1))
    strong_amounts = unique_amounts(strong)
    if len(strong_amounts) == 1:
        return strong_amounts[0], "label"
    if len(strong_amounts) > 1:
        return None, f"ambiguous labeled amounts: {', '.join(str(value) for value in strong_amounts)}"

    currency_candidates = re.findall(
        r"(?:CNY|RMB|[¥￥])\s*([0-9]{1,9}(?:,[0-9]{3})*\.\d{2})",
        compact,
        flags=re.IGNORECASE,
    )
    currency_amounts = unique_amounts(currency_candidates)
    if len(currency_amounts) == 1:
        return currency_amounts[0], "currency"
    if len(currency_amounts) > 1:
        return None, f"ambiguous currency amounts: {', '.join(str(value) for value in currency_amounts)}"
    return None, "invoice total not recognized"


def extract_invoice_number(text):
    patterns = [
        r"(?:发票号码|发票号)\s*[:：]?\s*([0-9A-Z-]{6,30})",
        r"Invoice\s*(?:No\.?|Number)\s*[:：]?\s*([0-9A-Z-]{6,30})",
    ]
    for pattern in patterns:
        match = re.search(pattern, str(text or ""), flags=re.IGNORECASE)
        if match:
            return match.group(1).upper()
    twenty_digit_numbers = list(dict.fromkeys(re.findall(r"(?<!\d)(\d{20})(?!\d)", str(text or ""))))
    if len(twenty_digit_numbers) == 1:
        return twenty_digit_numbers[0]
    return ""


def extract_invoice_date(text):
    patterns = [
        r"(?:开票日期|发票日期|Date)\s*[:：]?\s*(20\d{2})[年\-/.](\d{1,2})[月\-/.](\d{1,2})日?",
        r"(?:开票日期|发票日期|Date)\s*[:：]?\s*(20\d{2})(\d{2})(\d{2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, str(text or ""), flags=re.IGNORECASE)
        if not match:
            continue
        try:
            return datetime(int(match.group(1)), int(match.group(2)), int(match.group(3))).date().isoformat()
        except ValueError:
            return ""
    return ""


def normalized_entity(value):
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", str(value or "")).lower()


def extract_buyer_name(text):
    source = str(text or "")
    patterns = [
        r"(?:购买方|购方)(?:信息)?[\s\S]{0,80}?名称\s*[:：]\s*([^\n\r]{2,80})",
        r"(?:购买方名称|购方名称)\s*[:：]\s*([^\n\r]{2,80})",
    ]
    for pattern in patterns:
        match = re.search(pattern, source, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip().split("  ", 1)[0][:80]
    return ""


def collect_invoice_items(manifest, base_dir):
    declared = {}
    for raw in manifest.get("invoice_items", []) or []:
        if not isinstance(raw, dict) or not raw.get("path"):
            continue
        path = resolve_path(raw["path"], base_dir)
        item = dict(raw)
        item["path"] = str(path)
        declared[str(path)] = item

    ordered = []
    seen = set()

    by_path = {}

    def add(value, role, voucher_id=""):
        if not value:
            return
        path = resolve_path(value, base_dir)
        key = str(path)
        if key in seen:
            if voucher_id:
                voucher_ids = by_path[key].setdefault("voucher_ids", [])
                if voucher_id not in voucher_ids:
                    voucher_ids.append(voucher_id)
            return
        seen.add(key)
        item = dict(declared.get(key, {}))
        item["path"] = key
        item["role"] = item.get("role") or role
        voucher_ids = list(item.get("voucher_ids", []) or [])
        if voucher_id and voucher_id not in voucher_ids:
            voucher_ids.append(voucher_id)
        item["voucher_ids"] = voucher_ids
        ordered.append(item)
        by_path[key] = item

    for value in manifest.get("invoices", []) or []:
        add(value, declared.get(str(resolve_path(value, base_dir)), {}).get("role", "original"))
    for entry in manifest.get("entries", []) or []:
        for value in entry.get("invoices", []) or []:
            add(value, "original", entry.get("voucher_id", ""))
    for key, item in declared.items():
        if key not in seen:
            add(key, item.get("role", "original"))
    return ordered


def invoice_requirement(manifest, expense_total):
    requirement = dict(manifest.get("invoice_requirement") or {})
    policy = requirement.get("policy") or "full_amount"
    if policy == "none":
        required = Decimal("0.00")
    elif requirement.get("required_amount") not in (None, ""):
        required = money(requirement["required_amount"])
    else:
        required = expense_total
    requirement["policy"] = policy
    requirement["required_amount"] = f"{required:.2f}"
    return requirement, required


def write_markdown(path, report):
    rows = []
    for index, item in enumerate(report["invoice_items"], start=1):
        role = "原始发票" if item.get("role") == "original" else "替代发票"
        amount = item.get("amount") or "待确认"
        filename = Path(item.get("path", "")).name.replace("|", "_")
        vouchers = "、".join(item.get("voucher_ids", []) or []) or "未逐笔关联"
        rows.append(
            f"| {index} | {role} | {filename} | {amount} | {item.get('invoice_number', '未识别') or '未识别'} | "
            f"{item.get('invoice_date', '未识别') or '未识别'} | {item.get('entity_check', '')} | {vouchers} | {item.get('status', '')} |"
        )
    row_text = "\n".join(rows) if rows else "| - | - | 暂无发票 | 0.00 | - | - | - | - | needs_invoices |"
    warning_rows = "\n".join(
        f"- {Path(item.get('path', '')).name}: {item.get('message', '')}"
        for item in report.get("validation_warnings", [])
    ) or "- 无"
    text = f"""# 发票覆盖与缺口清单

项目：{report.get('project_name', '')}
报销人：{report.get('reimburser', '')}
报销总额：{report['expense_total']}
应提供发票金额：{report['required_amount']}
已识别发票金额：{report['recognized_amount']}
尚缺发票金额：{report['missing_amount']}
超额覆盖金额：{report['excess_amount']}
替代发票：{report['replacement_invoice_count']} 张，合计 {report['replacement_invoice_amount']}
字段识别警告：{len(report.get('validation_warnings', []))} 项
当前状态：{report['status']}

## 发票明细

| 序号 | 类型 | 文件 | 识别金额 | 发票号码 | 开票日期 | 购买方核对 | 关联凭证 | 状态 |
|---:|---|---|---:|---|---|---|---|---|
{row_text}

## 字段识别警告

{warning_rows}

## 使用边界

替代发票必须真实、可核验、未在本次报销中重复，并符合公司财务制度。系统只记录其为替代发票，不会把它改写成某笔消费的原始发票。

本报告执行文件指纹、号码重复、金额、日期窗口和购买方文字核对，不等同于税务平台的发票真伪查验。
"""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def main():
    args = parse_args()
    manifest_path = Path(args.manifest).expanduser().resolve()
    output = Path(args.out).expanduser().resolve()
    markdown_output = Path(args.markdown_out).expanduser().resolve() if args.markdown_out else output.parent / "发票缺口清单.md"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    base_dir = manifest_path.parent
    expense_total = sum((money(entry.get("amount", 0)) for entry in manifest.get("entries", []) or []), Decimal("0.00"))
    requirement, required_amount = invoice_requirement(manifest, expense_total)
    items = collect_invoice_items(manifest, base_dir)
    expense_dates = []
    for entry in manifest.get("entries", []) or []:
        try:
            expense_dates.append(datetime.strptime(str(entry.get("date", "")), "%Y-%m-%d").date())
        except ValueError:
            continue
    earliest_invoice_date = min(expense_dates) - timedelta(days=31) if expense_dates else None
    latest_invoice_date = max(expense_dates) + timedelta(days=180) if expense_dates else None
    expected_entity = str((manifest.get("approval_metadata") or {}).get("invoice_entity", "")).strip()

    image_paths = [Path(item["path"]) for item in items if Path(item["path"]).suffix.lower() in IMAGE_EXTS and Path(item["path"]).exists()]
    try:
        image_text = run_apple_vision_ocr(image_paths, Path(__file__).resolve().parent)
    except Exception:
        image_text = {}

    processed = []
    hash_owner = {}
    number_owner = {}
    recognized_amount = Decimal("0.00")
    replacement_amount = Decimal("0.00")
    replacement_count = 0
    unresolved = []
    validation_warnings = []

    for index, raw in enumerate(items, start=1):
        item = dict(raw)
        path = Path(item["path"])
        role = item.get("role", "original")
        item["role"] = role
        item["status"] = "recognized"
        item["notes"] = ""
        item["authenticity_check"] = "not_performed"
        text = ""

        def add_note(message):
            item["notes"] = "; ".join(part for part in [item.get("notes", ""), message] if part)

        if role not in VALID_ROLES:
            item["status"] = "needs_review"
            item["notes"] = f"invalid role: {role}"
        elif not path.exists():
            item["status"] = "needs_review"
            item["notes"] = "invoice file not found"
        else:
            digest = file_hash(path)
            item["sha256"] = digest
            if digest in hash_owner:
                item["status"] = "duplicate"
                item["notes"] = f"same file content as invoice {hash_owner[digest]}"
            else:
                hash_owner[digest] = index

            if path.suffix.lower() in PDF_EXTS:
                try:
                    text = read_pdf_text(path)
                except Exception as exc:
                    item["notes"] = f"PDF text extraction failed: {exc}"
            elif path.suffix.lower() in IMAGE_EXTS:
                text = image_text.get(str(path.resolve()), "")
            else:
                item["status"] = "needs_review"
                item["notes"] = f"unsupported invoice extension: {path.suffix.lower()}"

        invoice_number = extract_invoice_number(text)
        if invoice_number:
            item["invoice_number"] = invoice_number
            item["invoice_number_status"] = "recognized"
            if invoice_number in number_owner and item["status"] == "recognized":
                item["status"] = "duplicate"
                add_note(f"same invoice number as invoice {number_owner[invoice_number]}")
            else:
                number_owner[invoice_number] = index
        else:
            item["invoice_number"] = ""
            item["invoice_number_status"] = "not_detected"
            validation_warnings.append({"index": index, "path": str(path), "field": "invoice_number", "message": "invoice number not detected"})

        invoice_date = extract_invoice_date(text)
        item["invoice_date"] = invoice_date
        if invoice_date:
            parsed_date = datetime.strptime(invoice_date, "%Y-%m-%d").date()
            if earliest_invoice_date and latest_invoice_date and not (earliest_invoice_date <= parsed_date <= latest_invoice_date):
                item["invoice_date_status"] = "outside_expected_window"
                if item["status"] == "recognized":
                    item["status"] = "needs_review"
                add_note(
                    f"invoice date {invoice_date} is outside expected window "
                    f"{earliest_invoice_date.isoformat()} to {latest_invoice_date.isoformat()}"
                )
            else:
                item["invoice_date_status"] = "reasonable"
        else:
            item["invoice_date_status"] = "not_detected"
            validation_warnings.append({"index": index, "path": str(path), "field": "invoice_date", "message": "invoice date not detected"})

        buyer_name = extract_buyer_name(text)
        item["buyer_name"] = buyer_name
        normalized_text = normalized_entity(text)
        if not expected_entity:
            item["entity_check"] = "not_required"
        elif normalized_entity(expected_entity) and normalized_entity(expected_entity) in normalized_text:
            item["entity_check"] = "matched"
        elif buyer_name:
            item["entity_check"] = "mismatch"
            if item["status"] == "recognized":
                item["status"] = "needs_review"
            add_note(f"buyer name {buyer_name!r} does not match expected invoice entity {expected_entity!r}")
        else:
            item["entity_check"] = "not_detected"
            validation_warnings.append({"index": index, "path": str(path), "field": "invoice_entity", "message": "buyer name not detected"})

        provided_amount = item.get("amount")
        if provided_amount not in (None, ""):
            try:
                amount = money(provided_amount)
                source = item.get("amount_source") or "provided"
            except (InvalidOperation, ValueError, TypeError):
                amount = None
                source = "invalid"
        else:
            amount, source = extract_invoice_amount(text)

        if amount is None or amount <= 0:
            if item["status"] == "recognized":
                item["status"] = "needs_review"
            if not item["notes"]:
                item["notes"] = str(source)
            item["amount"] = ""
        else:
            item["amount"] = f"{amount:.2f}"
            item["amount_source"] = source

        if item["status"] == "recognized":
            recognized_amount += amount
            if role == "replacement":
                replacement_count += 1
                replacement_amount += amount
        else:
            unresolved.append({"index": index, "path": str(path), "status": item["status"], "notes": item["notes"]})
        processed.append(item)

    expense_by_amount = {}
    for entry in manifest.get("entries", []) or []:
        try:
            entry_amount = money(entry.get("amount"))
        except (InvalidOperation, ValueError, TypeError):
            continue
        expense_by_amount.setdefault(entry_amount, []).append(entry.get("voucher_id", ""))
    for item in processed:
        if item.get("status") != "recognized" or item.get("role") != "original" or item.get("voucher_ids"):
            continue
        matches = [value for value in expense_by_amount.get(money(item.get("amount")), []) if value]
        if len(matches) == 1:
            item["voucher_ids"] = matches
            item["voucher_match_source"] = "unique_amount"

    missing = max(required_amount - recognized_amount, Decimal("0.00"))
    excess = max(recognized_amount - required_amount, Decimal("0.00"))
    if unresolved:
        status = "needs_invoice_review"
    elif missing > 0:
        status = "needs_invoices"
    else:
        status = "covered"

    coverage = {
        "status": status,
        "policy": requirement["policy"],
        "expense_total": f"{expense_total:.2f}",
        "required_amount": f"{required_amount:.2f}",
        "recognized_amount": f"{recognized_amount:.2f}",
        "missing_amount": f"{missing:.2f}",
        "excess_amount": f"{excess:.2f}",
        "invoice_count": len(processed),
        "recognized_invoice_count": sum(item["status"] == "recognized" for item in processed),
        "replacement_invoice_count": replacement_count,
        "replacement_invoice_amount": f"{replacement_amount:.2f}",
        "validation_warnings": validation_warnings,
        "authenticity_check": "not_performed",
        "unresolved": unresolved,
        "report": str(output),
        "markdown_report": str(markdown_output),
    }
    report = {
        "project_name": manifest.get("project_name", ""),
        "reimburser": manifest.get("reimburser", ""),
        **coverage,
        "invoice_items": processed,
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    write_markdown(markdown_output, report)

    if args.update_manifest:
        manifest["invoice_requirement"] = requirement
        manifest["invoice_items"] = processed
        manifest["invoice_coverage"] = coverage
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
