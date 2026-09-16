#!/usr/bin/env python3
import argparse
import json
import re
from decimal import Decimal, InvalidOperation
from pathlib import Path


DEFAULT_CATEGORIES = ["交通费", "餐饮费", "设备费", "场地费", "演员费", "其他费用"]


def parse_args():
    parser = argparse.ArgumentParser(description="Assign voucher IDs and write the pre-DingTalk expense confirmation and exception reports.")
    parser.add_argument("--manifest", required=True, help="Reimbursement manifest JSON.")
    parser.add_argument("--out-dir", required=True, help="Folder for human-readable workflow reports.")
    parser.add_argument("--update-manifest", action="store_true", help="Write voucher IDs, exceptions, and report paths back to the manifest.")
    return parser.parse_args()


def money(value):
    return Decimal(str(value)).quantize(Decimal("0.01"))


def sort_key(entry):
    return (
        entry.get("date") or "9999-12-31",
        entry.get("time") or "23:59:59",
        entry.get("screenshot") or entry.get("source") or "",
    )


def ensure_voucher_ids(manifest):
    entries = list(manifest.get("entries", []) or [])
    entries.sort(key=sort_key)
    existing = [str(entry.get("voucher_id", "")).strip() for entry in entries]
    valid_existing = (
        len(existing) == len(set(existing))
        and all(re.fullmatch(r"FY\d{3,}", value) for value in existing)
    )
    if not valid_existing:
        for index, entry in enumerate(entries, start=1):
            entry["voucher_id"] = f"FY{index:03d}"
    manifest["entries"] = entries
    return entries


def amount_candidates(text):
    values = []
    for match in re.finditer(r"(?:CNY|RMB|[¥￥$])\s*([0-9]{1,8}(?:,[0-9]{3})*\.\d{2})", text, flags=re.IGNORECASE):
        try:
            value = money(match.group(1).replace(",", ""))
        except (InvalidOperation, ValueError):
            continue
        if value not in values:
            values.append(value)
    return values


def detected_exceptions(entry):
    text = " ".join([
        str(entry.get("ocr_text", "")),
        Path(str(entry.get("screenshot", ""))).name,
        str(entry.get("notes", "")),
    ])
    lowered = text.lower()
    issues = []

    def add(code, severity, message):
        issues.append({"code": code, "severity": severity, "message": message})

    if re.search(r"已退款|退款成功|退款中|交易关闭|已撤销|refund", lowered, flags=re.IGNORECASE):
        add("refund_or_reversal", "blocking", "支付凭证可能包含退款、撤销或关闭交易，需确认实际可报销金额。")
    if re.search(r"个人消费|私人|自用|\bAA\b|拼单", text, flags=re.IGNORECASE):
        add("personal_or_split_expense", "blocking", "支付凭证可能包含个人消费、AA 或拼单，需说明可报销范围。")
    if re.search(r"(?:USD|HKD|EUR|JPY|美元|港币|欧元|日元)|(?<![A-Z])\$\s*\d", text, flags=re.IGNORECASE):
        add("foreign_currency", "blocking", "检测到外币标记，需确认币种、汇率和人民币报销金额。")
    candidates = amount_candidates(text)
    if len(candidates) > 1:
        add("multiple_amounts", "blocking", f"凭证中识别到多个金额：{', '.join(f'{value:.2f}' for value in candidates)}，需确认实际支付金额。")
    if re.search(r"优惠券|补贴|折扣|红包|立减", text):
        add("discount_present", "warning", "凭证包含优惠或补贴信息，已按实际支付金额报销，请复核。")
    return issues


def refresh_exceptions(manifest):
    exceptions = []
    issue_index = 1
    for entry in manifest.get("entries", []) or []:
        resolution = str(entry.get("exception_resolution", "")).strip()
        detected = detected_exceptions(entry)
        entry["detected_exception_codes"] = [item["code"] for item in detected]
        for item in detected:
            exceptions.append({
                "exception_id": f"EX{issue_index:03d}",
                "voucher_id": entry.get("voucher_id", ""),
                **item,
                "resolved": bool(resolution),
                "resolution": resolution,
            })
            issue_index += 1
    manifest["expense_exceptions"] = exceptions
    return exceptions


def invoice_status(manifest, entry):
    if entry.get("invoices"):
        return "已逐笔关联"
    policy = (manifest.get("invoice_requirement") or {}).get("policy", "full_amount")
    coverage = manifest.get("invoice_coverage") or {}
    if policy == "none":
        return "无需发票"
    if coverage.get("status") == "covered":
        return "总额已覆盖"
    if coverage.get("status") == "needs_invoices":
        return "发票不足"
    if coverage.get("status") == "needs_invoice_review":
        return "发票待复核"
    return "待核对"


def confirmation_status(manifest, total):
    confirmation = manifest.get("expense_confirmation") or {}
    try:
        confirmed_amount = money(confirmation.get("confirmed_amount"))
    except (InvalidOperation, ValueError, TypeError):
        confirmed_amount = None
    approval = manifest.get("approval_metadata") or {}
    if approval.get("dingding_number") and confirmed_amount == total:
        return "已通过钉钉金额确认"
    return "待确认后提交钉钉"


def write_confirmation(path, manifest, exceptions):
    entries = manifest.get("entries", []) or []
    total = sum((money(entry.get("amount", 0)) for entry in entries), Decimal("0.00"))
    rows = []
    for entry in entries:
        screenshot_status = "已提供" if entry.get("screenshot") else "缺少"
        rows.append(
            f"| {entry.get('voucher_id', '')} | {entry.get('date', '')} | {entry.get('purpose', '')} | "
            f"{entry.get('expense_type', '')} | {money(entry.get('amount', 0)):.2f} | {screenshot_status} | {invoice_status(manifest, entry)} |"
        )
    category_totals = {category: Decimal("0.00") for category in manifest.get("categories", []) or DEFAULT_CATEGORIES}
    for entry in entries:
        category = entry.get("expense_type", "")
        category_totals.setdefault(category, Decimal("0.00"))
        category_totals[category] += money(entry.get("amount", 0))
    category_lines = "\n".join(
        f"- {category}: {amount:.2f} 元"
        for category, amount in category_totals.items()
        if amount != Decimal("0.00")
    ) or "- 无"
    unresolved = [item for item in exceptions if item.get("severity") == "blocking" and not item.get("resolved")]
    text = f"""# 报销金额确认单

项目：{manifest.get('project_name', '')}
报销人：{manifest.get('reimburser', '')}
费用笔数：{len(entries)}
报销总额：{total:.2f} 元
确认状态：{confirmation_status(manifest, total)}
未解决异常：{len(unresolved)} 项

## 费用明细

| 凭证编号 | 日期 | 用途 | 费用类别 | 金额（元） | 支付凭证 | 发票状态 |
|---|---|---|---|---:|---|---|
{chr(10).join(rows)}

## 分类小计

{category_lines}

## 确认规则

- 报销金额以支付截图中的实际支付金额为准。
- 确认总额后再发起钉钉，钉钉审批金额必须与本单总额完全一致。
- 异常清单中存在未解决的阻断项时，不生成最终报销材料。
"""
    path.write_text(text, encoding="utf-8")


def write_exceptions(path, manifest, exceptions):
    rows = []
    for item in exceptions:
        status = "已说明" if item.get("resolved") else ("需处理" if item.get("severity") == "blocking" else "请复核")
        rows.append(
            f"| {item['exception_id']} | {item.get('voucher_id', '')} | {item.get('severity', '')} | "
            f"{item.get('message', '').replace('|', '_')} | {status} | {item.get('resolution', '').replace('|', '_')} |"
        )
    row_text = "\n".join(rows) if rows else "| - | - | - | 未发现异常 | 通过 | - |"
    unresolved = [item for item in exceptions if item.get("severity") == "blocking" and not item.get("resolved")]
    text = f"""# 报销异常清单

项目：{manifest.get('project_name', '')}
阻断异常：{len(unresolved)} 项
当前状态：{'需处理后继续' if unresolved else '可继续'}

| 异常编号 | 凭证编号 | 级别 | 问题 | 状态 | 处理说明 |
|---|---|---|---|---|---|
{row_text}

检测范围包括退款或撤销、个人或拆分消费、外币、多金额以及优惠补贴。自动检测用于提示复核，不替代原始凭证判断。
"""
    path.write_text(text, encoding="utf-8")


def main():
    args = parse_args()
    manifest_path = Path(args.manifest).expanduser().resolve()
    out_dir = Path(args.out_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    ensure_voucher_ids(manifest)
    exceptions = refresh_exceptions(manifest)
    confirmation_path = out_dir / "报销金额确认单.md"
    exceptions_path = out_dir / "报销异常清单.md"
    write_confirmation(confirmation_path, manifest, exceptions)
    write_exceptions(exceptions_path, manifest, exceptions)
    manifest["workflow_reports"] = {
        **(manifest.get("workflow_reports") or {}),
        "expense_confirmation": str(confirmation_path),
        "expense_exceptions": str(exceptions_path),
    }
    if args.update_manifest:
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    unresolved = [item for item in exceptions if item.get("severity") == "blocking" and not item.get("resolved")]
    result = {
        "status": "needs_review" if unresolved else "ready",
        "manifest": str(manifest_path),
        "confirmation_report": str(confirmation_path),
        "exception_report": str(exceptions_path),
        "voucher_count": len(manifest.get("entries", []) or []),
        "unresolved_exception_count": len(unresolved),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    from runtime_utils import configure_utf8_stdio

    configure_utf8_stdio()
    main()
