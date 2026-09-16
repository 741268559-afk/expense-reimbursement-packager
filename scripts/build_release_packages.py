#!/usr/bin/env python3
import argparse
import hashlib
import json
import re
import zipfile
from pathlib import Path


EXCLUDED_PARTS = {".git", ".venv", "__pycache__", "dist", "output", "outputs"}
EXCLUDED_SUFFIXES = {".pyc", ".pyo"}
WORKBUDDY_EXCLUDED_ROOTS = {".github", ".gitignore", "adapters", "agents", "AGENTS.md", "CLAUDE.md", "README.md"}


def parse_args():
    parser = argparse.ArgumentParser(description="Build universal and WorkBuddy Skill ZIP packages.")
    parser.add_argument("--out", required=True, help="Output directory.")
    parser.add_argument("--version", help="Release version. Defaults to VERSION.")
    return parser.parse_args()


def should_include(path, root, workbuddy=False):
    relative = path.relative_to(root)
    if any(part in EXCLUDED_PARTS for part in relative.parts) or path.suffix.lower() in EXCLUDED_SUFFIXES:
        return False
    if workbuddy and relative.parts and relative.parts[0] in WORKBUDDY_EXCLUDED_ROOTS:
        return False
    return True


def skill_body(skill_text):
    match = re.match(r"^---\n.*?\n---\n", skill_text, flags=re.DOTALL)
    if not match:
        raise SystemExit("SKILL.md does not contain valid YAML frontmatter.")
    return skill_text[match.end():]


def workbuddy_skill(skill_text, version):
    frontmatter = f"""---
name: expense-reimbursement-packager
display_name: 费用报销材料整理
display_name_en: Expense Reimbursement Packager
description: 整理支付截图、发票、审批信息和费用报销单，生成可打印、可核对的财务报销材料包。
description_zh: 从支付截图、发票、审批信息和 Excel 模板生成经过金额与附件校验的报销材料。
description_en: Build and verify finance-ready reimbursement packages from payment evidence, invoices, approval data, and Excel forms.
category: data-analysis
version: {version}
author: Expense Reimbursement Packager contributors
---
"""
    return frontmatter + skill_body(skill_text)


def manifest_for_version(text, version):
    return re.sub(r"(?m)^version:\s*.*$", f"version: {version}", text, count=1)


def build_zip(root, output, version, workbuddy=False):
    files = [
        path
        for path in sorted(root.rglob("*"))
        if path.is_file() and should_include(path, root, workbuddy=workbuddy)
    ]
    prefix = root.name
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(prefix + "/", b"")
        for path in files:
            relative = path.relative_to(root)
            arcname = str(Path(prefix) / relative).replace("\\", "/")
            if relative.as_posix() == "SKILL.md" and workbuddy:
                data = workbuddy_skill(path.read_text(encoding="utf-8"), version).encode("utf-8")
                archive.writestr(arcname, data)
            elif relative.as_posix() == "manifest.yaml":
                data = manifest_for_version(path.read_text(encoding="utf-8"), version).encode("utf-8")
                archive.writestr(arcname, data)
            else:
                archive.write(path, arcname)
    with zipfile.ZipFile(output) as archive:
        bad = archive.testzip()
        if bad:
            raise SystemExit(f"ZIP integrity failure in {output}: {bad}")
    return {
        "path": str(output),
        "files": len(files),
        "bytes": output.stat().st_size,
        "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
    }


def main():
    args = parse_args()
    root = Path(__file__).resolve().parent.parent
    version = args.version or (root / "VERSION").read_text(encoding="utf-8").strip()
    if not re.fullmatch(r"\d+\.\d+\.\d+", version):
        raise SystemExit(f"Invalid semantic version: {version}")
    output_dir = Path(args.out).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    universal = output_dir / f"expense-reimbursement-packager-universal-v{version}.zip"
    workbuddy = output_dir / f"expense-reimbursement-packager-workbuddy-v{version}.zip"
    result = {
        "status": "passed",
        "version": version,
        "universal": build_zip(root, universal, version, workbuddy=False),
        "workbuddy": build_zip(root, workbuddy, version, workbuddy=True),
    }
    checksum_file = output_dir / f"expense-reimbursement-packager-v{version}-SHA256SUMS.txt"
    checksum_file.write_text(
        f"{result['universal']['sha256']}  {universal.name}\n"
        f"{result['workbuddy']['sha256']}  {workbuddy.name}\n",
        encoding="ascii",
    )
    result["checksums"] = str(checksum_file)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    from runtime_utils import configure_utf8_stdio

    configure_utf8_stdio()
    main()
