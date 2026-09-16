#!/usr/bin/env python3
import argparse
import json
import os
import subprocess
import sys
import venv
from pathlib import Path


def parse_args():
    parser = argparse.ArgumentParser(description="Create an isolated runtime for the reimbursement packager.")
    parser.add_argument("--venv", help="Virtual environment folder. Defaults to .venv inside the Skill directory.")
    parser.add_argument("--requirements", help="Requirements file. Defaults to the repository requirements.txt.")
    return parser.parse_args()


def venv_python(folder):
    return folder / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def main():
    args = parse_args()
    root = Path(__file__).resolve().parent.parent
    environment = Path(args.venv).expanduser().resolve() if args.venv else root / ".venv"
    requirements = Path(args.requirements).expanduser().resolve() if args.requirements else root / "requirements.txt"
    if not requirements.is_file():
        raise SystemExit(f"Requirements file not found: {requirements}")
    if not venv_python(environment).is_file():
        venv.EnvBuilder(with_pip=True).create(environment)
    python = venv_python(environment)
    subprocess.run(
        [str(python), "-m", "pip", "install", "--disable-pip-version-check", "-r", str(requirements)],
        check=True,
    )
    result = {
        "status": "ready",
        "python": str(python),
        "venv": str(environment),
        "requirements": str(requirements),
        "next_step": f"{python} {root / 'scripts' / 'check_readiness.py'} --out readiness.json",
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
