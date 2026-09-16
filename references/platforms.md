# Platform And Agent Compatibility

## Supported Hosts

The core workflow is ordinary Python and does not depend on Codex APIs. It can be called by any Agent that can read `SKILL.md`, access local files, and execute terminal commands.

- Codex and Agent Skills-compatible hosts: install the universal folder or ZIP in the host's Skill directory.
- WorkBuddy personal: upload the WorkBuddy ZIP through Add Skill > Upload Skill.
- WorkBuddy Enterprise: upload the WorkBuddy ZIP in Skill Management. The package includes `manifest.yaml`.
- Claude Code and other terminal Agents: point the Agent at `SKILL.md`, or install the folder in the host's local skills directory when supported.

## Python

Python 3.10 or newer is required.

Create an isolated environment on either platform:

```text
python scripts/bootstrap.py
```

The command prints JSON containing the virtual-environment Python path. Use that path to run subsequent scripts.

Typical macOS command:

```bash
./.venv/bin/python scripts/check_readiness.py --out readiness.json
```

Typical Windows PowerShell command:

```powershell
.\.venv\Scripts\python.exe scripts\check_readiness.py --out readiness.json
```

## OCR

`--ocr auto` selects the first available local backend:

1. Apple Vision on macOS when Swift is available.
2. Tesseract on Windows or macOS.
3. If neither is available, the run stops using OCR and relies on filename parsing plus the editable Excel review workbook.

Tesseract environment variables:

- `TESSERACT_CMD`: full executable path when Tesseract is not on `PATH`.
- `TESSERACT_LANG`: language expression, normally `chi_sim+eng`.
- `TESSERACT_PSM`: page segmentation mode, default `6`.

Install examples:

```bash
brew install tesseract tesseract-lang
```

```powershell
winget install --id UB-Mannheim.TesseractOCR
$env:TESSERACT_CMD = "C:\Program Files\Tesseract-OCR\tesseract.exe"
$env:TESSERACT_LANG = "chi_sim+eng"
```

OCR is always a draft. Amounts, dates, currencies, refunds, and categories must still pass human review and package verification.

## Configuration Locations

Explicit command arguments and environment variables have priority. Automatic config discovery also checks:

- Windows: `%APPDATA%\expense-reimbursement-packager\expense-reimbursement\`
- macOS: `~/Library/Application Support/expense-reimbursement-packager/expense-reimbursement/`
- Linux: `${XDG_CONFIG_HOME:-~/.config}/expense-reimbursement-packager/expense-reimbursement/`
- Legacy Codex location: `${CODEX_HOME:-~/.codex}/expense-reimbursement/`

Set `EXPENSE_REIMBURSEMENT_HOME` to use the same custom configuration root with any Agent.

## File Compatibility

- Windows-invalid filename characters are sanitized in generated outputs.
- All subprocesses use argument arrays rather than shell command strings.
- Chinese paths and filenames are handled as Unicode.
- HEIC/HEIF conversion uses `pillow-heif` on Windows and macOS, with macOS `sips` as a fallback.
- ZIP packages use standard UTF-8 filenames.
