#!/usr/bin/env python3
import json
import os
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path


HEIC_EXTS = {".heic", ".heif"}
OCR_MODES = ("auto", "apple-vision", "tesseract", "none")


class OCRUnavailable(RuntimeError):
    pass


def find_tesseract():
    configured = os.environ.get("TESSERACT_CMD", "").strip()
    candidates = [Path(configured).expanduser()] if configured else []
    discovered = shutil.which("tesseract")
    if discovered:
        candidates.append(Path(discovered))
    if os.name == "nt":
        for env_name in ("ProgramFiles", "ProgramW6432", "ProgramFiles(x86)", "LOCALAPPDATA"):
            base = os.environ.get(env_name)
            if not base:
                continue
            root = Path(base)
            candidates.extend([
                root / "Tesseract-OCR" / "tesseract.exe",
                root / "Programs" / "Tesseract-OCR" / "tesseract.exe",
            ])
    seen = set()
    for candidate in candidates:
        key = str(candidate)
        if key in seen:
            continue
        seen.add(key)
        if candidate.is_file():
            return candidate.resolve()
    return None


def apple_vision_available(script_dir):
    return (
        platform.system() == "Darwin"
        and shutil.which("swift") is not None
        and (Path(script_dir) / "vision_ocr.swift").is_file()
    )


def tesseract_languages(executable):
    if not executable:
        return []
    result = subprocess.run(
        [str(executable), "--list-langs"],
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    if result.returncode != 0:
        return []
    return [
        line.strip()
        for line in result.stdout.splitlines()
        if line.strip() and not line.lower().startswith("list of available languages")
    ]


def choose_tesseract_language(executable):
    configured = os.environ.get("TESSERACT_LANG", "").strip()
    if configured:
        return configured
    languages = set(tesseract_languages(executable))
    if {"chi_sim", "eng"}.issubset(languages):
        return "chi_sim+eng"
    if "chi_sim" in languages:
        return "chi_sim"
    if "eng" in languages:
        return "eng"
    return next(iter(sorted(languages)), "eng")


def backend_status(script_dir):
    tesseract = find_tesseract()
    apple = apple_vision_available(script_dir)
    automatic = "apple-vision" if apple else "tesseract" if tesseract else "none"
    return {
        "platform": platform.system(),
        "apple_vision": apple,
        "tesseract": str(tesseract) if tesseract else "",
        "tesseract_languages": tesseract_languages(tesseract),
        "auto_backend": automatic,
    }


def run_apple_vision_ocr(images, script_dir):
    if not apple_vision_available(script_dir):
        raise OCRUnavailable("Apple Vision OCR requires macOS, Swift, and vision_ocr.swift.")
    result = subprocess.run(
        [shutil.which("swift"), str(Path(script_dir) / "vision_ocr.swift"), *[str(path) for path in images]],
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
    )
    if result.returncode != 0:
        raise OCRUnavailable(result.stderr.strip() or "Apple Vision OCR failed.")
    data = json.loads(result.stdout)
    return {Path(item["path"]).resolve(): item for item in data}


def prepare_tesseract_image(path, temp_dir):
    path = Path(path)
    if path.suffix.lower() not in HEIC_EXTS:
        return path
    try:
        from pillow_heif import register_heif_opener
        from PIL import Image
    except ImportError as exc:
        raise OCRUnavailable("HEIC OCR requires pillow-heif. Install requirements.txt.") from exc
    register_heif_opener()
    output = Path(temp_dir) / f"{path.stem}.png"
    with Image.open(path) as image:
        image.convert("RGB").save(output, "PNG")
    return output


def run_tesseract_ocr(images):
    executable = find_tesseract()
    if not executable:
        raise OCRUnavailable(
            "Tesseract OCR was not found. Add it to PATH or set TESSERACT_CMD to the executable."
        )
    language = choose_tesseract_language(executable)
    page_segmentation = os.environ.get("TESSERACT_PSM", "6").strip() or "6"
    output = {}
    with tempfile.TemporaryDirectory(prefix="expense_ocr_") as temp_dir:
        for path in images:
            source = Path(path).resolve()
            prepared = prepare_tesseract_image(source, temp_dir)
            result = subprocess.run(
                [str(executable), str(prepared), "stdout", "-l", language, "--psm", page_segmentation],
                text=True,
                encoding="utf-8",
                errors="replace",
                capture_output=True,
            )
            if result.returncode != 0:
                raise OCRUnavailable(
                    f"Tesseract OCR failed for {source.name}: {result.stderr.strip()}"
                )
            lines = [line.strip() for line in result.stdout.splitlines() if line.strip()]
            output[source] = {
                "path": str(source),
                "backend": "tesseract",
                "language": language,
                "lines": [{"text": line} for line in lines],
            }
    return output


def run_ocr_images(images, mode, script_dir):
    images = [Path(path).resolve() for path in images]
    if not images or mode == "none":
        return {}
    if mode not in OCR_MODES:
        raise ValueError(f"Unsupported OCR mode: {mode}")

    if mode == "apple-vision":
        return run_apple_vision_ocr(images, script_dir)
    if mode == "tesseract":
        return run_tesseract_ocr(images)

    attempts = []
    if apple_vision_available(script_dir):
        attempts.append(("apple-vision", lambda: run_apple_vision_ocr(images, script_dir)))
    if find_tesseract():
        attempts.append(("tesseract", lambda: run_tesseract_ocr(images)))
    if not attempts:
        raise OCRUnavailable(
            "No OCR backend is available. Install Swift on macOS or Tesseract on Windows/macOS, "
            "or use --ocr none and review the generated workbook manually."
        )

    errors = []
    for name, runner in attempts:
        try:
            return runner()
        except Exception as exc:
            errors.append(f"{name}: {exc}")
    raise OCRUnavailable("; ".join(errors))
