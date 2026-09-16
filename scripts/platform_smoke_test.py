#!/usr/bin/env python3
import json
import platform
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pillow_heif import from_pillow

from build_reimbursement_pack import convert_heic_image, safe_name
from ocr_utils import backend_status, run_ocr_images
from package_from_folder import default_config_dirs


def test_font():
    candidates = [
        Path("C:/Windows/Fonts/arial.ttf"),
        Path("/System/Library/Fonts/Supplemental/Arial.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]
    for path in candidates:
        if path.is_file():
            return ImageFont.truetype(str(path), 72)
    return ImageFont.load_default()


def main():
    script_dir = Path(__file__).resolve().parent
    checks = []
    with tempfile.TemporaryDirectory(prefix="expense_platform_smoke_") as temp:
        root = Path(temp)
        image = Image.new("RGB", (1200, 400), "white")
        ImageDraw.Draw(image).text((50, 120), "AMOUNT 12.30", font=test_font(), fill="black")
        png = root / "receipt.png"
        image.save(png)
        heic = root / "receipt.heic"
        from_pillow(image).save(heic)
        converted = root / "converted"
        converted.mkdir()
        converted_path = convert_heic_image(heic, {"converted": converted})
        if not converted_path.is_file() or converted_path.suffix.lower() != ".png":
            raise SystemExit("HEIC conversion did not produce PNG output.")
        checks.append({"name": "heic_conversion", "status": "passed"})

        ocr = backend_status(script_dir)
        if ocr["tesseract"]:
            result = run_ocr_images([png], "tesseract", script_dir)
            text = "\n".join(line["text"] for line in result[png.resolve()]["lines"])
            if not text.strip():
                raise SystemExit("Tesseract returned no text for the smoke-test image.")
            checks.append({"name": "tesseract_ocr", "status": "passed", "text": text})
        else:
            checks.append({"name": "tesseract_ocr", "status": "skipped"})

    sanitized = safe_name('project:*?"<>|/\\name')
    if any(character in sanitized for character in '\\/:*?"<>|'):
        raise SystemExit(f"Windows filename sanitization failed: {sanitized}")
    checks.append({"name": "windows_filename_sanitization", "status": "passed", "value": sanitized})

    result = {
        "status": "passed",
        "platform": platform.system(),
        "python": sys.version.split()[0],
        "config_dirs": [str(path) for path in default_config_dirs()],
        "checks": checks,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    from runtime_utils import configure_utf8_stdio

    configure_utf8_stdio()
    main()
