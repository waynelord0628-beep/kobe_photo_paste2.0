from __future__ import annotations

import base64
import io
import zipfile
from pathlib import Path

from docx import Document
from PIL import Image, ImageDraw, ImageFont

from template_export import EXPORT_MODES, export_word


ROOT = Path(__file__).resolve().parent
WORKSPACE = ROOT.parents[1]
OUT_DIR = WORKSPACE / "outputs" / "kobe_template_modes"


def make_photo(index: int, orientation: str = "portrait") -> dict:
    colors = [
        (42, 93, 151),
        (190, 82, 65),
        (51, 138, 97),
        (224, 170, 55),
        (113, 78, 154),
        (45, 139, 165),
    ]
    size = (1200, 800) if orientation == "landscape" else (800, 1200)
    image = Image.new("RGB", size, colors[(index - 1) % len(colors)])
    draw = ImageDraw.Draw(image)
    try:
        font = ImageFont.truetype("arial.ttf", 82)
        small = ImageFont.truetype("arial.ttf", 38)
    except OSError:
        font = ImageFont.load_default()
        small = ImageFont.load_default()
    if orientation == "landscape":
        draw.rectangle((36, 36, 1164, 764), outline=(255, 255, 255), width=8)
        draw.text((76, 110), f"橫式照片 {index}", fill=(255, 255, 255), font=font)
        draw.text((80, 260), f"Kobe Word 模板測試 {index}", fill=(245, 245, 245), font=small)
    else:
        draw.rectangle((36, 36, 764, 1164), outline=(255, 255, 255), width=8)
        draw.text((76, 120), f"直式照片 {index}", fill=(255, 255, 255), font=font)
        draw.text((80, 270), f"Kobe Word 模板測試 {index}", fill=(245, 245, 245), font=small)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    data = base64.b64encode(buffer.getvalue()).decode("ascii")
    return {
        "url": f"data:image/png;base64,{data}",
        "location": f"測試地點 {index}",
        "description": f"{'橫式' if orientation == 'landscape' else '直式'}照片排版範例 {index}",
    }


def inspect_docx(path: Path) -> str:
    document = Document(str(path))
    with zipfile.ZipFile(path) as archive:
        media = len([name for name in archive.namelist() if name.startswith("word/media/")])
    return f"tables={len(document.tables)} media={media} sections={len(document.sections)}"


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    portrait_photos = [make_photo(i, "portrait") for i in range(1, 7)]
    landscape_photos = [make_photo(i, "landscape") for i in range(1, 7)]
    for mode, label in EXPORT_MODES.items():
        photos = landscape_photos if mode == "two_landscape" else portrait_photos
        test_label = "橫式測試" if mode == "two_landscape" else "直式測試"
        path = export_word(
            template_dir=ROOT / "templates",
            output_dir=OUT_DIR,
            mode=mode,
            title=f"Kobe_{label}_{test_label}",
            photos=photos,
        )
        print(f"{label}: {path} | {inspect_docx(path)}")


if __name__ == "__main__":
    main()
