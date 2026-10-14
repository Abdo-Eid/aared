"""Crop two review captures to their content and place them side by side.

Usage:
    python pair-images.py <excel.png> <aared.png> <output.png> <excel-label> <aared-label>

Each capture is trimmed to its non-white content (plus a small margin), so
the A4 page's empty space is dropped. Both are scaled to the shorter height,
preserving aspect ratio, then placed side by side with a label above each.
Prints {"width","height"} for the caller.

Needs Pillow (`pip install Pillow`), in addition to the PyMuPDF the
rasterizer uses.
"""

import json
import sys

from PIL import Image, ImageChops, ImageDraw, ImageFont

MARGIN = 16
LABEL_HEIGHT = 30
GAP = 12
OUTER_MARGIN = 12
WHITE = (255, 255, 255)


def trimmed(path: str) -> Image.Image:
    image = Image.open(path).convert("RGB")
    background = Image.new("RGB", image.size, WHITE)
    box = ImageChops.difference(image, background).getbbox()
    if box is None:
        raise SystemExit(f"aared pair: {path} has no content to keep")
    left, top, right, bottom = box
    return image.crop((
        max(left - MARGIN, 0),
        max(top - MARGIN, 0),
        min(right + MARGIN, image.width),
        min(bottom + MARGIN, image.height),
    ))


def labelled(image: Image.Image, label: str, height: int, font: ImageFont.ImageFont | ImageFont.FreeTypeFont) -> Image.Image:
    if image.height != height:
        width = round(image.width * height / image.height)
        image = image.resize((width, height), Image.Resampling.LANCZOS)
    bar = Image.new("RGB", (image.width, LABEL_HEIGHT), (34, 34, 34))
    ImageDraw.Draw(bar).text((10, 6), label, fill=WHITE, font=font)
    out = Image.new("RGB", (image.width, LABEL_HEIGHT + image.height), WHITE)
    out.paste(bar, (0, 0))
    out.paste(image, (0, LABEL_HEIGHT))
    return out


def main() -> None:
    if len(sys.argv) != 6:
        raise SystemExit("Usage: python pair-images.py <excel.png> <aared.png> <output.png> <excel-label> <aared-label>")
    excel_path, aared_path, output, excel_label, aared_label = sys.argv[1:]
    excel = trimmed(excel_path)
    aared = trimmed(aared_path)
    height = min(excel.height, aared.height)
    font = ImageFont.load_default(size=16)
    first = labelled(excel, excel_label, height, font)
    second = labelled(aared, aared_label, height, font)
    canvas = Image.new("RGB", (first.width + GAP + second.width, LABEL_HEIGHT + height + 2 * OUTER_MARGIN), WHITE)
    canvas.paste(first, (OUTER_MARGIN, OUTER_MARGIN))
    canvas.paste(second, (OUTER_MARGIN + first.width + GAP, OUTER_MARGIN))
    canvas.save(output)
    print(json.dumps({"width": canvas.width, "height": canvas.height}))


if __name__ == "__main__":
    main()
