"""Rasterize Word's exported PDF to one PNG per page and dump exact text geometry.

Usage: python rasterize.py <input.pdf> <output-dir> <dpi>

Writes word-page-N.png for every page and word-spans.json: every text span with
its font, size and position in CSS pixels at 96 dpi (the unit the bench and the
renderer work in), so a layout number can be read from Word instead of guessed
from a screenshot. Prints {"pages": [{"width","height"}, ...]} for the caller.
"""

import json
import sys
from pathlib import Path

import fitz  # PyMuPDF

CSS_PX_PER_PT = 96 / 72


def main() -> None:
    pdf_path, output_dir, dpi = sys.argv[1], Path(sys.argv[2]), float(sys.argv[3])
    document = fitz.open(pdf_path)
    pages = []
    spans = []
    for number, page in enumerate(document, start=1):
        pixmap = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
        pixmap.save(str(output_dir / f"word-page-{number}.png"))
        pages.append({"width": pixmap.width, "height": pixmap.height})
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                for span in line["spans"]:
                    x0, y0, x1, y1 = (value * CSS_PX_PER_PT for value in span["bbox"])
                    spans.append({
                        "page": number,
                        "text": span["text"],
                        "font": span["font"],
                        "sizePx": span["size"] * CSS_PX_PER_PT,
                        "x0": x0, "y0": y0, "x1": x1, "y1": y1,
                        "baselinePx": span["origin"][1] * CSS_PX_PER_PT,
                    })
    (output_dir / "word-spans.json").write_text(json.dumps(spans, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({"pages": pages}))


if __name__ == "__main__":
    main()
