"""Rasterize Office-exported PDFs to PNG with PyMuPDF (MuPDF engine).

Usage:
    python rasterize.py --mode sheet <input.pdf> <output.png> <dpi>
    python rasterize.py --mode document <input.pdf> <output-dir> <dpi>

sheet: one sheet (or print-area part) PDF to one PNG. Every export fits
    exactly one A4 page, so a multi-page PDF means the fit-to-page
    assumption broke and fails loud instead of silently pairing the wrong
    page. Prints {"width","height"} for the caller.

document: Word's exported PDF to one PNG per page, plus exact text
    geometry. Writes word-page-N.png for every page, word-spans.json for
    every text span and word-words.json for every token box, in CSS pixels
    at 96 dpi (the unit the bench and the renderer work in), so layout
    positions can be compared with Word instead of guessed from a
    screenshot. Prints {"pages": [{"width","height"}, ...]}.
"""

import json
import sys
from pathlib import Path

import fitz  # PyMuPDF

CSS_PX_PER_PT = 96 / 72


def rasterize_sheet(pdf_path: str, png_path: str, dpi: float) -> None:
    document = fitz.open(pdf_path)
    if document.page_count != 1:
        raise SystemExit(
            f"aared excel rasterize: expected 1 page (fit-to-page), "
            f"got {document.page_count} in {pdf_path}"
        )
    pixmap = document[0].get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
    pixmap.save(png_path)
    print(json.dumps({"width": pixmap.width, "height": pixmap.height}))


def rasterize_document(pdf_path: str, output_dir: str, dpi: float) -> None:
    out = Path(output_dir)
    document = fitz.open(pdf_path)
    pages = []
    spans = []
    words = []
    for number, page in enumerate(document, start=1):
        pixmap = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
        pixmap.save(str(out / f"word-page-{number}.png"))
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
        for x0, y0, x1, y1, text, block_no, line_no, word_no in page.get_text("words"):
            words.append({
                "page": number,
                "text": text,
                "x0": x0 * CSS_PX_PER_PT,
                "y0": y0 * CSS_PX_PER_PT,
                "x1": x1 * CSS_PX_PER_PT,
                "y1": y1 * CSS_PX_PER_PT,
                "block": block_no,
                "line": line_no,
                "word": word_no,
            })
    (out / "word-spans.json").write_text(json.dumps(spans, ensure_ascii=False), encoding="utf-8")
    (out / "word-words.json").write_text(json.dumps(words, ensure_ascii=False), encoding="utf-8")
    print(json.dumps({"pages": pages}))


def main() -> None:
    if len(sys.argv) >= 3 and sys.argv[1] == "--mode":
        mode, rest = sys.argv[2], sys.argv[3:]
    else:
        raise SystemExit(
            "aared rasterize: want --mode <sheet|document> "
            "(sheet: <input.pdf> <output.png> <dpi>; "
            "document: <input.pdf> <output-dir> <dpi>)"
        )
    if mode == "sheet" and len(rest) == 3:
        rasterize_sheet(rest[0], rest[1], float(rest[2]))
    elif mode == "document" and len(rest) == 3:
        rasterize_document(rest[0], rest[1], float(rest[2]))
    else:
        raise SystemExit(f"aared rasterize: unknown mode '{mode}' (want sheet|document)")


if __name__ == "__main__":
    main()
