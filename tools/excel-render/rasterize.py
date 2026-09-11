"""Rasterize one sheet (or part) PDF to PNG with PyMuPDF (MuPDF engine).

Usage: python rasterize.py <input.pdf> <output.png> <dpi>

Every export fits exactly one A4 page, so a multi-page PDF means the
fit-to-page assumption broke and fails loud instead of silently pairing
the wrong page. Prints {"width","height"} for the caller.
"""

import json
import sys

import fitz  # PyMuPDF


def main() -> None:
    pdf_path, png_path, dpi = sys.argv[1], sys.argv[2], float(sys.argv[3])
    document = fitz.open(pdf_path)
    if document.page_count != 1:
        raise SystemExit(
            f"aared excel rasterize: expected 1 page (fit-to-page), "
            f"got {document.page_count} in {pdf_path}"
        )
    pixmap = document[0].get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
    pixmap.save(png_path)
    print(json.dumps({"width": pixmap.width, "height": pixmap.height}))


if __name__ == "__main__":
    main()
