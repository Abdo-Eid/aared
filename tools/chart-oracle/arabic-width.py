"""Score the engine's joined-letter Arabic widths against PowerPoint's PDF of the Arabic cases.

The evidence behind the Arabic rule in packages/aared/src/core/chart/metrics.ts:
PowerPoint sets Arabic chart text in Arial (its fallback when the chart's Latin
face has no Arabic) with every letter in its joining form, so a joined run is
about half the width of the isolated letters, and a word's width depends on
its letters (Oct 2026: "al-mabi'at", 8 letters, 31.4pt at 12pt = 2.617 em, the
engine says 2.633 em). The estimate here is `joined_width` from gen_metrics.py,
the same rule as the engine's, read straight from the font file.

  python tools/chart-oracle/arabic-width.py      (after export-office.ps1)

Scores every case marked `"script": "arabic"` in cases.json, apart from the
Latin ones: per text span, Office's width against the estimate, and the mean
ratio (1.0 = the rule matches Office).
"""
import json
import os

import fitz

import gen_metrics

HERE = os.path.dirname(os.path.abspath(__file__))
meta = json.load(open(os.path.join(HERE, "cases.json"), encoding="utf-8"))
doc = fitz.open(os.path.join(HERE, "out", "chart-oracle.pdf"))

forms, _, _ = gen_metrics.letter_order()
FACE = gen_metrics.arabic_face(gen_metrics.ARABIC_FACES["arial"], forms)


def is_arabic(ch: str) -> bool:
    return 0x0600 <= ord(ch) <= 0x06FF


def estimate_em(text: str) -> float:
    """Width in em of a span of Arabic words in Arial: joined letters, and the Arabic face's own spaces."""
    return gen_metrics.joined_width(text, FACE)


rows = []
for index, (case, frame) in enumerate(zip(meta["cases"], meta["frames"])):
    # Cases authored after the last Office export have no PDF page yet.
    if case.get("script") != "arabic" or index >= len(doc):
        continue
    page = doc[index]
    x0 = frame["x"] * 72
    for block in page.get_text("dict")["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                text = span["text"].strip()
                a, _, c, _ = span["bbox"]
                # Pure Arabic spans only: a span with Latin letters or digits mixes two faces.
                if not text or not all(is_arabic(ch) or ch == " " for ch in text):
                    continue
                if not (x0 - 2 <= a and c <= x0 + frame["w"] * 72 + 2):
                    continue
                size = span["size"]
                rows.append((case["id"], text, span["font"], round(size, 2), round(c - a, 2), round(estimate_em(text) * size, 2)))

for row in rows:
    print(row)
if rows:
    ratio = sum(r[5] for r in rows) / sum(r[4] for r in rows)
    print(f"engine / Office width over {len(rows)} spans: {ratio:.3f}")
