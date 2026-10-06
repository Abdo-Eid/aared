"""Read word-arabic-fallback.pdf (one sample per page, in word-arabic-fallback.order.tsv
order), take each page's real text width, and compare fonts against Word's own Arial.

Also prints which font the PDF says drew the text, so a font with no Arabic glyphs
(Cambria, Calibri) shows what Word substituted.
"""
import collections
import math
import os

import fitz

here = os.path.dirname(os.path.abspath(__file__))
order = [line.rstrip("\n").split("\t", 1) for line in open(f"{here}/word-arabic-fallback.order.tsv", encoding="utf-8-sig") if line.strip()]

pages = []  # (width, fonts)
doc = fitz.open(f"{here}/word-arabic-fallback.pdf")
for page in doc:
    spans = [s for b in page.get_text("dict")["blocks"] for l in b.get("lines", []) for s in l["spans"] if s["text"].strip()]
    if not spans:
        pages.append((0.0, collections.Counter()))
        continue
    # A page holds exactly one single-line paragraph: its width is the extent of all its spans.
    x0 = min(s["bbox"][0] for s in spans)
    x1 = max(s["bbox"][2] for s in spans)
    pages.append((x1 - x0, collections.Counter(s["font"] for s in spans)))
print(f"{len(order)} samples expected, {len(pages)} pages found")
if len(pages) != len(order):
    raise SystemExit("page count mismatch: a sample wrapped or vanished; refusing to guess the mapping")

by_font = collections.OrderedDict()
drawn = collections.defaultdict(collections.Counter)
for (font, _), (width, fonts) in zip(order, pages):
    by_font.setdefault(font, []).append(width)
    drawn[font].update(fonts)

ref = by_font["Arial"]
print(f"\n{'font':20} {'ratio vs Arial':>15} {'sd':>8}   font(s) the PDF says drew the text")
for font, widths in by_font.items():
    mean = sum(widths) / sum(ref)
    ratios = [w / r for w, r in zip(widths, ref)]
    sd = math.sqrt(sum((x - mean) ** 2 for x in ratios) / len(ratios))
    names = ", ".join(f"{n}" for n, _ in drawn[font].most_common(2))
    print(f"{font:20} {mean:15.4f} {sd:8.4f}   {names}")
