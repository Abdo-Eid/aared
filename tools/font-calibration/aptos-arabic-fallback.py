"""Read aptos-arabic-fallback.{excel,word}.pdf: for every labelled case, which font the
PDF says drew the Arabic, its size, and its total width against the same app's Arial
case. Spans are assigned to the label on their row (nearest vertical centre): PyMuPDF
splits one row into several lines and blocks.
"""
import collections
import os
import re

import fitz

here = os.path.dirname(os.path.abspath(__file__))
ARABIC = re.compile(r"[؀-ۿﭐ-﷿ﹰ-﻿]")

for app in ("excel", "word"):
    path = f"{here}/aptos-arabic-fallback.{app}.pdf"
    if not os.path.exists(path):
        print(f"== {app}: no PDF")
        continue
    fonts = collections.defaultdict(collections.Counter)
    sizes = collections.defaultdict(set)
    widths = collections.defaultdict(float)
    for page in fitz.open(path):
        spans = [s for b in page.get_text("dict")["blocks"] for l in b.get("lines", []) for s in l["spans"] if s["text"].strip()]
        labels = [s for s in spans if "|" in s["text"]]
        rows = collections.defaultdict(list)
        for s in spans:
            if "|" in s["text"] or not labels:
                continue
            cy = (s["bbox"][1] + s["bbox"][3]) / 2
            label = min(labels, key=lambda l: abs((l["bbox"][1] + l["bbox"][3]) / 2 - cy))
            rows[label["text"].strip()].append(s)
        for label, row in rows.items():
            key = "|".join(label.split("|")[1:]).strip()
            arabic = [s for s in row if ARABIC.search(s["text"])]
            for s in arabic:
                fonts[key][s["font"]] += len(s["text"])
                sizes[key].add(round(s["size"], 2))
            if arabic:
                widths[key] += max(s["bbox"][2] for s in row) - min(s["bbox"][0] for s in row)
    print(f"== {app}")
    for key in fonts:
        mode = key.rsplit("|", 1)[0] if "|" in key else ""
        base = f"{mode}|Arial" if mode else "Arial"
        ratio = widths[key] / widths[base] if widths.get(base) else float("nan")
        print(f"  {key:28} width vs Arial {ratio:6.3f}   sizes {sorted(sizes[key])}   drawn in {dict(fonts[key].most_common(3))}")
