"""Score the engine against PowerPoint's PDF of the oracle deck.

  python tools/chart-oracle/score.py      (after export-office.ps1, render.ts, measure.ts)

Arguments default to out/chart-oracle.pdf, cases.json, out/engine.json,
out/measured.json and out/score.json; pass them to score another run.
Needs PyMuPDF (the same Python the Excel review route uses).

Per case:
  words   Office words (PyMuPDF) vs the engine's words as the browser laid
          them out (measure.ts); matched by string, nearest centre first.
  axis    the numeric labels in the left- or right-most column of each side,
          compared as sets: does the value axis carry the same numbers?
  plot    the inner plot box from gridline extents on both sides.
"""
import json, os, re, sys
from collections import defaultdict
import fitz

PT = 72.0
NUM = re.compile(r"^-?[\$]?[\d,]*\.?\d+%?$")


def office_case(page, f):
    x0, y0, w, h = f["x"] * PT, f["y"] * PT, f["w"] * PT, f["h"] * PT
    inside = lambda b: b[0] >= x0 - 2 and b[2] <= x0 + w + 2 and b[1] >= y0 - 2 and b[3] <= y0 + h + 2
    words = [dict(text=t, x0=a - x0, y0=b - y0, x1=c - x0, y1=d - y0) for a, b, c, d, t, *_ in page.get_text("words") if inside((a, b, c, d))]
    hs, vs = [], []
    for d in page.get_drawings():
        r = d["rect"]
        if not inside((r.x0, r.y0, r.x1, r.y1)) or d.get("color") is None:
            continue
        for item in d["items"]:
            if item[0] != "l":
                continue
            p, q = item[1], item[2]
            if abs(p.y - q.y) < 0.6 and abs(q.x - p.x) > w * 0.3:
                hs.append((min(p.x, q.x) - x0, p.y - y0, max(p.x, q.x) - x0))
            elif abs(p.x - q.x) < 0.6 and abs(q.y - p.y) > h * 0.3:
                vs.append((p.x - x0, min(p.y, q.y) - y0, max(p.y, q.y) - y0))
    return words, grid_box(hs, vs)


def grid_box(hs, vs):
    if not hs and not vs:
        return None
    xs0 = [a for a, _, _ in hs] + [x for x, _, _ in vs]
    xs1 = [c for _, _, c in hs] + [x for x, _, _ in vs]
    ys0 = [y for _, y, _ in hs] + [a for _, a, _ in vs]
    ys1 = [y for _, y, _ in hs] + [b for _, _, b in vs]
    return [min(xs0), min(ys0), max(xs1), max(ys1)]


SEG = re.compile(r"M(-?[\d.]+) (-?[\d.]+)([HV])(-?[\d.]+)")


def engine_grid(entry, w, h):
    hs, vs = [], []
    for p in entry["primitives"]:
        if p["kind"] != "path" or not p.get("stroke") or p.get("fill"):
            continue
        for x, y, k, e in SEG.findall(p["path"]):
            x, y, e = float(x), float(y), float(e)
            if k == "H" and abs(e - x) > w * 0.3:
                hs.append((min(x, e), y, max(x, e)))
            elif k == "V" and abs(e - y) > h * 0.3:
                vs.append((x, min(y, e), max(y, e)))
    return grid_box(hs, vs)


def centre(t):
    return ((t["x0"] + t["x1"]) / 2, (t["y0"] + t["y1"]) / 2)


def match(ref, eng):
    pool = defaultdict(list)
    for t in eng:
        pool[t["text"]].append(t)
    errs, missing = [], []
    for r in ref:
        cands = pool.get(r["text"])
        if not cands:
            missing.append(r["text"])
            continue
        rc = centre(r)
        best = min(cands, key=lambda e: abs(centre(e)[0] - rc[0]) + abs(centre(e)[1] - rc[1]))
        cands.remove(best)
        ec = centre(best)
        errs.append((ec[0] - rc[0], ec[1] - rc[1]))
    extra = [t["text"] for v in pool.values() for t in v]
    return errs, missing, extra


def axis_numbers(words, w):
    """Numbers in the outermost left and right label columns (the value axes)."""
    nums = [t for t in words if NUM.match(t["text"])]
    left = sorted({t["text"] for t in nums if t["x1"] < w * 0.2})
    right = sorted({t["text"] for t in nums if t["x0"] > w * 0.8})
    return left, right


def main(pdf_path, cases_path, engine_path, measured_path, out_path):
    meta = json.load(open(cases_path, encoding="utf-8"))
    engine = json.load(open(engine_path, encoding="utf-8"))
    measured = json.load(open(measured_path, encoding="utf-8"))
    doc = fitz.open(pdf_path)
    rows = []
    for i, (case, frame) in enumerate(zip(meta["cases"], meta["frames"])):
        # Cases authored after the last Office export have no slide, no engine drawing and no PDF page yet.
        if i >= len(engine) or i >= len(doc) or i >= len(measured):
            continue
        if engine[i] is None:
            rows.append(dict(id=case["id"], drawn=False))
            continue
        w, h = frame["w"] * PT, frame["h"] * PT
        ref_words, ref_plot = office_case(doc[i], frame)
        eng_words = measured[i]
        errs, missing, extra = match(ref_words, eng_words)
        eng_plot = engine_grid(engine[i], w, h)
        plot = None if ref_plot is None or eng_plot is None else [round(e - r, 1) for r, e in zip(ref_plot, eng_plot)]
        ra, ea = axis_numbers(ref_words, w), axis_numbers(eng_words, w)
        rows.append(dict(
            id=case["id"], script=case.get("script", "latin"), drawn=True, words=len(ref_words), matched=len(errs),
            dx=round(sum(abs(a) for a, _ in errs) / max(len(errs), 1), 1),
            dy=round(sum(abs(b) for _, b in errs) / max(len(errs), 1), 1),
            worst=round(max([max(abs(a), abs(b)) for a, b in errs] or [0]), 1),
            plot=plot, ref_plot=ref_plot and [round(v, 1) for v in ref_plot], eng_plot=eng_plot and [round(v, 1) for v in eng_plot],
            ref_words=ref_words, eng_words=eng_words, axis_ok=ra == ea, axis_office=ra, axis_engine=ea,
            missing=missing[:6], extra=extra[:6]))
    json.dump(rows, open(out_path, "w", encoding="utf-8"), indent=1, ensure_ascii=False)
    drawn = [r for r in rows if r["drawn"]]
    print(f"{'case':24} {'match':>9} {'dx':>5} {'dy':>5} {'worst':>6}  {'plot (dl,dt,dr,db)':24} axis")
    for r in drawn:
        ax = "ok" if r["axis_ok"] else f"office {r['axis_office']} engine {r['axis_engine']}"
        print(f"{r['id']:24} {r['matched']:>4}/{r['words']:<4} {r['dx']:5} {r['dy']:5} {r['worst']:6}  {str(r['plot']):24} {ax}")
    # Latin and Arabic cases are scored apart: the Arabic ones measure joined-letter widths and the complex-script face.
    for script in ("latin", "arabic"):
        group = [r for r in drawn if r["script"] == script]
        if not group:
            continue
        n = len(group)
        big = [r for r in group if r["plot"] and max(abs(v) for v in r["plot"]) > 6]
        print(f"\n{script:6} {n} cases | mean dx {sum(r['dx'] for r in group) / n:.2f} dy {sum(r['dy'] for r in group) / n:.2f} pt"
              f" | words matched {sum(r['matched'] for r in group)}/{sum(r['words'] for r in group)}"
              f" | value axis same {sum(r['axis_ok'] for r in group)}/{n}"
              f" | plot box off >6pt {len(big)}/{n}")


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    out = os.path.abspath(os.path.join(here, "..", "out", "charts"))
    defaults = [os.path.join(out, "chart-oracle.pdf"), os.path.join(here, "cases.json"), os.path.join(out, "engine.json"),
                os.path.join(out, "measured.json"), os.path.join(out, "score.json")]
    main(*(sys.argv[1:6] + defaults[len(sys.argv) - 1:]))
