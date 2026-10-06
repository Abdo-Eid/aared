"""Emit compact advance-width tables (per mille of the em) for the faces
Office charts are most often set in.

Latin: widths come from metric-compatible clones (Carlito = Calibri,
Liberation = Arial / Times New Roman), whose advance widths are identical to
the originals by design.

Arabic: joined letters change width with their joining form, so the table is
per letter and form (isolated / initial / medial / final), read from the
glyphs a font draws for the Arabic Presentation Forms-B (and -A) code points,
which are the shaped glyphs themselves. It comes from the real Windows faces
(Arial, Tahoma; Times New Roman's Arabic glyphs are Arial's, byte for byte in
advance) because the Liberation clones carry no Arabic shaping.

  python gen_metrics.py > ../../packages/aared/src/core/chart/metrics-data.ts
  python gen_metrics.py --arabic-only        (just the Arabic exports)

`joined_width` below is the algorithm of `mixedEm` in
packages/aared/src/core/chart/metrics.ts for text inside an Arabic run (it always
measures spaces and punctuation in the Arabic face); arabic-width.py imports it
to score the engine's estimate against PowerPoint's own PDF.
"""
import os
import sys
import unicodedata

from fontTools.ttLib import TTFont

B = "/usr/share/fonts/truetype/"
W = os.environ.get("WINDOWS_FONTS", "C:/Windows/Fonts/")


def pick(*paths):
  for p in paths:
    if os.path.exists(p):
      return p
  return paths[0]


LATIN_FACES = {
  "Calibri": (pick(B + "crosextra/Carlito-Regular.ttf", W + "calibri.ttf"), pick(B + "crosextra/Carlito-Bold.ttf", W + "calibrib.ttf")),
  "Arial": (pick(B + "liberation/LiberationSans-Regular.ttf", W + "arial.ttf"), pick(B + "liberation/LiberationSans-Bold.ttf", W + "arialbd.ttf")),
  "Times New Roman": (pick(B + "liberation/LiberationSerif-Regular.ttf", W + "times.ttf"), pick(B + "liberation/LiberationSerif-Bold.ttf", W + "timesbd.ttf")),
}
# Faces whose Arabic glyphs are measured. Times New Roman's letters come out identical to Arial's
# (only its space and punctuation advances differ); a face that is not listed measures as Arial.
ARABIC_FACES = {"arial": W + "arial.ttf", "times new roman": W + "times.ttf", "tahoma": W + "tahoma.ttf"}

CODES = list(range(0x20, 0x7f)) + list(range(0xa0, 0x100))
D = "0123456789abcdefghijklmnopqrstuvwxyz"


def enc(v):
  v = max(0, min(1295, round(v)))
  return D[v // 36] + D[v % 36]


def face(path):
  f = TTFont(path)
  em = f["head"].unitsPerEm
  cmap = f.getBestCmap()
  hmtx = f["hmtx"]
  avg = sum(hmtx[cmap[c]][0] for c in range(0x61, 0x7b)) / 26
  out = ""
  for c in CODES:
    g = cmap.get(c)
    out += enc((hmtx[g][0] if g else avg) * 1000 / em)
  os2 = f["OS/2"]
  return out, round(os2.usWinAscent * 1000 / em), round(os2.usWinDescent * 1000 / em)


# ---------------------------------------------------------------- Arabic

# Letters the table covers: the Arabic block's letters plus the Persian and
# Urdu ones a chart is likely to carry. Anything else Arabic falls to one
# fallback advance in the engine.
LETTERS = (
  list(range(0x0621, 0x063B)) + list(range(0x0641, 0x064B)) +
  [0x0671, 0x0679, 0x067E, 0x0686, 0x0688, 0x0691, 0x0698, 0x06A4, 0x06A9, 0x06AF, 0x06BA, 0x06BE, 0x06C1, 0x06CC, 0x06D2]
)
# Joining-neutral signs and punctuation that still have an advance; tatweel
# joins on both sides, so it takes the dual-joining type with one width.
SIGNS = [0x060C, 0x061B, 0x061F, 0x066A, 0x066B, 0x066C]
TATWEEL = 0x0640
# Per mille of the em: the advance of an Arabic-script character outside the tables (ARABIC_FALLBACK in metrics.ts).
ARABIC_FALLBACK = 450
# ASCII spaces and punctuation, as inclusive ranges. Beside Arabic letters they are set in the Arabic face
# (a space inside an Arabic run is that face's space), so the face's own advances are tabulated.
NEUTRAL_RANGES = [(0x20, 0x2F), (0x3A, 0x40), (0x5B, 0x60), (0x7B, 0x7E)]
NEUTRALS = [c for a, b in NEUTRAL_RANGES for c in range(a, b + 1)]
LAM = 0x0644
ALEFS = [0x0622, 0x0623, 0x0625, 0x0627]
FORMS = ["isolated", "initial", "medial", "final"]


def presentation_forms():
  """base code point -> {form: presentation code point}, from Unicode's own decompositions."""
  forms = {}
  for pf in list(range(0xFB50, 0xFE00)) + list(range(0xFE70, 0xFEFD)):
    parts = unicodedata.decomposition(chr(pf)).split()
    if len(parts) == 2 and parts[0][1:-1] in FORMS:
      forms.setdefault(int(parts[1], 16), {})[parts[0][1:-1]] = pf
  return forms


def joining_type(base, forms):
  f = forms.get(base, {})
  if "initial" in f and "medial" in f:
    return "d"
  if "final" in f:
    return "r"
  return "n"


def marks():
  """Ranges (inclusive) of the combining marks in U+0600-06FF: they take no advance and never break a join."""
  out, start = [], None
  for c in range(0x0600, 0x0701):
    mark = c < 0x0700 and unicodedata.category(chr(c)) == "Mn"
    if mark and start is None:
      start = c
    if not mark and start is not None:
      out.append((start, c - 1))
      start = None
  return out


MARK_SET = {c for a, b in marks() for c in range(a, b + 1)}


def arabic_face(path, forms):
  """Advances (per mille) for one font: {letters: {code: [iso, init, med, fin]}, lamalef: [...], digits: (a, b), signs}."""
  f = TTFont(path)
  em = f["head"].unitsPerEm
  cmap = f.getBestCmap()
  hmtx = f["hmtx"]

  def adv(code):
    g = cmap.get(code)
    return None if g is None else hmtx[g][0] * 1000 / em

  def form_adv(base, form):
    pf = forms.get(base, {}).get(form)
    v = adv(pf) if pf is not None else None
    if v is None:
      v = adv(base)  # the font lacks the shaped glyph: the plain letter's advance stands in
    if v is None:
      raise SystemExit(f"{path}: no glyph for U+{base:04X} {form}")
    return v

  letters = {}
  for base in LETTERS:
    kind = joining_type(base, forms)
    iso = form_adv(base, "isolated")
    row = [iso, iso, iso, iso]
    if kind in ("d", "r"):
      row[3] = form_adv(base, "final")
    if kind == "d":
      row[1] = form_adv(base, "initial")
      row[2] = form_adv(base, "medial")
    letters[base] = (kind, row)
  tat = adv(TATWEEL)
  letters[TATWEEL] = ("d", [tat] * 4)
  for s in SIGNS:
    letters[s] = ("n", [adv(s)] * 4)
  # Lam + alef is one glyph: FEF5/FEF7/FEF9/FEFB isolated, the next code point its final form.
  lamalef = []
  for i in range(4):
    lamalef += [adv(0xFEF5 + 2 * i), adv(0xFEF6 + 2 * i)]
  digits = (sum(adv(0x0660 + i) for i in range(10)) / 10, sum(adv(0x06F0 + i) for i in range(10)) / 10)
  neutral = {c: adv(c) for c in NEUTRALS}
  return dict(letters=letters, lamalef=lamalef, digits=digits, neutral=neutral)


def letter_order():
  forms = presentation_forms()
  codes = sorted(set(LETTERS + SIGNS + [TATWEEL]))
  return forms, codes, [("d" if c == TATWEEL else "n" if c in SIGNS else joining_type(c, forms)) for c in codes]


def encode_face(data, codes):
  out = ""
  for c in codes:
    kind, row = data["letters"][c]
    n = {"d": 4, "r": 2, "n": 1}[kind]
    # d: isolated initial medial final; r: isolated final; n: isolated
    vals = row if n == 4 else [row[0], row[3]] if n == 2 else [row[0]]
    out += "".join(enc(v) for v in vals)
  out += "".join(enc(v) for v in data["lamalef"])
  out += "".join(enc(v) for v in data["digits"])
  out += "".join(enc(data["neutral"][c]) for c in NEUTRALS)
  return out


def joined_width(text, data):
  """Width in em of Arabic `text` in one face: the algorithm of mixedEm() in metrics.ts, inside an Arabic run."""
  letters = data["letters"]
  cps =[ord(ch) for ch in text]

  def kind(c):
    if c in MARK_SET:
      return "t"
    if c == 0x200D:
      return "d"
    if c in letters:
      return letters[c][0]
    return "n"

  total = 0.0
  i = 0
  n = len(cps)
  while i < n:
    c = cps[i]
    if c in MARK_SET or c in (0x200C, 0x200D):
      i += 1
      continue
    if 0x0660 <= c <= 0x0669:
      total += data["digits"][0]
      i += 1
      continue
    if 0x06F0 <= c <= 0x06F9:
      total += data["digits"][1]
      i += 1
      continue
    if c in data["neutral"]:
      total += data["neutral"][c]  # a space or punctuation mark inside an Arabic run is the Arabic face's own
      i += 1
      continue
    if c not in letters:
      # An Arabic-script character the tables lack takes the engine's fallback advance; anything else is not Arabic.
      if 0x0600 <= c <= 0x06FF or 0x0750 <= c <= 0x077F or 0x08A0 <= c <= 0x08FF or 0xFB50 <= c <= 0xFDFF or 0xFE70 <= c <= 0xFEFF:
        total += ARABIC_FALLBACK
      i += 1
      continue
    k = letters[c][0]
    p = i - 1
    while p >= 0 and kind(cps[p]) == "t":
      p -= 1
    q = i + 1
    while q < n and kind(cps[q]) == "t":
      q += 1
    prev_joins = p >= 0 and kind(cps[p]) == "d"
    if c == LAM and q < n and cps[q] in ALEFS:
      total += data["lamalef"][ALEFS.index(cps[q]) * 2 + (1 if prev_joins else 0)]
      i = q + 1
      continue
    next_joins = q < n and kind(cps[q]) in ("d", "r")
    row = letters[c][1]
    if k == "d":
      total += row[2] if prev_joins and next_joins else row[3] if prev_joins else row[1] if next_joins else row[0]
    elif k == "r":
      total += row[3] if prev_joins else row[0]
    else:
      total += row[0]
    i += 1
  return total / 1000


def arabic_exports():
  forms, codes, kinds = letter_order()
  # Letters: three characters each, two base-36 digits of (code - 0x0600), then the joining type.
  letters = "".join(enc(c - 0x0600) + k for c, k in zip(codes, kinds))
  tables, seen = {}, {}
  for name, path in ARABIC_FACES.items():
    if not os.path.exists(path):
      print(f"skipping {name}: {path} not found", file=sys.stderr)
      continue
    data = arabic_face(path, forms)
    s = encode_face(data, codes)
    if s in seen:
      print(f"{name}: Arabic advances identical to {seen[s]}, not emitted", file=sys.stderr)
      continue
    seen[s] = name
    tables[name] = s
  lines = ["/** Arabic letters the width tables cover: per letter, two base-36 digits of (code - 0x0600) then its joining type (d dual, r right, n none). */",
           f'export const ARABIC_LETTERS = "{letters}";',
           "/**",
           " * Arabic advance widths per mille of the em, two base-36 digits each, per face. Per letter of ARABIC_LETTERS in order:",
           " * d gives isolated, initial, medial, final; r gives isolated, final; n gives isolated. Then lam-alef (alef, alef with",
           " * madda, hamza above, hamza below; isolated and final each), then the mean Arabic-Indic digit and the mean extended one,",
           " * then the face's own advance for each character of ARABIC_NEUTRALS (spaces and ASCII punctuation inside an Arabic run).",
           " * Times New Roman's Arabic letters are Arial's, byte for byte; only its neutrals differ.",
           " */",
           "export const ARABIC_TABLES: Record<string, string> = {"]
  lines += [f'  "{name}": "{s}",' for name, s in tables.items()]
  lines.append("};")
  ranges = ", ".join(f"0x{a:04x}, 0x{b:04x}" for a, b in marks())
  neutrals = ", ".join(f"0x{a:02x}, 0x{b:02x}" for a, b in NEUTRAL_RANGES)
  lines += ["/** The characters whose Arabic-face advance each table ends with, as inclusive [first, last] pairs of code points, in order. */",
            f"export const ARABIC_NEUTRALS: readonly number[] = [{neutrals}];","/** Combining marks in U+0600-06FF as inclusive [first, last] pairs: no advance, and a join passes through them. */",
            f"export const ARABIC_MARKS: readonly number[] = [{ranges}];"]
  return lines


def latin_exports():
  lines = []
  for name, (reg, bold) in LATIN_FACES.items():
    r, asc, desc = face(reg)
    b, _, _ = face(bold)
    lines.append(f'  "{name.lower()}": {{ ascent: {asc}, descent: {desc}, regular: "{r}", bold: "{b}" }},')
  return ["/** Advance widths, per mille of the em, two base-36 digits per code point: U+0020-007E then U+00A0-00FF. */",
          "export const FACE_TABLES: Record<string, { ascent: number; descent: number; regular: string; bold: string }> = {",
          *lines, "};"]


if __name__ == "__main__":
  sys.stdout.reconfigure(newline="\n")
  args = sys.argv[1:]
  if "--arabic-only" in args:
    print("\n".join(arabic_exports()))
  else:
    print("/* Generated by tools/chart-oracle/gen_metrics.py - do not edit. */")
    print("\n".join(latin_exports()))
    print("\n".join(arabic_exports()))
