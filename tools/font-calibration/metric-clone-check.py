"""One-off: confirm Arimo/Tinos/Cousine (Google Fonts) share Arial/Times New
Roman/Courier New's advance widths, the same way Liberation does — so the
substitution table fix (Google Fonts hosts Arimo, not Liberation) still keeps
metric compatibility. Downloads each woff2, converts with fonttools, and
diffs hmtx against the local system font. Network only here, not in tests.
"""
import glob
import io
import subprocess
import sys
import tempfile
import urllib.request

from fontTools.ttLib import TTFont

PAIRS = [
    ("Arimo", "Arial"),
    ("Tinos", "Times New Roman"),
    ("Cousine", "Courier New"),
]

SYSTEM_NAME_TO_FILE = {}
for path in glob.glob(r"C:\Windows\Fonts\*.ttf") + glob.glob(r"C:\Windows\Fonts\*.ttc"):
    try:
        f = TTFont(path, fontNumber=0, lazy=True)
        name = f["name"].getDebugName(1) or ""
        SYSTEM_NAME_TO_FILE.setdefault(name, path)
    except Exception:
        pass


def google_woff2_url(family: str) -> str:
    req = urllib.request.Request(
        f"https://fonts.googleapis.com/css2?family={family.replace(' ', '+')}:ital,wght@0,400&display=swap",
        headers={"User-Agent": "Mozilla/5.0"},
    )
    css = urllib.request.urlopen(req).read().decode()
    for line in css.splitlines():
        if "url(" in line:
            return line.split("url(")[1].split(")")[0]
    raise RuntimeError(f"no url in css for {family}")


def widths(font: TTFont) -> dict:
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]
    upm = font["head"].unitsPerEm
    out = {}
    for cp in range(0x20, 0x7F):
        glyph = cmap.get(cp)
        if glyph is None:
            continue
        out[chr(cp)] = hmtx[glyph][0] / upm
    return out


for google_name, system_name in PAIRS:
    system_path = SYSTEM_NAME_TO_FILE.get(system_name)
    if system_path is None:
        print(system_name, "-> not installed locally, skip")
        continue
    url = google_woff2_url(google_name)
    data = urllib.request.urlopen(url).read()
    tmp_woff2 = tempfile.mktemp(suffix=".woff2")
    tmp_ttf = tempfile.mktemp(suffix=".ttf")
    with open(tmp_woff2, "wb") as fh:
        fh.write(data)
    font = TTFont(tmp_woff2)
    font.flavor = None
    font.save(tmp_ttf)

    google_font = TTFont(tmp_ttf, lazy=True)
    system_font = TTFont(system_path, fontNumber=0, lazy=True)
    a = widths(google_font)
    b = widths(system_font)
    diffs = {ch: (a[ch], b[ch]) for ch in a if ch in b and abs(a[ch] - b[ch]) > 1e-4}
    print(f"{google_name} vs {system_name}: {len(a)} glyphs compared, {len(diffs)} differ")
    for ch, (x, y) in list(diffs.items())[:5]:
        print("  ", repr(ch), x, "vs", y)
