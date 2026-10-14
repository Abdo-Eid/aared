"""
Build large search-benchmark fixtures by replicating parts of the real Office
fixtures in the repository's fixtures/ (slides, sheet rows, document body). The output is
generated, untracked (`out/`), and only meant for timing `search()`.

    python tools/benchmarks/search/make-fixtures.py [--slides 300] [--sheets 9] [--sheet-copies 60] [--docx-copies 40]
"""
import argparse
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[3]
FIXTURES = ROOT / "fixtures"
OUT = ROOT / "tools" / "out" / "search-fixtures"


def read_zip(path):
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist() if not i.is_dir()}


def write_zip(path, parts):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml first, as Office writes it.
        for name in sorted(parts, key=lambda n: (n != "[Content_Types].xml", n)):
            z.writestr(name, parts[name])
    print(f"wrote {path.relative_to(ROOT)} ({path.stat().st_size / 1e6:.1f} MB)")


def make_pptx(slides):
    parts = read_zip(FIXTURES / "post-production-review.pptx")
    sources = sorted(int(m.group(1)) for n in parts if (m := re.fullmatch(r"ppt/slides/slide(\d+)\.xml", n)))
    pres = parts["ppt/presentation.xml"].decode()
    rels = parts["ppt/_rels/presentation.xml.rels"].decode()
    types = parts["[Content_Types].xml"].decode()
    next_slide = max(sources) + 1
    next_id = max(int(i) for i in re.findall(r'<p:sldId id="(\d+)"', pres)) + 1
    new_ids, new_rels, new_types = [], [], []
    for k in range(slides - len(sources)):
        src = sources[k % len(sources)]
        num = next_slide + k
        parts[f"ppt/slides/slide{num}.xml"] = parts[f"ppt/slides/slide{src}.xml"]
        parts[f"ppt/slides/_rels/slide{num}.xml.rels"] = parts[f"ppt/slides/_rels/slide{src}.xml.rels"]
        rid = f"rId{1000 + k}"
        new_rels.append(
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{num}.xml"/>'
        )
        new_ids.append(f'<p:sldId id="{next_id + k}" r:id="{rid}"/>')
        new_types.append(
            f'<Override PartName="/ppt/slides/slide{num}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        )
    pres = pres.replace("</p:sldIdLst>", "".join(new_ids) + "</p:sldIdLst>")
    parts["ppt/presentation.xml"] = pres.encode()
    parts["ppt/_rels/presentation.xml.rels"] = rels.replace("</Relationships>", "".join(new_rels) + "</Relationships>").encode()
    parts["[Content_Types].xml"] = types.replace("</Types>", "".join(new_types) + "</Types>").encode()
    write_zip(OUT / f"large-{slides}-slides.pptx", parts)


# Parts that reference rows/rels of the original sheet; dropped from the copies.
STRIP_BLOCKS = re.compile(r"<(mergeCells|conditionalFormatting|dataValidations|hyperlinks|tableParts)\b[^>]*>.*?</\1>", re.S)
STRIP_EMPTY = re.compile(r"<(legacyDrawing|drawing|autoFilter)\b[^>]*/>")


def strip_row_bound_parts(xml):
    return STRIP_EMPTY.sub("", STRIP_BLOCKS.sub("", xml))


def make_xlsx(extra_sheets, copies):
    parts = read_zip(FIXTURES / "xlsx-coverage.xlsx")
    source = parts["xl/worksheets/sheet3.xml"].decode()
    head, rest = source.split("<sheetData>", 1)
    data, tail = rest.split("</sheetData>", 1)
    rows = re.findall(r"<row .*?</row>", data, re.S)
    height = max(int(re.search(r'<row r="(\d+)"', r).group(1)) for r in rows)

    def replicate(times):
        out = []
        for c in range(times):
            shift = c * height
            for row in rows:
                row = re.sub(r'<row r="(\d+)"', lambda m: f'<row r="{int(m.group(1)) + shift}"', row)
                row = re.sub(r' r="([A-Z]+)(\d+)"', lambda m: f' r="{m.group(1)}{int(m.group(2)) + shift}"', row)
                out.append(row)
        return "".join(out)

    body = replicate(copies)
    head = re.sub(r'<dimension ref="[A-Z]+\d+:([A-Z]+)\d+"/>', lambda m: f'<dimension ref="A1:{m.group(1)}{height * copies}"/>', head)
    sheet = head + "<sheetData>" + body + "</sheetData>" + strip_row_bound_parts(tail)
    workbook = parts["xl/workbook.xml"].decode()
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    types = parts["[Content_Types].xml"].decode()
    sheets, new_rels, new_types = [], [], []
    for k in range(extra_sheets):
        num = 100 + k
        parts[f"xl/worksheets/sheet{num}.xml"] = sheet.encode()
        rid = f"rId{1000 + k}"
        sheets.append(f'<sheet name="Bulk {k + 1}" sheetId="{100 + k}" r:id="{rid}"/>')
        new_rels.append(
            f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{num}.xml"/>'
        )
        new_types.append(
            f'<Override PartName="/xl/worksheets/sheet{num}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    parts["xl/workbook.xml"] = workbook.replace("</sheets>", "".join(sheets) + "</sheets>").encode()
    parts["xl/_rels/workbook.xml.rels"] = rels.replace("</Relationships>", "".join(new_rels) + "</Relationships>").encode()
    parts["[Content_Types].xml"] = types.replace("</Types>", "".join(new_types) + "</Types>").encode()
    write_zip(OUT / f"large-{extra_sheets + 4}-sheets.xlsx", parts)


def make_docx(copies):
    parts = read_zip(FIXTURES / "docx-coverage.docx")
    doc = parts["word/document.xml"].decode()
    start = doc.index("<w:body>") + len("<w:body>")
    final = doc.rindex("<w:sectPr")
    end = doc.index("</w:body>")
    body, final_sect = doc[start:final], doc[final:end]
    parts["word/document.xml"] = (doc[:start] + body * copies + final_sect + doc[end:]).encode()
    write_zip(OUT / f"large-{copies}x.docx", parts)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--slides", type=int, default=300)
    ap.add_argument("--sheets", type=int, default=9, help="extra sheets on top of the 4 original")
    ap.add_argument("--sheet-copies", type=int, default=60, help="times the Shot Log rows repeat per extra sheet")
    ap.add_argument("--docx-copies", type=int, default=40)
    args = ap.parse_args()
    OUT.mkdir(parents=True, exist_ok=True)
    make_pptx(args.slides)
    make_xlsx(args.sheets, args.sheet_copies)
    make_docx(args.docx_copies)
