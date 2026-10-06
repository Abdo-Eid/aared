"""Remove the org sensitivity label Office (PowerPoint, Word, Excel) adds on save, in place.

Removes only: shapes whose nvPr carries a p184:classification extension (the
"Public" content marking PowerPoint puts on the master); in a Word package,
the header/footer runs holding the marking text box (its wp:docPr `descr` is
the label's marking text), then any even/first-page header or footer left
with no run (Word adds those only to carry the marking); and MSIP_Label_* / ClassificationContentMarking*
custom properties (the whole custom.xml part, its relationship and content
type when nothing else is left) — this last path is format-agnostic, and in
Excel it is the whole job (the label lives only in those properties).
Every other part is copied byte for byte.
"""
import re
import shutil
import sys
import tempfile
import zipfile

PROP = re.compile(r'<property [^>]*name="(?:MSIP_Label|ClassificationContentMarking)[^"]*"[^>]*>.*?</property>', re.S)


def strip_shapes(xml: str) -> str:
    out, pos = [], 0
    for match in re.finditer(r"<p:sp>.*?</p:sp>", xml, re.S):
        if "classification" in match.group(0):
            out.append(xml[pos:match.start()])
            pos = match.end()
    out.append(xml[pos:])
    return "".join(out)


MARKING_TEXT = re.compile(r'name="ClassificationContentMarkingHeaderText"[^>]*>\s*<vt:lpwstr>([^<]*)</vt:lpwstr>')
# Word writes the marking as one run: a no-proof drawing whose AlternateContent is never nested.
# The run's properties vary (a right-to-left document adds `w:rtl` and a language), so any `w:rPr` is allowed.
MARKING_RUN = re.compile(r"<w:r>(?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?<mc:AlternateContent>.*?</mc:AlternateContent></w:r>", re.S)


def strip_word_marking(xml: str, text: str) -> str:
    return MARKING_RUN.sub(lambda run: "" if f'descr="{text}"' in run.group(0) else run.group(0), xml)


def drop_empty_alternates(parts: dict, drop: set) -> None:
    """Word adds even and first-page headers/footers just to carry the marking.

    Once the marking is gone such a part holds no run: drop its reference, its
    relationship, its content type and the part itself.
    """
    rels_name = "word/_rels/document.xml.rels"
    if "word/document.xml" not in parts or rels_name not in parts:
        return
    document = parts["word/document.xml"].decode("utf-8")
    rels = parts[rels_name].decode("utf-8")
    types = parts["[Content_Types].xml"].decode("utf-8")
    for ref in re.findall(r'<w:(?:header|footer)Reference w:type="(?:even|first)" r:id="\w+"/>', document):
        rid_match = re.search(r'r:id="(\w+)"', ref)
        assert rid_match is not None
        rid = rid_match.group(1)
        rel = re.search(rf'<Relationship [^>]*Id="{rid}"[^>]*/>', rels)
        if rel is None:
            continue
        target_match = re.search(r'Target="([^"]+)"', rel.group(0))
        assert target_match is not None
        name = "word/" + target_match.group(1)
        if name not in parts or re.search(rb"<w:r[ >]", parts[name]):
            continue
        document = document.replace(ref, "")
        rels = rels.replace(rel.group(0), "")
        types = re.sub(rf'<Override PartName="/{re.escape(name)}"[^>]*/>', "", types)
        drop.add(name)
    parts["word/document.xml"] = document.encode("utf-8")
    parts[rels_name] = rels.encode("utf-8")
    parts["[Content_Types].xml"] = types.encode("utf-8")


def main(path: str) -> None:
    src = zipfile.ZipFile(path)
    parts = {name: src.read(name) for name in src.namelist()}
    infos = {info.filename: info for info in src.infolist()}
    src.close()
    drop = set()
    for name, data in parts.items():
        if re.match(r"ppt/(slideMasters|slideLayouts|slides)/[^/]+\.xml$", name) and b"classification" in data:
            parts[name] = strip_shapes(data.decode("utf-8")).encode("utf-8")
    custom = parts.get("docProps/custom.xml")
    marking = None if custom is None else MARKING_TEXT.search(custom.decode("utf-8"))
    # A file already stripped once has no custom.xml: a header Word added since still carries "Public".
    marking_text = marking.group(1) if marking is not None else ("Public" if "word/document.xml" in parts else None)
    if marking_text is not None:
        for name, data in parts.items():
            if re.match(r"word/(header|footer)\d+\.xml$", name):
                parts[name] = strip_word_marking(data.decode("utf-8"), marking_text).encode("utf-8")
    drop_empty_alternates(parts, drop)
    if custom is not None:
        text = PROP.sub("", custom.decode("utf-8"))
        if "<property " in text:
            parts["docProps/custom.xml"] = text.encode("utf-8")
        else:
            drop.add("docProps/custom.xml")
            parts["_rels/.rels"] = re.sub(r'<Relationship [^>]*Target="docProps/custom.xml"[^>]*/>', "", parts["_rels/.rels"].decode("utf-8")).encode("utf-8")
            parts["[Content_Types].xml"] = re.sub(r'<Override PartName="/docProps/custom.xml"[^>]*/>', "", parts["[Content_Types].xml"].decode("utf-8")).encode("utf-8")
    tmp = tempfile.mktemp(suffix=path[path.rfind("."):])
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in parts.items():
            if name not in drop:
                out.writestr(infos[name], data)
    shutil.move(tmp, path)
    print("stripped", path, "dropped", sorted(drop))


if __name__ == "__main__":
    main(sys.argv[1])
