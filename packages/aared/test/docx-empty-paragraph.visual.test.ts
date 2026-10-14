import { expect } from "bun:test";
import { visual, withVisualPage } from "./harness.js";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { serveDocx } from "./docx-harness.js";
import { storedZip } from "./zip.js";

const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
const paragraph = (text: string, pPr = ""): string =>
  `<w:p><w:pPr>${pPr}</w:pPr>${text === "" ? "" : `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>${text}</w:t></w:r>`}</w:p>`;

/**
 * Word 16 laid out this very document (probe `docx-empty-paragraph`): each text
 * paragraph is one line, an empty one adds exactly one more line, an empty one
 * with a 24pt paragraph mark adds a 24pt line, and one with 12pt before and
 * after adds those 24pt on top of its line. The checks are proportions of the
 * control pitch, so they hold at any font metrics.
 */
visual("visual DOM output: empty paragraphs take a line, sized by their paragraph mark, like Word", async () => {
  const body =
    paragraph("ALPHA") + paragraph("BRAVO") +
    paragraph("CHARLIE") + paragraph("") + paragraph("DELTA") +
    paragraph("ECHO") + paragraph("") + paragraph("") + paragraph("") + paragraph("FOXTROT") +
    paragraph("GOLF") + paragraph("", `<w:rPr><w:sz w:val="48"/></w:rPr>`) + paragraph("HOTEL") +
    paragraph("INDIA") + paragraph("", `<w:spacing w:before="240" w:after="240"/>`) + paragraph("JULIET");
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const bytes = storedZip({
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/document.xml": document,
    "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    "word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
  });
  const dir = await mkdtemp(join(tmpdir(), "aared-empty-"));
  const path = join(dir, "empty.docx");
  await writeFile(path, bytes);
  const { server, sections } = await serveDocx(path, `<!doctype html><style>html,body{margin:0}#host{position:absolute;left:0;top:0}</style><main id="host"></main><script type="module">
      import { render } from "/docx.mjs";
      window.paint = async (model) => {
        await render({ document: model, target: document.querySelector("#host"), pagination: { mode: "layout" } });
        await document.fonts.ready;
        return [...document.querySelectorAll('[data-aared-element="paragraph"]')]
          .map((p) => ({ text: p.textContent.trim(), top: p.getBoundingClientRect().top, empty: p.querySelector("br") !== null }));
      };
    </script>`);
  try {
  await withVisualPage({ server, width: 1000, height: 1200 }, async (view) => {
    const rows = (await view.evaluate(`paint(${JSON.stringify(sections)})`)) as Array<{ text: string; top: number; empty: boolean }>;
    const top = (text: string): number => rows.find((row) => row.text === text)!.top;
    const pitch = top("BRAVO") - top("ALPHA");
    expect(pitch).toBeGreaterThan(10);
    // Control: text paragraphs one line apart; one empty paragraph is one line more, three are three.
    expect((top("DELTA") - top("CHARLIE")) / pitch).toBeCloseTo(2, 1);
    expect((top("FOXTROT") - top("ECHO")) / pitch).toBeCloseTo(4, 1);
    // A 24pt paragraph mark: Word's gap was 53.6px against a 16.8px line (3.19 lines); the empty line is 24pt, not 11pt.
    expect((top("HOTEL") - top("GOLF")) / pitch).toBeCloseTo(53.6 / 16.83, 1);
    // 12pt before and after: Word's gap was 65.76px (3.91 lines).
    expect((top("JULIET") - top("INDIA")) / pitch).toBeCloseTo(65.76 / 16.83, 1);
    expect(rows.filter((row) => row.text === "" && row.empty)).toHaveLength(6);
  });
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}, 60_000);
