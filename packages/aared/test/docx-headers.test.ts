import { expect, test } from "bun:test";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { paginate } from "../src/core/docx/pagination.js";
import type { FlowNode } from "../src/core/model.js";
import { OpcPackage } from "../src/core/opc.js";
import { SIZE, build, fourSectionDocument, page, refs } from "./docx-headers-doc.js";

// Which header a page shows is a rule of Word's, so the expected tables below
// are what Word 16 printed for these two documents (probe `docx-first-header`:
// the same XML exported to PDF and read back page by page).

const text = (nodes: FlowNode[] | undefined): string =>
  (nodes ?? []).map((node) => (node.type === "paragraph" || node.type === "heading" ? node.runs.map((run) => run.text).join("") : node.type)).join("|");

async function pagesOf(bytes: Uint8Array): Promise<Array<{ header: string; number: number }>> {
  const pkg = await OpcPackage.open(bytes.buffer as ArrayBuffer);
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  return paginate(sections).pages.map((p) => ({ header: text(p.header), number: p.number }));
}

test("first-page, even-page and inherited headers pick the page Word picks them for", async () => {
  const pages = await pagesOf(fourSectionDocument());
  expect(pages).toEqual([
    { header: "A-FIRST", number: 1 }, // titlePg: the section's first page
    { header: "A-EVEN", number: 2 },
    { header: "A-DEFAULT", number: 3 },
    { header: "A-EVEN", number: 4 },
    { header: "A-DEFAULT", number: 5 }, // B names nothing and has no titlePg: A's default and even headers
    { header: "A-EVEN", number: 6 },
    { header: "A-FIRST", number: 7 }, // C has titlePg but no first header of its own: A's, not its C-DEFAULT
    { header: "A-EVEN", number: 8 },
    { header: "D-DEFAULT", number: 1 }, // D restarts at 1, so its page 2 is even, and takes the inherited even header
    { header: "A-EVEN", number: 2 },
    { header: "D-DEFAULT", number: 3 },
  ]);
});

test("a first page with no first header anywhere is blank, and without evenAndOddHeaders every other page is the default", async () => {
  const only = `<w:sectPr>${refs([["default", "rId1"]])}${SIZE}<w:titlePg/></w:sectPr>`;
  const pages = await pagesOf(build(page("X1") + page("X2") + page("X3", true) + only, ""));
  expect(pages).toEqual([
    { header: "", number: 1 },
    { header: "A-DEFAULT", number: 2 },
    { header: "A-DEFAULT", number: 3 },
  ]);
});
