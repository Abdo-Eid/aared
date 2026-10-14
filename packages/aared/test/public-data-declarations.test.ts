import { expect, test } from "bun:test";
import { mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";

const ROOT = resolve(import.meta.dir, "../../..");
const DIST = resolve(import.meta.dir, "../dist");
const OUT = join(import.meta.dir, ".data-declarations");

/** Follow every dependency from the actual package-exported declaration roots. */
function declarationGraph(): Map<string, string> {
  const pkg = JSON.parse(readFileSync(resolve(import.meta.dir, "../package.json"), "utf8")) as { exports: Record<string, { types?: string } | string> };
  const seen = new Map<string, string>();
  const visit = (file: string): void => {
    if (seen.has(file)) return;
    const source = readFileSync(file, "utf8").replace(/\/\*[\s\S]*?\*\/|\/\/[^\n]*/g, "");
    seen.set(file, source);
    for (const match of source.matchAll(/(?:from\s*|import\s*\()?["'](\.{1,2}\/[^"']+)["']/g)) {
      const path = match[1]!;
      visit(resolve(dirname(file), path.replace(/\.mjs$/, ".d.mts")));
    }
  };
  for (const entry of Object.values(pkg.exports)) {
    if (typeof entry === "object" && entry.types) visit(resolve(DIST, "..", entry.types));
  }
  return seen;
}

test("built public declaration dependencies expose no rendering models or session machinery", () => {
  const model = readFileSync(resolve(import.meta.dir, "../src/core/model.ts"), "utf8");
  const internal = [...model.matchAll(/^export (?:interface|type) (\w+)/gm)].map(match => match[1]!)
    .filter(name => name !== "CellRange" && name !== "RenderWarning");
  internal.push("ChartSpecModel", "WorkerSession", "HeadingOutline", "PaintSlideInput");
  const graph = declarationGraph();
  expect(graph.size).toBeGreaterThan(7);
  for (const [file, source] of graph) for (const name of internal) {
    expect(source, `${file} leaks ${name}`).not.toMatch(new RegExp(`\\b${name}\\b`));
  }
});

test("a consumer can inspect supported data but cannot access anonymous paint fields or model rendering", () => {
  mkdirSync(OUT, { recursive: true });
  try {
    writeFileSync(join(OUT, "consumer.ts"), `
import { open as openXlsx, render as renderXlsx, type CellData, type SheetData, type CellRange } from '@aared/viewer/xlsx';
import { open as openPptx, render as renderPptx, type SlideText, type PptxDocument } from '@aared/viewer/pptx';
import { open as openDocx, render as renderDocx, type SectionInfo } from '@aared/viewer/docx';
import type { PreviewHandle } from '@aared/viewer/preview';
declare const file: File, target: HTMLElement, viewer: PreviewHandle;
const workbook = await openXlsx(file);
const sheet = workbook.sheets.get(0);
const data: SheetData = await sheet.read();
const range: CellRange = data.range;
const cell: CellData = data.rows[0]!.cells[0]!.cell;
cell.formatted satisfies string;
// @ts-expect-error styles are not a data contract
cell.style;
// @ts-expect-error conditional paint is private
cell.conditional;
// @ts-expect-error drawings are private
data.drawings;
// @ts-expect-error track dimensions are private
data.columnWidths;
// @ts-expect-error row layout is private
data.rows[0]!.height;
// @ts-expect-error data cannot be rendered
await renderXlsx({sheet:data,target});
const deck = await openPptx(file);
const text: SlideText = await deck.slides.get(0).read();
text.paragraphs satisfies string[];
// @ts-expect-error slide elements are private
text.elements;
// @ts-expect-error semantic text cannot be rendered
await renderPptx({slide:text,target});
// @ts-expect-error documents have no public constructors
new PptxDocument();
const doc = await openDocx(file);
const section: SectionInfo = doc.sections.get(0);
section.page.width satisfies number;
// @ts-expect-error flow nodes are private
section.root;
// @ts-expect-error headers are private rich content
section.header;
// @ts-expect-error section metadata cannot be rendered
await renderDocx({document:[section],target});
// @ts-expect-error internal outline is not public
doc.outline;
// @ts-expect-error implementation/session methods are private
sheet.sheetDocument();
if (viewer.format === 'xlsx') {
  const borrowed = await viewer.content.sheets.get(0).read();
  borrowed satisfies SheetData;
  viewer.addEventListener('cellactivate', event => {
    event.detail.cell satisfies CellData;
    // @ts-expect-error the event no longer exposes the model
    event.detail.model;
  });
}
if (viewer.format === 'pptx') {
  const borrowed = await viewer.content.slides.get(0).read();
  // @ts-expect-error borrowed reads do not expose models
  borrowed.elements;
}
if (viewer.format === 'docx') {
  // @ts-expect-error borrowed sections do not expose flow roots
  viewer.content.sections.get(0).root;
}
// @ts-expect-error the core no longer exports cell models
import type { CellModel } from '@aared/viewer/core';
// @ts-expect-error synchronous model painting is not public
import { paintSlide } from '@aared/viewer/pptx';
`);
    writeFileSync(join(OUT, "tsconfig.json"), JSON.stringify({
      extends: join(ROOT, "tsconfig.base.json"),
      compilerOptions: { noEmit: true, declaration: false, declarationMap: false,
        paths: { "@aared/viewer/*": [join(DIST, "*", "index.d.mts")] } },
      include: ["consumer.ts"],
    }));
    const tsc = join(ROOT, "node_modules", ".bin", process.platform === "win32" ? "tsc.exe" : "tsc");
    const result = Bun.spawnSync([tsc, "-p", join(OUT, "tsconfig.json")]);
    expect(result.stdout.toString() + result.stderr.toString()).toBe("");
    expect(result.exitCode).toBe(0);
  } finally { rmSync(OUT, { recursive: true, force: true }); }
}, 120000);
