import { createHash } from "node:crypto";
import { writeFile } from "node:fs/promises";
import { basename, join, resolve } from "node:path";
import { open } from "../../packages/aared/src/xlsx/index.js";
import { collectStrings } from "../../packages/aared/src/core/collect.js";
import { absentFontsModule } from "../../packages/aared/test/absent-fonts.js";
import { within } from "../../packages/aared/test/harness.js";
import { buildXlsxRenderer, hostPage, writeReview, type AaredPart, type AaredSheet } from "../../packages/aared/test/xlsx-review.js";
import { parseParts, renderExcel, type ExcelPartSpec } from "./office-export.js";
import { command, requireRasterizer, partSlug } from "../shared/office.js";

const PAIR = join(import.meta.dir, "pair-images.py");

/**
 * Pair any workbook's native Excel print with aared's A4 render, the same
 * review the committed-fixture visual suite writes, for a file outside
 * `fixtures/`. Windows, Excel and PyMuPDF are required (see
 * `tools/review/xlsx/README.md`).
 *
 *   bun tools/review/xlsx.ts <workbook.xlsx> [--part sheet:name=range ...] [--pair-full] [--absent-fonts]
 *
 * Output goes to `tools/out/xlsx/<stem>/`: the native pages, the aared
 * pages (`aared-sheet-N.png`, `aared-sheet-N-part-<name>.png`), `review.html`,
 * and cropped, side-by-side pairs for curated parts
 * (`pair-sheet-N-part-<name>.png`). Add `--pair-full` to also pair complete
 * sheets. `--part` uses the 1-based sheet position among visible sheets, as
 * `office-export.ts` does. `--absent-fonts` adds a third column rendered as a
 * device without Office's fonts would (network).
 */

const VIEW = { width: 810, height: 1218 };

export async function captureXlsx(file: string, options: {
  absentFonts?: boolean; pairFull?: boolean; parts?: ExcelPartSpec[];
} = {}): Promise<{ review: string; outputDir: string; pairs: string[] }> {
  const absentFonts = options.absentFonts ?? false;
  const pairFull = options.pairFull ?? false;
  const parts = options.parts ?? [];
  const fixture = resolve(file);
  const bytes = await Bun.file(fixture).arrayBuffer();
  const fixtureHash = createHash("sha256").update(new Uint8Array(bytes)).digest("hex");
  const { python } = await requireRasterizer("xlsx capture");
  if (parts.length > 0 || pairFull) {
    try { await command("pairing prerequisites", python, ["-c", "import PIL"]); }
    catch { throw new Error("XLSX PNG pairs require Pillow; run pip install Pillow."); }
  }
  const reference = await renderExcel(fixture, parts);

  /** Crop the Excel print and aared capture, then place them side by side. */
  async function pair(excelFile: string, aaredFile: string, output: string): Promise<string> {
    await command("pairing", python, [
      PAIR, join(reference.outputDir, excelFile), join(reference.outputDir, aaredFile),
      join(reference.outputDir, output), "Excel", "aared",
    ]);
    return output;
  }

  const workbook = await open(bytes);
  const models: unknown[] = [];
  const titles: string[] = [];
  const media: Record<string, string> = {};
  try {
    for (const [index, handle] of workbook.sheets.list().entries()) {
      if (workbook.info.sheetStates?.[index] !== undefined && workbook.info.sheetStates[index] !== "visible") continue;
      models.push(await handle.read());
      titles.push(workbook.info.titles[index] ?? `sheet-${index + 1}`);
    }
    if (reference.sheetCount !== models.length) throw new Error(`Excel printed ${reference.sheetCount} visible sheets, aared read ${models.length}`);
    if (parts.some((part) => part.sheet > models.length)) throw new Error("Requested XLSX part names a nonexistent visible sheet.");
    const ids = new Set<string>();
    for (const key of ["mediaId", "pictureMediaId", "fallbackMediaId", "pictureFallbackMediaId"]) collectStrings(models, key, ids);
    for (const id of ids) media[id] = Buffer.from(await workbook.media.get(id)).toString("base64");
  } finally {
    workbook.close();
  }

  const module = await buildXlsxRenderer();
  const absentModule = absentFonts ? await absentFontsModule() : "";
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const path = new URL(request.url).pathname;
      if (path === "/xlsx.mjs") return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      if (path === "/absent-fonts.mjs") return new Response(absentModule, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      return new Response(hostPage(absentFonts), { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });

  const sheets: AaredSheet[] = [];
  const pairs: string[] = [];
  try {
    await using view = new Bun.WebView(VIEW);
    await within("loading the WebView host", view.navigate(server.url.href), 30_000);
    const mediaJson = JSON.stringify(media);
    for (const [index, model] of models.entries()) {
      const label = `${basename(fixture)} sheet ${index + 1}`;
      const full = await within(`rendering ${label}`, view.evaluate(`renderFull(${JSON.stringify(model)}, ${mediaJson})`), 30_000) as {
        direction: string | null; canvasCount: number; scale: number;
      };
      if (full.canvasCount !== 0) throw new Error(`aared capture: ${label} painted a canvas; the grid must be DOM`);
      const nativeDirection = reference.sheets[index]?.direction ?? null;
      if (full.direction !== nativeDirection) throw new Error(`aared capture: ${label} is ${full.direction} in aared but ${nativeDirection} in Excel`);
      const fullFile = `aared-sheet-${index + 1}.png`;
      await writeFile(join(reference.outputDir, fullFile), await within(`capturing ${label}`, view.screenshot({ encoding: "buffer", format: "png" }), 30_000));
      if (absentFonts) {
        await within(`rendering ${label} without Office fonts`, view.evaluate(`renderFull(${JSON.stringify(model)}, ${mediaJson}, true)`), 60_000);
        await writeFile(join(reference.outputDir, `aared-absent-sheet-${index + 1}.png`), await within(`capturing ${label} without Office fonts`, view.screenshot({ encoding: "buffer", format: "png" }), 30_000));
      }

      const sheetParts: AaredPart[] = [];
      for (const part of parts.filter((entry) => entry.sheet === index + 1)) {
        // bandPart shifts/hides the given DOM; each part needs a fresh full render.
        await within(`rendering ${part.name}`, view.evaluate(`renderFull(${JSON.stringify(model)}, ${mediaJson})`), 30_000);
        const band = await within(`isolating ${part.name}`, view.evaluate(`bandPart(${JSON.stringify(model)}, ${JSON.stringify(part.range)})`), 30_000) as { scale: number };
        const slug = partSlug(part.name);
        const file = `aared-sheet-${index + 1}-part-${slug}.png`;
        await writeFile(join(reference.outputDir, file), await within(`capturing ${part.name}`, view.screenshot({ encoding: "buffer", format: "png" }), 30_000));
        sheetParts.push({ name: part.name, range: part.range, file, scale: band.scale });
        if (absentFonts) {
          await within(`rendering ${part.name} without Office fonts`, view.evaluate(`renderFull(${JSON.stringify(model)}, ${mediaJson}, true)`), 60_000);
          await within(`isolating ${part.name} without Office fonts`, view.evaluate(`bandPart(${JSON.stringify(model)}, ${JSON.stringify(part.range)})`), 30_000);
          await writeFile(join(reference.outputDir, `aared-absent-${file.slice("aared-".length)}`), await within(`capturing ${part.name} without Office fonts`, view.screenshot({ encoding: "buffer", format: "png" }), 30_000));
        }
        const nativePart = reference.sheets[index]?.parts.find((entry) => entry.name === part.name);
        if (nativePart === undefined) throw new Error(`Native XLSX export is missing part ${part.name}.`);
        pairs.push(await pair(nativePart.file, file, `pair-sheet-${index + 1}-part-${slug}.png`));
      }
      const nativeFull = reference.sheets[index]?.file ?? null;
      if (pairFull && nativeFull !== null) pairs.push(await pair(nativeFull, fullFile, `pair-sheet-${index + 1}.png`));
      sheets.push({ name: titles[index] ?? `sheet-${index + 1}`, direction: full.direction, file: fullFile, scale: full.scale, parts: sheetParts });
    }
  } finally {
    server.stop();
  }
  await writeReview(reference, basename(fixture), fixtureHash, sheets, absentFonts);
  return { review: join(reference.outputDir, "review.html"), outputDir: reference.outputDir, pairs: pairs.map((file) => join(reference.outputDir, file)) };
}

if (import.meta.main) {
  const argv = process.argv.slice(2);
  const file = argv[0];
  if (file === undefined || file.startsWith("--")) throw new Error("Usage: bun tools/review/xlsx.ts <workbook.xlsx> [--part sheet:name=range ...] [--pair-full] [--absent-fonts]");
  for (let i = 1; i < argv.length; i++) {
    if (argv[i] === "--part") { i++; continue; }
    if (argv[i] !== "--pair-full" && argv[i] !== "--absent-fonts") throw new Error(`Unknown XLSX capture option: ${argv[i]}`);
  }
  console.log(JSON.stringify(await captureXlsx(file, {
    parts: parseParts(argv), absentFonts: argv.includes("--absent-fonts"), pairFull: argv.includes("--pair-full"),
  })));
}
