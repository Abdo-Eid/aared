import { readdir, rm } from "node:fs/promises";
import { join, resolve } from "node:path";
import { command, outputDirectory, requireRasterizer } from "../shared.js";

const OUTPUT_ROOT = join(import.meta.dir, "out");
/** 96 dpi is CSS pixels: a Word page rasterizes at the size aared paints it. */
const RASTER_DPI = 96;

export interface WordPageRender {
  file: string;
  widthPx: number;
  heightPx: number;
}

export interface WordRender {
  outputDir: string;
  wordVersion: string;
  pageCount: number;
  pages: WordPageRender[];
  /** JSON list of every text span with font, size and position in CSS pixels. */
  spansFile: string;
  rasterizer: { tool: "PyMuPDF"; version: string; dpi: number };
}

/**
 * Export a real DOCX with the installed Word (fields updated, print pipeline),
 * then rasterize every page with PyMuPDF and dump exact text geometry. Windows
 * only. The fixture is never opened in place, so it never gets a lock file.
 */
export async function renderWord(fixturePath: string, dpi = RASTER_DPI): Promise<WordRender> {
  if (process.platform !== "win32") {
    throw new Error("aared Word export: native Word export requires Windows with Microsoft Word installed");
  }
  if (!Number.isInteger(dpi) || dpi <= 0) throw new Error("aared Word export: dpi must be a positive integer");
  const pwsh = Bun.which("pwsh");
  if (pwsh === null) throw new Error("aared Word export: pwsh is required to automate Microsoft Word");
  const { python, version } = await requireRasterizer("Word");

  const fixture = resolve(fixturePath);
  const outputDir = outputDirectory(fixture, OUTPUT_ROOT);
  await rm(outputDir, { recursive: true, force: true });
  await command("Word", pwsh, [
    "-NoProfile", "-NonInteractive", "-STA", "-File", join(import.meta.dir, "export.ps1"),
    "-InputPath", fixture, "-OutputDir", outputDir,
  ]);

  const metadata = await Bun.file(join(outputDir, "word-export.json")).json() as { wordVersion: string; pageCount: number; pdf: string };
  const raster = JSON.parse(await command("Word", python, [
    join(import.meta.dir, "rasterize.py"), join(outputDir, metadata.pdf), outputDir, String(dpi),
  ])) as { pages: Array<{ width: number; height: number }> };
  const files = (await readdir(outputDir))
    .filter((name) => /^word-page-\d+\.png$/.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  if (files.length !== metadata.pageCount || raster.pages.length !== metadata.pageCount) {
    throw new Error(`aared Word export: Word reported ${metadata.pageCount} pages but ${files.length} PNGs were written`);
  }
  const pages = files.map((file, index) => ({ file, widthPx: raster.pages[index]!.width, heightPx: raster.pages[index]!.height }));
  const rasterizer = { tool: "PyMuPDF" as const, version, dpi };
  await Bun.write(join(outputDir, "word-export.json"), JSON.stringify({ ...metadata, pages, rasterizer }, null, 2));
  return { outputDir, wordVersion: metadata.wordVersion, pageCount: metadata.pageCount, pages, spansFile: join(outputDir, "word-spans.json"), rasterizer };
}

if (import.meta.main) {
  const [fixture, ...args] = process.argv.slice(2);
  if (fixture === undefined) throw new Error("Usage: bun tools/word-render/index.ts <fixture.docx> [--dpi <pixels>]");
  const dpiIndex = args.indexOf("--dpi");
  if (dpiIndex !== -1 && args[dpiIndex + 1] === undefined) throw new Error("aared Word export: --dpi requires a pixel value");
  console.log(JSON.stringify(await renderWord(fixture, dpiIndex === -1 ? RASTER_DPI : Number(args[dpiIndex + 1]))));
}
