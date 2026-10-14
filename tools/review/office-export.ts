import { readdir, rm } from "node:fs/promises";
import { basename, join, resolve } from "node:path";
import { parseArgs } from "node:util";
import { command, outputDirectory, requireRasterizer, partSlug } from "../shared/office.js";
import { reviewOutputRoot } from "../shared/paths.js";

/**
 * Native Office export driver for the three review pipelines.
 *
 * One skeleton drives all three formats: guard the platform, clean the
 * pipeline's `out/` directory, spawn that format's `export.ps1`, then check
 * the PNG count against the metadata the script wrote. Only the Office
 * call itself is format-specific (different object models, print vs slide
 * vs layout semantics) and stays in each pipeline's own `export.ps1`.
 *
 *   bun tools/review/office-export.ts --format pptx <fixture.pptx> [--width <pixels>]
 *   bun tools/review/office-export.ts --format xlsx <fixture.xlsx> [--part sheet:name=range ...] [--dpi <dpi>]
 *   bun tools/review/office-export.ts --format docx <fixture.docx> [--dpi <dpi>]
 */

const POWERPOINT_WIDTH = 1280;
const EXCEL_RASTER_DPI = 150;
/** 96 dpi is CSS pixels: a Word page rasterizes at the size aared paints it. */
const WORD_RASTER_DPI = 96;

const RASTERIZE = join(import.meta.dir, "rasterize.py");

function pipelineDir(format: "pptx" | "xlsx" | "docx"): string {
  return join(import.meta.dir, format);
}

async function windowsOffice(label: string, app: string): Promise<string> {
  if (process.platform !== "win32") {
    throw new Error(`aared ${label} export: native ${app} export requires Windows with Microsoft ${app} installed`);
  }
  const pwsh = Bun.which("pwsh");
  if (pwsh === null) throw new Error(`aared ${label} export: pwsh is required to automate Microsoft ${app}`);
  return pwsh;
}

function numericSort(left: string, right: string): number {
  return left.localeCompare(right, undefined, { numeric: true });
}

export interface PowerPointRender {
  outputDir: string;
  powerpointVersion: string;
  slideCount: number;
  width: number;
  height: number;
  pages: string[];
}

/** Export a real PPTX with the installed PowerPoint renderer. Windows only. */
export async function renderPowerPoint(fixturePath: string, width = POWERPOINT_WIDTH): Promise<PowerPointRender> {
  const pwsh = await windowsOffice("PowerPoint", "PowerPoint");
  if (!Number.isInteger(width) || width <= 0) {
    throw new Error("aared PowerPoint export: width must be a positive integer");
  }

  const dir = pipelineDir("pptx");
  const fixture = resolve(fixturePath);
  const outputDir = outputDirectory(fixture, reviewOutputRoot("pptx"));
  if (!(await Bun.file(fixture).exists())) throw new Error(`aared PowerPoint export: input does not exist: ${fixture}`);
  await rm(outputDir, { recursive: true, force: true });
  await command("PowerPoint", pwsh, ["-NoProfile", "-NonInteractive", "-STA", "-File", join(dir, "export.ps1"), "-InputPath", fixture, "-OutputDir", outputDir, "-Width", String(width)]);

  const metadata = await Bun.file(join(outputDir, "powerpoint-export.json")).json() as Omit<PowerPointRender, "outputDir" | "pages">;
  const pages = (await readdir(outputDir))
    .filter((name) => /^powerpoint-slide-\d+\.png$/.test(name))
    .sort(numericSort)
    .map((name) => join(outputDir, name));
  if (pages.length !== metadata.slideCount) {
    throw new Error(`aared PowerPoint export: expected ${metadata.slideCount} PNGs but found ${pages.length}`);
  }
  return { outputDir, ...metadata, pages };
}

export interface ExcelPartSpec {
  /** 1-based position among the visible sheets (hidden ones are skipped), matching the exported sheet index. */
  sheet: number;
  name: string;
  range: string;
}

export interface ExcelPageRender {
  file: string;
  widthPx: number;
  heightPx: number;
}

export interface ExcelPartRender extends ExcelPageRender {
  name: string;
  range: string;
}

export interface ExcelSheetRender {
  index: number;
  name: string;
  direction: "rtl" | "ltr";
  usedRange: string;
  empty: boolean;
  file: string | null;
  widthPx: number;
  heightPx: number;
  parts: ExcelPartRender[];
}

export interface ExcelRasterizer {
  tool: "PyMuPDF";
  version: string;
  dpi: number;
}

export interface ExcelRender {
  outputDir: string;
  excelVersion: string;
  sheetCount: number;
  sheets: ExcelSheetRender[];
  rasterizer: ExcelRasterizer;
  pages: string[];
}

interface ExportedPart {
  name: string;
  range: string;
  file: string;
}

interface ExportedSheet {
  index: number;
  name: string;
  direction: "rtl" | "ltr";
  usedRange: string;
  empty: boolean;
  file: string | null;
  parts: ExportedPart[];
}

async function rasterizeSheet(python: string, outputDir: string, pdfFile: string, dpi: number): Promise<ExcelPageRender> {
  const pdf = join(outputDir, pdfFile);
  const png = join(outputDir, pdfFile.replace(/\.pdf$/i, ".png"));
  const dims = JSON.parse(await command("Excel", python, [
    RASTERIZE, "--mode", "sheet", pdf, png, String(dpi),
  ])) as { width: number; height: number };
  if (!Number.isInteger(dims.width) || !Number.isInteger(dims.height) || dims.width <= 0 || dims.height <= 0) {
    throw new Error(`aared Excel export: rasterizer reported invalid dimensions for ${pdfFile}`);
  }
  await rm(pdf, { force: true });
  return { file: basename(png), widthPx: dims.width, heightPx: dims.height };
}

/**
 * Export a real XLSX with the installed Excel renderer (A4 print path, fit
 * sheet on one page) plus the curated print-area parts, then rasterize every
 * PDF with PyMuPDF. Windows only.
 */
export async function renderExcel(fixturePath: string, parts: ExcelPartSpec[] = [], dpi = EXCEL_RASTER_DPI): Promise<ExcelRender> {
  validateParts(parts);
  const pwsh = await windowsOffice("Excel", "Excel");
  if (!Number.isInteger(dpi) || dpi <= 0) {
    throw new Error("aared Excel export: dpi must be a positive integer");
  }
  const { python, version: pymupdfVersion } = await requireRasterizer("Excel");

  const dir = pipelineDir("xlsx");
  const fixture = resolve(fixturePath);
  const outputDir = outputDirectory(fixture, reviewOutputRoot("xlsx"));
  if (!(await Bun.file(fixture).exists())) throw new Error(`aared Excel export: input does not exist: ${fixture}`);
  await rm(outputDir, { recursive: true, force: true });
  await command("Excel", pwsh, [
    "-NoProfile", "-NonInteractive", "-STA", "-File", join(dir, "export.ps1"),
    "-InputPath", fixture, "-OutputDir", outputDir,
    "-PartsJson", JSON.stringify(parts),
  ]);

  const metadata = await Bun.file(join(outputDir, "excel-export.json")).json() as {
    excelVersion: string;
    sheetCount: number;
    sheets: ExportedSheet[];
  };
  const sheets: ExcelSheetRender[] = [];
  for (const sheet of metadata.sheets) {
    if (sheet.empty || sheet.file === null) {
      if (sheet.parts.length > 0) throw new Error(`aared Excel export: sheet ${sheet.index} is empty but has curated parts`);
      sheets.push({ ...sheet, file: null, widthPx: 0, heightPx: 0, parts: [] });
      continue;
    }
    const full = await rasterizeSheet(python, outputDir, sheet.file, dpi);
    const renderedParts: ExcelPartRender[] = [];
    for (const part of sheet.parts) {
      const page = await rasterizeSheet(python, outputDir, part.file, dpi);
      renderedParts.push({ ...page, name: part.name, range: part.range });
    }
    sheets.push({ ...sheet, file: full.file, widthPx: full.widthPx, heightPx: full.heightPx, parts: renderedParts });
  }
  const rasterizer: ExcelRasterizer = { tool: "PyMuPDF", version: pymupdfVersion, dpi };
  await Bun.write(join(outputDir, "excel-export.json"), JSON.stringify({
    excelVersion: metadata.excelVersion,
    sheetCount: metadata.sheetCount,
    sheets,
    rasterizer,
  }, null, 2));
  const pages = (await readdir(outputDir))
    .filter((name) => /^excel-sheet-\d+(-part-[A-Za-z0-9._-]+)?\.png$/.test(name))
    .sort(numericSort)
    .map((name) => join(outputDir, name));
  const expected = sheets.flatMap((sheet) => [sheet.file, ...sheet.parts.map((part) => part.file)]).filter((file) => file !== null).length;
  if (pages.length !== expected) {
    throw new Error(`aared Excel export: expected ${expected} PNGs but found ${pages.length}`);
  }
  return { outputDir, excelVersion: metadata.excelVersion, sheetCount: metadata.sheetCount, sheets, rasterizer, pages };
}

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
export async function renderWord(fixturePath: string, dpi = WORD_RASTER_DPI): Promise<WordRender> {
  const pwsh = await windowsOffice("Word", "Word");
  if (!Number.isInteger(dpi) || dpi <= 0) throw new Error("aared Word export: dpi must be a positive integer");
  const { python, version } = await requireRasterizer("Word");

  const dir = pipelineDir("docx");
  const fixture = resolve(fixturePath);
  const outputDir = outputDirectory(fixture, reviewOutputRoot("docx"));
  if (!(await Bun.file(fixture).exists())) throw new Error(`aared Word export: input does not exist: ${fixture}`);
  await rm(outputDir, { recursive: true, force: true });
  await command("Word", pwsh, [
    "-NoProfile", "-NonInteractive", "-STA", "-File", join(dir, "export.ps1"),
    "-InputPath", fixture, "-OutputDir", outputDir,
  ]);

  const metadata = await Bun.file(join(outputDir, "word-export.json")).json() as { wordVersion: string; pageCount: number; pdf: string };
  const raster = JSON.parse(await command("Word", python, [
    RASTERIZE, "--mode", "document", join(outputDir, metadata.pdf), outputDir, String(dpi),
  ])) as { pages: Array<{ width: number; height: number }> };
  const files = (await readdir(outputDir))
    .filter((name) => /^word-page-\d+\.png$/.test(name))
    .sort(numericSort);
  if (files.length !== metadata.pageCount || raster.pages.length !== metadata.pageCount) {
    throw new Error(`aared Word export: Word reported ${metadata.pageCount} pages but ${files.length} PNGs were written`);
  }
  const pages = files.map((file, index) => ({ file, widthPx: raster.pages[index]!.width, heightPx: raster.pages[index]!.height }));
  const rasterizer = { tool: "PyMuPDF" as const, version, dpi };
  await Bun.write(join(outputDir, "word-export.json"), JSON.stringify({ ...metadata, pages, rasterizer }, null, 2));
  return { outputDir, wordVersion: metadata.wordVersion, pageCount: metadata.pageCount, pages, spansFile: join(outputDir, "word-spans.json"), rasterizer };
}

/** Parse repeated `--part sheetIndex:name=range` arguments. */
export function parseParts(args: string[]): ExcelPartSpec[] {
  const parts: ExcelPartSpec[] = [];
  for (let i = 0; i < args.length; i++) {
    if (args[i] !== "--part") continue;
    const spec = args[i + 1];
    if (spec === undefined) throw new Error("aared Excel export: --part requires sheetIndex:name=range");
    const separator = spec.indexOf(":");
    const assignment = spec.indexOf("=");
    if (separator === -1 || assignment === -1 || assignment < separator) {
      throw new Error(`aared Excel export: malformed --part '${spec}' (want sheetIndex:name=range)`);
    }
    parts.push({
      sheet: Number(spec.slice(0, separator)),
      name: spec.slice(separator + 1, assignment),
      range: spec.slice(assignment + 1),
    });
    i++;
  }
  validateParts(parts);
  return parts;
}

function validateParts(parts: readonly ExcelPartSpec[]): void {
  if (parts.some((part) => !Number.isInteger(part.sheet) || part.sheet <= 0 || part.name === "" || part.range === "")) {
    throw new Error("aared Excel export: malformed --part (want sheetIndex:name=range)");
  }
  const names = new Set<string>();
  for (const part of parts) {
    const key = `${part.sheet}:${partSlug(part.name).toLowerCase()}`;
    if (names.has(key)) throw new Error(`aared Excel export: part names produce the same file name on sheet ${part.sheet}`);
    names.add(key);
  }
}

if (import.meta.main) {
  const [formatFlag, format, fixture, ...args] = process.argv.slice(2);
  if (formatFlag !== "--format" || format === undefined || fixture === undefined) {
    throw new Error("Usage: bun tools/review/office-export.ts --format <pptx|xlsx|docx> <fixture> [options]");
  }
  if (format === "pptx") {
    const { values } = parseArgs({ args, options: { width: { type: "string" } } });
    const width = values.width === undefined ? POWERPOINT_WIDTH : Number(values.width);
    console.log(JSON.stringify(await renderPowerPoint(fixture, width)));
  } else if (format === "xlsx") {
    const { values } = parseArgs({ args, options: { dpi: { type: "string" }, part: { type: "string", multiple: true } } });
    console.log(JSON.stringify(await renderExcel(fixture, parseParts((values.part ?? []).flatMap((part) => ["--part", part])), values.dpi === undefined ? EXCEL_RASTER_DPI : Number(values.dpi))));
  } else if (format === "docx") {
    const { values } = parseArgs({ args, options: { dpi: { type: "string" } } });
    console.log(JSON.stringify(await renderWord(fixture, values.dpi === undefined ? WORD_RASTER_DPI : Number(values.dpi))));
  } else {
    throw new Error(`aared export: unknown format '${format ?? ""}' (want pptx|xlsx|docx)`);
  }
}
