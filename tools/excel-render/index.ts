import { readdir, rm } from "node:fs/promises";
import { basename, join, resolve } from "node:path";
import { command, outputDirectory, requireRasterizer } from "../shared.js";

const OUTPUT_ROOT = join(import.meta.dir, "out");
const RASTER_DPI = 150;

export interface ExcelPartSpec {
  /** 1-based tab position, matching the exported sheet index. */
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

async function rasterize(python: string, outputDir: string, pdfFile: string, dpi: number): Promise<ExcelPageRender> {
  const pdf = join(outputDir, pdfFile);
  const png = join(outputDir, pdfFile.replace(/\.pdf$/i, ".png"));
  const dims = JSON.parse(await command("Excel", python, [
    join(import.meta.dir, "rasterize.py"), pdf, png, String(dpi),
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
export async function renderExcel(fixturePath: string, parts: ExcelPartSpec[] = [], dpi = RASTER_DPI): Promise<ExcelRender> {
  if (process.platform !== "win32") {
    throw new Error("aared Excel export: native Excel export requires Windows with Microsoft Excel installed");
  }
  if (!Number.isInteger(dpi) || dpi <= 0) {
    throw new Error("aared Excel export: dpi must be a positive integer");
  }
  const pwsh = Bun.which("pwsh");
  if (pwsh === null) throw new Error("aared Excel export: pwsh is required to automate Microsoft Excel");
  const { python, version: pymupdfVersion } = await requireRasterizer("Excel");

  const fixture = resolve(fixturePath);
  const outputDir = outputDirectory(fixture, OUTPUT_ROOT);
  await rm(outputDir, { recursive: true, force: true });
  await command("Excel", pwsh, [
    "-NoProfile", "-NonInteractive", "-STA", "-File", join(import.meta.dir, "export.ps1"),
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
    const full = await rasterize(python, outputDir, sheet.file, dpi);
    const renderedParts: ExcelPartRender[] = [];
    for (const part of sheet.parts) {
      const page = await rasterize(python, outputDir, part.file, dpi);
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
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))
    .map((name) => join(outputDir, name));
  const expected = sheets.flatMap((sheet) => [sheet.file, ...sheet.parts.map((part) => part.file)]).filter((file) => file !== null).length;
  if (pages.length !== expected) {
    throw new Error(`aared Excel export: expected ${expected} PNGs but found ${pages.length}`);
  }
  return { outputDir, excelVersion: metadata.excelVersion, sheetCount: metadata.sheetCount, sheets, rasterizer, pages };
}

if (import.meta.main) {
  const [fixture, ...args] = process.argv.slice(2);
  if (fixture === undefined) {
    throw new Error("Usage: bun tools/excel-render/index.ts <fixture.xlsx> [--part sheetIndex:name=range ...] [--dpi <pixels>]");
  }
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
  if (parts.some((part) => !Number.isInteger(part.sheet) || part.sheet <= 0 || part.name === "" || part.range === "")) {
    throw new Error("aared Excel export: malformed --part (want sheetIndex:name=range)");
  }
  const dpiIndex = args.indexOf("--dpi");
  const dpi = dpiIndex === -1 ? RASTER_DPI : Number(args[dpiIndex + 1]);
  if (dpiIndex !== -1 && args[dpiIndex + 1] === undefined) {
    throw new Error("aared Excel export: --dpi requires a pixel value");
  }
  console.log(JSON.stringify(await renderExcel(fixture, parts, dpi)));
}
