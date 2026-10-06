import { describe, expect, test } from "bun:test";
import { AaredError, isAbort } from "../src/core/errors.js";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { parseXlsxSheet } from "../src/core/xlsx/sheet.js";
import { readWorkbook } from "../src/core/xlsx/workbook.js";
import { parseXml } from "../src/core/xml.js";
import { MAX_ZIP_ENTRY_BYTES, getEntryData, readEntries } from "../src/core/zip.js";
import { FIXTURES, withPart } from "./mutate.js";
import { storedZip } from "./zip.js";

// Untrusted input: a file can claim any count, length or size. Whatever it
// claims, reading it must end in a result or a typed error, never a hang, an
// out-of-memory or a bare RangeError that the Worker would report as an aared
// bug. Hostile zips are synthetic (a pure edge case, AGENTS.md § Testing rule);
// the sweeps start from the real coverage fixtures.

/** Deterministic PRNG so a failing seed reproduces. */
function mulberry32(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function isTyped(error: unknown): boolean {
  return error instanceof AaredError || isAbort(error);
}

/** The code a read ends with, or "ok". */
function codeOf(run: () => unknown): string {
  try {
    run();
    return "ok";
  } catch (error) {
    return error instanceof AaredError ? error.code : `untyped ${String(error)}`;
  }
}

describe("hostile XML", () => {
  test("a character reference past U+10FFFF is a damaged document, not a viewer failure", () => {
    for (const reference of ["&#x110000;", "&#99999999;", `&#${"9".repeat(400)};`]) {
      expect(codeOf(() => parseXml(`<a>${reference}</a>`))).toBe("invalid-document");
      expect(codeOf(() => parseXml(`<a b="${reference}"/>`))).toBe("invalid-document");
    }
    expect(parseXml("<a>&#x10FFFF;&#65;</a>").text).toBe(`${String.fromCodePoint(0x10ffff)}A`);
  });

  test("elements nested absurdly deep end in a typed error instead of overflowing the stack", () => {
    const deep = (levels: number): string => `${"<a>".repeat(levels)}${"</a>".repeat(levels)}`;
    expect(codeOf(() => parseXml(deep(100_000)))).toBe("invalid-document");
    expect(codeOf(() => parseXml(deep(300)))).toBe("ok");
  });
});

describe("hostile ZIP structure", () => {
  const one = (): Uint8Array => storedZip({ "a.xml": "<a/>" });
  /** Offsets in the one-entry archive built above: the End Of Central Directory record is the last 22 bytes. */
  const eocd = (zip: Uint8Array): number => zip.byteLength - 22;
  const centralStart = (zip: Uint8Array): number => new DataView(zip.buffer, zip.byteOffset).getUint32(eocd(zip) + 16, true);

  test("a count larger than the directory holds is a corrupt archive, not a RangeError", () => {
    const zip = one();
    new DataView(zip.buffer).setUint16(eocd(zip) + 10, 3, true);
    expect(codeOf(() => readEntries(zip))).toBe("invalid-document");
  });

  test("a name length that runs past the directory is a corrupt archive", () => {
    const zip = one();
    new DataView(zip.buffer).setUint16(centralStart(zip) + 28, 60_000, true);
    expect(codeOf(() => readEntries(zip))).toBe("invalid-document");
  });

  test("ZIP64 markers are unsupported-feature and encrypted entries are encrypted-document", () => {
    const zip64 = one();
    new DataView(zip64.buffer).setUint32(eocd(zip64) + 16, 0xffff_ffff, true);
    expect(() => readEntries(zip64)).toThrow("ZIP64");
    expect(codeOf(() => readEntries(zip64))).toBe("unsupported-feature");
    const encrypted = one();
    new DataView(encrypted.buffer).setUint16(centralStart(encrypted) + 8, 0x0001, true);
    expect(() => readEntries(encrypted)).toThrow("encrypted");
    expect(codeOf(() => readEntries(encrypted))).toBe("encrypted-document");
  });

  test("an entry that declares more than the size limit is refused without reading it", async () => {
    const zip = one();
    new DataView(zip.buffer).setUint32(centralStart(zip) + 24, MAX_ZIP_ENTRY_BYTES + 1, true);
    const [entry] = readEntries(zip);
    await expect(getEntryData(zip, entry!)).rejects.toThrow("preview limit");
  });

  test("a header that lies about the inflated size still yields the real bytes", async () => {
    const payload = new TextEncoder().encode("<root>" + "x".repeat(100_000) + "</root>");
    const deflated = new Uint8Array(await new Response(new Blob([payload]).stream().pipeThrough(new CompressionStream("deflate-raw"))).arrayBuffer());
    const name = new TextEncoder().encode("a.xml");
    const local = new Uint8Array(30 + name.length + deflated.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(26, name.length, true);
    local.set(name, 30);
    local.set(deflated, 30 + name.length);
    const central = new Uint8Array(46 + name.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(10, 8, true);
    cv.setUint32(20, deflated.length, true);
    cv.setUint32(24, 10, true); // claims 10 bytes; the stream holds 100,013
    cv.setUint16(28, name.length, true);
    central.set(name, 46);
    const end = new Uint8Array(22);
    const ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, 1, true);
    ev.setUint16(10, 1, true);
    ev.setUint32(12, central.length, true);
    ev.setUint32(16, local.length, true);
    const zip = new Uint8Array(local.length + central.length + end.length);
    zip.set(local, 0);
    zip.set(central, local.length);
    zip.set(end, local.length + central.length);
    const [entry] = readEntries(zip);
    expect(await getEntryData(zip, entry!)).toEqual(payload);
  });
});

type Parse = (pkg: OpcPackage) => Promise<unknown>;

const FORMATS: Array<{ file: string; parse: Parse; flips: number; mutations: number }> = [
  {
    file: "xlsx-coverage.xlsx",
    flips: 60,
    mutations: 40,
    parse: async (pkg) => {
      const workbook = await readWorkbook(pkg);
      for (let index = 0; index < workbook.sheets.length; index++) await parseXlsxSheet(pkg, workbook, index);
    },
  },
  {
    file: "post-production-tracker.xlsx",
    flips: 40,
    mutations: 30,
    parse: async (pkg) => {
      const workbook = await readWorkbook(pkg);
      for (let index = 0; index < workbook.sheets.length; index++) await parseXlsxSheet(pkg, workbook, index);
    },
  },
  {
    file: "pptx-coverage.pptx",
    flips: 20,
    mutations: 15,
    parse: async (pkg) => {
      const deck = await indexPptx(pkg);
      for (let index = 0; index < deck.slideParts.length; index++) await parsePptxSlide(pkg, deck, index);
    },
  },
  {
    file: "docx-coverage.docx",
    flips: 20,
    mutations: 15,
    parse: async (pkg) => {
      await parseDocxSections(pkg, await indexDocx(pkg));
    },
  },
];

/** Run one parse and report how it ended, so a failure names the file, seed and iteration. */
async function outcome(label: string, bytes: Uint8Array, parse: Parse): Promise<"ok" | "typed"> {
  try {
    await parse(await OpcPackage.open(bytes.slice().buffer as ArrayBuffer));
    return "ok";
  } catch (error) {
    if (isTyped(error)) return "typed";
    const detail = error instanceof Error ? `${error.name}: ${error.message}\n${error.stack ?? ""}` : String(error);
    throw new Error(`${label}: untyped failure ${detail}`);
  }
}

describe("seeded corruption sweeps over the coverage fixtures", () => {
  for (const { file, parse, flips, mutations } of FORMATS) {
    test(`${file}: random byte flips end in a result or a typed error`, async () => {
      const original = new Uint8Array(await Bun.file(`${FIXTURES}/${file}`).arrayBuffer());
      const random = mulberry32(0x5eed + file.length);
      let typed = 0;
      for (let i = 0; i < flips; i++) {
        const bytes = original.slice();
        for (let k = 0, n = 1 + Math.floor(random() * 3); k < n; k++) {
          bytes[Math.floor(random() * bytes.length)] = Math.floor(random() * 256);
        }
        if ((await outcome(`${file} flip #${i}`, bytes, parse)) === "typed") typed++;
      }
      // Most flips land in compressed data; the sweep is only useful if it does reject some.
      expect(typed).toBeGreaterThan(0);
    }, 60_000);

    test(`${file}: damaged XML text (cut, deleted span, swapped character) ends in a result or a typed error`, async () => {
      const source = await OpcPackage.open(await Bun.file(`${FIXTURES}/${file}`).arrayBuffer());
      const parts: Record<string, Uint8Array> = {};
      const xml: string[] = [];
      for (const name of source.partNames()) {
        parts[name] = await source.readPart(name);
        if (/\.(xml|rels)$/.test(name)) xml.push(name);
      }
      const random = mulberry32(0xc0de + file.length);
      const decoder = new TextDecoder();
      const encoder = new TextEncoder();
      for (let i = 0; i < mutations; i++) {
        const name = xml[Math.floor(random() * xml.length)]!;
        const text = decoder.decode(parts[name]);
        const at = Math.floor(random() * text.length);
        const kind = Math.floor(random() * 3);
        const mutated =
          kind === 0
            ? text.slice(0, at)
            : kind === 1
              ? text.slice(0, at) + text.slice(at + 1 + Math.floor(random() * 200))
              : text.slice(0, at) + String.fromCharCode(33 + Math.floor(random() * 90)) + text.slice(at + 1);
        const bytes = storedZip({ ...parts, [name]: encoder.encode(mutated) });
        await outcome(`${file} mutation #${i} (${["cut", "delete", "swap"][kind]} in ${name} at ${at})`, bytes, parse);
      }
    }, 60_000);
  }
});

async function timed<T>(run: () => Promise<T>): Promise<{ value: T; ms: number }> {
  const start = performance.now();
  const value = await run();
  return { value, ms: performance.now() - start };
}

// Counts a file claims that drive a loop or an allocation. Each of these was
// measured before it was bounded: a 60-byte <col> took 21 s and then ran the
// Worker out of memory; a whole-sheet merge ran the page out of memory; a
// table span of 3000 x 3000 froze it for seconds. (The PPTX span clamp is
// pinned in pptx-fidelity.test.ts beside the table parser's other cases.)
describe("hostile counts", () => {
  const SHEET = "xl/worksheets/sheet1.xml";
  const parseFirstSheet = async (pkg: OpcPackage) => parseXlsxSheet(pkg, await readWorkbook(pkg), 0);
  const insertCols = (cols: string) => (xml: string) =>
    xml.replace(/<cols>[\s\S]*?<\/cols>/, "").replace("<sheetData", `<cols>${cols}</cols><sheetData`);
  const notice = (sheet: Awaited<ReturnType<typeof parseXlsxSheet>>, what: string) =>
    sheet.unsupported.some((entry) => entry.detail.includes(what));

  test("a <col> with a negative min, a zero min or a min above its max is skipped, not looped over", async () => {
    const pkg = await withPart(
      "xlsx-coverage.xlsx",
      SHEET,
      insertCols(`<col min="-100000000" max="1" width="20" customWidth="1"/><col min="0" max="3" width="20"/><col min="9" max="2" width="20"/>`),
    );
    const { value: sheet, ms } = await timed(() => parseFirstSheet(pkg));
    expect(ms).toBeLessThan(1000);
    expect(notice(sheet, "column range")).toBe(true);
  });

  test("tens of thousands of full-width <col> ranges stop at one pass over the grid", async () => {
    const pkg = await withPart("xlsx-coverage.xlsx", SHEET, insertCols(`<col min="1" max="16384" width="9"/>`.repeat(50_000)));
    const { value: sheet, ms } = await timed(() => parseFirstSheet(pkg));
    expect(ms).toBeLessThan(3000);
    expect(notice(sheet, "column range")).toBe(true);
  });

  test("a merge over the whole sheet is cut to the sheet's extent, not indexed cell by cell", async () => {
    const pkg = await withPart("xlsx-coverage.xlsx", SHEET, (xml) =>
      xml.replace(/<mergeCells[\s\S]*?<\/mergeCells>/, "").replace("</sheetData>", `</sheetData><mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>`),
    );
    const { value: sheet, ms } = await timed(() => parseFirstSheet(pkg));
    expect(ms).toBeLessThan(1000);
    expect(sheet.merges).toHaveLength(1);
    expect(sheet.merges[0]!.end).toEqual({ row: sheet.extent.rows - 1, col: sheet.extent.cols - 1 });
  });

  test("with no dimension, a whole-sheet merge does not size the sheet: it is cut to the content", async () => {
    const pkg = await withPart("xlsx-coverage.xlsx", SHEET, (xml) =>
      xml
        .replace(/<dimension[^>]*\/>/, "")
        .replace(/<mergeCells[\s\S]*?<\/mergeCells>/, "")
        .replace("</sheetData>", `</sheetData><mergeCells count="1"><mergeCell ref="A1:XFD1048576"/></mergeCells>`),
    );
    const sheet = await parseFirstSheet(pkg);
    expect(sheet.extent.rows).toBeLessThan(1000);
    expect(sheet.extent.cols).toBeLessThan(100);
    expect(sheet.merges).toEqual([{ start: { row: 0, col: 0 }, end: { row: sheet.extent.rows - 1, col: sheet.extent.cols - 1 } }]);
  });

  test("a DOCX gridSpan past the table's columns is cut to the columns that are left", async () => {
    const pkg = await withPart("docx-coverage.docx", "word/document.xml", (xml) =>
      xml.replace(/<w:gridSpan w:val="\d+"\/>/, `<w:gridSpan w:val="99999999"/>`),
    );
    const { value: sections, ms } = await timed(async () => parseDocxSections(pkg, await indexDocx(pkg)));
    expect(ms).toBeLessThan(2000);
    const spans: number[] = [];
    for (const section of sections) {
      for (const node of section.root as unknown as Array<{ type?: string; rows?: Array<{ cells: Array<{ columnSpan?: number }> }> }>) {
        if (node.type !== "table" || node.rows === undefined) continue;
        for (const row of node.rows) for (const cell of row.cells) if (cell.columnSpan !== undefined) spans.push(cell.columnSpan);
      }
    }
    expect(spans.length).toBeGreaterThan(0);
    expect(Math.max(...spans)).toBeLessThan(100);
  });
});
