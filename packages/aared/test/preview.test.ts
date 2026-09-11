import { expect, test } from "bun:test";
import { AaredError, isAaredError } from "../src/core/index.js";
import { detect, preview } from "../src/preview/index.js";

const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-xlsx-rtl-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-docx-bidi-coverage.docx`;

/** `preview()` paints, so these headless tests only reach paths that reject
 * before any DOM is touched (dispatch, validation, corrupt input). The stub
 * target only needs `replaceChildren` for `destroy()` cleanup. */
function stubTarget(): HTMLElement {
  return { replaceChildren(): void {} } as unknown as HTMLElement;
}

async function rejection(promise: Promise<unknown>): Promise<Error> {
  try {
    await promise;
  } catch (error) {
    return error instanceof Error ? error : new Error(String(error));
  }
  throw new Error("expected the request to reject");
}

test("detect identifies a PPTX fixture", async () => {
  expect(await detect(await Bun.file(PPTX_FIXTURE).arrayBuffer())).toBe("pptx");
});

test("detect identifies an XLSX fixture", async () => {
  expect(await detect(await Bun.file(XLSX_FIXTURE).arrayBuffer())).toBe("xlsx");
});

test("detect identifies a DOCX fixture", async () => {
  expect(await detect(await Bun.file(DOCX_FIXTURE).arrayBuffer())).toBe("docx");
});

test("detect maps corrupt archives to invalid-document errors", async () => {
  const error = await rejection(detect(new ArrayBuffer(8)));
  expect(error).toBeInstanceOf(AaredError);
  expect((error as AaredError).code).toBe("invalid-document");
});

test("detect honors an aborted signal", async () => {
  const controller = new AbortController();
  controller.abort();
  const error = await rejection(
    detect(await Bun.file(PPTX_FIXTURE).arrayBuffer(), { signal: controller.signal }),
  );
  expect(error.name).toBe("AbortError");
});

test("preview with an asserted wrong type rejects before painting", async () => {
  const error = await rejection(
    preview(await Bun.file(PPTX_FIXTURE).arrayBuffer(), stubTarget(), { type: "xlsx" }),
  );
  expect(error).toBeInstanceOf(AaredError);
  expect(isAaredError(error) && error.code).toBe("unsupported-format");
  expect(error.message).toContain("aared/pptx");
});

test("preview maps corrupt archives to invalid-document errors", async () => {
  const error = await rejection(preview(new ArrayBuffer(8), stubTarget()));
  expect(error).toBeInstanceOf(AaredError);
  expect((error as AaredError).code).toBe("invalid-document");
});

test("preview rejects an out-of-range initial index without painting", async () => {
  const error = await rejection(
    preview(await Bun.file(PPTX_FIXTURE).arrayBuffer(), stubTarget(), { index: 99 }),
  );
  expect(error).toBeInstanceOf(AaredError);
  expect((error as AaredError).code).toBe("not-found");
  expect(error.message).toContain("99");
});

const CRC_TABLE = new Uint32Array(256).map((_, n) => {
  let c = n;
  for (let bit = 0; bit < 8; bit++) c = c & 1 ? 0xedb8_8320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});

function crc32(bytes: Uint8Array): number {
  let crc = 0xffff_ffff;
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffff_ffff) >>> 0;
}

/** Minimal stored-entry (uncompressed) ZIP: local headers + central directory + EOCD. */
function storedZip(parts: Array<[name: string, text: string]>): ArrayBuffer {
  const encoder = new TextEncoder();
  const entries = parts.map(([name, text]) => ({ name: encoder.encode(name), data: encoder.encode(text) }));
  const local = new Uint8Array(entries.reduce((sum, e) => sum + 30 + e.name.length + e.data.length, 0));
  const central = new Uint8Array(entries.reduce((sum, e) => sum + 46 + e.name.length, 0));
  const localView = new DataView(local.buffer);
  const centralView = new DataView(central.buffer);
  let localPos = 0;
  let centralPos = 0;
  for (const entry of entries) {
    const crc = crc32(entry.data);
    localView.setUint32(localPos, 0x0403_4b50, true);
    localView.setUint16(localPos + 4, 20, true);
    localView.setUint16(localPos + 8, 0, true);
    localView.setUint32(localPos + 14, crc, true);
    localView.setUint32(localPos + 18, entry.data.length, true);
    localView.setUint32(localPos + 22, entry.data.length, true);
    localView.setUint16(localPos + 26, entry.name.length, true);
    local.set(entry.name, localPos + 30);
    local.set(entry.data, localPos + 30 + entry.name.length);
    centralView.setUint32(centralPos, 0x0201_4b50, true);
    centralView.setUint16(centralPos + 4, 20, true);
    centralView.setUint16(centralPos + 6, 20, true);
    centralView.setUint16(centralPos + 10, 0, true);
    centralView.setUint32(centralPos + 16, crc, true);
    centralView.setUint32(centralPos + 20, entry.data.length, true);
    centralView.setUint32(centralPos + 24, entry.data.length, true);
    centralView.setUint16(centralPos + 28, entry.name.length, true);
    centralView.setUint32(centralPos + 42, localPos, true);
    central.set(entry.name, centralPos + 46);
    localPos += 30 + entry.name.length + entry.data.length;
    centralPos += 46 + entry.name.length;
  }
  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, 0x0605_4b50, true);
  eocdView.setUint16(8, entries.length, true);
  eocdView.setUint16(10, entries.length, true);
  eocdView.setUint32(12, central.length, true);
  eocdView.setUint32(16, local.length, true);
  const out = new Uint8Array(local.length + central.length + eocd.length);
  out.set(local, 0);
  out.set(central, local.length);
  out.set(eocd, local.length + central.length);
  return out.buffer;
}

/** A structurally valid PPTX whose presentation part lists no slides. */
function emptyDeckBytes(): ArrayBuffer {
  return storedZip([
    [
      "[Content_Types].xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`,
    ],
    [
      "_rels/.rels",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`,
    ],
    [
      "ppt/presentation.xml",
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`,
    ],
  ]);
}

test("preview maps a valid but empty deck to an invalid-document error", async () => {
  const error = await rejection(preview(emptyDeckBytes(), stubTarget()));
  expect(error).toBeInstanceOf(AaredError);
  expect((error as AaredError).code).toBe("invalid-document");
  expect(error.message).toContain("no slides");
});
