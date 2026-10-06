import { expect, test } from "bun:test";
import { AaredError, isAaredError } from "../src/core/index.js";
import { detect, preview } from "../src/preview/index.js";
import { rejection } from "./harness.js";
import { storedZipFromTuples } from "./zip.js";

const PPTX_FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;
const DOCX_FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

/** `preview()` paints, so these headless tests only reach paths that reject
 * before any DOM is touched (dispatch, validation, corrupt input). The stub
 * target only needs `replaceChildren` for `destroy()` cleanup. */
function stubTarget(): HTMLElement {
  return { replaceChildren(): void {} } as unknown as HTMLElement;
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
    preview({ source: await Bun.file(PPTX_FIXTURE).arrayBuffer(), target: stubTarget(), type: "xlsx" }),
  );
  expect(error).toBeInstanceOf(AaredError);
  expect(isAaredError(error) && error.code).toBe("unsupported-format");
  expect(error.message).toContain("aared/pptx");
});

test("preview maps corrupt archives to invalid-document errors", async () => {
  const error = await rejection(preview({ source: new ArrayBuffer(8), target: stubTarget() }));
  expect(error).toBeInstanceOf(AaredError);
  expect((error as AaredError).code).toBe("invalid-document");
});

test("preview rejects an out-of-range initial index without painting", async () => {
  const error = await rejection(
    preview({ source: await Bun.file(PPTX_FIXTURE).arrayBuffer(), target: stubTarget(), index: 99 }),
  );
  expect(error).toBeInstanceOf(AaredError);
  expect((error as AaredError).code).toBe("not-found");
  expect(error.message).toContain("99");
});

/** A structurally valid PPTX whose presentation part lists no slides. */
function emptyDeckBytes(): ArrayBuffer {
  return storedZipFromTuples([
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
  const error = await rejection(preview({ source: emptyDeckBytes(), target: stubTarget() }));
  expect(error).toBeInstanceOf(AaredError);
  expect((error as AaredError).code).toBe("invalid-document");
  expect(error.message).toContain("no slides");
});
