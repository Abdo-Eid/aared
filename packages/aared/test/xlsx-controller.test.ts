import { expect, test } from "bun:test";
import { open } from "../src/xlsx/index.js";

const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;

test("sheetStates are exposed in document info", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.info.sheetStates).toBeDefined();
    expect(workbook.info.sheetStates).toHaveLength(6);
    for (const state of workbook.info.sheetStates!) {
      expect(state).toBe("visible");
    }
  } finally {
    workbook.close();
  }
});

test("count matches sheet collection length", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.info.count).toBe(workbook.sheets.list().length);
  } finally {
    workbook.close();
  }
});

test("search with abort signal rejects", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const controller = new AbortController();
    const promise = workbook.sheets.get("Summary").search("budget", { signal: controller.signal });
    controller.abort();
    const error = await promise.catch((e) => e);
    expect(error.name).toBe("AbortError");
  } finally {
    workbook.close();
  }
});
