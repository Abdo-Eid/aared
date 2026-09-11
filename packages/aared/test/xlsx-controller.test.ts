import { expect, test } from "bun:test";
import { open } from "../src/xlsx/index.js";

const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/arabic-xlsx-rtl-coverage.xlsx`;

test("sheetStates are exposed in document info", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    expect(workbook.info.sheetStates).toBeDefined();
    expect(workbook.info.sheetStates).toHaveLength(4);
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

test("search returns cell addresses in A1 notation", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const matches = await workbook.search("budget");
    for (const match of matches) {
      expect(match.cell).toMatch(/^[A-Z]+[0-9]+$/);
      expect(match.sheet).toBeTruthy();
      expect(match.text.length).toBeGreaterThan(0);
    }
  } finally {
    workbook.close();
  }
});

test("search with abort signal rejects", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const controller = new AbortController();
    const promise = workbook.search("budget", { signal: controller.signal });
    controller.abort();
    const error = await promise.catch((e) => e);
    expect(error.name).toBe("AbortError");
  } finally {
    workbook.close();
  }
});
