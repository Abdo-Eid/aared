import { expect, test } from "bun:test";
import { open } from "../src/xlsx/index.js";

const XLSX_FIXTURE = `${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`;

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

test("search names each cell as { address, row, col } in its sheet's { index, title }", async () => {
  const workbook = await open(await Bun.file(XLSX_FIXTURE).arrayBuffer());
  try {
    const matches = await workbook.sheets.get("Summary").search("budget");
    expect(matches.length).toBeGreaterThan(0);
    for (const match of matches) {
      expect(match.address).toMatch(/^[A-Z]+[0-9]+$/);
      expect(match.sheet.title).toBe(workbook.sheets.get(match.sheet.index).info.title);
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
    const promise = workbook.sheets.get("Summary").search("budget", { signal: controller.signal });
    controller.abort();
    const error = await promise.catch((e) => e);
    expect(error.name).toBe("AbortError");
  } finally {
    workbook.close();
  }
});
