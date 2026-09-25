import { describe, expect, test } from "bun:test";
import { DEFAULT_ZOOM, announceUnit, followSheet, forwardRenderEvents, unitRef } from "../src/format-route.js";
import type { AnyDocument } from "../src/format-route.js";

/** Collect the events a target raises, as `[type, detail]`. */
function record(target: EventTarget, types: string[]): Array<[string, unknown]> {
  const seen: Array<[string, unknown]> = [];
  for (const type of types) target.addEventListener(type, (event) => seen.push([type, (event as CustomEvent).detail]));
  return seen;
}

const ALL = ["indexchange", "slidechange", "sheetchange"];
const deckOf = (titles: string[]) => ({ info: { titles } }) as unknown as AnyDocument;
const deck = deckOf(["Slide 1", "Slide 2", "Slide 3"]);
const workbook = deckOf(["Summary", "Log"]);
const document = deckOf(["Coverage document"]);

describe("announcing a unit", () => {
  test("a slide raises indexchange then slidechange, naming the slide as { index, title }", () => {
    const target = new EventTarget();
    const seen = record(target, ALL);
    announceUnit(target, "pptx", deck, 2);
    expect(seen).toEqual([["indexchange", { index: 2 }], ["slidechange", { index: 2, title: "Slide 3" }]]);
  });

  test("a sheet raises indexchange then sheetchange, naming the sheet the same way", () => {
    const target = new EventTarget();
    const seen = record(target, ALL);
    announceUnit(target, "xlsx", workbook, 1);
    expect(seen).toEqual([["indexchange", { index: 1 }], ["sheetchange", { index: 1, title: "Log" }]]);
  });

  test("a document raises only indexchange", () => {
    const target = new EventTarget();
    const seen = record(target, ALL);
    announceUnit(target, "docx", document, 0);
    expect(seen).toEqual([["indexchange", { index: 0 }]]);
  });

  test("a unit reference is the goTo key plus the label from info.titles", () => {
    expect(unitRef(workbook, 0)).toEqual({ index: 0, title: "Summary" });
  });
});

describe("default zoom", () => {
  test("a slide opens whole, a sheet at 100% like Excel, a document to the width", () => {
    expect(DEFAULT_ZOOM).toEqual({ pptx: "fit-page", xlsx: 1, docx: "fit-width" });
  });
});

describe("forwarding render events", () => {
  const FORWARDED = ["zoomchange", "linkactivate", "cellactivate", "selectionchange"];

  test("re-raises each forwarded event on the facade with its detail", () => {
    const handle = new EventTarget();
    const facade = new EventTarget();
    const seen = record(facade, [...FORWARDED, "other"]);
    forwardRenderEvents(handle, facade, new AbortController().signal);
    for (const [index, type] of FORWARDED.entries()) handle.dispatchEvent(new CustomEvent(type, { detail: index }));
    expect(seen).toEqual(FORWARDED.map((type, index) => [type, index]));
  });

  test("leaves the handle's own sheetchange to the facade's unit announcement", () => {
    const handle = new EventTarget();
    const facade = new EventTarget();
    const seen = record(facade, ["sheetchange", "render"]);
    forwardRenderEvents(handle, facade, new AbortController().signal);
    handle.dispatchEvent(new CustomEvent("sheetchange", { detail: { index: 1, title: "Log" } }));
    handle.dispatchEvent(new CustomEvent("render"));
    expect(seen).toEqual([]);
  });

  test("stops forwarding once the render is replaced", () => {
    const handle = new EventTarget();
    const facade = new EventTarget();
    const seen = record(facade, ["zoomchange"]);
    const controller = new AbortController();
    forwardRenderEvents(handle, facade, controller.signal);
    handle.dispatchEvent(new CustomEvent("zoomchange", { detail: { zoom: 1 } }));
    controller.abort();
    handle.dispatchEvent(new CustomEvent("zoomchange", { detail: { zoom: 2 } }));
    expect(seen).toEqual([["zoomchange", { zoom: 1 }]]);
  });
});

describe("following the sheet tabs", () => {
  test("hands the facade the index a tab click switched to, until the render is replaced", () => {
    const handle = new EventTarget();
    const followed: number[] = [];
    const controller = new AbortController();
    followSheet(handle, controller.signal, (index) => followed.push(index));
    handle.dispatchEvent(new CustomEvent("sheetchange", { detail: { index: 2, title: "Notes" } }));
    controller.abort();
    handle.dispatchEvent(new CustomEvent("sheetchange", { detail: { index: 0, title: "Summary" } }));
    expect(followed).toEqual([2]);
  });
});
