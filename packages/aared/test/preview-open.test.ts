import { expect, test } from "bun:test";
import { detect, preview } from "../src/preview/index.js";
import { openViewer } from "../src/embed-viewer.js";
import { WorkerSession } from "../src/core/session.js";
import type { ViewerRoute } from "../src/format-route.js";
import { rejection, within } from "./harness.js";

const FIXTURE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

/** Observe the real Worker boundary, not an in-memory simulated parser. */
function recordingWorker() {
  const requests: string[] = [];
  const transfers: number[] = [];
  let terminations = 0;
  const factory = (): Worker => {
    const worker = new Worker(new URL("../dist/core/worker.mjs", import.meta.url), { type: "module" });
    const post = worker.postMessage.bind(worker);
    worker.postMessage = (message: unknown, options?: Transferable[] | StructuredSerializeOptions): void => {
      requests.push((message as { type: string }).type);
      transfers.push(Array.isArray(options) ? options.length : options?.transfer?.length ?? 0);
      if (Array.isArray(options)) post(message, options);
      else post(message, options);
    };
    const terminate = worker.terminate.bind(worker);
    worker.terminate = (): void => { terminations++; terminate(); };
    return worker;
  };
  return { factory, requests, transfers, get terminations() { return terminations; } };
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}

/** These lifecycle tests deliberately stop before renderer DOM access. */
const target = (): HTMLElement => ({ replaceChildren() {} }) as unknown as HTMLElement;

test("auto-preview transfers one archive, never detects then reopens, and preserves the caller's buffer", async () => {
  const recorder = recordingWorker();
  const bytes = await Bun.file(FIXTURE).arrayBuffer();
  const length = bytes.byteLength;
  const error = await rejection(preview({ source: bytes, target: target(), index: 999, worker: { create: recorder.factory } }));
  expect(error).toMatchObject({ code: "not-found" });
  expect(bytes.byteLength).toBe(length);
  expect(new Uint8Array(bytes, 0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
  expect(recorder.requests).toEqual(["open", "close"]);
  expect(recorder.transfers).toEqual([1, 0]);
  expect(recorder.terminations).toBe(1);
});

test("preview and detect transfer their internally owned Blob bytes without another copy", async () => {
  for (const operation of ["preview", "detect"] as const) {
    const recorder = recordingWorker();
    const owned = await Bun.file(FIXTURE).arrayBuffer();
    const source = { size: owned.byteLength, arrayBuffer: async () => owned } as Blob;
    if (operation === "detect") expect(await detect(source, { worker: { create: recorder.factory } })).toBe("pptx");
    else expect(await rejection(preview({ source, target: target(), index: 999, worker: { create: recorder.factory } })))
      .toMatchObject({ code: "not-found" });
    expect(owned.byteLength).toBe(0);
    expect(recorder.requests).toEqual([operation === "detect" ? "detect" : "open", "close"]);
    expect(recorder.transfers).toEqual([1, 0]);
    expect(recorder.terminations).toBe(1);
  }
});

test("asserted mismatches close the opened session without loading a route", async () => {
  const recorder = recordingWorker();
  let routed = false;
  const error = await rejection(openViewer({
    scope: "aared/preview", source: await Bun.file(FIXTURE).arrayBuffer(), target: target(),
    options: { type: "xlsx", worker: { create: recorder.factory } },
    route: async () => { routed = true; throw new Error("must not import"); },
  }));
  expect(error).toMatchObject({ code: "unsupported-format" });
  expect(error.message).toContain("aared/xlsx:");
  expect(error.message).toContain("aared/pptx");
  expect(routed).toBe(false);
  expect(recorder.requests).toEqual(["open", "close"]);
});

test("cancellation closes and rejects immediately while a lazy route is still loading", async () => {
  const recorder = recordingWorker();
  const loading = deferred<void>();
  const route = deferred<ViewerRoute>();
  const controller = new AbortController();
  const opening = openViewer({
    scope: "aared/preview", source: await Bun.file(FIXTURE).arrayBuffer(), target: target(),
    options: { signal: controller.signal, worker: { create: recorder.factory } },
    route: (format) => { expect(format).toBe("pptx"); loading.resolve(); return route.promise; },
  });
  await loading.promise;
  controller.abort();
  expect(recorder.requests).toEqual(["open", "close"]);
  expect(await within("cancelling a pending route", rejection(opening), 1000)).toMatchObject({ code: "aborted" });
  expect(recorder.terminations).toBe(1);
  // A failed chunk arriving after cancellation is observed, not unhandled.
  route.reject(new Error("late chunk failure"));
  await Promise.resolve();
});

test("a failed lazy route closes the session before rejecting", async () => {
  const recorder = recordingWorker();
  const failure = new Error("chunk failed");
  expect(await rejection(openViewer({
    scope: "aared/preview", source: await Bun.file(FIXTURE).arrayBuffer(), target: target(),
    options: { worker: { create: recorder.factory } }, route: async () => { throw failure; },
  }))).toBe(failure);
  expect(recorder.requests).toEqual(["open", "close"]);
  expect(recorder.terminations).toBe(1);
});

test("a failed session adoption closes before rejecting and never renders", async () => {
  const recorder = recordingWorker();
  const failure = new Error("document initialization failed");
  expect(await rejection(openViewer({
    scope: "aared/preview", source: await Bun.file(FIXTURE).arrayBuffer(), target: target(),
    options: { worker: { create: recorder.factory } },
    route: { format: "pptx", adopt: async () => { throw failure; }, content: () => { throw new Error("must not expose content"); }, render: async () => { throw new Error("must not render"); } },
  }))).toBe(failure);
  expect(recorder.requests).toEqual(["open", "close"]);
  expect(recorder.terminations).toBe(1);
});

test("cancellation during adoption closes the session and observes a late adapter failure", async () => {
  const recorder = recordingWorker();
  const adopting = deferred<void>();
  const adopted = deferred<never>();
  const controller = new AbortController();
  const opening = openViewer({
    scope: "aared/preview", source: await Bun.file(FIXTURE).arrayBuffer(), target: target(),
    options: { signal: controller.signal, worker: { create: recorder.factory } },
    route: { format: "pptx", adopt: () => { adopting.resolve(); return adopted.promise; }, content: () => { throw new Error("must not expose content"); }, render: async () => { throw new Error("must not render"); } },
  });
  await adopting.promise;
  controller.abort();
  expect(await within("cancelling adoption", rejection(opening), 1000)).toMatchObject({ code: "aborted" });
  expect(recorder.requests).toEqual(["open", "close"]);
  adopted.reject(new Error("late adoption failure"));
  await Promise.resolve();
});

test("an adapter that aborts synchronously still has its rejected promise observed", async () => {
  const recorder = recordingWorker();
  const controller = new AbortController();
  const error = await rejection(openViewer({
    scope: "aared/preview", source: await Bun.file(FIXTURE).arrayBuffer(), target: target(),
    options: { signal: controller.signal, worker: { create: recorder.factory } },
    route: () => { controller.abort(); return Promise.reject(new Error("cancelled import")); },
  }));
  expect(error).toMatchObject({ code: "aborted" });
  expect(recorder.requests).toEqual(["open", "close"]);
});

test("owned-buffer open rejects an already aborted signal before transferring", async () => {
  const recorder = recordingWorker();
  const owned = await Bun.file(FIXTURE).arrayBuffer();
  const length = owned.byteLength;
  expect(await rejection(WorkerSession.openBytes("aared/preview", owned, {
    worker: { create: recorder.factory }, signal: AbortSignal.abort(),
  }))).toMatchObject({ code: "aborted" });
  expect(owned.byteLength).toBe(length);
  expect(recorder.requests).toEqual([]);
});

test("PPTX adoption installs deferred geometry and resolves the same fonts and models as open", async () => {
  const { open, PptxDocument } = await import("../src/pptx/index.js");
  const { route } = await import("../src/pptx/viewer.js");
  const bytes = await Bun.file(FIXTURE).arrayBuffer();
  const ordinaryFonts = new Set<string>();
  const adoptedFonts = new Set<string>();
  const ordinary = await open(bytes, { fontResolver: (family) => { ordinaryFonts.add(family); return undefined; } });
  const session = await WorkerSession.open("aared/preview", bytes, {
    fontResolver: (family) => { adoptedFonts.add(family); return undefined; },
  });
  try {
    const adopted = await route.adopt(session, undefined) as InstanceType<typeof PptxDocument>;
    expect(adopted.info).toEqual(ordinary.info);
    let found = false;
    for (const slide of adopted.slides.list()) {
      const model = await slide.read();
      const reference = await ordinary.slides.get(slide.info.index).read();
      expect(model).toEqual(reference);
      const arrow = model.elements.find((element) => element.kind === "shape" && element.geometry?.preset === "upArrow");
      if (arrow?.kind !== "shape") continue;
      expect(arrow.geometry?.deferred).toBeUndefined();
      expect(arrow.geometry?.fallback).toBeUndefined();
      expect(arrow.geometry?.path.length).toBeGreaterThan(0);
      found = true;
      break;
    }
    expect(found).toBe(true);
    expect([...adoptedFonts].sort()).toEqual([...ordinaryFonts].sort());
    expect(adoptedFonts.size).toBeGreaterThan(0);
  } finally { ordinary.close(); session.close(); }
});

test("DOCX adoption shares eager section loading, headings, search and font resolution with open", async () => {
  const { open, DocxDocument } = await import("../src/docx/index.js");
  const { route } = await import("../src/docx/viewer.js");
  const bytes = await Bun.file(`${import.meta.dir}/../../../fixtures/docx-coverage.docx`).arrayBuffer();
  const ordinaryFonts = new Set<string>();
  const adoptedFonts = new Set<string>();
  const ordinary = await open(bytes, { fontResolver: (family) => { ordinaryFonts.add(family); return undefined; } });
  const session = await WorkerSession.open("aared/preview", bytes, {
    fontResolver: (family) => { adoptedFonts.add(family); return undefined; },
  });
  try {
    const adopted = await route.adopt(session, undefined) as InstanceType<typeof DocxDocument>;
    expect(adopted.sections.list()).toEqual(ordinary.sections.list());
    expect(adopted.headings.list()).toEqual(ordinary.headings.list());
    expect(await adopted.search("Project")).toEqual(await ordinary.search("Project"));
    expect([...adoptedFonts].sort()).toEqual([...ordinaryFonts].sort());
    expect(adoptedFonts.size).toBeGreaterThan(0);
  } finally { ordinary.close(); session.close(); }
});

test("XLSX adoption preserves sheet metadata and cached-value reads", async () => {
  const { open, XlsxDocument } = await import("../src/xlsx/index.js");
  const { route } = await import("../src/xlsx/viewer.js");
  const bytes = await Bun.file(`${import.meta.dir}/../../../fixtures/xlsx-coverage.xlsx`).arrayBuffer();
  const ordinary = await open(bytes);
  const session = await WorkerSession.open("aared/preview", bytes);
  try {
    const adopted = await route.adopt(session, undefined) as InstanceType<typeof XlsxDocument>;
    expect(adopted.info).toEqual(ordinary.info);
    expect(adopted.sheets.list().map((sheet) => sheet.info)).toEqual(ordinary.sheets.list().map((sheet) => sheet.info));
    expect(await adopted.sheets.get("Summary").read()).toEqual(await ordinary.sheets.get("Summary").read());
  } finally { ordinary.close(); session.close(); }
});

test("cancelling DOCX during eager font resolution closes the session before any paint", async () => {
  const { route } = await import("../src/docx/viewer.js");
  const recorder = recordingWorker();
  const requested = deferred<void>();
  const font = deferred<FontFace | undefined>();
  const controller = new AbortController();
  const opening = openViewer({
    scope: "aared/docx", source: await Bun.file(`${import.meta.dir}/../../../fixtures/docx-coverage.docx`).arrayBuffer(),
    target: target(), route,
    options: {
      signal: controller.signal, worker: { create: recorder.factory },
      fontResolver: () => { requested.resolve(); return font.promise; },
    },
  });
  await requested.promise;
  controller.abort();
  expect(await within("cancelling DOCX font initialization", rejection(opening), 1000)).toMatchObject({ code: "aborted" });
  expect(recorder.requests).toEqual(["open", "unit", "close"]);
  expect(recorder.terminations).toBe(1);
  font.resolve(undefined);
});
