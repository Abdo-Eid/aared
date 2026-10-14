import { expect, spyOn, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxLayout, parsePptxSlide } from "../src/core/pptx.js";
import { open, readSlideModel } from "../src/pptx/renderer.js";

const COVERAGE = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;
const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;
const LAYOUT = "ppt/slideLayouts/slideLayout2.xml";

type LayoutSource = { layout(id: string, options?: { signal?: AbortSignal }): Promise<unknown> };

function sessionOf(deck: object): LayoutSource {
  return (deck as { session: LayoutSource }).session;
}

test("layout graphics never include placeholders, and a layout id is set only when there is something to fetch", async () => {
  for (const path of [COVERAGE, PRODUCTION]) {
    const pkg = await OpcPackage.open(await Bun.file(path).arrayBuffer());
    const deck = await indexPptx(pkg);
    for (const [index] of deck.slideParts.entries()) {
      const slide = await parsePptxSlide(pkg, deck, index);
      if (slide.layoutId === undefined) continue;
      const layout = await parsePptxLayout(pkg, deck, slide.layoutId);
      expect(layout.graphics.length).toBeGreaterThan(0);
      for (const graphic of layout.graphics) {
        expect((graphic as { placeholder?: unknown }).placeholder).toBeUndefined();
      }
    }
  }
  // A layout with no graphics of its own gets no id, so no extra request.
  const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
  const deck = await indexPptx(pkg);
  expect((await parsePptxSlide(pkg, deck, 0)).layoutId).toBeUndefined();
});

test("a layout is parsed once per package", async () => {
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const deck = await indexPptx(pkg);
  const first = await parsePptxLayout(pkg, deck, LAYOUT);
  expect(first.id).toBe(LAYOUT);
  expect(await parsePptxLayout(pkg, deck, LAYOUT)).toBe(first);
});

test("concurrent readers share one layout request and one object", async () => {
  const deck = await open(await Bun.file(COVERAGE).arrayBuffer());
  try {
    const post = spyOn(Worker.prototype, "postMessage");
    try {
      const session = sessionOf(deck);
      const [a, b, c] = await Promise.all([session.layout(LAYOUT), session.layout(LAYOUT), session.layout(LAYOUT)]);
      expect(b).toBe(a);
      expect(c).toBe(a);
      // A later reader reuses it too.
      expect(await session.layout(LAYOUT)).toBe(a);
      const requests = post.mock.calls.filter(([message]) => (message as { type?: string }).type === "layout");
      expect(requests).toHaveLength(1);
    } finally {
      post.mockRestore();
    }
  } finally {
    deck.close();
  }
});

test("asking for something that is not a layout in this deck is a not-found error", async () => {
  const deck = await open(await Bun.file(COVERAGE).arrayBuffer());
  try {
    const session = sessionOf(deck);
    const error = await session.layout("ppt/slides/slide1.xml").catch((e: unknown) => e);
    expect((error as { code?: string }).code).toBe("not-found");
    // A failed fetch is not cached: a valid layout still loads afterwards.
    expect(await session.layout(LAYOUT)).toMatchObject({ id: LAYOUT });
  } finally {
    deck.close();
  }
});

test("search reads slides without fetching layouts", async () => {
  const deck = await open(await Bun.file(COVERAGE).arrayBuffer());
  try {
    const post = spyOn(Worker.prototype, "postMessage");
    try {
      await deck.search("الإنتاج");
      expect(post.mock.calls.some(([message]) => (message as { type?: string }).type === "layout")).toBe(false);
    } finally {
      post.mockRestore();
    }
  } finally {
    deck.close();
  }
});

test("a reader that aborts while the layout loads does not fail the next reader", async () => {
  const deck = await open(await Bun.file(COVERAGE).arrayBuffer());
  try {
    const session = sessionOf(deck);
    const controller = new AbortController();
    const aborted = session.layout(LAYOUT, { signal: controller.signal }).catch((error: unknown) => error);
    controller.abort();
    expect(((await aborted) as Error).name).toBe("AbortError");
    // The shared fetch was never tied to that caller's signal.
    expect(await session.layout(LAYOUT)).toMatchObject({ id: LAYOUT });
  } finally {
    deck.close();
  }
});

// Slides 15 to 18 of the coverage deck (`tools/probe/pptx-second-master`) sit on a second master that has a
// logo, a picture background and its own theme, with two layouts: A (shows the master's graphics) and B
// (`showMasterSp="0"`, hides them). Slide 15 and 16 are on A, 17 on B, and 18 is on A with its own
// `showMasterSp="0"`.
const SECOND_MASTER = { a1: 14, a2: 15, b: 16, hidden: 17 };

const namesOf = (graphics: Array<{ name: string }>): string[] => graphics.map((graphic) => graphic.name);

test("private slide model reads share one layout object", async () => {
  const deck = await open(await Bun.file(COVERAGE).arrayBuffer());
  try {
    const [first, second, other] = await Promise.all([SECOND_MASTER.a1, SECOND_MASTER.a2, SECOND_MASTER.b].map((index) => readSlideModel(deck.slides.get(index))));
    expect(first!.layout).toBeDefined();
    expect(second!.layout).toBe(first!.layout);
    expect(other!.layout).not.toBe(first!.layout);
  } finally {
    deck.close();
  }
});

test("master graphics come first, then the layout's; a layout with showMasterSp=0 drops the master's", async () => {
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const deck = await indexPptx(pkg);
  const [a, b] = await Promise.all([SECOND_MASTER.a1, SECOND_MASTER.b].map(async (index) => {
    const slide = await parsePptxSlide(pkg, deck, index);
    return parsePptxLayout(pkg, deck, slide.layoutId!);
  }));
  expect(namesOf(a!.graphics)).toEqual(["Master logo", "Layout A graphic"]);
  expect(namesOf(b!.graphics)).toEqual(["Layout B graphic"]);
});

test("slide showMasterSp=0 sets hideBackgroundGraphics, and only that slide", async () => {
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const deck = await indexPptx(pkg);
  const flags = await Promise.all([SECOND_MASTER.a1, SECOND_MASTER.a2, SECOND_MASTER.b, SECOND_MASTER.hidden].map(async (index) => (await parsePptxSlide(pkg, deck, index)).hideBackgroundGraphics));
  expect(flags).toEqual([undefined, undefined, undefined, true]);
  // The hidden slide still names its layout: the renderer, not the parser, leaves the graphics out.
  expect((await parsePptxSlide(pkg, deck, SECOND_MASTER.hidden)).layoutId).toBeDefined();
});

test("a picture background set on the master is inherited by every slide on it, as one shared image fetched once", async () => {
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const deck = await indexPptx(pkg);
  const backgrounds = await Promise.all([SECOND_MASTER.a1, SECOND_MASTER.a2, SECOND_MASTER.b, SECOND_MASTER.hidden].map(async (index) => (await parsePptxSlide(pkg, deck, index)).background));
  for (const background of backgrounds) expect(background).toMatchObject({ kind: "picture", contentType: "image/png", shared: true });
  const ids = new Set(backgrounds.map((background) => (background as { mediaId: string }).mediaId));
  expect(ids.size).toBe(1);
  // The first master's slides do not inherit it.
  expect((await parsePptxSlide(pkg, deck, 0)).background).not.toMatchObject({ kind: "picture" });
  // Shared media goes through one request per session, whoever asks.
  const live = await open(await Bun.file(COVERAGE).arrayBuffer());
  try {
    const post = spyOn(Worker.prototype, "postMessage");
    try {
      const session = live as unknown as { session: { sharedMedia(id: string): Promise<ArrayBuffer> } };
      const [mediaId] = ids;
      const [x, y] = await Promise.all([session.session.sharedMedia(mediaId!), session.session.sharedMedia(mediaId!)]);
      expect(y).toBe(x);
      expect(post.mock.calls.filter(([message]) => (message as { type?: string }).type === "media")).toHaveLength(1);
    } finally {
      post.mockRestore();
    }
  } finally {
    live.close();
  }
});

// A slide resolves its colours against the theme of the master its layout belongs to: accent 1 is a dark
// green on the second master (`1E6B3C`) and blue on the first (`156082`).
test("a slide on the second master resolves accent 1 from its own theme, and the first master's slides keep theirs", async () => {
  const pkg = await OpcPackage.open(await Bun.file(COVERAGE).arrayBuffer());
  const deck = await indexPptx(pkg);
  expect([...deck.masters.values()].map((master) => master.theme.colors.accent1)).toEqual(["#156082", "#1E6B3C"]);
  expect(deck.theme.colors.accent1).toBe("#156082");
  const slide = await parsePptxSlide(pkg, deck, SECOND_MASTER.a1);
  const box = slide.elements.find((element) => element.kind === "shape" && element.name === "Accent box");
  if (box?.kind !== "shape") throw new Error("fixture structure changed: expected the Accent box on slide 15");
  expect(box.fill).toEqual({ kind: "solid", color: "#1E6B3C" });
  // A shape filled with accent 1 on a first-master slide still takes the blue one: the table-fills slide's header cells.
  const first = await parsePptxSlide(pkg, deck, 9);
  expect(JSON.stringify(first.elements)).toContain("#156082");
  expect(JSON.stringify(first.elements)).not.toContain("#1E6B3C");
});
