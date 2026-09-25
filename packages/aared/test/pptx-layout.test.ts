import { expect, spyOn, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxLayout, parsePptxSlide } from "../src/core/pptx.js";
import { open } from "../src/pptx/index.js";

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

// These need a fixture with real master/layout content (a master logo or
// shape, a layout or slide that hides background graphics, a picture
// background). The current fixtures have none, so they are tracked here
// instead of silently untested. See docs/ROADMAP.md "Known fidelity gaps".
const pending = (): void => {};
test.todo("slides on one layout carry the same `layout` object from SlideHandle.read()", pending);
test.todo("master shapes paint under slide content; layout showMasterSp=0 drops them", pending);
test.todo("slide showMasterSp=0 sets hideBackgroundGraphics and hides layout and master graphics", pending);
test.todo("a picture background paints and is fetched once when inherited from the master", pending);
test.todo("p:clrMapOvr on a layout or slide overrides the master colour map", pending);
