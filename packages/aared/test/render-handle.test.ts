import { describe, expect, test } from "bun:test";
import { RenderHandleBase, fitScale, safeScale, sameZoom } from "../src/render-handle.js";
import type { ZoomValue } from "../src/core/index.js";

class Handle extends RenderHandleBase<{ theme?: string }> {
  constructor(zoom: ZoomValue = "fit-width") {
    super(zoom);
  }

  setAppearance(appearance: { theme?: string }): void {
    if (this.destroyed) return;
    this.applyAppearance(appearance);
  }
}

function events(handle: Handle): Array<unknown> {
  const seen: unknown[] = [];
  handle.addEventListener("zoomchange", (event) => seen.push((event as CustomEvent).detail));
  return seen;
}

describe("render handle", () => {
  test("opens at the zoom it was given, silently, and announces each real change once", () => {
    const handle = new Handle("fit-page");
    const seen = events(handle);
    expect(handle.zoom).toBe("fit-page");
    handle.setZoom(1.5);
    handle.setZoom("fit-page");
    expect(handle.zoom).toBe("fit-page");
    expect(seen).toEqual([{ zoom: 1.5 }, { zoom: "fit-page" }]);
  });

  test("setting the zoom it already has re-applies it but announces nothing", () => {
    const handle = new Handle(1.5);
    const seen = events(handle);
    let applied = 0;
    handle.attach({ applyZoom: () => { applied++; return 1.5; }, destroy: () => {} });
    handle.setZoom(1.5);
    handle.setZoom("fit-width");
    handle.setZoom("fit-width");
    expect(applied).toBe(3);
    expect(seen).toEqual([{ zoom: "fit-width" }]);
  });

  test("scale reports what the painter actually applied", () => {
    const handle = new Handle("fit-width");
    expect(handle.scale).toBe(1);
    handle.attach({ applyZoom: () => 0.62, destroy: () => {} });
    handle.refit();
    expect(handle.scale).toBe(0.62);
  });

  test("a zoom set before the painter attaches is applied by its first paint", () => {
    const handle = new Handle();
    handle.setZoom(2);
    let applied = 0;
    handle.attach({ applyZoom: () => { applied++; return 2; }, destroy: () => {} });
    // Attaching does not apply anything itself: the painter refits once it has painted.
    expect(applied).toBe(0);
    expect(handle.zoom).toBe(2);
    handle.setZoom(3);
    expect(applied).toBe(1);
  });

  test("appearance updates are ignored until wired, then reach the painter", () => {
    const handle = new Handle();
    const received: unknown[] = [];
    handle.setAppearance({ theme: "dark" });
    handle.attach({ applyZoom: () => 1, destroy: () => {}, applyAppearance: (appearance) => void received.push(appearance) });
    handle.setAppearance({ theme: "light" });
    expect(received).toEqual([{ theme: "light" }]);
  });

  test("destroy runs the painter's teardown once, and everything after it is inert", () => {
    const handle = new Handle();
    const seen = events(handle);
    let torn = 0;
    let applied = 0;
    handle.attach({ applyZoom: () => { applied++; return 1; }, destroy: () => void torn++ });
    expect(handle.destroyed).toBe(false);
    handle.destroy();
    handle.destroy();
    expect(torn).toBe(1);
    expect(handle.destroyed).toBe(true);
    handle.setZoom(2);
    handle.refit();
    expect(applied).toBe(0);
    expect(seen).toEqual([]);
  });

  test("a painter that tears itself down marks the handle destroyed", () => {
    const handle = new Handle();
    handle.markDestroyed();
    expect(handle.destroyed).toBe(true);
  });
});

describe("zoom comparison", () => {
  test("fit modes compare by name, numbers by the scale they paint", () => {
    expect(sameZoom("fit-width", "fit-width")).toBe(true);
    expect(sameZoom("fit-width", "fit-page")).toBe(false);
    expect(sameZoom(1.5, 1.5)).toBe(true);
    expect(sameZoom(1, "fit-width")).toBe(false);
    // 0 and NaN both paint at 1.
    expect(sameZoom(0, Number.NaN)).toBe(true);
    expect(sameZoom(0, 1)).toBe(true);
  });
});

describe("fit scale", () => {
  const natural = { width: 1000, height: 500 };

  test("fit-width matches widths", () => {
    expect(fitScale("fit-width", { width: 500, height: 100 }, natural)).toBe(0.5);
  });

  test("fit-page also keeps the whole height in view", () => {
    expect(fitScale("fit-page", { width: 500, height: 100 }, natural)).toBe(0.2);
    expect(fitScale("fit-page", { width: 500, height: 1000 }, natural)).toBe(0.5);
  });

  test("fit-page falls back to fit-width when the room offers no height", () => {
    expect(fitScale("fit-page", { width: 500, height: 0 }, natural)).toBe(0.5);
  });

  test("no width to fit yet (not laid out) gives no scale", () => {
    expect(fitScale("fit-width", { width: 0, height: 100 }, natural)).toBeUndefined();
    expect(fitScale("fit-page", { width: 500, height: 100 }, { width: 0, height: 0 })).toBeUndefined();
  });
});

describe("zoom guard", () => {
  test("a usable numeric zoom is itself", () => {
    expect(safeScale(1)).toBe(1);
    expect(safeScale(0.25)).toBe(0.25);
    expect(safeScale(3)).toBe(3);
  });

  test("a zero, negative or non-finite zoom falls back to 1 instead of collapsing the content", () => {
    for (const zoom of [0, -1, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      expect(safeScale(zoom)).toBe(1);
    }
  });
});
