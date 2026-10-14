import { expect } from "bun:test";
import { visual, withVisualPage } from "./harness.js";
import { serveDocx } from "./docx-harness.js";

const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

/**
 * `fit-width`/`fit-page` scale layout pages against the measured container —
 * the fix for pages overflowing (or under-filling) a host narrower or
 * shorter than the page's own authored size. Flow mode has no fixed-width
 * page to fit, so both fit modes are a no-op there; only a numeric zoom
 * pins the CSS `zoom` on either mode.
 */
async function serveZoomHarness(): ReturnType<typeof serveDocx> {
  return serveDocx(
    FIXTURE_PATH,
    `<!doctype html><style>html,body{margin:0}</style>
      <div id="narrow" style="width:420px"></div>
      <div id="wide" style="width:1400px"></div>
      <div id="short" style="width:1400px;height:300px;overflow:auto"></div>
      <div id="flow" style="width:420px"></div>
      <script type="module">
        import { render } from "/docx.mjs";
        window.measure = async (model, id, options, zoom) => {
          const target = document.getElementById(id);
          const handle = await render({ document: model, target, ...options });
          if (zoom !== undefined) handle.setZoom(zoom);
          // The fixture mixes portrait and landscape pages; fit scale is
          // computed against the widest one, matching what the renderer
          // does, so measure that specific page rather than the first.
          // Page sizes are inches; the painted box is CSS pixels (96 per inch).
          const widestInches = handle.pages.length > 0 ? Math.max(...handle.pages.map((p) => p.width)) : null;
          const pageWidth = widestInches === null ? null : widestInches * 96;
          const pageHeight = handle.pages.length > 0 ? Math.max(...handle.pages.map((p) => p.height)) * 96 : null;
          const widest = handle.pages.find((p) => p.width === widestInches);
          // Public page indexes are zero-based; the painted page's data-page
          // marker carries the 1-based page number.
          const page = widest !== undefined
            ? target.querySelector('[data-aared-element="page"][data-page="' + (widest.index + 1) + '"]')
            : target.querySelector('[data-aared-element="document"]');
          const rect = page.getBoundingClientRect();
          const result = {
            zoom: handle.zoom,
            pageWidth,
            pageHeight,
            width: Math.round(rect.width),
            height: Math.round(rect.height),
            // clientWidth/clientHeight, not getBoundingClientRect(): the
            // product's own fit measures the content box (excluding a
            // scrollbar's width), and the test must match what it measures.
            containerWidth: target.clientWidth,
            containerHeight: target.clientHeight,
          };
          handle.destroy();
          return result;
        };
      </script>`,
  );
}

visual("visual DOM output: DOCX layout pages fit-width to a narrower or wider host", async () => {
  const { server, sections } = await serveZoomHarness();
  await withVisualPage({ server, width: 1600, height: 1200 }, async (view) => {
    const model = JSON.stringify(sections);
    const options = JSON.stringify({ pagination: { mode: "layout" } });
    const narrow = await view.evaluate(`measure(${model}, "narrow", ${options})`) as {
      zoom: string; pageWidth: number; width: number; containerWidth: number;
    };
    // Default zoom is fit-width: the page's rendered box matches the host,
    // not its own authored (unscaled) size.
    expect(narrow.zoom).toBe("fit-width");
    expect(narrow.width).toBeLessThanOrEqual(narrow.containerWidth);
    expect(Math.abs(narrow.width - narrow.containerWidth)).toBeLessThanOrEqual(2);
    expect(narrow.width).toBeLessThan(narrow.pageWidth);

    const wide = await view.evaluate(`measure(${model}, "wide", ${options})`) as {
      width: number; pageWidth: number; containerWidth: number;
    };
    // A host wider than the authored page scales it up to match, not left
    // stuck at its natural size with dead space around it.
    expect(Math.abs(wide.width - wide.containerWidth)).toBeLessThanOrEqual(2);
    expect(wide.width).toBeGreaterThan(wide.pageWidth);
  });
});

visual("visual DOM output: DOCX fit-page also bounds height, and a numeric zoom still pins CSS zoom", async () => {
  const { server, sections } = await serveZoomHarness();
  await withVisualPage({ server, width: 1600, height: 1200 }, async (view) => {
    const model = JSON.stringify(sections);
    const options = JSON.stringify({ pagination: { mode: "layout" } });
    const short = await view.evaluate(`measure(${model}, "short", ${options}, "fit-page")`) as {
      zoom: string; width: number; height: number; pageWidth: number; pageHeight: number;
    };
    expect(short.zoom).toBe("fit-page");
    // Bound by the container's height (300px), not its full 1400px width.
    expect(short.height).toBeLessThanOrEqual(300);
    expect(short.width).toBeLessThan(short.pageWidth);

    // A host with no height of its own offers none to fit: fit-page falls
    // back to fit-width instead of reading back the pages' own height.
    const unbounded = await view.evaluate(`measure(${model}, "narrow", ${options}, "fit-page")`) as {
      width: number; containerWidth: number;
    };
    expect(Math.abs(unbounded.width - unbounded.containerWidth)).toBeLessThanOrEqual(2);

    const pinned = await view.evaluate(`measure(${model}, "narrow", ${options}, 1)`) as {
      zoom: number; width: number; pageWidth: number;
    };
    expect(pinned.zoom).toBe(1);
    expect(Math.abs(pinned.width - pinned.pageWidth)).toBeLessThanOrEqual(1);
  });
});

visual("visual DOM output: DOCX flow mode has no fixed page, so fit modes leave it flowing at the container width", async () => {
  const { server, sections } = await serveZoomHarness();
  await withVisualPage({ server, width: 1600, height: 1200 }, async (view) => {
    const model = JSON.stringify(sections);
    const flow = await view.evaluate(`measure(${model}, "flow", {})`) as {
      pageWidth: number | null; width: number; containerWidth: number;
    };
    expect(flow.pageWidth).toBeNull();
    expect(Math.abs(flow.width - flow.containerWidth)).toBeLessThanOrEqual(2);
  });
});
