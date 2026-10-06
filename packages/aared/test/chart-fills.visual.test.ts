import { expect, test } from "bun:test";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import type { ChartModel } from "../src/core/model.js";
import { browserModule } from "./harness.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const PRODUCTION = `${import.meta.dir}/../../../fixtures/post-production-review.pptx`;
/** A 1x1 PNG: the picture a chart fill names. */
const PNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

visual("visual DOM output: chart fills paint through the shared fill painter, and what cannot paint is marked, not replaced", async () => {
  const pkg = await OpcPackage.open(await Bun.file(PRODUCTION).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slide = await parsePptxSlide(pkg, deck, 3);
  const chart = slide.elements.find((element): element is ChartModel => element.kind === "chart");
  const spec = chart?.spec;
  if (spec === undefined) throw new Error("fixture structure changed: expected a chart on slide 4");
  const [logged, planned] = spec.groups[0]!.series;
  // The real chart, with its fills swapped for each kind the reader now keeps.
  logged!.look = { ...logged!.look, fill: { kind: "pattern", preset: "dkDnDiag", foreground: "#156082", background: "#FFFFFF" } };
  planned!.look = {
    ...planned!.look,
    fill: {
      kind: "path-gradient",
      path: "circle",
      focus: { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 },
      extent: { left: 0, top: 0, right: 0, bottom: 0 },
      stops: [{ position: 0, color: "#FFFFFF" }, { position: 1, color: "#C9372C" }],
    },
  };
  spec.look = { ...spec.look, fill: { kind: "picture", mediaId: "ppt/media/chart-fill.png", contentType: "image/png" } };
  spec.plot = { ...spec.plot, look: { fill: { kind: "unsupported", source: "group" } } };

  const module = await browserModule(`${import.meta.dir}/../src/pptx/index.ts`);
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      if (new URL(request.url).pathname === "/pptx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><style>html,body{margin:0}#host{width:1280px;height:720px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        const decode = (b64) => Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)).buffer;
        window.renderFixture = async (model, withMedia) => {
          window.rendered?.destroy();
          window.rendered = await render({
            slide: model,
            target: document.querySelector("#host"),
            ...(withMedia ? { mediaResolver: async () => decode("${PNG}") } : {}),
          });
          const deadline = Date.now() + 8000;
          while (document.querySelector("[data-aared-fill-loading]") !== null && Date.now() < deadline) {
            await new Promise((resolve) => setTimeout(resolve, 25));
          }
        };
        window.describePaints = () => {
          const svg = document.querySelector('[data-aared-element="chart-geometry"]');
          const serverOf = (path) => {
            const fill = path.getAttribute("fill") ?? "";
            return fill.startsWith("url(") ? svg.querySelector("#" + CSS.escape(fill.slice(5, -1))) : null;
          };
          const paths = [...svg.querySelectorAll(":scope > path")];
          const byTag = (tag) => paths.filter((path) => serverOf(path)?.localName === tag);
          const radial = byTag("radialGradient")[0];
          const radialServer = radial ? serverOf(radial) : null;
          const box = radial?.getBBox();
          const move = /translate\\(([-\\d.]+) ([-\\d.]+)\\)/.exec(radialServer?.getAttribute("gradientTransform") ?? "");
          const patterns = byTag("pattern");
          const patternServer = patterns.map(serverOf).find((server) => server.querySelector("image") === null && server.querySelectorAll(":scope > rect").length > 1);
          const picture = patterns.map(serverOf).find((server) => server.querySelector("image") !== null);
          const marked = paths.filter((path) => path.hasAttribute("data-aared-fill-fallback"));
          return {
            role: svg.getAttribute("role"),
            label: svg.getAttribute("aria-label"),
            desc: svg.querySelector("desc")?.textContent,
            radialCount: byTag("radialGradient").length,
            radialCenter: radialServer && box && move ? [Number(radialServer.getAttribute("cx")) + Number(move[1]), Number(radialServer.getAttribute("cy")) + Number(move[2]), box.x + box.width / 2, box.y + box.height / 2] : null,
            patternCell: patternServer ? [Number(patternServer.getAttribute("width")), Number(patternServer.getAttribute("height")), patternServer.getAttribute("patternTransform")] : null,
            patternColors: patternServer ? [...new Set([...patternServer.querySelectorAll("rect")].map((rect) => rect.getAttribute("fill")))] : null,
            pictureHref: picture?.querySelector("image")?.getAttribute("href")?.slice(0, 5) ?? null,
            pictureSize: picture ? [Number(picture.getAttribute("width")), Number(picture.getAttribute("height"))] : null,
            marked: marked.map((path) => [path.getAttribute("data-aared-fill-fallback"), path.getAttribute("fill")]),
            svgMark: svg.dataset.aaredFillFallback ?? null,
            loading: svg.hasAttribute("data-aared-fill-loading"),
            canvasCount: document.querySelectorAll("canvas").length,
          };
        };
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);
    await view.evaluate(`renderFixture(${JSON.stringify(slide)}, true)`);
    const out = await view.evaluate(`describePaints()`) as {
      role: string; label: string; desc: string; radialCount: number; radialCenter: number[] | null;
      patternCell: [number, number, string | null] | null; patternColors: string[] | null;
      pictureHref: string | null; pictureSize: number[] | null; marked: string[][]; svgMark: string | null; loading: boolean; canvasCount: number;
    };

    // The paint servers are the shape painter's own: the 8 px preset tile (6 view-box points, a point being 0.75 px),
    // anchored to the chart's origin, in the authored colors.
    expect(out.patternCell).toEqual([6, 6, null]);
    expect(out.patternColors).toEqual(expect.arrayContaining(["#FFFFFF", "#156082"]));
    // A radial gradient, centered on the shape it fills (its server is built at the box's own origin, then moved onto it).
    expect(out.radialCount).toBeGreaterThanOrEqual(3);
    const [cx, cy, boxX, boxY] = out.radialCenter!;
    expect(cx).toBeCloseTo(boxX!, 0);
    expect(cy).toBeCloseTo(boxY!, 0);
    // The chart area's picture fill loaded its media (a blob: URL) and the loading flag is gone.
    expect(out.pictureHref).toBe("blob:");
    expect(out.pictureSize![0]).toBeGreaterThan(300);
    expect(out.loading).toBe(false);
    // The fill the reader could not express (a group fill) is marked, left unfilled: never a grey stand-in.
    expect(out.marked.length).toBeGreaterThanOrEqual(1);
    for (const [reason, fill] of out.marked) {
      expect(reason).toBe("group");
      expect(fill).toBe("none");
    }
    expect(out.svgMark).toBe("group");
    // Named and described for assistive technology.
    expect(out.role).toBe("img");
    expect(out.label).toBe("Chart");
    expect(out.desc).toBe("Column chart. 2 series: Logged, Planned. 3 categories: Ep 04, Ep 05, Ep 06.");
    expect(out.canvasCount).toBe(0);

    // With no media resolver a picture fill has nothing to load: marked, not replaced.
    await view.evaluate(`renderFixture(${JSON.stringify(slide)}, false)`);
    const bare = await view.evaluate(`describePaints()`) as typeof out;
    expect(bare.pictureHref).toBeNull();
    expect(bare.marked.map(([reason]) => reason)).toEqual(expect.arrayContaining(["missing-resolver", "group"]));
  } finally {
    server.stop();
  }
});
