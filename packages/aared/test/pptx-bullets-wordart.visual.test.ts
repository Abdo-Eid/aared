import { expect, test } from "bun:test";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { findCoverageSlide } from "./fixture.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

visual("visual DOM output: explicit bullet paint and WordArt markers re-read from the DOM", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const [{ slide: bulletsSlide }, { slide: wordartSlide }] = await Promise.all([
    findCoverageSlide(slides, "bullets"),
    findCoverageSlide(slides, "wordart"),
  ]);
  const bulletSvg = await pkg.readPart("ppt/media/image10.svg");
  const bulletB64 = Buffer.from(bulletSvg).toString("base64");
  const build = await Bun.build({
    entrypoints: [`${import.meta.dir}/../src/pptx/index.ts`],
    target: "browser",
    format: "esm",
  });
  if (!build.success) throw new Error(build.logs.map((log) => log.message).join("\n"));
  const module = await build.outputs[0]!.text();
  const server = Bun.serve({
    port: 0,
    fetch(request) {
      const url = new URL(request.url);
      if (url.pathname === "/pptx.mjs") {
        return new Response(module, { headers: { "content-type": "text/javascript; charset=utf-8" } });
      }
      return new Response(`<!doctype html><style>html,body{margin:0}#host{width:1280px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.renderedFixture = null;
        window.renderFixture = async (model, media) => {
          window.renderedFixture?.destroy();
          const bytes = media ? Uint8Array.from(atob(media.b64), (c) => c.charCodeAt(0)).buffer : undefined;
          window.renderedFixture = await render({
            slide: model,
            target: document.querySelector("#host"),
            ...(bytes === undefined ? {} : { mediaResolver: async () => bytes }),
          });
        };
        // The same render inside a shadow root, as <aared-viewer> paints: a url(#id) reference resolves only in its own tree scope.
        window.renderShadowFixture = async (model) => {
          window.renderedFixture?.destroy();
          let host = document.querySelector("#shadow-host");
          if (host === null) {
            host = document.createElement("div");
            host.id = "shadow-host";
            host.style.width = "1280px";
            document.body.append(host);
            const target = document.createElement("div");
            target.id = "shadow-target";
            host.attachShadow({ mode: "open" }).append(target);
          }
          window.renderedFixture = await render({ slide: model, target: host.shadowRoot.querySelector("#shadow-target") });
        };
        window.destroyFixture = () => window.renderedFixture?.destroy();
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);

    // Bullets without a media resolver: explicit colors/sizes paint, the
    // picture bullet falls back to its character with a marker.
    await view.evaluate(`renderFixture(${JSON.stringify(bulletsSlide)})`);
    const bulletsOutput = await view.evaluate(`(() => {
      const line = (text) => [...document.querySelectorAll("p")].find((p) => p.textContent.includes(text));
      const markerOf = (text) => line(text)?.querySelector('[data-aared-element="bullet-marker"]');
      const runOf = (text) => {
        const spans = [...(line(text)?.querySelectorAll('[data-aared-element="paragraph-text"] span') ?? [])];
        return spans.find((s) => s.textContent.includes(text.split(" ")[0]));
      };
      const red = markerOf("نقطة حمراء كبيرة");
      const redRun = runOf("نقطة حمراء كبيرة");
      const wing = markerOf("نقطة بخط مختلف");
      const pictured = markerOf("نقطة مصورة");
      return {
        red: red ? { color: getComputedStyle(red).color, size: getComputedStyle(red).fontSize } : null,
        redRunSize: redRun ? getComputedStyle(redRun).fontSize : null,
        wing: wing ? { text: wing.textContent, family: getComputedStyle(wing).fontFamily } : null,
        pictured: pictured ? {
          text: pictured.textContent,
          picture: pictured.dataset.aaredBulletPicture,
          fallback: pictured.dataset.aaredBulletPictureFallback,
          img: pictured.querySelector("img") !== null,
        } : null,
      };
    })()`) as {
      red: { color: string; size: string } | null;
      redRunSize: string | null;
      wing: { text: string | null; family: string } | null;
      pictured: { text: string | null; picture?: string; fallback?: string; img: boolean } | null;
    };
    expect(bulletsOutput.red?.color).toBe("rgb(255, 0, 0)");
    // 150% of the run size: markers scale, they do not inherit 1:1.
    const markerPx = Number.parseFloat(bulletsOutput.red?.size ?? "0");
    const runPx = Number.parseFloat(bulletsOutput.redRunSize ?? "0");
    expect(markerPx).toBeGreaterThan(0);
    expect(runPx).toBeGreaterThan(0);
    expect(markerPx / runPx).toBeCloseTo(1.5, 2);
    // The Wingdings square ("q") paints as its Unicode shape, not in Wingdings:
    // a device without Wingdings would otherwise show the letter.
    expect(bulletsOutput.wing?.text).toBe("❑");
    expect(bulletsOutput.wing?.family).not.toContain("Wingdings");
    expect(bulletsOutput.pictured).toMatchObject({ text: "•", fallback: "missing-resolver", img: false });

    // Bullets with a media resolver: the picture bullet paints an image.
    await view.evaluate(`renderFixture(${JSON.stringify(bulletsSlide)}, { b64: "${bulletB64}" })`);
    const picturedLoaded = await view.evaluate(`(() => new Promise((resolve) => {
      const check = () => {
        const img = document.querySelector('[data-aared-element="bullet-picture"]');
        if (img?.dataset.aaredMediaLoaded === "true") resolve(true);
        else setTimeout(check, 25);
      };
      setTimeout(() => resolve(false), 3000);
      check();
    }))()`) as boolean;
    expect(picturedLoaded).toBe(true);
    const picturedMeta = await view.evaluate(`(() => {
      const marker = [...document.querySelectorAll('[data-aared-element="bullet-marker"]')]
        .find((m) => m.dataset.aaredBulletPicture !== undefined);
      return marker ? { picture: marker.dataset.aaredBulletPicture, img: marker.querySelector("img") !== null } : null;
    })()`) as { picture: string; img: boolean } | null;
    expect(picturedMeta).toMatchObject({ picture: "ppt/media/image10.svg", img: true });

    // WordArt: outline strokes, gradient marking, glow, warp, shape links.
    await view.evaluate(`renderFixture(${JSON.stringify(wordartSlide)})`);
    const wordartOutput = await view.evaluate(`(() => {
      const runWith = (text) => {
        const line = [...document.querySelectorAll("p")].find((p) => p.textContent.includes(text));
        return [...(line?.querySelectorAll('[data-aared-element="paragraph-text"] span') ?? [])]
          .find((s) => s.textContent.trim() !== "");
      };
      const graded = runWith("متدرج");
      const glowing = runWith("متوهج");
      const hollow = runWith("تعبئة");
      const warped = [...document.querySelectorAll("p")].find((p) => p.textContent.includes("مقوس"));
      const shapeLink = document.querySelector('a[data-aared-element="shape-link"]');
      return {
        graded: graded ? {
          clip: getComputedStyle(graded).backgroundClip,
          color: getComputedStyle(graded).color,
          marker: graded.dataset.aaredTextGradient,
        } : null,
        glowing: glowing ? { shadow: getComputedStyle(glowing).textShadow } : null,
        hollow: hollow ? {
          filter: getComputedStyle(hollow).filter,
          stroke: hollow.getAttribute("style") ?? "",
          outline: hollow.dataset.aaredTextOutline,
          nofill: hollow.dataset.aaredTextNofill,
          // A url(#id) filter resolves in the run's own tree scope: the filter sits inside the run and that scope finds it.
          filterInRun: hollow.querySelector("filter") !== null,
          filterInRunScope: hollow.querySelector("filter") !== null && hollow.getRootNode().getElementById(hollow.querySelector("filter").id) !== null,
        } : null,
        warp: warped?.closest("[data-aared-text-warp]")?.dataset.aaredTextWarp ?? null,
        shapeLink: shapeLink ? { href: shapeLink.getAttribute("href"), text: shapeLink.textContent?.slice(0, 12) } : null,
      };
    })()`) as {
      graded: { clip: string; color: string; marker?: string } | null;
      glowing: { shadow: string } | null;
      hollow: { filter: string; stroke: string; outline?: string; nofill?: string; filterInRun: boolean; filterInRunScope: boolean } | null;
      warp: string | null;
      shapeLink: { href: string | null; text?: string } | null;
    };
    // Centered path="circle" gradient text paints as a real CSS radial
    // gradient clipped to the glyphs, not approximated as linear.
    expect(wordartOutput.graded?.marker).toBe("radial");
    expect(wordartOutput.graded?.clip).toBe("text");
    expect(wordartOutput.graded?.color).toBe("rgba(0, 0, 0, 0)");
    expect(wordartOutput.glowing?.shadow).toContain("255, 106, 28");
    // Hollow WordArt: an SVG filter traces the run's merged glyph silhouette
    // (dilate + XOR against the original alpha) instead of stroking each
    // glyph separately, so touching/overlapping Arabic glyphs read as one
    // outline with no seam at the joins — see text-outline.ts.
    expect(wordartOutput.hollow?.filter).toMatch(/^url\("?#aared-hollow-/);
    expect(wordartOutput.hollow?.stroke).not.toContain("-webkit-text-stroke");
    expect(wordartOutput.hollow).toMatchObject({ outline: "true", nofill: "true", filterInRun: true, filterInRunScope: true });
    expect(wordartOutput.warp).toBe("textArchUp");
    expect(wordartOutput.shapeLink?.href).toBe("https://example.com/");

    // The same slide inside a shadow root (how <aared-viewer> paints): the run's filter must be found
    // from the shadow root, or the run would paint as solid text.
    await view.evaluate(`renderShadowFixture(${JSON.stringify(wordartSlide)})`);
    const shadowHollow = await view.evaluate(`(() => {
      const root = document.querySelector("#shadow-host")?.shadowRoot;
      const line = [...(root?.querySelectorAll("p") ?? [])].find((p) => p.textContent.includes("تعبئة"));
      const run = [...(line?.querySelectorAll('[data-aared-element="paragraph-text"] span') ?? [])].find((s) => s.textContent.trim() !== "");
      const id = /#([^")]+)/.exec(run?.style.filter ?? "")?.[1] ?? "";
      return {
        inShadowRoot: run?.getRootNode() === root,
        filterInRun: run?.querySelector("filter") != null,
        resolves: id !== "" && root.getElementById(id) !== null,
      };
    })()`) as { inShadowRoot: boolean; filterInRun: boolean; resolves: boolean };
    expect(shadowHollow).toEqual({ inShadowRoot: true, filterInRun: true, resolves: true });
  } finally {
    server.stop();
  }
});
