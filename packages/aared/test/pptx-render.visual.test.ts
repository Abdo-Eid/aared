import { expect, test } from "bun:test";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { findCoverageSlide } from "./fixture.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;

visual("visual DOM output: PPTX preserves RTL bullets and logical RTL table order", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE_PATH).arrayBuffer());
  const deck = await indexPptx(pkg);
  const slides = await Promise.all(deck.slideParts.map((_, index) => parsePptxSlide(pkg, deck, index)));
  const [{ slide: textSlide }, { slide: tableSlide }, { slide: shapesSlide }, { slide: gradientSlide }] = await Promise.all([
    findCoverageSlide(slides, "rtl-text"),
    findCoverageSlide(slides, "rtl-table"),
    findCoverageSlide(slides, "shapes"),
    findCoverageSlide(slides, "gradients"),
  ]);
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
      return new Response(`<!doctype html><style>html,body{margin:0;font-family:Manrope,ui-sans-serif,system-ui,sans-serif}#host{width:1280px}</style><main id="host"></main><script type="module">
        import { render } from "/pptx.mjs";
        window.renderedFixture = null;
        window.renderFixture = async (model) => {
          window.renderedFixture?.destroy();
          window.renderedFixture = await render({ slide: model, target: document.querySelector("#host") });
        };
        window.destroyFixture = () => window.renderedFixture?.destroy();
      </script>`, { headers: { "content-type": "text/html; charset=utf-8" } });
    },
  });
  try {
    await using view = new Bun.WebView({ width: 1280, height: 720 });
    await view.navigate(server.url.href);
    await view.evaluate(`renderFixture(${JSON.stringify(textSlide)})`);
    const textOutput = await view.evaluate(`(() => {
      const control = [...document.querySelectorAll('[data-aared-element="shape"] p')]
        .find((line) => line.textContent.includes("Next review: Thursday 12:00"));
      const bulletLine = [...document.querySelectorAll('[data-aared-element="shape"] p')]
        .find((line) => line.textContent.includes("تم تثبيت"));
      const marker = bulletLine?.querySelector('[data-aared-element="bullet-marker"]');
      const bulletText = bulletLine?.querySelector('[data-aared-element="paragraph-text"]');
      const controlMarker = control?.querySelector('[data-aared-element="bullet-marker"]');
      const controlText = control?.querySelector('[data-aared-element="paragraph-text"]');
      const rect = (element) => {
        const { left, right, top, bottom } = element.getBoundingClientRect();
        return { left, right, top, bottom };
      };
      const titleRun = document.querySelector('[data-aared-element="shape-text"] span span');
      const subtitleRun = [...document.querySelectorAll('[data-aared-element="shape-text"] span span')]
        .find((span) => span.textContent === "تم");
      const titleLine = [...document.querySelectorAll('[data-aared-element="shape"] p')]
        .find((line) => line.textContent.includes("تقرير الإنتاج الأسبوعي"));
      return {
        markerSide: marker?.dataset.aaredBulletSide,
        bulletGeometry: marker && bulletText ? { marker: rect(marker), text: rect(bulletText) } : null,
        control: control ? {
          direction: control.dataset.aaredDirection,
          align: control.dataset.aaredAlign,
          geometry: controlMarker && controlText ? { marker: rect(controlMarker), text: rect(controlText) } : null,
        } : null,
        titleFallback: document.querySelector('[data-aared-element="shape"]')?.dataset.aaredGeometryFallback,
        titleVertical: document.querySelector('[data-aared-element="shape-text"]')?.dataset.aaredVerticalAlign,
        titleStyle: titleRun ? { fontSize: getComputedStyle(titleRun).fontSize, fontFamily: getComputedStyle(titleRun).fontFamily, fontWeight: getComputedStyle(titleRun).fontWeight } : null,
        titleGeometry: titleRun ? rect(titleRun) : null,
        subtitleStyle: subtitleRun ? { fontSize: getComputedStyle(subtitleRun).fontSize, fontFamily: getComputedStyle(subtitleRun).fontFamily } : null,
        titleParagraphStyle: titleLine ? {
          fontSize: getComputedStyle(titleLine).fontSize,
          lineHeight: titleLine.style.lineHeight,
          sourceLineSpacing: titleLine.dataset.aaredLineSpacing,
        } : null,
        titleTextGeometry: titleLine?.closest('[data-aared-element="shape-text"]')
          ? rect(titleLine.closest('[data-aared-element="shape-text"]'))
          : null,
      };
    })()`) as {
      markerSide?: string;
      bulletGeometry: { marker: { left: number; right: number; top: number; bottom: number }; text: { left: number; right: number; top: number; bottom: number } } | null;
      control: {
        direction?: string;
        align?: string;
        geometry: { marker: { left: number; right: number; top: number; bottom: number }; text: { left: number; right: number; top: number; bottom: number } } | null;
      } | null;
      titleFallback?: string;
      titleVertical?: string;
      titleStyle: { fontSize: string; fontFamily: string; fontWeight: string } | null;
      titleGeometry: { left: number; right: number; top: number; bottom: number } | null;
      subtitleStyle: { fontSize: string; fontFamily: string } | null;
      titleParagraphStyle: { fontSize: string; lineHeight: string; sourceLineSpacing?: string } | null;
      titleTextGeometry: { left: number; right: number; top: number; bottom: number } | null;
    };
    expect(textOutput.markerSide).toBe("right");
    expect(textOutput.bulletGeometry).not.toBeNull();
    expect(textOutput.bulletGeometry!.marker.right - textOutput.bulletGeometry!.marker.left).toBeGreaterThan(0);
    expect(textOutput.bulletGeometry!.marker.left).toBeGreaterThanOrEqual(textOutput.bulletGeometry!.text.right);
    expect(textOutput.control?.direction).toBe("rtl");
    expect(textOutput.control?.align).toBe("left");
    expect(textOutput.control?.geometry).not.toBeNull();
    expect(textOutput.control!.geometry!.marker.right - textOutput.control!.geometry!.marker.left).toBeGreaterThan(0);
    expect(textOutput.control!.geometry!.marker.left).toBeGreaterThanOrEqual(textOutput.control!.geometry!.text.right);
    expect(textOutput.control!.geometry!.marker.left - textOutput.control!.geometry!.text.right).toBeLessThanOrEqual(20);
    expect(textOutput.titleFallback).toBe("missing");
    expect(textOutput.titleVertical).toBe("bottom");
    expect(textOutput.titleStyle).toMatchObject({ fontSize: "80px", fontWeight: "700" });
    expect(textOutput.titleStyle?.fontFamily).toContain("Times New Roman");
    expect(textOutput.subtitleStyle).toMatchObject({ fontSize: "32px" });
    expect(textOutput.subtitleStyle?.fontFamily).toContain("Arial");
    expect(textOutput.titleParagraphStyle).toEqual({ fontSize: "80px", lineHeight: "1.1", sourceLineSpacing: "0.9" });
    expect(textOutput.titleTextGeometry).not.toBeNull();
    expect(textOutput.titleGeometry!.bottom).toBeLessThanOrEqual(textOutput.titleTextGeometry!.bottom);

    await view.evaluate(`(() => new Promise((resolve) => {
      document.querySelector("#host").style.width = "640px";
      requestAnimationFrame(() => requestAnimationFrame(resolve));
    }))()`);
    const responsiveOutput = await view.evaluate(`(() => {
      const title = document.querySelector('[data-aared-element="shape-text"] span span');
      const titleShape = title?.closest('[data-aared-element="shape"]');
      const marker = document.querySelector('[data-aared-element="bullet-marker"]');
      const rect = (element) => {
        const { left, right, top, bottom } = element.getBoundingClientRect();
        return { left, right, top, bottom };
      };
      return {
        titleFontSize: title ? getComputedStyle(title).fontSize : undefined,
        markerFontSize: marker ? getComputedStyle(marker).fontSize : undefined,
        titleGeometry: title ? rect(title) : null,
        sceneTransform: document.querySelector('[data-aared-element="slide-scene"]')?.style.transform,
        titleFits: title && titleShape
          ? rect(title).left >= rect(titleShape).left && rect(title).right <= rect(titleShape).right
          : false,
      };
    })()`) as { titleFontSize?: string; markerFontSize?: string; titleGeometry: { left: number; right: number; top: number; bottom: number } | null; sceneTransform?: string; titleFits: boolean };
    expect(responsiveOutput.titleFontSize).toBe("80px");
    expect(responsiveOutput.markerFontSize).toBe("32px");
    expect(responsiveOutput.titleGeometry).not.toBeNull();
    expect(textOutput.titleGeometry).not.toBeNull();
    expect(responsiveOutput.titleGeometry!.right - responsiveOutput.titleGeometry!.left)
      .toBeCloseTo((textOutput.titleGeometry!.right - textOutput.titleGeometry!.left) / 2, 2);
    expect(responsiveOutput.sceneTransform).toBe("scale(0.5)");
    expect(responsiveOutput.titleFits).toBe(true);

    await view.evaluate(`renderFixture(${JSON.stringify({
      size: { width: 10, height: 5.625 },
      elements: [{
        kind: "shape",
        id: "directionless",
        name: "directionless",
        bounds: { x: 0, y: 0, width: 10, height: 5.625 },
        text: {
          paragraphs: [
            { runs: [{ text: "Inherited direction" }] },
            { bullet: { char: "•" }, runs: [{ text: "Inherited bullet" }] },
            { direction: "ltr", bullet: { char: "•" }, runs: [{ text: "Explicit LTR bullet" }] },
          ],
        },
      }],
    })})`);
    const directionOutput = await view.evaluate(`(() => {
      const inherited = [...document.querySelectorAll("p")].find((line) => line.textContent.includes("Inherited direction"));
      const inheritedBullet = [...document.querySelectorAll("p")].find((line) => line.textContent.includes("Inherited bullet"));
      const explicitLtr = [...document.querySelectorAll("p")].find((line) => line.textContent.includes("Explicit LTR bullet"));
      const geometry = (line) => {
        const marker = line?.querySelector('[data-aared-element="bullet-marker"]')?.getBoundingClientRect();
        const text = line?.querySelector('[data-aared-element="paragraph-text"]')?.getBoundingClientRect();
        return marker && text ? { markerLeft: marker.left, markerRight: marker.right, textLeft: text.left, textRight: text.right } : null;
      };
      return {
        inherited: inherited ? { style: inherited.style.direction, data: inherited.dataset.aaredDirection, computed: getComputedStyle(inherited).direction } : null,
        inheritedBullet: inheritedBullet ? { side: inheritedBullet.querySelector('[data-aared-element="bullet-marker"]')?.dataset.aaredBulletSide, geometry: geometry(inheritedBullet) } : null,
        explicitLtr: explicitLtr ? { side: explicitLtr.querySelector('[data-aared-element="bullet-marker"]')?.dataset.aaredBulletSide, geometry: geometry(explicitLtr) } : null,
      };
    })()`) as {
      inherited: { style: string; data?: string | undefined; computed: string } | null;
      inheritedBullet: { side?: string; geometry: { markerLeft: number; markerRight: number; textLeft: number; textRight: number } | null } | null;
      explicitLtr: { side?: string; geometry: { markerLeft: number; markerRight: number; textLeft: number; textRight: number } | null } | null;
    };
    expect(directionOutput.inherited).toEqual({ style: "", data: undefined, computed: "ltr" });
    expect(directionOutput.inheritedBullet?.side).toBeUndefined();
    expect(directionOutput.inheritedBullet?.geometry?.markerRight).toBeLessThanOrEqual(directionOutput.inheritedBullet!.geometry!.textLeft);
    expect(directionOutput.explicitLtr?.side).toBe("left");
    expect(directionOutput.explicitLtr?.geometry?.markerRight).toBeLessThanOrEqual(directionOutput.explicitLtr!.geometry!.textLeft);

    await view.evaluate(`(() => new Promise((resolve) => {
      document.querySelector("#host").style.width = "1280px";
      requestAnimationFrame(() => requestAnimationFrame(resolve));
    }))()`);
    await view.evaluate(`renderFixture(${JSON.stringify(tableSlide)})`);
    const tableOutput = await view.evaluate(`({
      direction: document.querySelector('[data-aared-element="table"]')?.dataset.aaredTableDirection,
      cells: [...document.querySelectorAll('[data-aared-element="table-cell"]')]
        .slice(0, 3).map((cell) => cell.textContent),
      cellDirections: [...document.querySelectorAll('[data-aared-element="table-cell"] p')]
        .slice(0, 3).map((line) => line.dataset.aaredDirection),
       cellAlignments: [...document.querySelectorAll('[data-aared-element="table-cell"]')]
         .slice(0, 3).map((cell) => cell.style.justifyContent),
       cellBackgrounds: [...document.querySelectorAll('[data-aared-element="table-cell"]')]
         .map((cell) => getComputedStyle(cell).backgroundColor),
       cellLefts: [...document.querySelectorAll('[data-aared-element="table-cell"]')]
         .slice(0, 3).map((cell) => cell.getBoundingClientRect().left),
       rowTracks: document.querySelector('[data-aared-element="table"]')?.style.gridTemplateRows,
    })`) as { direction?: string; cells: string[]; cellDirections: string[]; cellAlignments: string[]; cellBackgrounds: string[]; cellLefts: number[]; rowTracks?: string };
    expect(tableOutput.direction).toBe("rtl");
    // DOM order remains logical XML order. CSS direction mirrors display;
    // reordering this array would silently change what each column means.
    expect(tableOutput.cells).toEqual(["الاسم", "الدور", "الحالة"]);
    expect(tableOutput.cellDirections).toEqual(["ltr", "ltr", "ltr"]);
    expect(tableOutput.cellAlignments).toEqual(["center", "center", "center"]);
    expect(tableOutput.cellBackgrounds).toEqual([
      "rgb(21, 96, 130)", "rgb(21, 96, 130)", "rgb(21, 96, 130)",
      "rgb(204, 210, 216)", "rgb(204, 210, 216)", "rgb(204, 210, 216)",
      "rgb(231, 234, 237)", "rgb(231, 234, 237)", "rgb(231, 234, 237)",
      "rgb(204, 210, 216)", "rgb(204, 210, 216)", "rgb(204, 210, 216)",
    ]);
    expect(tableOutput.cellLefts[0]).toBeGreaterThan(tableOutput.cellLefts[1]!);
    expect(tableOutput.cellLefts[1]).toBeGreaterThan(tableOutput.cellLefts[2]!);
    const rowTracks = tableOutput.rowTracks?.split(" ").map(Number.parseFloat);
    expect(rowTracks).toEqual([25, 25, 25, 25]);

    const tableGeometry = await view.evaluate(`(() => {
      const rect = (element) => {
        const { left, top, width, height } = element.getBoundingClientRect();
        return { left, top, width, height };
      };
      const table = document.querySelector('[data-aared-element="table"]');
      const cell = document.querySelector('[data-aared-element="table-cell"]');
      const lastCell = [...document.querySelectorAll('[data-aared-element="table-cell"]')].at(-1);
      const wrapped = [...document.querySelectorAll('[data-aared-element="table-cell"] p')]
        .find((line) => line.textContent.includes("Shot 12"));
      const range = document.createRange();
      if (wrapped) range.selectNodeContents(wrapped);
      return {
        table: table ? rect(table) : null,
        cell: cell ? rect(cell) : null,
        lastCell: lastCell ? rect(lastCell) : null,
        wrappedLineCount: wrapped ? range.getClientRects().length : 0,
      };
    })()`) as {
      table: { left: number; top: number; width: number; height: number } | null;
      cell: { left: number; top: number; width: number; height: number } | null;
      lastCell: { left: number; top: number; width: number; height: number } | null;
      wrappedLineCount: number;
    };
    expect(tableGeometry.table).not.toBeNull();
    expect(tableGeometry.cell).not.toBeNull();
    expect(tableGeometry.lastCell).not.toBeNull();
    expect(Math.abs(
      (tableGeometry.lastCell!.top + tableGeometry.lastCell!.height)
      - (tableGeometry.table!.top + tableGeometry.table!.height),
    )).toBeLessThanOrEqual(2);

    await view.evaluate(`(() => new Promise((resolve) => {
      document.querySelector("#host").style.width = "640.5px";
      requestAnimationFrame(() => requestAnimationFrame(resolve));
    }))()`);
    const fractionalGeometry = await view.evaluate(`(() => {
      const rect = (element) => {
        const { left, top, width, height } = element.getBoundingClientRect();
        return { left, top, width, height };
      };
      const table = document.querySelector('[data-aared-element="table"]');
      const cell = document.querySelector('[data-aared-element="table-cell"]');
      const lastCell = [...document.querySelectorAll('[data-aared-element="table-cell"]')].at(-1);
      const wrapped = [...document.querySelectorAll('[data-aared-element="table-cell"] p')]
        .find((line) => line.textContent.includes("Shot 12"));
      const range = document.createRange();
      if (wrapped) range.selectNodeContents(wrapped);
      return {
        table: table ? rect(table) : null,
        cell: cell ? rect(cell) : null,
        lastCell: lastCell ? rect(lastCell) : null,
        wrappedLineCount: wrapped ? range.getClientRects().length : 0,
        sceneTransform: document.querySelector('[data-aared-element="slide-scene"]')?.style.transform,
      };
    })()`) as {
      table: { left: number; top: number; width: number; height: number } | null;
      cell: { left: number; top: number; width: number; height: number } | null;
      lastCell: { left: number; top: number; width: number; height: number } | null;
      wrappedLineCount: number;
      sceneTransform?: string;
    };
    const fractionalScale = 640.5 / 1280;
    for (const key of ["left", "top", "width", "height"] as const) {
      expect(fractionalGeometry.table![key]).toBeCloseTo(tableGeometry.table![key] * fractionalScale, 2);
      expect(fractionalGeometry.cell![key]).toBeCloseTo(tableGeometry.cell![key] * fractionalScale, 2);
      expect(fractionalGeometry.lastCell![key]).toBeCloseTo(tableGeometry.lastCell![key] * fractionalScale, 2);
    }
    expect(fractionalGeometry.wrappedLineCount).toBe(tableGeometry.wrappedLineCount);
    expect(Number(/scale\((.+)\)/.exec(fractionalGeometry.sceneTransform ?? "")?.[1]))
      .toBeCloseTo(fractionalScale, 5);

    await view.evaluate(`(() => new Promise((resolve) => {
      document.querySelector("#host").style.width = "1023px";
      requestAnimationFrame(() => requestAnimationFrame(resolve));
    }))()`);
    await view.evaluate(`renderFixture(${JSON.stringify(shapesSlide)})`);
    const shapeOutput = await view.evaluate(`(() => {
      const shapes = [...document.querySelectorAll('[data-aared-element="shape"]')];
      const rotated = shapes.find((shape) => shape.style.transform.includes("rotate(330deg)"));
      const titleLine = [...document.querySelectorAll('[data-aared-element="shape-text"] p')]
        .find((line) => line.textContent.includes("أشكال ورسومات"));
      const titleText = titleLine?.closest('[data-aared-element="shape-text"]');
      const titleRange = document.createRange();
      if (titleLine) titleRange.selectNodeContents(titleLine);
      const rect = (element) => {
        const { left, right, top, bottom } = element.getBoundingClientRect();
        return { left, right, top, bottom };
      };
        return {
        shapeCount: shapes.length,
        svgCount: document.querySelectorAll('[data-aared-element="shape-geometry"]').length,
        canvasCount: document.querySelectorAll("canvas").length,
        title: shapes.map((shape) => shape.querySelector('[data-aared-element="shape-text"]')?.textContent).find((text) => text?.includes("أشكال ورسومات")),
        blueOutlinedShape: shapes.some((shape) => {
          const path = shape.querySelector("path");
          return path?.getAttribute("fill") === "#73A0B4" && path.getAttribute("stroke") === "#156082";
        }),
         rotatedTransform: rotated?.style.transform,
        titleInsets: shapes.at(-1)?.querySelector('[data-aared-element="shape-text"]')?.style.paddingLeft,
        titleFontFamily: titleLine?.querySelector("span span") ? getComputedStyle(titleLine.querySelector("span span")).fontFamily : undefined,
        titleGeometry: titleLine && titleText ? { range: rect(titleRange), container: rect(titleText) } : null,
         fallback: shapes.find((shape) => shape.dataset.aaredGeometryFallback === "unsupported")?.dataset.aaredGeometryFallback,
        fallbackFill: shapes.find((shape) => shape.dataset.aaredGeometryFallback === "unsupported")?.querySelector("path")?.getAttribute("fill"),
      };
    })()`) as {
      shapeCount: number;
      svgCount: number;
      canvasCount: number;
      title?: string;
        blueOutlinedShape: boolean;
      rotatedTransform?: string;
      titleInsets?: string;
      titleFontFamily?: string;
      titleGeometry: { range: { left: number; right: number; top: number; bottom: number }; container: { left: number; right: number; top: number; bottom: number } } | null;
      fallback?: string;
      fallbackFill?: string;
    };
    expect(shapeOutput.shapeCount).toBeGreaterThanOrEqual(17);
    expect(shapeOutput.svgCount).toBe(shapeOutput.shapeCount);
    expect(shapeOutput.canvasCount).toBe(0);
    expect(shapeOutput.title).toContain("أشكال ورسومات");
    expect(shapeOutput.blueOutlinedShape).toBe(true);
    expect(shapeOutput.rotatedTransform).toContain("rotate(330deg)");
    expect(shapeOutput.titleFontFamily).toContain("Arial");
    expect(shapeOutput.titleGeometry).not.toBeNull();
    expect(shapeOutput.titleGeometry!.range.left).toBeGreaterThanOrEqual(shapeOutput.titleGeometry!.container.left);
    expect(shapeOutput.titleGeometry!.range.right).toBeLessThanOrEqual(shapeOutput.titleGeometry!.container.right);
    expect(shapeOutput.fallback).toBe("unsupported");
    expect(shapeOutput.fallbackFill).toBe("#73A0B4");

    await view.evaluate(`renderFixture(${JSON.stringify(gradientSlide)})`);
    const gradientOutput = await view.evaluate(`(() => {
      const gradient = document.querySelector('linearGradient');
      const titleLine = [...document.querySelectorAll('[data-aared-element="shape-text"] p')]
        .find((line) => line.textContent.includes("تدرجات لونية"));
      const titleText = titleLine?.closest('[data-aared-element="shape-text"]');
      const titleRange = document.createRange();
      if (titleLine) titleRange.selectNodeContents(titleLine);
      const rect = (element) => {
        const { left, right, top, bottom } = element.getBoundingClientRect();
        return { left, right, top, bottom };
      };
      const stops = gradient ? [...gradient.querySelectorAll('stop')].map((stop) => ({
        offset: stop.getAttribute('offset'),
        color: stop.getAttribute('stop-color'),
      })) : [];
      const gradientPath = gradient?.closest('svg')?.querySelector('path');
      const pathFallback = [...document.querySelectorAll('[data-aared-element="shape"]')]
        .find((shape) => shape.dataset.aaredFillFallback === 'gradient-path');
      return {
        gradientId: gradient?.id,
        gradientCoordinates: gradient ? [gradient.getAttribute('x1'), gradient.getAttribute('y1'), gradient.getAttribute('x2'), gradient.getAttribute('y2')] : null,
        gradientPathFill: gradientPath?.getAttribute('fill'),
        stops,
        pathFallback: pathFallback?.dataset.aaredFillFallback,
        pathFallbackPaint: pathFallback?.querySelector('path')?.getAttribute('fill'),
        canvasCount: document.querySelectorAll('canvas').length,
        title: [...document.querySelectorAll('[data-aared-element="shape-text"]')]
          .map((element) => element.textContent)
          .find((text) => text.includes("تدرجات لونية")),
        titleFontFamily: titleLine?.querySelector("span span") ? getComputedStyle(titleLine.querySelector("span span")).fontFamily : undefined,
        titleGeometry: titleLine && titleText ? { range: rect(titleRange), container: rect(titleText) } : null,
      };
    })()`) as {
      gradientId?: string;
      gradientCoordinates: string[] | null;
      gradientPathFill?: string;
      stops: Array<{ offset: string | null; color: string | null }>;
      pathFallback?: string;
      pathFallbackPaint?: string;
      canvasCount: number;
      title?: string;
      titleFontFamily?: string;
      titleGeometry: { range: { left: number; right: number; top: number; bottom: number }; container: { left: number; right: number; top: number; bottom: number } } | null;
    };
    expect(gradientOutput.gradientId).toMatch(/^aared-gradient-[a-z0-9]+-\d+-\d+$/);
    expect(gradientOutput.gradientPathFill).toBe(`url(#${gradientOutput.gradientId})`);
    expect(gradientOutput.gradientCoordinates).toEqual(["0", "0", "1", "1"]);
    expect(gradientOutput.stops).toEqual([
      { offset: "0%", color: "#156082" },
      { offset: "100%", color: "#E8A33D" },
    ]);
    expect(gradientOutput.pathFallback).toBe("gradient-path");
    expect(gradientOutput.pathFallbackPaint).toBe("none");
    expect(gradientOutput.canvasCount).toBe(0);
    expect(gradientOutput.title).toContain("تدرجات لونية");
    expect(gradientOutput.titleFontFamily).toContain("Arial");
    expect(gradientOutput.titleGeometry).not.toBeNull();
    expect(gradientOutput.titleGeometry!.range.left).toBeGreaterThanOrEqual(gradientOutput.titleGeometry!.container.left);
    expect(gradientOutput.titleGeometry!.range.right).toBeLessThanOrEqual(gradientOutput.titleGeometry!.container.right);

    await view.evaluate(`renderFixture(${JSON.stringify(gradientSlide)})`);
    const repeatedGradientId = await view.evaluate(`document.querySelector('linearGradient')?.id`) as string | undefined;
    expect(repeatedGradientId).toBe(gradientOutput.gradientId);
    const destroyedOutput = await view.evaluate(`(() => new Promise((resolve) => {
      const scene = document.querySelector('[data-aared-element="slide-scene"]');
      const before = scene?.style.transform;
      destroyFixture();
      document.querySelector("#host").style.width = "320px";
      requestAnimationFrame(() => requestAnimationFrame(() => resolve({ before, after: scene?.style.transform })));
    }))()`) as { before?: string; after?: string };
    expect(destroyedOutput.after).toBe(destroyedOutput.before);

  } finally {
    server.stop();
  }
});
