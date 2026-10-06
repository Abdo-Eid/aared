import { expect, test } from "bun:test";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { OpcPackage } from "../src/core/opc.js";
import { mixResolvedColorsLinear } from "../src/core/color.js";
import { findCoverageSlide } from "./fixture.js";

const visual = process.env.AARED_VISUAL === "1" ? test : test.skip;
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/pptx-coverage.pptx`;

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
        titleBoxPainted: (() => {
          // A shape with no authored geometry ("missing") is an invisible
          // container in Office: no fallback marker, no outline.
          const shape = document.querySelector('[data-aared-element="shape"]');
          const path = shape?.querySelector("path");
          return path ? { fill: path.getAttribute("fill"), stroke: path.getAttribute("stroke") } : null;
        })(),
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
      titleBoxPainted: { fill?: string | null; stroke?: string | null } | null;
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
    expect(textOutput.titleFallback).toBeUndefined();
    expect(textOutput.titleBoxPainted).toEqual({ fill: "none", stroke: null });
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
       rowHeights: [...document.querySelectorAll('[data-aared-element="table-cell"]')]
         .filter((_, index) => index % 3 === 0).map((cell) => cell.getBoundingClientRect().height),
       sceneScale: document.querySelector('[data-aared-element="slide-scene"]').getBoundingClientRect().width / 1280,
    })`) as { direction?: string; cells: string[]; cellDirections: string[]; cellAlignments: string[]; cellBackgrounds: string[]; cellLefts: number[]; rowHeights: number[]; sceneScale: number };
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
    // Each painted row is at least its authored height (a row only ever
    // grows past it, to hold its text), read back from layout.
    const table = tableSlide.elements.find((element) => element.kind === "table");
    if (table?.kind !== "table") throw new Error("fixture structure changed: expected a table");
    for (const [index, height] of tableOutput.rowHeights.entries()) {
      expect(height).toBeGreaterThanOrEqual((table.rows[index]!.height ?? 0) * 96 * tableOutput.sceneScale - 1);
    }

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
        titleFontFamily: titleLine?.querySelector("span span") ? getComputedStyle(titleLine.querySelector("span span")).fontFamily : undefined,
        titleGeometry: titleLine && titleText ? { range: rect(titleRange), container: rect(titleText) } : null,
         fallback: shapes.find((shape) => shape.dataset.aaredGeometryFallback === "unsupported")?.dataset.aaredGeometryFallback,
      };
    })()`) as {
      shapeCount: number;
      svgCount: number;
      canvasCount: number;
      title?: string;
        blueOutlinedShape: boolean;
      rotatedTransform?: string;
      titleFontFamily?: string;
      titleGeometry: { range: { left: number; right: number; top: number; bottom: number }; container: { left: number; right: number; top: number; bottom: number } } | null;
      fallback?: string;
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
    // Unsupported geometry is an empty dashed outline, never a filled
    // rectangle that reads as a wrong shape.
    const fallbackShape = await view.evaluate(`(() => {
      const shape = [...document.querySelectorAll('[data-aared-shape-id]')]
        .find((s) => s.dataset.aaredGeometryFallback === "unsupported");
      const path = shape?.querySelector("path");
      return { fill: path?.getAttribute("fill"), dashed: path?.hasAttribute("stroke-dasharray") === true };
    })()`) as { fill?: string; dashed: boolean };
    expect(fallbackShape.fill).toBe("none");
    expect(fallbackShape.dashed).toBe(true);
    // Degenerate-axis shapes paint: the ellipse resolves to non-zero area,
    // the hairline connector gets a real svg viewport (its stroke extends
    // past it via overflow), and the star renders.
    const painted = await view.evaluate(`(() => {
      const byShape = (id) => document.querySelector('[data-aared-shape-id="' + id + '"]');
      const svgOf = (id) => byShape(id)?.querySelector('svg[data-aared-element="shape-geometry"]');
      const rectOf = (id) => byShape(id)?.querySelector("path")?.getBoundingClientRect();
      const ellipse = rectOf("33");
      const lineSvg = byShape("49")?.querySelector("svg");
      const star = rectOf("51");
      return {
        ellipse: ellipse ? { w: ellipse.width, h: ellipse.height } : null,
        lineViewport: lineSvg ? { h: lineSvg.clientHeight, stroke: lineSvg.querySelector("path")?.getAttribute("stroke-width") } : null,
        star: star ? { w: star.width, h: star.height } : null,
      };
    })()`) as { ellipse: { w: number; h: number } | null; lineViewport: { h: number; stroke?: string } | null; star: { w: number; h: number } | null };
    expect(painted.ellipse?.h ?? 0).toBeGreaterThan(10);
    expect(painted.lineViewport?.h ?? 0).toBeGreaterThanOrEqual(1);
    expect(Number(painted.lineViewport?.stroke)).toBeGreaterThan(0);
    expect(painted.star?.h ?? 0).toBeGreaterThan(10);

    await view.evaluate(`renderFixture(${JSON.stringify(gradientSlide)})`);
    const gradientOutput = await view.evaluate(`(() => {
      const gradient = document.querySelector('linearGradient');
      const radial = document.querySelector('radialGradient');
      const titleLine = [...document.querySelectorAll('[data-aared-element="shape-text"] p')]
        .find((line) => line.textContent.includes("تدرجات لونية"));
      const titleText = titleLine?.closest('[data-aared-element="shape-text"]');
      const titleRange = document.createRange();
      if (titleLine) titleRange.selectNodeContents(titleLine);
      const rect = (element) => {
        const { left, right, top, bottom } = element.getBoundingClientRect();
        return { left, right, top, bottom };
      };
      const stopsOf = (el) => el ? [...el.querySelectorAll('stop')].map((stop) => ({
        offset: stop.getAttribute('offset'),
        color: stop.getAttribute('stop-color'),
      })) : [];
      const gradientPath = gradient?.closest('svg')?.querySelector('path');
      const radialPath = radial?.closest('svg')?.querySelector('path');
      return {
        gradientId: gradient?.id,
        gradientCoordinates: gradient ? [gradient.getAttribute('gradientUnits'), ...['x1', 'y1', 'x2', 'y2'].map((name) => Number(gradient.getAttribute(name)))] : null,
        gradientPathFill: gradientPath?.getAttribute('fill'),
        stops: stopsOf(gradient),
        radialId: radial?.id,
        radialPathFill: radialPath?.getAttribute('fill'),
        radialStops: stopsOf(radial),
        canvasCount: document.querySelectorAll('canvas').length,
        title: [...document.querySelectorAll('[data-aared-element="shape-text"]')]
          .map((element) => element.textContent)
          .find((text) => text.includes("تدرجات لونية")),
        titleFontFamily: titleLine?.querySelector("span span") ? getComputedStyle(titleLine.querySelector("span span")).fontFamily : undefined,
        titleGeometry: titleLine && titleText ? { range: rect(titleRange), container: rect(titleText) } : null,
      };
    })()`) as {
      gradientId?: string;
      gradientCoordinates: [string | null, number, number, number, number] | null;
      gradientPathFill?: string;
      stops: Array<{ offset: string | null; color: string | null }>;
      radialId?: string;
      radialPathFill?: string;
      radialStops: Array<{ offset: string | null; color: string | null }>;
      canvasCount: number;
      title?: string;
      titleFontFamily?: string;
      titleGeometry: { range: { left: number; right: number; top: number; bottom: number }; container: { left: number; right: number; top: number; bottom: number } } | null;
    };
    expect(gradientOutput.gradientId).toMatch(/^aared-gradient-[a-z0-9]+-\d+-\d+$/);
    expect(gradientOutput.gradientPathFill).toBe(`url(#${gradientOutput.gradientId})`);
    // Unscaled (no `scaled` attribute): a true 45° in the shape's own units,
    // running top-left to bottom-right through the box's centre.
    const [units, x1, y1, x2, y2] = gradientOutput.gradientCoordinates!;
    expect(units).toBe("userSpaceOnUse");
    expect(x2 - x1).toBeCloseTo(y2 - y1, 0);
    expect((x1 + x2) / 2).toBeCloseTo(2340000 / 2, 0);
    expect((y1 + y2) / 2).toBeCloseTo(1368000 / 2, 0);
    // Each authored segment is expanded into PowerPoint's blend: the authored
    // colours at the ends, and at the segment's middle an equal mix in linear
    // light (smoothstep(0.5) = 0.5), not the sRGB average.
    const blended = (stops: Array<{ offset: string | null; color: string | null }>, from: string, to: string) => {
      expect(stops).toHaveLength(17);
      expect(stops[0]).toEqual({ offset: "0%", color: from });
      expect(stops[8]).toEqual({ offset: "50%", color: mixResolvedColorsLinear(from, to, 0.5) });
      expect(stops[16]).toEqual({ offset: "100%", color: to });
    };
    blended(gradientOutput.stops, "#156082", "#E8A33D");
    // Centered `path="circle"` gradient (the oval) paints a real SVG
    // radialGradient instead of vanishing as an unsupported fallback.
    expect(gradientOutput.radialId).toMatch(/^aared-gradient-[a-z0-9]+-\d+-\d+$/);
    expect(gradientOutput.radialPathFill).toBe(`url(#${gradientOutput.radialId})`);
    blended(gradientOutput.radialStops, "#156082", "#73A0B4");
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
