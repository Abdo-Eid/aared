import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { firstChild, parseXml } from "../src/core/xml.js";
import type { XmlElement } from "../src/core/xml.js";
import { readChart, readText } from "../src/core/chart/read.js";
import type { Ctx } from "../src/core/chart/read.js";
import { layoutChart } from "../src/core/chart/draw.js";
import { lineHeight } from "../src/core/chart/metrics.js";
import { themeForColorMap } from "../src/core/theme.js";

const theme = themeForColorMap({
  colors: { dk1: "#000000", lt1: "#FFFFFF", accent1: "#4472C4" },
  fonts: { major: { latin: "Arial", complexScript: "", scripts: {} }, minor: { latin: "Arial", complexScript: "", scripts: {} } },
}, undefined);
const base = Object.freeze({ size: 18, bold: true, italic: false, color: "#000000", family: "Arial" });
const ctx: Ctx = { theme, style: 2, text: base, owner: undefined, budget: { points: 0, read: 0 } };

function body(fit: string, def = "", run = ""): XmlElement {
  return structuredClone(parseXml(`<c:txPr xmlns:c="c" xmlns:a="a"><a:bodyPr>${fit}</a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr ${def}/></a:pPr><a:r><a:rPr ${run}/><a:t>TXPR</a:t></a:r></a:p></c:txPr>`));
}

// Native Excel 16.0 build 20430 PDF probe: 24pt -> 22.08 PDF pt at the
// fixed print scale; fontScale=50000 -> 11.04. With sz absent, 16.56 ->
// 8.28. Non-default scale survives saveback. These are numerical mutations,
// not new format fixtures; the integration cases use the existing workbook.
describe("chart text resolves the stored Office autofit scale", () => {
  test("scales after explicit paragraph/run size or inherited default resolution, without mutating the base", () => {
    const fit = '<a:normAutofit fontScale="50000"/>';
    expect(readText(body(fit, 'sz="2400"'), base, ctx).size).toBe(12);
    expect(readText(body(fit, 'sz="2400"', 'sz="3000"'), base, ctx).size).toBe(15);
    expect(readText(body(fit), base, ctx).size).toBe(9);
    expect(base.size).toBe(18);
  });

  test("absent/default scale, noAutofit and spAutoFit preserve the resolved size", () => {
    for (const fit of ["", "<a:normAutofit/>", '<a:normAutofit fontScale="100000"/>', "<a:noAutofit/>", "<a:spAutoFit/>"]) {
      expect(readText(body(fit, 'sz="2400"'), base, ctx).size).toBe(24);
      expect(readText(body(fit), base, ctx).size).toBe(18);
    }
  });

  test("line-spacing reduction shortens the leading without shrinking glyphs; malformed scales do not poison layout", () => {
    const look = readText(body('<a:normAutofit lnSpcReduction="20000"/>', 'sz="2400"'), base, ctx);
    expect(look.size).toBe(24);
    expect(look.lineSpacingReduction).toBe(0.2);
    // Excel measured 25.39 PDF pt -> 20.30 for this title: 80% of the leading.
    expect(lineHeight(look)).toBeCloseTo(lineHeight({ ...base, size: 24 }) * 0.8, 6);
    expect(readText(body('<a:normAutofit lnSpcReduction="100000"/>'), base, ctx).lineSpacingReduction).toBeUndefined();
    for (const scale of ["", "NaN", "Infinity", "-50000", "0", "100001"]) {
      expect(readText(body(`<a:normAutofit fontScale="${scale}"/>`), base, ctx).size).toBe(18);
    }
    expect(readText(body('<a:normAutofit fontScale="50000" lnSpcReduction="20000"/>'), base, ctx).size).toBe(9);
  });

  test("the shared reader handles classic, ChartEx and rich text bodies with the same rotation and styling", () => {
    for (const name of ["c:txPr", "cx:txPr", "c:rich", "cx:rich"]) {
      const xml = body('<a:normAutofit fontScale="50000"/>', 'sz="2400" i="1"');
      xml.name = name;
      firstChild(xml, "bodyPr")!.attrs.rot = "5400000";
      expect(readText(xml, base, ctx)).toMatchObject({ size: 12, rotation: 90, italic: true, bold: true, family: "Arial" });
    }
  });
});

async function fixtureChart(): Promise<XmlElement> {
  const pkg = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/post-production-tracker.xlsx`).arrayBuffer());
  return structuredClone(await pkg.readXml("xl/charts/chart1.xml"));
}

function replace(parent: XmlElement, node: XmlElement): void {
  parent.children = [...parent.children.filter((c) => c.local !== node.local), node];
}

describe("chart autofit reaches resolved models and painted text primitives", () => {
  test("a linked title scales its authored or default size, independently of chart bounds", async () => {
    for (const explicit of [true, false]) {
      const root = await fixtureChart();
      const chart = firstChild(root, "chart")!;
      replace(chart, structuredClone(parseXml('<c:title xmlns:c="c"><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>TXPR TITLE</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>')));
      const title = firstChild(chart, "title")!;
      const props = body('<a:normAutofit fontScale="50000"/>', explicit ? 'sz="2400"' : "");
      replace(title, props);
      const scaled = readChart(root, theme)!;
      firstChild(props, "bodyPr")!.children = [];
      const normal = readChart(root, theme)!;
      expect(scaled.title!.text.size).toBe(normal.title!.text.size / 2);
      for (const frame of [{ width: 8, height: 5 }, { width: 12, height: 8 }]) {
        const painted = layoutChart(scaled, frame).primitives.find((p) => p.kind === "text" && p.text === "TXPR TITLE");
        expect(painted).toMatchObject({ kind: "text", fontSize: scaled.title!.text.size });
      }
    }
  });

  test("rich title text uses its own body scale, not the sibling txPr scale (native Excel measured)", async () => {
    const root = await fixtureChart();
    const chart = firstChild(root, "chart")!;
    const rich = body("", 'sz="2400"');
    rich.name = "c:rich";
    rich.local = "rich";
    const title = structuredClone(parseXml('<c:title xmlns:c="c"><c:tx/></c:title>'));
    firstChild(title, "tx")!.children = [rich];
    replace(title, body('<a:normAutofit fontScale="50000"/>', 'sz="2400"'));
    replace(chart, title);
    expect(readChart(root, theme)!.title!.text.size).toBe(24);
    firstChild(rich, "bodyPr")!.children = [parseXml('<a:normAutofit xmlns:a="a" fontScale="50000"/>')];
    const spec = readChart(root, theme)!;
    expect(spec.title!.text.size).toBe(12);
    expect(layoutChart(spec, { width: 8, height: 5 }).primitives.find((p) => p.kind === "text" && p.text === "TXPR"))
      .toMatchObject({ kind: "text", fontSize: 12 });
  });

  test("chart-space scaling cascades to axes; local axis and legend sizes resolve before their own scale", async () => {
    const root = await fixtureChart();
    replace(root, body('<a:normAutofit fontScale="50000"/>', 'sz="2000"'));
    const chart = firstChild(root, "chart")!;
    const plot = firstChild(chart, "plotArea")!;
    const axis = plot.children.find((c) => c.local === "catAx")!;
    replace(axis, body('<a:normAutofit fontScale="50000"/>', 'sz="2400"'));
    const legend = firstChild(chart, "legend") ?? structuredClone(parseXml('<c:legend xmlns:c="c"><c:legendPos val="r"/></c:legend>'));
    replace(chart, legend);
    replace(legend, body('<a:normAutofit fontScale="50000"/>', 'sz="1600"'));
    const spec = readChart(root, theme)!;
    expect(spec.text.size).toBe(10);
    expect(spec.axes.find((a) => a.kind === "cat")!.text.size).toBe(12);
    expect(spec.legend!.text.size).toBe(8);
  });
});
