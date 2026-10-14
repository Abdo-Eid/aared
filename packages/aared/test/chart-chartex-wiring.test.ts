import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import type { DocxChartNode, FlowNode } from "../src/core/model.js";
import { storedZip } from "./zip.js";

// How Word finds an Office 2016 chart (waterfall, treemap, ...): the graphic
// frame sits in `mc:AlternateContent/mc:Choice`, its part is a `chartEx`
// relationship, and its data is the cached `cx:lvl` form. A minimal synthetic
// package, as the other parser edge-case tests use: no Word-authored fixture
// carries such a chart (Word's own UI writes only the classic types, and its
// chart part is covered for Excel in chart-chartex.test.ts).
// The PPTX route reads the same shape from PowerPoint's own output, slide 14
// of pptx-coverage.pptx, in pptx-parser.test.ts.

const CHARTEX_REL = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const CHART = `<?xml version="1.0" encoding="UTF-8"?><cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  `<cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">Start</cx:pt><cx:pt idx="1">Up</cx:pt><cx:pt idx="2">Down</cx:pt></cx:lvl></cx:strDim>` +
  `<cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$4</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">100</cx:pt><cx:pt idx="1">60</cx:pt><cx:pt idx="2">-40</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>` +
  `<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:dataId val="0"/><cx:dataLabels pos="outEnd"><cx:visibility value="1"/></cx:dataLabels></cx:series></cx:plotAreaRegion>` +
  `<cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>`;

const RELS_HEAD = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

describe("DOCX", () => {
  test("an Office 2016 chart in mc:AlternateContent reads as a chart node through the chartEx relationship", async () => {
    const WORD = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${REL_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
    const body = `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="cx1"><w:drawing><wp:inline><wp:extent cx="4572000" cy="2743200"/><wp:docPr id="1" name="Chart 1" descr="Waterfall"/>` +
      `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" r:id="rId9"/></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>` +
      `<mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r></w:p>`;
    const zip = storedZip({
      "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
        `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      "_rels/.rels": `${RELS_HEAD}<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
      "word/document.xml": `<w:document ${WORD}><w:body>${body}</w:body></w:document>`,
      "word/_rels/document.xml.rels": `${RELS_HEAD}<Relationship Id="rId9" Type="${CHARTEX_REL}" Target="charts/chartEx1.xml"/></Relationships>`,
      "word/charts/chartEx1.xml": CHART,
    });
    const pkg = await OpcPackage.open(zip.buffer as ArrayBuffer);
    const sections = await parseDocxSections(pkg, await indexDocx(pkg));
    const nodes: FlowNode[] = sections.flatMap((s) => [...s.root]);
    const chart = nodes.find((n): n is DocxChartNode => n.type === "chart");
    if (chart === undefined) throw new Error(`expected a chart node, got ${nodes.map((n) => n.type).join(",")}`);
    expect(chart.width).toBeCloseTo(5, 5);
    expect(chart.spec.groups[0]!.series[0]!.values).toEqual([100, 60, -40]);
  });
});
