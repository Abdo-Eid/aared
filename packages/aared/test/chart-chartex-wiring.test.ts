import { describe, expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexPptx, parsePptxSlide } from "../src/core/pptx.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import type { ChartModel, DocxChartNode, FlowNode } from "../src/core/model.js";
import { storedZip } from "./zip.js";

// How each format finds an Office 2016 chart (waterfall, treemap, ...): the
// graphic frame sits in `mc:AlternateContent/mc:Choice`, its part is a
// `chartEx` relationship, and its data is the cached `cx:lvl` form PowerPoint
// and Word embed. Minimal synthetic packages, as the other parser edge-case
// tests use: no Office output of this form is available to read it from
// (only Excel could be driven; its part, covered in chart-chartex.test.ts,
// has no cache).

const THEME_PATH = `${import.meta.dir}/../../../tools/chart-oracle/chartex-oracle.xlsx`;
const CHARTEX_REL = "http://schemas.microsoft.com/office/2014/relationships/chartEx";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const CHART = `<?xml version="1.0" encoding="UTF-8"?><cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  `<cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">Start</cx:pt><cx:pt idx="1">Up</cx:pt><cx:pt idx="2">Down</cx:pt></cx:lvl></cx:strDim>` +
  `<cx:numDim type="val"><cx:f>Sheet1!$B$2:$B$4</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">100</cx:pt><cx:pt idx="1">60</cx:pt><cx:pt idx="2">-40</cx:pt></cx:lvl></cx:numDim></cx:data></cx:chartData>` +
  `<cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall"><cx:dataId val="0"/><cx:dataLabels pos="outEnd"><cx:visibility value="1"/></cx:dataLabels></cx:series></cx:plotAreaRegion>` +
  `<cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis><cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea></cx:chart></cx:chartSpace>`;

async function themeXml(): Promise<string> {
  const pkg = await OpcPackage.open(await Bun.file(THEME_PATH).arrayBuffer());
  return pkg.readText("xl/theme/theme1.xml");
}

const RELS_HEAD = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;

async function pptxWith(slideTree: string): Promise<ArrayBuffer> {
  const ns = `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL_NS}" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`;
  const empty = `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>`;
  return storedZip({
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/></Types>`,
    "_rels/.rels": `${RELS_HEAD}<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
    "ppt/presentation.xml": `<p:presentation ${ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`,
    "ppt/_rels/presentation.xml.rels": `${RELS_HEAD}<Relationship Id="rId1" Type="${REL_NS}/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="${REL_NS}/slide" Target="slides/slide1.xml"/><Relationship Id="rId3" Type="${REL_NS}/theme" Target="theme/theme1.xml"/></Relationships>`,
    "ppt/slideMasters/slideMaster1.xml": `<p:sldMaster ${ns}>${empty}<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`,
    "ppt/slideMasters/_rels/slideMaster1.xml.rels": `${RELS_HEAD}<Relationship Id="rId1" Type="${REL_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="${REL_NS}/theme" Target="../theme/theme1.xml"/></Relationships>`,
    "ppt/slideLayouts/slideLayout1.xml": `<p:sldLayout ${ns}>${empty}</p:sldLayout>`,
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels": `${RELS_HEAD}<Relationship Id="rId1" Type="${REL_NS}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
    "ppt/theme/theme1.xml": await themeXml(),
    "ppt/slides/slide1.xml": `<p:sld ${ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${slideTree}</p:spTree></p:cSld></p:sld>`,
    "ppt/slides/_rels/slide1.xml.rels": `${RELS_HEAD}<Relationship Id="rId1" Type="${REL_NS}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId5" Type="${CHARTEX_REL}" Target="../charts/chartEx1.xml"/></Relationships>`,
    "ppt/charts/chartEx1.xml": CHART,
  }).buffer as ArrayBuffer;
}

const FRAME = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Waterfall 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="4572000" cy="2743200"/></p:xfrm>` +
  `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" r:id="rId5"/></a:graphicData></a:graphic></p:graphicFrame>`;

describe("PPTX", () => {
  test("an Office 2016 chart in mc:AlternateContent reads as a chart, and its fallback picture is not drawn beside it", async () => {
    const pkg = await OpcPackage.open(await pptxWith(
      `<mc:AlternateContent><mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1">${FRAME}</mc:Choice>` +
      `<mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="8" name="Fallback"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId9"/></p:blipFill><p:spPr/></p:pic></mc:Fallback></mc:AlternateContent>`,
    ));
    const slide = await parsePptxSlide(pkg, await indexPptx(pkg), 0);
    expect(slide.elements.map((e) => e.kind)).toEqual(["chart"]);
    const chart = slide.elements[0] as ChartModel;
    expect(chart).toMatchObject({ id: "7", name: "Waterfall 1", sourceChartId: "ppt/charts/chartEx1.xml", bounds: { x: 1, y: 1, width: 5, height: 3 } });
    expect(chart.spec!.groups[0]!.series[0]!.spans).toEqual([[0, 100], [100, 160], [160, 120]]);
    expect(chart.spec!.groups[0]!.series[0]!.categories).toEqual(["Start", "Up", "Down"]);
  });
});

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
