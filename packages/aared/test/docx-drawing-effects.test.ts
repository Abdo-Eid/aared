import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import type { DocxImageNode, DocxInlineImage, DocxTextBoxNode, FlowNode } from "../src/core/model.js";

const FIXTURE = `${import.meta.dir}/../../../fixtures/docx-coverage.docx`;

test("Word drawings preserve the same resolved effects as PowerPoint, including all picture routes", async () => {
  const pkg = await OpcPackage.open(await Bun.file(FIXTURE).arrayBuffer());
  const sections = await parseDocxSections(pkg, await indexDocx(pkg));
  const boxes: DocxTextBoxNode[] = [], pictures: Array<DocxImageNode | DocxInlineImage> = [];
  const visit = (nodes: FlowNode[]): void => {
    for (const node of nodes) {
      if (node.type === "image") pictures.push(node);
      else if (node.type === "textBox") { boxes.push(node); visit(node.content); }
      else if (node.type === "paragraph" || node.type === "heading") {
        for (const run of node.runs) if (run.image) pictures.push(run.image);
        for (const drawing of node.floats ?? []) {
          if (drawing.type === "textBox") { boxes.push(drawing); visit(drawing.content); }
          else if (drawing.type === "image") pictures.push(drawing);
        }
      } else if (node.type === "table") for (const row of node.rows) for (const cell of row.cells) visit(cell.nodes);
    }
  };
  for (const section of sections) visit(section.root);
  const text = (box: DocxTextBoxNode) => box.content.flatMap(n=>"runs" in n?n.runs.map(r=>r.text):[]).join('');
  const named = (name: string) => boxes.find(box=>text(box).startsWith(`Word effect ${name}`))!;
  expect(named('outer shadow').effects?.outerShadow).toMatchObject({blurPt:3,color:'rgb(0 0 0 / 0.6)',rotatesWithShape:false});
  expect(named('outer shadow').effects?.outerShadow?.dxPt).toBeCloseTo(5,3);
  expect(named('outer shadow').effects?.outerShadow?.dyPt).toBeCloseTo(4,3);
  expect(named('inner shadow').effects?.innerShadow?.blurPt).toBe(3);
  expect(named('glow').effects?.glow).toMatchObject({radiusPt:6,color:'rgb(220 120 10 / 0.8)'});
  expect(named('soft edge').effects?.softEdgePt).toBe(5);
  expect(named('reflection').effects?.reflection).toMatchObject({startOpacity:.5,endPosition:.55});
  expect(named('rotated shadow').transform?.rotation).toBe(12);
  for (const label of ['anchored','block','inline']) {
    const pic = pictures.find(p=>p.alt===`Word effect ${label} picture`)!;
    expect(pic).toBeDefined();
    expect(pic.effects).toBeDefined();
    expect(pic.effectInsets?.bottom).toBeGreaterThan(0);
  }
  expect(pictures.filter(p=>p.effects)).toHaveLength(3);
});
