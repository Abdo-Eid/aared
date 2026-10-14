import { expect, test } from "bun:test";
import { OpcPackage } from "../src/core/opc.js";
import { indexDocx, parseDocxSections } from "../src/core/docx/document.js";
import { resolveWordRunColor } from "../src/core/theme.js";
import { parseXml } from "../src/core/xml.js";
import type { DrawingTheme } from "../src/core/theme.js";
import type { DocumentSection, FlowNode, TextRunModel } from "../src/core/model.js";
import { withPart } from "./mutate.js";

export const WORD_INK = [
  ["Theme ink base", "#156082"], ["Theme ink shade 25", "#0F4761"],
  ["Theme ink tint 40", "#45B0E1"], ["Theme ink shade 50", "#0A2F40"],
  ["Theme ink tint 80", "#C1E4F5"],
] as const;

function paragraphs(sections: DocumentSection[]): Array<{ runs: TextRunModel[] }> {
  const found: Array<{ runs: TextRunModel[] }> = [];
  const walk = (nodes: FlowNode[]): void => {
    for (const node of nodes) {
      if (node.type === "paragraph" || node.type === "heading") found.push(node);
      else if (node.type === "table") for (const row of node.rows) for (const cell of row.cells) walk(cell.nodes);
    }
  };
  for (const section of sections) walk(section.root);
  return found;
}

test("Word-authored Latin and Arabic theme ink matches native Word's PDF colours", async () => {
  const pkg = await OpcPackage.open(await Bun.file(`${import.meta.dir}/../../../fixtures/docx-coverage.docx`).arrayBuffer());
  const nodes = paragraphs(await parseDocxSections(pkg, await indexDocx(pkg)));
  for (const [label, color] of WORD_INK) {
    const node = nodes.find((p) => p.runs.map((r) => r.text).join("").startsWith(label));
    expect(node).toBeDefined();
    expect(node!.runs.length).toBeGreaterThanOrEqual(2);
    for (const run of node!.runs.filter((r) => r.text.trim() !== "")) expect(run.color).toBe(color);
  }
});

test("a changed theme recalculates shade/tint rather than trusting stale cached RGB", async () => {
  const pkg = await withPart("docx-coverage.docx", "word/theme/theme1.xml", (xml) => {
    expect(xml).toContain('<a:accent1><a:srgbClr val="156082"/></a:accent1>');
    return xml.replace('<a:accent1><a:srgbClr val="156082"/></a:accent1>', '<a:accent1><a:srgbClr val="808080"/></a:accent1>');
  });
  const nodes = paragraphs(await parseDocxSections(pkg, await indexDocx(pkg)));
  const shaded = nodes.find((p) => p.runs.map((r) => r.text).join("").startsWith("Theme ink shade 25"));
  expect(shaded?.runs[0]?.color).toBe("#5F5F5F");
});

test("Word ink keeps literal/fallback colours and handles hex-byte boundaries and theme aliases", () => {
  const theme: DrawingTheme = { colors: { accent1: "#156082", dk1: "#000000" }, fonts: {
    major: { latin: "", complexScript: "", scripts: {} }, minor: { latin: "", complexScript: "", scripts: {} },
  } };
  const read = (attributes: string, value = theme) => resolveWordRunColor(parseXml(`<color ${attributes}/>`), value);
  expect(read('val="ABCDEF"')).toBe("#ABCDEF");
  expect(read('val="0F4761" themeColor="accent1" themeShade="BF"', { ...theme, colors: {} })).toBe("#0F4761");
  expect(read('val="FFFFFF" themeColor="dark1"')).toBe("#000000");
  expect(read('val="FFFFFF" themeColor="text1"')).toBe("#000000");
  expect(read('themeColor="accent1" themeShade="00"')).toBe("#000000");
  expect(read('themeColor="accent1" themeShade="FF"')).toBe("#156082");
  expect(read('themeColor="accent1" themeTint="00"')).toBe("#FFFFFF");
  expect(read('themeColor="accent1" themeTint="FF"')).toBe("#156082");
  expect(read('themeColor="accent1" themeShade="oops"')).toBe("#156082");
  expect(read('val="auto"')).toBeUndefined();
});
