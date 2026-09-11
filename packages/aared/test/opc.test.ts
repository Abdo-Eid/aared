import { describe, expect, test } from "bun:test";
import { OpcPackage, resolveTarget } from "../src/core/opc.js";
import { readEntries } from "../src/core/zip.js";
import { attr, parseXml, textContent } from "../src/core/xml.js";

// OPC reader. Format-shape assertions run against the real
// Arabic gate fixture; only the pure-function XML edge cases below are
// synthetic (allowed by AGENTS.md § Testing rule).
const FIXTURE_PATH = `${import.meta.dir}/../../../fixtures/arabic-pptx-rtl-coverage.pptx`;

async function fixtureBytes(): Promise<Uint8Array> {
  return new Uint8Array(await Bun.file(FIXTURE_PATH).arrayBuffer());
}

describe("zip central directory", () => {
  test("lists package parts without inflating them", async () => {
    const names = (await OpcPackage.open(await fixtureBytes())).partNames();
    expect(names).toContain("[Content_Types].xml");
    expect(names).toContain("ppt/presentation.xml");
    expect(names).toContain("ppt/slides/slide1.xml");
    expect(names).toContain("ppt/slides/slide2.xml");
  });

  test("inflates a deflated slide on demand", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    const xml = await pkg.readText("ppt/slides/slide2.xml");
    expect(xml).toContain('<a:tblPr rtl="1"');
    expect(xml).toContain("الاسم");
  });

  test("reads [Content_Types].xml", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    expect(await pkg.readText("[Content_Types].xml")).toContain("<Override");
  });

  test("rejects non-ZIP bytes with an actionable error", () => {
    expect(() => readEntries(new Uint8Array([0x50, 0x4b, 0x03]))).toThrow(/not a ZIP/);
    expect(() => readEntries(new Uint8Array([1, 2, 3, 4]))).toThrow(/not a ZIP/);
  });
});

describe("opc package", () => {
  test("resolves content types (Override, then Default by extension)", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    expect(pkg.contentTypeOf("ppt/slides/slide1.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    expect(pkg.contentTypeOf("/ppt/slides/slide1.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    expect(pkg.contentTypeOf("ppt/theme/theme1.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.theme+xml",
    );
  });

  test("package rels resolve the office document target", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    const rels = await pkg.packageRelationships();
    const office = rels.find((r) => r.type.endsWith("/officeDocument"));
    expect(office?.target).toBe("ppt/presentation.xml");
  });

  test("part rels resolve slide targets relative to their directory", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    const rels = await pkg.relationships("ppt/presentation.xml");
    const targets = rels.map((r) => r.target);
    expect(targets).toContain("ppt/slides/slide1.xml");
    expect(targets).toContain("ppt/slides/slide2.xml");
  });

  test("parts without a rels file yield [], not an error", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    expect(await pkg.relationships("[Content_Types].xml")).toEqual([]);
  });

  test("missing parts fail loudly", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    expect(pkg.readPart("ppt/slides/slide99.xml")).rejects.toThrow(/not found/);
  });

  test("resolveTarget handles relative, parent, and absolute targets", () => {
    expect(resolveTarget("ppt/slides", "slide1.xml")).toBe("ppt/slides/slide1.xml");
    expect(resolveTarget("ppt/slides", "../theme/theme1.xml")).toBe("ppt/theme/theme1.xml");
    expect(resolveTarget("", "ppt/presentation.xml")).toBe("ppt/presentation.xml");
    expect(resolveTarget("ppt/slides", "/ppt/media/a.png")).toBe("ppt/media/a.png");
  });
});

describe("xml parser edge cases (synthetic)", () => {
  test("attributes, nesting, entities, CDATA, comments", () => {
    const root = parseXml(
      `<?xml version="1.0"?><!-- top --><a:doc r:id='x1' plain="A &amp; B &#65;"><a:child>t1<![CDATA[<raw>&]]></a:child></a:doc>`,
    );
    expect(root.local).toBe("doc");
    expect(attr(root, "r:id")).toBe("x1");
    expect(attr(root, "plain")).toBe("A & B A");
    expect(textContent(root)).toBe("t1<raw>&");
  });

  test("self-closing elements and single-quote values", () => {
    const root = parseXml(`<Types><Override PartName='/a' ContentType='t'/></Types>`);
    expect(root.children).toHaveLength(1);
    const only = root.children[0];
    if (only === undefined) throw new Error("unreachable");
    expect(attr(only, "PartName")).toBe("/a");
  });

  test("mismatched tags and trailing content throw", () => {
    expect(() => parseXml("<a><b></a></b>")).toThrow(/mismatched/);
    expect(() => parseXml("<a/>trailing")).toThrow(/trailing/);
    expect(() => parseXml("")).toThrow(/no root element/);
  });
});
