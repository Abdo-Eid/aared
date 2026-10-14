import { describe, expect, test } from "bun:test";
import { OpcPackage, guessMediaContentType, officeFormatOf, resolveTarget } from "../src/core/opc.js";
import type { OfficeFormat } from "../src/core/opc.js";
import { readEntries } from "../src/core/zip.js";
import { attr, numberAttr, parseNumber, parseXml, textContent } from "../src/core/xml.js";

// OPC reader. Format-shape assertions run against the real
// Arabic gate fixture; only the pure-function XML edge cases below are
// synthetic (allowed by AGENTS.md § Testing rule).
const FIXTURES = `${import.meta.dir}/../../../fixtures`;
const FIXTURE_PATH = `${FIXTURES}/pptx-coverage.pptx`;

async function fixtureBytes(path = FIXTURE_PATH): Promise<Uint8Array> {
  return new Uint8Array(await Bun.file(path).arrayBuffer());
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

  test("resolveTarget handles relative, parent, and absolute targets", () => {
    expect(resolveTarget("ppt/slides", "slide1.xml")).toBe("ppt/slides/slide1.xml");
    expect(resolveTarget("ppt/slides", "../theme/theme1.xml")).toBe("ppt/theme/theme1.xml");
    expect(resolveTarget("", "ppt/presentation.xml")).toBe("ppt/presentation.xml");
    expect(resolveTarget("ppt/slides", "/ppt/media/a.png")).toBe("ppt/media/a.png");
  });

  test("readXml parses a part, and readXmlIfPresent returns undefined for an absent one", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    const slide = await pkg.readXml("ppt/slides/slide2.xml");
    expect(slide.local).toBe("sld");
    expect((await pkg.readXmlIfPresent("ppt/slides/slide2.xml"))?.local).toBe("sld");
    expect(await pkg.readXmlIfPresent("ppt/slides/slide99.xml")).toBeUndefined();
    await expect(pkg.readXml("ppt/slides/slide99.xml")).rejects.toThrow(/not found/);
  });

  test("the core-properties title is the file's own, trimmed; an empty one is no title", async () => {
    const titled = await OpcPackage.open(await fixtureBytes(`${FIXTURES}/docx-coverage.docx`));
    expect(await titled.coreTitle()).toBe("Production Brief — Sawt al-Nahr / موجز الإنتاج");
    // pptx-coverage.pptx writes `<dc:title></dc:title>`.
    expect(await (await OpcPackage.open(await fixtureBytes())).coreTitle()).toBeUndefined();
  });

  test("readXml hands an already-cancelled read back as an abort, never a tree", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    const controller = new AbortController();
    controller.abort();
    await expect(pkg.readXml("ppt/slides/slide2.xml", controller.signal)).rejects.toMatchObject({ name: "AbortError" });
    await expect(pkg.readXmlIfPresent("ppt/slides/slide2.xml", controller.signal)).rejects.toMatchObject({ name: "AbortError" });
  });
});

describe("office format identity", () => {
  const cases: Array<{ format: OfficeFormat; file: string }> = [
    { format: "pptx", file: "pptx-coverage.pptx" },
    { format: "xlsx", file: "xlsx-coverage.xlsx" },
    { format: "docx", file: "docx-coverage.docx" },
  ];

  test("officeDocument returns each format's main part and refuses the other two", async () => {
    for (const { format, file } of cases) {
      const pkg = await OpcPackage.open(await fixtureBytes(`${FIXTURES}/${file}`));
      const main = await pkg.officeDocument(format);
      expect(officeFormatOf(pkg.contentTypeOf(main))).toBe(format);
      for (const other of cases.filter((candidate) => candidate.format !== format)) {
        await expect(pkg.officeDocument(other.format)).rejects.toThrow(/^aared\/core: package is not (a PPTX presentation|an XLSX workbook|a Word document)$/);
      }
    }
  });

  test("Word and Excel templates open as their documents; a PowerPoint template stays unsupported", () => {
    const office = "application/vnd.openxmlformats-officedocument";
    expect(officeFormatOf(`${office}.wordprocessingml.template.main+xml`)).toBe("docx");
    expect(officeFormatOf(`${office}.spreadsheetml.template.main+xml`)).toBe("xlsx");
    expect(officeFormatOf(`${office}.presentationml.template.main+xml`)).toBeUndefined();
    expect(officeFormatOf("application/xml")).toBeUndefined();
    expect(officeFormatOf(undefined)).toBeUndefined();
  });
});

describe("media content types", () => {
  test("a declared type wins; an undeclared one is guessed from its extension, else octet-stream", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    // Declared by the package: the fixture's slide part has an Override, so the guess never applies.
    expect(pkg.mediaContentType("ppt/slides/slide1.xml")).toBe(
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    // Undeclared here, so the extension decides (case-insensitively), and unknown ones stay opaque.
    expect(pkg.mediaContentType("ppt/media/scan.EMF")).toBe("image/x-emf");
    expect(pkg.mediaContentType("ppt/media/photo.jpeg")).toBe("image/jpeg");
    expect(pkg.mediaContentType("ppt/media/blob.xyz")).toBe("application/octet-stream");
    expect(pkg.mediaContentType("ppt/media/noextension")).toBe("application/octet-stream");
  });

  test("the headless guess matches the package rule for undeclared parts", async () => {
    const pkg = await OpcPackage.open(await fixtureBytes());
    for (const part of ["a.png", "a.svg", "a.tif", "a.wmf", "a.bin", "a"]) {
      expect(pkg.mediaContentType(`ppt/media/${part}`)).toBe(guessMediaContentType(part));
    }
  });
});

describe("numeric attributes (synthetic)", () => {
  const element = parseXml(`<a:ln w="12700" bad="x"/>`);

  test("numberAttr reads a number, or undefined when the attribute is absent", () => {
    expect(numberAttr(element, "w", "a:ln")).toBe(12700);
    expect(numberAttr(element, "missing", "a:ln")).toBeUndefined();
  });

  test("a malformed value throws naming the attribute and its element", () => {
    expect(() => numberAttr(element, "bad", "a:ln")).toThrow('aared/core: invalid bad="x" on a:ln');
    expect(() => parseNumber("Infinity", "cx", "a:ext")).toThrow('aared/core: invalid cx="Infinity" on a:ext');
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
