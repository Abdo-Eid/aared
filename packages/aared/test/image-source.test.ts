import { describe, expect, test } from "bun:test";
import { resolveImageSource } from "../src/core/image-source.js";
import type { ImageSourceOwner } from "../src/core/image-source.js";
import { parseXml } from "../src/core/xml.js";

const IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

function owner(parts: string[] = ["media/a.png", "media/a.svg"]): ImageSourceOwner {
  return {
    relationships: [
      { id: "rId1", type: IMAGE, target: "media/a.png", targetMode: "Internal" },
      { id: "rId2", type: IMAGE, target: "media/a.svg", targetMode: "Internal" },
    ],
    mediaContentType: (target) => (target.endsWith(".svg") ? "image/svg+xml" : "image/png"),
    hasPart: (target) => parts.includes(target),
  };
}

const SVG_EXT =
  `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="RID"/></a:ext></a:extLst>`;
const RASTER_EXT =
  `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId1"/></a:ext></a:extLst>`;

const blip = (attrs: string, inner = "") =>
  parseXml(
    `<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
      `xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"${attrs}>${inner}</a:blip>`,
  );

describe("resolveImageSource", () => {
  test("a plain blip resolves exactly as before, with no fallback", () => {
    expect(resolveImageSource(blip(` r:embed="rId1"`), owner())).toEqual({
      mediaId: "media/a.png",
      contentType: "image/png",
    });
  });

  test("an embedded SVG wins with the raster as its fallback", () => {
    expect(
      resolveImageSource(blip(` r:embed="rId1"`, SVG_EXT.replace("RID", "rId2")), owner()),
    ).toEqual({
      mediaId: "media/a.svg",
      contentType: "image/svg+xml",
      fallbackMediaId: "media/a.png",
      fallbackContentType: "image/png",
    });
  });

  test("a lone svgBlip (Office's picture-bullet shape) paints alone", () => {
    expect(resolveImageSource(blip(``, SVG_EXT.replace("RID", "rId2")), owner())).toEqual({
      mediaId: "media/a.svg",
      contentType: "image/svg+xml",
    });
  });

  test("a dangling SVG reference keeps the main blip", () => {
    expect(
      resolveImageSource(blip(` r:embed="rId1"`, SVG_EXT.replace("RID", "rId2")), owner(["media/a.png"])),
    ).toEqual({
      mediaId: "media/a.png",
      contentType: "image/png",
    });
  });

  test("an svgBlip naming a raster part is not treated as vector", () => {
    expect(resolveImageSource(blip(` r:embed="rId1"`, RASTER_EXT), owner())).toEqual({
      mediaId: "media/a.png",
      contentType: "image/png",
    });
  });

  test("external and mistyped references never resolve", () => {
    const external: ImageSourceOwner = {
      relationships: [{ id: "rId9", type: IMAGE, target: "https://example.com/a.svg", targetMode: "External" }],
      mediaContentType: () => "image/svg+xml",
    };
    expect(resolveImageSource(blip(` r:embed="rId9"`), external)).toBeUndefined();
    const typed: ImageSourceOwner = {
      relationships: [{ id: "rId9", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", target: "media/a.svg", targetMode: "Internal" }],
      mediaContentType: () => "image/svg+xml",
    };
    expect(resolveImageSource(blip(` r:embed="rId9"`), typed)).toBeUndefined();
    expect(resolveImageSource(undefined, owner())).toBeUndefined();
    expect(resolveImageSource(blip(``), owner())).toBeUndefined();
  });
});
