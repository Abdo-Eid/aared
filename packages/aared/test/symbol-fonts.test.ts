import { describe, expect, test } from "bun:test";
import { symbolFontStandIn } from "../src/core/symbol-fonts.js";

describe("symbol-font stand-ins", () => {
  test("Wingdings and Symbol bullets become their Unicode shapes, plain or private-use, any case of the name", () => {
    expect(symbolFontStandIn("Wingdings", "q")).toBe("❑");
    expect(symbolFontStandIn("Wingdings", "ü")).toBe("✓");
    expect(symbolFontStandIn("wingdings", "")).toBe("▪︎");
    // Word's default first-level bullet.
    expect(symbolFontStandIn("Symbol", "")).toBe("•");
    expect(symbolFontStandIn(" Symbol ", "·")).toBe("•");
  });

  test("a shape a phone may draw as colour emoji carries the text-style selector", () => {
    expect(symbolFontStandIn("Wingdings", "þ")).toBe("☑︎");
    expect(symbolFontStandIn("Symbol", "¨")).toBe("♦︎");
  });

  test("anything else is left alone", () => {
    expect(symbolFontStandIn("Arial", "q")).toBeUndefined();
    expect(symbolFontStandIn(undefined, "")).toBeUndefined();
    expect(symbolFontStandIn("Wingdings 2", "P")).toBeUndefined();
    // One character without an entry keeps the whole text in its own font.
    expect(symbolFontStandIn("Wingdings", "qZ")).toBeUndefined();
    expect(symbolFontStandIn("Wingdings", "")).toBeUndefined();
  });
});
