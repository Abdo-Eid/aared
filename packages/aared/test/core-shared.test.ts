import { describe, expect, test } from "bun:test";
import {
  clampChannel,
  contrastRatio,
  linearToSrgb,
  mixColors,
  parseHexColor,
  parseHexDigits,
  relativeLuminance,
  srgbToLinear,
  toHexColor,
} from "../src/core/color.js";
import { automaticTextColor } from "../src/core/pptx.js";
import { resolveDrawingColor } from "../src/core/theme.js";
import { parseXml } from "../src/core/xml.js";
import { clipSnippet, findByTitle, foldQuery, foldText, runsText, searchKey } from "../src/core/search.js";
import { MINIMUM_LINE_HEIGHT, arabicShare, usesArabic } from "../src/core/typography.js";
import {
  EMU_PER_INCH,
  EMU_PER_PT,
  PT_PER_INCH,
  PX_PER_INCH,
  TWIPS_PER_INCH,
  TWIPS_PER_PT,
  emuToInches,
  emuToPt,
  inchesToCss,
  inchesToEmu,
  inchesToPt,
  ptToInches,
  ptToPx,
  twipsToInches,
  twipsToPx,
} from "../src/core/units.js";

// Pure-function edge cases: the shared leaf modules every parser and renderer
// imports. Their real-fixture coverage lives in the format suites that use them.

describe("units", () => {
  test("the four Office unit systems agree on one inch", () => {
    expect(PX_PER_INCH).toBe(96);
    expect(PT_PER_INCH).toBe(72);
    expect(EMU_PER_INCH).toBe(914400);
    expect(TWIPS_PER_INCH).toBe(1440);
    expect(EMU_PER_PT).toBe(12700);
    expect(TWIPS_PER_PT).toBe(20);
    expect(EMU_PER_INCH / PT_PER_INCH).toBe(EMU_PER_PT);
    expect(TWIPS_PER_INCH / PT_PER_INCH).toBe(TWIPS_PER_PT);
  });

  test("conversions land on the inch and round-trip", () => {
    expect(emuToInches(EMU_PER_INCH)).toBe(1);
    expect(twipsToInches(TWIPS_PER_INCH)).toBe(1);
    expect(ptToInches(PT_PER_INCH)).toBe(1);
    expect(emuToPt(EMU_PER_PT)).toBeCloseTo(1, 12);
    expect(inchesToPt(ptToInches(10.5))).toBeCloseTo(10.5, 12);
    expect(inchesToEmu(emuToInches(9525))).toBeCloseTo(9525, 9);
    expect(twipsToPx(TWIPS_PER_INCH)).toBe(PX_PER_INCH);
    expect(ptToPx(PT_PER_INCH)).toBe(PX_PER_INCH);
    expect(ptToPx(10.5)).toBeCloseTo(14, 12);
  });

  test("model inches format as CSS pixels", () => {
    expect(inchesToCss(0.5)).toBe("48px");
    expect(inchesToCss(0)).toBe("0px");
    expect(inchesToCss(-0.25)).toBe("-24px");
  });
});

describe("colour arithmetic", () => {
  test("hex parsing is strict: six digits, an optional # only where the caller allows it", () => {
    expect(parseHexDigits("3A4961")).toEqual([0x3a, 0x49, 0x61]);
    expect(parseHexColor("#3a4961")).toEqual([0x3a, 0x49, 0x61]);
    expect(parseHexColor("3a4961")).toEqual([0x3a, 0x49, 0x61]);
    expect(parseHexDigits("#3A4961")).toBeUndefined();
    for (const bad of ["", "auto", "FFF", "#FFF", "3A49612", "FF3A4961", "rgb(0 0 0)", "GG0000"]) {
      expect(parseHexColor(bad)).toBeUndefined();
    }
  });

  test("formatting rounds, clamps and upper-cases", () => {
    expect(toHexColor([0x3a, 0x49, 0x61])).toBe("#3A4961");
    expect(toHexColor([-5, 127.6, 300])).toBe("#0080FF");
    expect(clampChannel(-1)).toBe(0);
    expect(clampChannel(255.4)).toBe(255);
    expect(clampChannel(12.5)).toBe(13);
  });

  test("mixing is an sRGB lerp, clamped, and never turns bad input into NaN", () => {
    expect(mixColors("#000000", "#FFFFFF", 0)).toBe("#000000");
    expect(mixColors("#000000", "#FFFFFF", 1)).toBe("#FFFFFF");
    expect(mixColors("#000000", "#FFFFFF", 0.5)).toBe("#808080");
    expect(mixColors("#102030", "#FFFFFF", 5)).toBe("#FFFFFF");
    expect(mixColors("#102030", "#FFFFFF", -1)).toBe("#102030");
    expect(mixColors("rgba(1, 2, 3, 0.5)", "#FFFFFF", 0.5)).toBe("rgba(1, 2, 3, 0.5)");
    expect(mixColors("#102030", "not a colour", 0.5)).toBe("#102030");
  });

  test("linear light round-trips and luminance spans black to white", () => {
    for (const channel of [0, 1, 17, 128, 254, 255]) {
      expect(linearToSrgb(srgbToLinear(channel))).toBe(channel);
    }
    expect(linearToSrgb(-1)).toBe(0);
    expect(linearToSrgb(2)).toBe(255);
    expect(relativeLuminance([0, 0, 0])).toBe(0);
    expect(relativeLuminance([255, 255, 255])).toBeCloseTo(1, 12);
  });

  test("contrast ratio is the WCAG one, symmetric, and undefined for a colour that does not parse", () => {
    expect(contrastRatio("#000000", "#FFFFFF")).toBeCloseTo(21, 10);
    expect(contrastRatio("#FFFFFF", "#000000")).toBeCloseTo(21, 10);
    expect(contrastRatio("#777777", "#777777")).toBe(1);
    // The published ratio for #767676 on white is 4.54:1, the WCAG AA boundary grey.
    expect(contrastRatio("#767676", "#FFFFFF")).toBeCloseTo(4.54, 2);
    expect(contrastRatio("767676", "#ffffff")).toBeCloseTo(4.54, 2);
    expect(contrastRatio("transparent", "#FFFFFF")).toBeUndefined();
    expect(contrastRatio("#FFFFFF", "rgb(0, 0, 0)")).toBeUndefined();
  });

  test("lumOff is signed: a negative offset darkens, a positive one lightens, both stay inside 0..1", () => {
    const theme = { colors: {} } as unknown as Parameters<typeof resolveDrawingColor>[1];
    const fill = (modifier: string): string | undefined =>
      resolveDrawingColor(parseXml(`<a:solidFill xmlns:a="x"><a:srgbClr val="808080">${modifier}</a:srgbClr></a:solidFill>`), theme);
    const grey = (hex: string | undefined): number => parseHexColor(hex!)![0]!;
    expect(grey(fill(`<a:lumOff val="-20000"/>`))).toBeLessThan(0x80);
    expect(grey(fill(`<a:lumOff val="20000"/>`))).toBeGreaterThan(0x80);
    expect(fill(`<a:lumOff val="-500000"/>`)).toBe("#000000");
    expect(fill(`<a:lumOff val="500000"/>`)).toBe("#FFFFFF");
  });

  test("PowerPoint's automatic text colour follows the shared luminance", () => {
    // Ground truth (see automaticTextColor): coverage slide titles on #3A4961 paint white.
    expect(automaticTextColor({ kind: "solid", color: "#3A4961" })).toBe("#FFFFFF");
    expect(automaticTextColor({ kind: "solid", color: "3A4961" })).toBe("#FFFFFF");
    expect(automaticTextColor({ kind: "solid", color: "#FFFFFF" })).toBe("#000000");
    expect(automaticTextColor({ kind: "solid", color: "rgb(0 0 0 / 0.5)" })).toBe("#000000");
    expect(automaticTextColor(undefined)).toBe("#000000");
  });
});

describe("search text", () => {
  test("needle and haystack pass the same fold", () => {
    expect(foldText("Total ΑΒΓ Ünï")).toBe(foldText("total αβγ ünï"));
    expect(foldQuery("Slide")).toBe(foldText("SLIDE"));
    expect(foldQuery("")).toBeNull();
  });

  test("tashkeel and tatweel do not separate spellings", () => {
    const plain = searchKey("كتب");
    expect(searchKey("كَتَبَ")).toBe(plain); // fatha on every letter
    expect(searchKey("مُحَمَّد")).toBe(searchKey("محمد")); // damma, fatha, shadda
    expect(searchKey("كـــتب")).toBe(plain); // tatweel
    expect(searchKey("الله")).toBe(searchKey("اللّٰه")); // shadda and superscript alef
  });

  test("every alef form is one letter, and alef maqsura is ya", () => {
    const bare = searchKey("احمد");
    for (const spelling of ["أحمد", "إحمد", "آحمد", "ٱحمد"]) expect(searchKey(spelling)).toBe(bare);
    expect(searchKey("الإنتاج")).toBe(searchKey("الانتاج"));
    expect(searchKey("على")).toBe(searchKey("علي"));
  });

  test("Arabic-Indic and Persian digits are Western digits, separators included", () => {
    expect(searchKey("١٢٣")).toBe("123");
    expect(searchKey("۱۲۳")).toBe("123");
    expect(searchKey("٠٩")).toBe("09");
    expect(searchKey("١٬٢٣٤٫٥")).toBe("1,234.5");
    expect(searchKey("Q١ budget")).toBe("q1 budget");
  });

  test("presentation forms take their base letters", () => {
    // U+FEFB is the isolated lam-alef ligature a PDF paste leaves behind.
    expect(searchKey(String.fromCharCode(0xfefb))).toBe(searchKey("لا"));
    expect(searchKey(`${String.fromCharCode(0xfe8d)}${String.fromCharCode(0xfee1)}`)).toBe(searchKey("ام"));
  });

  test("merges that change a word stay apart: ta marbuta against ha, the hamza seats", () => {
    expect(searchKey("مدرسة")).not.toBe(searchKey("مدرسه"));
    expect(searchKey("مؤمن")).not.toBe(searchKey("مومن"));
    expect(searchKey("مسئول")).not.toBe(searchKey("مسيول"));
  });

  test("text with no Arabic takes the plain case fold", () => {
    expect(searchKey("Total Ünï 42")).toBe(foldText("Total Ünï 42"));
    expect(searchKey("")).toBe("");
  });

  test("a query of nothing but dropped marks matches nothing, not everything", () => {
    expect(foldQuery("َّ")).toBeNull();
    expect(foldQuery("ـ")).toBeNull();
    expect(foldQuery("الإنتاج")).toBe(searchKey("الانتاج"));
  });

  test("title lookup stays case-only: it never folds Arabic spellings", () => {
    const sheets = [{ title: "أحمد" }, { title: "Q1" }];
    expect(findByTitle(sheets, "q1", (sheet) => sheet.title)).toBe(sheets[1]);
    expect(findByTitle(sheets, "احمد", (sheet) => sheet.title)).toBeUndefined();
  });

  test("a block's text is its runs joined in order, and snippets are clipped", () => {
    expect(runsText([{ text: "ab" }, { text: "" }, { text: "cd" }])).toBe("abcd");
    expect(runsText([])).toBe("");
    expect(clipSnippet("x".repeat(500))).toHaveLength(120);
    expect(clipSnippet("short")).toBe("short");
  });
});

describe("typography", () => {
  test("every Arabic-script block picks the complex-script slot", () => {
    for (const sample of [
      "الاسم", // Arabic
      "ݐ", // Arabic Supplement
      "ࢠ", // Arabic Extended-A
      "ﭐ", // Arabic Presentation Forms-A
      "ﻻ", // Arabic Presentation Forms-B
      "Total: الاسم", // mixed
    ]) {
      expect(usesArabic(sample)).toBe(true);
    }
  });

  test("Latin, digits, Hebrew and empty text do not", () => {
    for (const sample of ["", "Total 42", "שלום", "ก"]) {
      expect(usesArabic(sample)).toBe(false);
    }
  });

  test("the test is repeatable: a shared pattern must not remember its last match", () => {
    expect([usesArabic("ع"), usesArabic("ع"), usesArabic("ع")]).toEqual([true, true, true]);
    expect([arabicShare("عع"), arabicShare("عع")]).toEqual([1, 1]);
  });

  test("the Arabic share of a string is a fraction of its characters", () => {
    expect(arabicShare("")).toBe(0);
    expect(arabicShare("abcd")).toBe(0);
    expect(arabicShare("عربي")).toBe(1);
    expect(arabicShare("abعع")).toBe(0.5);
    expect(arabicShare("ab عع")).toBe(0.4);
  });

  test("the line floor keeps Arabic ink inside its box", () => {
    expect(MINIMUM_LINE_HEIGHT).toBe(1.1);
  });
});
