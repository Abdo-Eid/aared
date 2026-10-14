import { describe, expect, test } from "bun:test";
import {
  formatCodeForId,
  formatNumberValue,
  formatSerialDate,
  isDateFormatCode,
  serialToDate,
} from "../src/core/xlsx/numberfmt.js";

// Pure number-format behavior. Synthetic inputs are allowed here by
// AGENTS.md § Testing rule (a number-format string parser is the example).

describe("XLSX number formats", () => {
  test("resolves custom codes over builtins", () => {
    const custom = new Map([[164, `#,##0" EGP"`]]);
    expect(formatCodeForId(164, custom)).toBe(`#,##0" EGP"`);
    expect(formatCodeForId(9, custom)).toBe("0%");
    expect(formatCodeForId(9999, custom)).toBe("General");
  });

  test("formats integers, decimals, and grouping", () => {
    expect(formatNumberValue(12, "General")).toBe("12");
    expect(formatNumberValue(74, "0")).toBe("74");
    expect(formatNumberValue(74, "0.0")).toBe("74.0");
    expect(formatNumberValue(2925, "#,##0")).toBe("2,925");
    expect(formatNumberValue(2925.5, "#,##0.00")).toBe("2,925.50");
    expect(formatNumberValue(-5, "0")).toBe("-5");
  });

  test("formats percents and literal suffixes", () => {
    expect(formatNumberValue(0.2436, "0.0%")).toBe("24.4%");
    expect(formatNumberValue(500000, `#,##0" EGP"`)).toBe("500,000 EGP");
    expect(formatNumberValue(6.5, `0.0" h"`)).toBe("6.5 h");
  });

  test("honors positive/negative/zero sections", () => {
    expect(formatNumberValue(0.2436, "0.0%;(0.0%);-")).toBe("24.4%");
    expect(formatNumberValue(-0.3074, "#,##0;(#,##0);-")).toBe("(0)");
    // The percent sign prints where the code writes it, inside the brackets,
    // as Excel prints it (post-production-tracker.xlsx's budget sheet).
    expect(formatNumberValue(-0.815, String.raw`0.0%;\(0.0%\);\-`)).toBe("(81.5%)");
    expect(formatNumberValue(-0.3074, "0.0%;(0.0%);-")).toBe("(30.7%)");
    expect(formatNumberValue(0.0123, "0.00E+00%")).toBe("1.23E+00%");
    expect(formatNumberValue(0, String.raw`0.0%;\(0.0%\);\-`)).toBe("-");
    expect(formatNumberValue(-5, "#,##0;-#,##0")).toBe("-5");
    expect(formatNumberValue(-5, "#,##0")).toBe("-5");
  });

  test("formats scientific and fractions", () => {
    expect(formatNumberValue(0.5, "0.00E+00")).toBe("5.00E-01");
    expect(formatNumberValue(2.5, "# ?/?")).toBe("2 1/2");
    expect(formatNumberValue(2.25, "# ??/??")).toBe("2 1/4");
  });

  test("never renders a wrong number for non-finite caches", () => {
    expect(formatNumberValue(Number.NaN, "0.00")).toBe("NaN");
    expect(formatNumberValue(Number.POSITIVE_INFINITY, "#,##0")).toBe("Infinity");
  });

  test("detects date codes without mistaking literal hours", () => {
    expect(isDateFormatCode("dd mmm yyyy")).toBe(true);
    expect(isDateFormatCode("h:mm")).toBe(true);
    expect(isDateFormatCode("[h]:mm")).toBe(true);
    expect(isDateFormatCode("0%")).toBe(false);
    expect(isDateFormatCode(`0.0" h"`)).toBe(false);
    expect(isDateFormatCode("General")).toBe(false);
  });

  test("the 1904 date system counts from 1904-01-01 (serial 0), not from the 1900 epoch", () => {
    expect(serialToDate(0, true)).toMatchObject({ year: 1904, month: 1, day: 1 });
    expect(serialToDate(1, true)).toMatchObject({ year: 1904, month: 1, day: 2 });
    // The same day is 1462 serials earlier in the 1904 system.
    expect(serialToDate(45000 - 1462, true)).toEqual(serialToDate(45000, false));
    expect(formatSerialDate(24107, "yyyy-mm-dd", true)).toBe("1970-01-01");
  });

  test("day names are written out, in English, from the real weekday", () => {
    // 2023-03-15 was a Wednesday; the 1904 system names the same day.
    expect(formatSerialDate(45000, "dddd, mmmm d", false)).toBe("Wednesday, March 15");
    expect(formatSerialDate(45000, "ddd dd/mm/yyyy", false)).toBe("Wed 15/03/2023");
    expect(formatSerialDate(45000 - 1462, "dddd", true)).toBe("Wednesday");
    expect(formatSerialDate(45004, "ddd", false)).toBe("Sun");
  });

  test("converts serials deterministically in UTC", () => {
    expect(serialToDate(45000, false)).toEqual({ year: 2023, month: 3, day: 15, weekday: 3, hours: 0, minutes: 0, seconds: 0, milliseconds: 0 });
    expect(formatSerialDate(45000, "dd mmm yyyy", false)).toBe("15 Mar 2023");
    expect(formatSerialDate(45000, "mm-dd-yy", false)).toBe("03-15-23");
    expect(formatSerialDate(0.5, "h:mm", false)).toBe("12:00");
    expect(formatSerialDate(1.5, "[h]:mm", false)).toBe("36:00");
    expect(formatSerialDate(45000.75, "m/d/yy h:mm", false)).toBe("3/15/23 18:00");
  });

  // Excel 16 showed 45123.756 (18:08:38.4) under built-in id 47 as 08:38.4
  // (probe xlsx-builtin-numfmt); ECMA-376's table prints that id as "mmss.0".
  test("shows fractional seconds, and built-in id 47 is mm:ss.0", () => {
    expect(formatCodeForId(47, new Map())).toBe("mm:ss.0");
    expect(formatNumberValue(45123.756, formatCodeForId(47, new Map()))).toBe("08:38.4");
    expect(formatNumberValue(45123.756, "mm:ss.00")).toBe("08:38.40");
    expect(formatNumberValue(45123.756, "h:mm:ss")).toBe("18:08:38");
    // Rounded at the shown precision first: 59.96s carries into the next minute.
    expect(formatNumberValue(59.96 / 86_400, "mm:ss.0")).toBe("01:00.0");
  });

  // Every expectation below is what Excel 16 itself displayed for the code and
  // value (probe `xlsx-builtin-numfmt`, whitespace collapsed).
  test("optional and padded digits: # drops a trailing zero, ? blanks it, only 0 prints an integer zero", () => {
    expect(formatNumberValue(0.5, "0.0#")).toBe("0.5");
    expect(formatNumberValue(0.004, "0.0#")).toBe("0.0");
    expect(formatNumberValue(0.5, "#.00")).toBe(".50");
    expect(formatNumberValue(0.5, "??.??").replace(/\s+/g, " ").trim()).toBe(".5");
    expect(formatNumberValue(0, "??.??").replace(/\s+/g, " ").trim()).toBe(".");
    expect(formatNumberValue(0, "#,###")).toBe("");
    expect(formatNumberValue(1234.5678, "#,###")).toBe("1,235");
    expect(formatNumberValue(0, '0.00;-0.00;"zero"')).toBe("zero");
  });

  test("ids a file names without a code resolve to the ECMA-376 currency, accounting and date codes", () => {
    const code = (id: number): string => formatCodeForId(id, new Map());
    // Excel reads these from the reader's regional settings; aared shows the spec's en-US form.
    expect(formatNumberValue(-1234.5678, code(5))).toBe("($1,235)");
    expect(formatNumberValue(1234.5678, code(7))).toBe("$1,234.57");
    expect(formatNumberValue(-1234.5678, code(41))).toBe("(1,235)");
    expect(formatNumberValue(0, code(43)).replace(/\s+/g, " ").trim()).toBe("-");
    expect(formatNumberValue(1234.5678, code(44))).toBe("$1,234.57");
    expect(formatNumberValue(45000.75, code(32))).toBe("18:00:00");
    for (const id of [27, 30, 36, 50, 58]) expect(formatNumberValue(45000, code(id))).toBe("03-15-23");
    expect(code(23)).toBe("General");
  });
});
