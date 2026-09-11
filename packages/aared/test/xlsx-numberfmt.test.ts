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
    expect(formatNumberValue(-0.815, String.raw`0.0%;\(0.0%\);\-`)).toBe("(81.5)%");
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

  test("converts serials deterministically in UTC", () => {
    expect(serialToDate(45000, false)).toEqual({ year: 2023, month: 3, day: 15, hours: 0, minutes: 0, seconds: 0 });
    expect(formatSerialDate(45000, "dd mmm yyyy", false)).toBe("15 Mar 2023");
    expect(formatSerialDate(45000, "mm-dd-yy", false)).toBe("03-15-23");
    expect(formatSerialDate(0.5, "h:mm", false)).toBe("12:00");
    expect(formatSerialDate(1.5, "[h]:mm", false)).toBe("36:00");
    expect(formatSerialDate(45000.75, "m/d/yy h:mm", false)).toBe("3/15/23 18:00");
  });
});
