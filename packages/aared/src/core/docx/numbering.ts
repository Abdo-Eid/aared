/**
 * Word `numbering.xml` reader: `w:numId` → abstract level formats for the
 * DOCX route. Headless and pure — list grouping/ordinals live in
 * `document.ts`, marker text resolves here.
 */

import { attr, childElements, firstChild, ownString, ownTree, textContent } from "../xml.js";
import type { XmlElement } from "../xml.js";
import { throwIfAborted } from "../errors.js";
import { twipsToInches } from "../units.js";

export interface NumberingLevel {
  /** Raw `w:numFmt` value (`bullet`, `decimal`, `lowerLetter`, …). */
  numFmt: string;
  /** Raw `w:lvlText` template (`●`, `%1.`, `%2)`). */
  lvlText: string;
  start: number;
  /** Level indent in inches, from `w:lvl > w:pPr > w:ind`. */
  indentLeft?: number;
  /** Hanging indent in inches (positive magnitude), for marker alignment. */
  indentHanging?: number;
  /** The level's marker run formatting (`w:lvl > w:rPr`), resolved by the caller with theme and styles. */
  rPr?: XmlElement;
  /** What follows the marker (`w:suff`); Word's default is a tab to the text indent. */
  suffix?: "tab" | "space" | "nothing";
}

export interface NumberingData {
  abstracts: Map<string, Map<number, NumberingLevel>>;
  /** Concrete `w:numId` → abstract `w:abstractNumId`. */
  nums: Map<string, string>;
  /** `w:lvlOverride` start overrides per `numId:ilvl`. */
  startOverrides: Map<string, number>;
}

function readLevel(level: XmlElement): { ilvl: number; parsed: NumberingLevel } | undefined {
  const ilvl = attr(level, "ilvl");
  if (ilvl === undefined) return undefined;
  const index = Number(ilvl);
  if (!Number.isInteger(index) || index < 0) return undefined;
  const start = firstChild(level, "start");
  const numFmt = firstChild(level, "numFmt");
  const lvlText = firstChild(level, "lvlText");
  const pPr = firstChild(level, "pPr");
  const ind = pPr === undefined ? undefined : firstChild(pPr, "ind");
  const left = ind === undefined ? undefined : attr(ind, "left");
  const hanging = ind === undefined ? undefined : attr(ind, "hanging");
  const leftDx = left === undefined ? Number.NaN : Number(left);
  const hangingDx = hanging === undefined ? Number.NaN : Number(hanging);
  const rPr = firstChild(level, "rPr");
  const suff = firstChild(level, "suff");
  const suffValue = suff === undefined ? undefined : attr(suff, "val");
  const suffix = suffValue === "space" || suffValue === "nothing" ? suffValue : undefined;
  // The session keeps numbering for its lifetime: own every string, so it doesn't pin numbering.xml.
  return {
    ilvl: index,
    parsed: {
      numFmt: ownString(numFmt === undefined ? "bullet" : attr(numFmt, "val") ?? "bullet"),
      lvlText: ownString(lvlText === undefined ? "" : attr(lvlText, "val") ?? textContent(lvlText)),
      start: start === undefined ? 1 : Number(attr(start, "val") ?? "1"),
      ...(Number.isFinite(leftDx) ? { indentLeft: twipsToInches(leftDx) } : {}),
      ...(Number.isFinite(hangingDx) ? { indentHanging: twipsToInches(hangingDx) } : {}),
      ...(rPr === undefined ? {} : { rPr: ownTree(rPr) }),
      ...(suffix === undefined ? {} : { suffix }),
    },
  };
}

export function parseNumbering(root: XmlElement | undefined, signal?: AbortSignal): NumberingData {
  throwIfAborted(signal);
  const data: NumberingData = { abstracts: new Map(), nums: new Map(), startOverrides: new Map() };
  if (root === undefined) return data;
  for (const abstract of childElements(root, "abstractNum")) {
    throwIfAborted(signal);
    const id = attr(abstract, "abstractNumId");
    if (id === undefined) continue;
    const levels = new Map<number, NumberingLevel>();
    for (const level of childElements(abstract, "lvl")) {
      const parsed = readLevel(level);
      if (parsed !== undefined) levels.set(parsed.ilvl, parsed.parsed);
    }
    data.abstracts.set(ownString(id), levels);
  }
  for (const num of childElements(root, "num")) {
    throwIfAborted(signal);
    const numId = attr(num, "numId");
    const abstractId = firstChild(num, "abstractNumId");
    if (numId === undefined || abstractId === undefined) continue;
    const target = attr(abstractId, "val");
    if (target === undefined) continue;
    data.nums.set(ownString(numId), ownString(target));
    for (const override of childElements(num, "lvlOverride")) {
      const ilvl = attr(override, "ilvl");
      const start = firstChild(override, "startOverride");
      const value = ilvl === undefined || start === undefined ? undefined : Number(attr(start, "val") ?? "NaN");
      if (value !== undefined && Number.isInteger(value)) data.startOverrides.set(`${numId}:${ilvl}`, value);
    }
  }
  throwIfAborted(signal);
  return data;
}

/** Bullet and suppressed levels render literally; every other format counts. */
export function isOrderedFormat(numFmt: string): boolean {
  return numFmt !== "bullet" && numFmt !== "none";
}

function toLetters(ordinal: number, upper: boolean): string {
  let value = Math.max(ordinal, 1);
  let out = "";
  while (value > 0) {
    value--;
    out = String.fromCharCode((value % 26) + (upper ? 65 : 97)) + out;
    value = Math.floor(value / 26);
  }
  return out;
}

function toRoman(ordinal: number, upper: boolean): string {
  const table: Array<[number, string]> = [
    [1000, "m"], [900, "cm"], [500, "d"], [400, "cd"],
    [100, "c"], [90, "xc"], [50, "l"], [40, "xl"],
    [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"],
  ];
  let value = Math.max(ordinal, 1);
  let out = "";
  for (const [amount, numeral] of table) {
    while (value >= amount) {
      out += numeral;
      value -= amount;
    }
  }
  return upper ? out.toUpperCase() : out;
}

function counterText(numFmt: string, ordinal: number): string {
  switch (numFmt) {
    case "decimal":
      return String(Math.max(ordinal, 1));
    case "lowerLetter":
      return toLetters(ordinal, false);
    case "upperLetter":
      return toLetters(ordinal, true);
    case "lowerRoman":
      return toRoman(ordinal, false);
    case "upperRoman":
      return toRoman(ordinal, true);
    default:
      // Unrecognized counting formats degrade to decimal rather than
      // dropping the marker: the list position stays visible and honest.
      return String(Math.max(ordinal, 1));
  }
}

/**
 * Resolve one item marker. `counters` holds the current 1-based ordinal per
 * level (index = `ilvl`); `%1`–`%9` placeholders substitute across levels so
 * `%1.%2)` renders `2.b)`. Each placeholder formats with its own level's
 * `w:numFmt`, matching Word's per-level counters.
 */
export function markerFor(
  level: NumberingLevel,
  levels: Map<number, NumberingLevel>,
  ilvl: number,
  counters: number[],
): string {
  if (!isOrderedFormat(level.numFmt)) return level.lvlText;
  return level.lvlText.replace(/%([1-9])/g, (_match, digit: string) => {
    const depth = Number(digit) - 1;
    const depthLevel = depth === ilvl ? level : levels.get(depth);
    const ordinal = counters[depth] ?? 1;
    return counterText(depthLevel?.numFmt ?? level.numFmt, ordinal);
  });
}

/** Starting ordinal for a level: `w:startOverride` wins over abstract `w:start`. */
export function startFor(data: NumberingData, numId: string, ilvl: number, level: NumberingLevel | undefined): number {
  const override = data.startOverrides.get(`${numId}:${ilvl}`);
  if (override !== undefined) return override;
  return level?.start ?? 1;
}
