/**
 * Word `styles.xml` reader: paragraph/character style inheritance and table
 * styles for the DOCX route. Headless, plain data — mirrors how `core/pptx.ts`
 * resolves DrawingML placeholder inheritance, but over `w:basedOn` chains.
 *
 * Precedence (Word): direct formatting > character style (+ its bases) >
 * paragraph style (+ its bases) > `w:docDefaults`. Duplicate `w:styleId`
 * definitions resolve last-wins; reference `.docx` files in the wild repeat
 * `Heading1`/`Heading2` blocks.
 *
 * A table style layers *into* that order rather than replacing it: its regions'
 * `w:rPr` sits below the paragraph style and its `w:tcPr`/`w:tblPr` below direct
 * cell formatting (probe `docx-table-style`).
 */

import { attr, childElements, firstChild, ownString, ownTree } from "../xml.js";
import type { XmlElement } from "../xml.js";
import { throwIfAborted } from "../errors.js";
import { TABLE_STYLE_REGIONS } from "../table-style.js";
import type { TableStyleRegionName } from "../table-style.js";

export interface DocxStyle {
  type: "paragraph" | "character";
  basedOn?: string;
  name?: string;
  pPr?: XmlElement;
  rPr?: XmlElement;
}

/**
 * One `w:tblStylePr` region of a table style, or the style's own whole-table
 * properties. Word authors these as separate `w:tblPr`/`w:tcPr`/`w:pPr`/`w:rPr`
 * trees per region rather than a merged one.
 */
export interface DocxTableStyleRegion {
  /** `w:tblPr`: the region's borders and cell margins. */
  tblPr?: XmlElement;
  /** `w:tcPr`: the region's cell shading and borders. */
  tcPr?: XmlElement;
  pPr?: XmlElement;
  rPr?: XmlElement;
}

/**
 * A table style: its whole-table properties plus its conditional regions.
 *
 * Word authors **no** per-cell shading for a style-driven table — the entire
 * banded look lives here (probe `docx-table-style`: four banded tables wrote
 * zero `w:shd` elements), which is why a parser that reads only direct
 * formatting paints such a table plain.
 */
export interface DocxTableStyle {
  basedOn?: string;
  name?: string;
  /** The style's own `w:tblPr`/`w:tcPr`/`w:pPr`/`w:rPr`, the `wholeTbl` region. */
  whole: DocxTableStyleRegion;
  regions: Partial<Record<TableStyleRegionName, DocxTableStyleRegion>>;
}

/** `w:tblStylePr/@w:type` spelled as `a:tblStyle` spells it, so one region resolver serves both. */
const TABLE_STYLE_REGION_TYPES: Record<string, TableStyleRegionName> = {
  wholeTable: "wholeTbl",
  band1Vert: "band1V",
  band2Vert: "band2V",
  band1Horz: "band1H",
  band2Horz: "band2H",
  firstCol: "firstCol",
  lastCol: "lastCol",
  firstRow: "firstRow",
  lastRow: "lastRow",
  nwCell: "nwCell",
  neCell: "neCell",
  swCell: "swCell",
  seCell: "seCell",
};

function readTableStyle(style: XmlElement): { id: string; parsed: DocxTableStyle } | undefined {
  const id = attr(style, "styleId");
  if (id === undefined) return undefined;
  const basedOn = firstChild(style, "basedOn");
  const name = firstChild(style, "name");
  const whole: DocxTableStyleRegion = {};
  // A style may carry `w:wholeTable` as a region *and* top-level properties;
  // Word writes the latter for `Table Grid`'s borders, so read both.
  for (const element of childElements(style, "tblPr")) whole.tblPr = ownTree(element);
  for (const element of childElements(style, "tcPr")) whole.tcPr = ownTree(element);
  for (const element of childElements(style, "pPr")) whole.pPr = ownTree(element);
  for (const element of childElements(style, "rPr")) whole.rPr = ownTree(element);
  const regions: Partial<Record<TableStyleRegionName, DocxTableStyleRegion>> = {};
  for (const region of childElements(style, "tblStylePr")) {
    const name = TABLE_STYLE_REGION_TYPES[attr(region, "type") ?? ""];
    if (name === undefined) continue;
    const parsed: DocxTableStyleRegion = {};
    for (const element of childElements(region, "tblPr")) parsed.tblPr = ownTree(element);
    for (const element of childElements(region, "tcPr")) parsed.tcPr = ownTree(element);
    for (const element of childElements(region, "pPr")) parsed.pPr = ownTree(element);
    for (const element of childElements(region, "rPr")) parsed.rPr = ownTree(element);
    regions[name] = parsed;
  }
  // The session keeps styles for its lifetime: own every string, so they don't pin styles.xml.
  return {
    id: ownString(id),
    parsed: {
      ...(basedOn === undefined ? {} : { basedOn: ownString(attr(basedOn, "val") ?? "") }),
      ...(name === undefined ? {} : { name: ownString(attr(name, "val") ?? "") }),
      whole,
      regions,
    },
  };
}

export interface StyleData {
  defaultPPr?: XmlElement;
  defaultRPr?: XmlElement;
  /**
   * The paragraph style Word marks `w:default="1"` (usually `Normal`): what a
   * paragraph with no `w:pStyle` uses, so its font, size and spacing are this
   * style's, not the bare document defaults'.
   */
  defaultParagraphStyleId?: string;
  paragraph: Map<string, DocxStyle>;
  character: Map<string, DocxStyle>;
  table: Map<string, DocxTableStyle>;
}

/** The spellings of an on/off attribute that mean on (`w:default="1"`, `"true"`, `"on"`). */
const ON_VALUES = new Set(["1", "true", "on"]);

function readStyle(style: XmlElement): { id: string; parsed: DocxStyle } | undefined {
  const id = attr(style, "styleId");
  const type = attr(style, "type");
  if (id === undefined) return undefined;
  if (type !== "paragraph" && type !== "character") return undefined;
  const basedOn = firstChild(style, "basedOn");
  const name = firstChild(style, "name");
  const pPr = firstChild(style, "pPr");
  const rPr = firstChild(style, "rPr");
  const basedOnId = basedOn === undefined ? undefined : attr(basedOn, "val");
  const displayName = name === undefined ? undefined : attr(name, "val");
  // The session keeps styles for its lifetime: own every string, so they don't pin styles.xml.
  return {
    id: ownString(id),
    parsed: {
      type,
      ...(basedOnId === undefined ? {} : { basedOn: ownString(basedOnId) }),
      ...(displayName === undefined ? {} : { name: ownString(displayName) }),
      ...(pPr === undefined ? {} : { pPr: ownTree(pPr) }),
      ...(rPr === undefined ? {} : { rPr: ownTree(rPr) }),
    },
  };
}

export function parseStyles(root: XmlElement | undefined, signal?: AbortSignal): StyleData {
  throwIfAborted(signal);
  const data: StyleData = { paragraph: new Map(), character: new Map(), table: new Map() };
  if (root === undefined) return data;
  const defaults = firstChild(root, "docDefaults");
  const rPrDefault = defaults === undefined ? undefined : firstChild(defaults, "rPrDefault");
  const pPrDefault = defaults === undefined ? undefined : firstChild(defaults, "pPrDefault");
  const defaultRPr = rPrDefault === undefined ? undefined : firstChild(rPrDefault, "rPr");
  const defaultPPr = pPrDefault === undefined ? undefined : firstChild(pPrDefault, "pPr");
  if (defaultRPr !== undefined) data.defaultRPr = ownTree(defaultRPr);
  if (defaultPPr !== undefined) data.defaultPPr = ownTree(defaultPPr);
  for (const style of childElements(root, "style")) {
    throwIfAborted(signal);
    const id = attr(style, "styleId");
    if (id === undefined) continue;
    const type = attr(style, "type");
    // Last definition wins: real files repeat Heading1/Heading2 blocks.
    if (type === "table") {
      const parsed = readTableStyle(style);
      if (parsed !== undefined) data.table.set(parsed.id, parsed.parsed);
      continue;
    }
    const parsed = readStyle(style);
    if (parsed === undefined) continue;
    if (parsed.parsed.type === "paragraph") {
      data.paragraph.set(parsed.id, parsed.parsed);
      if (ON_VALUES.has(attr(style, "default") ?? "")) data.defaultParagraphStyleId = parsed.id;
    } else {
      data.character.set(parsed.id, parsed.parsed);
    }
  }
  throwIfAborted(signal);
  return data;
}

/**
 * Style ancestry base-first (so callers apply direct formatting last).
 * Cycle-safe: a repeated `basedOn` target ends the chain instead of looping.
 */
function ancestry(styles: ReadonlyMap<string, DocxStyle>, styleId: string | undefined): DocxStyle[] {
  const chain: DocxStyle[] = [];
  const seen = new Set<string>();
  let current = styleId;
  while (current !== undefined && !seen.has(current)) {
    seen.add(current);
    const style = styles.get(current);
    if (style === undefined) break;
    chain.unshift(style);
    current = style.basedOn;
  }
  return chain;
}

export function paragraphAncestry(styles: StyleData, styleId: string | undefined): DocxStyle[] {
  return ancestry(styles.paragraph, styleId);
}

export function characterAncestry(styles: StyleData, styleId: string | undefined): DocxStyle[] {
  return ancestry(styles.character, styleId);
}

/**
 * Table style ancestry base-first, so a derived style's regions layer over its
 * base's. A `w:basedOn` cycle ends the chain rather than looping, like
 * {@link ancestry}.
 */
export function tableStyleAncestry(styles: StyleData, styleId: string | undefined): DocxTableStyle[] {
  const chain: DocxTableStyle[] = [];
  const seen = new Set<string>();
  let current = styleId;
  while (current !== undefined && !seen.has(current)) {
    seen.add(current);
    const style = styles.table.get(current);
    if (style === undefined) break;
    chain.unshift(style);
    current = style.basedOn;
  }
  return chain;
}

/**
 * The region properties in effect for one cell, in application order: the
 * `use` list weakest first, and within each region its `w:basedOn` bases first,
 * so a derived style wins over its base and a specific region over the whole
 * table.
 */
export function tableStyleRegionsFor(
  styles: StyleData,
  styleId: string | undefined,
  use: ReadonlyArray<{ name: TableStyleRegionName }>,
): DocxTableStyleRegion[] {
  const chain = tableStyleAncestry(styles, styleId);
  const resolved: DocxTableStyleRegion[] = [];
  for (const { name } of use) {
    for (const style of chain) {
      if (name === "wholeTbl") resolved.push(style.whole);
      else {
        const region = style.regions[name];
        if (region !== undefined) resolved.push(region);
      }
    }
  }
  return resolved;
}

/**
 * The `w:rPr` and `w:pPr` trees a cell's text takes from its table style, in
 * application order, so a caller can read them into its own resolvers below the
 * paragraph and character styles (probe `docx-table-style`: a `Heading 2`
 * paragraph in a header row kept its own colour over the region's, while the
 * region's bold still applied).
 */
export function tableStyleTextPropsFor(
  styles: StyleData,
  styleId: string | undefined,
  use: ReadonlyArray<{ name: TableStyleRegionName }>,
): { rPr: XmlElement[]; pPr: XmlElement[] } {
  const rPr: XmlElement[] = [];
  const pPr: XmlElement[] = [];
  for (const region of tableStyleRegionsFor(styles, styleId, use)) {
    if (region.rPr !== undefined) rPr.push(region.rPr);
    if (region.pPr !== undefined) pPr.push(region.pPr);
  }
  return { rPr, pPr };
}

const HEADING_STYLE_ID = /^Heading([1-9])$/;
const HEADING_STYLE_NAME = /^heading ([1-9])$/i;

/** Heading level 1–9 from a paragraph style id or display name, if it names one. */
export function headingLevelForStyle(styleId: string | undefined, styles: StyleData): number | undefined {
  if (styleId === undefined) return undefined;
  const idMatch = HEADING_STYLE_ID.exec(styleId);
  if (idMatch !== null) return Number(idMatch[1]);
  const name = styles.paragraph.get(styleId)?.name;
  if (name === undefined) return undefined;
  const nameMatch = HEADING_STYLE_NAME.exec(name.trim());
  return nameMatch === null ? undefined : Number(nameMatch[1]);
}
