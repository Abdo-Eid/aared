import { attr, childElements, firstChild } from "./xml.js";
import type { XmlElement } from "./xml.js";

/**
 * DrawingML geometry as a compact, structured-cloneable program string.
 *
 * `a:custGeom` in a slide and every entry of ECMA-376's
 * `presetShapeDefinitions.xml` are written in the same guide language
 * (`avLst`, `gdLst`, `pathLst`). This encoder turns either into one string so
 * a single evaluator (`drawingml-engine.ts`) serves both: the worker encodes
 * a slide's `custGeom` here, and `scripts/compile-presets.ts` encodes the ECMA
 * catalog with this same function at build time.
 *
 * Grammar — statements joined by `;`, tokens by one space:
 *
 *   a NAME VALUE           adjust-value default (`avLst` `val` entries)
 *   g NAME OP ARG…         guide, formula tokens verbatim (`*\/ w adj 100000`)
 *   p W H FLAGS CMD…       one path; W/H are its coordinate space or `-` for
 *                          the shape's own; FLAGS is `-` or any of `f` (no
 *                          fill) and `s` (no stroke)
 *
 * Path commands take guide names or numeric literals:
 *   M x y | L x y | A wR hR stAng swAng | Q x1 y1 x y | C x1 y1 x2 y2 x y | Z
 *
 * Internal to aared; the format is versioned with the catalog it ships with.
 */

const SAFE_TOKEN = /^[^\s;]+$/;

function tokens(value: string | undefined): string[] | undefined {
  if (value === undefined) return undefined;
  const parts = value.trim().split(/\s+/);
  return parts.length > 0 && parts.every((part) => SAFE_TOKEN.test(part)) ? parts : undefined;
}

function point(el: XmlElement | undefined): string[] | undefined {
  if (el === undefined) return undefined;
  const x = tokens(attr(el, "x"));
  const y = tokens(attr(el, "y"));
  return x?.length === 1 && y?.length === 1 ? [x[0]!, y[0]!] : undefined;
}

function pathCommands(path: XmlElement): string[] | undefined {
  const out: string[] = [];
  for (const command of path.children) {
    const pts = childElements(command, "pt").map(point);
    if (pts.some((pt) => pt === undefined)) return undefined;
    const coords = (pts as string[][]).flat();
    switch (command.local) {
      case "moveTo":
      case "lnTo":
        if (coords.length !== 2) return undefined;
        out.push(command.local === "moveTo" ? "M" : "L", ...coords);
        break;
      case "quadBezTo":
        if (coords.length !== 4) return undefined;
        out.push("Q", ...coords);
        break;
      case "cubicBezTo":
        if (coords.length !== 6) return undefined;
        out.push("C", ...coords);
        break;
      case "arcTo": {
        const args = ["wR", "hR", "stAng", "swAng"].map((name) => tokens(attr(command, name)));
        if (args.some((arg) => arg?.length !== 1)) return undefined;
        out.push("A", ...args.map((arg) => arg![0]!));
        break;
      }
      case "close":
        out.push("Z");
        break;
      default:
        // Unknown path verbs would silently change the outline: refuse.
        return undefined;
    }
  }
  return out;
}

/**
 * Encode a `custGeom` (or an ECMA preset definition element with the same
 * children) as a geometry program. Returns `undefined` when there is nothing
 * drawable or the source uses constructs the grammar cannot carry, so the
 * caller keeps its visible fallback instead of painting a wrong outline.
 */
export function encodeGeometry(geometry: XmlElement): string | undefined {
  const statements: string[] = [];
  const avLst = firstChild(geometry, "avLst");
  for (const guide of avLst === undefined ? [] : childElements(avLst, "gd")) {
    const name = tokens(attr(guide, "name"));
    const formula = tokens(attr(guide, "fmla"));
    if (name?.length !== 1 || formula === undefined) return undefined;
    statements.push(formula[0] === "val" && formula.length === 2 ? `a ${name[0]} ${formula[1]}` : `g ${name[0]} ${formula.join(" ")}`);
  }
  const gdLst = firstChild(geometry, "gdLst");
  for (const guide of gdLst === undefined ? [] : childElements(gdLst, "gd")) {
    const name = tokens(attr(guide, "name"));
    const formula = tokens(attr(guide, "fmla"));
    if (name?.length !== 1 || formula === undefined) return undefined;
    statements.push(`g ${name[0]} ${formula.join(" ")}`);
  }
  const pathLst = firstChild(geometry, "pathLst");
  let drawable = false;
  for (const path of pathLst === undefined ? [] : childElements(pathLst, "path")) {
    const commands = pathCommands(path);
    if (commands === undefined) return undefined;
    if (commands.length === 0) continue;
    const size = (name: string): string => {
      const value = Number(attr(path, name) ?? "0");
      return Number.isFinite(value) && value > 0 ? String(value) : "-";
    };
    const stroke = attr(path, "stroke");
    const flags = (attr(path, "fill") === "none" ? "f" : "") + (stroke === "false" || stroke === "0" ? "s" : "");
    statements.push(["p", size("w"), size("h"), flags === "" ? "-" : flags, ...commands].join(" "));
    drawable = true;
  }
  return drawable ? statements.join(";") : undefined;
}
