/**
 * Compile ECMA-376 `presetShapeDefinitions.xml` into the preset table
 * `src/core/presets.ts`.
 *
 *   bun scripts/compile-presets.ts path/to/presetShapeDefinitions.xml
 *   bun scripts/compile-presets.ts --check scripts/presetShapeDefinitions.xml
 *
 * `--check` regenerates in memory and fails if the committed table differs;
 * `test/generated-drift.test.ts` runs it, so an encoder change cannot leave
 * `src/core/presets.ts` stale.
 *
 * The full file ships in ECMA-376 Part 1's OfficeOpenXML-DrawingMLGeometries.zip.
 * `scripts/presetShapeDefinitions.xml` is a derived copy of it in the same
 * layout; its header names the one entry (`upArrow`) taken from LibreOffice.
 *
 * Each definition goes through the same `encodeGeometry` the worker applies
 * to a slide's `a:custGeom`, so catalog entries and custom geometry share one
 * program format and one evaluator. Presets the worker's hand-coded fast
 * path already resolves are left out: the table only carries what would
 * otherwise render as a fallback marker. The table is one static module that
 * only `aared/pptx` imports (DECISION_LOG §26 addendum), so xlsx/docx apps
 * never carry it and nothing downloads lazily.
 */
import { readFileSync, writeFileSync } from "node:fs";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { parseXml } from "../src/core/xml.js";
import { encodeGeometry } from "../src/core/drawingml-program.js";
import { geometryForPreset } from "../src/core/drawingml-shape.js";
import { evaluateGeometry } from "../src/core/drawingml-engine.js";

/**
 * Drop guides no path depends on. ECMA definitions also carry text-rectangle,
 * handle and connection-site guides (`il`, `ir`, `it`, `ib`…) that never
 * affect the outline; shipping them would only grow the table.
 *
 * Liveness runs to a fixpoint: a definition kept by name must keep its own
 * dependencies too, even when the name only becomes live through an earlier
 * statement visited later in the backward walk (circularArrow's doubled
 * dxI/dyI: the surviving late definition references wtI/htI, which an
 * incomplete single pass drops — and then every evaluation throws).
 * Duplicate names are kept as-is; the evaluator (like Office) processes
 * guides in order with later definitions winning.
 */
function pruneDeadGuides(program: string): string {
  const statements = program.split(";");
  const live = new Set<string>();
  for (const statement of statements) if (statement.startsWith("p ")) for (const token of statement.split(" ").slice(4)) live.add(token);
  let grew = true;
  while (grew) {
    grew = false;
    for (let index = statements.length - 1; index >= 0; index--) {
      const [kind, name, , ...args] = statements[index]!.split(" ");
      if (kind !== "g" || !live.has(name!)) continue;
      for (const token of args) {
        if (!live.has(token)) {
          live.add(token);
          grew = true;
        }
      }
    }
  }
  return statements.filter((statement) => {
    const [kind, name] = statement.split(" ");
    // Adjust values stay: an authored avLst may name them even when unused.
    return kind !== "g" || live.has(name!);
  }).join(";");
}

/**
 * A geometric fingerprint of rendered path data: every endpoint plus every
 * segment's midpoint (line midpoints pin down connectivity; arc and Bézier
 * midpoints pin down bulge direction), tagged with its paint layer. Two
 * outlines with the same corners but different edges — a plaque and a round
 * rect, an open elbow and a closed triangle — get different fingerprints.
 */
function fingerprint(layers: Record<string, string | undefined>): string {
  const marks: string[] = [];
  const mark = (layer: string, kind: string, x: number, y: number): void => {
    marks.push(`${layer}${kind}${Math.round(x / 40)},${Math.round(y / 40)}`);
  };
  for (const [layer, d] of Object.entries(layers)) {
    if (d === undefined || d === "") continue;
    const t = d.split(" ");
    let x = 0, y = 0, sx = 0, sy = 0;
    for (let i = 0; i < t.length;) {
      const verb = t[i++];
      const n = (count: number): number[] => t.slice(i, (i += count)).map(Number);
      if (verb === "M") { [x, y] = n(2); sx = x; sy = y; mark(layer, "p", x, y); }
      else if (verb === "L") { const [nx, ny] = n(2); mark(layer, "m", (x + nx!) / 2, (y + ny!) / 2); x = nx!; y = ny!; mark(layer, "p", x, y); }
      else if (verb === "Q") { const [x1, y1, nx, ny] = n(4); mark(layer, "m", (x + 2 * x1! + nx!) / 4, (y + 2 * y1! + ny!) / 4); x = nx!; y = ny!; mark(layer, "p", x, y); }
      else if (verb === "C") { const [x1, y1, x2, y2, nx, ny] = n(6); mark(layer, "m", (x + 3 * x1! + 3 * x2! + nx!) / 8, (y + 3 * y1! + 3 * y2! + ny!) / 8); x = nx!; y = ny!; mark(layer, "p", x, y); }
      else if (verb === "A") {
        const [rx, ry, , large, sweep, nx, ny] = n(7);
        const mid = arcMidpoint(x, y, rx!, ry!, large === 1, sweep === 1, nx!, ny!);
        mark(layer, "m", mid[0], mid[1]);
        x = nx!; y = ny!; mark(layer, "p", x, y);
      } else if (verb === "Z") {
        if (Math.abs(x - sx) > 1 || Math.abs(y - sy) > 1) mark(layer, "m", (x + sx) / 2, (y + sy) / 2);
        x = sx; y = sy;
      }
    }
  }
  return marks.sort().join(" ");
}

/** Midpoint of an axis-aligned SVG arc (SVG 1.1 F.6.5 endpoint-to-centre). */
function arcMidpoint(x1: number, y1: number, rx: number, ry: number, large: boolean, sweep: boolean, x2: number, y2: number): [number, number] {
  const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
  const scale = Math.max(1, Math.sqrt((dx * dx) / (rx * rx) + (dy * dy) / (ry * ry)));
  rx *= scale; ry *= scale;
  const num = rx * rx * ry * ry - rx * rx * dy * dy - ry * ry * dx * dx;
  const coef = (large === sweep ? -1 : 1) * Math.sqrt(Math.max(0, num / (rx * rx * dy * dy + ry * ry * dx * dx)));
  const cx = coef * ((rx * dy) / ry) + (x1 + x2) / 2;
  const cy = coef * (-(ry * dx) / rx) + (y1 + y2) / 2;
  const a1 = Math.atan2((y1 - cy) / ry, (x1 - cx) / rx);
  let delta = Math.atan2((y2 - cy) / ry, (x2 - cx) / rx) - a1;
  if (sweep && delta < 0) delta += 2 * Math.PI;
  if (!sweep && delta > 0) delta -= 2 * Math.PI;
  return [cx + rx * Math.cos(a1 + delta / 2), cy + ry * Math.sin(a1 + delta / 2)];
}

const SYMMETRIES = ["x", "y", "xy", "t", "tx", "ty", "txy"];

/** The `avLst` defaults of a program, which an alias must share with its base. */
function adjustDefaults(program: string): string {
  return program.split(";").filter((statement) => statement.startsWith("a ")).sort().join(";");
}

/**
 * Find presets that are a symmetry of another catalogued preset, verified by
 * evaluating both through the engine at several sizes and adjustments. Only
 * proven equivalences become aliases; a base is never itself an alias.
 */
function discoverAliases(programs: Map<string, string>): Map<string, string> {
  const samples: Array<[number, number, number]> = [[539640, 251640, 1], [251640, 739640, 0.7], [400000, 400000, 1.25], [812345, 190000, 0.55]];
  const aliases = new Map<string, string>();
  const bases = new Set<string>();
  const names = [...programs.keys()].sort();
  const render = (program: string, w: number, h: number, k: number, symmetry = ""): string | undefined => {
    const adjustments: Record<string, number> = {};
    for (const statement of program.split(";")) {
      const [kind, name, value] = statement.split(" ");
      if (kind === "a") adjustments[name!] = Number(value) * k;
    }
    const paths = evaluateGeometry(program, w, h, adjustments, symmetry);
    return paths === undefined ? undefined : fingerprint({ b: paths.path, f: paths.fillOnly, s: paths.strokeOnly });
  };
  for (const target of names) {
    if (bases.has(target)) continue;
    const own = programs.get(target)!;
    for (const base of names) {
      if (base === target || aliases.has(base) || adjustDefaults(programs.get(base)!) !== adjustDefaults(own)) continue;
      const symmetry = SYMMETRIES.find((candidate) => samples.every(([w, h, k]) => {
        const expected = render(own, w, h, k);
        return expected !== undefined && expected === render(programs.get(base)!, w, h, k, candidate);
      }));
      if (symmetry !== undefined) {
        aliases.set(target, `= ${symmetry} ${base}`);
        bases.add(base);
        break;
      }
    }
  }
  return aliases;
}

const check = process.argv.includes("--check");
const source = process.argv.slice(2).find((argument) => argument !== "--check");
if (source === undefined) {
  console.error("usage: compile-presets [--check] <presetShapeDefinitions.xml>");
  process.exit(1);
}
const out = join(dirname(fileURLToPath(import.meta.url)), "../src/core/presets.ts");
const root = parseXml(readFileSync(source, "utf8"));

const programs = new Map<string, string>();
const nativelyResolved: string[] = [];
const rejected: string[] = [];
for (const definition of root.children) {
  const name = definition.local;
  if (geometryForPreset(name, { width: 100000, height: 100000 }) !== undefined) {
    nativelyResolved.push(name);
    continue;
  }
  const encoded = encodeGeometry(definition);
  const program = encoded === undefined ? undefined : pruneDeadGuides(encoded);
  if (program === undefined) {
    rejected.push(name);
    continue;
  }
  programs.set(name, program);
}

const aliases = discoverAliases(programs);
const entries = [...programs].map(([name, program]) => [name, aliases.get(name) ?? program] as const).sort(([a], [b]) => a.localeCompare(b));
const body = entries.map(([name, program]) => `  ${JSON.stringify(name)}: ${JSON.stringify(program)},`).join("\n");
const text = `// Generated by scripts/compile-presets.ts from ${basename(source)}. Do not edit.

/**
 * Every ECMA-376 preset outside the worker's hand-coded fast path, as an
 * engine program or a proven symmetry alias (\`= <symmetry> <base>\`), keyed
 * by preset name. Only \`aared/pptx\` imports this, passing it to its session
 * at open; the flat worker and the xlsx/docx entries never contain it.
 */
export const presetPrograms: Readonly<Record<string, string>> = {
${body}
};
`;
if (check) {
  // `--check` writes nothing: it fails when the committed table is not what the
  // source XML and the current encoder produce (line endings aside).
  if (readFileSync(out, "utf8").replace(/\r\n/g, "\n") !== text) {
    console.error(`${out} is out of date: run \`bun scripts/compile-presets.ts ${source}\` and commit the result`);
    process.exit(1);
  }
  console.log(`${entries.length} presets match ${out}`);
  process.exit(0);
}
writeFileSync(out, text);

console.log(`${entries.length} presets, ${text.length} bytes → ${out}`);
if (aliases.size > 0) console.log(`symmetry aliases (${aliases.size}): ${[...aliases].map(([name, alias]) => `${name} ${alias}`).join(", ")}`);
console.log(`fast path (not catalogued): ${nativelyResolved.length} — ${nativelyResolved.join(" ")}`);
if (rejected.length > 0) console.log(`rejected (kept as fallback): ${rejected.join(" ")}`);
