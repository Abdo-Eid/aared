import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { geometryForPreset, unresolvedGeometry } from "../src/core/drawingml-shape.js";
import { evaluateGeometry } from "../src/core/drawingml-engine.js";
import { presetProgram, resolveDeferredGeometry } from "../src/core/drawingml-resolve.js";
import { collectDeferredGeometry } from "../src/core/collect.js";
import { presetPrograms } from "../src/core/presets.js";
import { parseXml } from "../src/core/xml.js";
import type { ShapeGeometryModel, SlideModel } from "../src/core/model.js";

interface OracleCase {
  preset: string;
  width: number;
  height: number;
  adjustments: Record<string, number>;
  paths: Array<Array<[string, number, number] | [string]>>;
}
const oracle = JSON.parse(readFileSync(new URL("./drawingml-oracle.json", import.meta.url), "utf8")) as { cases: OracleCase[] };
const seed = parseXml(readFileSync(new URL("../scripts/presetShapeDefinitions.xml", import.meta.url), "utf8"));
const W = 539640;
const H = 251640;
const catalogued = Object.keys(presetPrograms);
const programOf = (preset: string): string | undefined => presetProgram(preset, presetPrograms);

function pointsOf(d: string | undefined): Array<[number, number]> {
  if (d === undefined || d === "") return [];
  const arity: Record<string, number> = { M: 2, L: 2, Q: 4, C: 6, A: 7, Z: 0 };
  const tokens = d.split(" ");
  const out: Array<[number, number]> = [];
  for (let i = 0; i < tokens.length;) {
    const verb = tokens[i++]!;
    const args = tokens.slice(i, i + arity[verb]!).map(Number);
    i += arity[verb]!;
    if (verb === "Z") continue;
    out.push([args[args.length - 2]!, args[args.length - 1]!]);
  }
  return out;
}

describe("preset catalog", () => {
  test("carries every seed preset the fast path does not, and none it does", () => {
    const expected = seed.children
      .map((definition) => definition.local)
      .filter((name) => geometryForPreset(name, { width: 1, height: 1 }) === undefined);
    expect([...catalogued].sort()).toEqual([...expected].sort());
  });

  test("matches an independent implementation of the ECMA catalog, endpoint by endpoint", () => {
    expect(oracle.cases.length).toBeGreaterThan(300);
    for (const entry of oracle.cases) {
      let program = programOf(entry.preset);
      if (program === undefined) {
        // A fast-path preset (hand-coded, so not catalogued) answers to the
        // oracle too: same endpoints, in order.
        const fast = geometryForPreset(entry.preset, { width: entry.width, height: entry.height }, entry.adjustments);
        if (fast === undefined) throw new Error(`catalog lacks ${entry.preset}`);
        const want = entry.paths.flat().filter((command) => command.length === 3).map((command) => [command[1], command[2]]);
        expect(pointsOf(fast.path)).toEqual(want as Array<[number, number]>);
        continue;
      }
      // Aliases render another preset under a symmetry; endpoints compare
      // as multisets, so a mirrored start vertex still matches.
      let symmetry = "";
      if (program.startsWith("= ")) {
        const [, spelled, base] = program.split(" ");
        symmetry = spelled!;
        program = programOf(base!)!;
        expect(program.startsWith("= ")).toBe(false);
      }
      // Evaluate the whole program at once (segmentation differs by design:
      // mid-path closes, LO subpath splits, full-turn halves), flatten every
      // paint bucket into one endpoint multiset, and compare verb-insensitively.
      const label = `${entry.preset} ${entry.width}x${entry.height} ${JSON.stringify(entry.adjustments)}`;
      const evaluated = evaluateGeometry(program, entry.width, entry.height, entry.adjustments, symmetry);
      if (evaluated === undefined) throw new Error(`engine returned undefined for ${label}`);
      const got = [...pointsOf(evaluated.path), ...pointsOf(evaluated.fillOnly), ...pointsOf(evaluated.strokeOnly)];
      // An ECMA path flagged neither-fill-nor-stroke paints nothing; when
      // LibreOffice's segmentation aligns 1:1 with ECMA's, its record of
      // that subpath is not compared (flowChartMultidocument's invisible
      // third path).
      const ecmaPaths = program.split(";").filter((statement) => statement.startsWith("p "));
      const aligned = ecmaPaths.length === entry.paths.length;
      const want: Array<[number, number]> = [];
      entry.paths.forEach((subpath, index) => {
        const flags = ecmaPaths[index]?.split(" ")[3] ?? "-";
        if (aligned && flags.includes("f") && flags.includes("s")) return;
        for (const command of subpath) if (command.length === 3) want.push([command[1], command[2]]);
      });
      expect(`${label}: ${got.length}`).toBe(`${label}: ${want.length}`);
      const pool = [...want];
      let worst = 0;
      for (const [gx, gy] of got) {
        let best = -1, bestD = Infinity;
        for (let i = 0; i < pool.length; i++) {
          const d = Math.hypot(pool[i]![0] - gx, pool[i]![1] - gy);
          if (d < bestD) { bestD = d; best = i; }
        }
        worst = Math.max(worst, bestD);
        pool.splice(best, 1);
      }
      expect(`${label}: worst ${worst.toFixed(2)}`).toBe(`${label}: worst ${Math.min(worst, 2).toFixed(2)}`);
    }
  });

  test("stores proven symmetric images as aliases that never chain", () => {
    const aliases = catalogued.map((name) => [name, programOf(name)] as const);
    const aliased = aliases.filter(([, program]) => program!.startsWith("= "));
    expect(aliased.map(([name]) => name).sort()).toEqual(["actionButtonBackPrevious", "actionButtonBeginning", "actionButtonBlank", "curvedLeftArrow", "downArrow", "downArrowCallout", "flowChartExtract", "horizontalScroll", "leftBrace", "leftBracket", "leftRightArrow", "leftRightArrowCallout", "rightArrow", "rightArrowCallout", "upArrow", "upArrowCallout"]);
    for (const [, alias] of aliased) {
      const base = alias!.split(" ")[2]!;
      expect(programOf(base)!.startsWith("= ")).toBe(false);
    }
  });

  test("resolves an alias to exactly the outline of the definition it replaced", async () => {
    const original = seed.children.find((definition) => definition.local === "rightArrow")!;
    const { encodeGeometry } = await import("../src/core/drawingml-program.js");
    const direct = evaluateGeometry(encodeGeometry(original)!, W, H, { adj1: 70000, adj2: 90000 })!.path;
    const viaAlias = unresolvedGeometry({ width: W, height: H }, "rightArrow", { adj1: 70000, adj2: 90000 }, undefined);
    resolveDeferredGeometry([viaAlias], presetPrograms);
    const corners = (d: string): string => d.replace(/[MLZ]/g, " ").trim().split(/\s+/).reduce<string[]>((pairs, value, index, all) => (index % 2 === 0 ? [...pairs, `${Math.round(Number(value))},${Math.round(Number(all[index + 1]))}`] : pairs), []).sort().join(" ");
    expect(corners(viaAlias.path)).toBe(corners(direct));
  });

  test("looks presets up by name and reports unknown presets as absent", () => {
    expect(programOf("rightArrow")).toBe("= x leftArrow");
    expect(programOf("leftArrow")).toContain("p - - - M l vc");
    expect(programOf("noSuchPreset")).toBeUndefined();
    // Object prototype keys are not presets.
    expect(programOf("constructor")).toBeUndefined();
    // The fast-path presets never enter the catalog.
    expect(programOf("rect")).toBeUndefined();
  });
});

describe("deferred geometry resolution", () => {
  test("the worker defers unknown presets with their authored adjustments", () => {
    const geometry = unresolvedGeometry({ width: 1000, height: 500 }, "rightArrow", { adj1: 20000 }, undefined);
    expect(geometry.fallback).toBe("unsupported");
    expect(geometry.preset).toBe("rightArrow");
    expect(geometry.deferred).toEqual({ adjustments: { adj1: 20000 } });
    expect(geometry.path).toBe("M 0 0 H 1000 V 500 H 0 Z");
  });

  test("the worker encodes custGeom, and keeps a plain fallback when it cannot", () => {
    const custom = parseXml(`<custGeom><pathLst><path w="2" h="2"><moveTo><pt x="0" y="0"/></moveTo><lnTo><pt x="2" y="1"/></lnTo><lnTo><pt x="0" y="2"/></lnTo><close/></path></pathLst></custGeom>`);
    expect(unresolvedGeometry({ width: 10, height: 10 }, undefined, {}, custom).deferred).toEqual({ program: "p 2 2 - M 0 0 L 2 1 L 0 2 Z" });
    const empty = unresolvedGeometry({ width: 10, height: 10 }, undefined, {}, parseXml("<custGeom/>"));
    expect(empty).toEqual({ path: "M 0 0 H 10 V 10 H 0 Z", viewBox: { width: 10, height: 10 }, fallback: "custom" });
    expect(unresolvedGeometry(undefined, undefined, {}, undefined).fallback).toBe("missing");
  });

  test("finds deferred geometry anywhere in a slide, including nested groups", () => {
    const deferred = unresolvedGeometry({ width: 10, height: 10 }, "chevron", {}, undefined);
    const slide = {
      part: "/ppt/slides/slide1.xml",
      size: { width: 10, height: 7.5 },
      elements: [
        { kind: "shape", id: "1", name: "a", geometry: geometryForPreset("rect", { width: 10, height: 10 }) },
        { kind: "group", id: "2", name: "g", children: [{ kind: "shape", id: "3", name: "b", geometry: deferred }] },
      ],
    } as unknown as SlideModel;
    expect(collectDeferredGeometry(slide)).toEqual([deferred]);
  });

  test("resolves presets and custGeom in place and drops the fallback markers", () => {
    const arrow = unresolvedGeometry({ width: 539640, height: 251640 }, "rightArrow", {}, undefined);
    const custom = unresolvedGeometry(
      { width: 100, height: 100 },
      undefined,
      {},
      parseXml(`<custGeom><pathLst><path><moveTo><pt x="l" y="t"/></moveTo><lnTo><pt x="r" y="b"/></lnTo><lnTo><pt x="l" y="b"/></lnTo><close/></path></pathLst></custGeom>`),
    );
    const can = unresolvedGeometry({ width: 400000, height: 600000 }, "can", { adj: 30000 }, undefined);
    resolveDeferredGeometry([arrow, custom, can], presetPrograms);
    // rightArrow is stored as the mirror of leftArrow: same seven vertices,
    // starting at the tip. (Winding reverses; the fill is unaffected.)
    const corners = (d: string): string[] => [...d.matchAll(/[ML] (\S+) (\S+)/g)].map((m) => `${m[1]} ${m[2]}`).sort();
    expect(corners(arrow.path)).toEqual(corners("M 0 62910 L 413820 62910 L 413820 0 L 539640 125820 L 413820 251640 L 413820 188730 L 0 188730 Z"));
    expect(arrow.preset).toBe("rightArrow");
    expect(arrow.fallback).toBeUndefined();
    expect(arrow.deferred).toBeUndefined();
    expect(custom).toEqual({ path: "M 0 0 L 100 100 L 0 100 Z", viewBox: { width: 100, height: 100 } });
    expect(can.fallback).toBeUndefined();
    expect(can.fillOnly).toBeDefined();
    expect(can.strokeOnly).toBeDefined();
  });

  test("anything it cannot resolve keeps the worker's fallback, minus the transport field", () => {
    const unknown = unresolvedGeometry({ width: 10, height: 10 }, "noSuchPreset", {}, undefined);
    const broken: ShapeGeometryModel = { ...unresolvedGeometry({ width: 10, height: 10 }, undefined, {}, undefined), fallback: "custom", deferred: { program: "p - - - M nope 0" } };
    resolveDeferredGeometry([unknown, broken], presetPrograms);
    expect(unknown).toEqual({ path: "M 0 0 H 10 V 10 H 0 Z", viewBox: { width: 10, height: 10 }, fallback: "unsupported", preset: "noSuchPreset" });
    expect(broken).toEqual({ path: "M 0 0 H 10 V 10 H 0 Z", viewBox: { width: 10, height: 10 }, fallback: "custom" });
  });
});
