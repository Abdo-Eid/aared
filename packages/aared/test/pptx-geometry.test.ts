import { describe, expect, test } from "bun:test";
import { geometryForPreset } from "../src/core/drawingml-shape.js";

// Slide-3 proportions: 539640 × 251640 EMU (w ≈ 2.15 × h), like the real
// fixture shapes, so the ss-vs-w formula bases are actually distinguishable.
const W = 539640;
const H = 251640;

describe("DrawingML preset paths (ECMA-376 formulas)", () => {
  test("rtTriangle winds right-angle-bottom-left like Office", () => {
    expect(geometryForPreset("rtTriangle", { width: W, height: H })?.path).toBe(`M 0 ${H} L 0 0 L ${W} ${H} Z`);
  });

  test("triangle apex is width-based and clamped to the full width", () => {
    expect(geometryForPreset("triangle", { width: W, height: H }, { adj: 50000 })?.path).toBe(`M 0 ${H} L ${W / 2} 0 L ${W} ${H} Z`);
    expect(geometryForPreset("triangle", { width: W, height: H }, { adj: 120000 })?.path).toBe(`M 0 ${H} L ${W} 0 L ${W} ${H} Z`);
  });

  test("parallelogram skew uses the short side with the ECMA maxAdj clamp", () => {
    const geometry = geometryForPreset("parallelogram", { width: W, height: H }, { adj: 25000 });
    // ECMA: x2 = ss·a/1e5 = 62910, not w·a/1e5 = 134910.
    expect(geometry?.path).toBe(`M 0 ${H} L 62910 0 L ${W} 0 L ${W - 62910} ${H} Z`);
    // maxAdj = 100000·w/ss: a wide shape accepts adj values past 50000.
    const wide = geometryForPreset("parallelogram", { width: W, height: H }, { adj: 200000 });
    expect(wide?.path).toBe(`M 0 ${H} L ${2 * H} 0 L ${W} 0 L ${W - 2 * H} ${H} Z`);
  });

  test("hexagon inset is short-side based", () => {
    expect(geometryForPreset("hexagon", { width: W, height: H })?.path).toBe(
      `M 0 ${H / 2} L 62910 0 L ${W - 62910} 0 L ${W} ${H / 2} L ${W - 62910} ${H} L 62910 ${H} Z`,
    );
  });

  test("octagon keeps its short-side inset", () => {
    const inset = (H * 29289) / 100000;
    expect(geometryForPreset("octagon", { width: W, height: H })?.path).toBe(
      `M 0 ${inset} L ${inset} 0 L ${W - inset} 0 L ${W} ${inset} L ${W} ${H - inset} L ${W - inset} ${H} L ${inset} ${H} L 0 ${H - inset} Z`,
    );
  });

  test("ellipse draws two half-arcs between the horizontal extremes", () => {
    expect(geometryForPreset("ellipse", { width: W, height: H })?.path).toBe(
      `M 0 ${H / 2} A ${W / 2} ${H / 2} 0 0 1 ${W} ${H / 2} A ${W / 2} ${H / 2} 0 0 1 0 ${H / 2} Z`,
    );
  });

  test("line keeps the authored diagonal", () => {
    expect(geometryForPreset("line", { width: 648000, height: 360 })?.path).toBe("M 0 0 L 648000 360");
  });

  test("star5 draws ten points from the ECMA guides", () => {
    const geometry = geometryForPreset("star5", { width: 540000, height: 252000 });
    if (geometry === undefined) throw new Error("star5 missing");
    expect(geometry.preset).toBe("star5");
    expect(geometry.fallback).toBeUndefined();
    // Top point touches the box top; ten vertices + close.
    expect(geometry.path.split(" ").filter((token) => token === "L")).toHaveLength(9);
    expect(geometry.path.endsWith(" Z")).toBe(true);
    // The top point sits on the top edge: the hf/vf-enlarged ellipse center
    // (svc) minus its radius (shd2) is exactly 0.
    expect(geometry.path).toContain(`L ${540000 / 2} 0`);
  });

  test("unknown presets stay unresolved so the fallback marker applies", () => {
    expect(geometryForPreset("upArrow", { width: W, height: H })).toBeUndefined();
    expect(geometryForPreset("custGeom-like", { width: W, height: H })).toBeUndefined();
  });
});
