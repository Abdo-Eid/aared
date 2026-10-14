/**
 * Main-thread resolution of geometry the worker deferred: `custGeom`
 * programs, and presets outside the worker's hand-coded fast path.
 *
 * The session `import()`s this module only when a slide actually carries
 * deferred geometry. The preset table is not imported here: `@aared/viewer/pptx`
 * imports it statically and hands it to its session at open, so the table
 * ships only in the pptx entry and the flat worker never contains it.
 */
import type { ShapeGeometryModel } from "./model.js";
import { evaluateGeometry } from "./drawingml-engine.js";

/** The table's program (or alias) for `preset`; `undefined` when it is not catalogued. */
export function presetProgram(preset: string, programs: Readonly<Record<string, string>>): string | undefined {
  return Object.hasOwn(programs, preset) ? programs[preset] : undefined;
}

/**
 * Evaluate each geometry's deferred program in place: on success its `path`
 * (and any fill-only/stroke-only layers) replace the bounds fallback and the
 * `fallback` marker is dropped; otherwise the worker's visible fallback stays.
 * `deferred` is always removed, so consumers never see the transport field.
 */
export function resolveDeferredGeometry(geometries: readonly ShapeGeometryModel[], programs: Readonly<Record<string, string>>): void {
  for (const geometry of geometries) {
    const deferred = geometry.deferred;
    delete geometry.deferred;
    if (deferred === undefined) continue;
    let program = deferred.program ?? (geometry.preset === undefined ? undefined : presetProgram(geometry.preset, programs));
    // A catalog alias (`= x rightArrow`) is another preset under a symmetry
    // of the square; the compiler proved the equivalence and never chains.
    let symmetry = "";
    if (deferred.program === undefined && program?.startsWith("= ") === true) {
      const [, spelled, base] = program.split(" ");
      symmetry = spelled!;
      program = base === undefined ? undefined : presetProgram(base, programs);
      if (program?.startsWith("= ") === true) continue;
    }
    if (program === undefined) continue;
    const paths = evaluateGeometry(program, geometry.viewBox.width, geometry.viewBox.height, deferred.adjustments, symmetry);
    if (paths === undefined) continue;
    geometry.path = paths.path;
    if (paths.fillOnly !== undefined) geometry.fillOnly = paths.fillOnly;
    if (paths.strokeOnly !== undefined) geometry.strokeOnly = paths.strokeOnly;
    delete geometry.fallback;
  }
}
