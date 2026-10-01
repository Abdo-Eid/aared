/**
 * Which fonts a render actually painted with: load the faces a model's text
 * uses (`awaitDocumentFonts`), then measure whether each family painted with
 * something real or fell to the browser's own fallback (`detectFontFidelity`,
 * `fontFidelityWarnings`). Renderer-side, used by all three formats. Its own
 * module, apart from `text-style.ts`'s paragraph painting, so a format that
 * only needs the check (XLSX) does not download PPTX and DOCX text styling
 * with it: the bundler splits shared code by module, not by function.
 */

import type { RenderWarning } from "./core/model.js";
import { fontAlias, genericFontFamily, resolvesLikeGenericDefault } from "./core/typography.js";
import { collectFontUse } from "./core/collect.js";

// A handful of Arabic letters (U+0628-062E) so the width/subset probes below
// exercise that script too. Spelled through JSON.parse, like core/chars.ts,
// so no build step inlines the literal characters into the shipped script
// (see that file's doc comment — a wide character there makes the whole
// script two bytes per character in V8). Not semantically named individual
// glyphs (chars.ts is for those), so it stays local to this one sample.
const ARABIC_SAMPLE_LETTERS = /* @__PURE__ */ (JSON.parse('"\\u0628\\u062a\\u062b\\u062c\\u062d\\u062e"') as string);
/** Broad enough to trigger whichever `unicode-range` subset a registered face actually needs. */
const FONT_LOAD_SAMPLE_LATIN = "AaBbXxYyZz0129";
const FONT_LOAD_SAMPLE = `${FONT_LOAD_SAMPLE_LATIN} ${ARABIC_SAMPLE_LETTERS}`;
const FONT_LOAD_TIMEOUT_MS = 2000;

/**
 * Load every face a model's text paints in — each family in each variant it
 * uses (bold, italic), under its requested name and any delivered
 * substitute's alias — and wait for them, before a renderer measures painted
 * text (DOCX layout pagination's off-screen probe) or checks which fonts
 * painted (`detectFontFidelity`). A `FontResolver`'s delivered faces are
 * registered on `document.fonts` by the time a model reaches `render()` (the
 * session awaits the resolver's own promise settling, `core/session.ts`), but
 * `aared-fonts` never loads them itself: the browser fetches a face only when
 * text first needs it, so a document downloads only the subsets it paints
 * (DECISION_LOG §50). Measuring before the real bytes arrive reads whatever
 * was painted first instead — an installed real font, or the browser's own
 * unrelated fallback — which differs from what paints a moment later once
 * the intended face is ready (DECISION_LOG §45: this is what let pagination
 * measure with a swapped-in font after the page count was already decided).
 * Loading through `document.fonts.load` with a sample is also what makes the
 * download happen for exactly those faces and subsets.
 *
 * `model` is walked generically (`collectFontUse`), so this serves any
 * format's model shape. XLSX calls it only for the fidelity check: its row
 * heights and column widths come from the file, not a live measurement.
 *
 * Bounded by `timeoutMs`: an offline or blocked CDN must not hang a render,
 * so the caller measures whatever is ready at the deadline — the same
 * outcome as before this existed. Headless (no `document.fonts.load`) is a
 * no-op.
 */
export async function awaitDocumentFonts(model: unknown, timeoutMs = FONT_LOAD_TIMEOUT_MS): Promise<void> {
  if (typeof document === "undefined" || typeof document.fonts?.load !== "function") return;
  const { families, arabic } = collectFontUse(model);
  if (families.size === 0) return;
  // Arabic letters join the sample only when the document has Arabic text: they are what pulls in
  // an Arabic-limited fallback face (Arial's, Times New Roman's), which a Latin-only document
  // must not download or wait for.
  const sample = arabic ? FONT_LOAD_SAMPLE : FONT_LOAD_SAMPLE_LATIN;
  const loads: Array<Promise<unknown>> = [];
  for (const [family, variants] of families) {
    for (const variant of variants) {
      for (const name of [family, fontAlias(family)]) {
        loads.push(document.fonts.load(`${variant}16px "${name.replace(/"/g, '\\"')}"`, sample).catch(() => undefined));
      }
    }
  }
  await Promise.race([
    Promise.allSettled(loads).then(() => undefined),
    new Promise<void>((resolve) => setTimeout(resolve, timeoutMs)),
  ]);
}

export interface FontFidelity {
  family: string;
  /**
   * `"available"`: the requested name resolved to something real, installed
   * locally. `"substituted"`: nothing answered to the requested name, but a
   * delivered alias (see `fontAlias`) painted instead. `"unavailable"`:
   * neither — this family painted with the browser's own generic default,
   * unrelated to what the document asked for.
   */
  status: "available" | "substituted" | "unavailable";
}

let fidelityCanvas: HTMLCanvasElement | undefined;

/** A font stack's rendered width for `sample`, via `measureText` (real text shaping, no DOM insertion or reflow). */
function measuredWidth(fontSpec: string, sample: string): number {
  fidelityCanvas ??= document.createElement("canvas");
  const context = fidelityCanvas.getContext("2d");
  if (context === null) return 0;
  context.font = fontSpec;
  return context.measureText(sample).width;
}

/**
 * Whether each font family a model's text uses actually painted with
 * something real — the requested name itself, or a delivered substitute's
 * alias — or silently fell to the browser's own unrelated generic default.
 * Measured, not asked: `document.fonts.check()` cannot answer this at all —
 * verified empirically (`tools/probe/font-metrics/`), it reports `true` even
 * for a family name that exists nowhere, apparently whenever *anything*
 * could paint the fallback stack, not whether the named family specifically
 * resolved. This instead compares a rendered width against a
 * same-generic-category control (`genericFontFamily`'s classification, so a
 * genuinely serif face is never compared against a sans-serif default) that
 * is guaranteed never to resolve to anything of its own — a family that
 * measures the same as its own control never painted with anything of its
 * own. One real exception: Arial, Times New Roman and Courier New are
 * themselves what a browser's `sans-serif`/`serif`/`monospace` keyword
 * commonly already resolves to (`resolvesLikeGenericDefault` — this is *why*
 * those three were picked as the generics' archetypes), so measuring the
 * same as the control there is inconclusive, not evidence of absence: the
 * pixels are right regardless of which path produced them, and calling it
 * "unavailable" would be a false alarm on the most common Latin fonts of
 * all. Call once `awaitDocumentFonts` has settled, or a substitute mid-load
 * reads as unavailable.
 */
export function detectFontFidelity(model: unknown): FontFidelity[] {
  if (typeof document === "undefined") return [];
  const { families, arabic } = collectFontUse(model);
  // Measuring text through a font stack makes the browser fetch whichever face covers it, so the
  // Arabic letters are in the sample only when the document has Arabic to paint — the same rule
  // as `awaitDocumentFonts`, and for the same reason: a Latin-only document must not download
  // the Arabic-limited fallback face just to be checked.
  const sample = arabic ? FONT_LOAD_SAMPLE : FONT_LOAD_SAMPLE_LATIN;
  const out: FontFidelity[] = [];
  for (const [family, variants] of families) {
    if (resolvesLikeGenericDefault(family)) {
      out.push({ family, status: "available" });
      continue;
    }
    // Measured in a variant the document paints, the one `awaitDocumentFonts` loaded: a
    // bold-only family's regular face may never have been fetched.
    const variant = variants.values().next().value ?? "";
    const generic = genericFontFamily(family);
    const clean = family.replace(/["\\]/g, "");
    const controlWidth = measuredWidth(`${variant}16px ${generic}`, sample);
    const requestedWidth = measuredWidth(`${variant}16px "${clean}", ${generic}`, sample);
    if (Math.abs(requestedWidth - controlWidth) > 0.5) {
      out.push({ family, status: "available" });
      continue;
    }
    const aliasWidth = measuredWidth(`${variant}16px "${fontAlias(clean)}", ${generic}`, sample);
    out.push({ family, status: Math.abs(aliasWidth - controlWidth) > 0.5 ? "substituted" : "unavailable" });
  }
  return out;
}

/**
 * `RenderWarning`s for every family that painted with the browser's own
 * unrelated fallback — not even a delivered substitute. A working
 * substitution (`"substituted"` in `detectFontFidelity`) is not itself a
 * warning: it is the intended, working case for a font nobody has installed.
 */
export function fontFidelityWarnings(model: unknown): RenderWarning[] {
  return detectFontFidelity(model)
    .filter((entry) => entry.status === "unavailable")
    .map((entry) => ({
      code: "font-unavailable",
      message: `"${entry.family}" is not installed and no resolver delivered a substitute for it; text in this family is approximated with the browser's own fallback font.`,
    }));
}
