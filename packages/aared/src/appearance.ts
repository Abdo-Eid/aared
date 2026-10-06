/**
 * Viewer appearance shared by the format renderers: the light/dark resolve and
 * the design tokens each renderer publishes as `--aared-*` custom properties.
 *
 * Renderer-side (it reads `matchMedia`), never imported by `core`. The tokens
 * are renderer-owned design tokens, not consumer-overridable variables; the
 * consumer-facing inputs (`--aared-background`, `--aared-text`,
 * `--aared-page-shadow`) belong to `<aared-viewer>` alone, which is why the
 * text colour is published as `--aared-ink` and never as `--aared-text`.
 */

import { contrastRatio } from "./core/color.js";
import type { ViewerTheme } from "./core/index.js";

export type ResolvedTheme = "light" | "dark";

/** `system` follows the OS colour scheme; any other value is itself. Hosts without `matchMedia` resolve `system` to light. */
export function resolveTheme(theme: ViewerTheme | undefined): ResolvedTheme {
  if (theme === "dark") return "dark";
  if (theme === "system" && typeof matchMedia === "function" && matchMedia("(prefers-color-scheme: dark)").matches) return "dark";
  return "light";
}

/**
 * Call `onChange` each time the OS colour scheme flips, until the returned
 * function is called. The one place a viewer listens for it: `system` is
 * resolved when something paints, so whoever holds a `system` theme asks to
 * be told when that answer may have changed. A host without `matchMedia`
 * resolves `system` to light and never flips.
 */
export function watchSystemTheme(onChange: () => void): () => void {
  if (typeof matchMedia !== "function") return () => {};
  const query = matchMedia("(prefers-color-scheme: dark)");
  query.addEventListener("change", onChange);
  return () => query.removeEventListener("change", onChange);
}

interface Palette {
  /** Chrome surface for the resolved theme; document content always paints light source. */
  surface: string;
  /** Default text on that surface. */
  ink: string;
  border: string;
  /** Secondary text: notices, captions. */
  muted: string;
  /** Selection and focus. */
  accent: string;
}

/**
 * One palette for every format. The dark entries serve viewer chrome only
 * (tabs, notices, selection); document content always paints light source.
 * The dark surface is neutral so authored greys are not tinted; each muted
 * and accent value is the higher-contrast of the pair the renderers used to
 * paint separately.
 */
export const PALETTE: Record<ResolvedTheme, Palette> = {
  light: { surface: "#ffffff", ink: "#000000", border: "#d1d5db", muted: "#5b6470", accent: "#0b57d0" },
  dark: { surface: "#1a1a1a", ink: "#e8e8e8", border: "#3a3a3a", muted: "#a7aab3", accent: "#8ab4f8" },
};

/**
 * Automatic text over an authored `fill`: whichever palette ink reads better
 * on it, as Word's and Excel's Automatic colour does. A fill that does not
 * parse keeps the light ink.
 */
export function automaticInk(fill: string): string {
  const onLight = contrastRatio(fill, PALETTE.dark.ink);
  const onDark = contrastRatio(fill, PALETTE.light.ink);
  if (onLight === undefined || onDark === undefined) return PALETTE.dark.ink;
  return onDark >= onLight ? PALETTE.light.ink : PALETTE.dark.ink;
}

/** Publish the shared design tokens on `element`, for the resolved theme. */
export function paintTokens(element: HTMLElement, mode: ResolvedTheme): void {
  const palette = PALETTE[mode];
  element.style.setProperty("--aared-surface", palette.surface);
  element.style.setProperty("--aared-ink", palette.ink);
  element.style.setProperty("--aared-border", palette.border);
  element.style.setProperty("--aared-muted-text", palette.muted);
  element.style.setProperty("--aared-accent", palette.accent);
  element.style.setProperty("--aared-focus-ring", palette.accent);
}
