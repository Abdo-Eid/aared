import { afterEach, describe, expect, test } from "bun:test";
import { PALETTE, automaticInk, paintTokens, resolveTheme } from "../src/appearance.js";
import { contrastRatio } from "../src/core/color.js";

/** The custom properties a renderer publishes, read back from a stand-in element. */
function tokensFor(mode: "light" | "dark"): Record<string, string> {
  const properties: Record<string, string> = {};
  const element = { style: { setProperty: (name: string, value: string) => void (properties[name] = value) } };
  paintTokens(element as unknown as HTMLElement, mode);
  return properties;
}

const original = (globalThis as { matchMedia?: unknown }).matchMedia;
afterEach(() => {
  (globalThis as { matchMedia?: unknown }).matchMedia = original;
});

function systemPrefers(scheme: "light" | "dark"): void {
  (globalThis as { matchMedia?: unknown }).matchMedia = (query: string) => ({
    matches: query === "(prefers-color-scheme: dark)" && scheme === "dark",
  });
}

describe("theme resolution", () => {
  test("light, dark and an unset theme are themselves; unset means light", () => {
    expect(resolveTheme("light")).toBe("light");
    expect(resolveTheme("dark")).toBe("dark");
    expect(resolveTheme(undefined)).toBe("light");
  });

  test("system follows the OS colour scheme, and a host without matchMedia stays light", () => {
    systemPrefers("dark");
    expect(resolveTheme("system")).toBe("dark");
    systemPrefers("light");
    expect(resolveTheme("system")).toBe("light");
    (globalThis as { matchMedia?: unknown }).matchMedia = undefined;
    expect(resolveTheme("system")).toBe("light");
  });

  test("an explicit theme ignores the OS scheme", () => {
    systemPrefers("dark");
    expect(resolveTheme("light")).toBe("light");
    systemPrefers("light");
    expect(resolveTheme("dark")).toBe("dark");
  });
});

describe("design tokens", () => {
  test("every renderer publishes the same six tokens, and the text colour is ink, never the element's own input", () => {
    for (const mode of ["light", "dark"] as const) {
      expect(Object.keys(tokensFor(mode)).sort()).toEqual([
        "--aared-accent",
        "--aared-border",
        "--aared-focus-ring",
        "--aared-ink",
        "--aared-muted-text",
        "--aared-surface",
      ]);
      // `--aared-text` is `<aared-viewer>`'s consumer-facing stage colour; a renderer must not shadow it.
      expect(tokensFor(mode)["--aared-text"]).toBeUndefined();
    }
  });

  test("the palette is pinned, one value per token and theme", () => {
    expect(tokensFor("light")).toEqual({
      "--aared-surface": "#ffffff",
      "--aared-ink": "#000000",
      "--aared-border": "#d1d5db",
      "--aared-muted-text": "#5b6470",
      "--aared-accent": "#0b57d0",
      "--aared-focus-ring": "#0b57d0",
    });
    expect(tokensFor("dark")).toEqual({
      "--aared-surface": "#1a1a1a",
      "--aared-ink": "#e8e8e8",
      "--aared-border": "#3a3a3a",
      "--aared-muted-text": "#a7aab3",
      "--aared-accent": "#8ab4f8",
      "--aared-focus-ring": "#8ab4f8",
    });
  });

  test("dark text keeps AA contrast on the dark surface, and light muted text on white", () => {
    for (const mode of ["light", "dark"] as const) {
      const { ink, muted, accent, surface } = PALETTE[mode];
      expect(contrastRatio(ink, surface)!).toBeGreaterThanOrEqual(7);
      expect(contrastRatio(muted, surface)!).toBeGreaterThanOrEqual(4.5);
      expect(contrastRatio(accent, surface)!).toBeGreaterThanOrEqual(4.5);
    }
  });
});

describe("automatic ink", () => {
  test("a light authored fill takes the dark ink back, a dark one keeps the light ink", () => {
    // The Shot Log's cream zebra rows and its total row, and a white fill.
    for (const fill of ["#FBF9F5", "#F4F1EC", "#FFFFFF", "#F8E8CB"]) expect(automaticInk(fill)).toBe(PALETTE.light.ink);
    // The title band and a mid-dark teal.
    for (const fill of ["#12171F", "#156082", "#1A1A1A"]) expect(automaticInk(fill)).toBe(PALETTE.dark.ink);
  });

  test("it picks the better of the two inks at every grey, and never leaves one under 4:1", () => {
    // Two fixed inks bound the worst case: the crossover grey reads at about 4.1:1 either way.
    for (let level = 0; level <= 255; level++) {
      const hex = level.toString(16).padStart(2, "0");
      const fill = `#${hex}${hex}${hex}`;
      const picked = contrastRatio(fill, automaticInk(fill))!;
      const other = contrastRatio(fill, automaticInk(fill) === PALETTE.light.ink ? PALETTE.dark.ink : PALETTE.light.ink)!;
      expect(picked).toBeGreaterThanOrEqual(other);
      expect(picked).toBeGreaterThanOrEqual(4);
    }
  });

  test("a fill that does not parse keeps the light ink", () => {
    expect(automaticInk("transparent")).toBe(PALETTE.dark.ink);
    expect(automaticInk("rgb(1, 2, 3)")).toBe(PALETTE.dark.ink);
  });
});
