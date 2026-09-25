/**
 * Shared `search()` and lookup helpers for the three format entrypoints.
 * Internal: not part of the public surface. `search()` matching stays a
 * case-insensitive substring test; each format owns what text it searches and
 * how a hit is shaped. Lookups by label (`findByTitle`) are exact-then-folded
 * equality, never substring.
 */

/** Longest snippet a search match reports. */
const SNIPPET_LENGTH = 120;

/**
 * The one case fold every search compares through: needles and haystacks both
 * pass it, so changing the rule here changes both sides together.
 */
export function foldText(text: string): string {
  return text.toLowerCase();
}

/**
 * The one rule every lookup by label follows (`sheets.get("Q1")`,
 * `headings.get("Scope")`, `goToHeading("Scope")`): the item whose title is
 * exactly `title`, else the first whose title matches ignoring case. Never a
 * substring — partial matching belongs to `search()`. The caller raises its
 * own `not-found` when nothing matches.
 */
export function findByTitle<T>(items: readonly T[], title: string, titleOf: (item: T) => string): T | undefined {
  const exact = items.find((item) => titleOf(item) === title);
  if (exact !== undefined) return exact;
  const folded = foldText(title);
  return items.find((item) => foldText(titleOf(item)) === folded);
}

/** The folded needle for `query`, or `null` when it matches nothing. */
export function foldQuery(query: string): string | null {
  const needle = foldText(query);
  return needle === "" ? null : needle;
}

/** A block's plain text: its runs joined in order. */
export function runsText(runs: readonly { text: string }[]): string {
  return runs.map((run) => run.text).join("");
}

/** Trim a matching block to the reported snippet length. */
export function clipSnippet(text: string): string {
  return text.slice(0, SNIPPET_LENGTH);
}
