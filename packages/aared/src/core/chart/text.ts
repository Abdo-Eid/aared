/**
 * The words a chart carries, read straight from the resolved spec with no
 * layout: what search matches and what a screen reader is told. DOM-free, so
 * the same functions serve the page (a chart's accessible name), the Worker
 * (search's text reads) and plain unit tests.
 *
 * The spec keeps no rich text for individual data labels (a label states
 * which of value, category, series and percent it shows, never its own words),
 * so custom data-label text is not searchable.
 */

import type { ChartSpecModel, Group } from "./spec.js";

/** How many series names or categories a description lists before it says "and N more". */
const LISTED = 8;

/** The chart title as one line, or `undefined` when it has none. */
export function chartTitleText(spec: ChartSpecModel): string | undefined {
  const text = spec.title?.lines.join(" ").trim();
  return text === undefined || text === "" ? undefined : text;
}

/**
 * Every distinct piece of text a chart shows, in reading order: the title,
 * the titles of its visible axes (and a display-unit label), the series names,
 * then the category labels. One string per item, never merged, so a match
 * reports the item that holds it.
 */
export function chartText(spec: ChartSpecModel): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  const add = (text: string | undefined): void => {
    const trimmed = text?.trim();
    if (trimmed === undefined || trimmed === "" || seen.has(trimmed)) return;
    seen.add(trimmed);
    out.push(trimmed);
  };
  add(chartTitleText(spec));
  for (const axis of spec.axes) {
    if (axis.deleted) continue;
    add(axis.title?.lines.join(" "));
    add(axis.displayUnitLabel);
  }
  for (const group of spec.groups) for (const series of group.series) add(series.name);
  for (const group of spec.groups) for (const series of group.series) for (const category of series.categories) add(category);
  return out;
}

function kindName(group: Group): string {
  switch (group.kind) {
    case "bar":
      return group.direction === "bar" ? "bar" : "column";
    default:
      return group.kind;
  }
}

function groupName(group: Group): string {
  const grouping = group.kind === "bar" || group.kind === "line" || group.kind === "area"
    ? group.grouping === "stacked" ? "stacked " : group.grouping === "percentStacked" ? "100% stacked " : ""
    : "";
  return `${grouping}${kindName(group)}`;
}

function list(items: readonly string[]): string {
  const shown = items.slice(0, LISTED).join(", ");
  return items.length > LISTED ? `${shown} and ${items.length - LISTED} more` : shown;
}

/**
 * A one-sentence-per-fact summary for assistive technology: the plot type,
 * the series and the categories, e.g. `Stacked column chart. 2 series: Logged,
 * Planned. 3 categories: Ep 04, Ep 05, Ep 06.` Chart words (names, labels)
 * come from the file, so they keep their own language; the template is plain
 * ASCII.
 */
export function chartDescription(spec: ChartSpecModel): string {
  const names = [...new Set(spec.groups.map(groupName))];
  const kind = names.length === 0 ? "chart" : names.length === 1 ? `${names[0]} chart` : `combination chart (${names.join(", ")})`;
  const sentences = [`${kind.charAt(0).toUpperCase()}${kind.slice(1)}.`];
  const series = spec.groups.flatMap((group) => group.series.map((s) => s.name));
  if (series.length > 0) sentences.push(`${series.length} series: ${list(series)}.`);
  const categories = spec.groups.flatMap((group) => group.series[0]?.categories ?? []).filter((c) => c.trim() !== "");
  const distinct = [...new Set(categories)];
  if (distinct.length > 0) sentences.push(`${distinct.length} ${distinct.length === 1 ? "category" : "categories"}: ${list(distinct)}.`);
  return sentences.join(" ");
}
