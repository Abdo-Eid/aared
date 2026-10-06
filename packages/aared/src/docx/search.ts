/**
 * DOCX `search()` index: every searchable block of the cached sections,
 * flattened once in document order with its search key (`core/search.ts`), so
 * repeated searches are plain string scans. Internal: the public surface is
 * `DocxDocument.search()` and `DocxSearchMatch` in `./index.ts`.
 *
 * Covered: headings, paragraphs, list items, table-cell blocks (nested
 * tables included), footnote/endnote bodies, and chart text (title, axis
 * titles, series names, categories; one match per item, for charts in the flow
 * and charts anchored in a paragraph alike), plus each distinct header/footer
 * part once (including its tables, text boxes and charts). Not covered: image
 * alt text and custom data-label text
 * (the chart model keeps none).
 */

import { chartText } from "../core/chart/text.js";
import type { DocumentSection, DocxChartNode, DocxDrawingNode, DocxFootnoteNode, DocxTextBoxNode, FlowNode, TextRunModel } from "../core/model.js";
import { clipSnippet, runsText, searchKey } from "../core/search.js";
import { docxStories } from "./stories.js";
import type { DocxSearchMatch } from "./index.js";
import type { HeadingOutline } from "./outline.js";

/** One searchable block, plus the search key `search()` scans. */
export interface DocxSearchEntry {
  match: DocxSearchMatch;
  key: string;
}

/** Where a block sits: its section and the nearest preceding outline heading. */
interface Location {
  section: number;
  heading: { index: number; title: string } | null;
}

interface Walk {
  /** Shared and mutable: a body heading advances it for the rest of the section. */
  location: Location;
  /** Per-section counter giving each table its document-order index. */
  tables: { count: number };
  /** Per-section counter giving each text box its document-order index. */
  textBoxes: { count: number };
  /** Set inside a table cell; the innermost table wins. */
  table?: { index: number; row: number; col: number };
  /** Set inside a footnote/endnote body. */
  note?: { id: string; label: string };
  textBox?: { index: number };
  story?: { kind: "header" | "footer"; part: string; variant: "default" | "first" | "even" };
}

function paragraphText(runs: readonly TextRunModel[]): string {
  return runsText(runs).trim();
}

/** A fresh public match for an entry, so callers can't mutate the cached index. */
export function entryMatch(entry: DocxSearchEntry): DocxSearchMatch {
  const { match } = entry;
  return {
    ...match,
    heading: match.heading === null ? null : { ...match.heading },
    ...(match.table === undefined ? {} : { table: { ...match.table } }),
    ...(match.note === undefined ? {} : { note: { ...match.note } }),
    ...(match.textBox === undefined ? {} : { textBox: { ...match.textBox } }),
  };
}

export function buildDocxSearchIndex(sections: readonly DocumentSection[], outline: HeadingOutline): DocxSearchEntry[] {
  const entries: DocxSearchEntry[] = [];
  // Footnotes/endnotes live at the end of the LAST section's root, so a note
  // reports the section and heading of the run that first referenced it.
  const references = new Map<string, Location>();
  const notes: DocxFootnoteNode[] = [];

  const emit = (text: string, walk: Walk, kind: DocxSearchMatch["kind"]): void => {
    if (text === "") return;
    const noteKind = walk.note === undefined ? undefined : walk.note.id.startsWith("e") ? "endnote" : "footnote";
    entries.push({
      key: searchKey(text),
      match: {
        // A header/footer is the searchable story as a whole, regardless of
        // whether this block lives in a table, text box, or chart in that part.
        kind: walk.story?.kind ?? noteKind ?? (walk.table !== undefined ? "table" : walk.textBox !== undefined ? "textBox" : kind),
        section: walk.location.section,
        heading: walk.story === undefined ? walk.location.heading : null,
        text: clipSnippet(text),
        ...(walk.story !== undefined ? { story: { part: walk.story.part, variant: walk.story.variant } } : {}),
        ...(walk.story === undefined && walk.table !== undefined ? { table: { ...walk.table } } : {}),
        ...(walk.note === undefined ? {} : { note: { ...walk.note } }),
        ...(walk.story === undefined && walk.textBox !== undefined ? { textBox: { ...walk.textBox } } : {}),
      },
    });
  };

  const noteReferences = (runs: readonly TextRunModel[], walk: Walk): void => {
    for (const run of runs) {
      if (run.footnoteRef === undefined || references.has(run.footnoteRef)) continue;
      references.set(run.footnoteRef, { ...walk.location });
    }
  };

  /** A chart's words, one entry per item (title, axis titles, series, categories), after the block it sits in. */
  const emitChart = (chart: DocxChartNode, walk: Walk): void => {
    for (const text of chartText(chart.spec)) emit(text, walk, "chart");
  };
  const emitTextBox = (textBox: DocxTextBoxNode, walk: Walk): void => {
    const index = walk.textBoxes.count++;
    visit(textBox.content, { ...walk, textBox: { index } });
  };
  /** Charts anchored in a paragraph or heading (`floats`) belong to it, so they are searched where it is. */
  const emitFloats = (floats: readonly DocxDrawingNode[] | undefined, walk: Walk): void => {
    for (const float of floats ?? []) {
      if (float.type === "chart") emitChart(float, walk);
      else if (float.type === "textBox") emitTextBox(float, walk);
    }
  };

  const visit = (nodes: readonly FlowNode[], walk: Walk): void => {
    for (const node of nodes) {
      switch (node.type) {
        case "heading": {
          emit(node.text, walk, "heading");
          // Only outline headings advance the context; a heading inside a
          // table cell or note body isn't in the outline.
          const index = outline.indexOf.get(node);
          if (index !== undefined) walk.location.heading = { index, title: node.text };
          noteReferences(node.runs, walk);
          emitFloats(node.floats, walk);
          break;
        }
        case "paragraph":
          emit(paragraphText(node.runs), walk, "paragraph");
          noteReferences(node.runs, walk);
          emitFloats(node.floats, walk);
          break;
        case "chart":
          emitChart(node, walk);
          break;
        case "textBox":
          emitTextBox(node, walk);
          break;
        case "list":
          for (const item of node.items) {
            emit(paragraphText(item.paragraph.runs), walk, "list");
            noteReferences(item.paragraph.runs, walk);
            emitFloats(item.paragraph.floats, walk);
          }
          break;
        case "table": {
          const index = walk.note === undefined ? walk.tables.count++ : -1;
          for (const [row, rowNode] of node.rows.entries()) {
            for (const [col, cell] of rowNode.cells.entries()) {
              visit(cell.nodes, {
                ...walk,
                ...(walk.note === undefined ? { table: { index, row, col } } : {}),
              });
            }
          }
          break;
        }
        case "footnote":
          if (walk.note === undefined) notes.push(node);
          break;
        default:
          break;
      }
    }
  };

  for (const [section, { root }] of sections.entries()) {
    visit(root, { location: { section, heading: null }, tables: { count: 0 }, textBoxes: { count: 0 } });
  }
  const indexedStories = new Set<string>();
  for (const story of docxStories(sections)) {
    const identity = story.part;
    if (indexedStories.has(identity)) continue;
    indexedStories.add(identity);
    visit(story.nodes, {
      location: { section: story.section, heading: null },
      tables: { count: 0 },
      textBoxes: { count: 0 },
      story: { kind: story.kind, part: story.part, variant: story.variant },
    });
  }
  for (const note of notes) {
    const origin = references.get(note.id) ?? { section: Math.max(0, sections.length - 1), heading: null };
    visit(note.nodes, {
      location: { ...origin },
      tables: { count: 0 },
      textBoxes: { count: 0 },
      note: { id: note.id, label: note.label },
    });
  }
  return entries;
}
