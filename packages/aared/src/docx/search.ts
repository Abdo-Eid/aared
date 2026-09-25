/**
 * DOCX `search()` index: every searchable block of the cached sections,
 * flattened once in document order with its lowercased haystack, so repeated
 * searches are plain string scans. Internal: the public surface is
 * `DocxDocument.search()` and `DocxSearchMatch` in `./index.ts`.
 *
 * Covered: headings, paragraphs, list items, table-cell blocks (nested
 * tables included), and footnote/endnote bodies. Not covered: headers and
 * footers, text boxes (not parsed), and image alt text.
 */

import type { DocumentSection, DocxFootnoteNode, FlowNode, TextRunModel } from "../core/model.js";
import { clipSnippet, foldText, runsText } from "../core/search.js";
import type { DocxSearchMatch } from "./index.js";
import type { HeadingOutline } from "./outline.js";

/** One searchable block, plus the folded text `search()` scans. */
export interface DocxSearchEntry {
  match: DocxSearchMatch;
  lower: string;
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
  /** Set inside a table cell; the innermost table wins. */
  table?: { index: number; row: number; col: number };
  /** Set inside a footnote/endnote body. */
  note?: { id: string; label: string };
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
      lower: foldText(text),
      match: {
        kind: noteKind ?? (walk.table === undefined ? kind : "table"),
        section: walk.location.section,
        heading: walk.location.heading,
        text: clipSnippet(text),
        ...(walk.table === undefined ? {} : { table: { ...walk.table } }),
        ...(walk.note === undefined ? {} : { note: { ...walk.note } }),
      },
    });
  };

  const noteReferences = (runs: readonly TextRunModel[], walk: Walk): void => {
    for (const run of runs) {
      if (run.footnoteRef === undefined || references.has(run.footnoteRef)) continue;
      references.set(run.footnoteRef, { ...walk.location });
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
          break;
        }
        case "paragraph":
          emit(paragraphText(node.runs), walk, "paragraph");
          noteReferences(node.runs, walk);
          break;
        case "list":
          for (const item of node.items) {
            emit(paragraphText(item.paragraph.runs), walk, "list");
            noteReferences(item.paragraph.runs, walk);
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
    visit(root, { location: { section, heading: null }, tables: { count: 0 } });
  }
  for (const note of notes) {
    const origin = references.get(note.id) ?? { section: Math.max(0, sections.length - 1), heading: null };
    visit(note.nodes, {
      location: { ...origin },
      tables: { count: 0 },
      note: { id: note.id, label: note.label },
    });
  }
  return entries;
}
