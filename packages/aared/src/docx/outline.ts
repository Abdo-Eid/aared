/**
 * The DOCX body heading outline — the one index space `document.headings`,
 * `search()` matches, `goToHeading()` and the painted headings share.
 * Internal: the public surface is `HeadingInfo` and `document.headings` in
 * `./index.ts`.
 *
 * Only headings at the top level of a section's body count: a heading-styled
 * paragraph in a header, footer, table cell or note is not part of the
 * outline (it still paints as a heading, just unindexed). `indexOf` maps each
 * outline node to its index; pagination places blocks whole, so the same node
 * objects reach the painter in flow and layout mode alike.
 */

import type { DocumentSection, DocxHeadingNode } from "../core/model.js";
import type { HeadingInfo } from "./index.js";

export interface HeadingOutline {
  list: readonly HeadingInfo[];
  indexOf: ReadonlyMap<DocxHeadingNode, number>;
}

export function headingOutline(sections: readonly DocumentSection[]): HeadingOutline {
  const list: HeadingInfo[] = [];
  const indexOf = new Map<DocxHeadingNode, number>();
  for (const [section, { root }] of sections.entries()) {
    for (const node of root) {
      if (node.type !== "heading") continue;
      indexOf.set(node, list.length);
      list.push({ index: list.length, title: node.text, level: node.level, section });
    }
  }
  return { list, indexOf };
}
