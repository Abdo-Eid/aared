import type { DocumentSection, FlowNode, HeaderFooterParts } from "../core/model.js";

export interface DocxStory {
  section: number;
  kind: "header" | "footer";
  variant: "default" | "first" | "even";
  part: string;
  nodes: readonly FlowNode[];
}

/** Resolve part inheritance without changing the read-only Worker model. */
export function docxStories(sections: readonly DocumentSection[]): DocxStory[] {
  const result: DocxStory[] = [];
  const headers: HeaderFooterParts = {};
  const footers: HeaderFooterParts = {};
  const anonymous = new Map<readonly FlowNode[], string>();
  for (const [section, model] of sections.entries()) {
    for (const kind of ["header", "footer"] as const) {
      const inherited = kind === "header" ? headers : footers;
      const own = kind === "header" ? model.properties.headerParts : model.properties.footerParts;
      for (const variant of ["default", "first", "even"] as const) {
        if (own?.[variant] !== undefined) inherited[variant] = own[variant];
        const nodes = kind === "header"
          ? variant === "default" ? model.header : variant === "first" ? model.firstHeader : model.evenHeader
          : variant === "default" ? model.footer : variant === "first" ? model.firstFooter : model.evenFooter;
        if (nodes === undefined || nodes.length === 0) continue;
        let part = inherited[variant];
        if (part === undefined) {
          part = anonymous.get(nodes);
          if (part === undefined) {
            part = `resolved-story-${anonymous.size}`;
            anonymous.set(nodes, part);
          }
        }
        result.push({ section, kind, variant, part, nodes });
      }
    }
  }
  return result;
}
