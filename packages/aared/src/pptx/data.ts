/** Trimmed, nonempty slide paragraphs in element order; excludes master/layout graphics. */
export interface SlideText {
  paragraphs: string[];
  /** Speaker notes as plain text; absent when no notes body exists. */
  notes?: string;
}
