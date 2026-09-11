/**
 * Word `footnotes.xml` / `endnotes.xml` reader for the DOCX route. Entries
 * typed `separator` / `continuationSeparator` / `continuationNotice` are
 * Word's own rules, not content, and never reach the model. Ids are
 * namespaced (`f`/`e`) because both parts number from 1 independently.
 */

import { attr, childElements } from "../xml.js";
import type { XmlElement } from "../xml.js";
import { throwIfAborted } from "../zip.js";

const SEPARATOR_TYPES = new Set(["separator", "continuationSeparator", "continuationNotice"]);

export function parseNotes(
  root: XmlElement | undefined,
  prefix: "f" | "e",
  local: "footnote" | "endnote",
  signal?: AbortSignal,
): Map<string, XmlElement> {
  throwIfAborted(signal);
  const out = new Map<string, XmlElement>();
  if (root === undefined) return out;
  for (const note of childElements(root, local)) {
    throwIfAborted(signal);
    const id = attr(note, "id");
    if (id === undefined) continue;
    const type = attr(note, "type");
    if (type !== undefined && SEPARATOR_TYPES.has(type)) continue;
    out.set(`${prefix}${id}`, note);
  }
  throwIfAborted(signal);
  return out;
}
