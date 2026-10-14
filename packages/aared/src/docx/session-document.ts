/** Shared initialization for low-level opens and the viewer's already-open session. */
import { DocxDocument } from "./renderer.js";
import type { WorkerSession } from "../core/session.js";
import type { DocumentSection } from "../core/model.js";
import type { OpenOptions } from "../core/index.js";

export async function documentFromSession(session: WorkerSession, options?: OpenOptions): Promise<DocxDocument> {
  try {
    const signal = options?.signal;
    // Font resolution belongs to unit(): both entry paths await it before
    // constructing the document, and neither keeps sections in the Worker.
    const sections = await session.unit<DocumentSection[]>(0, signal === undefined ? undefined : { signal });
    return new DocxDocument(session, sections);
  } catch (error) {
    session.close();
    throw error;
  }
}
