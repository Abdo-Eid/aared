/**
 * One image reference for every `a:blip` in every format (a PPTX/XLSX picture,
 * a DOCX drawing, a picture bullet, a shared picture fill). Office stores an
 * inserted SVG beside its raster fallback: the main `a:blip/@r:embed` names
 * the PNG Office rasterized for compatibility, while the SVG lives behind the
 * `a:extLst/a:ext/asvg:svgBlip/@r:embed` extension (MS-ODRAWXML 2.26.1.1,
 * "Pictures"). The viewer paints the SVG and keeps the raster as the
 * fallback, so an SVG that cannot be read or decoded still shows Office's own
 * rendering of it. Headless: no DOM, no format parser imports.
 */

import { attr, childElements, firstChild } from "./xml.js";
import type { XmlElement } from "./xml.js";

export interface ImageSourceRelationship {
  id: string;
  type: string;
  target: string;
  targetMode: string;
}

/** What a picture's owning part supplies: its relationships, MIME lookup, and part existence. */
export interface ImageSourceOwner {
  relationships: readonly ImageSourceRelationship[];
  mediaContentType: (target: string) => string;
  /**
   * Whether the package carries the part. Absent where the caller holds no
   * package (DOCX part contexts): the SVG is then trusted on its
   * relationship alone, and the renderer's fallback covers a miss.
   */
  hasPart?: (target: string) => boolean;
}

/**
 * A picture's preferred source plus Office's own raster fallback. `mediaId`
 * is the SVG part when a valid embedded `svgBlip` exists, else the main
 * `blip`; the fallback is the main `blip` when it differs from the source.
 * Identifiers are OPC part paths, accepted by `MediaCollection.get()`.
 */
export interface ResolvedImageSource {
  mediaId: string;
  contentType: string;
  fallbackMediaId?: string;
  fallbackContentType?: string;
}

/** An internal image relationship's target part, `undefined` for external or untyped relationships. */
function imageTarget(owner: ImageSourceOwner, id: string | undefined): string | undefined {
  if (id === undefined) return undefined;
  const rel = owner.relationships.find((candidate) => candidate.id === id);
  if (rel === undefined || rel.targetMode === "External" || !rel.type.endsWith("/image")) return undefined;
  return rel.target;
}

/**
 * The `svgBlip` embed under `blip`'s extension list, if any. `attr()` matches
 * the local name, so the `asvg:` prefix (and any other prefix a file uses)
 * resolves. A `link` (an external SVG, not embedded) is never followed:
 * linked media is not previewed, as for every other external part.
 */
function svgBlipEmbed(blip: XmlElement): string | undefined {
  const list = firstChild(blip, "extLst");
  if (list === undefined) return undefined;
  for (const ext of childElements(list, "ext")) {
    const svg = firstChild(ext, "svgBlip");
    if (svg === undefined) continue;
    const embed = attr(svg, "embed");
    if (embed !== undefined) return embed;
  }
  return undefined;
}

/**
 * Resolve one `a:blip` element to its preferred source. An embedded SVG wins;
 * anything else (no extension, an external or missing SVG, a non-SVG target)
 * keeps the main `blip` exactly as before, so a file without SVGs resolves
 * byte-identically to how it always has. `undefined` when neither reference
 * resolves, so one bad reference never breaks its unit.
 */
export function resolveImageSource(blip: XmlElement | undefined, owner: ImageSourceOwner): ResolvedImageSource | undefined {
  if (blip === undefined) return undefined;
  const main = imageTarget(owner, attr(blip, "embed"));
  const svg = imageTarget(owner, svgBlipEmbed(blip));
  // The SVG candidate must exist where the caller can check: a dangling SVG
  // reference keeps the main `blip` (whose own missing-part behaviour — a
  // marked fallback box, or the caller's damaged report — stays exactly as
  // it was). The main reference itself is never existence-checked here.
  if (svg !== undefined && owner.mediaContentType(svg) === "image/svg+xml"
    && (owner.hasPart === undefined || owner.hasPart(svg))) {
    const out: ResolvedImageSource = { mediaId: svg, contentType: owner.mediaContentType(svg) };
    if (main !== undefined && main !== svg) {
      out.fallbackMediaId = main;
      out.fallbackContentType = owner.mediaContentType(main);
    }
    return out;
  }
  if (main === undefined) return undefined;
  return { mediaId: main, contentType: owner.mediaContentType(main) };
}
