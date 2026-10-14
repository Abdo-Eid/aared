/** Shared renderer presentation defaults; authored geometry stays inline. */
import { cssClass } from "./class-sheet.js";

/**
 * Office text must not inherit the application's emphasis, tracking or casing.
 * Keep family, size, ink and line pitch in their existing format/model homes;
 * in particular, exact-pitch runs deliberately use a zero line height.
 * This blocks inherited host traits, not selectors targeting descendants.
 */
export const CONTENT_TEXT = cssClass("font-weight:normal;font-style:normal;font-variant:normal;font-stretch:normal;letter-spacing:normal;word-spacing:normal;text-transform:none;text-indent:0");

/** A paint layer that follows its frame; clipping remains the frame's choice. */
export const FRAME_PAINT = cssClass("position:absolute;inset:0;width:100%;height:100%");
