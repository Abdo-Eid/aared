import { open, render } from "aared/pptx";
import { defaultResolver } from "aared-fonts";
import type { BoundsModel, ShapeModel, SlideElementModel, SlideModel } from "aared/core";
import type { DemoOutcome } from "../lib/demo-frame.js";

/**
 * The light table: one slide shown as two mounted transparencies inside
 * `stage` — `[data-panel="json"]` gets the parsed IR as real, hand-tokenized
 * JSON; `[data-panel="render"]` gets the live DOM/SVG `aared` actually
 * paints. Hovering (or focusing) a shape's row lights up that exact shape
 * in the render panel, and back — matched by the `id` the renderer already
 * stamps as `data-aared-shape-id` on every `kind: "shape"` wrapper
 * (packages/aared/src/pptx/index.ts). Tables/pictures/charts aren't
 * id-stamped by the renderer, so their JSON rows render read-only
 * (`data-static`) rather than pretending to link.
 */
export async function runLightTable(stage: HTMLElement, source: string | File): Promise<DemoOutcome> {
  const jsonPanel = stage.querySelector<HTMLElement>('[data-panel="json"]');
  const renderPanel = stage.querySelector<HTMLElement>('[data-panel="render"]');
  if (!jsonPanel || !renderPanel) {
    throw new Error("light table: stage is missing its [data-panel] mounts");
  }

  const deck = await open(source, { fontResolver: defaultResolver });
  try {
    const slide = deck.slides.get(0);
    const model = await slide.read();
    const viewer = await render({ slide, target: renderPanel });

    const rows = new Map<string, HTMLElement>();
    jsonPanel.replaceChildren(buildSlideTree(model, rows));

    const shapes = new Map<string, HTMLElement>();
    for (const id of rows.keys()) {
      const shape = renderPanel.querySelector<HTMLElement>(`[data-aared-shape-id="${CSS.escape(id)}"]`);
      if (shape) shapes.set(id, shape);
    }

    const setLit = (id: string, on: boolean): void => {
      rows.get(id)?.toggleAttribute("data-lit", on);
      shapes.get(id)?.toggleAttribute("data-lit", on);
    };

    const teardown: Array<() => void> = [];
    const wire = (el: HTMLElement, id: string): void => {
      const on = (): void => setLit(id, true);
      const off = (): void => setLit(id, false);
      el.addEventListener("mouseenter", on);
      el.addEventListener("mouseleave", off);
      el.addEventListener("focus", on);
      el.addEventListener("blur", off);
      teardown.push(() => {
        el.removeEventListener("mouseenter", on);
        el.removeEventListener("mouseleave", off);
        el.removeEventListener("focus", on);
        el.removeEventListener("blur", off);
      });
    };
    for (const [id, row] of rows) {
      wire(row, id);
      const shape = shapes.get(id);
      if (shape) wire(shape, id);
    }

    return {
      summary: { format: deck.info.format, units: deck.info.count },
      dispose: () => {
        for (const off of teardown) off();
        viewer.destroy();
        deck.close();
        jsonPanel.replaceChildren();
      },
    };
  } catch (error) {
    deck.close();
    throw error;
  }
}

// ---- JSON panel: a small hand-rolled projector + tokenizer -------------
//
// Not Shiki: the hover-link needs one live DOM node per element as the
// tree is built, and a full JSON.stringify of a real slide is multi-KB of
// SVG path data nobody wants to read. This projects only what a viewer
// would scan for — kind, id, name, bounds, a short text preview — and
// elides geometry to a byte count.

function buildSlideTree(model: SlideModel, rows: Map<string, HTMLElement>): DocumentFragment {
  const frag = document.createDocumentFragment();
  frag.append(
    line(0, punct("{")),
    kvLine(1, "part", str(model.part)),
    line(1, key("size"), punct(": "), punct("{ "), kv("width", num(model.size.width)), punct(", "), kv("height", num(model.size.height)), punct(" },")),
    line(1, key("elements"), punct(": [")),
  );
  model.elements.forEach((element, index) => {
    const isLast = index === model.elements.length - 1;
    frag.append(buildElementRow(element, 2, isLast, rows));
  });
  frag.append(line(1, punct("]")), line(0, punct("}")));
  return frag;
}

function buildElementRow(
  element: SlideElementModel,
  indent: number,
  isLast: boolean,
  rows: Map<string, HTMLElement>,
): HTMLElement {
  const isShape = element.kind === "shape";
  const row = document.createElement("div");
  row.className = "lt-json-row";
  row.dataset.aaredElement = element.kind;
  if (isShape) {
    row.dataset.aaredShapeId = element.id;
    row.tabIndex = 0;
    row.setAttribute("role", "button");
    row.setAttribute("aria-label", `Highlight shape ${element.name || element.id} on the light table`);
    rows.set(element.id, row);
  } else {
    row.dataset.static = "";
  }

  row.append(line(indent, punct("{")));
  row.append(kvLine(indent + 1, "kind", str(element.kind)));
  row.append(kvLine(indent + 1, "id", str(element.id)));
  row.append(kvLine(indent + 1, "name", str(element.name)));
  if (element.bounds) row.append(boundsLine(indent + 1, element.bounds));

  if (element.kind === "shape") {
    appendShapeFields(row, element, indent + 1);
  } else if (element.kind === "table") {
    row.append(kvLine(indent + 1, "direction", str(element.direction)));
    row.append(kvLine(indent + 1, "rows", num(element.rows.length)));
    row.append(kvLine(indent + 1, "columns", num(element.columns.length)));
  } else if (element.kind === "picture") {
    row.append(kvLine(indent + 1, "mediaId", str(element.mediaId)));
    row.append(kvLine(indent + 1, "contentType", str(element.contentType)));
  } else if (element.kind === "chart") {
    row.append(kvLine(indent + 1, "graphic", element.graphic ? str("resolved") : keyword("undefined")));
  }

  row.append(line(indent, punct(isLast ? "}" : "},")));
  return row;
}

function appendShapeFields(row: HTMLElement, shape: ShapeModel, indent: number): void {
  const preview = textPreview(shape);
  if (preview !== undefined) row.append(kvLine(indent, "text", str(preview)));
  if (shape.geometry) {
    row.append(
      line(
        indent,
        key("geometry"),
        punct(": { "),
        key("path"),
        punct(": "),
        elided(`« ${shape.geometry.path.length} chars »`),
        punct(" },"),
      ),
    );
  }
}

function textPreview(shape: ShapeModel): string | undefined {
  if (!shape.text) return undefined;
  const joined = shape.text.paragraphs
    .map((p) => p.runs.map((r) => r.text).join(""))
    .join(" ")
    .trim();
  if (!joined) return undefined;
  return joined.length > 64 ? `${joined.slice(0, 64)}…` : joined;
}

function boundsLine(indent: number, bounds: BoundsModel): HTMLElement {
  return line(
    indent,
    key("bounds"),
    punct(": { "),
    kv("x", num(round(bounds.x))),
    punct(", "),
    kv("y", num(round(bounds.y))),
    punct(", "),
    kv("width", num(round(bounds.width))),
    punct(", "),
    kv("height", num(round(bounds.height))),
    punct(" },"),
  );
}

function round(n: number): number {
  return Math.round(n * 100) / 100;
}

// ---- tiny token-span helpers --------------------------------------------

function line(indent: number, ...children: Array<Node | string>): HTMLElement {
  const el = document.createElement("div");
  el.className = "lt-json-line";
  el.style.paddingInlineStart = `${indent * 1.1}em`;
  el.append(...children);
  return el;
}

function kvLine(indent: number, k: string, valueNode: Node): HTMLElement {
  return line(indent, kv(k, valueNode), punct(","));
}

function kv(k: string, valueNode: Node): DocumentFragment {
  const frag = document.createDocumentFragment();
  frag.append(key(k), punct(": "), valueNode);
  return frag;
}

function key(k: string): HTMLElement {
  const span = document.createElement("span");
  span.className = "tok-key";
  span.textContent = `"${k}"`;
  return span;
}

function str(s: string): HTMLElement {
  const span = document.createElement("span");
  span.className = "tok-str";
  span.textContent = JSON.stringify(s);
  return span;
}

function num(n: number): HTMLElement {
  const span = document.createElement("span");
  span.className = "tok-num";
  span.textContent = String(n);
  return span;
}

function keyword(s: string): HTMLElement {
  const span = document.createElement("span");
  span.className = "tok-kw";
  span.textContent = s;
  return span;
}

function elided(s: string): HTMLElement {
  const span = document.createElement("span");
  span.className = "tok-elided";
  span.textContent = s;
  return span;
}

function punct(s: string): HTMLElement {
  const span = document.createElement("span");
  span.className = "tok-punct";
  span.textContent = s;
  return span;
}
