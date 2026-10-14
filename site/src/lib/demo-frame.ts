/**
 * Docs plumbing for the playground islands (not displayed as snippets).
 *
 * Each vanilla demo frame owns a stage, a status line, and a picker. This
 * helper wires them to a demo module: picker/file changes start a run,
 * the previous run is torn down first, and stale runs never paint. Frames
 * start lazily — a hidden tab costs no Worker until first visible.
 */

const UNIT_NAMES: Record<string, readonly [string, string]> = {
  pptx: ["slide", "slides"],
  xlsx: ["sheet", "sheets"],
  docx: ["document", "documents"],
};

/** "7 slides", "4 sheets", "1 document": what `info.count` counts, per format. */
export function describeUnits(format: string, count: number): string {
  const [one, many] = UNIT_NAMES[format] ?? ["unit", "units"];
  return `${count} ${count === 1 ? one : many}`;
}

export interface DemoSummary {
  format: string;
  units: number;
}

export interface DemoFixture {
  label: string;
  url: string;
  format: string;
}

export interface DemoOutcome {
  summary: DemoSummary;
  dispose(): void;
}

/**
 * `signal` aborts when a newer run starts in the same stage. A runner that
 * hands it to `preview()` gets the library's own cancel: the superseded viewer
 * is torn down at once, so it can never clear the stage after the newer run
 * has painted (a viewer's `destroy()` empties its target).
 */
export type DemoRunner = (stage: HTMLElement, source: string | File, signal: AbortSignal) => Promise<DemoOutcome>;

export type ResolveRunner = () => DemoRunner;

export function attachDemo(frame: HTMLElement, resolveRunner: ResolveRunner): void {
  const stage = frame.querySelector("[data-stage]");
  const status = frame.querySelector("[data-status]");
  const picker = frame.querySelector<HTMLSelectElement>("[data-pick]");
  const upload = frame.querySelector<HTMLInputElement>("[data-file]");
  if (!(stage instanceof HTMLElement) || !(status instanceof HTMLElement)) return;

  let dispose: (() => void) | null = null;
  let generation = 0;
  let running: AbortController | null = null;
  let current: string | File | null = frame.dataset.initial ?? null;
  let started = false;

  const paint = (source: string | File): void => {
    started = true;
    current = source;
    const run = resolveRunner();
    const gen = ++generation;
    // A lazy start and a picker change can overlap in a tab that was just shown:
    // cancel the earlier run before the new one begins to paint.
    running?.abort();
    const controller = (running = new AbortController());
    dispose?.();
    dispose = null;
    stage.dataset.state = "loading";
    status.dataset.state = "loading";
    delete status.dataset.format;
    delete status.dataset.units;
    status.textContent = "Loading…";
    // A runner that throws before returning its promise must land in the
    // error branch too, not leave the frame on "Loading…" forever.
    new Promise<DemoOutcome>((resolve) => resolve(run(stage, source, controller.signal))).then(
      ({ summary, dispose: done }) => {
        if (gen !== generation) {
          done();
          return;
        }
        dispose = done;
        stage.dataset.state = "ready";
        status.dataset.state = "ready";
        status.dataset.format = summary.format;
        status.dataset.units = String(summary.units);
        status.textContent = `${summary.format} · ${describeUnits(summary.format, summary.units)}`;
      },
      (error: unknown) => {
        if (gen !== generation) return;
        stage.dataset.state = "error";
        status.dataset.state = "error";
        status.textContent = error instanceof Error ? error.message : String(error);
      },
    );
  };

  picker?.addEventListener("change", () => {
    if (upload) upload.value = "";
    if (picker.value !== "") paint(picker.value);
  });
  upload?.addEventListener("change", () => {
    const file = upload.files?.[0];
    if (file) paint(file);
  });
  if (current === null) return;
  const initial = current;
  const start = (): void => {
    if (started) return;
    started = true;
    paint(initial);
  };
  if (frame.offsetParent !== null) {
    start();
  } else {
    const observer = new IntersectionObserver((entries) => {
      if (entries.some((entry) => entry.isIntersecting)) {
        observer.disconnect();
        start();
      }
    });
    observer.observe(frame);
  }
}
