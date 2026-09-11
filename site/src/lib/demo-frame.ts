/**
 * Docs plumbing for the playground islands (not displayed as snippets).
 *
 * Each vanilla demo frame owns a stage, a status line, and a picker. This
 * helper wires them to a demo module: picker/file changes start a run,
 * the previous run is torn down first, and stale runs never paint. Frames
 * start lazily — a hidden tab costs no Worker until first visible.
 */

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

export type DemoRunner = (stage: HTMLElement, source: string | File) => Promise<DemoOutcome>;

export type ResolveRunner = () => DemoRunner;

export function attachDemo(frame: HTMLElement, resolveRunner: ResolveRunner): void {
  const stage = frame.querySelector("[data-stage]");
  const status = frame.querySelector("[data-status]");
  const picker = frame.querySelector<HTMLSelectElement>("[data-pick]");
  const upload = frame.querySelector<HTMLInputElement>("[data-file]");
  if (!(stage instanceof HTMLElement) || !(status instanceof HTMLElement)) return;

  let dispose: (() => void) | null = null;
  let generation = 0;
  let current: string | File | null = frame.dataset.initial ?? null;
  let started = false;

  const paint = (source: string | File): void => {
    current = source;
    const run = resolveRunner();
    const gen = ++generation;
    dispose?.();
    dispose = null;
    stage.dataset.state = "loading";
    status.dataset.state = "loading";
    delete status.dataset.format;
    delete status.dataset.units;
    status.textContent = "Loading…";
    run(stage, source).then(
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
        status.textContent = `${summary.format} · ${summary.units} unit${summary.units === 1 ? "" : "s"}`;
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
  for (const pill of frame.querySelectorAll<HTMLElement>("[data-format-pill]")) {
    pill.addEventListener("click", () => {
      frame.dataset.format = pill.dataset.formatPill ?? "";
      for (const other of frame.querySelectorAll("[data-format-pill]")) {
        other.toggleAttribute("data-active", other === pill);
      }
      if (current !== null) paint(current);
    });
  }

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
