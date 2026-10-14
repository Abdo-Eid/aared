/** One cached, retryable lazy fill painter for charts and Word shapes. */
export type FillPainter = typeof import("./fill-paint.js");
let painter: FillPainter | undefined;
let loading: Promise<FillPainter> | undefined;

export function loadedFillPainter(): FillPainter | undefined {
  return painter;
}

export function loadFillPainter(): Promise<FillPainter> {
  loading ??= import("./fill-paint.js").then(
    (module) => (painter = module),
    (error: unknown) => {
      loading = undefined;
      throw error;
    },
  );
  return loading;
}
