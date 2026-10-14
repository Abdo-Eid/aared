/** One cached, retryable effect painter; plain Word documents never load it. */
type EffectPainter = typeof import("./effect-paint.js");
let painter: EffectPainter | undefined;
let loading: Promise<EffectPainter> | undefined;

export function loadedEffectPainter(): EffectPainter | undefined {
  return painter;
}

export function loadEffectPainter(): Promise<EffectPainter> {
  loading ??= import("./effect-paint.js").then(
    module => (painter = module),
    (error: unknown) => { loading = undefined; throw error; },
  );
  return loading;
}
