// Lazy singleton loader for CanvasKit (Skia WASM).
//
// CanvasKit is ~7MB and only loaded when the canvas renderer is requested.
// The HTML renderer path has zero extra cost.

type CanvasKit = unknown;

let ck: CanvasKit | null = null;
let loading: Promise<CanvasKit> | null = null;

/**
 * Get the CanvasKit singleton, lazily loading on first call.
 *
 * The WASM binary is resolved relative to the canvaskit-wasm package.
 * Override locateFile if your deployment bundles it differently.
 */
export async function getCanvasKit(): Promise<CanvasKit> {
  if (ck) return ck;
  if (loading) return loading;

  loading = (async () => {
    const mod = await import("canvaskit-wasm");
    const init = mod.default;
    ck = await init({
      locateFile: (file: string) => `/node_modules/canvaskit-wasm/bin/${file}`,
    });
    return ck!;
  })();

  return loading;
}
