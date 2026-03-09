// Type declarations for the doc-view public API.

export { DocView } from "../pkg/offidized_docview.js";
export type { InitInput, InitOutput } from "../pkg/offidized_docview.js";

/** Initialize the WASM module. Safe to call multiple times. */
export function init(
  input?: import("../pkg/offidized_docview.js").InitInput,
): Promise<void>;

export type RenderMode = "continuous" | "paginated";

export interface MountedViewer {
  /** Load a .docx file from bytes. Returns parse time in ms. */
  load(data: Uint8Array): number;
  /** Switch between continuous and paginated modes. */
  setMode(mode: RenderMode): void;
  /** Get the current render mode. */
  getMode(): RenderMode;
  /** Destroy the viewer and clean up resources. */
  destroy(): void;
}

/**
 * Mount a docview instance into a container element.
 * Creates the renderer and returns a controller.
 *
 * @param opts.renderer - `"canvas"` (default) uses CanvasKit rendering;
 *   `"html"` uses DOM-based rendering.
 */
export function mount(
  container: HTMLElement,
  opts?: { renderer?: "html" | "canvas" },
): Promise<MountedViewer>;

/** `<doc-view>` custom element for drop-in usage. */
export declare class DocViewElement extends HTMLElement {
  static readonly observedAttributes: string[];
  /** Load .docx from bytes. */
  load(data: Uint8Array): number;
  /** Set render mode. */
  setMode(mode: RenderMode): void;
}
