export { XlView } from "../pkg/offidized_xlview.js";
export type { InitInput, InitOutput } from "../pkg/offidized_xlview.js";

/** Initialize the WASM module. Safe to call multiple times. */
export function init(
  input?: import("../pkg/offidized_xlview.js").InitInput,
): Promise<void>;

export interface MountedViewer {
  /** Load an XLSX file from bytes */
  load(data: Uint8Array): void;
  /** Destroy the viewer and release resources */
  destroy(): void;
  /** Access the underlying XlView instance */
  readonly viewer: import("../pkg/offidized_xlview.js").XlView;
}

/**
 * Mount an offidized-xlview instance into a container element.
 * Creates canvases, sets up resize handling, and returns a controller.
 */
export function mount(container: HTMLElement): Promise<MountedViewer>;

/** `<xl-view>` custom element for drop-in usage */
export declare class XlViewElement extends HTMLElement {
  static readonly observedAttributes: string[];
  /** Load XLSX from bytes */
  load(data: Uint8Array): void;
  /** Access underlying XlView instance (null before mount completes) */
  readonly xlview: import("../pkg/offidized_xlview.js").XlView | null;
}
