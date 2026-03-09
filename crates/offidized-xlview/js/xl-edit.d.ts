export { XlEdit } from "../pkg/offidized_xlview.js";
export type { InitInput, InitOutput } from "../pkg/offidized_xlview.js";

/** Initialize the WASM module. Safe to call multiple times. */
export function init(
  input?: import("../pkg/offidized_xlview.js").InitInput,
): Promise<void>;

export interface MountedEditor {
  /** Load an XLSX file from bytes */
  load(data: Uint8Array): void;
  /** Save modified XLSX -- returns the bytes */
  save(): Uint8Array;
  /** Download modified XLSX as a file */
  download(filename?: string): void;
  /** Destroy the editor and release resources */
  destroy(): void;
  /** Access the underlying XlEdit instance */
  readonly editor: import("../pkg/offidized_xlview.js").XlEdit;
  /** Optional save callback invoked on Ctrl+S / Cmd+S */
  onSave?: () => void;
  /** Optional callback invoked when dirty state changes */
  onDirtyChange?: (dirty: boolean) => void;
}

/**
 * Mount an xl-edit instance into a container element.
 * Creates canvases, sets up resize handling + editing, and returns a controller.
 *
 * Keyboard shortcuts:
 * - Double-click a cell to edit
 * - Enter commits and moves down (Shift+Enter moves up)
 * - Tab commits and moves right (Shift+Tab moves left)
 * - Escape cancels editing
 * - Typing a printable character on a selected cell starts editing (replaces content)
 * - Delete/Backspace clears the selected cell
 * - Ctrl+Z / Cmd+Z to undo, Ctrl+Shift+Z / Cmd+Shift+Z / Ctrl+Y / Cmd+Y to redo
 * - Ctrl+S / Cmd+S triggers onSave callback and dispatches "xlview-save" event
 *
 * Custom events dispatched on the container:
 * - "xlview-save" -- when Ctrl+S / Cmd+S is pressed
 * - "xlview-dirty" -- after any mutating operation; detail: { dirty: boolean }
 */
export function mountEditor(container: HTMLElement): Promise<MountedEditor>;
