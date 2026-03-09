import _init, { XlEdit } from "../pkg/offidized_xlview.js";
import type { InitInput } from "../pkg/offidized_xlview.js";

export { XlEdit };

// Reuse init from xl-view (safe to call multiple times)
let _initPromise: Promise<void> | null = null;
export function init(input?: InitInput): Promise<void> {
  if (!_initPromise) _initPromise = _init(input).then(() => {});
  return _initPromise;
}

// Helper: read viewport dimensions from scroll container (or fallback)
function viewportSize(container: HTMLElement) {
  const el =
    (container.querySelector("[data-xlview-scroll]") as HTMLElement) ??
    container;
  const w = el.clientWidth || el.getBoundingClientRect().width;
  const h = el.clientHeight || el.getBoundingClientRect().height;
  const dpr = window.devicePixelRatio || 1;
  return {
    w,
    h,
    dpr,
    pw: Math.max(1, Math.round(w * dpr)),
    ph: Math.max(1, Math.round(h * dpr)),
  };
}

/** Check whether a key event represents a printable character (not a modifier/control key). */
function isPrintableKey(event: KeyboardEvent): boolean {
  if (event.ctrlKey || event.metaKey || event.altKey) return false;
  // Single character keys are printable
  if (event.key.length === 1) return true;
  return false;
}

/** Dispatch a custom event on the container. */
function emitCustomEvent(
  container: HTMLElement,
  name: string,
  detail?: unknown,
) {
  container.dispatchEvent(new CustomEvent(name, { detail, bubbles: true }));
}

/** Fire dirty-state change event after any mutating operation. */
function emitDirtyEvent(container: HTMLElement, editor: XlEdit) {
  try {
    emitCustomEvent(container, "xlview-dirty", { dirty: editor.is_dirty() });
  } catch {
    /* editor may be freed */
  }
}

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
  readonly editor: XlEdit;
  /** Optional save callback invoked on Ctrl+S / Cmd+S */
  onSave?: () => void;
  /** Optional callback invoked when dirty state changes */
  onDirtyChange?: (dirty: boolean) => void;
}

export async function mountEditor(
  container: HTMLElement,
): Promise<MountedEditor> {
  await init();

  const dpr = window.devicePixelRatio || 1;
  const pw = Math.max(1, Math.round((container.clientWidth || 300) * dpr));
  const ph = Math.max(1, Math.round((container.clientHeight || 150) * dpr));

  const base = document.createElement("canvas");
  const overlay = document.createElement("canvas");
  base.width = pw;
  base.height = ph;
  overlay.width = pw;
  overlay.height = ph;

  // Ensure positioning context
  const pos = getComputedStyle(container).position;
  if (!pos || pos === "static") container.style.position = "relative";

  container.appendChild(base);
  container.appendChild(overlay);

  const editor = new XlEdit(base, overlay, dpr);

  // RAF-batched render callback
  let pending = false;
  editor.set_render_callback(() => {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => {
      pending = false;
      try {
        editor.render();
      } catch {
        /* freed */
      }
    });
  });

  // Resize handler
  const doResize = () => {
    const { w, h, dpr: d, pw: rpw, ph: rph } = viewportSize(container);
    overlay.width = rpw;
    overlay.height = rph;
    overlay.style.width = w + "px";
    overlay.style.height = h + "px";
    editor.resize(rpw, rph, d);
  };
  await new Promise<void>((r) =>
    requestAnimationFrame(() => {
      doResize();
      r();
    }),
  );

  // ResizeObserver
  let destroyed = false;
  const ro = new ResizeObserver(() => {
    if (!destroyed) doResize();
  });
  ro.observe(container);

  // --- Mounted object (created early so handlers can reference it) ---
  const mounted: MountedEditor = {
    load(data: Uint8Array) {
      editor.load(data);
      requestAnimationFrame(doResize);
    },
    save(): Uint8Array {
      return editor.save();
    },
    download(filename = "edited.xlsx") {
      const bytes = editor.save();
      const blob = new Blob([bytes as unknown as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    },
    destroy() {
      destroyed = true;
      ro.disconnect();
      observer.disconnect();
      container.removeEventListener("dblclick", onDblClick);
      container.removeEventListener("keydown", onContainerKeydown);
      editor.free();
    },
    get editor() {
      return editor;
    },
    onSave: undefined,
    onDirtyChange: undefined,
  };

  // --- Helpers for dirty-state notification ---
  let lastDirty: boolean | null = null;
  function notifyDirty() {
    try {
      const dirty = editor.is_dirty();
      emitDirtyEvent(container, editor);
      if (dirty !== lastDirty) {
        lastDirty = dirty;
        mounted.onDirtyChange?.(dirty);
      }
    } catch {
      /* editor may be freed */
    }
  }

  // --- Input element event wiring ---
  let inputEl: HTMLInputElement | null = null;

  function commitFromInput() {
    if (!editor.is_editing()) return;
    const value = editor.input_value();
    if (value != null) {
      editor.commit_edit(value);
      notifyDirty();
    }
  }

  function attachInputHandlers(input: HTMLInputElement) {
    if (inputEl === input) return;
    inputEl = input;

    input.addEventListener("blur", () => {
      // Small delay so click-based commits (Enter/Tab) fire first
      setTimeout(() => {
        if (editor.is_editing()) commitFromInput();
      }, 0);
    });

    input.addEventListener("keydown", (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        editor.cancel_edit();
        event.preventDefault();
        return;
      }

      if (event.key === "Enter" || event.key === "Tab") {
        // Commit current edit
        commitFromInput();

        // Determine navigation direction
        const sel = editor.get_selection();
        if (sel && sel.length >= 2) {
          let newRow = sel[0];
          let newCol = sel[1];

          if (event.key === "Enter") {
            // Enter: down, Shift+Enter: up
            newRow = event.shiftKey ? Math.max(0, newRow - 1) : newRow + 1;
          } else {
            // Tab: right, Shift+Tab: left
            newCol = event.shiftKey ? Math.max(0, newCol - 1) : newCol + 1;
          }

          editor.set_selection(newRow, newCol);
          editor.begin_edit(newRow, newCol);

          // Re-attach input handlers if a new input was created
          const inp = container.querySelector("input");
          if (inp) attachInputHandlers(inp);
        }

        event.preventDefault();
        return;
      }
    });
  }

  // Watch for Rust-created <input> appearing in the container
  const observer = new MutationObserver((mutations) => {
    for (const m of mutations) {
      for (let i = 0; i < m.addedNodes.length; i++) {
        const node = m.addedNodes[i];
        if (node instanceof HTMLInputElement) attachInputHandlers(node);
      }
    }
  });
  observer.observe(container, { childList: true, subtree: true });

  // Also check for any input already present
  const existing = container.querySelector("input");
  if (existing) attachInputHandlers(existing);

  // --- Container-level keydown for type-to-replace, Delete, undo/redo, Ctrl+S ---
  function onContainerKeydown(event: KeyboardEvent) {
    // --- Ctrl+S / Cmd+S: save ---
    if ((event.ctrlKey || event.metaKey) && event.key === "s") {
      event.preventDefault();
      emitCustomEvent(container, "xlview-save");
      mounted.onSave?.();
      return;
    }

    // --- Ctrl+Z / Cmd+Z: undo, Ctrl+Shift+Z / Cmd+Shift+Z: redo ---
    if ((event.ctrlKey || event.metaKey) && event.key === "z") {
      event.preventDefault();
      if (event.shiftKey) {
        editor.redo();
      } else {
        editor.undo();
      }
      notifyDirty();
      return;
    }

    // --- Ctrl+Y / Cmd+Y: redo ---
    if ((event.ctrlKey || event.metaKey) && event.key === "y") {
      event.preventDefault();
      editor.redo();
      notifyDirty();
      return;
    }

    // Everything below only applies when NOT currently editing
    if (editor.is_editing()) return;

    // --- Delete / Backspace: clear selected cell ---
    if (event.key === "Delete" || event.key === "Backspace") {
      editor.delete_selected_cell();
      notifyDirty();
      event.preventDefault();
      return;
    }

    // --- Type-to-replace: printable character starts editing ---
    if (isPrintableKey(event)) {
      const sel = editor.get_selection();
      if (sel && sel.length >= 2) {
        const typed = event.key;
        editor.begin_edit(sel[0], sel[1]);

        // Re-attach input handlers if a new input was created
        const inp = container.querySelector("input");
        if (inp) attachInputHandlers(inp);

        // After a microtask, replace the input value with just the typed character
        setTimeout(() => {
          const input = container.querySelector("input");
          if (input) {
            input.value = typed;
            // Trigger an input event so Rust side picks up the new value
            input.dispatchEvent(new Event("input", { bubbles: true }));
          }
        }, 0);

        event.preventDefault();
        return;
      }
    }
  }

  container.addEventListener("keydown", onContainerKeydown);

  // Make container focusable so it can receive keydown events
  if (!container.hasAttribute("tabindex")) {
    container.setAttribute("tabindex", "0");
  }

  // Double-click to begin editing
  function onDblClick(event: MouseEvent) {
    const rect = container.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const cell = editor.cell_at_point(x, y);
    if (cell && cell.length >= 2) {
      editor.begin_edit(cell[0], cell[1]);
      // Re-check for input after begin_edit creates it
      const inp = container.querySelector("input");
      if (inp) attachInputHandlers(inp);
    }
  }
  container.addEventListener("dblclick", onDblClick);

  return mounted;
}
