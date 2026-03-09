// Contenteditable-based editor component for offidized-docview.
//
// Thin wrapper that wires together the renderer adapter (HtmlRenderer or
// CanvasRenderer), input adapter, and editor controller.

import _init, { DocEdit } from "../pkg/offidized_docview.js";
import { VIEWER_CSS } from "./styles.ts";
import { HtmlRenderer, EDITOR_CSS } from "./html/html-renderer.ts";
import { HtmlInput } from "./html/html-input.ts";
import { createEditorController } from "./editor.ts";
import {
  configureFontManifestFallback,
  type LocalFontManifestProvider,
} from "./canvas/font-manager.ts";

// Re-export types from adapter.ts for external consumers.
export type {
  FormatAction,
  FormattingState,
  ImageBlockAlignment,
  InlineImageInsertPayload,
  SelectedInlineImageState,
  TextStylePatch,
  ParagraphStylePatch,
  SyncConfig,
  DocEditorController,
  TableCellState,
} from "./adapter.ts";
import type { SyncConfig, DocEditorController } from "./adapter.ts";

/** Options bag for mountEditor (new form). */
export interface MountEditorOpts {
  renderer?: "html" | "canvas";
  featureFlags?: {
    /**
     * Controls default renderer choice when `renderer` is omitted.
     * Defaults to `true` for backward compatibility.
     */
    canvasDefaultRenderer?: boolean;
  };
  /**
   * Optional local/offline font manifest provider used by CanvasKit startup.
   */
  localFontManifestProvider?: LocalFontManifestProvider | null;
  /**
   * Keep remote bundled-font fallback enabled by default.
   */
  allowRemoteFontFallback?: boolean;
  sync?: SyncConfig;
}

type InitInput = Parameters<typeof _init>[0];

// Deduplicate init() -- safe to call from multiple <doc-edit> elements.
let _initPromise: Promise<void> | null = null;

/** Initialize the WASM module. Safe to call multiple times. */
export function init(input?: InitInput): Promise<void> {
  if (!_initPromise) {
    _initPromise = _init(input).then(() => {});
  }
  return _initPromise!;
}

export { DocEdit };

function parseBooleanAttribute(value: string | null): boolean | undefined {
  if (value === null) return undefined;
  const normalized = value.trim().toLowerCase();
  if (!normalized) return true;
  if (
    normalized === "0" ||
    normalized === "false" ||
    normalized === "off" ||
    normalized === "no"
  ) {
    return false;
  }
  return true;
}

function resolveRendererType(opts?: MountEditorOpts): "html" | "canvas" {
  if (opts?.renderer) return opts.renderer;
  return (opts?.featureFlags?.canvasDefaultRenderer ?? true)
    ? "canvas"
    : "html";
}

/** Inject VIEWER_CSS + EDITOR_CSS into document.head (deduplicated). */
function ensureStyles(): void {
  if (document.querySelector("style[data-docedit]")) return;
  const style = document.createElement("style");
  style.setAttribute("data-docedit", "");
  style.textContent = VIEWER_CSS + EDITOR_CSS;
  document.head.appendChild(style);
}

// ---------------------------------------------------------------------------
// mountEditor
// ---------------------------------------------------------------------------

/**
 * Mount an editing instance into a container element.
 *
 * Accepts either the legacy `SyncConfig` directly (backward compat) or a
 * `MountEditorOpts` bag with `renderer` and `sync` fields.
 *
 * When `renderer === "canvas"` (default), lazy-imports CanvasRenderer +
 * CanvasInput. `renderer === "html"` uses HtmlRenderer + HtmlInput.
 */
export async function mountEditor(
  container: HTMLElement,
  syncConfigOrOpts?: SyncConfig | MountEditorOpts,
): Promise<DocEditorController> {
  await init();
  ensureStyles();

  // Disambiguate legacy SyncConfig vs opts bag.
  let syncConfig: SyncConfig | undefined;
  let opts: MountEditorOpts | undefined;

  if (syncConfigOrOpts) {
    if ("roomId" in syncConfigOrOpts) {
      // Legacy: mountEditor(container, { roomId: "..." })
      syncConfig = syncConfigOrOpts as SyncConfig;
    } else {
      // New: mountEditor(container, { renderer: "canvas", sync: {...} })
      opts = syncConfigOrOpts as MountEditorOpts;
      syncConfig = opts.sync;
    }
  }

  const rendererType = resolveRendererType(opts);

  if (rendererType === "canvas") {
    configureFontManifestFallback({
      localManifestProvider: opts?.localFontManifestProvider,
      allowRemoteFallback: opts?.allowRemoteFontFallback,
    });
    const { CanvasRenderer } = await import("./canvas/canvas-renderer.ts");
    const { CanvasInput } = await import("./canvas/canvas-input.ts");
    const renderer = await CanvasRenderer.create(container);
    const input = new CanvasInput(container);
    return createEditorController(renderer, input, container, syncConfig);
  }

  const renderer = new HtmlRenderer(container);
  const input = new HtmlInput(renderer.getInputElement());
  return createEditorController(renderer, input, container, syncConfig);
}

// ---------------------------------------------------------------------------
// <doc-edit> custom element
// ---------------------------------------------------------------------------

/** `<doc-edit>` custom element with Shadow DOM for document editing. */
class DocEditElement extends HTMLElement {
  private _controller: DocEditorController | null = null;
  private _shadow: ShadowRoot;
  private _container: HTMLDivElement;

  static get observedAttributes(): string[] {
    return ["src"];
  }

  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: "open" });

    // Inject styles into shadow DOM (HtmlRenderer injects into document.head,
    // but Shadow DOM needs its own copy).
    const style = document.createElement("style");
    style.textContent = VIEWER_CSS + EDITOR_CSS;
    this._shadow.appendChild(style);

    this._container = document.createElement("div");
    this._container.className = "docview-root";
    this._container.style.cssText =
      "width:100%;height:100%;overflow:auto;position:relative;";
    this._shadow.appendChild(this._container);
  }

  async connectedCallback(): Promise<void> {
    const rendererAttr = this.getAttribute("renderer");
    const renderer =
      rendererAttr === "canvas" || rendererAttr === "html"
        ? rendererAttr
        : undefined;
    const canvasDefaultRenderer = parseBooleanAttribute(
      this.getAttribute("canvas-default-renderer"),
    );
    this._controller = await mountEditor(this._container, {
      renderer,
      featureFlags:
        canvasDefaultRenderer === undefined
          ? undefined
          : { canvasDefaultRenderer },
    });

    const src = this.getAttribute("src");
    if (src) this._loadUrl(src);

    this.dispatchEvent(new Event("ready"));
  }

  disconnectedCallback(): void {
    this._controller?.destroy();
    this._controller = null;
  }

  attributeChangedCallback(
    name: string,
    old: string | null,
    val: string | null,
  ): void {
    if (name === "src" && val && val !== old && this._controller) {
      this._loadUrl(val);
    }
  }

  private async _loadUrl(url: string): Promise<void> {
    const res = await fetch(url);
    if (!res.ok) {
      this.dispatchEvent(
        new CustomEvent("error", { detail: `HTTP ${res.status}` }),
      );
      return;
    }
    this._controller?.load(new Uint8Array(await res.arrayBuffer()));
  }

  /** Load .docx from bytes. Returns parse time in ms. */
  load(data: Uint8Array): number {
    return this._controller?.load(data) ?? 0;
  }

  /** Export the current document as .docx bytes. */
  save(): Uint8Array {
    return this._controller?.save() ?? new Uint8Array(0);
  }

  /** Check if the document has unsaved changes. */
  isDirty(): boolean {
    return this._controller?.isDirty() ?? false;
  }
}

customElements.define("doc-edit", DocEditElement);
export { DocEditElement };
