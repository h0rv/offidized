// Web component and mount() API for offidized-docview.

import _init, { DocView } from "../pkg/offidized_docview.js";
import type { DocViewModel } from "./types.ts";
import { DocRenderer, type RenderMode } from "./renderer.ts";
import { VIEWER_CSS } from "./styles.ts";
import {
  configureFontManifestFallback,
  type LocalFontManifestProvider,
} from "./canvas/font-manager.ts";

type InitInput = Parameters<typeof _init>[0];

// Deduplicate init() — safe to call from multiple <doc-view> elements.
let _initPromise: Promise<void> | null = null;

/** Initialize the WASM module. Safe to call multiple times. */
export function init(input?: InitInput): Promise<void> {
  if (!_initPromise) {
    _initPromise = _init(input ?? "/pkg/offidized_docview_bg.wasm").then(
      () => {},
    );
  }
  return _initPromise!;
}

export { DocView };

export interface RendererFeatureFlags {
  /**
   * Controls which renderer is used when `opts.renderer` is omitted.
   * Defaults to `true` for backward compatibility.
   */
  canvasDefaultRenderer?: boolean;
}

export interface MountViewerOpts {
  renderer?: "html" | "canvas";
  featureFlags?: RendererFeatureFlags;
  /**
   * Optional local/offline font manifest provider used by CanvasKit startup.
   * Set this before first canvas mount to avoid remote-only font dependency.
   */
  localFontManifestProvider?: LocalFontManifestProvider | null;
  /**
   * Keep remote bundled-font fallback enabled by default.
   * Set to `false` only if your local manifest is complete.
   */
  allowRemoteFontFallback?: boolean;
}

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

function resolveRendererType(opts?: MountViewerOpts): "html" | "canvas" {
  if (opts?.renderer) return opts.renderer;
  return (opts?.featureFlags?.canvasDefaultRenderer ?? true)
    ? "canvas"
    : "html";
}

/** Inject VIEWER_CSS into document.head (deduplicated). */
function ensureStyles(): void {
  if (document.querySelector("style[data-docview]")) return;
  const style = document.createElement("style");
  style.setAttribute("data-docview", "");
  style.textContent = VIEWER_CSS;
  document.head.appendChild(style);
}

export interface MountedViewer {
  /** Load a .docx file from bytes. Returns the parse time in ms. */
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
 * @param opts.renderer - `"canvas"` (default) lazy-loads the CanvasKit-based
 *   renderer; `"html"` uses DOM-based rendering.
 */
export async function mount(
  container: HTMLElement,
  opts?: MountViewerOpts,
): Promise<MountedViewer> {
  await init();
  ensureStyles();

  const renderTarget = document.createElement("div");
  renderTarget.className = "docview-root";
  container.appendChild(renderTarget);

  const docview = new DocView();
  const rendererType = resolveRendererType(opts);

  if (rendererType === "canvas") {
    configureFontManifestFallback({
      localManifestProvider: opts?.localFontManifestProvider,
      allowRemoteFallback: opts?.allowRemoteFontFallback,
    });
    const { CanvasRenderer } = await import("./canvas/canvas-renderer.ts");
    const canvasRenderer = await CanvasRenderer.create(renderTarget);

    return {
      load(data: Uint8Array): number {
        const t0 = performance.now();
        const model = docview.parse(data) as DocViewModel;
        const elapsed = performance.now() - t0;
        canvasRenderer.renderModel(model);
        return Math.round(elapsed);
      },
      setMode(mode: RenderMode): void {
        // Canvas renderer is currently paginated-only.
        void mode;
      },
      getMode(): RenderMode {
        return "paginated";
      },
      destroy(): void {
        canvasRenderer.destroy();
        container.removeChild(renderTarget);
        docview.free();
      },
    };
  }

  const renderer = new DocRenderer(renderTarget);

  return {
    load(data: Uint8Array): number {
      const t0 = performance.now();
      const model = docview.parse(data) as DocViewModel;
      const elapsed = performance.now() - t0;
      renderer.setModel(model);
      return Math.round(elapsed);
    },
    setMode(mode: RenderMode): void {
      renderer.setMode(mode);
    },
    getMode(): RenderMode {
      return renderer.getMode();
    },
    destroy(): void {
      container.removeChild(renderTarget);
      docview.free();
    },
  };
}

/** `<doc-view>` custom element with Shadow DOM. */
class DocViewElement extends HTMLElement {
  private _mounted: MountedViewer | null = null;
  private _shadow: ShadowRoot;
  private _container: HTMLDivElement;

  static get observedAttributes(): string[] {
    return ["src", "mode", "renderer"];
  }

  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: "open" });

    // Inject styles into shadow DOM
    const style = document.createElement("style");
    style.textContent = VIEWER_CSS;
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
    this._mounted = await mount(this._container, {
      renderer,
      featureFlags:
        canvasDefaultRenderer === undefined
          ? undefined
          : { canvasDefaultRenderer },
    });

    const mode = this.getAttribute("mode");
    if (mode === "paginated" || mode === "continuous") {
      this._mounted.setMode(mode);
    }

    const src = this.getAttribute("src");
    if (src) this._loadUrl(src);

    this.dispatchEvent(new Event("ready"));
  }

  disconnectedCallback(): void {
    this._mounted?.destroy();
    this._mounted = null;
  }

  attributeChangedCallback(
    name: string,
    old: string | null,
    val: string | null,
  ): void {
    if (name === "src" && val && val !== old && this._mounted) {
      this._loadUrl(val);
    }
    if (
      name === "mode" &&
      val &&
      val !== old &&
      this._mounted &&
      (val === "continuous" || val === "paginated")
    ) {
      this._mounted.setMode(val);
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
    this._mounted?.load(new Uint8Array(await res.arrayBuffer()));
  }

  /** Load .docx from bytes. */
  load(data: Uint8Array): number {
    return this._mounted?.load(data) ?? 0;
  }

  /** Set render mode. */
  setMode(mode: RenderMode): void {
    this._mounted?.setMode(mode);
  }
}

customElements.define("doc-view", DocViewElement);
export { DocViewElement };
