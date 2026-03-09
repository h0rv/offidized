import _init, { XlView } from "../pkg/offidized_xlview.js";
import type { InitInput } from "../pkg/offidized_xlview.js";

// Re-export WASM bindings
export { XlView };

// Deduplicate init() — safe to call from multiple <xl-view> elements
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

export interface MountedViewer {
  load(data: Uint8Array): void;
  destroy(): void;
  readonly viewer: XlView;
}

export async function mount(container: HTMLElement): Promise<MountedViewer> {
  await init();

  // Create canvases with initial dimensions
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

  const viewer = XlView.newWithOverlay(base, overlay, dpr);

  // RAF-batched render callback (same pattern as demo/viewer.ts)
  let pending = false;
  viewer.set_render_callback(() => {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => {
      pending = false;
      try {
        viewer.render();
      } catch {
        /* freed */
      }
    });
  });

  // After setup_native_scroll restructures DOM, wait a frame for layout
  const doResize = () => {
    const { w, h, dpr: d, pw: rpw, ph: rph } = viewportSize(container);
    overlay.width = rpw;
    overlay.height = rph;
    overlay.style.width = w + "px";
    overlay.style.height = h + "px";
    viewer.resize(rpw, rph, d);
  };
  await new Promise<void>((r) =>
    requestAnimationFrame(() => {
      doResize();
      r();
    }),
  );

  // ResizeObserver for automatic resize
  let destroyed = false;
  const ro = new ResizeObserver(() => {
    if (!destroyed) doResize();
  });
  ro.observe(container);

  return {
    load(data: Uint8Array) {
      viewer.load(data);
      requestAnimationFrame(doResize); // recalc scroll spacer after load
    },
    destroy() {
      destroyed = true;
      ro.disconnect();
      viewer.free();
    },
    get viewer() {
      return viewer;
    },
  };
}

// Custom Element
class XlViewElement extends HTMLElement {
  private _mounted: MountedViewer | null = null;
  private _shadow: ShadowRoot;
  private _container: HTMLDivElement;

  static get observedAttributes() {
    return ["src"];
  }

  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: "open" });
    this._container = document.createElement("div");
    this._container.style.cssText =
      "width:100%;height:100%;position:relative;overflow:hidden";
    this._shadow.appendChild(this._container);
  }

  async connectedCallback() {
    this._mounted = await mount(this._container);
    const src = this.getAttribute("src");
    if (src) this._loadUrl(src);
    this.dispatchEvent(new Event("ready"));
  }

  disconnectedCallback() {
    this._mounted?.destroy();
    this._mounted = null;
  }

  attributeChangedCallback(
    name: string,
    old: string | null,
    val: string | null,
  ) {
    if (name === "src" && val && val !== old && this._mounted)
      this._loadUrl(val);
  }

  private async _loadUrl(url: string) {
    const res = await fetch(url);
    if (!res.ok) {
      this.dispatchEvent(
        new CustomEvent("error", { detail: `HTTP ${res.status}` }),
      );
      return;
    }
    this._mounted?.load(new Uint8Array(await res.arrayBuffer()));
  }

  /** Load XLSX from bytes */
  load(data: Uint8Array) {
    this._mounted?.load(data);
  }

  /** Access underlying XlView */
  get xlview(): XlView | null {
    return this._mounted?.viewer ?? null;
  }
}

customElements.define("xl-view", XlViewElement);
export { XlViewElement };
