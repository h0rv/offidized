// Web component and mount() API for offidized-pptview.

import _init, { PptView } from "../pkg/offidized_pptview.js";
import type { PresentationViewModel } from "./types.ts";
import { SlideRenderer } from "./renderer.ts";
import { Navigator } from "./navigator.ts";
import { VIEWER_CSS } from "./styles.ts";

type InitInput = Parameters<typeof _init>[0];

// Deduplicate init() — safe to call from multiple <ppt-view> elements.
let _initPromise: Promise<void> | null = null;

/** Initialize the WASM module. Safe to call multiple times. */
export function init(input?: InitInput): Promise<void> {
  if (!_initPromise) {
    _initPromise = _init(input ?? "/pkg/offidized_pptview_bg.wasm").then(
      () => {},
    );
  }
  return _initPromise!;
}

export { PptView };

/** Inject VIEWER_CSS into document.head (deduplicated). */
function ensureStyles(): void {
  if (document.querySelector("style[data-pptview]")) return;
  const style = document.createElement("style");
  style.setAttribute("data-pptview", "");
  style.textContent = VIEWER_CSS;
  document.head.appendChild(style);
}

export interface MountedViewer {
  /** Load a .pptx file from bytes. Returns the parse time in ms. */
  load(data: Uint8Array): number;
  /** Navigate to a specific slide (0-indexed). */
  goToSlide(index: number): void;
  /** Go to the next slide. */
  nextSlide(): void;
  /** Go to the previous slide. */
  prevSlide(): void;
  /** Get the total number of slides. */
  slideCount(): number;
  /** Get the current slide index (0-indexed). */
  currentSlide(): number;
  /** Set a callback for slide change events. */
  onSlideChange(cb: (index: number) => void): void;
  /** Destroy the viewer and clean up resources. */
  destroy(): void;
}

/**
 * Mount a pptview instance into a container element.
 * Creates the renderer, filmstrip, and returns a controller.
 */
export async function mount(container: HTMLElement): Promise<MountedViewer> {
  await init();
  ensureStyles();

  const root = document.createElement("div");
  root.className = "pptview-root";
  container.appendChild(root);

  // Filmstrip sidebar
  const filmstrip = document.createElement("div");
  filmstrip.className = "pptview-filmstrip";
  root.appendChild(filmstrip);

  // Main slide area
  const slideArea = document.createElement("div");
  slideArea.className = "pptview-slide-area";
  root.appendChild(slideArea);

  const renderer = new SlideRenderer(slideArea);
  const navigator = new Navigator(filmstrip, renderer);
  navigator.attachKeyboard(window);

  const pptview = new PptView();

  return {
    load(data: Uint8Array): number {
      const t0 = performance.now();
      const model = pptview.parse(data) as PresentationViewModel;
      const elapsed = performance.now() - t0;
      renderer.setModel(model);
      navigator.setModel(model);
      return Math.round(elapsed);
    },
    goToSlide(index: number): void {
      navigator.goToSlide(index);
    },
    nextSlide(): void {
      navigator.nextSlide();
    },
    prevSlide(): void {
      navigator.prevSlide();
    },
    slideCount(): number {
      return navigator.slideCount();
    },
    currentSlide(): number {
      return navigator.getCurrentSlide();
    },
    onSlideChange(cb: (index: number) => void): void {
      navigator.setOnChange(cb);
    },
    destroy(): void {
      navigator.detachKeyboard(window);
      navigator.destroy();
      container.removeChild(root);
      pptview.free();
    },
  };
}

/** `<ppt-view>` custom element with Shadow DOM. */
class PptViewElement extends HTMLElement {
  private _mounted: MountedViewer | null = null;
  private _shadow: ShadowRoot;
  private _container: HTMLDivElement;

  static get observedAttributes(): string[] {
    return ["src"];
  }

  constructor() {
    super();
    this._shadow = this.attachShadow({ mode: "open" });

    // Inject styles into shadow DOM
    const style = document.createElement("style");
    style.textContent = VIEWER_CSS;
    this._shadow.appendChild(style);

    this._container = document.createElement("div");
    this._container.style.cssText =
      "width:100%;height:100%;overflow:hidden;position:relative;";
    this._shadow.appendChild(this._container);
  }

  async connectedCallback(): Promise<void> {
    this._mounted = await mount(this._container);

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

  /** Load .pptx from bytes. */
  load(data: Uint8Array): number {
    return this._mounted?.load(data) ?? 0;
  }

  /** Navigate to a specific slide. */
  goToSlide(index: number): void {
    this._mounted?.goToSlide(index);
  }
}

customElements.define("ppt-view", PptViewElement);
export { PptViewElement };
