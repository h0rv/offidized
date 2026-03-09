// Canvas page pool with viewport virtualization.
//
// Only creates canvas elements for pages visible in the scroll viewport
// (plus a 1-page buffer above and below). Recycles canvases as the user
// scrolls, keeping DOM node count bounded regardless of document length.

/** Clamp raster DPR to keep buffers sharp without runaway memory use. */
function currentDevicePixelRatio(): number {
  const dpr = window.devicePixelRatio || 1;
  return Math.min(3, Math.max(1, dpr));
}

/** Gap between pages in pixels (matches .docview-paginated gap). */
const PAGE_GAP_PX = 28;

interface PageSlot {
  /** Page width in CSS points. */
  widthPt: number;
  /** Page height in CSS points. */
  heightPt: number;
  /** The canvas element, if currently materialized. */
  canvas: HTMLCanvasElement | null;
  /** The placeholder div that reserves vertical space. */
  placeholder: HTMLDivElement;
}

/**
 * Manages a pool of page canvas elements with viewport-based virtualization.
 *
 * Pages outside the visible scroll region (plus a 1-page buffer) are
 * represented by lightweight placeholder divs. When a page enters the
 * visible range, a canvas element is created (or recycled) and sized at
 * the current device pixel ratio for sharp rendering.
 */
export class CanvasPagePool {
  private scrollContainer: HTMLElement;
  private wrapper: HTMLElement;
  private slots: PageSlot[] = [];
  private visibleStart = 0;
  private visibleEnd = -1;
  private recycledCanvases: HTMLCanvasElement[] = [];
  private rasterDpr = currentDevicePixelRatio();

  constructor(scrollContainer: HTMLElement, wrapper: HTMLElement) {
    this.scrollContainer = scrollContainer;
    this.wrapper = wrapper;
  }

  /**
   * Configure page geometry. Rebuilds all placeholder slots and resets
   * visibility. Must be called before `updateVisibility()`.
   */
  setPageCount(
    count: number,
    pageWidths: number[],
    pageHeights: number[],
  ): void {
    this.rasterDpr = currentDevicePixelRatio();
    // Tear down existing slots
    this.destroySlots();

    for (let i = 0; i < count; i++) {
      const widthPt = pageWidths[i] ?? 612;
      const heightPt = pageHeights[i] ?? 792;

      // Placeholder div reserves the full page box.
      const placeholder = document.createElement("div");
      placeholder.style.width = widthPt + "pt";
      placeholder.style.height = heightPt + "pt";
      placeholder.style.position = "relative";
      placeholder.style.margin = "0 auto";
      placeholder.style.flexShrink = "0";

      this.wrapper.appendChild(placeholder);

      this.slots.push({
        widthPt,
        heightPt,
        canvas: null,
        placeholder,
      });
    }

    // Style the wrapper as a flex column centered layout
    this.wrapper.style.display = "flex";
    this.wrapper.style.flexDirection = "column";
    this.wrapper.style.alignItems = "center";
    this.wrapper.style.gap = PAGE_GAP_PX + "px";
    this.wrapper.style.paddingTop = PAGE_GAP_PX + "px";
    this.wrapper.style.paddingBottom = PAGE_GAP_PX + "px";

    this.visibleStart = 0;
    this.visibleEnd = -1;
  }

  /**
   * Returns the range of page indices currently considered visible
   * (including the 1-page buffer).
   */
  getVisibleRange(): { start: number; end: number } {
    return { start: this.visibleStart, end: this.visibleEnd };
  }

  /**
   * Returns the canvas for the given page index, or null if the page
   * is not currently in the visible range.
   */
  getCanvas(pageIndex: number): HTMLCanvasElement | null {
    const slot = this.slots[pageIndex];
    return slot?.canvas ?? null;
  }

  /**
   * Recalculate which pages are visible based on the scroll container's
   * current scroll position and viewport height. Creates canvases for
   * newly visible pages and recycles canvases for pages that scrolled
   * out of view.
   */
  updateVisibility(): void {
    if (this.slots.length === 0) return;
    this.syncDevicePixelRatio();

    const viewportHeight = this.scrollContainer.clientHeight;

    // Find visible range by checking placeholder positions
    let newStart = this.slots.length;
    let newEnd = -1;

    for (let i = 0; i < this.slots.length; i++) {
      const slot = this.slots[i]!;
      const rect = slot.placeholder.getBoundingClientRect();
      const containerRect = this.scrollContainer.getBoundingClientRect();

      // Page position relative to the scroll container's viewport
      const pageTop = rect.top - containerRect.top;
      const pageBottom = rect.bottom - containerRect.top;

      // Page is visible if it overlaps the viewport
      if (pageBottom > 0 && pageTop < viewportHeight) {
        if (i < newStart) newStart = i;
        if (i > newEnd) newEnd = i;
      }
    }

    // If nothing visible, default to first page
    if (newEnd < newStart) {
      newStart = 0;
      newEnd = 0;
    }

    // Expand by 1-page buffer
    const bufferedStart = Math.max(0, newStart - 1);
    const bufferedEnd = Math.min(this.slots.length - 1, newEnd + 1);

    // Recycle canvases that moved out of range
    for (let i = this.visibleStart; i <= this.visibleEnd; i++) {
      if (i < bufferedStart || i > bufferedEnd) {
        this.detachCanvas(i);
      }
    }

    // Attach canvases for pages that moved into range
    for (let i = bufferedStart; i <= bufferedEnd; i++) {
      if (!this.slots[i]!.canvas) {
        this.attachCanvas(i);
      }
    }

    this.visibleStart = bufferedStart;
    this.visibleEnd = bufferedEnd;
  }

  /**
   * Resizes attached canvases when device pixel ratio changes.
   * Returns true when at least one canvas buffer was resized.
   */
  syncDevicePixelRatio(): boolean {
    const nextDpr = currentDevicePixelRatio();
    if (Math.abs(nextDpr - this.rasterDpr) < 0.01) return false;
    this.rasterDpr = nextDpr;

    for (const slot of this.slots) {
      if (slot.canvas) {
        this.sizeCanvasBuffer(slot.canvas, slot);
      }
    }
    return true;
  }

  /** Clean up all canvases and DOM nodes. */
  destroy(): void {
    this.destroySlots();
    this.recycledCanvases.length = 0;
    this.wrapper.innerHTML = "";
  }

  // ---- internal ----

  private attachCanvas(index: number): void {
    const slot = this.slots[index]!;
    if (slot.canvas) return;

    const canvas =
      this.recycledCanvases.pop() ?? document.createElement("canvas");

    // CSS size matches page dimensions in points
    canvas.style.width = slot.widthPt + "pt";
    canvas.style.height = slot.heightPt + "pt";
    canvas.style.display = "block";
    canvas.style.position = "absolute";
    canvas.style.top = "0";
    canvas.style.left = "0";
    canvas.style.cursor = "text";
    canvas.style.background = "var(--docview-page-bg, #fff)";
    canvas.style.boxShadow =
      "var(--docview-page-shadow, 0 0 0 .75pt #d1d1d1, 0 2px 8px 2px rgba(60,64,67,.16))";
    this.sizeCanvasBuffer(canvas, slot);

    // Paper shadow — reuses the .docview-page look
    canvas.className = "docview-page";

    slot.placeholder.appendChild(canvas);
    slot.canvas = canvas;
  }

  private detachCanvas(index: number): void {
    const slot = this.slots[index];
    if (!slot?.canvas) return;

    slot.placeholder.removeChild(slot.canvas);
    this.recycledCanvases.push(slot.canvas);
    slot.canvas = null;
  }

  private sizeCanvasBuffer(canvas: HTMLCanvasElement, slot: PageSlot): void {
    // Convert pt to CSS px: 1pt = 96/72 px
    const cssWidthPx = slot.widthPt * (96 / 72);
    const cssHeightPx = slot.heightPt * (96 / 72);
    canvas.width = Math.ceil(cssWidthPx * this.rasterDpr);
    canvas.height = Math.ceil(cssHeightPx * this.rasterDpr);
  }

  private destroySlots(): void {
    for (let i = 0; i < this.slots.length; i++) {
      const slot = this.slots[i]!;
      if (slot.canvas) {
        slot.placeholder.removeChild(slot.canvas);
      }
      if (slot.placeholder.parentNode === this.wrapper) {
        this.wrapper.removeChild(slot.placeholder);
      }
    }
    this.slots.length = 0;
    this.recycledCanvases.length = 0;
  }
}
