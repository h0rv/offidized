// HTML (DOM-based) RendererAdapter implementation for offidized-docview.
//
// Wraps the existing EditRenderer + paginator logic, exposing it through the
// RendererAdapter interface so the editor controller is renderer-agnostic.

import type { DocViewModel } from "../types.ts";
import type {
  RendererAdapter,
  DocPosition,
  DocSelection,
  HitTestResult,
  Rect,
  InlineImageHit,
  TableCellHit,
  TableCellPosition,
} from "../adapter.ts";
import { DocRenderer } from "../renderer.ts";
import { VIEWER_CSS } from "../styles.ts";
import {
  charOffsetInElement,
  cursorRectForDocPosition,
  findBodyItemElement,
  resolveSelEndpoint,
  restoreCursor,
  restoreSelection,
  selectionRectsForDocSelection,
  selectionToDocPositions,
} from "../position.ts";

// ---------------------------------------------------------------------------
// Editor CSS (extends VIEWER_CSS with editing-specific styles)
// ---------------------------------------------------------------------------

export const EDITOR_CSS = `
/* Editing surface — transparent wrapper for page cards */
.docedit-surface {
  outline: none;
  cursor: text;
  white-space: pre-wrap;
  word-wrap: break-word;
  overflow-wrap: break-word;
  width: 612pt;
  margin: 24px auto;
  padding: 0;
  background: transparent;
  font-family: "Calibri", "Segoe UI", sans-serif;
  font-size: 11pt;
  line-height: 1.15;
  caret-color: #000;
  box-sizing: border-box;
}

.docedit-surface:focus {
  outline: none;
}

/* Individual page card (letter-sized: 8.5 x 11 in = 612 x 792 pt) */
.docedit-page {
  min-height: 792pt;
  padding: 72pt 72pt;
  background: var(--docview-page-bg, #fff);
  box-shadow: var(--docview-page-shadow, 0 0 0 .75pt #d1d1d1, 0 2px 8px 2px rgba(60,64,67,.16));
  box-sizing: border-box;
  position: relative;
}

/* Gap between page cards — shows container background */
.docedit-page-gap {
  height: 8px;
  user-select: none;
}

/* Empty paragraph placeholder */
.docedit-surface p:empty::before {
  content: " ";
  white-space: pre;
}
`;

/** Inject VIEWER_CSS + EDITOR_CSS into document.head (deduplicated). */
function ensureStyles(): void {
  if (document.querySelector("style[data-docedit]")) return;
  const style = document.createElement("style");
  style.setAttribute("data-docedit", "");
  style.textContent = VIEWER_CSS + EDITOR_CSS;
  document.head.appendChild(style);
}

// ---------------------------------------------------------------------------
// EditRenderer: wraps DocRenderer and adds data-body-index attributes
// ---------------------------------------------------------------------------

/**
 * Renders a DocViewModel into DOM elements annotated with `data-body-index`
 * attributes on each paragraph/table element. This enables mapping browser
 * DOM selections back to CRDT positions.
 */
class EditRenderer {
  private renderer: DocRenderer;
  private container: HTMLElement;
  private model: DocViewModel | null = null;

  constructor(container: HTMLElement) {
    this.container = container;
    this.renderer = new DocRenderer(container);
  }

  renderModel(model: DocViewModel): void {
    this.model = model;
    this.container.innerHTML = "";

    const firstPage = document.createElement("div");
    firstPage.className = "docedit-page";

    for (let i = 0; i < model.body.length; i++) {
      const item = model.body[i]!;
      const el = this.renderer.renderBodyItem(item, model);
      el.setAttribute("data-body-index", String(i));
      if (el.textContent === "" && el.childNodes.length === 0) {
        el.appendChild(document.createElement("br"));
      }
      const brs = el.querySelectorAll("br");
      for (let b = 0; b < brs.length; b++) {
        const br = brs[b]!;
        const next = br.nextSibling;
        if (
          !next ||
          next.nodeType !== Node.TEXT_NODE ||
          next.textContent !== "\u200B"
        ) {
          br.after(document.createTextNode("\u200B"));
        }
      }
      firstPage.appendChild(el);
    }

    if (model.body.length === 0) {
      const p = document.createElement("p");
      p.setAttribute("data-body-index", "0");
      p.appendChild(document.createElement("br"));
      firstPage.appendChild(p);
    }

    this.container.appendChild(firstPage);
    void this.container.offsetHeight;
    this.paginatePages();
  }

  private paginatePages(): void {
    const PAGE_CONTENT_PX = 648 * (96 / 72);

    const singlePage = this.container.querySelector(
      ".docedit-page",
    ) as HTMLElement | null;
    if (!singlePage) return;

    const paragraphs = Array.from(
      singlePage.querySelectorAll("[data-body-index]"),
    ) as HTMLElement[];
    if (paragraphs.length === 0) return;

    const pagePaddingTop = parseFloat(getComputedStyle(singlePage).paddingTop);
    const pageContentStart =
      singlePage.getBoundingClientRect().top + pagePaddingTop;

    const pageGroups: HTMLElement[][] = [[]];
    let pageStartY = 0;

    for (const p of paragraphs) {
      const rect = p.getBoundingClientRect();
      const topInContent = rect.top - pageContentStart;
      const bottomInContent = rect.bottom - pageContentStart;
      const bottomOnPage = bottomInContent - pageStartY;
      const bodyIndex = Number(p.getAttribute("data-body-index") ?? "-1");
      const item =
        bodyIndex >= 0 && this.model ? this.model.body[bodyIndex] : undefined;
      const hasPageBreak =
        item?.type === "paragraph" && (item.pageBreakBefore ?? false);

      if (hasPageBreak && pageGroups[pageGroups.length - 1]!.length > 0) {
        pageStartY = topInContent;
        pageGroups.push([]);
      }

      if (
        bottomOnPage > PAGE_CONTENT_PX &&
        pageGroups[pageGroups.length - 1]!.length > 0
      ) {
        pageStartY = topInContent;
        pageGroups.push([]);
      }

      pageGroups[pageGroups.length - 1]!.push(p);
    }

    if (pageGroups.length <= 1) return;

    singlePage.remove();

    for (let i = 0; i < pageGroups.length; i++) {
      if (i > 0) {
        const gap = document.createElement("div");
        gap.className = "docedit-page-gap";
        gap.contentEditable = "false";
        this.container.appendChild(gap);
      }

      const pageDiv = document.createElement("div");
      pageDiv.className = "docedit-page";
      for (const p of pageGroups[i]!) {
        pageDiv.appendChild(p);
      }
      this.container.appendChild(pageDiv);
    }
  }

  getModel(): DocViewModel | null {
    return this.model;
  }
}

// ---------------------------------------------------------------------------
// HtmlRenderer: RendererAdapter implementation
// ---------------------------------------------------------------------------

export class HtmlRenderer implements RendererAdapter {
  private surface: HTMLElement;
  private editRenderer: EditRenderer;
  private scrollContainer: HTMLElement;

  constructor(container: HTMLElement) {
    ensureStyles();

    this.scrollContainer = container;

    this.surface = document.createElement("div");
    this.surface.className = "docview-root docedit-surface";
    this.surface.contentEditable = "true";
    this.surface.spellcheck = false;
    container.appendChild(this.surface);

    this.editRenderer = new EditRenderer(this.surface);
  }

  renderModel(model: DocViewModel): void {
    this.editRenderer.renderModel(model);
  }

  destroy(): void {
    this.surface.remove();
  }

  hitTest(x: number, y: number): HitTestResult | null {
    // Use caretPositionFromPoint (standard) or caretRangeFromPoint (WebKit)
    let node: Node | null = null;
    let offset = 0;

    if ("caretPositionFromPoint" in document) {
      const pos = (
        document as unknown as {
          caretPositionFromPoint(
            x: number,
            y: number,
          ): { offsetNode: Node; offset: number } | null;
        }
      ).caretPositionFromPoint(x, y);
      if (pos) {
        node = pos.offsetNode;
        offset = pos.offset;
      }
    } else if ("caretRangeFromPoint" in document) {
      const range = (
        document as unknown as {
          caretRangeFromPoint(x: number, y: number): Range | null;
        }
      ).caretRangeFromPoint(x, y);
      if (range) {
        node = range.startContainer;
        offset = range.startOffset;
      }
    }

    if (!node) return null;

    const resolved = resolveSelEndpoint(node, offset);
    if (!resolved) return null;

    return {
      bodyIndex: resolved.bodyIndex,
      charOffset: resolved.charOffset,
      affinity: "leading",
    };
  }

  setCursor(pos: DocPosition): void {
    restoreCursor(this.surface, pos.bodyIndex, pos.charOffset);
  }

  setSelection(sel: DocSelection): void {
    restoreSelection(
      this.surface,
      sel.anchor.bodyIndex,
      sel.anchor.charOffset,
      sel.focus.bodyIndex,
      sel.focus.charOffset,
    );
  }

  getSelection(): DocSelection | null {
    const sel = window.getSelection();
    if (!sel || !sel.anchorNode) return null;
    if (!this.surface.contains(sel.anchorNode)) return null;
    return selectionToDocPositions(sel);
  }

  getCursorRect(): Rect | null {
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return null;

    const range = sel.getRangeAt(0);
    let rect = range.getBoundingClientRect();

    if (rect.height === 0 && sel.anchorNode) {
      const el = findBodyItemElement(sel.anchorNode);
      if (el) rect = el.getBoundingClientRect();
      else return null;
    }

    return { x: rect.left, y: rect.top, w: rect.width, h: rect.height };
  }

  getCursorRectForPosition(pos: DocPosition): Rect | null {
    const rect = cursorRectForDocPosition(
      this.surface,
      pos.bodyIndex,
      pos.charOffset,
    );
    if (!rect) return null;
    return { x: rect.left, y: rect.top, w: rect.width, h: rect.height };
  }

  getSelectionRects(sel: DocSelection): Rect[] {
    return selectionRectsForDocSelection(
      this.surface,
      sel.anchor.bodyIndex,
      sel.anchor.charOffset,
      sel.focus.bodyIndex,
      sel.focus.charOffset,
    ).map((rect) => ({
      x: rect.left,
      y: rect.top,
      w: rect.width,
      h: rect.height,
    }));
  }

  private findInlineImageElement(
    bodyIndex: number,
    charOffset: number,
  ): HTMLImageElement | null {
    const bodyEl = this.surface.querySelector(
      `[data-body-index="${bodyIndex}"]`,
    ) as HTMLElement | null;
    if (!bodyEl) return null;
    const images = bodyEl.querySelectorAll(
      "img.docview-inline-image",
    ) as NodeListOf<HTMLImageElement>;
    for (const image of images) {
      const parent = image.parentNode;
      if (!parent) continue;
      const childIndex = Array.from(parent.childNodes).indexOf(image);
      if (childIndex < 0) continue;
      const imageOffset = charOffsetInElement(bodyEl, parent, childIndex);
      if (imageOffset === charOffset) return image;
    }
    return null;
  }

  getInlineImageAtPoint(x: number, y: number): InlineImageHit | null {
    const hit = document.elementFromPoint(x, y) as HTMLElement | null;
    const image = hit?.closest?.(".docview-inline-image");
    if (!(image instanceof HTMLImageElement)) return null;
    const bodyEl = findBodyItemElement(image);
    if (!bodyEl) return null;
    const bodyIndex = Number.parseInt(
      bodyEl.getAttribute("data-body-index") ?? "-1",
      10,
    );
    if (!Number.isFinite(bodyIndex) || bodyIndex < 0) return null;
    const parent = image.parentNode;
    if (!parent) return null;
    const childIndex = Array.from(parent.childNodes).indexOf(image);
    if (childIndex < 0) return null;
    const charOffset = charOffsetInElement(bodyEl, parent, childIndex);
    const rect = image.getBoundingClientRect();
    return {
      bodyIndex,
      charOffset,
      imageIndex: Number.parseInt(image.dataset.docviewImageIndex ?? "0", 10),
      rect: { x: rect.left, y: rect.top, w: rect.width, h: rect.height },
    };
  }

  getInlineImageRect(pos: DocPosition): Rect | null {
    const image = this.findInlineImageElement(pos.bodyIndex, pos.charOffset);
    if (!image) return null;
    const rect = image.getBoundingClientRect();
    return { x: rect.left, y: rect.top, w: rect.width, h: rect.height };
  }

  getTableCellAtPoint(x: number, y: number): TableCellHit | null {
    const hit = document.elementFromPoint(x, y) as HTMLElement | null;
    const cell = hit?.closest?.(
      "[data-docview-table-cell='1']",
    ) as HTMLElement | null;
    if (!cell) return null;
    const bodyItem = cell.closest("[data-body-index]") as HTMLElement | null;
    if (!bodyItem) return null;
    const bodyIndex = Number.parseInt(
      bodyItem.getAttribute("data-body-index") ?? "-1",
      10,
    );
    const row = Number.parseInt(cell.dataset.docviewTableRow ?? "-1", 10);
    const col = Number.parseInt(cell.dataset.docviewTableCol ?? "-1", 10);
    if (!Number.isFinite(bodyIndex) || bodyIndex < 0) return null;
    if (!Number.isFinite(row) || row < 0) return null;
    if (!Number.isFinite(col) || col < 0) return null;
    const rect = cell.getBoundingClientRect();
    return {
      bodyIndex,
      row,
      col,
      rect: { x: rect.left, y: rect.top, w: rect.width, h: rect.height },
    };
  }

  getTableCellRect(cell: TableCellPosition): Rect | null {
    const bodyItem = this.surface.querySelector(
      `[data-body-index="${cell.bodyIndex}"]`,
    ) as HTMLElement | null;
    if (!bodyItem) return null;
    const td = bodyItem.querySelector(
      `[data-docview-table-cell="1"][data-docview-table-row="${cell.row}"][data-docview-table-col="${cell.col}"]`,
    ) as HTMLElement | null;
    if (!td) return null;
    const rect = td.getBoundingClientRect();
    return { x: rect.left, y: rect.top, w: rect.width, h: rect.height };
  }

  getInputElement(): HTMLElement {
    return this.surface;
  }

  isFocused(): boolean {
    return document.activeElement === this.surface;
  }

  focus(): void {
    this.surface.focus({ preventScroll: true });
  }

  getScrollContainer(): HTMLElement {
    return this.scrollContainer;
  }
}
