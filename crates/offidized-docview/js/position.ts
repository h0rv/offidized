// DOM position utilities for mapping between browser Selection and CRDT positions.
//
// Extracted from doc-edit.ts — these are DOM-specific helpers shared between
// the HTML renderer adapter and the editor controller.

import type { DocPosition, DocSelection } from "./adapter.ts";

/**
 * Find the nearest ancestor element with a `data-body-index` attribute.
 */
export function findBodyItemElement(node: Node): HTMLElement | null {
  let current: Node | null = node;
  while (current) {
    if (
      current instanceof HTMLElement &&
      current.hasAttribute("data-body-index")
    ) {
      return current;
    }
    current = current.parentNode;
  }
  return null;
}

/** Check if an element is a CRDT sentinel (line break or tab). */
export function isSentinelElement(node: Node): boolean {
  if (node instanceof HTMLBRElement) return true;
  if (
    node instanceof HTMLElement &&
    node.classList.contains("docview-inline-image")
  ) {
    return true;
  }
  if (node instanceof HTMLElement && node.classList.contains("docview-tab"))
    return true;
  return false;
}

function isNumberingElement(node: Node): boolean {
  return (
    node instanceof HTMLElement && node.classList.contains("docview-numbering")
  );
}

function isInsideNumbering(node: Node): boolean {
  if (isNumberingElement(node)) return true;
  return (
    node.parentNode instanceof HTMLElement &&
    node.parentNode.classList.contains("docview-numbering")
  );
}

/**
 * Check if a text node is a zero-width-space cursor anchor inserted after
 * a `<br>` sentinel.  These are NOT CRDT content — they exist only so the
 * browser can compute a proper caret height on new lines.
 */
export function isCursorAnchor(node: Node): boolean {
  if (node.nodeType !== Node.TEXT_NODE) return false;
  if (node.textContent !== "\u200B") return false;
  const prev = node.previousSibling;
  return prev instanceof HTMLBRElement;
}

/**
 * Count CRDT character length of a node and its descendants.
 * Like textLength but also counts <br> and .docview-tab as 1 char each.
 * Skips zero-width-space cursor anchors (not CRDT content).
 */
export function crdtCharLength(node: Node): number {
  if (isNumberingElement(node)) return 0;
  if (node.nodeType === Node.TEXT_NODE) {
    if (isInsideNumbering(node)) return 0;
    if (isCursorAnchor(node)) return 0;
    return (node.textContent ?? "").length;
  }
  if (isSentinelElement(node)) return 1;
  let len = 0;
  for (let i = 0; i < node.childNodes.length; i++) {
    len += crdtCharLength(node.childNodes[i]!);
  }
  return len;
}

/**
 * Count the character offset from the start of a body item element
 * to the given node/offset within it.
 *
 * Walks the DOM tree in document order, counting text node characters
 * and sentinel elements (<br>, .docview-tab) as 1 CRDT character each.
 */
export function charOffsetInElement(
  root: HTMLElement,
  targetNode: Node,
  targetOffset: number,
): number {
  let count = 0;

  const walker = document.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT,
  );

  let node = walker.nextNode();
  while (node) {
    if (isNumberingElement(node)) {
      if (node === targetNode) {
        return count;
      }
      node = walker.nextNode();
      continue;
    }

    // Skip zero-width-space cursor anchors — not CRDT content.
    if (isCursorAnchor(node)) {
      if (node === targetNode) {
        return count;
      }
      node = walker.nextNode();
      continue;
    }

    if (node === targetNode) {
      if (node.nodeType === Node.TEXT_NODE) {
        return count + targetOffset;
      } else {
        let childCount = 0;
        for (let i = 0; i < targetOffset && i < node.childNodes.length; i++) {
          childCount += crdtCharLength(node.childNodes[i]!);
        }
        return count + childCount;
      }
    }

    if (node.nodeType === Node.TEXT_NODE) {
      if (isInsideNumbering(node)) {
        node = walker.nextNode();
        continue;
      }
      count += (node.textContent ?? "").length;
    } else if (isSentinelElement(node)) {
      count += 1;
    }

    node = walker.nextNode();
  }

  // If we reach here, the target is the root element itself
  if (targetNode === root) {
    let childCount = 0;
    for (let i = 0; i < targetOffset && i < root.childNodes.length; i++) {
      childCount += crdtCharLength(root.childNodes[i]!);
    }
    return childCount;
  }

  return count;
}

/**
 * Resolve a selection endpoint (node + offset) to a body-item element,
 * body index, and character offset within it.
 *
 * When the node is ABOVE body items (e.g. the surface div after Ctrl+A),
 * we resolve using the child at the given offset:
 *  - offset 0 → first body item, char 0
 *  - offset == childCount → last body item, end of content
 */
export function resolveSelEndpoint(
  node: Node,
  offset: number,
): { el: HTMLElement; bodyIndex: number; charOffset: number } | null {
  const el = findBodyItemElement(node);
  if (el) {
    const bodyIndex = parseInt(el.getAttribute("data-body-index") ?? "", 10);
    if (isNaN(bodyIndex)) return null;
    return { el, bodyIndex, charOffset: charOffsetInElement(el, node, offset) };
  }

  // Node is above body items (e.g. the surface itself after Ctrl+A).
  if (node instanceof HTMLElement) {
    const children = node.querySelectorAll("[data-body-index]");
    if (children.length === 0) return null;

    if (offset === 0) {
      const first = children[0] as HTMLElement;
      const idx = parseInt(first.getAttribute("data-body-index") ?? "0", 10);
      return { el: first, bodyIndex: idx, charOffset: 0 };
    }

    const directChildren = Array.from(node.childNodes);
    const boundaryChild =
      offset < directChildren.length ? directChildren[offset] : null;
    if (boundaryChild instanceof HTMLElement) {
      const boundaryEl =
        boundaryChild.matches?.("[data-body-index]") === true
          ? boundaryChild
          : (boundaryChild.querySelector?.(
              "[data-body-index]",
            ) as HTMLElement | null);
      if (boundaryEl) {
        const idx = parseInt(
          boundaryEl.getAttribute("data-body-index") ?? "0",
          10,
        );
        return { el: boundaryEl, bodyIndex: idx, charOffset: 0 };
      }
    }

    const last = children[children.length - 1] as HTMLElement;
    const idx = parseInt(last.getAttribute("data-body-index") ?? "0", 10);
    return { el: last, bodyIndex: idx, charOffset: crdtCharLength(last) };
  }

  return null;
}

/**
 * Resolve a CRDT (bodyIndex, charOffset) to a DOM position {node, offset}.
 *
 * Walks text nodes and sentinel elements (<br>, .docview-tab), counting
 * each sentinel as 1 CRDT character and skipping ZWSP cursor anchors.
 * Returns null if the body element isn't found.
 */
export function resolveDomPosition(
  surface: HTMLElement,
  bodyIndex: number,
  charOffset: number,
): { node: Node; offset: number } | null {
  const el = surface.querySelector(
    `[data-body-index="${bodyIndex}"]`,
  ) as HTMLElement | null;
  if (!el) return null;

  const walker = document.createTreeWalker(
    el,
    NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT,
  );
  let remaining = charOffset;
  let node = walker.nextNode();

  while (node) {
    if (isCursorAnchor(node)) {
      node = walker.nextNode();
      continue;
    }

    if (node.nodeType === Node.TEXT_NODE) {
      const len = (node.textContent ?? "").length;
      if (remaining <= len) {
        return { node, offset: remaining };
      }
      remaining -= len;
    } else if (isSentinelElement(node)) {
      if (remaining < 1) {
        const parent = node.parentNode;
        if (parent) {
          const idx = Array.from(parent.childNodes).indexOf(node as ChildNode);
          return { node: parent, offset: idx };
        }
      }
      remaining -= 1;
      if (remaining === 0) {
        const nextSibling = node.nextSibling;
        if (
          nextSibling?.nodeType === Node.TEXT_NODE &&
          nextSibling.textContent === "\u200B"
        ) {
          return { node: nextSibling, offset: 1 };
        }
        const parent = node.parentNode;
        if (parent) {
          const idx = Array.from(parent.childNodes).indexOf(node as ChildNode);
          return { node: parent, offset: idx + 1 };
        }
      }
    }
    node = walker.nextNode();
  }

  return { node: el, offset: el.childNodes.length };
}

/**
 * Place a collapsed cursor at (bodyIndex, charOffset) after a re-render.
 */
export function restoreCursor(
  surface: HTMLElement,
  bodyIndex: number,
  charOffset: number,
): void {
  const pos = resolveDomPosition(surface, bodyIndex, charOffset);
  if (!pos) return;
  const sel = window.getSelection();
  if (!sel) return;
  sel.removeAllRanges();
  const range = document.createRange();
  range.setStart(pos.node, pos.offset);
  range.collapse(true);
  sel.addRange(range);
}

/**
 * Restore a (possibly non-collapsed) selection after a re-render.
 * Used by formatting operations to keep the highlighted range visible.
 */
export function restoreSelection(
  surface: HTMLElement,
  anchorBody: number,
  anchorOff: number,
  focusBody: number,
  focusOff: number,
): void {
  const a = resolveDomPosition(surface, anchorBody, anchorOff);
  const f = resolveDomPosition(surface, focusBody, focusOff);
  if (!a || !f) return;
  const sel = window.getSelection();
  if (!sel) return;
  const isBackward =
    anchorBody > focusBody ||
    (anchorBody === focusBody && anchorOff > focusOff);
  sel.removeAllRanges();
  if (isBackward && typeof sel.setBaseAndExtent === "function") {
    try {
      sel.setBaseAndExtent(a.node, a.offset, f.node, f.offset);
      return;
    } catch {
      // Fall back to a normalized forward range below.
    }
  }
  const range = document.createRange();
  if (isBackward) {
    range.setStart(f.node, f.offset);
    range.setEnd(a.node, a.offset);
  } else {
    range.setStart(a.node, a.offset);
    range.setEnd(f.node, f.offset);
  }
  sel.addRange(range);
}

export function selectionRectsForDocSelection(
  surface: HTMLElement,
  anchorBody: number,
  anchorOff: number,
  focusBody: number,
  focusOff: number,
): DOMRect[] {
  const isBackward =
    anchorBody > focusBody ||
    (anchorBody === focusBody && anchorOff > focusOff);
  const startBody = isBackward ? focusBody : anchorBody;
  const startOff = isBackward ? focusOff : anchorOff;
  const endBody = isBackward ? anchorBody : focusBody;
  const endOff = isBackward ? anchorOff : focusOff;
  const a = resolveDomPosition(surface, startBody, startOff);
  const f = resolveDomPosition(surface, endBody, endOff);
  if (!a || !f) return [];
  const range = document.createRange();
  range.setStart(a.node, a.offset);
  range.setEnd(f.node, f.offset);
  return Array.from(range.getClientRects());
}

export function cursorRectForDocPosition(
  surface: HTMLElement,
  bodyIndex: number,
  charOffset: number,
): DOMRect | null {
  const pos = resolveDomPosition(surface, bodyIndex, charOffset);
  if (!pos) return null;
  const range = document.createRange();
  range.setStart(pos.node, pos.offset);
  range.collapse(true);
  let rect = range.getBoundingClientRect();
  if (rect.height === 0) {
    const el = surface.querySelector(
      `[data-body-index="${bodyIndex}"]`,
    ) as HTMLElement | null;
    if (!el) return null;
    rect = el.getBoundingClientRect();
  }
  return rect;
}

/**
 * Convert a browser Selection to DocPosition pairs.
 * Returns null if the selection is outside the editing surface.
 */
export function selectionToDocPositions(sel: Selection): DocSelection | null {
  if (!sel.anchorNode || !sel.focusNode) return null;

  const a = resolveSelEndpoint(sel.anchorNode, sel.anchorOffset);
  const f = resolveSelEndpoint(sel.focusNode, sel.focusOffset);
  if (!a || !f) return null;

  return {
    anchor: { bodyIndex: a.bodyIndex, charOffset: a.charOffset },
    focus: { bodyIndex: f.bodyIndex, charOffset: f.charOffset },
  };
}

/** Safely get body index and char offset for a selection endpoint. */
export function selEndpoint(node: Node | null, offset: number): DocPosition {
  if (!node) return { bodyIndex: 0, charOffset: 0 };
  const resolved = resolveSelEndpoint(node, offset);
  if (!resolved) return { bodyIndex: 0, charOffset: 0 };
  return { bodyIndex: resolved.bodyIndex, charOffset: resolved.charOffset };
}
