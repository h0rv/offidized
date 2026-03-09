// Renderer-agnostic editor controller for offidized-docview.
//
// Takes a RendererAdapter + InputAdapter and orchestrates all editing logic:
// intent generation, CRDT interaction, re-rendering, cursor restoration,
// formatting state, and sync.

import { DocEdit } from "../pkg/offidized_docview.js";
import type { DocViewModel, RunModel } from "./types.ts";
import type {
  RendererAdapter,
  InputAdapter,
  NavigationKey,
  NavigationPayload,
  NormalizedInput,
  DocPosition,
  DocSelection,
  DocEditorController,
  FormattingState,
  FormatAction,
  ImageBlockAlignment,
  ListKind,
  ParagraphAlignment,
  TextStylePatch,
  ParagraphStylePatch,
  SyncConfig,
  PointerDownPayload,
  InlineImageInsertPayload,
  SelectedInlineImageState,
  TableCellState,
  TableCellPosition,
} from "./adapter.ts";
import { SyncProvider, type RemoteAwarenessPeer } from "./sync-provider.ts";
import { resolveSelEndpoint, selectionToDocPositions } from "./position.ts";
import { EditorPerfTracker } from "./editor/perf.ts";

/**
 * Create a renderer-agnostic editor controller.
 *
 * Wires the RendererAdapter and InputAdapter together with the CRDT,
 * handling all editing orchestration.
 */
export function createEditorController(
  renderer: RendererAdapter,
  input: InputAdapter,
  container: HTMLElement,
  syncConfig?: SyncConfig,
): DocEditorController {
  type PendingTextAttrs = Record<string, string | number | boolean | null>;

  let docEdit: DocEdit | null = null;
  let currentModel: DocViewModel | null = null;
  let destroyed = false;
  let syncProvider: SyncProvider | null = null;
  const perfTracker = new EditorPerfTracker(180);
  const useNativeSelection = renderer.getInputElement().isContentEditable;
  let remotePeers: ReadonlyArray<RemoteAwarenessPeer> = [];
  let presenceOverlay: HTMLDivElement | null = null;
  let presenceRaf = 0;
  let imageSelectionOverlay: HTMLDivElement | null = null;
  let imageSelectionBox: HTMLDivElement | null = null;
  let imageResizeHandle: HTMLDivElement | null = null;
  let imageResizeSession: {
    startX: number;
    startY: number;
    startWidthPt: number;
    startHeightPt: number;
    startRect: { x: number; y: number; w: number; h: number };
    bodyIndex: number;
    charOffset: number;
  } | null = null;

  // --- Pending format state ---
  let pendingTextAttrs: PendingTextAttrs = {};
  let activeTableCell: TableCellPosition | null = null;
  let suppressTableCellBlurCommit = false;
  const tableCellCallbacks: Array<(state: TableCellState | null) => void> = [];
  const tableCellEditor = document.createElement("textarea");
  tableCellEditor.dataset.docviewTableCellEditor = "1";
  tableCellEditor.spellcheck = false;
  tableCellEditor.style.position = "absolute";
  tableCellEditor.style.display = "none";
  tableCellEditor.style.zIndex = "30";
  tableCellEditor.style.resize = "none";
  tableCellEditor.style.overflow = "hidden";
  tableCellEditor.style.padding = "4px 6px";
  tableCellEditor.style.border = "2px solid #1a73e8";
  tableCellEditor.style.borderRadius = "2px";
  tableCellEditor.style.background = "#fff";
  tableCellEditor.style.boxShadow = "0 1px 2px rgba(0,0,0,0.14)";
  tableCellEditor.style.font = '11pt "Calibri", "Segoe UI", sans-serif';
  tableCellEditor.style.lineHeight = "1.15";
  tableCellEditor.style.whiteSpace = "pre-wrap";
  container.appendChild(tableCellEditor);

  function togglePendingFormat(action: string): void {
    const state = queryFormattingState();
    const isEnabled =
      action === "bold"
        ? !!state.bold
        : action === "italic"
          ? !!state.italic
          : action === "underline"
            ? !!state.underline
            : action === "strike"
              ? !!state.strike
              : false;
    pendingTextAttrs[action] = isEnabled ? null : true;
  }

  function consumePendingAttrs(): PendingTextAttrs | null {
    const keys = Object.keys(pendingTextAttrs);
    if (keys.length === 0) return null;
    const attrs = { ...pendingTextAttrs };
    pendingTextAttrs = {};
    return attrs;
  }

  function clearPendingFormat(): void {
    pendingTextAttrs = {};
  }

  function normalizeColor(color: string): string {
    return color.replace(/^#/, "").toUpperCase();
  }

  function toIntentTextAttrs(patch: TextStylePatch): PendingTextAttrs {
    const attrs: PendingTextAttrs = {};
    if (patch.bold !== undefined) attrs.bold = patch.bold;
    if (patch.italic !== undefined) attrs.italic = patch.italic;
    if (patch.underline !== undefined) attrs.underline = patch.underline;
    if (patch.strike !== undefined) attrs.strike = patch.strike;
    if (patch.fontFamily !== undefined) attrs.fontFamily = patch.fontFamily;
    if (patch.fontSizePt !== undefined) {
      attrs.fontSizePt =
        patch.fontSizePt == null ? null : Math.max(1, patch.fontSizePt);
    }
    if (patch.color !== undefined) {
      attrs.color =
        typeof patch.color === "string" ? normalizeColor(patch.color) : null;
    }
    if (patch.highlight !== undefined) {
      attrs.highlight = patch.highlight ?? null;
    }
    if (patch.hyperlink !== undefined) {
      attrs.hyperlink = patch.hyperlink ?? null;
    }
    return attrs;
  }

  function toIntentParagraphAttrs(
    patch: ParagraphStylePatch,
  ): Record<string, string | number | null> {
    const attrs: Record<string, string | number | null> = {};
    if (patch.headingLevel !== undefined) {
      attrs.headingLevel =
        patch.headingLevel == null
          ? null
          : Math.min(9, Math.max(1, patch.headingLevel));
    }
    if (patch.alignment !== undefined) {
      attrs.alignment = patch.alignment ?? null;
    }
    if (patch.spacingBeforePt !== undefined) {
      attrs.spacingBeforePt =
        patch.spacingBeforePt == null
          ? null
          : Math.max(0, patch.spacingBeforePt);
    }
    if (patch.spacingAfterPt !== undefined) {
      attrs.spacingAfterPt =
        patch.spacingAfterPt == null ? null : Math.max(0, patch.spacingAfterPt);
    }
    if (patch.indentLeftPt !== undefined) {
      attrs.indentLeftPt =
        patch.indentLeftPt == null ? null : Math.max(0, patch.indentLeftPt);
    }
    if (patch.indentFirstLinePt !== undefined) {
      attrs.indentFirstLinePt =
        patch.indentFirstLinePt == null
          ? null
          : Math.max(0, patch.indentFirstLinePt);
    }
    if (patch.lineSpacingMultiple !== undefined) {
      attrs.lineSpacingMultiple =
        patch.lineSpacingMultiple == null
          ? null
          : Math.max(0.5, patch.lineSpacingMultiple);
    }
    if (patch.numberingKind !== undefined) {
      attrs.numberingKind = patch.numberingKind ?? null;
    }
    if (patch.numberingNumId !== undefined) {
      attrs.numberingNumId =
        patch.numberingNumId == null
          ? null
          : Math.max(1, Math.round(patch.numberingNumId));
    }
    if (patch.numberingIlvl !== undefined) {
      attrs.numberingIlvl =
        patch.numberingIlvl == null
          ? null
          : Math.max(0, Math.min(8, Math.round(patch.numberingIlvl)));
    }
    return attrs;
  }

  function mergePendingTextAttrs(patch: TextStylePatch): void {
    const attrs = toIntentTextAttrs(patch);
    for (const [key, value] of Object.entries(attrs)) {
      if (value === null) {
        pendingTextAttrs[key] = null;
      } else {
        pendingTextAttrs[key] = value;
      }
    }
  }

  function applyPendingFormattingState(state: FormattingState): void {
    if (pendingTextAttrs.bold !== undefined) {
      if (pendingTextAttrs.bold === null) delete state.bold;
      else state.bold = pendingTextAttrs.bold === true;
    }
    if (pendingTextAttrs.italic !== undefined) {
      if (pendingTextAttrs.italic === null) delete state.italic;
      else state.italic = pendingTextAttrs.italic === true;
    }
    if (pendingTextAttrs.underline !== undefined) {
      if (pendingTextAttrs.underline === null) delete state.underline;
      else state.underline = pendingTextAttrs.underline === true;
    }
    if (pendingTextAttrs.strike !== undefined) {
      if (pendingTextAttrs.strike === null) delete state.strike;
      else state.strike = pendingTextAttrs.strike === true;
    }
    if (pendingTextAttrs.fontFamily !== undefined) {
      if (pendingTextAttrs.fontFamily === null) delete state.fontFamily;
      else if (typeof pendingTextAttrs.fontFamily === "string") {
        state.fontFamily = pendingTextAttrs.fontFamily;
      }
    }
    if (pendingTextAttrs.fontSizePt !== undefined) {
      if (pendingTextAttrs.fontSizePt === null) delete state.fontSizePt;
      else if (typeof pendingTextAttrs.fontSizePt === "number") {
        state.fontSizePt = pendingTextAttrs.fontSizePt;
      }
    }
    if (pendingTextAttrs.color !== undefined) {
      if (pendingTextAttrs.color === null) delete state.color;
      else if (typeof pendingTextAttrs.color === "string") {
        state.color = pendingTextAttrs.color;
      }
    }
    if (pendingTextAttrs.highlight !== undefined) {
      if (pendingTextAttrs.highlight === null) delete state.highlight;
      else if (typeof pendingTextAttrs.highlight === "string") {
        state.highlight = pendingTextAttrs.highlight;
      }
    }
    if (pendingTextAttrs.hyperlink !== undefined) {
      if (pendingTextAttrs.hyperlink === null) delete state.hyperlink;
      else if (typeof pendingTextAttrs.hyperlink === "string") {
        state.hyperlink = pendingTextAttrs.hyperlink;
      }
    }
  }

  function seedPendingTextAttrsFromState(state: FormattingState): void {
    pendingTextAttrs = {};
    if (state.bold) pendingTextAttrs.bold = true;
    if (state.italic) pendingTextAttrs.italic = true;
    if (state.underline) pendingTextAttrs.underline = true;
    if (state.strike) pendingTextAttrs.strike = true;
    if (state.fontFamily) pendingTextAttrs.fontFamily = state.fontFamily;
    if (typeof state.fontSizePt === "number") {
      pendingTextAttrs.fontSizePt = state.fontSizePt;
    }
    if (state.color) pendingTextAttrs.color = normalizeColor(state.color);
    if (state.highlight) pendingTextAttrs.highlight = state.highlight;
    if (state.hyperlink) pendingTextAttrs.hyperlink = state.hyperlink;
  }

  // --- Formatting state tracking ---

  let formattingCallbacks: Array<(state: FormattingState) => void> = [];

  interface OrderedSelectionRange {
    start: DocPosition;
    end: DocPosition;
    collapsed: boolean;
  }

  function comparePositions(a: DocPosition, b: DocPosition): number {
    if (a.bodyIndex !== b.bodyIndex) return a.bodyIndex - b.bodyIndex;
    return a.charOffset - b.charOffset;
  }

  function samePosition(a: DocPosition, b: DocPosition): boolean {
    return a.bodyIndex === b.bodyIndex && a.charOffset === b.charOffset;
  }

  function normalizeSelection(sel: DocSelection): OrderedSelectionRange {
    if (comparePositions(sel.anchor, sel.focus) <= 0) {
      return {
        start: { ...sel.anchor },
        end: { ...sel.focus },
        collapsed: samePosition(sel.anchor, sel.focus),
      };
    }
    return {
      start: { ...sel.focus },
      end: { ...sel.anchor },
      collapsed: samePosition(sel.anchor, sel.focus),
    };
  }

  function getParagraphBodyIndices(): number[] {
    if (!currentModel) return [];
    const indices: number[] = [];
    for (let i = 0; i < currentModel.body.length; i++) {
      const item = currentModel.body[i];
      if (item?.type === "paragraph") indices.push(i);
    }
    return indices;
  }

  function tableItemAt(
    bodyIndex: number,
  ): Extract<DocViewModel["body"][number], { type: "table" }> | null {
    const item = currentModel?.body[bodyIndex];
    if (!item || item.type !== "table") return null;
    return item;
  }

  function tableCellTextAt(cell: TableCellPosition): string {
    return (
      tableItemAt(cell.bodyIndex)?.rows[cell.row]?.cells[cell.col]?.text ?? ""
    );
  }

  function tableCellCount(bodyIndex: number): { rows: number; cols: number } {
    const table = tableItemAt(bodyIndex);
    return {
      rows: table?.rows.length ?? 0,
      cols: table?.rows[0]?.cells.length ?? 0,
    };
  }

  function supportsTableCellEditing(): boolean {
    return (
      typeof renderer.getTableCellAtPoint === "function" &&
      typeof renderer.getTableCellRect === "function"
    );
  }

  function getActiveTableCellState(): TableCellState | null {
    if (!activeTableCell) return null;
    const { rows, cols } = tableCellCount(activeTableCell.bodyIndex);
    return {
      ...activeTableCell,
      rect: renderer.getTableCellRect?.(activeTableCell) ?? null,
      text: tableCellTextAt(activeTableCell),
      rowCount: rows,
      colCount: cols,
    };
  }

  function fireActiveTableCellChange(): void {
    const state = getActiveTableCellState();
    for (const cb of tableCellCallbacks) {
      cb(state);
    }
  }

  function refreshActiveTableCell(): void {
    if (!activeTableCell) return;
    if (!tableItemAt(activeTableCell.bodyIndex)) {
      closeTableCellEditor({ commit: false });
      return;
    }
    positionTableCellEditor();
    fireActiveTableCellChange();
  }

  function positionTableCellEditor(): void {
    if (!activeTableCell || !renderer.getTableCellRect) return;
    const rect = renderer.getTableCellRect(activeTableCell);
    if (!rect) {
      tableCellEditor.style.display = "none";
      return;
    }
    const containerRect = container.getBoundingClientRect();
    tableCellEditor.style.left = `${rect.x - containerRect.left + container.scrollLeft}px`;
    tableCellEditor.style.top = `${rect.y - containerRect.top + container.scrollTop}px`;
    tableCellEditor.style.width = `${Math.max(48, rect.w)}px`;
    tableCellEditor.style.height = `${Math.max(28, rect.h)}px`;
    tableCellEditor.style.display = "block";
    fireActiveTableCellChange();
  }

  function applyTableCellEditorValue(options?: {
    flushFullState?: boolean;
    preserveSelection?: boolean;
  }): boolean {
    if (!activeTableCell) return false;
    const nextText = tableCellEditor.value;
    if (nextText === tableCellTextAt(activeTableCell)) {
      return true;
    }
    const cell = { ...activeTableCell };
    const selectionStart = tableCellEditor.selectionStart ?? nextText.length;
    const selectionEnd = tableCellEditor.selectionEnd ?? selectionStart;
    const shouldPreserveSelection = options?.preserveSelection ?? false;
    const wasFocused =
      typeof document !== "undefined" &&
      document.activeElement === tableCellEditor;
    const ok = applyAndRender(
      JSON.stringify({
        type: "setTableCellText",
        bodyIndex: cell.bodyIndex,
        row: cell.row,
        col: cell.col,
        text: nextText,
      }),
    );
    if (!ok) return false;
    if (options?.flushFullState) {
      syncProvider?.flushFullState();
    }
    tableCellEditor.value = nextText;
    if (wasFocused) {
      tableCellEditor.focus({ preventScroll: true });
      if (
        shouldPreserveSelection &&
        typeof tableCellEditor.setSelectionRange === "function"
      ) {
        tableCellEditor.setSelectionRange(selectionStart, selectionEnd);
      }
    }
    return true;
  }

  function closeTableCellEditor(options?: { commit?: boolean }): void {
    const shouldCommit = options?.commit ?? true;
    if (activeTableCell && shouldCommit) {
      applyTableCellEditorValue({ flushFullState: true });
    }
    activeTableCell = null;
    tableCellEditor.style.display = "none";
    tableCellEditor.value = "";
    fireActiveTableCellChange();
  }

  function moveTableCellSelection(delta: -1 | 1): TableCellPosition | null {
    if (!activeTableCell) return null;
    const { rows, cols } = tableCellCount(activeTableCell.bodyIndex);
    if (rows <= 0 || cols <= 0) return null;
    const linear =
      activeTableCell.row * cols + activeTableCell.col + (delta < 0 ? -1 : 1);
    if (linear < 0 || linear >= rows * cols) return null;
    return {
      bodyIndex: activeTableCell.bodyIndex,
      row: Math.floor(linear / cols),
      col: linear % cols,
    };
  }

  function openTableCellEditor(cell: TableCellPosition): void {
    if (!supportsTableCellEditing()) return;
    const table = tableItemAt(cell.bodyIndex);
    if (!table) return;
    const rect = renderer.getTableCellRect?.(cell) ?? null;
    if (!rect) return;
    activeTableCell = cell;
    clearImageSelectionOverlay();
    tableCellEditor.value = tableCellTextAt(cell);
    positionTableCellEditor();
    tableCellEditor.focus({ preventScroll: true });
    tableCellEditor.select();
    clearPendingFormat();
    fireFormattingChange();
    fireActiveTableCellChange();
  }

  tableCellEditor.addEventListener("blur", () => {
    if (suppressTableCellBlurCommit) return;
    closeTableCellEditor({ commit: true });
  });

  tableCellEditor.addEventListener("keydown", (event) => {
    if (!activeTableCell) return;
    if (event.key === "Escape") {
      event.preventDefault();
      closeTableCellEditor({ commit: false });
      renderer.focus();
      ensureCanvasInputFocus();
      return;
    }
    if (event.key === "Tab") {
      event.preventDefault();
      const next = moveTableCellSelection(event.shiftKey ? -1 : 1);
      suppressTableCellBlurCommit = true;
      closeTableCellEditor({ commit: true });
      suppressTableCellBlurCommit = false;
      if (next) {
        openTableCellEditor(next);
      } else {
        renderer.focus();
        ensureCanvasInputFocus();
      }
    }
  });

  tableCellEditor.addEventListener("input", () => {
    if (!activeTableCell) return;
    applyTableCellEditorValue({ preserveSelection: true });
  });

  tableCellEditor.addEventListener("copy", (event) => {
    const data = event.clipboardData;
    if (!data) return;
    const start = tableCellEditor.selectionStart ?? 0;
    const end = tableCellEditor.selectionEnd ?? start;
    if (start === end) return;
    data.setData("text/plain", tableCellEditor.value.slice(start, end));
    event.preventDefault();
  });

  tableCellEditor.addEventListener("cut", (event) => {
    const data = event.clipboardData;
    const start = tableCellEditor.selectionStart ?? 0;
    const end = tableCellEditor.selectionEnd ?? start;
    if (!data || start === end) return;
    data.setData("text/plain", tableCellEditor.value.slice(start, end));
    tableCellEditor.setRangeText("", start, end, "start");
    event.preventDefault();
    applyTableCellEditorValue({ preserveSelection: true });
  });

  tableCellEditor.addEventListener("paste", (event) => {
    const text = event.clipboardData?.getData("text/plain") ?? "";
    if (!text) return;
    const start = tableCellEditor.selectionStart ?? 0;
    const end = tableCellEditor.selectionEnd ?? start;
    tableCellEditor.setRangeText(text, start, end, "end");
    event.preventDefault();
    applyTableCellEditorValue({ preserveSelection: true });
  });

  function imageRunAt(
    bodyIndex: number,
    charOffset: number,
  ):
    | (RunModel & {
        inlineImage: NonNullable<RunModel["inlineImage"]>;
      })
    | null {
    const paragraph = paragraphItemAt(bodyIndex);
    if (!paragraph) return null;
    let cursor = 0;
    for (const run of paragraph.runs) {
      if (run.inlineImage) {
        if (cursor === charOffset) {
          return run as RunModel & {
            inlineImage: NonNullable<RunModel["inlineImage"]>;
          };
        }
        cursor += 1;
        continue;
      }
      if (run.hasBreak || run.hasTab) {
        cursor += (run.text?.length ?? 0) + 1;
      } else if (run.footnoteRef != null || run.endnoteRef != null) {
        cursor += String(run.footnoteRef ?? run.endnoteRef ?? "").length;
      } else {
        cursor += run.text?.length ?? 0;
      }
    }
    return null;
  }

  function getSelectedInlineImage(): SelectedInlineImageState | null {
    const sel = renderer.getSelection();
    if (!sel) return null;
    const range = normalizeSelection({
      anchor: clampPosition(sel.anchor),
      focus: clampPosition(sel.focus),
    });
    if (
      range.start.bodyIndex !== range.end.bodyIndex ||
      range.end.charOffset !== range.start.charOffset + 1
    ) {
      return null;
    }
    const run = imageRunAt(range.start.bodyIndex, range.start.charOffset);
    if (!run?.inlineImage) return null;
    return {
      bodyIndex: range.start.bodyIndex,
      charOffset: range.start.charOffset,
      imageIndex: run.inlineImage.imageIndex,
      widthPt: run.inlineImage.widthPt,
      heightPt: run.inlineImage.heightPt,
      rect:
        renderer.getInlineImageRect?.({
          bodyIndex: range.start.bodyIndex,
          charOffset: range.start.charOffset,
        }) ?? null,
    };
  }

  function ensureImageSelectionOverlay(): HTMLDivElement {
    if (imageSelectionOverlay && imageSelectionBox && imageResizeHandle) {
      return imageSelectionOverlay;
    }
    if (
      typeof getComputedStyle === "function" &&
      getComputedStyle(container).position === "static"
    ) {
      container.style.position = "relative";
    }
    const overlay = document.createElement("div");
    overlay.dataset.docviewImageSelectionLayer = "1";
    overlay.style.position = "absolute";
    overlay.style.inset = "0";
    overlay.style.pointerEvents = "none";
    overlay.style.overflow = "visible";
    overlay.style.zIndex = "26";
    overlay.style.display = "none";

    const box = document.createElement("div");
    box.dataset.docviewSelectedInlineImage = "1";
    box.style.position = "absolute";
    box.style.border = "2px solid #1a73e8";
    box.style.borderRadius = "2px";
    box.style.boxShadow = "0 0 0 1px rgba(26,115,232,0.12)";
    box.style.background = "transparent";
    box.style.pointerEvents = "none";

    const handle = document.createElement("div");
    handle.dataset.docviewImageResizeHandle = "1";
    handle.style.position = "absolute";
    handle.style.width = "10px";
    handle.style.height = "10px";
    handle.style.right = "-6px";
    handle.style.bottom = "-6px";
    handle.style.borderRadius = "999px";
    handle.style.background = "#1a73e8";
    handle.style.border = "2px solid #fff";
    handle.style.boxShadow = "0 1px 2px rgba(0,0,0,0.2)";
    handle.style.cursor = "nwse-resize";
    handle.style.pointerEvents = "auto";
    handle.addEventListener("mousedown", (event) => {
      if (destroyed) return;
      const selected = getSelectedInlineImage();
      if (!selected?.rect) return;
      event.preventDefault();
      event.stopPropagation();
      closeTableCellEditor({ commit: true });
      imageResizeSession = {
        startX: event.clientX,
        startY: event.clientY,
        startWidthPt: selected.widthPt,
        startHeightPt: selected.heightPt,
        startRect: { ...selected.rect },
        bodyIndex: selected.bodyIndex,
        charOffset: selected.charOffset,
      };
      if (imageSelectionOverlay) {
        imageSelectionOverlay.style.display = "block";
      }
    });

    box.appendChild(handle);
    overlay.appendChild(box);
    container.appendChild(overlay);
    imageSelectionOverlay = overlay;
    imageSelectionBox = box;
    imageResizeHandle = handle;
    return overlay;
  }

  function clearImageSelectionOverlay(): void {
    if (!imageSelectionOverlay || !imageSelectionBox) return;
    imageSelectionOverlay.style.display = "none";
    imageSelectionBox.style.width = "0px";
    imageSelectionBox.style.height = "0px";
  }

  function updateImageSelectionOverlay(): void {
    if (imageResizeSession && imageSelectionBox) {
      const containerRect = container.getBoundingClientRect();
      imageSelectionBox.style.left = `${imageResizeSession.startRect.x - containerRect.left + container.scrollLeft}px`;
      imageSelectionBox.style.top = `${imageResizeSession.startRect.y - containerRect.top + container.scrollTop}px`;
      imageSelectionBox.style.width = `${imageResizeSession.startRect.w}px`;
      imageSelectionBox.style.height = `${imageResizeSession.startRect.h}px`;
      return;
    }

    const selected = getSelectedInlineImage();
    if (!selected?.rect) {
      clearImageSelectionOverlay();
      return;
    }
    ensureImageSelectionOverlay();
    if (!imageSelectionBox || !imageSelectionOverlay) return;
    const containerRect = container.getBoundingClientRect();
    imageSelectionBox.style.left = `${selected.rect.x - containerRect.left + container.scrollLeft}px`;
    imageSelectionBox.style.top = `${selected.rect.y - containerRect.top + container.scrollTop}px`;
    imageSelectionBox.style.width = `${Math.max(1, selected.rect.w)}px`;
    imageSelectionBox.style.height = `${Math.max(1, selected.rect.h)}px`;
    imageSelectionOverlay.style.display = "block";
  }

  function onImageResizeMove(event: MouseEvent): void {
    if (!imageResizeSession || !imageSelectionBox) return;
    const deltaXPt = ((event.clientX - imageResizeSession.startX) * 72) / 96;
    const deltaYPt = ((event.clientY - imageResizeSession.startY) * 72) / 96;
    const nextWidthPt = Math.max(8, imageResizeSession.startWidthPt + deltaXPt);
    const nextHeightPt = Math.max(
      8,
      imageResizeSession.startHeightPt + deltaYPt,
    );
    const containerRect = container.getBoundingClientRect();
    imageSelectionBox.style.left = `${imageResizeSession.startRect.x - containerRect.left + container.scrollLeft}px`;
    imageSelectionBox.style.top = `${imageResizeSession.startRect.y - containerRect.top + container.scrollTop}px`;
    imageSelectionBox.style.width = `${(nextWidthPt * 96) / 72}px`;
    imageSelectionBox.style.height = `${(nextHeightPt * 96) / 72}px`;
  }

  function onImageResizeEnd(event: MouseEvent): void {
    if (!imageResizeSession) return;
    const deltaXPt = ((event.clientX - imageResizeSession.startX) * 72) / 96;
    const deltaYPt = ((event.clientY - imageResizeSession.startY) * 72) / 96;
    const nextWidthPt = Math.max(8, imageResizeSession.startWidthPt + deltaXPt);
    const nextHeightPt = Math.max(
      8,
      imageResizeSession.startHeightPt + deltaYPt,
    );
    imageResizeSession = null;
    resizeSelectedInlineImage(nextWidthPt, nextHeightPt);
    updateImageSelectionOverlay();
  }

  if (typeof window !== "undefined") {
    window.addEventListener("mousemove", onImageResizeMove);
    window.addEventListener("mouseup", onImageResizeEnd);
  }

  function resolveParagraphBodyIndex(bodyIndex: number): number {
    const paragraphIndices = getParagraphBodyIndices();
    if (paragraphIndices.length === 0) return Math.max(0, bodyIndex);
    if (paragraphIndices.includes(bodyIndex)) return bodyIndex;
    let nearest = paragraphIndices[0]!;
    let bestDistance = Math.abs(nearest - bodyIndex);
    for (let i = 1; i < paragraphIndices.length; i++) {
      const candidate = paragraphIndices[i]!;
      const distance = Math.abs(candidate - bodyIndex);
      if (
        distance < bestDistance ||
        (distance === bestDistance && candidate < nearest)
      ) {
        nearest = candidate;
        bestDistance = distance;
      }
    }
    return nearest;
  }

  function paragraphItemAt(
    bodyIndex: number,
  ): Extract<DocViewModel["body"][number], { type: "paragraph" }> | null {
    const resolved = resolveParagraphBodyIndex(bodyIndex);
    const item = currentModel?.body[resolved];
    if (!item || item.type !== "paragraph") return null;
    return item;
  }

  function paragraphTextAt(bodyIndex: number): string {
    const item = paragraphItemAt(bodyIndex);
    if (!item) return "";

    let text = "";
    for (const run of item.runs) {
      if (run.inlineImage) {
        text += "\uFFFC";
      } else if (run.hasBreak) {
        text += (run.text ?? "") + "\n";
      } else if (run.hasTab) {
        text += (run.text ?? "") + "\t";
      } else if (run.footnoteRef != null || run.endnoteRef != null) {
        text += String(run.footnoteRef ?? run.endnoteRef ?? "");
      } else {
        text += run.text ?? "";
      }
    }

    return text;
  }

  function paragraphCharCount(bodyIndex: number): number {
    return paragraphTextAt(bodyIndex).length;
  }

  function isWhitespaceChar(ch: string): boolean {
    return /\s/u.test(ch);
  }

  function isWordModifier(payload: NavigationPayload): boolean {
    return payload.alt || payload.ctrl;
  }

  function wordSelectionAt(pos: DocPosition): DocSelection {
    const paragraph = clampPosition(pos);
    const text = paragraphTextAt(paragraph.bodyIndex);
    const len = text.length;
    if (len === 0) {
      return {
        anchor: { bodyIndex: paragraph.bodyIndex, charOffset: 0 },
        focus: { bodyIndex: paragraph.bodyIndex, charOffset: 0 },
      };
    }

    let index = Math.max(0, Math.min(paragraph.charOffset, len - 1));
    if (paragraph.charOffset >= len) index = len - 1;

    const current = text[index] ?? "";
    const matchWhitespace = isWhitespaceChar(current);
    let start = index;
    while (
      start > 0 &&
      isWhitespaceChar(text[start - 1] ?? "") === matchWhitespace
    ) {
      start -= 1;
    }

    let end = index + 1;
    while (end < len && isWhitespaceChar(text[end] ?? "") === matchWhitespace) {
      end += 1;
    }

    return {
      anchor: { bodyIndex: paragraph.bodyIndex, charOffset: start },
      focus: { bodyIndex: paragraph.bodyIndex, charOffset: end },
    };
  }

  function paragraphSelectionAt(pos: DocPosition): DocSelection {
    const paragraph = clampPosition(pos);
    return {
      anchor: { bodyIndex: paragraph.bodyIndex, charOffset: 0 },
      focus: {
        bodyIndex: paragraph.bodyIndex,
        charOffset: paragraphCharCount(paragraph.bodyIndex),
      },
    };
  }

  function paragraphListKindAt(bodyIndex: number): ListKind | undefined {
    const item = paragraphItemAt(bodyIndex);
    const format = item?.numbering?.format;
    return format === "bullet" || format === "decimal" ? format : undefined;
  }

  function paragraphListIdAt(bodyIndex: number): number | undefined {
    const item = paragraphItemAt(bodyIndex);
    return item?.numbering?.numId;
  }

  function paragraphListLevelAt(bodyIndex: number): number | undefined {
    const item = paragraphItemAt(bodyIndex);
    return item?.numbering?.level;
  }

  function paragraphNumberingAttrsAt(
    bodyIndex: number,
  ): { kind: string; numId: number; level: number } | null {
    if (!docEdit) return null;
    const attrs = (docEdit.formattingAt(DocEdit.encodePosition(bodyIndex, 0)) ??
      {}) as Record<string, unknown>;
    const kind =
      typeof attrs.numberingKind === "string"
        ? attrs.numberingKind
        : paragraphItemAt(bodyIndex)?.numbering?.format;
    const numId =
      typeof attrs.numberingNumId === "number"
        ? attrs.numberingNumId
        : paragraphListIdAt(bodyIndex);
    const level =
      typeof attrs.numberingIlvl === "number"
        ? attrs.numberingIlvl
        : paragraphListLevelAt(bodyIndex);
    if (
      typeof kind !== "string" ||
      typeof numId !== "number" ||
      typeof level !== "number"
    ) {
      return null;
    }
    return {
      kind,
      numId: Math.max(1, Math.round(numId)),
      level: Math.max(0, Math.min(8, Math.round(level))),
    };
  }

  function clampPosition(pos: DocPosition): DocPosition {
    const resolvedBodyIndex = resolveParagraphBodyIndex(pos.bodyIndex);
    const maxOffset = paragraphCharCount(resolvedBodyIndex);
    return {
      bodyIndex: resolvedBodyIndex,
      charOffset: Math.max(0, Math.min(pos.charOffset, maxOffset)),
    };
  }

  function firstParagraphPosition(): DocPosition {
    const paragraphIndices = getParagraphBodyIndices();
    if (paragraphIndices.length === 0) return { bodyIndex: 0, charOffset: 0 };
    return { bodyIndex: paragraphIndices[0]!, charOffset: 0 };
  }

  function lastParagraphPosition(): DocPosition {
    const paragraphIndices = getParagraphBodyIndices();
    if (paragraphIndices.length === 0) return { bodyIndex: 0, charOffset: 0 };
    const bodyIndex = paragraphIndices[paragraphIndices.length - 1]!;
    return { bodyIndex, charOffset: paragraphCharCount(bodyIndex) };
  }

  function selectionOrCursor(): DocSelection {
    const sel = renderer.getSelection();
    if (sel) {
      return {
        anchor: clampPosition(sel.anchor),
        focus: clampPosition(sel.focus),
      };
    }
    const fallback = clampPosition(firstParagraphPosition());
    return { anchor: fallback, focus: fallback };
  }

  function currentAwarenessSelection(): Record<string, unknown> | null {
    if (!currentModel) return null;
    const selection = selectionOrCursor();
    return {
      anchor: {
        bodyIndex: selection.anchor.bodyIndex,
        charOffset: selection.anchor.charOffset,
      },
      focus: {
        bodyIndex: selection.focus.bodyIndex,
        charOffset: selection.focus.charOffset,
      },
    };
  }

  function ensurePresenceOverlay(): HTMLDivElement {
    if (presenceOverlay) return presenceOverlay;
    if (
      typeof getComputedStyle === "function" &&
      getComputedStyle(container).position === "static"
    ) {
      container.style.position = "relative";
    }
    const overlay = document.createElement("div");
    overlay.dataset.doceditRemotePresenceLayer = "1";
    overlay.dataset.docviewRemotePresenceLayer = "1";
    overlay.dataset.docviewRemotePresenceVisible = "0";
    overlay.dataset.docviewRemotePresenceCount = "0";
    overlay.style.position = "absolute";
    overlay.style.inset = "0";
    overlay.style.pointerEvents = "none";
    overlay.style.overflow = "visible";
    overlay.style.zIndex = "25";
    overlay.style.display = "none";
    container.appendChild(overlay);
    presenceOverlay = overlay;
    return overlay;
  }

  function clearPresenceOverlay(): void {
    if (!presenceOverlay) return;
    presenceOverlay.replaceChildren();
    presenceOverlay.dataset.docviewRemotePresenceVisible = "0";
    presenceOverlay.dataset.docviewRemotePresenceCount = "0";
    presenceOverlay.style.display = "none";
    emitEvent("docedit-presence", { visiblePeers: 0 });
  }

  function colorForPeer(senderId: string): string {
    let hash = 0;
    for (let i = 0; i < senderId.length; i++) {
      hash = (hash * 31 + senderId.charCodeAt(i)) | 0;
    }
    const hue = Math.abs(hash) % 360;
    return `hsl(${hue} 75% 48%)`;
  }

  function parsePeerSelection(
    state: Record<string, unknown>,
  ): DocSelection | null {
    const anchor = state.anchor as Record<string, unknown> | undefined;
    const focus = state.focus as Record<string, unknown> | undefined;
    if (
      !anchor ||
      !focus ||
      typeof anchor.bodyIndex !== "number" ||
      typeof anchor.charOffset !== "number" ||
      typeof focus.bodyIndex !== "number" ||
      typeof focus.charOffset !== "number"
    ) {
      return null;
    }
    return {
      anchor: clampPosition({
        bodyIndex: anchor.bodyIndex,
        charOffset: anchor.charOffset,
      }),
      focus: clampPosition({
        bodyIndex: focus.bodyIndex,
        charOffset: focus.charOffset,
      }),
    };
  }

  function renderRemotePresence(): void {
    presenceRaf = 0;
    if (destroyed || !currentModel) return;
    const peers = remotePeers
      .map((peer) => ({
        senderId: peer.senderId,
        selection: parsePeerSelection(peer.state),
      }))
      .filter(
        (
          peer,
        ): peer is {
          senderId: string;
          selection: DocSelection;
        } => !!peer.selection,
      );

    if (peers.length === 0) {
      clearPresenceOverlay();
      return;
    }

    const overlay = ensurePresenceOverlay();
    overlay.replaceChildren();
    const containerRect = container.getBoundingClientRect();
    let visiblePeers = 0;

    for (const peer of peers) {
      const selection = peer.selection;
      const color = colorForPeer(peer.senderId);
      const peerRoot = document.createElement("div");
      peerRoot.dataset.doceditRemotePresence = "1";
      peerRoot.dataset.docviewRemotePresencePeer = "1";
      peerRoot.dataset.peerId = peer.senderId;
      peerRoot.style.position = "absolute";
      peerRoot.style.inset = "0";
      let peerVisible = false;

      const selectionRects = renderer.getSelectionRects(selection);
      for (const rect of selectionRects) {
        if (rect.w <= 0 || rect.h <= 0) continue;
        const block = document.createElement("div");
        block.dataset.docviewRemoteSelection = "1";
        block.style.position = "absolute";
        block.style.left = `${rect.x - containerRect.left + container.scrollLeft}px`;
        block.style.top = `${rect.y - containerRect.top + container.scrollTop}px`;
        block.style.width = `${Math.max(1, rect.w)}px`;
        block.style.height = `${Math.max(1, rect.h)}px`;
        block.style.background = color;
        block.style.opacity = "0.16";
        block.style.borderRadius = "2px";
        peerRoot.appendChild(block);
        peerVisible = true;
      }

      const caretRect = renderer.getCursorRectForPosition(selection.focus);
      if (caretRect) {
        const caret = document.createElement("div");
        caret.dataset.docviewRemoteCursor = "1";
        caret.style.position = "absolute";
        caret.style.left = `${caretRect.x - containerRect.left + container.scrollLeft}px`;
        caret.style.top = `${caretRect.y - containerRect.top + container.scrollTop}px`;
        caret.style.width = `${Math.max(2, Math.ceil(caretRect.w || 1))}px`;
        caret.style.height = `${Math.max(12, Math.round(caretRect.h))}px`;
        caret.style.background = color;
        caret.style.borderRadius = "1px";
        peerRoot.appendChild(caret);
        peerVisible = true;

        const label = document.createElement("div");
        label.dataset.docviewRemoteCursorLabel = "1";
        label.textContent = peer.senderId.slice(0, 6);
        label.style.position = "absolute";
        label.style.left = `${caretRect.x - containerRect.left + container.scrollLeft}px`;
        label.style.top = `${caretRect.y - containerRect.top + container.scrollTop - 18}px`;
        label.style.padding = "1px 4px";
        label.style.background = color;
        label.style.color = "#fff";
        label.style.fontSize = "10px";
        label.style.lineHeight = "14px";
        label.style.borderRadius = "4px";
        label.style.whiteSpace = "nowrap";
        peerRoot.appendChild(label);
      }

      if (!peerVisible) continue;
      overlay.appendChild(peerRoot);
      visiblePeers += 1;
    }

    overlay.dataset.docviewRemotePresenceVisible = visiblePeers > 0 ? "1" : "0";
    overlay.dataset.docviewRemotePresenceCount = String(visiblePeers);
    overlay.style.display = visiblePeers > 0 ? "block" : "none";
    emitEvent("docedit-presence", { visiblePeers });
  }

  function scheduleRemotePresenceRender(): void {
    if (remotePeers.length === 0) {
      if (presenceRaf !== 0 && typeof cancelAnimationFrame === "function") {
        cancelAnimationFrame(presenceRaf);
      }
      presenceRaf = 0;
      clearPresenceOverlay();
      return;
    }
    if (presenceRaf !== 0) return;
    if (typeof requestAnimationFrame === "function") {
      presenceRaf = requestAnimationFrame(renderRemotePresence);
      return;
    }
    presenceRaf = -1;
    queueMicrotask(renderRemotePresence);
  }

  function updateLocalAwareness(): void {
    syncProvider?.setLocalAwareness(currentAwarenessSelection());
  }

  const onPresenceViewportChange = (): void => {
    scheduleRemotePresenceRender();
    positionTableCellEditor();
    updateImageSelectionOverlay();
  };

  if (typeof container.addEventListener === "function") {
    container.addEventListener("scroll", onPresenceViewportChange, {
      passive: true,
    });
  }
  if (typeof window !== "undefined") {
    window.addEventListener("resize", onPresenceViewportChange);
  }

  function moveByCharacter(base: DocPosition, delta: -1 | 1): DocPosition {
    const pos = clampPosition(base);
    const paragraphIndices = getParagraphBodyIndices();
    if (paragraphIndices.length === 0) return pos;

    const currentParagraphIdx = paragraphIndices.indexOf(pos.bodyIndex);
    if (currentParagraphIdx < 0) return pos;

    if (delta < 0) {
      if (pos.charOffset > 0) {
        return { bodyIndex: pos.bodyIndex, charOffset: pos.charOffset - 1 };
      }
      if (currentParagraphIdx === 0) return pos;
      const prevBody = paragraphIndices[currentParagraphIdx - 1]!;
      return { bodyIndex: prevBody, charOffset: paragraphCharCount(prevBody) };
    }

    const currentMax = paragraphCharCount(pos.bodyIndex);
    if (pos.charOffset < currentMax) {
      return { bodyIndex: pos.bodyIndex, charOffset: pos.charOffset + 1 };
    }
    if (currentParagraphIdx === paragraphIndices.length - 1) return pos;
    const nextBody = paragraphIndices[currentParagraphIdx + 1]!;
    return { bodyIndex: nextBody, charOffset: 0 };
  }

  function moveVerticalByHitTest(
    base: DocPosition,
    direction: "up" | "down" | "pageUp" | "pageDown",
  ): DocPosition {
    const start = clampPosition(base);
    const rect = renderer.getCursorRect();
    if (!rect) return start;

    const x = rect.x + Math.max(rect.w, 1) / 2;
    const y = rect.y + rect.h / 2;
    const lineStep = Math.max(rect.h, 16);
    const pageStep = Math.max(container.clientHeight - lineStep * 2, lineStep);
    const up = direction === "up" || direction === "pageUp";
    const primaryDelta =
      direction === "up" || direction === "down"
        ? lineStep * 1.25
        : pageStep * 0.9;
    const attempts = [primaryDelta, primaryDelta * 2];

    for (const step of attempts) {
      const targetY = y + (up ? -step : step);
      const hit = renderer.hitTest(x, targetY);
      if (!hit) continue;
      const candidate = clampPosition({
        bodyIndex: hit.bodyIndex,
        charOffset: hit.charOffset,
      });
      if (!samePosition(candidate, start)) return candidate;
    }

    return up ? firstParagraphPosition() : lastParagraphPosition();
  }

  function moveForNavigation(
    base: DocPosition,
    key: NavigationKey,
    payload?: NavigationPayload,
  ): DocPosition {
    const start = clampPosition(base);
    switch (key) {
      case "ArrowLeft":
        if (payload && isWordModifier(payload)) return moveByWord(start, -1);
        return moveByCharacter(start, -1);
      case "ArrowRight":
        if (payload && isWordModifier(payload)) return moveByWord(start, 1);
        return moveByCharacter(start, 1);
      case "Home":
        return { bodyIndex: start.bodyIndex, charOffset: 0 };
      case "End":
        return {
          bodyIndex: start.bodyIndex,
          charOffset: paragraphCharCount(start.bodyIndex),
        };
      case "ArrowUp":
        return moveVerticalByHitTest(start, "up");
      case "ArrowDown":
        return moveVerticalByHitTest(start, "down");
      case "PageUp":
        return moveVerticalByHitTest(start, "pageUp");
      case "PageDown":
        return moveVerticalByHitTest(start, "pageDown");
    }
  }

  function collapsePositionForNavigation(
    range: OrderedSelectionRange,
    key: NavigationKey,
  ): DocPosition {
    switch (key) {
      case "ArrowLeft":
      case "ArrowUp":
      case "Home":
      case "PageUp":
        return range.start;
      case "ArrowRight":
      case "ArrowDown":
      case "End":
      case "PageDown":
        return range.end;
    }
  }

  function selectedPlainText(): string {
    if (!currentModel) return "";
    const sel = renderer.getSelection();
    if (!sel) return "";

    const range = normalizeSelection({
      anchor: clampPosition(sel.anchor),
      focus: clampPosition(sel.focus),
    });
    if (range.collapsed) return "";

    const paragraphIndices = getParagraphBodyIndices().filter((idx) => {
      return idx >= range.start.bodyIndex && idx <= range.end.bodyIndex;
    });
    if (paragraphIndices.length === 0) return "";

    const chunks: string[] = [];
    for (const bodyIndex of paragraphIndices) {
      const paragraphText = paragraphTextAt(bodyIndex);
      let start = 0;
      let end = paragraphText.length;
      if (bodyIndex === range.start.bodyIndex) start = range.start.charOffset;
      if (bodyIndex === range.end.bodyIndex) end = range.end.charOffset;
      start = Math.max(0, Math.min(start, paragraphText.length));
      end = Math.max(start, Math.min(end, paragraphText.length));
      chunks.push(paragraphText.slice(start, end));
    }

    return chunks.join("\n");
  }

  function escapeHtml(text: string): string {
    return text
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;");
  }

  function escapeHtmlAttr(text: string): string {
    return escapeHtml(text).replaceAll('"', "&quot;");
  }

  function highlightCssColor(name: string): string {
    const map: Record<string, string> = {
      yellow: "#ffff00",
      green: "#00ff00",
      cyan: "#00ffff",
      magenta: "#ff00ff",
      red: "#ff0000",
      blue: "#0000ff",
      darkBlue: "#000080",
      darkCyan: "#008080",
      darkGreen: "#008000",
      darkMagenta: "#800080",
      darkRed: "#800000",
      darkYellow: "#808000",
      lightGray: "#c0c0c0",
      darkGray: "#808080",
      black: "#000000",
    };
    return map[name] ?? "#ffff00";
  }

  function styleAttr(style: string[]): string {
    return style.length > 0
      ? ` style="${escapeHtmlAttr(style.join("; "))}"`
      : "";
  }

  type RichPasteInlineState = {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    fontFamily?: string;
    fontSizePt?: number;
    color?: string;
    highlight?: string;
    hyperlink?: string;
  };

  interface RichPasteSpan {
    text: string;
    attrs: RichPasteInlineState;
  }

  interface RichPasteBlock {
    text: string;
    spans: RichPasteSpan[];
    paragraph: ParagraphStylePatch;
    listInstanceId?: number;
  }

  interface RichPasteDocument {
    blocks: RichPasteBlock[];
    plainText: string;
  }

  function tokenizeRichHtml(html: string): Array<
    | {
        type: "text";
        text: string;
      }
    | {
        type: "start";
        name: string;
        attrs: Record<string, string>;
      }
    | {
        type: "end";
        name: string;
      }
  > {
    const tokens: Array<
      | {
          type: "text";
          text: string;
        }
      | {
          type: "start";
          name: string;
          attrs: Record<string, string>;
        }
      | {
          type: "end";
          name: string;
        }
    > = [];
    const normalized = html.replaceAll(/\r\n?/g, "\n");
    const voidTags = new Set(["br", "img", "meta", "hr", "input", "link"]);
    let index = 0;
    while (index < normalized.length) {
      const lt = normalized.indexOf("<", index);
      if (lt < 0) {
        const text = normalized.slice(index);
        if (text) tokens.push({ type: "text", text });
        break;
      }
      if (lt > index) {
        tokens.push({ type: "text", text: normalized.slice(index, lt) });
      }
      if (normalized.startsWith("<!--", lt)) {
        const commentEnd = normalized.indexOf("-->", lt + 4);
        index = commentEnd >= 0 ? commentEnd + 3 : normalized.length;
        continue;
      }
      const gt = normalized.indexOf(">", lt + 1);
      if (gt < 0) break;
      const rawTag = normalized.slice(lt + 1, gt).trim();
      index = gt + 1;
      if (!rawTag || rawTag.startsWith("!")) continue;
      if (rawTag.startsWith("/")) {
        const name = rawTag.slice(1).trim().split(/\s+/u, 1)[0]?.toLowerCase();
        if (name) tokens.push({ type: "end", name });
        continue;
      }
      const selfClosing = rawTag.endsWith("/");
      const nameMatch = /^([^\s/>]+)/u.exec(rawTag);
      const name = nameMatch?.[1]?.toLowerCase();
      if (!name) continue;
      const attrs: Record<string, string> = {};
      const nameToken = nameMatch?.[0];
      if (!nameToken) continue;
      const attrSource = rawTag.slice(nameToken.length).replace(/\/\s*$/u, "");
      const attrRegex =
        /([^\s=/>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/gu;
      for (const match of attrSource.matchAll(attrRegex)) {
        const attrName = match[1]?.toLowerCase();
        if (!attrName) continue;
        attrs[attrName] = match[2] ?? match[3] ?? match[4] ?? "";
      }
      tokens.push({ type: "start", name, attrs });
      if (selfClosing || voidTags.has(name)) {
        tokens.push({ type: "end", name });
      }
    }
    return tokens;
  }

  function decodeRichHtmlEntities(text: string): string {
    return text.replaceAll(
      /&(#x?[0-9a-f]+|nbsp|amp|lt|gt|quot|apos);/giu,
      (entity, code: string) => {
        const lower = code.toLowerCase();
        if (lower === "nbsp") return " ";
        if (lower === "amp") return "&";
        if (lower === "lt") return "<";
        if (lower === "gt") return ">";
        if (lower === "quot") return '"';
        if (lower === "apos") return "'";
        const value = lower.startsWith("#x")
          ? Number.parseInt(lower.slice(2), 16)
          : lower.startsWith("#")
            ? Number.parseInt(lower.slice(1), 10)
            : Number.NaN;
        return Number.isFinite(value) ? String.fromCodePoint(value) : entity;
      },
    );
  }

  function parseStyleMap(style: string | undefined): Record<string, string> {
    const map: Record<string, string> = {};
    if (!style) return map;
    for (const entry of style.split(";")) {
      const colon = entry.indexOf(":");
      if (colon <= 0) continue;
      const key = entry.slice(0, colon).trim().toLowerCase();
      const value = entry.slice(colon + 1).trim();
      if (!key || !value) continue;
      map[key] = value;
    }
    return map;
  }

  function parseCssColor(value: string | undefined): string | undefined {
    if (!value) return undefined;
    const trimmed = value.trim().toLowerCase();
    if (!trimmed || trimmed === "transparent" || trimmed === "inherit") {
      return undefined;
    }
    const named: Record<string, string> = {
      black: "000000",
      white: "FFFFFF",
      red: "FF0000",
      blue: "0000FF",
      green: "008000",
      yellow: "FFFF00",
      cyan: "00FFFF",
      aqua: "00FFFF",
      magenta: "FF00FF",
      fuchsia: "FF00FF",
      gray: "808080",
      grey: "808080",
      lightgray: "C0C0C0",
      lightgrey: "C0C0C0",
      orange: "FFA500",
    };
    if (trimmed in named) return named[trimmed];
    const hex = trimmed.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/iu);
    if (hex) {
      const raw = hex[1]!.toUpperCase();
      return raw.length === 3
        ? raw
            .split("")
            .map((ch) => `${ch}${ch}`)
            .join("")
        : raw;
    }
    const rgb = trimmed.match(
      /^rgba?\(\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})\s*,\s*([0-9]{1,3})/u,
    );
    if (!rgb) return undefined;
    return [rgb[1], rgb[2], rgb[3]]
      .map((part) =>
        Math.max(0, Math.min(255, Number.parseInt(part ?? "0", 10)))
          .toString(16)
          .padStart(2, "0")
          .toUpperCase(),
      )
      .join("");
  }

  function parseHighlightName(value: string | undefined): string | undefined {
    const hex = parseCssColor(value);
    if (!hex) return undefined;
    const direct: Record<string, string> = {
      FFFF00: "yellow",
      "00FF00": "green",
      "00FFFF": "cyan",
      FF00FF: "magenta",
      FF0000: "red",
      "0000FF": "blue",
      "000080": "darkBlue",
      "008080": "darkCyan",
      "008000": "darkGreen",
      "800080": "darkMagenta",
      "800000": "darkRed",
      "808000": "darkYellow",
      C0C0C0: "lightGray",
      "808080": "darkGray",
      "000000": "black",
    };
    return direct[hex];
  }

  function parseCssFontSizePt(value: string | undefined): number | undefined {
    if (!value) return undefined;
    const match = value.trim().match(/^(-?\d+(?:\.\d+)?)(pt|px)?$/iu);
    if (!match) return undefined;
    const amount = Number.parseFloat(match[1] ?? "");
    if (!Number.isFinite(amount) || amount <= 0) return undefined;
    const unit = (match[2] ?? "px").toLowerCase();
    const pt = unit === "pt" ? amount : (amount * 72) / 96;
    return Math.round(pt * 100) / 100;
  }

  function parseFontFamily(value: string | undefined): string | undefined {
    if (!value) return undefined;
    const first = value.split(",")[0]?.trim();
    if (!first) return undefined;
    return first.replace(/^["']|["']$/gu, "");
  }

  function extractInlineStateFromTag(
    name: string,
    attrs: Record<string, string>,
  ): RichPasteInlineState {
    const style = parseStyleMap(attrs.style);
    const next: RichPasteInlineState = {};

    if (name === "strong" || name === "b") next.bold = true;
    if (name === "em" || name === "i") next.italic = true;
    if (name === "u") next.underline = true;
    if (name === "s" || name === "strike" || name === "del") next.strike = true;
    if (name === "a" && attrs.href) next.hyperlink = attrs.href;
    if (name === "font") {
      if (attrs.color) next.color = parseCssColor(attrs.color);
      if (attrs.face) next.fontFamily = parseFontFamily(attrs.face);
    }

    const fontWeight = style["font-weight"]?.toLowerCase();
    if (
      fontWeight === "bold" ||
      fontWeight === "bolder" ||
      (fontWeight != null && Number.parseInt(fontWeight, 10) >= 600)
    ) {
      next.bold = true;
    }
    if (style["font-style"]?.toLowerCase() === "italic") next.italic = true;
    const textDecoration =
      style["text-decoration"]?.toLowerCase() ??
      style["text-decoration-line"]?.toLowerCase() ??
      "";
    if (textDecoration.includes("underline")) next.underline = true;
    if (textDecoration.includes("line-through")) next.strike = true;
    const fontFamily = parseFontFamily(style["font-family"]);
    if (fontFamily) next.fontFamily = fontFamily;
    const fontSizePt = parseCssFontSizePt(style["font-size"]);
    if (typeof fontSizePt === "number") next.fontSizePt = fontSizePt;
    const color = parseCssColor(style.color);
    if (color) next.color = color;
    const highlight = parseHighlightName(
      style["background-color"] ?? style.background,
    );
    if (highlight) next.highlight = highlight;
    return next;
  }

  function extractParagraphAttrsFromTag(
    name: string,
    attrs: Record<string, string>,
    listStack: Array<{ kind: ListKind; listInstanceId: number }>,
  ): { paragraph: ParagraphStylePatch; listInstanceId?: number } {
    const style = parseStyleMap(attrs.style);
    const paragraph: ParagraphStylePatch = {};
    const headingMatch = name.match(/^h([1-6])$/u);
    if (headingMatch) {
      paragraph.headingLevel = Number.parseInt(headingMatch[1] ?? "1", 10);
    }
    const alignment = (style["text-align"] ?? attrs.align ?? "")
      .trim()
      .toLowerCase();
    if (
      alignment === "left" ||
      alignment === "center" ||
      alignment === "right" ||
      alignment === "justify"
    ) {
      paragraph.alignment = alignment;
    }
    if (name === "li") {
      const currentList = listStack[listStack.length - 1];
      if (currentList) {
        paragraph.numberingKind = currentList.kind;
        paragraph.numberingIlvl = Math.max(0, listStack.length - 1);
        return { paragraph, listInstanceId: currentList.listInstanceId };
      }
    }
    return { paragraph };
  }

  function compactInlinePasteAttrs(
    state: RichPasteInlineState,
  ): TextStylePatch {
    const attrs: TextStylePatch = {};
    if (state.bold) attrs.bold = true;
    if (state.italic) attrs.italic = true;
    if (state.underline) attrs.underline = true;
    if (state.strike) attrs.strike = true;
    if (state.fontFamily) attrs.fontFamily = state.fontFamily;
    if (typeof state.fontSizePt === "number")
      attrs.fontSizePt = state.fontSizePt;
    if (state.color) attrs.color = state.color;
    if (state.highlight) attrs.highlight = state.highlight;
    if (state.hyperlink) attrs.hyperlink = state.hyperlink;
    return attrs;
  }

  function sameInlinePasteAttrs(
    left: RichPasteInlineState,
    right: RichPasteInlineState,
  ): boolean {
    return (
      left.bold === right.bold &&
      left.italic === right.italic &&
      left.underline === right.underline &&
      left.strike === right.strike &&
      left.fontFamily === right.fontFamily &&
      left.fontSizePt === right.fontSizePt &&
      left.color === right.color &&
      left.highlight === right.highlight &&
      left.hyperlink === right.hyperlink
    );
  }

  function isBlockPasteTag(name: string): boolean {
    return (
      name === "p" ||
      name === "li" ||
      name === "blockquote" ||
      /^h[1-6]$/u.test(name)
    );
  }

  function parseRichClipboardHtml(html: string): RichPasteDocument | null {
    if (!html || !html.includes("<")) return null;
    const tokens = tokenizeRichHtml(html);
    if (tokens.length === 0) return null;

    const inlineStack: RichPasteInlineState[] = [{}];
    const listStack: Array<{ kind: ListKind; listInstanceId: number }> = [];
    const blockTagStack: string[] = [];
    const blocks: RichPasteBlock[] = [];
    let currentBlock: RichPasteBlock | null = null;
    let nextListInstanceId = 1;

    const ensureBlock = (
      paragraph?: ParagraphStylePatch,
      listInstanceId?: number,
    ): RichPasteBlock => {
      if (!currentBlock) {
        currentBlock = {
          text: "",
          spans: [],
          paragraph: paragraph ? { ...paragraph } : {},
          listInstanceId,
        };
      } else {
        if (paragraph && Object.keys(currentBlock.paragraph).length === 0) {
          currentBlock.paragraph = { ...paragraph };
        }
        if (listInstanceId != null && currentBlock.listInstanceId == null) {
          currentBlock.listInstanceId = listInstanceId;
        }
      }
      return currentBlock;
    };

    const pushText = (value: string): void => {
      const text = decodeRichHtmlEntities(value);
      if (!text) return;
      if (/^\s+$/u.test(text) && !currentBlock) return;
      const block = ensureBlock();
      block.text += text;
      const attrs = { ...inlineStack[inlineStack.length - 1]! };
      const lastSpan = block.spans[block.spans.length - 1];
      if (lastSpan && sameInlinePasteAttrs(lastSpan.attrs, attrs)) {
        lastSpan.text += text;
      } else {
        block.spans.push({ text, attrs });
      }
    };

    const finishBlock = (): void => {
      if (!currentBlock) return;
      if (
        currentBlock.text.length > 0 ||
        Object.keys(currentBlock.paragraph).length > 0 ||
        currentBlock.listInstanceId != null
      ) {
        blocks.push(currentBlock);
      }
      currentBlock = null;
    };

    for (const token of tokens) {
      if (token.type === "text") {
        pushText(token.text);
        continue;
      }
      if (token.type === "start") {
        if (token.name === "ul" || token.name === "ol") {
          listStack.push({
            kind: token.name === "ol" ? "decimal" : "bullet",
            listInstanceId: nextListInstanceId++,
          });
        }
        if (token.name === "br") {
          finishBlock();
          continue;
        }
        if (isBlockPasteTag(token.name)) {
          finishBlock();
          const blockInfo = extractParagraphAttrsFromTag(
            token.name,
            token.attrs,
            listStack,
          );
          currentBlock = {
            text: "",
            spans: [],
            paragraph: blockInfo.paragraph,
            listInstanceId: blockInfo.listInstanceId,
          };
          blockTagStack.push(token.name);
        }
        const currentInline = inlineStack[inlineStack.length - 1] ?? {};
        inlineStack.push({
          ...currentInline,
          ...extractInlineStateFromTag(token.name, token.attrs),
        });
        continue;
      }

      const ended = token.name;
      if (ended === "br") {
        continue;
      }
      if (ended === "ul" || ended === "ol") {
        listStack.pop();
      }
      if (inlineStack.length > 1) {
        inlineStack.pop();
      }
      if (
        blockTagStack.length > 0 &&
        blockTagStack[blockTagStack.length - 1] === ended
      ) {
        blockTagStack.pop();
        finishBlock();
      }
    }

    finishBlock();
    const normalizedBlocks = blocks.filter((block) => {
      return (
        block.text.length > 0 ||
        Object.keys(block.paragraph).length > 0 ||
        block.listInstanceId != null
      );
    });
    if (normalizedBlocks.length === 0) return null;

    return {
      plainText: normalizedBlocks.map((block) => block.text).join("\n"),
      blocks: normalizedBlocks.map((block) => ({
        ...block,
        spans: block.spans
          .filter((span) => span.text.length > 0)
          .map((span) => ({ text: span.text, attrs: { ...span.attrs } })),
        paragraph: { ...block.paragraph },
      })),
    };
  }

  function richPasteEndPosition(
    start: DocPosition,
    blocks: RichPasteBlock[],
  ): DocPosition {
    if (blocks.length === 0) return { ...start };
    if (blocks.length === 1) {
      return {
        bodyIndex: start.bodyIndex,
        charOffset: start.charOffset + blocks[0]!.text.length,
      };
    }
    const last = blocks[blocks.length - 1]!;
    return {
      bodyIndex: start.bodyIndex + blocks.length - 1,
      charOffset: last.text.length,
    };
  }

  function applyRichPasteDocument(
    parsed: RichPasteDocument,
    range: OrderedSelectionRange,
    positions: { anchor: string; focus: string },
  ): boolean {
    if (!docEdit || parsed.blocks.length === 0) return false;
    const intents: string[] = [
      JSON.stringify({
        type: "insertFromPaste",
        data: parsed.plainText,
        anchor: positions.anchor,
        focus: positions.focus,
      }),
    ];
    const listIdMap = new Map<number, number>();
    let nextNumId = nextSyntheticListId();

    parsed.blocks.forEach((block, blockIndex) => {
      const bodyIndex = range.start.bodyIndex + blockIndex;
      const paragraphAttrs = { ...block.paragraph };
      if (block.listInstanceId != null) {
        const existing = listIdMap.get(block.listInstanceId);
        const numberingNumId = existing ?? nextNumId++;
        listIdMap.set(block.listInstanceId, numberingNumId);
        paragraphAttrs.numberingNumId = numberingNumId;
      }
      const intentParagraphAttrs = toIntentParagraphAttrs(paragraphAttrs);
      if (Object.keys(intentParagraphAttrs).length > 0) {
        intents.push(
          JSON.stringify({
            type: "setParagraphAttrs",
            anchor: DocEdit.encodePosition(bodyIndex, 0),
            focus: DocEdit.encodePosition(bodyIndex, block.text.length),
            attrs: intentParagraphAttrs,
          }),
        );
      }

      let offset = blockIndex === 0 ? range.start.charOffset : 0;
      for (const span of block.spans) {
        if (!span.text) continue;
        const textAttrs = toIntentTextAttrs(
          compactInlinePasteAttrs(span.attrs),
        );
        const spanStart = offset;
        const spanEnd = spanStart + span.text.length;
        offset = spanEnd;
        if (Object.keys(textAttrs).length === 0) continue;
        intents.push(
          JSON.stringify({
            type: "setTextAttrs",
            anchor: DocEdit.encodePosition(bodyIndex, spanStart),
            focus: DocEdit.encodePosition(bodyIndex, spanEnd),
            attrs: textAttrs,
          }),
        );
      }
    });

    const end = richPasteEndPosition(range.start, parsed.blocks);
    return applyBatchAndRender(intents, end.bodyIndex, end.charOffset);
  }

  function serializeRunFragment(run: RunModel, text: string): string {
    const effectiveText = text
      .replaceAll("\t", "&#9;")
      .replaceAll("\n", "<br>");
    let content = effectiveText.length > 0 ? escapeHtml(effectiveText) : "";
    content = content
      .replaceAll("&amp;#9;", "&#9;")
      .replaceAll("&lt;br&gt;", "<br>");

    if (!content) {
      if (run.hasBreak) content = "<br>";
      else if (run.hasTab) content = "&#9;";
      else return "";
    }

    const styles: string[] = [];
    if (run.bold) styles.push("font-weight: 700");
    if (run.italic) styles.push("font-style: italic");
    const textDecorations: string[] = [];
    if (run.underline) textDecorations.push("underline");
    if (run.strikethrough) textDecorations.push("line-through");
    if (textDecorations.length > 0) {
      styles.push(`text-decoration: ${textDecorations.join(" ")}`);
    }
    if (run.fontFamily) {
      styles.push(`font-family: '${run.fontFamily.replaceAll("'", "\\'")}'`);
    }
    if (typeof run.fontSizePt === "number") {
      styles.push(`font-size: ${run.fontSizePt}pt`);
    }
    if (run.color) {
      styles.push(`color: #${run.color}`);
    }
    if (run.highlight) {
      styles.push(`background-color: ${highlightCssColor(run.highlight)}`);
    }

    if (styles.length > 0) {
      content = `<span${styleAttr(styles)}>${content}</span>`;
    }

    if (run.hyperlink) {
      content = `<a href="${escapeHtmlAttr(run.hyperlink)}">${content}</a>`;
    }

    return content;
  }

  function selectedHtml(): string {
    if (!currentModel) return "";
    const sel = renderer.getSelection();
    if (!sel) return "";

    const range = normalizeSelection({
      anchor: clampPosition(sel.anchor),
      focus: clampPosition(sel.focus),
    });
    if (range.collapsed) return "";

    const blocks: string[] = [];
    const paragraphIndices = getParagraphBodyIndices().filter((idx) => {
      return idx >= range.start.bodyIndex && idx <= range.end.bodyIndex;
    });

    for (const bodyIndex of paragraphIndices) {
      const paragraph = paragraphItemAt(bodyIndex);
      if (!paragraph) continue;
      const startOffset =
        bodyIndex === range.start.bodyIndex ? range.start.charOffset : 0;
      const endOffset =
        bodyIndex === range.end.bodyIndex
          ? range.end.charOffset
          : paragraphCharCount(bodyIndex);
      let cursor = 0;
      let content = "";

      for (const run of paragraph.runs) {
        const rawText = run.hasBreak
          ? `${run.text ?? ""}\n`
          : run.hasTab
            ? `${run.text ?? ""}\t`
            : run.inlineImage
              ? "\uFFFC"
              : run.footnoteRef != null || run.endnoteRef != null
                ? String(run.footnoteRef ?? run.endnoteRef ?? "")
                : (run.text ?? "");
        const runStart = cursor;
        const runEnd = cursor + rawText.length;
        cursor = runEnd;
        if (runEnd <= startOffset || runStart >= endOffset) continue;
        const sliceStart = Math.max(0, startOffset - runStart);
        const sliceEnd = Math.max(
          sliceStart,
          Math.min(rawText.length, endOffset - runStart),
        );
        const fragment = rawText.slice(sliceStart, sliceEnd);
        content += serializeRunFragment(run, fragment);
      }

      const paragraphStyles: string[] = [];
      if (paragraph.alignment) {
        paragraphStyles.push(`text-align: ${paragraph.alignment}`);
      }
      if (typeof paragraph.spacingBeforePt === "number") {
        paragraphStyles.push(`margin-top: ${paragraph.spacingBeforePt}pt`);
      }
      if (typeof paragraph.spacingAfterPt === "number") {
        paragraphStyles.push(`margin-bottom: ${paragraph.spacingAfterPt}pt`);
      }
      if (typeof paragraph.indents?.leftPt === "number") {
        paragraphStyles.push(`margin-left: ${paragraph.indents.leftPt}pt`);
      }
      if (typeof paragraph.indents?.firstLinePt === "number") {
        paragraphStyles.push(`text-indent: ${paragraph.indents.firstLinePt}pt`);
      }
      if (typeof paragraph.lineSpacing?.value === "number") {
        if (paragraph.lineSpacing.rule === "auto") {
          paragraphStyles.push(`line-height: ${paragraph.lineSpacing.value}`);
        } else {
          paragraphStyles.push(`line-height: ${paragraph.lineSpacing.value}pt`);
        }
      }

      const tag =
        typeof paragraph.headingLevel === "number" &&
        paragraph.headingLevel >= 1 &&
        paragraph.headingLevel <= 6
          ? `h${paragraph.headingLevel}`
          : "p";
      blocks.push(
        `<${tag}${styleAttr(paragraphStyles)}>${content || "<br>"}</${tag}>`,
      );
    }

    return blocks.join("");
  }

  function selectAllInCanvas(): void {
    const start = clampPosition(firstParagraphPosition());
    const end = clampPosition(lastParagraphPosition());
    if (samePosition(start, end)) {
      renderer.setCursor(start);
      return;
    }
    renderer.setSelection({
      anchor: start,
      focus: end,
    });
  }

  function selectAllInNative(): void {
    const surface = renderer.getInputElement();
    const sel = window.getSelection();
    if (!surface || !sel) return;
    const range = document.createRange();
    range.selectNodeContents(surface);
    sel.removeAllRanges();
    sel.addRange(range);
  }

  function moveByWord(base: DocPosition, delta: -1 | 1): DocPosition {
    const pos = clampPosition(base);
    const paragraphIndices = getParagraphBodyIndices();
    if (paragraphIndices.length === 0) return pos;

    const currentParagraphIdx = paragraphIndices.indexOf(pos.bodyIndex);
    if (currentParagraphIdx < 0) return pos;

    const text = paragraphTextAt(pos.bodyIndex);
    const len = text.length;

    if (delta < 0) {
      if (pos.charOffset === 0) {
        if (currentParagraphIdx === 0) return pos;
        const prevBody = paragraphIndices[currentParagraphIdx - 1]!;
        return {
          bodyIndex: prevBody,
          charOffset: paragraphCharCount(prevBody),
        };
      }
      let index = Math.max(0, Math.min(pos.charOffset, len));
      while (index > 0 && isWhitespaceChar(text[index - 1] ?? "")) {
        index -= 1;
      }
      while (index > 0 && !isWhitespaceChar(text[index - 1] ?? "")) {
        index -= 1;
      }
      return { bodyIndex: pos.bodyIndex, charOffset: index };
    }

    if (pos.charOffset >= len) {
      if (currentParagraphIdx === paragraphIndices.length - 1) return pos;
      const nextBody = paragraphIndices[currentParagraphIdx + 1]!;
      return { bodyIndex: nextBody, charOffset: 0 };
    }

    let index = Math.max(0, Math.min(pos.charOffset, len));
    if (!isWhitespaceChar(text[index] ?? "")) {
      while (index < len && !isWhitespaceChar(text[index] ?? "")) {
        index += 1;
      }
    }
    while (index < len && isWhitespaceChar(text[index] ?? "")) {
      index += 1;
    }
    return { bodyIndex: pos.bodyIndex, charOffset: index };
  }

  function deleteWordStart(base: DocPosition): DocPosition {
    const start = moveByWord(base, -1);
    if (start.bodyIndex !== base.bodyIndex) return start;
    return start;
  }

  function deleteWordBackwardEnd(base: DocPosition): DocPosition {
    const pos = clampPosition(base);
    const text = paragraphTextAt(pos.bodyIndex);
    const len = text.length;
    let index = pos.charOffset;
    while (index < len && isWhitespaceChar(text[index] ?? "")) {
      index += 1;
    }
    return { bodyIndex: pos.bodyIndex, charOffset: index };
  }

  function deleteWordEnd(base: DocPosition): DocPosition {
    const pos = clampPosition(base);
    const text = paragraphTextAt(pos.bodyIndex);
    const len = text.length;
    let index = pos.charOffset;
    while (index < len && isWhitespaceChar(text[index] ?? "")) {
      index += 1;
    }
    while (index < len && !isWhitespaceChar(text[index] ?? "")) {
      index += 1;
    }
    while (index < len && isWhitespaceChar(text[index] ?? "")) {
      index += 1;
    }
    return { bodyIndex: pos.bodyIndex, charOffset: index };
  }

  function collectSelectedParagraphs(
    range: OrderedSelectionRange,
  ): Array<Extract<DocViewModel["body"][number], { type: "paragraph" }>> {
    return getParagraphBodyIndices()
      .filter(
        (idx) => idx >= range.start.bodyIndex && idx <= range.end.bodyIndex,
      )
      .map((idx) => paragraphItemAt(idx))
      .filter(
        (
          paragraph,
        ): paragraph is Extract<
          DocViewModel["body"][number],
          { type: "paragraph" }
        > => paragraph != null,
      );
  }

  function collectSelectedRuns(
    range: OrderedSelectionRange,
  ): Array<RunModel & { _bodyIndex: number }> {
    const runs: Array<RunModel & { _bodyIndex: number }> = [];
    for (
      let bodyIndex = range.start.bodyIndex;
      bodyIndex <= range.end.bodyIndex;
      bodyIndex += 1
    ) {
      const paragraph = paragraphItemAt(bodyIndex);
      if (!paragraph) continue;
      const startOffset =
        bodyIndex === range.start.bodyIndex ? range.start.charOffset : 0;
      const endOffset =
        bodyIndex === range.end.bodyIndex
          ? range.end.charOffset
          : paragraphCharCount(bodyIndex);
      let cursor = 0;
      for (const run of paragraph.runs) {
        const rawText = run.hasBreak
          ? `${run.text ?? ""}\n`
          : run.hasTab
            ? `${run.text ?? ""}\t`
            : run.inlineImage
              ? "\uFFFC"
              : run.footnoteRef != null || run.endnoteRef != null
                ? String(run.footnoteRef ?? run.endnoteRef ?? "")
                : (run.text ?? "");
        const runStart = cursor;
        const runEnd = cursor + rawText.length;
        cursor = runEnd;
        if (runEnd <= startOffset || runStart >= endOffset) continue;
        runs.push({ ...run, _bodyIndex: bodyIndex });
      }
    }
    return runs;
  }

  function uniformValue<T>(
    values: T[],
    equals: (a: T, b: T) => boolean = Object.is,
  ): T | undefined {
    if (values.length === 0) return undefined;
    const first = values[0];
    if (first === undefined) return undefined;
    for (let i = 1; i < values.length; i += 1) {
      if (!equals(first, values[i] as T)) return undefined;
    }
    return first;
  }

  function queryFormattingStateForRange(
    range: OrderedSelectionRange,
  ): FormattingState {
    const paragraphs = collectSelectedParagraphs(range);
    const runs = collectSelectedRuns(range);
    const state: FormattingState = {};

    const allTrue = (values: boolean[]) =>
      values.length > 0 && values.every(Boolean);
    if (allTrue(runs.map((run) => !!run.bold))) state.bold = true;
    if (allTrue(runs.map((run) => !!run.italic))) state.italic = true;
    if (allTrue(runs.map((run) => !!run.underline))) state.underline = true;
    if (allTrue(runs.map((run) => !!run.strikethrough))) state.strike = true;

    const fontFamily = uniformValue(
      runs.map((run) =>
        typeof run.fontFamily === "string" && run.fontFamily.length > 0
          ? run.fontFamily
          : undefined,
      ),
    );
    if (fontFamily) state.fontFamily = fontFamily;

    const fontSizePt = uniformValue(
      runs.map((run) =>
        typeof run.fontSizePt === "number" && Number.isFinite(run.fontSizePt)
          ? run.fontSizePt
          : undefined,
      ),
    );
    if (typeof fontSizePt === "number") state.fontSizePt = fontSizePt;

    const color = uniformValue(
      runs.map((run) =>
        typeof run.color === "string" && run.color.length > 0
          ? run.color
          : undefined,
      ),
    );
    if (color) state.color = normalizeColor(color);

    const highlight = uniformValue(
      runs.map((run) =>
        typeof run.highlight === "string" && run.highlight.length > 0
          ? run.highlight
          : undefined,
      ),
    );
    if (highlight) state.highlight = highlight;

    const hyperlink = uniformValue(
      runs.map((run) =>
        typeof run.hyperlink === "string" && run.hyperlink.length > 0
          ? run.hyperlink
          : undefined,
      ),
    );
    if (hyperlink) state.hyperlink = hyperlink;

    const headingLevel = uniformValue(
      paragraphs.map((paragraph) =>
        typeof paragraph.headingLevel === "number" &&
        Number.isFinite(paragraph.headingLevel)
          ? paragraph.headingLevel
          : undefined,
      ),
    );
    if (typeof headingLevel === "number") state.headingLevel = headingLevel;

    const alignment = uniformValue(
      paragraphs.map((paragraph) =>
        paragraph.alignment === "left" ||
        paragraph.alignment === "center" ||
        paragraph.alignment === "right" ||
        paragraph.alignment === "justify"
          ? paragraph.alignment
          : undefined,
      ),
    );
    if (alignment) state.alignment = alignment;

    const spacingBeforePt = uniformValue(
      paragraphs.map((paragraph) =>
        typeof paragraph.spacingBeforePt === "number" &&
        Number.isFinite(paragraph.spacingBeforePt)
          ? paragraph.spacingBeforePt
          : undefined,
      ),
    );
    if (typeof spacingBeforePt === "number")
      state.spacingBeforePt = spacingBeforePt;

    const spacingAfterPt = uniformValue(
      paragraphs.map((paragraph) =>
        typeof paragraph.spacingAfterPt === "number" &&
        Number.isFinite(paragraph.spacingAfterPt)
          ? paragraph.spacingAfterPt
          : undefined,
      ),
    );
    if (typeof spacingAfterPt === "number")
      state.spacingAfterPt = spacingAfterPt;

    const indentLeftPt = uniformValue(
      paragraphs.map((paragraph) =>
        typeof paragraph.indents?.leftPt === "number" &&
        Number.isFinite(paragraph.indents.leftPt)
          ? paragraph.indents.leftPt
          : undefined,
      ),
    );
    if (typeof indentLeftPt === "number") state.indentLeftPt = indentLeftPt;

    const indentFirstLinePt = uniformValue(
      paragraphs.map((paragraph) =>
        typeof paragraph.indents?.firstLinePt === "number" &&
        Number.isFinite(paragraph.indents.firstLinePt)
          ? paragraph.indents.firstLinePt
          : undefined,
      ),
    );
    if (typeof indentFirstLinePt === "number") {
      state.indentFirstLinePt = indentFirstLinePt;
    }

    const lineSpacingMultiple = uniformValue(
      paragraphs.map((paragraph) =>
        paragraph.lineSpacing?.rule === "auto" &&
        typeof paragraph.lineSpacing.value === "number" &&
        Number.isFinite(paragraph.lineSpacing.value)
          ? paragraph.lineSpacing.value
          : undefined,
      ),
    );
    if (typeof lineSpacingMultiple === "number") {
      state.lineSpacingMultiple = lineSpacingMultiple;
    }

    const listKind = uniformValue(
      paragraphs.map((paragraph) =>
        paragraph.numbering?.format === "bullet" ||
        paragraph.numbering?.format === "decimal"
          ? paragraph.numbering.format
          : undefined,
      ),
    );
    if (listKind) state.listKind = listKind;

    const listLevel = uniformValue(
      paragraphs.map((paragraph) =>
        typeof paragraph.numbering?.level === "number" &&
        Number.isFinite(paragraph.numbering.level)
          ? paragraph.numbering.level
          : undefined,
      ),
    );
    if (typeof listLevel === "number") state.listLevel = listLevel;

    return state;
  }

  function queryFormattingState(): FormattingState {
    const state: FormattingState = {};
    const currentSelection = renderer.getSelection();
    if (currentSelection) {
      const normalized = normalizeSelection({
        anchor: clampPosition(currentSelection.anchor),
        focus: clampPosition(currentSelection.focus),
      });
      if (!normalized.collapsed) {
        Object.assign(state, queryFormattingStateForRange(normalized));
        applyPendingFormattingState(state);
        return state;
      }
    }
    if (docEdit) {
      let formatPos: DocPosition | null = null;
      if (!useNativeSelection) {
        const sel = renderer.getSelection();
        if (sel) formatPos = clampPosition(sel.focus);
      } else {
        const winSel = window.getSelection();
        if (winSel && winSel.anchorNode) {
          const ep = resolveSelEndpoint(winSel.anchorNode, winSel.anchorOffset);
          if (ep)
            formatPos = { bodyIndex: ep.bodyIndex, charOffset: ep.charOffset };
        }
      }
      if (formatPos) {
        const pos = DocEdit.encodePosition(
          formatPos.bodyIndex,
          formatPos.charOffset,
        );
        const raw = docEdit.formattingAt(pos);
        if (raw && typeof raw === "object") {
          const obj = raw as Record<string, unknown>;
          if (obj.bold) state.bold = true;
          if (obj.italic) state.italic = true;
          if (obj.underline) state.underline = true;
          if (obj.strike) state.strike = true;
          if (typeof obj.fontFamily === "string") {
            state.fontFamily = obj.fontFamily;
          }
          if (typeof obj.fontSizePt === "number") {
            state.fontSizePt = obj.fontSizePt;
          }
          if (typeof obj.color === "string") {
            state.color = normalizeColor(obj.color);
          }
          if (typeof obj.highlight === "string") {
            state.highlight = obj.highlight;
          }
          if (typeof obj.hyperlink === "string") {
            state.hyperlink = obj.hyperlink;
          }
          if (typeof obj.headingLevel === "number") {
            state.headingLevel = obj.headingLevel;
          }
          if (
            obj.alignment === "left" ||
            obj.alignment === "center" ||
            obj.alignment === "right" ||
            obj.alignment === "justify"
          ) {
            state.alignment = obj.alignment as ParagraphAlignment;
          }
          if (typeof obj.spacingBeforePt === "number") {
            state.spacingBeforePt = obj.spacingBeforePt;
          }
          if (typeof obj.spacingAfterPt === "number") {
            state.spacingAfterPt = obj.spacingAfterPt;
          }
          if (typeof obj.indentLeftPt === "number") {
            state.indentLeftPt = obj.indentLeftPt;
          }
          if (typeof obj.indentFirstLinePt === "number") {
            state.indentFirstLinePt = obj.indentFirstLinePt;
          }
          if (typeof obj.lineSpacingMultiple === "number") {
            state.lineSpacingMultiple = obj.lineSpacingMultiple;
          }
          if (
            obj.numberingKind === "bullet" ||
            obj.numberingKind === "decimal"
          ) {
            state.listKind = obj.numberingKind;
          }
          if (typeof obj.numberingIlvl === "number") {
            state.listLevel = obj.numberingIlvl;
          }
        }
      }
    }
    applyPendingFormattingState(state);
    return state;
  }

  function fireFormattingChange(): void {
    const state = queryFormattingState();
    updateImageSelectionOverlay();
    for (const cb of formattingCallbacks) {
      cb(state);
    }
  }

  // --- Sync ---

  function connectSync(): void {
    if (!syncConfig || !docEdit) return;
    if (!syncProvider) {
      syncProvider = new SyncProvider(docEdit, syncConfig.roomId, {
        wsUrl: syncConfig.wsUrl,
      });
      syncProvider.onRemote(() => rerenderFromState());
      syncProvider.onReplace((snapshot, generation) => {
        replaceFromSnapshot(snapshot, generation);
      });
      syncProvider.onAwareness((peers) => {
        remotePeers = peers;
        scheduleRemotePresenceRender();
      });
      syncProvider.onStatus((status) => {
        emitEvent("docedit-sync", status);
      });
      syncProvider.connect();
    } else {
      syncProvider.attachDocEdit(docEdit);
    }
    updateLocalAwareness();
  }

  function replaceFromSnapshot(snapshot: Uint8Array, generation: number): void {
    const nextDocEdit = new DocEdit(snapshot);
    closeTableCellEditor({ commit: false });
    clearPendingFormat();
    docEdit = nextDocEdit;
    syncProvider?.attachDocEdit(nextDocEdit, generation);
    rerenderFromState();
    emitEvent("docedit-dirty", { dirty: false });
    fireFormattingChange();
  }

  function loadEditorState(
    nextDocEdit: DocEdit,
    perfLabel: "load" | "loadBlank" | "replace" | "replaceBlank",
    broadcastReplaceBytes?: Uint8Array,
  ): number {
    const t0 = performance.now();
    closeTableCellEditor({ commit: false });
    clearPendingFormat();
    docEdit = nextDocEdit;
    const tViewStart = performance.now();
    const model = docEdit.viewModel() as DocViewModel;
    const tAfterView = performance.now();
    currentModel = model;
    renderer.renderModel(model);
    updateImageSelectionOverlay();
    const tAfterRender = performance.now();
    emitEvent("docedit-dirty", { dirty: docEdit.isDirty() });
    updateLocalAwareness();
    scheduleRemotePresenceRender();
    recordPerf(
      perfLabel,
      t0,
      0,
      tAfterView - tViewStart,
      tAfterRender - tAfterView,
    );
    connectSync();
    if (broadcastReplaceBytes) {
      syncProvider?.broadcastReplace(broadcastReplaceBytes);
      emitEvent("docedit-dirty", { dirty: false });
    }
    fireFormattingChange();
    return Math.round(tAfterRender - t0);
  }

  function rerenderFromState(): void {
    if (!docEdit) return;
    try {
      const t0 = performance.now();
      const savedScroll = container.scrollTop;
      const tViewStart = performance.now();
      const model = docEdit.viewModel() as DocViewModel;
      const tAfterView = performance.now();
      currentModel = model;
      renderer.renderModel(model);
      refreshActiveTableCell();
      updateImageSelectionOverlay();
      const tAfterRender = performance.now();
      container.scrollTop = savedScroll;
      emitEvent("docedit-dirty", { dirty: docEdit.isDirty() });
      updateLocalAwareness();
      scheduleRemotePresenceRender();
      recordPerf(
        "remoteViewUpdate",
        t0,
        0,
        tAfterView - tViewStart,
        tAfterRender - tAfterView,
      );
    } catch (err) {
      console.error("rerenderFromState failed:", err);
    }
  }

  // --- Helpers ---

  interface SavedSelection {
    anchorBody: number;
    anchorOff: number;
    focusBody: number;
    focusOff: number;
  }

  function captureSelection(): SavedSelection | null {
    const sel = renderer.getSelection();
    if (!sel) return null;
    return {
      anchorBody: sel.anchor.bodyIndex,
      anchorOff: sel.anchor.charOffset,
      focusBody: sel.focus.bodyIndex,
      focusOff: sel.focus.charOffset,
    };
  }

  function applyAndRender(
    intentJson: string,
    cursorBodyIndex?: number,
    cursorCharOffset?: number,
    selection?: SavedSelection,
  ): boolean {
    if (!docEdit) return false;
    try {
      const t0 = performance.now();
      const savedScroll = container.scrollTop;
      const tApplyStart = performance.now();
      docEdit.applyIntent(intentJson);
      const tAfterApply = performance.now();
      const model = docEdit.viewModel() as DocViewModel;
      const tAfterView = performance.now();
      currentModel = model;
      renderer.renderModel(model);
      refreshActiveTableCell();
      updateImageSelectionOverlay();
      const tAfterRender = performance.now();
      container.scrollTop = savedScroll;

      if (selection) {
        renderer.setSelection({
          anchor: {
            bodyIndex: selection.anchorBody,
            charOffset: selection.anchorOff,
          },
          focus: {
            bodyIndex: selection.focusBody,
            charOffset: selection.focusOff,
          },
        });
      } else if (cursorBodyIndex != null && cursorCharOffset != null) {
        renderer.setCursor({
          bodyIndex: cursorBodyIndex,
          charOffset: cursorCharOffset,
        });
      }

      if (!renderer.isFocused()) {
        renderer.focus();
      }
      ensureCanvasInputFocus();
      scrollCursorIntoView();
      emitEvent("docedit-dirty", { dirty: docEdit.isDirty() });
      syncProvider?.broadcastUpdate();
      updateLocalAwareness();
      scheduleRemotePresenceRender();
      recordPerf(
        "applyIntent",
        t0,
        tAfterApply - tApplyStart,
        tAfterView - tAfterApply,
        tAfterRender - tAfterView,
      );
      return true;
    } catch (err) {
      console.error("applyIntent failed:", err);
      return false;
    }
  }

  function applyBatchAndRender(
    intentJsons: string[],
    cursorBodyIndex?: number,
    cursorCharOffset?: number,
    selection?: SavedSelection,
  ): boolean {
    if (!docEdit || intentJsons.length === 0) return false;
    try {
      const t0 = performance.now();
      const savedScroll = container.scrollTop;
      const tApplyStart = performance.now();
      for (const intentJson of intentJsons) {
        docEdit.applyIntent(intentJson);
      }
      const tAfterApply = performance.now();
      const model = docEdit.viewModel() as DocViewModel;
      const tAfterView = performance.now();
      currentModel = model;
      renderer.renderModel(model);
      refreshActiveTableCell();
      const tAfterRender = performance.now();
      container.scrollTop = savedScroll;

      if (selection) {
        renderer.setSelection({
          anchor: {
            bodyIndex: selection.anchorBody,
            charOffset: selection.anchorOff,
          },
          focus: {
            bodyIndex: selection.focusBody,
            charOffset: selection.focusOff,
          },
        });
      } else if (cursorBodyIndex != null && cursorCharOffset != null) {
        renderer.setCursor({
          bodyIndex: cursorBodyIndex,
          charOffset: cursorCharOffset,
        });
      }

      if (!renderer.isFocused()) {
        renderer.focus();
      }
      ensureCanvasInputFocus();
      scrollCursorIntoView();
      emitEvent("docedit-dirty", { dirty: docEdit.isDirty() });
      syncProvider?.broadcastUpdate();
      updateLocalAwareness();
      scheduleRemotePresenceRender();
      recordPerf(
        "applyIntentBatch",
        t0,
        tAfterApply - tApplyStart,
        tAfterView - tAfterApply,
        tAfterRender - tAfterView,
      );
      return true;
    } catch (err) {
      console.error("applyIntentBatch failed:", err);
      return false;
    }
  }

  function scrollCursorIntoView(): void {
    const rect = renderer.getCursorRect();
    if (!rect) return;

    const cr = container.getBoundingClientRect();
    const margin = 40;

    if (rect.y + rect.h > cr.bottom - margin) {
      container.scrollTop += rect.y + rect.h - cr.bottom + margin;
    } else if (rect.y < cr.top + margin) {
      container.scrollTop -= cr.top + margin - rect.y;
    }
  }

  function emitEvent(name: string, detail?: unknown): void {
    container.dispatchEvent(new CustomEvent(name, { detail, bubbles: true }));
  }

  function recordPerf(
    op: string,
    startMs: number,
    applyMs: number,
    viewModelMs: number,
    renderMs: number,
  ): void {
    const totalMs = Math.max(0, performance.now() - startMs);
    const summary = perfTracker.record({
      op,
      applyMs,
      viewModelMs,
      renderMs,
      totalMs,
      atMs: Date.now(),
    });
    emitEvent("docedit-perf", {
      op,
      sample: {
        applyMs,
        viewModelMs,
        renderMs,
        totalMs,
      },
      summary,
    });
  }

  function ensureCanvasInputFocus(): void {
    const canvasInput = container.querySelector<HTMLTextAreaElement>(
      'textarea[data-docview-canvas-input="1"]',
    );
    if (!canvasInput) return;
    const root = container.getRootNode() as Document | ShadowRoot;
    const active =
      "activeElement" in root ? root.activeElement : document.activeElement;
    if (active === canvasInput) return;
    canvasInput.focus({ preventScroll: true });
  }

  /** Get current selection as encoded CRDT position strings. */
  function getPositions(): { anchor: string; focus: string } {
    const sel = renderer.getSelection();
    if (sel) {
      const anchor = clampPosition(sel.anchor);
      const focus = clampPosition(sel.focus);
      return {
        anchor: DocEdit.encodePosition(anchor.bodyIndex, anchor.charOffset),
        focus: DocEdit.encodePosition(focus.bodyIndex, focus.charOffset),
      };
    }
    // Fallback: try from window.getSelection() for HTML renderer
    const winSel = window.getSelection();
    if (winSel && winSel.anchorNode && winSel.focusNode) {
      const docSel = selectionToDocPositions(winSel);
      if (docSel) {
        return {
          anchor: DocEdit.encodePosition(
            docSel.anchor.bodyIndex,
            docSel.anchor.charOffset,
          ),
          focus: DocEdit.encodePosition(
            docSel.focus.bodyIndex,
            docSel.focus.charOffset,
          ),
        };
      }
    }
    return {
      anchor: DocEdit.encodePosition(0, 0),
      focus: DocEdit.encodePosition(0, 0),
    };
  }

  function applyFormat(formatType: string, attrName: string): void {
    closeTableCellEditor({ commit: true });
    const positions = getPositions();

    if (positions.anchor === positions.focus) {
      togglePendingFormat(attrName);
      fireFormattingChange();
      return;
    }

    const saved = captureSelection();
    const intent = JSON.stringify({
      type: formatType,
      anchor: positions.anchor,
      focus: positions.focus,
    });
    applyAndRender(intent, undefined, undefined, saved ?? undefined);
    fireFormattingChange();
  }

  function setTextStyle(patch: TextStylePatch): void {
    if (!docEdit) return;
    closeTableCellEditor({ commit: true });
    const attrs = toIntentTextAttrs(patch);
    if (Object.keys(attrs).length === 0) return;

    const positions = getPositions();
    if (positions.anchor === positions.focus) {
      mergePendingTextAttrs(patch);
      fireFormattingChange();
      return;
    }

    const saved = captureSelection();
    const intent = JSON.stringify({
      type: "setTextAttrs",
      anchor: positions.anchor,
      focus: positions.focus,
      attrs,
    });
    applyAndRender(intent, undefined, undefined, saved ?? undefined);
    fireFormattingChange();
  }

  function setParagraphStyle(patch: ParagraphStylePatch): void {
    if (!docEdit) return;
    closeTableCellEditor({ commit: true });
    const attrs = toIntentParagraphAttrs(patch);
    if (Object.keys(attrs).length === 0) return;

    const positions = getPositions();
    const saved = captureSelection();
    const intent = JSON.stringify({
      type: "setParagraphAttrs",
      anchor: positions.anchor,
      focus: positions.focus,
      attrs,
    });
    applyAndRender(intent, undefined, undefined, saved ?? undefined);
    fireFormattingChange();
  }

  function selectedParagraphBodyIndices(): number[] {
    const range = normalizeSelection(selectionOrCursor());
    return getParagraphBodyIndices().filter((idx) => {
      return idx >= range.start.bodyIndex && idx <= range.end.bodyIndex;
    });
  }

  function nextSyntheticListId(): number {
    let maxId = 0;
    for (const item of currentModel?.body ?? []) {
      if (item?.type !== "paragraph") continue;
      maxId = Math.max(maxId, item.numbering?.numId ?? 0);
    }
    return maxId + 1;
  }

  function resolveTargetListId(
    kind: ListKind,
    paragraphBodyIndices: number[],
  ): number {
    const firstBodyIndex = paragraphBodyIndices[0];
    if (firstBodyIndex == null) return nextSyntheticListId();
    const previousParagraphs = getParagraphBodyIndices().filter(
      (idx) => idx < firstBodyIndex,
    );
    const previousBodyIndex = previousParagraphs[previousParagraphs.length - 1];
    if (
      previousBodyIndex != null &&
      paragraphListKindAt(previousBodyIndex) === kind
    ) {
      const prevId = paragraphListIdAt(previousBodyIndex);
      if (typeof prevId === "number" && Number.isFinite(prevId)) {
        return prevId;
      }
    }
    const nextParagraphs = getParagraphBodyIndices().filter(
      (idx) => idx > firstBodyIndex,
    );
    const nextBodyIndex = nextParagraphs[0];
    if (nextBodyIndex != null && paragraphListKindAt(nextBodyIndex) === kind) {
      const nextId = paragraphListIdAt(nextBodyIndex);
      if (typeof nextId === "number" && Number.isFinite(nextId)) {
        return nextId;
      }
    }
    return nextSyntheticListId();
  }

  function collectContiguousListNormalizationPatches(
    kind: ListKind,
    paragraphBodyIndices: number[],
    listId: number,
  ): Array<{
    bodyIndex: number;
    attrs: Record<string, number | string>;
  }> {
    const selected = new Set(paragraphBodyIndices);
    const allParagraphs = getParagraphBodyIndices();
    const lastSelected = paragraphBodyIndices[paragraphBodyIndices.length - 1];
    if (lastSelected == null) return [];
    const patches: Array<{
      bodyIndex: number;
      attrs: Record<string, number | string>;
    }> = [];

    for (const bodyIndex of allParagraphs) {
      if (bodyIndex <= lastSelected) continue;
      if (selected.has(bodyIndex)) continue;
      if (paragraphListKindAt(bodyIndex) !== kind) break;
      const currentId = paragraphListIdAt(bodyIndex);
      const currentLevel = paragraphListLevelAt(bodyIndex) ?? 0;
      if (currentId === listId) continue;
      patches.push({
        bodyIndex,
        attrs: {
          numberingKind: kind,
          numberingNumId: listId,
          numberingIlvl: currentLevel,
        },
      });
    }

    return patches;
  }

  function adjustSelectedListLevel(delta: -1 | 1): boolean {
    if (!docEdit) return false;
    const paragraphBodyIndices = selectedParagraphBodyIndices();
    if (paragraphBodyIndices.length === 0) return false;

    const patches: Array<{
      bodyIndex: number;
      attrs: Record<string, string | number>;
    }> = [];
    for (const bodyIndex of paragraphBodyIndices) {
      const numbering = paragraphNumberingAttrsAt(bodyIndex);
      if (!numbering) return false;
      const nextLevel = Math.max(0, Math.min(8, numbering.level + delta));
      if (nextLevel === numbering.level) continue;
      patches.push({
        bodyIndex,
        attrs: {
          numberingKind: numbering.kind,
          numberingNumId: numbering.numId,
          numberingIlvl: nextLevel,
        },
      });
    }

    const saved = captureSelection();
    if (patches.length === 0) {
      if (saved) {
        renderer.setSelection({
          anchor: {
            bodyIndex: saved.anchorBody,
            charOffset: saved.anchorOff,
          },
          focus: {
            bodyIndex: saved.focusBody,
            charOffset: saved.focusOff,
          },
        });
      }
      if (!renderer.isFocused()) renderer.focus();
      ensureCanvasInputFocus();
      fireFormattingChange();
      return true;
    }

    const intents = patches.map(({ bodyIndex, attrs }) =>
      JSON.stringify({
        type: "setParagraphAttrs",
        anchor: DocEdit.encodePosition(bodyIndex, 0),
        focus: DocEdit.encodePosition(bodyIndex, 0),
        attrs,
      }),
    );
    applyBatchAndRender(intents, undefined, undefined, saved ?? undefined);
    fireFormattingChange();
    return true;
  }

  function toggleList(kind: ListKind): void {
    if (!docEdit) return;
    closeTableCellEditor({ commit: true });
    const paragraphBodyIndices = selectedParagraphBodyIndices();
    if (paragraphBodyIndices.length === 0) return;

    const allAlreadySameKind = paragraphBodyIndices.every(
      (bodyIndex) => paragraphListKindAt(bodyIndex) === kind,
    );

    if (allAlreadySameKind) {
      setParagraphStyle({
        numberingKind: null,
        numberingNumId: null,
        numberingIlvl: null,
      });
      return;
    }

    const listId = resolveTargetListId(kind, paragraphBodyIndices);
    const attrs = {
      headingLevel: null,
      numberingKind: kind,
      numberingNumId: listId,
      numberingIlvl: 0,
    };
    const normalizationPatches = collectContiguousListNormalizationPatches(
      kind,
      paragraphBodyIndices,
      listId,
    );
    if (normalizationPatches.length === 0) {
      setParagraphStyle(attrs);
      return;
    }

    const positions = getPositions();
    const saved = captureSelection();
    const intents = [
      JSON.stringify({
        type: "setParagraphAttrs",
        anchor: positions.anchor,
        focus: positions.focus,
        attrs,
      }),
      ...normalizationPatches.map(({ bodyIndex, attrs }) =>
        JSON.stringify({
          type: "setParagraphAttrs",
          anchor: DocEdit.encodePosition(bodyIndex, 0),
          focus: DocEdit.encodePosition(bodyIndex, 0),
          attrs,
        }),
      ),
    ];
    applyBatchAndRender(intents, undefined, undefined, saved ?? undefined);
    fireFormattingChange();
  }

  function nearestTableBodyIndex(preferredBodyIndex: number): number | null {
    if (!currentModel) return null;
    let bestBodyIndex: number | null = null;
    let bestDistance = Number.POSITIVE_INFINITY;
    for (let i = 0; i < currentModel.body.length; i += 1) {
      if (currentModel.body[i]?.type !== "table") continue;
      const distance = Math.abs(i - preferredBodyIndex);
      if (
        distance < bestDistance ||
        (distance === bestDistance &&
          (bestBodyIndex == null || i < bestBodyIndex))
      ) {
        bestBodyIndex = i;
        bestDistance = distance;
      }
    }
    return bestBodyIndex;
  }

  function insertInlineImage(payload: InlineImageInsertPayload): boolean {
    if (!docEdit) return false;
    closeTableCellEditor({ commit: true });
    const sel = selectionOrCursor();
    const range = normalizeSelection(sel);
    const anchor = clampPosition(sel.anchor);
    const focus = clampPosition(sel.focus);
    const ok = applyAndRender(
      JSON.stringify({
        type: "insertInlineImage",
        anchor: DocEdit.encodePosition(anchor.bodyIndex, anchor.charOffset),
        focus: DocEdit.encodePosition(focus.bodyIndex, focus.charOffset),
        data_uri: payload.dataUri,
        width_pt: Math.max(1, payload.widthPt),
        height_pt: Math.max(1, payload.heightPt),
        ...(payload.name ? { name: payload.name } : {}),
        ...(payload.description ? { description: payload.description } : {}),
      }),
      range.start.bodyIndex,
      range.start.charOffset + 1,
    );
    if (ok) {
      syncProvider?.flushFullState();
    }
    return ok;
  }

  async function blobToDataUri(blob: Blob): Promise<string> {
    const bytes = new Uint8Array(await blob.arrayBuffer());
    let binary = "";
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
    }
    return `data:${blob.type || "application/octet-stream"};base64,${btoa(binary)}`;
  }

  async function measureInlineImagePt(
    dataUri: string,
  ): Promise<{ widthPt: number; heightPt: number }> {
    if (typeof Image === "undefined") {
      return { widthPt: 24, heightPt: 24 };
    }
    return await new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        const widthPx = Math.max(1, img.naturalWidth || img.width || 1);
        const heightPx = Math.max(1, img.naturalHeight || img.height || 1);
        resolve({
          widthPt: Math.max(1, (widthPx * 72) / 96),
          heightPt: Math.max(1, (heightPx * 72) / 96),
        });
      };
      img.onerror = () => resolve({ widthPt: 24, heightPt: 24 });
      img.src = dataUri;
    });
  }

  async function insertPastedImage(file: File): Promise<void> {
    const dataUri = await blobToDataUri(file);
    const size = await measureInlineImagePt(dataUri);
    insertInlineImage({
      dataUri,
      widthPt: size.widthPt,
      heightPt: size.heightPt,
      name: file.name || undefined,
      description: file.name || undefined,
    });
  }

  function resizeSelectedInlineImage(
    widthPt: number,
    heightPt: number,
  ): boolean {
    if (!docEdit) return false;
    closeTableCellEditor({ commit: true });
    const selected = getSelectedInlineImage();
    if (!selected) return false;
    const nextWidthPt = Math.max(8, widthPt);
    const nextHeightPt = Math.max(8, heightPt);
    const saved = captureSelection();
    const ok = applyAndRender(
      JSON.stringify({
        type: "setTextAttrs",
        anchor: DocEdit.encodePosition(selected.bodyIndex, selected.charOffset),
        focus: DocEdit.encodePosition(
          selected.bodyIndex,
          selected.charOffset + 1,
        ),
        attrs: {
          widthPt: nextWidthPt,
          heightPt: nextHeightPt,
        },
      }),
      undefined,
      undefined,
      saved ?? undefined,
    );
    if (ok) {
      fireFormattingChange();
      updateImageSelectionOverlay();
    }
    return ok;
  }

  function setSelectedInlineImageAlignment(
    alignment: ImageBlockAlignment,
  ): boolean {
    const selected = getSelectedInlineImage();
    if (!selected) return false;
    setParagraphStyle({ alignment });
    return true;
  }

  function insertTable(rows: number, columns: number): boolean {
    if (!docEdit) return false;
    closeTableCellEditor({ commit: true });
    const anchor = selectionOrCursor().anchor;
    const ok = applyAndRender(
      JSON.stringify({
        type: "insertTable",
        anchor: DocEdit.encodePosition(anchor.bodyIndex, anchor.charOffset),
        rows: Math.max(1, Math.round(rows)),
        columns: Math.max(1, Math.round(columns)),
      }),
    );
    if (!ok) return false;
    syncProvider?.flushFullState();
    const bodyIndex = nearestTableBodyIndex(anchor.bodyIndex);
    if (bodyIndex != null) {
      openTableCellEditor({ bodyIndex, row: 0, col: 0 });
    }
    return true;
  }

  function insertTableRow(): boolean {
    if (!docEdit || !activeTableCell) return false;
    const target = { ...activeTableCell };
    const ok = applyAndRender(
      JSON.stringify({
        type: "insertTableRow",
        bodyIndex: target.bodyIndex,
        row: target.row + 1,
      }),
    );
    if (!ok) return false;
    syncProvider?.flushFullState();
    openTableCellEditor({
      bodyIndex: target.bodyIndex,
      row: target.row + 1,
      col: target.col,
    });
    return true;
  }

  function removeTableRow(): boolean {
    if (!docEdit || !activeTableCell) return false;
    const target = { ...activeTableCell };
    const { rows, cols } = tableCellCount(target.bodyIndex);
    if (rows <= 1) return false;
    const nextRow = Math.max(0, Math.min(target.row, rows - 2));
    const nextCol = Math.min(target.col, Math.max(0, cols - 1));
    const ok = applyAndRender(
      JSON.stringify({
        type: "removeTableRow",
        bodyIndex: target.bodyIndex,
        row: target.row,
      }),
    );
    if (!ok) return false;
    syncProvider?.flushFullState();
    openTableCellEditor({
      bodyIndex: target.bodyIndex,
      row: nextRow,
      col: nextCol,
    });
    return true;
  }

  function insertTableColumn(): boolean {
    if (!docEdit || !activeTableCell) return false;
    const target = { ...activeTableCell };
    const ok = applyAndRender(
      JSON.stringify({
        type: "insertTableColumn",
        bodyIndex: target.bodyIndex,
        col: target.col + 1,
      }),
    );
    if (!ok) return false;
    syncProvider?.flushFullState();
    openTableCellEditor({
      bodyIndex: target.bodyIndex,
      row: target.row,
      col: target.col + 1,
    });
    return true;
  }

  function removeTableColumn(): boolean {
    if (!docEdit || !activeTableCell) return false;
    const target = { ...activeTableCell };
    const { rows, cols } = tableCellCount(target.bodyIndex);
    if (cols <= 1) return false;
    const nextRow = Math.min(target.row, Math.max(0, rows - 1));
    const nextCol = Math.max(0, Math.min(target.col, cols - 2));
    const ok = applyAndRender(
      JSON.stringify({
        type: "removeTableColumn",
        bodyIndex: target.bodyIndex,
        col: target.col,
      }),
    );
    if (!ok) return false;
    syncProvider?.flushFullState();
    openTableCellEditor({
      bodyIndex: target.bodyIndex,
      row: nextRow,
      col: nextCol,
    });
    return true;
  }

  function setActiveTableCellText(text: string): boolean {
    if (!activeTableCell) return false;
    tableCellEditor.value = text;
    return applyTableCellEditorValue({ flushFullState: true });
  }

  function moveActiveTableCell(deltaRow: number, deltaCol: number): boolean {
    if (!activeTableCell) return false;
    const { rows, cols } = tableCellCount(activeTableCell.bodyIndex);
    if (rows <= 0 || cols <= 0) return false;
    const nextRow = Math.max(
      0,
      Math.min(rows - 1, activeTableCell.row + deltaRow),
    );
    const nextCol = Math.max(
      0,
      Math.min(cols - 1, activeTableCell.col + deltaCol),
    );
    if (nextRow === activeTableCell.row && nextCol === activeTableCell.col) {
      return false;
    }
    openTableCellEditor({
      bodyIndex: activeTableCell.bodyIndex,
      row: nextRow,
      col: nextCol,
    });
    return true;
  }

  function selectInlineImageAt(hit: {
    bodyIndex: number;
    charOffset: number;
  }): void {
    renderer.setSelection({
      anchor: {
        bodyIndex: hit.bodyIndex,
        charOffset: hit.charOffset,
      },
      focus: {
        bodyIndex: hit.bodyIndex,
        charOffset: hit.charOffset + 1,
      },
    });
  }

  // --- Wire up InputAdapter ---

  let pointerSelecting = false;
  let pointerAnchor: { bodyIndex: number; charOffset: number } | null = null;
  let pendingPointerMove: { x: number; y: number } | null = null;
  let pointerMoveRaf = 0;

  function flushPointerMove(): void {
    pointerMoveRaf = 0;
    if (destroyed || !docEdit) return;
    if (useNativeSelection) return;
    if (!pointerSelecting || !pointerAnchor || !pendingPointerMove) return;

    const { x, y } = pendingPointerMove;
    pendingPointerMove = null;

    const hit = renderer.hitTest(x, y);
    if (!hit) return;

    const sameAsAnchor =
      hit.bodyIndex === pointerAnchor.bodyIndex &&
      hit.charOffset === pointerAnchor.charOffset;

    if (sameAsAnchor) {
      renderer.setCursor(pointerAnchor);
      updateLocalAwareness();
      scheduleRemotePresenceRender();
      return;
    }

    renderer.setSelection({
      anchor: pointerAnchor,
      focus: {
        bodyIndex: hit.bodyIndex,
        charOffset: hit.charOffset,
      },
    });
    updateLocalAwareness();
    scheduleRemotePresenceRender();
  }

  input.onInput((normalized: NormalizedInput) => {
    if (!docEdit || destroyed) return;

    const sel = selectionOrCursor();
    const range = normalizeSelection(sel);
    const anchor = sel.anchor;
    const focus = sel.focus;
    const positions = {
      anchor: DocEdit.encodePosition(anchor.bodyIndex, anchor.charOffset),
      focus: DocEdit.encodePosition(focus.bodyIndex, focus.charOffset),
    };
    const rangeStart = range.start;

    switch (normalized.type) {
      case "insertText":
      case "insertFromComposition": {
        const text = normalized.data;
        if (!text) return;
        const pending = consumePendingAttrs();
        const collapsedIntentType =
          normalized.type === "insertFromComposition"
            ? "insertFromComposition"
            : "insertText";
        if (!range.collapsed) {
          const replaceIntentObj: Record<string, unknown> = {
            type: "insertFromPaste",
            data: text,
            anchor: positions.anchor,
            focus: positions.focus,
          };
          if (pending) replaceIntentObj.attrs = pending;
          applyAndRender(
            JSON.stringify(replaceIntentObj),
            rangeStart.bodyIndex,
            rangeStart.charOffset + text.length,
          );
          break;
        }
        const intentObj: Record<string, unknown> = {
          type: collapsedIntentType,
          data: text,
          anchor: positions.anchor,
        };
        if (pending) intentObj.attrs = pending;
        applyAndRender(
          JSON.stringify(intentObj),
          anchor.bodyIndex,
          anchor.charOffset + text.length,
        );
        break;
      }
      case "deleteContentBackward": {
        const intent = JSON.stringify({
          type: "deleteContentBackward",
          anchor: positions.anchor,
          focus: positions.focus,
        });
        const nextCursor = range.collapsed
          ? moveByCharacter(rangeStart, -1)
          : rangeStart;
        applyAndRender(intent, nextCursor.bodyIndex, nextCursor.charOffset);
        break;
      }
      case "deleteContentForward": {
        const intent = JSON.stringify({
          type: "deleteContentForward",
          anchor: positions.anchor,
          focus: positions.focus,
        });
        applyAndRender(intent, rangeStart.bodyIndex, rangeStart.charOffset);
        break;
      }
      case "deleteWordBackward": {
        const deleteEnd = range.collapsed
          ? deleteWordBackwardEnd(rangeStart)
          : range.end;
        const deleteStart = range.collapsed
          ? deleteWordStart(rangeStart)
          : rangeStart;
        const intent = JSON.stringify({
          type: "deleteByCut",
          anchor: DocEdit.encodePosition(
            deleteStart.bodyIndex,
            deleteStart.charOffset,
          ),
          focus: DocEdit.encodePosition(
            deleteEnd.bodyIndex,
            deleteEnd.charOffset,
          ),
        });
        applyAndRender(intent, deleteStart.bodyIndex, deleteStart.charOffset);
        break;
      }
      case "deleteWordForward": {
        const deleteEnd = range.collapsed
          ? deleteWordEnd(rangeStart)
          : range.end;
        const intent = JSON.stringify({
          type: "deleteByCut",
          anchor: positions.anchor,
          focus: DocEdit.encodePosition(
            deleteEnd.bodyIndex,
            deleteEnd.charOffset,
          ),
        });
        applyAndRender(intent, rangeStart.bodyIndex, rangeStart.charOffset);
        break;
      }
      case "insertParagraph": {
        const currentParagraph = paragraphItemAt(anchor.bodyIndex);
        const currentParagraphText = paragraphTextAt(anchor.bodyIndex);
        const currentListKind = paragraphListKindAt(anchor.bodyIndex);
        if (
          range.collapsed &&
          currentParagraph &&
          currentListKind &&
          currentParagraphText.length === 0
        ) {
          setParagraphStyle({
            numberingKind: null,
            numberingNumId: null,
            numberingIlvl: null,
          });
          break;
        }
        seedPendingTextAttrsFromState(queryFormattingState());
        const intent = JSON.stringify({
          type: "insertParagraph",
          anchor: positions.anchor,
        });
        applyAndRender(intent, anchor.bodyIndex + 1, 0);
        fireFormattingChange();
        break;
      }
      case "insertLineBreak": {
        const intent = JSON.stringify({
          type: "insertLineBreak",
          anchor: positions.anchor,
        });
        applyAndRender(intent, anchor.bodyIndex, anchor.charOffset + 1);
        break;
      }
      case "insertFromPaste": {
        const text = normalized.data ?? "";
        const html = normalized.html ?? "";
        if (html) {
          const parsed = parseRichClipboardHtml(html);
          if (parsed && parsed.blocks.length > 0) {
            const ok = applyRichPasteDocument(parsed, range, positions);
            if (ok) break;
          }
        }
        if (!text) return;
        const intent = JSON.stringify({
          type: "insertFromPaste",
          data: text,
          anchor: positions.anchor,
          focus: positions.focus,
        });
        applyAndRender(
          intent,
          rangeStart.bodyIndex,
          rangeStart.charOffset + text.length,
        );
        break;
      }
      case "deleteByCut": {
        const intent = JSON.stringify({
          type: "deleteByCut",
          anchor: positions.anchor,
          focus: positions.focus,
        });
        applyAndRender(intent, rangeStart.bodyIndex, rangeStart.charOffset);
        break;
      }
      case "historyUndo":
        applyAndRender(JSON.stringify({ type: "historyUndo" }));
        break;
      case "historyRedo":
        applyAndRender(JSON.stringify({ type: "historyRedo" }));
        break;
      case "insertTab": {
        if (adjustSelectedListLevel(normalized.shift ? -1 : 1)) {
          break;
        }
        const intent = JSON.stringify({
          type: "insertTab",
          anchor: DocEdit.encodePosition(anchor.bodyIndex, anchor.charOffset),
        });
        applyAndRender(intent, anchor.bodyIndex, anchor.charOffset + 1);
        break;
      }
    }
  });

  input.onShortcut((key: string, shift: boolean) => {
    if (!docEdit || destroyed) return;

    switch (key) {
      case "b":
        applyFormat("formatBold", "bold");
        break;
      case "i":
        applyFormat("formatItalic", "italic");
        break;
      case "u":
        applyFormat("formatUnderline", "underline");
        break;
      case "5":
        if (shift) applyFormat("formatStrikethrough", "strike");
        break;
      case "s":
        emitEvent("docedit-save");
        break;
      case "a":
        if (useNativeSelection) selectAllInNative();
        else selectAllInCanvas();
        if (!renderer.isFocused()) renderer.focus();
        ensureCanvasInputFocus();
        clearPendingFormat();
        fireFormattingChange();
        break;
    }
  });

  input.onPointerDown((payload: PointerDownPayload) => {
    if (destroyed || !docEdit) return;
    const tableCell = renderer.getTableCellAtPoint?.(payload.x, payload.y);
    if (tableCell) {
      pointerSelecting = false;
      pointerAnchor = null;
      pendingPointerMove = null;
      openTableCellEditor(tableCell);
      requestAnimationFrame(() => {
        const current = activeTableCell;
        if (
          !current ||
          current.bodyIndex !== tableCell.bodyIndex ||
          current.row !== tableCell.row ||
          current.col !== tableCell.col
        ) {
          openTableCellEditor(tableCell);
          return;
        }
        tableCellEditor.focus({ preventScroll: true });
      });
      return;
    }
    if (activeTableCell) {
      closeTableCellEditor({ commit: true });
    }
    const inlineImage = renderer.getInlineImageAtPoint?.(payload.x, payload.y);
    if (inlineImage) {
      pointerSelecting = false;
      pointerAnchor = null;
      pendingPointerMove = null;
      if (payload.shift) {
        const current = selectionOrCursor();
        renderer.setSelection({
          anchor: clampPosition(current.anchor),
          focus: {
            bodyIndex: inlineImage.bodyIndex,
            charOffset: inlineImage.charOffset + 1,
          },
        });
      } else {
        selectInlineImageAt(inlineImage);
      }
      if (!renderer.isFocused()) {
        renderer.focus();
      }
      ensureCanvasInputFocus();
      clearPendingFormat();
      fireFormattingChange();
      updateLocalAwareness();
      scheduleRemotePresenceRender();
      return;
    }
    if (useNativeSelection) return;
    // Use hitTest to place the cursor at the clicked position.
    // This is essential for the canvas renderer where there's no
    // contenteditable to handle cursor placement natively.
    const hit = renderer.hitTest(payload.x, payload.y);
    if (hit) {
      const hitPos = { bodyIndex: hit.bodyIndex, charOffset: hit.charOffset };
      const clickCount = Math.max(1, payload.clickCount || 1);
      pointerAnchor = { ...hitPos };
      if (clickCount >= 3) {
        pointerSelecting = false;
        renderer.setSelection(paragraphSelectionAt(hitPos));
      } else if (clickCount === 2) {
        pointerSelecting = false;
        renderer.setSelection(wordSelectionAt(hitPos));
      } else {
        pointerSelecting = true;
        renderer.setCursor({
          bodyIndex: hit.bodyIndex,
          charOffset: hit.charOffset,
        });
      }
      if (!renderer.isFocused()) {
        renderer.focus();
      }
      ensureCanvasInputFocus();
      clearPendingFormat();
      fireFormattingChange();
      updateLocalAwareness();
      scheduleRemotePresenceRender();
    }
  });

  input.onPointerMove((x: number, y: number) => {
    if (destroyed || !docEdit) return;
    if (activeTableCell) return;
    if (useNativeSelection) return;
    if (!pointerSelecting || !pointerAnchor) return;
    pendingPointerMove = { x, y };
    if (pointerMoveRaf !== 0) return;
    pointerMoveRaf = requestAnimationFrame(flushPointerMove);
  });

  input.onPointerUp(() => {
    if (activeTableCell) return;
    if (pointerMoveRaf !== 0) {
      cancelAnimationFrame(pointerMoveRaf);
      pointerMoveRaf = 0;
    }
    flushPointerMove();
    pendingPointerMove = null;
    pointerSelecting = false;
    pointerAnchor = null;
    clearPendingFormat();
    fireFormattingChange();
    updateLocalAwareness();
    scheduleRemotePresenceRender();
  });

  input.onSelectionChange(() => {
    if (destroyed || !docEdit) return;
    fireFormattingChange();
    updateImageSelectionOverlay();
    updateLocalAwareness();
    scheduleRemotePresenceRender();
  });

  input.onNavigate((payload: NavigationPayload) => {
    if (destroyed || !docEdit) return;
    if (useNativeSelection) return;

    const selection = selectionOrCursor();
    const range = normalizeSelection(selection);

    if (!payload.shift && !range.collapsed) {
      const collapse = clampPosition(
        collapsePositionForNavigation(range, payload.key),
      );
      renderer.setCursor(collapse);
      if (!renderer.isFocused()) renderer.focus();
      ensureCanvasInputFocus();
      scrollCursorIntoView();
      clearPendingFormat();
      fireFormattingChange();
      updateLocalAwareness();
      scheduleRemotePresenceRender();
      return;
    }

    const movingPos = clampPosition(selection.focus);
    const nextPos = moveForNavigation(movingPos, payload.key, payload);

    if (payload.shift) {
      const anchorPos = clampPosition(selection.anchor);
      if (samePosition(anchorPos, nextPos)) {
        renderer.setCursor(nextPos);
      } else {
        renderer.setSelection({ anchor: anchorPos, focus: nextPos });
      }
    } else {
      renderer.setCursor(nextPos);
    }

    if (!renderer.isFocused()) renderer.focus();
    ensureCanvasInputFocus();
    scrollCursorIntoView();
    clearPendingFormat();
    fireFormattingChange();
    updateLocalAwareness();
    scheduleRemotePresenceRender();
  });

  input.onRequestCopyText(() => selectedPlainText());
  input.onRequestCutText(() => selectedPlainText());
  input.onRequestCopyHtml(() => selectedHtml());
  input.onRequestCutHtml(() => selectedHtml());
  input.onPasteImage((file) => insertPastedImage(file));

  // --- Build controller ---

  const controller: DocEditorController = {
    load(data: Uint8Array): number {
      return loadEditorState(new DocEdit(data), "load");
    },

    loadBlank(): void {
      loadEditorState(DocEdit.blank(), "loadBlank");
    },

    replace(data: Uint8Array): number {
      const copy = data.slice();
      return loadEditorState(new DocEdit(copy), "replace", copy);
    },

    replaceBlank(): void {
      const nextDocEdit = DocEdit.blank();
      const snapshot = nextDocEdit.save();
      loadEditorState(nextDocEdit, "replaceBlank", snapshot);
    },

    save(): Uint8Array {
      if (!docEdit) return new Uint8Array(0);
      const bytes = docEdit.save();
      emitEvent("docedit-dirty", { dirty: false });
      return bytes;
    },

    isDirty(): boolean {
      if (!docEdit) return false;
      return docEdit.isDirty();
    },

    format(action: FormatAction): void {
      if (!docEdit) return;
      const typeMap: Record<FormatAction, string> = {
        bold: "formatBold",
        italic: "formatItalic",
        underline: "formatUnderline",
        strikethrough: "formatStrikethrough",
      };
      const attrMap: Record<FormatAction, string> = {
        bold: "bold",
        italic: "italic",
        underline: "underline",
        strikethrough: "strike",
      };
      applyFormat(typeMap[action], attrMap[action]);
      fireFormattingChange();
    },

    setTextStyle(patch: TextStylePatch): void {
      setTextStyle(patch);
    },

    setParagraphStyle(patch: ParagraphStylePatch): void {
      setParagraphStyle(patch);
    },

    toggleList(kind: ListKind): void {
      toggleList(kind);
    },

    insertInlineImage(payload: InlineImageInsertPayload): boolean {
      return insertInlineImage(payload);
    },

    getSelectedInlineImage(): SelectedInlineImageState | null {
      return getSelectedInlineImage();
    },

    resizeSelectedInlineImage(widthPt: number, heightPt: number): boolean {
      return resizeSelectedInlineImage(widthPt, heightPt);
    },

    setSelectedInlineImageAlignment(alignment: ImageBlockAlignment): boolean {
      return setSelectedInlineImageAlignment(alignment);
    },

    insertTable(rows: number, columns: number): boolean {
      return insertTable(rows, columns);
    },

    insertTableRow(): boolean {
      return insertTableRow();
    },

    removeTableRow(): boolean {
      return removeTableRow();
    },

    insertTableColumn(): boolean {
      return insertTableColumn();
    },

    removeTableColumn(): boolean {
      return removeTableColumn();
    },

    getActiveTableCell(): TableCellState | null {
      return getActiveTableCellState();
    },

    onActiveTableCellChange(cb: (state: TableCellState | null) => void): void {
      tableCellCallbacks.push(cb);
    },

    setActiveTableCellText(text: string): boolean {
      return setActiveTableCellText(text);
    },

    moveActiveTableCell(deltaRow: number, deltaCol: number): boolean {
      return moveActiveTableCell(deltaRow, deltaCol);
    },

    clearActiveTableCell(): void {
      closeTableCellEditor({ commit: true });
    },

    getFormattingState(): FormattingState {
      return queryFormattingState();
    },

    onFormattingChange(cb: (state: FormattingState) => void): void {
      formattingCallbacks.push(cb);
    },

    getSelectionState(): DocSelection | null {
      return renderer.getSelection();
    },

    getDocEdit(): DocEdit | null {
      return docEdit;
    },

    setAwarenessPausedForTests(paused: boolean): void {
      syncProvider?.setAwarenessPausedForTests(paused);
    },

    setTransportPausedForTests(paused: boolean): void {
      syncProvider?.setTransportPausedForTests(paused);
    },

    destroy(): void {
      destroyed = true;
      closeTableCellEditor({ commit: false });
      if (pointerMoveRaf !== 0 && typeof cancelAnimationFrame === "function") {
        cancelAnimationFrame(pointerMoveRaf);
        pointerMoveRaf = 0;
      }
      if (pointerMoveRaf !== 0) {
        pointerMoveRaf = 0;
      }
      if (presenceRaf !== 0 && typeof cancelAnimationFrame === "function") {
        cancelAnimationFrame(presenceRaf);
        presenceRaf = 0;
      }
      if (presenceRaf !== 0) {
        presenceRaf = 0;
      }
      pendingPointerMove = null;
      pointerSelecting = false;
      pointerAnchor = null;
      syncProvider?.destroy();
      syncProvider = null;
      if (typeof container.removeEventListener === "function") {
        container.removeEventListener("scroll", onPresenceViewportChange);
      }
      if (typeof window !== "undefined") {
        window.removeEventListener("resize", onPresenceViewportChange);
        window.removeEventListener("mousemove", onImageResizeMove);
        window.removeEventListener("mouseup", onImageResizeEnd);
      }
      input.destroy();
      renderer.destroy();
      if (docEdit) {
        docEdit.free();
        docEdit = null;
      }
      currentModel = null;
      formattingCallbacks = [];
      remotePeers = [];
      presenceOverlay?.remove();
      presenceOverlay = null;
      if (imageSelectionOverlay) {
        if (typeof imageSelectionOverlay.remove === "function") {
          imageSelectionOverlay.remove();
        } else {
          imageSelectionOverlay.parentNode?.removeChild(imageSelectionOverlay);
        }
      }
      imageSelectionOverlay = null;
      imageSelectionBox = null;
      imageResizeHandle = null;
      imageResizeSession = null;
    },
  };

  return controller;
}
