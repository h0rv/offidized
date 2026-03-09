// HTML (contenteditable) InputAdapter implementation for offidized-docview.
//
// Wraps beforeinput/keydown/mousedown/selectionchange events on the
// contenteditable surface and normalizes them into InputAdapter callbacks.

import type {
  InputAdapter,
  NavigationKey,
  NavigationPayload,
  NormalizedInput,
} from "../adapter.ts";

export class HtmlInput implements InputAdapter {
  private surface: HTMLElement;
  private inputHandler: ((input: NormalizedInput) => void) | null = null;
  private shortcutHandler: ((key: string, shift: boolean) => void) | null =
    null;
  private navigateHandler: ((payload: NavigationPayload) => void) | null = null;
  private requestCopyTextHandler: (() => string | null | undefined) | null =
    null;
  private requestCutTextHandler: (() => string | null | undefined) | null =
    null;
  private requestCopyHtmlHandler: (() => string | null | undefined) | null =
    null;
  private requestCutHtmlHandler: (() => string | null | undefined) | null =
    null;
  private pasteImageHandler: ((file: File) => void | Promise<void>) | null =
    null;
  private pointerDownHandler:
    | ((payload: {
        x: number;
        y: number;
        clickCount: number;
        shift: boolean;
      }) => void)
    | null = null;
  private pointerMoveHandler: ((x: number, y: number) => void) | null = null;
  private pointerUpHandler: (() => void) | null = null;
  private selectionChangeHandler: (() => void) | null = null;
  private suppressedBeforeInputTypes = new Set<string>();
  private composing = false;
  private compositionText = "";
  private compositionCommitted = false;
  private lastCompositionInputType: string | null = null;

  // Bound listeners for cleanup
  private _onBeforeInput: (e: InputEvent) => void;
  private _onKeyDown: (e: KeyboardEvent) => void;
  private _onMouseDown: (e: MouseEvent) => void;
  private _onMouseMove: (e: MouseEvent) => void;
  private _onMouseUp: () => void;
  private _onKeyUp: (e: KeyboardEvent) => void;
  private _onSelectionChange: () => void;
  private _onCompositionStart: () => void;
  private _onCompositionUpdate: (e: CompositionEvent) => void;
  private _onCompositionEnd: (e: CompositionEvent) => void;
  private _onCopy: (e: ClipboardEvent) => void;
  private _onCut: (e: ClipboardEvent) => void;
  private _onPaste: (e: ClipboardEvent) => void;

  constructor(surface: HTMLElement) {
    this.surface = surface;

    this._onBeforeInput = this.handleBeforeInput.bind(this);
    this._onKeyDown = this.handleKeyDown.bind(this);
    this._onMouseDown = this.handleMouseDown.bind(this);
    this._onMouseMove = this.handleMouseMove.bind(this);
    this._onMouseUp = this.handleMouseUp.bind(this);
    this._onKeyUp = this.handleKeyUp.bind(this);
    this._onSelectionChange = this.handleSelectionChange.bind(this);
    this._onCompositionStart = this.handleCompositionStart.bind(this);
    this._onCompositionUpdate = this.handleCompositionUpdate.bind(this);
    this._onCompositionEnd = this.handleCompositionEnd.bind(this);
    this._onCopy = this.handleCopy.bind(this);
    this._onCut = this.handleCut.bind(this);
    this._onPaste = this.handlePaste.bind(this);

    surface.addEventListener("beforeinput", this._onBeforeInput);
    surface.addEventListener("keydown", this._onKeyDown);
    surface.addEventListener("mousedown", this._onMouseDown);
    surface.addEventListener("mousemove", this._onMouseMove);
    surface.addEventListener("mouseup", this._onMouseUp);
    surface.addEventListener("keyup", this._onKeyUp);
    surface.addEventListener("compositionstart", this._onCompositionStart);
    surface.addEventListener("compositionupdate", this._onCompositionUpdate);
    surface.addEventListener("compositionend", this._onCompositionEnd);
    surface.addEventListener("copy", this._onCopy);
    surface.addEventListener("cut", this._onCut);
    surface.addEventListener("paste", this._onPaste);
    document.addEventListener("selectionchange", this._onSelectionChange);
  }

  onInput(handler: (input: NormalizedInput) => void): void {
    this.inputHandler = handler;
  }

  onShortcut(handler: (key: string, shift: boolean) => void): void {
    this.shortcutHandler = handler;
  }

  onNavigate(handler: (payload: NavigationPayload) => void): void {
    this.navigateHandler = handler;
  }

  onRequestCopyText(handler: () => string | null | undefined): void {
    this.requestCopyTextHandler = handler;
  }

  onRequestCutText(handler: () => string | null | undefined): void {
    this.requestCutTextHandler = handler;
  }

  onRequestCopyHtml(handler: () => string | null | undefined): void {
    this.requestCopyHtmlHandler = handler;
  }

  onRequestCutHtml(handler: () => string | null | undefined): void {
    this.requestCutHtmlHandler = handler;
  }

  onPointerDown(
    handler: (payload: {
      x: number;
      y: number;
      clickCount: number;
      shift: boolean;
    }) => void,
  ): void {
    this.pointerDownHandler = handler;
  }

  onPointerMove(handler: (x: number, y: number) => void): void {
    this.pointerMoveHandler = handler;
  }

  onPointerUp(handler: () => void): void {
    this.pointerUpHandler = handler;
  }

  onSelectionChange(handler: () => void): void {
    this.selectionChangeHandler = handler;
  }

  onPasteImage(handler: (file: File) => void | Promise<void>): void {
    this.pasteImageHandler = handler;
  }

  destroy(): void {
    this.surface.removeEventListener("beforeinput", this._onBeforeInput);
    this.surface.removeEventListener("keydown", this._onKeyDown);
    this.surface.removeEventListener("mousedown", this._onMouseDown);
    this.surface.removeEventListener("mousemove", this._onMouseMove);
    this.surface.removeEventListener("mouseup", this._onMouseUp);
    this.surface.removeEventListener("keyup", this._onKeyUp);
    this.surface.removeEventListener(
      "compositionstart",
      this._onCompositionStart,
    );
    this.surface.removeEventListener(
      "compositionupdate",
      this._onCompositionUpdate,
    );
    this.surface.removeEventListener("compositionend", this._onCompositionEnd);
    this.surface.removeEventListener("copy", this._onCopy);
    this.surface.removeEventListener("cut", this._onCut);
    this.surface.removeEventListener("paste", this._onPaste);
    document.removeEventListener("selectionchange", this._onSelectionChange);
    this.inputHandler = null;
    this.shortcutHandler = null;
    this.navigateHandler = null;
    this.requestCopyTextHandler = null;
    this.requestCutTextHandler = null;
    this.requestCopyHtmlHandler = null;
    this.requestCutHtmlHandler = null;
    this.pasteImageHandler = null;
    this.pointerDownHandler = null;
    this.pointerMoveHandler = null;
    this.pointerUpHandler = null;
    this.selectionChangeHandler = null;
    this.suppressedBeforeInputTypes.clear();
    this.composing = false;
    this.compositionText = "";
    this.compositionCommitted = false;
    this.lastCompositionInputType = null;
  }

  // --- Private event handlers ---

  private handleBeforeInput(e: InputEvent): void {
    const mapped = this.mapInputType(e);
    if (!mapped) return;
    e.preventDefault();
    this.inputHandler?.(mapped);
  }

  private mapInputType(e: InputEvent): NormalizedInput | null {
    if (this.suppressedBeforeInputTypes.delete(e.inputType)) {
      return null;
    }

    switch (e.inputType) {
      case "insertText":
        if (this.isCompositionInputEvent(e)) return null;
        return e.data ? { type: "insertText", data: e.data } : null;
      case "deleteContentBackward":
        return { type: "deleteContentBackward" };
      case "deleteContentForward":
        return { type: "deleteContentForward" };
      case "deleteWordBackward":
        return { type: "deleteWordBackward" };
      case "deleteWordForward":
        return { type: "deleteWordForward" };
      case "insertParagraph":
        return { type: "insertParagraph" };
      case "insertLineBreak":
        return { type: "insertLineBreak" };
      case "insertFromPaste": {
        const text = e.dataTransfer?.getData("text/plain") ?? e.data ?? "";
        const html = e.dataTransfer?.getData("text/html") ?? "";
        return text || html
          ? { type: "insertFromPaste", data: text, html }
          : null;
      }
      case "insertCompositionText":
        this.lastCompositionInputType = e.inputType;
        this.compositionText = this.resolveCompositionUpdateText(e.data);
        return null;
      case "insertFromComposition": {
        this.lastCompositionInputType = e.inputType;
        const text = this.resolveCommittedCompositionText(e.data);
        this.compositionCommitted = text.length > 0;
        return text ? { type: "insertFromComposition", data: text } : null;
      }
      case "deleteCompositionText":
      case "deleteByComposition":
        this.lastCompositionInputType = e.inputType;
        this.compositionText = "";
        return null;
      case "deleteByCut":
        return { type: "deleteByCut" };
      case "historyUndo":
        return { type: "historyUndo" };
      case "historyRedo":
        return { type: "historyRedo" };
      default:
        return null;
    }
  }

  private handleKeyDown(e: KeyboardEvent): void {
    if (this.isCompositionKeyEvent(e)) return;

    if (this.isWordDeleteKey(e, "Backspace")) {
      e.preventDefault();
      this.suppressBeforeInputType("deleteWordBackward");
      this.inputHandler?.({ type: "deleteWordBackward" });
      return;
    }
    if (this.isWordDeleteKey(e, "Delete")) {
      e.preventDefault();
      this.suppressBeforeInputType("deleteWordForward");
      this.inputHandler?.({ type: "deleteWordForward" });
      return;
    }

    if (e.key === "Backspace") {
      e.preventDefault();
      this.suppressBeforeInputType("deleteContentBackward");
      this.inputHandler?.({ type: "deleteContentBackward" });
      return;
    }
    if (e.key === "Delete") {
      e.preventDefault();
      this.suppressBeforeInputType("deleteContentForward");
      this.inputHandler?.({ type: "deleteContentForward" });
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      this.suppressBeforeInputType(
        e.shiftKey ? "insertLineBreak" : "insertParagraph",
      );
      this.inputHandler?.({
        type: e.shiftKey ? "insertLineBreak" : "insertParagraph",
      });
      return;
    }

    // Tab key — no meta key required
    if (e.key === "Tab") {
      e.preventDefault();
      this.inputHandler?.({ type: "insertTab", shift: e.shiftKey });
      return;
    }

    if (this.isNavigationKey(e.key)) {
      this.navigateHandler?.({
        key: e.key,
        shift: e.shiftKey,
        meta: e.metaKey,
        alt: e.altKey,
        ctrl: e.ctrlKey,
      });
    }

    if (!this.isMetaKey(e)) return;

    const key = e.key.toLowerCase();
    switch (key) {
      case "b":
      case "i":
      case "u":
        e.preventDefault();
        this.shortcutHandler?.(key, e.shiftKey);
        break;
      case "5":
        if (e.shiftKey) {
          e.preventDefault();
          this.shortcutHandler?.(key, e.shiftKey);
        }
        break;
      case "z":
        e.preventDefault();
        if (e.shiftKey) {
          this.inputHandler?.({ type: "historyRedo" });
        } else {
          this.inputHandler?.({ type: "historyUndo" });
        }
        break;
      case "y":
        e.preventDefault();
        this.inputHandler?.({ type: "historyRedo" });
        break;
      case "s":
        e.preventDefault();
        this.shortcutHandler?.("s", e.shiftKey);
        break;
      case "a":
        e.preventDefault();
        this.shortcutHandler?.("a", e.shiftKey);
        break;
      default:
        break;
    }
  }

  private handleMouseDown(e: MouseEvent): void {
    const target = e.target as HTMLElement;
    if (!target) return;

    const tableCell = target.closest?.("[data-docview-table-cell='1']");
    if (tableCell) {
      // Prevent native contenteditable focus/selection from racing the overlay
      // textarea when the controller opens a table cell editor for this click.
      e.preventDefault();
    }

    // Clicked on a page gap — redirect to nearest paragraph
    if (target.classList.contains("docedit-page-gap")) {
      e.preventDefault();
      const prevPage = target.previousElementSibling;
      if (prevPage) {
        const lastPara = this.lastBodyItemIn(prevPage);
        if (lastPara) {
          const sel = window.getSelection();
          if (sel) {
            sel.removeAllRanges();
            const range = document.createRange();
            range.selectNodeContents(lastPara);
            range.collapse(false);
            sel.addRange(range);
          }
        }
      }
      return;
    }

    // Clicked on empty page whitespace
    if (target.classList.contains("docedit-page")) {
      requestAnimationFrame(() => {
        const sel = window.getSelection();
        if (!sel || !sel.anchorNode) return;
        if (sel.anchorNode === target || sel.anchorNode === this.surface) {
          const lastPara = this.lastBodyItemIn(target);
          if (lastPara) {
            sel.removeAllRanges();
            const range = document.createRange();
            range.selectNodeContents(lastPara);
            range.collapse(false);
            sel.addRange(range);
          }
        }
      });
    }

    this.pointerDownHandler?.({
      x: e.clientX,
      y: e.clientY,
      clickCount: Math.max(1, e.detail || 1),
      shift: e.shiftKey,
    });
  }

  private handleMouseUp(): void {
    this.pointerUpHandler?.();
  }

  private handleMouseMove(e: MouseEvent): void {
    if ((e.buttons & 1) === 0) return;
    this.pointerMoveHandler?.(e.clientX, e.clientY);
  }

  private handleKeyUp(e: KeyboardEvent): void {
    if (this.isCompositionKeyEvent(e)) return;
    // Arrow keys, Home, End, PageUp, PageDown — notify selection change
    if (
      e.key.startsWith("Arrow") ||
      e.key === "Home" ||
      e.key === "End" ||
      e.key === "PageUp" ||
      e.key === "PageDown"
    ) {
      this.pointerUpHandler?.();
    }
  }

  private handleCompositionStart(): void {
    this.composing = true;
    this.compositionText = "";
    this.compositionCommitted = false;
    this.lastCompositionInputType = null;
  }

  private handleCompositionUpdate(e: CompositionEvent): void {
    if (!this.composing) return;
    this.compositionText = this.resolveCompositionUpdateText(e.data);
  }

  private handleCompositionEnd(e: CompositionEvent): void {
    const committedText =
      this.lastCompositionInputType === "deleteCompositionText" ||
      this.lastCompositionInputType === "deleteByComposition"
        ? ""
        : this.resolveCompositionEndText(e.data);
    const shouldCommit = !this.compositionCommitted && committedText.length > 0;
    this.composing = false;
    this.compositionText = "";
    this.compositionCommitted = false;
    this.lastCompositionInputType = null;
    if (shouldCommit) {
      this.inputHandler?.({
        type: "insertFromComposition",
        data: committedText,
      });
    }
  }

  private handleSelectionChange(): void {
    const sel = window.getSelection();
    if (!sel || !sel.anchorNode) return;
    if (!this.surface.contains(sel.anchorNode)) return;
    this.selectionChangeHandler?.();
  }

  private handleCopy(e: ClipboardEvent): void {
    const text = this.requestCopyTextHandler?.();
    if (text == null) return;
    e.preventDefault();
    const html = this.requestCopyHtmlHandler?.();
    if (html) {
      e.clipboardData?.setData("text/html", html);
    }
    e.clipboardData?.setData("text/plain", text);
  }

  private handleCut(e: ClipboardEvent): void {
    const text = this.requestCutTextHandler?.();
    if (text == null) return;
    e.preventDefault();
    const html = this.requestCutHtmlHandler?.();
    if (html) {
      e.clipboardData?.setData("text/html", html);
    }
    e.clipboardData?.setData("text/plain", text);
    if (!text) return;
    this.suppressBeforeInputType("deleteByCut");
    this.inputHandler?.({ type: "deleteByCut" });
  }

  private handlePaste(e: ClipboardEvent): void {
    const imageFile = this.extractPastedImageFile(e.clipboardData);
    if (imageFile) {
      e.preventDefault();
      void this.pasteImageHandler?.(imageFile);
      return;
    }
    const text = e.clipboardData?.getData("text/plain") ?? "";
    const html = e.clipboardData?.getData("text/html") ?? "";
    if (!text && !html) return;
    e.preventDefault();
    this.suppressBeforeInputType("insertFromPaste");
    this.inputHandler?.({ type: "insertFromPaste", data: text, html });
  }

  private lastBodyItemIn(root: Element): HTMLElement | null {
    const bodyItems = root.querySelectorAll("[data-body-index]");
    const last = bodyItems[bodyItems.length - 1];
    return last instanceof HTMLElement ? last : null;
  }

  private suppressBeforeInputType(inputType: string): void {
    this.suppressedBeforeInputTypes.add(inputType);
  }

  private extractPastedImageFile(data: DataTransfer | null): File | null {
    if (!data) return null;
    const items = Array.from(data.items ?? []);
    for (const item of items) {
      if (item.kind !== "file" || !item.type.startsWith("image/")) continue;
      const file = item.getAsFile();
      if (file) return file;
    }
    const files = Array.from(data.files ?? []);
    return files.find((file) => file.type.startsWith("image/")) ?? null;
  }

  private isMetaKey(e: KeyboardEvent): boolean {
    return e.metaKey || e.ctrlKey;
  }

  private isCompositionInputEvent(e: InputEvent): boolean {
    return this.composing || e.isComposing;
  }

  private isCompositionKeyEvent(e: KeyboardEvent): boolean {
    const keyCode = (
      e as KeyboardEvent & {
        keyCode?: number;
      }
    ).keyCode;
    return (
      this.composing ||
      e.isComposing ||
      e.key === "Process" ||
      e.key === "Dead" ||
      keyCode === 229
    );
  }

  private resolveCompositionUpdateText(
    data: string | null | undefined,
  ): string {
    const currentText = this.surface.textContent ?? "";
    return data ?? currentText ?? this.compositionText;
  }

  private resolveCommittedCompositionText(
    data: string | null | undefined,
  ): string {
    return this.firstNonEmptyString(
      data,
      this.surface.textContent ?? "",
      this.compositionText,
    );
  }

  private resolveCompositionEndText(data: string | null | undefined): string {
    if (typeof data === "string" && data.length > 0) {
      return data;
    }
    if (this.lastCompositionInputType === "insertFromComposition") {
      return this.firstNonEmptyString(
        data,
        this.surface.textContent ?? "",
        this.compositionText,
      );
    }
    return "";
  }

  private firstNonEmptyString(
    ...values: Array<string | null | undefined>
  ): string {
    for (const value of values) {
      if (typeof value === "string" && value.length > 0) {
        return value;
      }
    }
    return "";
  }

  private isWordDeleteKey(
    e: KeyboardEvent,
    key: "Backspace" | "Delete",
  ): boolean {
    return e.key === key && (e.altKey || (e.ctrlKey && !e.metaKey));
  }

  private isNavigationKey(key: string): key is NavigationKey {
    return (
      key === "ArrowLeft" ||
      key === "ArrowRight" ||
      key === "ArrowUp" ||
      key === "ArrowDown" ||
      key === "Home" ||
      key === "End" ||
      key === "PageUp" ||
      key === "PageDown"
    );
  }
}
