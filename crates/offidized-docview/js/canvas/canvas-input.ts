// Canvas InputAdapter implementation for offidized-docview.
//
// Since CanvasKit renders to opaque <canvas> elements with no contenteditable
// surface, keyboard input is captured via a hidden <textarea> positioned
// offscreen. Mouse events are captured on the scroll container. This gives us
// full beforeinput/IME/composition/clipboard support without a DOM text layer.

import type {
  InputAdapter,
  NavigationKey,
  NavigationPayload,
  NormalizedInput,
} from "../adapter.ts";

export class CanvasInput implements InputAdapter {
  private static readonly DEFAULT_INPUT_HEIGHT_PX = 20;

  private scrollContainer: HTMLElement;
  private textarea: HTMLTextAreaElement;
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
  private dragging = false;
  private composing = false;
  private compositionText = "";
  private compositionCommitted = false;
  private lastCompositionInputType: string | null = null;
  private suppressedBeforeInputTypes = new Set<string>();
  private suppressNextPasteEvent = false;
  private clearTextareaRaf = 0;
  private positionTextareaRaf = 0;
  private lastTextareaAnchor: {
    left: number;
    top: number;
    height: number;
  } | null = null;

  // Bound listeners for cleanup
  private _onBeforeInput: (e: InputEvent) => void;
  private _onKeyDown: (e: KeyboardEvent) => void;
  private _onKeyUp: (e: KeyboardEvent) => void;
  private _onMouseDown: (e: MouseEvent) => void;
  private _onMouseMove: (e: MouseEvent) => void;
  private _onMouseUp: (e: MouseEvent) => void;
  private _onCompositionStart: () => void;
  private _onCompositionUpdate: (e: CompositionEvent) => void;
  private _onCompositionEnd: (e: CompositionEvent) => void;
  private _onCopy: (e: ClipboardEvent) => void;
  private _onCut: (e: ClipboardEvent) => void;
  private _onPaste: (e: ClipboardEvent) => void;
  private _onFocus: () => void;
  private _onScroll: () => void;

  constructor(scrollContainer: HTMLElement) {
    this.scrollContainer = scrollContainer;

    // Create hidden textarea for keyboard/IME/clipboard capture
    this.textarea = document.createElement("textarea");
    this.textarea.style.position = "absolute";
    this.textarea.style.left = "0";
    this.textarea.style.top = "0";
    this.textarea.style.width = "1px";
    this.textarea.style.height = `${CanvasInput.DEFAULT_INPUT_HEIGHT_PX}px`;
    this.textarea.style.opacity = "0";
    this.textarea.style.overflow = "hidden";
    this.textarea.style.pointerEvents = "none";
    this.textarea.style.margin = "0";
    this.textarea.style.padding = "0";
    this.textarea.style.border = "0";
    this.textarea.style.background = "transparent";
    this.textarea.style.color = "transparent";
    this.textarea.style.caretColor = "transparent";
    this.textarea.style.resize = "none";
    this.textarea.dataset.docviewCanvasInput = "1";
    this.textarea.setAttribute("autocomplete", "off");
    this.textarea.setAttribute("autocorrect", "off");
    this.textarea.setAttribute("autocapitalize", "off");
    this.textarea.setAttribute("spellcheck", "false");
    this.textarea.tabIndex = 0;
    scrollContainer.appendChild(this.textarea);

    // Bind handlers
    this._onBeforeInput = this.handleBeforeInput.bind(this);
    this._onKeyDown = this.handleKeyDown.bind(this);
    this._onKeyUp = this.handleKeyUp.bind(this);
    this._onMouseDown = this.handleMouseDown.bind(this);
    this._onMouseMove = this.handleMouseMove.bind(this);
    this._onMouseUp = this.handleMouseUp.bind(this);
    this._onCompositionStart = this.handleCompositionStart.bind(this);
    this._onCompositionUpdate = this.handleCompositionUpdate.bind(this);
    this._onCompositionEnd = this.handleCompositionEnd.bind(this);
    this._onCopy = this.handleCopy.bind(this);
    this._onCut = this.handleCut.bind(this);
    this._onPaste = this.handlePaste.bind(this);
    this._onFocus = this.handleFocus.bind(this);
    this._onScroll = this.handleScroll.bind(this);

    // Wire textarea events
    this.textarea.addEventListener("beforeinput", this._onBeforeInput);
    this.textarea.addEventListener("keydown", this._onKeyDown);
    this.textarea.addEventListener("keyup", this._onKeyUp);
    this.textarea.addEventListener(
      "compositionstart",
      this._onCompositionStart,
    );
    this.textarea.addEventListener(
      "compositionupdate",
      this._onCompositionUpdate,
    );
    this.textarea.addEventListener("compositionend", this._onCompositionEnd);
    this.textarea.addEventListener("copy", this._onCopy);
    this.textarea.addEventListener("cut", this._onCut);
    this.textarea.addEventListener("paste", this._onPaste);
    this.textarea.addEventListener("focus", this._onFocus);

    // Wire scroll container mouse events
    this.scrollContainer.addEventListener("mousedown", this._onMouseDown);
    this.scrollContainer.addEventListener("scroll", this._onScroll);
    window.addEventListener("mousemove", this._onMouseMove);
    window.addEventListener("mouseup", this._onMouseUp);

    this.scheduleTextareaPosition();
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

  /** Returns the hidden textarea (useful for the renderer to focus). */
  getInputElement(): HTMLElement {
    return this.textarea;
  }

  destroy(): void {
    this.textarea.removeEventListener("beforeinput", this._onBeforeInput);
    this.textarea.removeEventListener("keydown", this._onKeyDown);
    this.textarea.removeEventListener("keyup", this._onKeyUp);
    this.textarea.removeEventListener(
      "compositionstart",
      this._onCompositionStart,
    );
    this.textarea.removeEventListener(
      "compositionupdate",
      this._onCompositionUpdate,
    );
    this.textarea.removeEventListener("compositionend", this._onCompositionEnd);
    this.textarea.removeEventListener("copy", this._onCopy);
    this.textarea.removeEventListener("cut", this._onCut);
    this.textarea.removeEventListener("paste", this._onPaste);
    this.textarea.removeEventListener("focus", this._onFocus);
    this.scrollContainer.removeEventListener("mousedown", this._onMouseDown);
    this.scrollContainer.removeEventListener("scroll", this._onScroll);
    window.removeEventListener("mousemove", this._onMouseMove);
    window.removeEventListener("mouseup", this._onMouseUp);

    if (this.textarea.parentNode) {
      this.textarea.parentNode.removeChild(this.textarea);
    }

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
    this.dragging = false;
    this.composing = false;
    this.compositionText = "";
    this.compositionCommitted = false;
    this.lastCompositionInputType = null;
    this.suppressedBeforeInputTypes.clear();
    this.suppressNextPasteEvent = false;
    this.cancelScheduledClear();
    this.cancelScheduledTextareaPosition();
  }

  // --- Private event handlers ---

  private handleBeforeInput(e: InputEvent): void {
    const mapped = this.mapInputType(e);
    if (!mapped) return;

    e.preventDefault();
    this.inputHandler?.(mapped);
    this.scheduleTextareaPosition();
    if (mapped.type === "insertFromPaste") {
      this.suppressNextPasteEvent = true;
      requestAnimationFrame(() => {
        this.suppressNextPasteEvent = false;
      });
    }

    if (!this.shouldPreserveTextareaValue(e.inputType)) {
      this.clearTextarea();
    }
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
      this.scheduleTextareaPosition();
      return;
    }
    if (this.isWordDeleteKey(e, "Delete")) {
      e.preventDefault();
      this.suppressBeforeInputType("deleteWordForward");
      this.inputHandler?.({ type: "deleteWordForward" });
      this.scheduleTextareaPosition();
      return;
    }

    // Tab key — no meta key required
    if (e.key === "Tab") {
      e.preventDefault();
      this.suppressBeforeInputType("insertText");
      this.inputHandler?.({ type: "insertTab", shift: e.shiftKey });
      this.scheduleTextareaPosition();
      return;
    }

    // Keep editing functional even when beforeinput is not emitted reliably.
    if (e.key === "Backspace") {
      e.preventDefault();
      this.suppressBeforeInputType("deleteContentBackward");
      this.inputHandler?.({ type: "deleteContentBackward" });
      this.scheduleTextareaPosition();
      return;
    }
    if (e.key === "Delete") {
      e.preventDefault();
      this.suppressBeforeInputType("deleteContentForward");
      this.inputHandler?.({ type: "deleteContentForward" });
      this.scheduleTextareaPosition();
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
      this.scheduleTextareaPosition();
      return;
    }

    if (this.isNavigationKey(e.key)) {
      e.preventDefault();
      this.navigateHandler?.({
        key: e.key,
        shift: e.shiftKey,
        meta: e.metaKey,
        alt: e.altKey,
        ctrl: e.ctrlKey,
      });
      this.scheduleTextareaPosition();
      return;
    }

    if (!this.isMetaKey(e)) return;

    const key = e.key.toLowerCase();
    switch (key) {
      case "b":
      case "i":
      case "u":
        e.preventDefault();
        this.shortcutHandler?.(key, e.shiftKey);
        this.scheduleTextareaPosition();
        break;
      case "5":
        if (e.shiftKey) {
          e.preventDefault();
          this.shortcutHandler?.(key, e.shiftKey);
          this.scheduleTextareaPosition();
        }
        break;
      case "z":
        e.preventDefault();
        this.emitHistoryFallback(e.shiftKey ? "historyRedo" : "historyUndo");
        break;
      case "y":
        e.preventDefault();
        this.emitHistoryFallback("historyRedo");
        break;
      case "a":
        e.preventDefault();
        this.shortcutHandler?.("a", e.shiftKey);
        this.scheduleTextareaPosition();
        break;
      case "s":
        e.preventDefault();
        this.shortcutHandler?.("s", e.shiftKey);
        break;
      default:
        break;
    }
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
      this.selectionChangeHandler?.();
      this.scheduleTextareaPosition();
    }
  }

  private handleMouseDown(e: MouseEvent): void {
    if (e.button !== 0) return;
    e.preventDefault();
    this.dragging = true;
    this.pointerDownHandler?.({
      x: e.clientX,
      y: e.clientY,
      clickCount: Math.max(1, e.detail || 1),
      shift: e.shiftKey,
    });

    const activeElement =
      typeof document !== "undefined" ? document.activeElement : null;
    if (
      activeElement instanceof HTMLTextAreaElement &&
      activeElement.dataset.docviewTableCellEditor === "1" &&
      activeElement.style.display !== "none"
    ) {
      return;
    }

    // Focus the hidden textarea so subsequent keystrokes are captured
    this.textarea.focus({ preventScroll: true });
    this.scheduleTextareaPosition();
  }

  private handleMouseMove(e: MouseEvent): void {
    if (!this.dragging) return;
    // Some synthetic/headless mousemove events report buttons=0 even during
    // an active drag. Only auto-finish for trusted events.
    if (e.isTrusted && (e.buttons & 1) === 0) {
      this.finishDrag(e.clientX, e.clientY);
      return;
    }
    this.pointerMoveHandler?.(e.clientX, e.clientY);
    this.scheduleTextareaPosition();
  }

  private handleMouseUp(e: MouseEvent): void {
    if (e.button !== 0) return;
    if (!this.dragging) return;
    this.finishDrag(e.clientX, e.clientY);
  }

  private handleCompositionStart(): void {
    this.cancelScheduledClear();
    this.composing = true;
    this.compositionText = this.textarea.value;
    this.compositionCommitted = false;
    this.lastCompositionInputType = null;
    this.scheduleTextareaPosition();
  }

  private handleCompositionUpdate(e: CompositionEvent): void {
    if (!this.composing) return;
    this.compositionText = this.resolveCompositionUpdateText(e.data);
    this.scheduleTextareaPosition();
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
    this.clearTextarea();
    this.scheduleTextareaPosition();
  }

  private handleCopy(e: ClipboardEvent): void {
    const text = this.requestCopyTextHandler?.() ?? "";
    e.preventDefault();
    const html = this.requestCopyHtmlHandler?.();
    if (html) {
      e.clipboardData?.setData("text/html", html);
    }
    e.clipboardData?.setData("text/plain", text);
  }

  private handleCut(e: ClipboardEvent): void {
    const text = this.requestCutTextHandler?.() ?? "";
    e.preventDefault();
    const html = this.requestCutHtmlHandler?.();
    if (html) {
      e.clipboardData?.setData("text/html", html);
    }
    e.clipboardData?.setData("text/plain", text);
    if (text) {
      this.inputHandler?.({ type: "deleteByCut" });
    }
  }

  private handlePaste(e: ClipboardEvent): void {
    const imageFile = this.extractPastedImageFile(e.clipboardData);
    if (imageFile) {
      e.preventDefault();
      void this.pasteImageHandler?.(imageFile);
      return;
    }
    if (this.suppressNextPasteEvent) {
      this.suppressNextPasteEvent = false;
      return;
    }
    const text = e.clipboardData?.getData("text/plain") ?? "";
    const html = e.clipboardData?.getData("text/html") ?? "";
    if (!text && !html) return;
    e.preventDefault();
    this.suppressBeforeInputType("insertFromPaste");
    this.inputHandler?.({ type: "insertFromPaste", data: text, html });
    this.clearTextarea();
    this.scheduleTextareaPosition();
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

  private finishDrag(x: number, y: number): void {
    this.pointerMoveHandler?.(x, y);
    this.dragging = false;
    this.pointerUpHandler?.();
    this.selectionChangeHandler?.();
    this.scheduleTextareaPosition();
  }

  private clearTextarea(): void {
    this.cancelScheduledClear();
    if (this.composing) return;
    // Defer clearing so the browser finishes processing the input event.
    this.clearTextareaRaf = requestAnimationFrame(() => {
      this.clearTextareaRaf = 0;
      if (this.composing) return;
      this.textarea.value = "";
    });
  }

  private cancelScheduledClear(): void {
    if (this.clearTextareaRaf === 0) return;
    cancelAnimationFrame(this.clearTextareaRaf);
    this.clearTextareaRaf = 0;
  }

  private handleFocus(): void {
    this.scheduleTextareaPosition();
  }

  private handleScroll(): void {
    this.scheduleTextareaPosition();
  }

  private scheduleTextareaPosition(): void {
    if (this.positionTextareaRaf !== 0) return;
    this.positionTextareaRaf = requestAnimationFrame(() => {
      this.positionTextareaRaf = 0;
      this.positionTextarea();
    });
  }

  private cancelScheduledTextareaPosition(): void {
    if (this.positionTextareaRaf === 0) return;
    cancelAnimationFrame(this.positionTextareaRaf);
    this.positionTextareaRaf = 0;
  }

  private positionTextarea(): void {
    const anchor = this.resolveTextareaAnchor();
    const left = Math.max(0, anchor.left);
    const top = Math.max(0, anchor.top);
    const height = Math.max(1, anchor.height);

    this.textarea.style.left = `${left}px`;
    this.textarea.style.top = `${top}px`;
    this.textarea.style.height = `${height}px`;
    this.textarea.style.fontSize = `${Math.max(16, Math.round(height))}px`;
    this.textarea.style.lineHeight = `${height}px`;

    this.lastTextareaAnchor = { left, top, height };
  }

  private resolveTextareaAnchor(): {
    left: number;
    top: number;
    height: number;
  } {
    const caretOverlay = this.findCaretOverlay();
    if (caretOverlay && caretOverlay.style.display !== "none") {
      const left = this.parsePx(caretOverlay.style.left);
      const top = this.parsePx(caretOverlay.style.top);
      const height = this.parsePx(
        caretOverlay.style.height,
        CanvasInput.DEFAULT_INPUT_HEIGHT_PX,
      );
      if (Number.isFinite(left) && Number.isFinite(top)) {
        return { left, top, height };
      }
    }

    if (this.lastTextareaAnchor) {
      return this.lastTextareaAnchor;
    }

    return {
      left: this.scrollContainer.scrollLeft || 0,
      top: this.scrollContainer.scrollTop || 0,
      height: CanvasInput.DEFAULT_INPUT_HEIGHT_PX,
    };
  }

  private findCaretOverlay(): HTMLElement | null {
    const queryable = this.scrollContainer as HTMLElement & {
      querySelector?: (selector: string) => HTMLElement | null;
      children?: ArrayLike<Element>;
    };
    if (typeof queryable.querySelector === "function") {
      const overlay = queryable.querySelector(
        '[data-docview-caret-overlay="1"]',
      );
      if (overlay) return overlay as HTMLElement;
    }

    const children = queryable.children;
    if (!children) return null;
    for (let i = 0; i < children.length; i += 1) {
      const child = children[i] as HTMLElement & {
        dataset?: DOMStringMap;
      };
      if (child?.dataset?.docviewCaretOverlay === "1") {
        return child;
      }
    }
    return null;
  }

  private parsePx(value: string | null | undefined, fallback = 0): number {
    if (typeof value !== "string" || value.length === 0) {
      return fallback;
    }
    const parsed = Number.parseFloat(value);
    return Number.isFinite(parsed) ? parsed : fallback;
  }

  private isMetaKey(e: KeyboardEvent): boolean {
    return e.metaKey || e.ctrlKey;
  }

  private isWordDeleteKey(
    e: KeyboardEvent,
    key: "Backspace" | "Delete",
  ): boolean {
    return e.key === key && (e.altKey || (e.ctrlKey && !e.metaKey));
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

  private shouldPreserveTextareaValue(inputType: string): boolean {
    return (
      inputType === "insertCompositionText" ||
      inputType === "insertFromComposition" ||
      inputType === "deleteCompositionText" ||
      this.composing
    );
  }

  private resolveCompositionUpdateText(
    data: string | null | undefined,
  ): string {
    return data ?? this.textarea.value ?? this.compositionText;
  }

  private resolveCommittedCompositionText(
    data: string | null | undefined,
  ): string {
    return this.firstNonEmptyString(
      data,
      this.textarea.value,
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
        this.textarea.value,
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

  private emitHistoryFallback(type: "historyUndo" | "historyRedo"): void {
    this.suppressBeforeInputType(type);
    this.inputHandler?.({ type });
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

  private suppressBeforeInputType(inputType: string): void {
    this.suppressedBeforeInputTypes.add(inputType);
    requestAnimationFrame(() => {
      this.suppressedBeforeInputTypes.delete(inputType);
    });
  }
}
