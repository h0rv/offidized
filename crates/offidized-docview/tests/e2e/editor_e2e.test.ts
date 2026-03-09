// Headless E2E tests for the WASM document editor.
//
// These run in Bun — no browser, no DOM. DocEdit is pure WASM.
// The WASM module is loaded synchronously from disk via initSync().

import { describe, test, expect } from "bun:test";
import { createEditorController } from "../../js/editor.ts";
import type {
  DocEditorController,
  DocSelection,
  InputAdapter,
  PointerDownPayload,
  NormalizedInput,
  RendererAdapter,
} from "../../js/adapter.ts";
import {
  DocEdit,
  vm,
  firstParaText,
  applyAndVm,
  paraText,
  paraRuns,
  paraModel,
} from "./helpers.ts";
import type { DocViewModel, VmParagraph, VmRun } from "./helpers.ts";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Encode a CRDT position (body_index, char_offset) → base64 string. */
function pos(bodyIndex: number, charOffset: number): string {
  return DocEdit.encodePosition(bodyIndex, charOffset);
}

/** Get runs from the first paragraph. */
function firstRuns(model: DocViewModel): VmRun[] {
  const para = model.body[0];
  return (para?.runs ?? []) as VmRun[];
}

function formattingAt(
  editor: DocEdit,
  bodyIndex: number,
  charOffset: number,
): Record<string, unknown> {
  return (editor.formattingAt(pos(bodyIndex, charOffset)) ?? {}) as Record<
    string,
    unknown
  >;
}

function expectDecimalListParagraph(
  paragraph: VmParagraph | undefined,
  expectedText: string,
  expectedOrdinal: number,
  expectedNumId: number,
): void {
  const text = (paragraph?.runs ?? []).map((run) => run.text ?? "").join("");
  expect(text).toBe(expectedText);
  expect(paragraph?.numbering?.format).toBe("decimal");
  expect(paragraph?.numbering?.numId).toBe(expectedNumId);
  expect(paragraph?.numbering?.level).toBe(0);
  expect(paragraph?.numbering?.text).toBe(`${expectedOrdinal}.`);
}

function decimalListParagraphs(editor: DocEdit): VmParagraph[] {
  return vm(editor).body.filter((item): item is VmParagraph => {
    const paragraph = item as VmParagraph;
    return (
      paragraph.type === "paragraph" &&
      paragraph.numbering?.format === "decimal"
    );
  });
}

function numberedParagraphs(editor: DocEdit): VmParagraph[] {
  return vm(editor).body.filter((item): item is VmParagraph => {
    const paragraph = item as VmParagraph;
    return paragraph.type === "paragraph" && paragraph.numbering != null;
  });
}

function syncEditors(source: DocEdit, peer: DocEdit): void {
  peer.applyUpdate(source.encodeDiff(peer.encodeStateVector()));
}

function fullResyncEditors(left: DocEdit, right: DocEdit): void {
  right.applyUpdate(left.encodeStateAsUpdate());
  left.applyUpdate(right.encodeStateAsUpdate());
}

const TINY_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aXioAAAAASUVORK5CYII=";

if (typeof CustomEvent === "undefined") {
  class TestCustomEvent<T = unknown> {
    type: string;
    detail: T | undefined;
    bubbles: boolean;

    constructor(type: string, init?: { detail?: T; bubbles?: boolean }) {
      this.type = type;
      this.detail = init?.detail;
      this.bubbles = init?.bubbles ?? false;
    }
  }

  (globalThis as { CustomEvent?: unknown }).CustomEvent = TestCustomEvent;
}

if (typeof document === "undefined") {
  const createMockElement = (): HTMLElement => {
    const listeners = new Map<string, Array<(event: Event) => void>>();
    const element = {
      style: {},
      dataset: {},
      value: "",
      spellcheck: false,
      addEventListener: (type: string, handler: (event: Event) => void) => {
        const handlers = listeners.get(type) ?? [];
        handlers.push(handler);
        listeners.set(type, handlers);
      },
      removeEventListener: (type: string, handler: (event: Event) => void) => {
        const handlers = listeners.get(type);
        if (!handlers) return;
        listeners.set(
          type,
          handlers.filter((item) => item !== handler),
        );
      },
      dispatchEvent: (event: Event) => {
        for (const handler of listeners.get(event.type) ?? []) {
          handler(event);
        }
        return !event.defaultPrevented;
      },
      focus: () => undefined,
      select: () => undefined,
      appendChild: () => undefined,
    };
    return element as unknown as HTMLElement;
  };

  (globalThis as { document?: unknown }).document = {
    activeElement: null,
    createElement: () => createMockElement(),
  };
}

class FakeRenderer implements RendererAdapter {
  model: DocViewModel | null = null;
  selection: DocSelection | null = null;
  focused = false;
  nextHit: {
    bodyIndex: number;
    charOffset: number;
    affinity?: "leading" | "trailing";
  } | null = null;
  private readonly inputElement: HTMLElement;
  private readonly scrollContainer: HTMLElement;

  constructor(scrollContainer: HTMLElement) {
    this.scrollContainer = scrollContainer;
    this.inputElement = {
      isContentEditable: false,
      focus: () => {
        this.focused = true;
      },
    } as unknown as HTMLElement;
  }

  renderModel(model: DocViewModel): void {
    this.model = model;
  }

  destroy(): void {}

  hitTest(): {
    bodyIndex: number;
    charOffset: number;
    affinity: "leading" | "trailing";
  } {
    if (this.nextHit) {
      return {
        bodyIndex: this.nextHit.bodyIndex,
        charOffset: this.nextHit.charOffset,
        affinity: this.nextHit.affinity ?? "trailing",
      };
    }
    const focus = this.selection?.focus ?? { bodyIndex: 0, charOffset: 0 };
    return { ...focus, affinity: "trailing" };
  }

  setCursor(pos: { bodyIndex: number; charOffset: number }): void {
    this.selection = {
      anchor: { ...pos },
      focus: { ...pos },
    };
  }

  setSelection(sel: DocSelection): void {
    this.selection = {
      anchor: { ...sel.anchor },
      focus: { ...sel.focus },
    };
  }

  getSelection(): DocSelection | null {
    if (!this.selection) return null;
    return {
      anchor: { ...this.selection.anchor },
      focus: { ...this.selection.focus },
    };
  }

  getCursorRect(): { x: number; y: number; w: number; h: number } {
    const focus = this.selection?.focus ?? { bodyIndex: 0, charOffset: 0 };
    return { x: 24, y: focus.bodyIndex * 24, w: 1, h: 18 };
  }

  getCursorRectForPosition(pos: { bodyIndex: number; charOffset: number }): {
    x: number;
    y: number;
    w: number;
    h: number;
  } {
    return { x: 24 + pos.charOffset * 8, y: pos.bodyIndex * 24, w: 1, h: 18 };
  }

  getSelectionRects(sel: DocSelection): Array<{
    x: number;
    y: number;
    w: number;
    h: number;
  }> {
    const start = Math.min(sel.anchor.charOffset, sel.focus.charOffset);
    const end = Math.max(sel.anchor.charOffset, sel.focus.charOffset);
    if (sel.anchor.bodyIndex !== sel.focus.bodyIndex || start === end)
      return [];
    return [
      {
        x: 24 + start * 8,
        y: sel.anchor.bodyIndex * 24,
        w: Math.max(1, (end - start) * 8),
        h: 18,
      },
    ];
  }

  getInlineImageAtPoint(): {
    bodyIndex: number;
    charOffset: number;
    imageIndex: number;
    rect: { x: number; y: number; w: number; h: number };
  } | null {
    if (!this.model) return null;
    for (
      let bodyIndex = 0;
      bodyIndex < this.model.body.length;
      bodyIndex += 1
    ) {
      const item = this.model.body[bodyIndex];
      if (!item || item.type !== "paragraph") continue;
      let charOffset = 0;
      for (const run of item.runs ?? []) {
        const typedRun = run as VmRun & {
          inlineImage?: { imageIndex?: number };
          footnoteRef?: number;
          endnoteRef?: number;
        };
        if (typedRun.inlineImage) {
          return {
            bodyIndex,
            charOffset,
            imageIndex: typedRun.inlineImage.imageIndex ?? 0,
            rect: { x: 64, y: bodyIndex * 24, w: 24, h: 18 },
          };
        }
        if (typedRun.hasBreak || typedRun.hasTab) {
          charOffset += (typedRun.text?.length ?? 0) + 1;
        } else if (
          typedRun.footnoteRef != null ||
          typedRun.endnoteRef != null
        ) {
          charOffset += String(
            typedRun.footnoteRef ?? typedRun.endnoteRef ?? "",
          ).length;
        } else {
          charOffset += typedRun.text?.length ?? 0;
        }
      }
    }
    return null;
  }

  getInlineImageRect(pos: {
    bodyIndex: number;
    charOffset: number;
  }): { x: number; y: number; w: number; h: number } | null {
    const hit = this.getInlineImageAtPoint();
    if (
      hit &&
      hit.bodyIndex === pos.bodyIndex &&
      hit.charOffset === pos.charOffset
    ) {
      return hit.rect;
    }
    return null;
  }

  getTableCellAtPoint(): {
    bodyIndex: number;
    row: number;
    col: number;
    rect: { x: number; y: number; w: number; h: number };
  } | null {
    const table =
      this.model?.body.findIndex((item) => item.type === "table") ?? -1;
    if (table < 0) return null;
    const rect = this.getTableCellRect({ bodyIndex: table, row: 0, col: 0 });
    if (!rect) return null;
    return { bodyIndex: table, row: 0, col: 0, rect };
  }

  getTableCellRect(cell: {
    bodyIndex: number;
    row: number;
    col: number;
  }): { x: number; y: number; w: number; h: number } | null {
    const bodyItem = this.model?.body[cell.bodyIndex];
    if (!bodyItem || bodyItem.type !== "table") return null;
    const table = bodyItem as {
      rows?: Array<{ cells?: unknown[] }>;
    };
    const row = table.rows?.[cell.row];
    if (!row || !row.cells?.[cell.col]) return null;
    return {
      x: 120 + cell.col * 96,
      y: 48 + cell.row * 28 + cell.bodyIndex * 32,
      w: 90,
      h: 24,
    };
  }

  getInputElement(): HTMLElement {
    return this.inputElement;
  }

  isFocused(): boolean {
    return this.focused;
  }

  focus(): void {
    this.focused = true;
  }

  getScrollContainer(): HTMLElement {
    return this.scrollContainer;
  }
}

class FakeInput implements InputAdapter {
  private inputHandler: ((input: NormalizedInput) => void) | null = null;
  private shortcutHandler: ((key: string, shift: boolean) => void) | null =
    null;
  private navigateHandler:
    | ((payload: {
        key:
          | "ArrowLeft"
          | "ArrowRight"
          | "ArrowUp"
          | "ArrowDown"
          | "Home"
          | "End"
          | "PageUp"
          | "PageDown";
        shift: boolean;
        meta: boolean;
        alt: boolean;
        ctrl: boolean;
      }) => void)
    | null = null;
  private selectionChangeHandler: (() => void) | null = null;
  private pointerDownHandler: ((payload: PointerDownPayload) => void) | null =
    null;
  private pointerMoveHandler: ((x: number, y: number) => void) | null = null;
  private pointerUpHandler: (() => void) | null = null;
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

  onInput(handler: (input: NormalizedInput) => void): void {
    this.inputHandler = handler;
  }

  onShortcut(handler: (key: string, shift: boolean) => void): void {
    this.shortcutHandler = handler;
  }

  onNavigate(
    handler: (payload: {
      key:
        | "ArrowLeft"
        | "ArrowRight"
        | "ArrowUp"
        | "ArrowDown"
        | "Home"
        | "End"
        | "PageUp"
        | "PageDown";
      shift: boolean;
      meta: boolean;
      alt: boolean;
      ctrl: boolean;
    }) => void,
  ): void {
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

  onPointerDown(handler: (payload: PointerDownPayload) => void): void {
    this.pointerDownHandler = handler;
  }

  onPasteImage(handler: (file: File) => void | Promise<void>): void {
    this.pasteImageHandler = handler;
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

  destroy(): void {}

  emitInput(input: NormalizedInput): void {
    this.inputHandler?.(input);
  }

  emitNavigate(payload: {
    key:
      | "ArrowLeft"
      | "ArrowRight"
      | "ArrowUp"
      | "ArrowDown"
      | "Home"
      | "End"
      | "PageUp"
      | "PageDown";
    shift: boolean;
    meta: boolean;
    alt: boolean;
    ctrl: boolean;
  }): void {
    this.navigateHandler?.(payload);
  }

  emitShortcut(key: string, shift = false): void {
    this.shortcutHandler?.(key, shift);
  }

  emitPointerDown(payload: PointerDownPayload): void {
    this.pointerDownHandler?.(payload);
  }

  emitPointerMove(x: number, y: number): void {
    this.pointerMoveHandler?.(x, y);
  }

  emitPointerUp(): void {
    this.pointerUpHandler?.();
  }

  emitSelectionChange(): void {
    this.selectionChangeHandler?.();
  }

  requestCopyText(): string | null | undefined {
    return this.requestCopyTextHandler?.();
  }

  requestCutText(): string | null | undefined {
    return this.requestCutTextHandler?.();
  }

  requestCopyHtml(): string | null | undefined {
    return this.requestCopyHtmlHandler?.();
  }

  requestCutHtml(): string | null | undefined {
    return this.requestCutHtmlHandler?.();
  }

  async emitPasteImage(file: File): Promise<void> {
    await this.pasteImageHandler?.(file);
  }
}

function makeControllerHarness(): {
  controller: DocEditorController;
  renderer: FakeRenderer;
  input: FakeInput;
  docEdit: () => DocEdit;
  destroy: () => void;
} {
  const root = { activeElement: null };
  const container = {
    scrollTop: 0,
    scrollLeft: 0,
    clientHeight: 600,
    style: { position: "" },
    querySelector: () => null,
    querySelectorAll: () => [],
    addEventListener: () => undefined,
    removeEventListener: () => undefined,
    appendChild: () => undefined,
    getRootNode: () => root,
    getBoundingClientRect: () => ({
      top: 0,
      bottom: 600,
      left: 0,
      right: 800,
      x: 0,
      y: 0,
      width: 800,
      height: 600,
    }),
    dispatchEvent: () => true,
  } as unknown as HTMLElement;
  const renderer = new FakeRenderer(container);
  const input = new FakeInput();
  const controller = createEditorController(renderer, input, container);
  controller.loadBlank();
  renderer.setCursor({ bodyIndex: 0, charOffset: 0 });
  input.emitSelectionChange();
  return {
    controller,
    renderer,
    input,
    docEdit: () => controller.getDocEdit() as DocEdit,
    destroy: () => controller.destroy(),
  };
}

// ---------------------------------------------------------------------------
// 1. Blank document
// ---------------------------------------------------------------------------

describe("blank document", () => {
  test("creates a document with 1 empty paragraph", () => {
    const editor = DocEdit.blank();
    const model = vm(editor);
    expect(model.body.length).toBe(1);
    expect(model.body[0]!.type).toBe("paragraph");
    const text = firstParaText(editor);
    expect(text).toBe("");
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 2–4. Insert text
// ---------------------------------------------------------------------------

describe("insert text", () => {
  test("inserts text at end", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    expect(firstParaText(editor)).toBe("Hello");
    editor.free();
  });

  test("sequential inserts accumulate", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "insertText",
      data: " world",
      anchor: pos(0, 5),
    });
    expect(firstParaText(editor)).toBe("Hello world");
    editor.free();
  });

  test("insert at middle of text", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Helloworld",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertText", data: " ", anchor: pos(0, 5) });
    expect(firstParaText(editor)).toBe("Hello world");
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 5–9. Delete operations
// ---------------------------------------------------------------------------

describe("delete operations", () => {
  test("delete backward removes last char", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(0, 5),
      focus: pos(0, 5),
    });
    expect(firstParaText(editor)).toBe("Hell");
    editor.free();
  });

  test("backspace at start is a no-op", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(0, 0),
      focus: pos(0, 0),
    });
    expect(firstParaText(editor)).toBe("Hello");
    editor.free();
  });

  test("delete forward removes char at cursor", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "deleteContentForward",
      anchor: pos(0, 0),
      focus: pos(0, 0),
    });
    expect(firstParaText(editor)).toBe("ello");
    editor.free();
  });

  test("range delete removes selected text", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(0, 0),
      focus: pos(0, 6),
    });
    expect(firstParaText(editor)).toBe("world");
    editor.free();
  });

  test("delete by cut removes selection", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "deleteByCut",
      anchor: pos(0, 5),
      focus: pos(0, 11),
    });
    expect(firstParaText(editor)).toBe("Hello");
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 10. Line break
// ---------------------------------------------------------------------------

describe("line break", () => {
  test("inserts a line break sentinel", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "insertLineBreak",
      anchor: pos(0, 5),
    });
    const runs = firstRuns(model);
    const hasBreakRun = runs.some((r) => r.hasBreak === true);
    expect(hasBreakRun).toBe(true);
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 11. Sentinel protection
// ---------------------------------------------------------------------------

describe("sentinel protection", () => {
  test("cannot backspace a sentinel (line break)", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "AB", anchor: pos(0, 0) });
    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 2) });
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(0, 3),
      focus: pos(0, 3),
    });
    const model = vm(editor);
    const runs = firstRuns(model);
    const hasBreak = runs.some((r) => r.hasBreak === true);
    expect(hasBreak).toBe(true);
    editor.free();
  });

  test("cannot forward-delete a sentinel", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "AB", anchor: pos(0, 0) });
    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 2) });
    applyAndVm(editor, {
      type: "deleteContentForward",
      anchor: pos(0, 2),
      focus: pos(0, 2),
    });
    const model = vm(editor);
    const runs = firstRuns(model);
    const hasBreak = runs.some((r) => r.hasBreak === true);
    expect(hasBreak).toBe(true);
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 12–15. Formatting
// ---------------------------------------------------------------------------

describe("formatting", () => {
  test("format bold marks range as bold", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });
    const runs = firstRuns(model);
    const boldRun = runs.find((r) => r.bold === true);
    expect(boldRun).toBeDefined();
    expect(boldRun!.text).toBe("Hello");
    editor.free();
  });

  test("format italic marks range as italic", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "formatItalic",
      anchor: pos(0, 6),
      focus: pos(0, 11),
    });
    const runs = firstRuns(model);
    const italicRun = runs.find((r) => r.italic === true);
    expect(italicRun).toBeDefined();
    expect(italicRun!.text).toBe("world");
    editor.free();
  });

  test("format underline marks range as underline", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "formatUnderline",
      anchor: pos(0, 0),
      focus: pos(0, 11),
    });
    const runs = firstRuns(model);
    const ulRun = runs.find((r) => r.underline === true);
    expect(ulRun).toBeDefined();
    editor.free();
  });

  test("bold + italic stacking on same range", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });
    const model = applyAndVm(editor, {
      type: "formatItalic",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });
    const runs = firstRuns(model);
    const run = runs.find((r) => r.text === "Hello");
    expect(run).toBeDefined();
    expect(run!.bold).toBe(true);
    expect(run!.italic).toBe(true);
    editor.free();
  });

  test("format strikethrough marks range as strikethrough", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "formatStrikethrough",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });
    const runs = firstRuns(model);
    const strikeRun = runs.find((r) => r.strikethrough === true);
    expect(strikeRun).toBeDefined();
    expect(strikeRun!.text).toBe("Hello");
    editor.free();
  });

  test("format collapsed range (cursor) is a no-op on CRDT", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    editor.save();
    expect(editor.isDirty()).toBe(false);

    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 3),
      focus: pos(0, 3),
    });
    expect(editor.isDirty()).toBe(false);
    editor.free();
  });

  test("insertText with explicit attrs applies formatting", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Bold",
      anchor: pos(0, 0),
      attrs: { bold: true },
    } as any);
    const model = vm(editor);
    const runs = firstRuns(model);
    expect(runs.length).toBeGreaterThan(0);
    const run0 = runs[0]!;
    expect(run0.text).toBe("Bold");
    expect(run0.bold).toBe(true);
    editor.free();
  });

  test("setTextAttrs applies font family, size, and color", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Styled",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "setTextAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 6),
      attrs: {
        fontFamily: "Noto Serif",
        fontSizePt: 18,
        color: "3366FF",
      },
    } as const);
    const run = firstRuns(model)[0]!;
    expect(run.text).toBe("Styled");
    expect(run.fontFamily).toBe("Noto Serif");
    expect(run.fontSizePt).toBe(18);
    expect(run.color).toBe("3366FF");
    editor.free();
  });

  test("setTextAttrs applies hyperlink", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Link",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "setTextAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 4),
      attrs: {
        hyperlink: "https://example.com",
      },
    } as const);
    const run = firstRuns(model)[0]!;
    expect(run.text).toBe("Link");
    expect(run.hyperlink).toBe("https://example.com");
    editor.free();
  });

  test("setTextAttrs applies highlight", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Marked",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "setTextAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 6),
      attrs: { highlight: "yellow" },
    });
    const run = firstRuns(model)[0]!;
    expect(run.text).toBe("Marked");
    expect(run.highlight).toBe("yellow");
    editor.free();
  });

  test("setParagraphAttrs applies heading level", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Title",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 5),
      attrs: { headingLevel: 2 },
    } as const);
    expect(model.body[0]?.type).toBe("paragraph");
    expect((model.body[0] as { headingLevel?: number }).headingLevel).toBe(2);
    editor.free();
  });

  test("setParagraphAttrs applies paragraph alignment", () => {
    const alignments = ["left", "center", "right", "justify"] as const;

    for (const alignment of alignments) {
      const editor = DocEdit.blank();
      applyAndVm(editor, {
        type: "insertText",
        data: alignment,
        anchor: pos(0, 0),
      });
      const model = applyAndVm(editor, {
        type: "setParagraphAttrs",
        anchor: pos(0, 0),
        focus: pos(0, alignment.length),
        attrs: { alignment },
      } as const);
      expect(model.body[0]?.type).toBe("paragraph");
      expect((model.body[0] as { alignment?: string }).alignment).toBe(
        alignment,
      );
      editor.free();
    }
  });

  test("setParagraphAttrs applies paragraph spacing and indent", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Block",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 5),
      attrs: {
        spacingBeforePt: 12,
        spacingAfterPt: 18,
        indentLeftPt: 36,
        indentFirstLinePt: 18,
      },
    } as const);
    const paragraph = model.body[0] as {
      spacingBeforePt?: number;
      spacingAfterPt?: number;
      indents?: { leftPt?: number; firstLinePt?: number };
    };
    expect(paragraph.spacingBeforePt).toBe(12);
    expect(paragraph.spacingAfterPt).toBe(18);
    expect(paragraph.indents?.leftPt).toBe(36);
    expect(paragraph.indents?.firstLinePt).toBe(18);
    editor.free();
  });

  test("setParagraphAttrs applies line spacing multiple", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Line space",
      anchor: pos(0, 0),
    });
    const model = applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 10),
      attrs: {
        lineSpacingMultiple: 1.5,
      },
    } as const);
    const paragraph = model.body[0] as {
      lineSpacing?: { value?: number; rule?: string };
    };
    expect(paragraph.lineSpacing?.value).toBe(1.5);
    expect(paragraph.lineSpacing?.rule).toBe("auto");
    editor.free();
  });

  test("insertText with bold+italic attrs", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Both",
      anchor: pos(0, 0),
      attrs: { bold: true, italic: true },
    } as any);
    const model = vm(editor);
    const runs = firstRuns(model);
    expect(runs.length).toBeGreaterThan(0);
    const run0 = runs[0]!;
    expect(run0.text).toBe("Both");
    expect(run0.bold).toBe(true);
    expect(run0.italic).toBe(true);
    editor.free();
  });

  test("insertText inherits formatting from character to the left", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });
    applyAndVm(editor, {
      type: "insertText",
      data: " world",
      anchor: pos(0, 5),
    });
    const model = vm(editor);
    const runs = firstRuns(model);
    const allBold = runs.every((r) => r.bold === true);
    expect(allBold).toBe(true);
    const allText = runs.map((r) => r.text).join("");
    expect(allText).toBe("Hello world");
    editor.free();
  });

  test("all four format types on same text", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "Test", anchor: pos(0, 0) });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 4),
    });
    applyAndVm(editor, {
      type: "formatItalic",
      anchor: pos(0, 0),
      focus: pos(0, 4),
    });
    applyAndVm(editor, {
      type: "formatUnderline",
      anchor: pos(0, 0),
      focus: pos(0, 4),
    });
    applyAndVm(editor, {
      type: "formatStrikethrough",
      anchor: pos(0, 0),
      focus: pos(0, 4),
    });
    const runs = firstRuns(vm(editor));
    const run = runs.find((r) => r.text === "Test");
    expect(run).toBeDefined();
    expect(run!.bold).toBe(true);
    expect(run!.italic).toBe(true);
    expect(run!.underline).toBe(true);
    expect(run!.strikethrough).toBe(true);
    editor.free();
  });

  test("format partial overlap creates split runs", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "ABCDEF",
      anchor: pos(0, 0),
    });
    // Bold AB, Italic CD — overlapping at nothing, but creates distinct runs
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 2),
    });
    applyAndVm(editor, {
      type: "formatItalic",
      anchor: pos(0, 2),
      focus: pos(0, 4),
    });
    const runs = firstRuns(vm(editor));
    // Should have at least 3 runs: bold "AB", italic "CD", plain "EF"
    expect(runs.length).toBeGreaterThanOrEqual(3);
    const run0 = runs[0]!;
    const run1 = runs[1]!;
    const run2 = runs[2]!;
    expect(run0.text).toBe("AB");
    expect(run0.bold).toBe(true);
    expect(run1.text).toBe("CD");
    expect(run1.italic).toBe(true);
    expect(run2.text).toBe("EF");
    expect(run2.bold).toBeUndefined();
    expect(run2.italic).toBeUndefined();
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 16. Paste with selection
// ---------------------------------------------------------------------------

describe("paste", () => {
  test("paste replaces selected text", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "insertFromPaste",
      data: "there",
      anchor: pos(0, 6),
      focus: pos(0, 11),
    });
    expect(firstParaText(editor)).toBe("Hello there");
    editor.free();
  });

  test("multi-paragraph paste splits paragraphs on newlines", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "ab",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "insertFromPaste",
      data: "X\nY",
      anchor: pos(0, 1),
      focus: pos(0, 1),
    });
    expect(vm(editor).body.length).toBe(2);
    expect(paraText(editor, 0)).toBe("aX");
    expect(paraText(editor, 1)).toBe("Yb");
    editor.free();
  });

  test("multi-paragraph paste replaces multi-paragraph selection", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "abcd",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 4) });
    applyAndVm(editor, {
      type: "insertText",
      data: "efgh",
      anchor: pos(1, 0),
    });
    applyAndVm(editor, {
      type: "insertFromPaste",
      data: "X\nY",
      anchor: pos(0, 1),
      focus: pos(1, 2),
    });
    expect(vm(editor).body.length).toBe(2);
    expect(paraText(editor, 0)).toBe("aX");
    expect(paraText(editor, 1)).toBe("Ygh");
    editor.free();
  });

  test("paste at collapsed cursor inserts without deleting", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "insertFromPaste",
      data: " world",
      anchor: pos(0, 5),
      focus: pos(0, 5),
    });
    expect(firstParaText(editor)).toBe("Hello world");
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 17. Empty insert is a no-op
// ---------------------------------------------------------------------------

describe("empty insert", () => {
  test("inserting empty string does nothing", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "", anchor: pos(0, 0) });
    expect(editor.isDirty()).toBe(false);
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 18. isDirty tracking
// ---------------------------------------------------------------------------

describe("isDirty tracking", () => {
  test("clean → edit → dirty → save → clean", () => {
    const editor = DocEdit.blank();
    expect(editor.isDirty()).toBe(false);

    applyAndVm(editor, { type: "insertText", data: "Hi", anchor: pos(0, 0) });
    expect(editor.isDirty()).toBe(true);

    editor.save();
    expect(editor.isDirty()).toBe(false);
    editor.free();
  });

  test("multiple edits still dirty until save", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "A", anchor: pos(0, 0) });
    applyAndVm(editor, { type: "insertText", data: "B", anchor: pos(0, 1) });
    applyAndVm(editor, { type: "insertText", data: "C", anchor: pos(0, 2) });
    expect(editor.isDirty()).toBe(true);
    editor.save();
    expect(editor.isDirty()).toBe(false);
    editor.free();
  });

  test("formatting change marks dirty", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    editor.save();
    expect(editor.isDirty()).toBe(false);
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });
    expect(editor.isDirty()).toBe(true);
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 19. Save roundtrip
// ---------------------------------------------------------------------------

describe("save roundtrip", () => {
  test("edit → save → reload → verify content", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Roundtrip",
      anchor: pos(0, 0),
    });

    const bytes = editor.save();
    expect(bytes.length).toBeGreaterThan(0);

    const editor2 = new DocEdit(bytes);
    expect(firstParaText(editor2)).toBe("Roundtrip");

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves bold formatting", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);

    const runs = firstRuns(vm(editor2));
    const boldRun = runs.find((r) => r.bold === true);
    expect(boldRun).toBeDefined();
    expect(boldRun!.text).toBe("Hello");

    // Non-bold part
    const plainRun = runs.find((r) => r.text?.includes("world"));
    expect(plainRun).toBeDefined();
    expect(plainRun!.bold).toBeUndefined();

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves italic formatting", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "formatItalic",
      anchor: pos(0, 6),
      focus: pos(0, 11),
    });

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);

    const runs = firstRuns(vm(editor2));
    const italicRun = runs.find((r) => r.italic === true);
    expect(italicRun).toBeDefined();
    expect(italicRun!.text).toBe("world");

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves stacked formatting", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "Test", anchor: pos(0, 0) });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 4),
    });
    applyAndVm(editor, {
      type: "formatItalic",
      anchor: pos(0, 0),
      focus: pos(0, 4),
    });

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);

    const runs = firstRuns(vm(editor2));
    const run = runs.find((r) => r.text === "Test");
    expect(run).toBeDefined();
    expect(run!.bold).toBe(true);
    expect(run!.italic).toBe(true);

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves font family, size, and color", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Styled",
      anchor: pos(0, 0),
      attrs: {
        fontFamily: "Noto Serif",
        fontSizePt: 18,
        color: "3366FF",
      },
    } as any);

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);

    const run = firstRuns(vm(editor2))[0]!;
    expect(run.text).toBe("Styled");
    expect(run.fontFamily).toBe("Noto Serif");
    expect(run.fontSizePt).toBe(18);
    expect(run.color).toBe("3366FF");

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves hyperlink", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Link",
      anchor: pos(0, 0),
      attrs: {
        hyperlink: "https://example.com",
      },
    } as any);

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);

    const run = firstRuns(vm(editor2))[0]!;
    expect(run.text).toBe("Link");
    expect(run.hyperlink).toBe("https://example.com");

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves highlight", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Marked",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "setTextAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 6),
      attrs: { highlight: "yellow" },
    });

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);
    const run = paraRuns(editor2, 0)[0]!;
    expect(run.text).toBe("Marked");
    expect(run.highlight).toBe("yellow");

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves heading level", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Heading",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 7),
      attrs: { headingLevel: 1 },
    } as const);

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);
    const model = vm(editor2);
    expect((model.body[0] as { headingLevel?: number }).headingLevel).toBe(1);

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves paragraph alignment", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Centered",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 8),
      attrs: { alignment: "center" },
    } as const);

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);
    const paragraph = paraModel(editor2, 0);
    expect(paragraph?.alignment).toBe("center");
    expect(paraText(editor2, 0)).toBe("Centered");

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves paragraph spacing and indent", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Block",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 5),
      attrs: {
        spacingBeforePt: 12,
        spacingAfterPt: 18,
        indentLeftPt: 36,
        indentFirstLinePt: 18,
      },
    } as const);

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);
    const paragraph = paraModel(editor2, 0);
    expect(paragraph?.spacingBeforePt).toBe(12);
    expect(paragraph?.spacingAfterPt).toBe(18);
    expect(paragraph?.indents?.leftPt).toBe(36);
    expect(paragraph?.indents?.firstLinePt).toBe(18);

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves line spacing multiple", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Line space",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 10),
      attrs: {
        lineSpacingMultiple: 1.5,
      },
    } as const);

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);
    const paragraph = paraModel(editor2, 0);
    expect(paragraph?.lineSpacing?.value).toBe(1.5);
    expect(paragraph?.lineSpacing?.rule).toBe("auto");

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves a bulleted list paragraph", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Bullet item",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 11),
      attrs: {
        numberingKind: "bullet",
        numberingNumId: 1,
        numberingIlvl: 0,
      },
    });

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);

    const para = paraModel(editor2, 0);
    expect(firstParaText(editor2)).toBe("Bullet item");
    expect(para?.numbering?.format).toBe("bullet");
    expect(para?.numbering?.level).toBe(0);
    expect(para?.numbering?.text.length ?? 0).toBeGreaterThan(0);
    expect(formattingAt(editor2, 0, 0).numberingKind).toBe("bullet");

    editor.free();
    editor2.free();
  });

  test("save roundtrip preserves numbered list sequencing across paragraphs", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "One",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 3) });
    applyAndVm(editor, {
      type: "insertText",
      data: "Two",
      anchor: pos(1, 0),
    });
    applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(1, 3),
      attrs: {
        numberingKind: "decimal",
        numberingNumId: 7,
        numberingIlvl: 0,
      },
    });

    const bytes = editor.save();
    const editor2 = new DocEdit(bytes);

    const para0 = paraModel(editor2, 0);
    const para1 = paraModel(editor2, 1);
    expect(paraText(editor2, 0)).toBe("One");
    expect(paraText(editor2, 1)).toBe("Two");
    expect(para0?.numbering?.format).toBe("decimal");
    expect(para0?.numbering?.text).toBe("1.");
    expect(para1?.numbering?.format).toBe("decimal");
    expect(para1?.numbering?.text).toBe("2.");
    expect(para1?.numbering?.numId).toBe(para0?.numbering?.numId);

    editor.free();
    editor2.free();
  });

  test("controller-authored list survives save and reload", () => {
    const harness = makeControllerHarness();

    harness.controller.toggleList("bullet");
    harness.input.emitInput({ type: "insertText", data: "From controller" });

    const bytes = harness.controller.save();
    const editor2 = new DocEdit(bytes);

    const para = paraModel(editor2, 0);
    expect(firstParaText(editor2)).toBe("From controller");
    expect(para?.numbering?.format).toBe("bullet");
    expect(para?.numbering?.level).toBe(0);
    expect(para?.numbering?.text.length ?? 0).toBeGreaterThan(0);

    harness.destroy();
    editor2.free();
  });
});

// ---------------------------------------------------------------------------
// 20. Position encoding
// ---------------------------------------------------------------------------

describe("position encoding", () => {
  test("encodePosition works end-to-end in intents", () => {
    const editor = DocEdit.blank();
    const anchor = DocEdit.encodePosition(0, 0);
    expect(typeof anchor).toBe("string");
    expect(anchor.length).toBeGreaterThan(0);

    editor.applyIntent(
      JSON.stringify({ type: "insertText", data: "Test", anchor }),
    );
    expect(firstParaText(editor)).toBe("Test");
    editor.free();
  });

  test("different positions produce different encodings", () => {
    const p1 = DocEdit.encodePosition(0, 0);
    const p2 = DocEdit.encodePosition(0, 5);
    const p3 = DocEdit.encodePosition(1, 0);
    expect(p1).not.toBe(p2);
    expect(p1).not.toBe(p3);
    expect(p2).not.toBe(p3);
  });
});

// ---------------------------------------------------------------------------
// 21. Borrow regression test
// ---------------------------------------------------------------------------

describe("borrow regression", () => {
  test("applyIntent → isDirty → viewModel sequentially without panic", () => {
    const editor = DocEdit.blank();
    editor.applyIntent(
      JSON.stringify({ type: "insertText", data: "Test", anchor: pos(0, 0) }),
    );
    const dirty = editor.isDirty();
    const model = vm(editor);

    expect(dirty).toBe(true);
    expect(model.body.length).toBe(1);
    expect(firstParaText(editor)).toBe("Test");
    editor.free();
  });

  test("rapid successive calls do not panic", () => {
    const editor = DocEdit.blank();
    for (let i = 0; i < 10; i++) {
      editor.applyIntent(
        JSON.stringify({
          type: "insertText",
          data: String(i),
          anchor: pos(0, i),
        }),
      );
      editor.isDirty();
      editor.viewModel();
    }
    expect(firstParaText(editor)).toBe("0123456789");
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 22. Offset overflow regression
// ---------------------------------------------------------------------------

describe("offset overflow regression", () => {
  test("insert text char by char then line break", () => {
    const editor = DocEdit.blank();
    for (let i = 0; i < 5; i++) {
      applyAndVm(editor, {
        type: "insertText",
        data: "Hello"[i],
        anchor: pos(0, i),
      });
    }
    expect(firstParaText(editor)).toBe("Hello");
    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 5) });
    applyAndVm(editor, { type: "insertText", data: "W", anchor: pos(0, 6) });
    editor.free();
  });

  test("insert at offset beyond text length (browser stale cursor)", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "Hi", anchor: pos(0, 0) });
    applyAndVm(editor, { type: "insertText", data: "!", anchor: pos(0, 100) });
    expect(firstParaText(editor)).toBe("Hi!");
    editor.free();
  });

  test("line break at offset beyond text length", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "Hi", anchor: pos(0, 0) });
    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 100) });
    const model = vm(editor);
    const runs = model.body[0]?.runs ?? [];
    expect(runs.some((r: any) => r.hasBreak)).toBe(true);
    editor.free();
  });

  test("rapid edit-viewModel cycles don't corrupt state", () => {
    const editor = DocEdit.blank();
    for (let i = 0; i < 20; i++) {
      editor.applyIntent(
        JSON.stringify({
          type: "insertText",
          data: String.fromCharCode(65 + (i % 26)),
          anchor: DocEdit.encodePosition(0, i),
        }),
      );
      editor.viewModel();
    }
    editor.applyIntent(
      JSON.stringify({
        type: "insertLineBreak",
        anchor: DocEdit.encodePosition(0, 20),
      }),
    );
    editor.viewModel();
    editor.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "X",
        anchor: DocEdit.encodePosition(0, 21),
      }),
    );
    editor.viewModel();
    editor.free();
  });

  test("format then insert with attributes", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello World",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });
    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 5) });
    applyAndVm(editor, {
      type: "insertText",
      data: "After",
      anchor: pos(0, 6),
    });
    editor.free();
  });

  test("delete then insert line break", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello World",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(0, 5),
      focus: pos(0, 11),
    });
    expect(firstParaText(editor)).toBe("Hello");
    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 5) });
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 23. Browser flow simulation: type → Enter → type
// ---------------------------------------------------------------------------

describe("type-enter-type browser flow", () => {
  test("type Hi, press Enter (line break), type after", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "H", anchor: pos(0, 0) });
    applyAndVm(editor, { type: "insertText", data: "i", anchor: pos(0, 1) });
    expect(firstParaText(editor)).toBe("Hi");

    const modelAfterEnter = applyAndVm(editor, {
      type: "insertLineBreak",
      anchor: pos(0, 2),
    });

    const runs = firstRuns(modelAfterEnter);
    expect(runs.length).toBeGreaterThanOrEqual(1);
    const breakRun = runs.find((r) => r.hasBreak);
    expect(breakRun).toBeDefined();

    // Type after line break
    applyAndVm(editor, { type: "insertText", data: "X", anchor: pos(0, 3) });
    applyAndVm(editor, { type: "insertText", data: "Y", anchor: pos(0, 4) });

    const runsAfterXY = firstRuns(vm(editor));
    const allText = runsAfterXY.map((r) => r.text).join("");
    expect(allText).toContain("Hi");
    expect(allText).toContain("XY");

    editor.free();
  });

  test("sentinel run text is empty (not newline)", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "Hi", anchor: pos(0, 0) });
    const model = applyAndVm(editor, {
      type: "insertLineBreak",
      anchor: pos(0, 2),
    });

    const runs = firstRuns(model);
    const breakRun = runs.find((r) => r.hasBreak);
    expect(breakRun).toBeDefined();
    expect(breakRun!.text).toBe("");
    editor.free();
  });

  test("typing after line break produces correct runs", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "AB", anchor: pos(0, 0) });
    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 2) });
    applyAndVm(editor, { type: "insertText", data: "C", anchor: pos(0, 3) });
    const model = applyAndVm(editor, {
      type: "insertText",
      data: "D",
      anchor: pos(0, 4),
    });

    const runs = firstRuns(model);
    const textParts: string[] = [];
    let hasBreak = false;
    for (const r of runs) {
      if (r.text) textParts.push(r.text);
      if (r.hasBreak) hasBreak = true;
    }
    expect(textParts.join("")).toBe("ABCD");
    expect(hasBreak).toBe(true);
    editor.free();
  });

  test("multiple consecutive Enter presses", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "A", anchor: pos(0, 0) });

    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 1) });
    let runs = firstRuns(vm(editor));
    expect(runs.filter((r) => r.hasBreak).length).toBe(1);

    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 2) });
    runs = firstRuns(vm(editor));
    expect(runs.filter((r) => r.hasBreak).length).toBe(2);

    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 3) });
    runs = firstRuns(vm(editor));
    expect(runs.filter((r) => r.hasBreak).length).toBe(3);

    applyAndVm(editor, { type: "insertText", data: "B", anchor: pos(0, 4) });
    runs = firstRuns(vm(editor));
    const allText = runs.map((r) => r.text).join("");
    expect(allText).toContain("A");
    expect(allText).toContain("B");
    expect(runs.filter((r) => r.hasBreak).length).toBe(3);

    editor.free();
  });

  test("Enter on empty document (no text typed yet)", () => {
    const editor = DocEdit.blank();

    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 0) });
    let runs = firstRuns(vm(editor));
    expect(runs.some((r) => r.hasBreak)).toBe(true);

    applyAndVm(editor, { type: "insertLineBreak", anchor: pos(0, 1) });
    runs = firstRuns(vm(editor));
    expect(runs.filter((r) => r.hasBreak).length).toBe(2);

    applyAndVm(editor, { type: "insertText", data: "X", anchor: pos(0, 2) });
    runs = firstRuns(vm(editor));
    expect(runs.map((r) => r.text).join("")).toContain("X");

    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 24. InsertParagraph (Enter splits paragraph)
// ---------------------------------------------------------------------------

describe("insertParagraph", () => {
  test("split at end creates new empty paragraph", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });

    const model = vm(editor);
    expect(model.body.length).toBe(2);
    expect(paraText(editor, 0)).toBe("Hello");
    expect(paraText(editor, 1)).toBe("");
    editor.free();
  });

  test("split at start moves all content to new paragraph", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 0) });

    const model = vm(editor);
    expect(model.body.length).toBe(2);
    expect(paraText(editor, 0)).toBe("");
    expect(paraText(editor, 1)).toBe("Hello");
    editor.free();
  });

  test("split in the middle distributes content", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });

    const model = vm(editor);
    expect(model.body.length).toBe(2);
    expect(paraText(editor, 0)).toBe("Hello");
    expect(paraText(editor, 1)).toBe(" world");
    editor.free();
  });

  test("split empty paragraph creates two empty paragraphs", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 0) });

    const model = vm(editor);
    expect(model.body.length).toBe(2);
    expect(paraText(editor, 0)).toBe("");
    expect(paraText(editor, 1)).toBe("");
    editor.free();
  });

  test("split preserves bold formatting on moved content", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    // Bold the entire text
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 11),
    });
    // Split at offset 5
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });

    const model = vm(editor);
    expect(model.body.length).toBe(2);

    // First paragraph: "Hello" should still be bold
    const runs0 = paraRuns(editor, 0);
    expect(runs0.every((r) => r.bold === true)).toBe(true);
    expect(runs0.map((r) => r.text).join("")).toBe("Hello");

    // Second paragraph: " world" should still be bold
    const runs1 = paraRuns(editor, 1);
    expect(runs1.every((r) => r.bold === true)).toBe(true);
    expect(runs1.map((r) => r.text).join("")).toBe(" world");
    editor.free();
  });

  test("split preserves mixed formatting", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "ABCDEF",
      anchor: pos(0, 0),
    });
    // Bold "ABC", italic "DEF"
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 3),
    });
    applyAndVm(editor, {
      type: "formatItalic",
      anchor: pos(0, 3),
      focus: pos(0, 6),
    });
    // Split between "C" and "D" (offset 3)
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 3) });

    const model = vm(editor);
    expect(model.body.length).toBe(2);

    // First paragraph: "ABC" bold
    const runs0 = paraRuns(editor, 0);
    expect(runs0.map((r) => r.text).join("")).toBe("ABC");
    expect(runs0.length).toBeGreaterThan(0);
    expect(runs0[0]!.bold).toBe(true);

    // Second paragraph: "DEF" italic
    const runs1 = paraRuns(editor, 1);
    expect(runs1.map((r) => r.text).join("")).toBe("DEF");
    expect(runs1.length).toBeGreaterThan(0);
    expect(runs1[0]!.italic).toBe(true);
    editor.free();
  });

  test("multiple paragraph splits", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Line1Line2Line3",
      anchor: pos(0, 0),
    });

    // Split after "Line1" (offset 5)
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });
    expect(vm(editor).body.length).toBe(2);
    expect(paraText(editor, 0)).toBe("Line1");
    expect(paraText(editor, 1)).toBe("Line2Line3");

    // Split the second paragraph after "Line2" (offset 5 in para 1)
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(1, 5) });
    expect(vm(editor).body.length).toBe(3);
    expect(paraText(editor, 0)).toBe("Line1");
    expect(paraText(editor, 1)).toBe("Line2");
    expect(paraText(editor, 2)).toBe("Line3");

    editor.free();
  });

  test("type in new paragraph after split", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });

    // Type in the new (second) paragraph
    applyAndVm(editor, {
      type: "insertText",
      data: "World",
      anchor: pos(1, 0),
    });
    expect(paraText(editor, 0)).toBe("Hello");
    expect(paraText(editor, 1)).toBe("World");
    editor.free();
  });

  test("bodyLength reflects paragraph count", () => {
    const editor = DocEdit.blank();
    expect(editor.bodyLength()).toBe(1);

    applyAndVm(editor, { type: "insertText", data: "A", anchor: pos(0, 0) });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 1) });
    expect(editor.bodyLength()).toBe(2);

    applyAndVm(editor, { type: "insertParagraph", anchor: pos(1, 0) });
    expect(editor.bodyLength()).toBe(3);

    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 25. Undo / Redo
// ---------------------------------------------------------------------------

describe("undo / redo", () => {
  test("undo keeps prior typing separate from paste", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "insertFromPaste",
      data: " world",
      anchor: pos(0, 5),
      focus: pos(0, 5),
    });
    expect(firstParaText(editor)).toBe("Hello world");

    applyAndVm(editor, { type: "historyUndo" });
    expect(firstParaText(editor)).toBe("Hello");

    applyAndVm(editor, { type: "historyRedo" });
    expect(firstParaText(editor)).toBe("Hello world");
    editor.free();
  });

  test("undo restores replaced selection in one step", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "insertFromPaste",
      data: "there",
      anchor: pos(0, 6),
      focus: pos(0, 11),
    });
    expect(firstParaText(editor)).toBe("Hello there");

    applyAndVm(editor, { type: "historyUndo" });
    expect(firstParaText(editor)).toBe("Hello world");
    editor.free();
  });

  test("undo keeps prior typing separate from paragraph split", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });
    expect(paraText(editor, 0)).toBe("Hello");
    expect(paraText(editor, 1)).toBe(" world");

    applyAndVm(editor, { type: "historyUndo" });
    expect(firstParaText(editor)).toBe("Hello world");
    expect(vm(editor).body.length).toBe(1);

    applyAndVm(editor, { type: "historyRedo" });
    expect(paraText(editor, 0)).toBe("Hello");
    expect(paraText(editor, 1)).toBe(" world");
    editor.free();
  });

  test("undo keeps prior typing separate from formatting", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });
    expect(firstParaText(editor)).toBe("Hello");
    expect(firstRuns(vm(editor))[0]?.bold).toBe(true);

    applyAndVm(editor, { type: "historyUndo" });
    expect(firstParaText(editor)).toBe("Hello");
    expect(firstRuns(vm(editor))[0]?.bold).toBeUndefined();

    applyAndVm(editor, { type: "historyRedo" });
    expect(firstRuns(vm(editor))[0]?.bold).toBe(true);
    editor.free();
  });

  test("undo keeps prior typing separate from committed composition text", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "A",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "insertFromComposition",
      data: "漢",
      anchor: pos(0, 1),
    });
    expect(firstParaText(editor)).toBe("A漢");

    applyAndVm(editor, { type: "historyUndo" });
    expect(firstParaText(editor)).toBe("A");

    applyAndVm(editor, { type: "historyRedo" });
    expect(firstParaText(editor)).toBe("A漢");
    editor.free();
  });

  test("undo reverses text insert", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    expect(firstParaText(editor)).toBe("Hello");

    // Wait a moment to ensure distinct undo capture
    applyAndVm(editor, { type: "historyUndo" });
    expect(firstParaText(editor)).toBe("");
    editor.free();
  });

  test("redo restores undone change", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "historyUndo" });
    expect(firstParaText(editor)).toBe("");

    applyAndVm(editor, { type: "historyRedo" });
    expect(firstParaText(editor)).toBe("Hello");
    editor.free();
  });

  test("undo on blank document is a no-op", () => {
    const editor = DocEdit.blank();
    // Should not throw
    applyAndVm(editor, { type: "historyUndo" });
    expect(firstParaText(editor)).toBe("");
    editor.free();
  });

  test("redo with nothing to redo is a no-op", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "historyRedo" });
    expect(firstParaText(editor)).toBe("");
    editor.free();
  });

  test("undo → viewModel → isDirty sequence doesn't panic", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "Test", anchor: pos(0, 0) });
    applyAndVm(editor, { type: "historyUndo" });
    // Rapid interleaving of different method calls
    const model = vm(editor);
    const dirty = editor.isDirty();
    expect(model.body.length).toBe(1);
    // After undo, the doc may or may not be dirty depending on undo tracking.
    // The important thing is no panic.
    expect(typeof dirty).toBe("boolean");
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 26. formattingAt (toolbar highlight query)
// ---------------------------------------------------------------------------

describe("formattingAt", () => {
  test("returns empty object at position 0", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    const fmt = editor.formattingAt(pos(0, 0));
    // At position 0, there's nothing to the left → no formatting
    expect(fmt.bold).toBeUndefined();
    expect(fmt.italic).toBeUndefined();
    editor.free();
  });

  test("detects bold at cursor inside bold text", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });

    // Cursor at offset 3 (inside "Hello") — char to the left is bold
    const fmt = editor.formattingAt(pos(0, 3));
    expect(fmt.bold).toBe(true);
    editor.free();
  });

  test("detects highlight at cursor inside highlighted text", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "setTextAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 11),
      attrs: { highlight: "yellow" },
    });
    const fmt = editor.formattingAt(pos(0, 5));
    expect(fmt.highlight).toBe("yellow");
    editor.free();
  });

  test("detects no bold at cursor outside bold text", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });

    // Cursor at offset 8 (inside "world") — not bold
    const fmt = editor.formattingAt(pos(0, 8));
    expect(fmt.bold).toBeUndefined();
    editor.free();
  });

  test("detects multiple formatting attrs", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "Test", anchor: pos(0, 0) });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 4),
    });
    applyAndVm(editor, {
      type: "formatItalic",
      anchor: pos(0, 0),
      focus: pos(0, 4),
    });

    const fmt = editor.formattingAt(pos(0, 2));
    expect(fmt.bold).toBe(true);
    expect(fmt.italic).toBe(true);
    editor.free();
  });

  test("detects formatting at boundary (right at end of bold)", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello world",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 5),
    });

    // Offset 5 → char to the left is 'o' in "Hello" (bold)
    const fmt = editor.formattingAt(pos(0, 5));
    expect(fmt.bold).toBe(true);

    // Offset 6 → char to the left is ' ' (space, not bold)
    const fmt2 = editor.formattingAt(pos(0, 6));
    expect(fmt2.bold).toBeUndefined();
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 27. Lists
// ---------------------------------------------------------------------------

describe("lists", () => {
  test("setParagraphAttrs applies bullet numbering and reports list formatting state", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Item 1",
      anchor: pos(0, 0),
    });

    applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 6),
      attrs: {
        numberingKind: "bullet",
        numberingNumId: 1,
        numberingIlvl: 0,
      },
    });

    const para = paraModel(editor, 0);
    expect(para?.numbering?.format).toBe("bullet");
    expect(para?.numbering?.numId).toBe(1);
    expect(para?.numbering?.level).toBe(0);
    expect(para?.numbering?.text.length ?? 0).toBeGreaterThan(0);

    const fmt = formattingAt(editor, 0, 0);
    expect(fmt.numberingKind).toBe("bullet");
    expect(fmt.numberingIlvl).toBe(0);
    editor.free();
  });

  test("setParagraphAttrs applies decimal numbering across consecutive paragraphs", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "One",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 3) });
    applyAndVm(editor, {
      type: "insertText",
      data: "Two",
      anchor: pos(1, 0),
    });

    applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(1, 3),
      attrs: {
        numberingKind: "decimal",
        numberingNumId: 7,
        numberingIlvl: 0,
      },
    });

    const para0 = paraModel(editor, 0);
    const para1 = paraModel(editor, 1);
    expect(para0?.numbering?.format).toBe("decimal");
    expect(para0?.numbering?.text).toBe("1.");
    expect(para1?.numbering?.format).toBe("decimal");
    expect(para1?.numbering?.text).toBe("2.");
    editor.free();
  });

  test("insertParagraph continues list numbering in the next paragraph", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "First item",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 10),
      attrs: {
        numberingKind: "decimal",
        numberingNumId: 3,
        numberingIlvl: 0,
      },
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 10) });

    const para0 = paraModel(editor, 0);
    const para1 = paraModel(editor, 1);
    expect(vm(editor).body.length).toBe(2);
    expect(para0?.numbering?.numId).toBe(3);
    expect(para0?.numbering?.text).toBe("1.");
    expect(para1?.numbering?.numId).toBe(3);
    expect(para1?.numbering?.format).toBe("decimal");
    expect(para1?.numbering?.text).toBe("2.");
    editor.free();
  });

  test("list numbering and bold formatting sync across CRDT peers", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();
    applyAndVm(editor1, {
      type: "insertText",
      data: "Shared item",
      anchor: pos(0, 0),
    });
    applyAndVm(editor1, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 11),
      attrs: {
        numberingKind: "bullet",
        numberingNumId: 5,
        numberingIlvl: 0,
      },
    });
    applyAndVm(editor1, {
      type: "setTextAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 6),
      attrs: { bold: true },
    });

    editor2.applyUpdate(editor1.encodeStateAsUpdate());

    const syncedPara = (vm(editor2).body.find((item) => {
      if (item.type !== "paragraph") return false;
      const paragraph = item as VmParagraph;
      return paragraph.runs.map((run) => run.text).join("") === "Shared item";
    }) ?? null) as VmParagraph | null;
    expect(syncedPara?.numbering?.format).toBe("bullet");
    expect(syncedPara?.numbering?.numId).toBe(5);
    expect(syncedPara?.numbering?.text.length ?? 0).toBeGreaterThan(0);
    expect(syncedPara?.runs[0]?.bold).toBe(true);

    editor1.free();
    editor2.free();
  });
});

describe("list controller semantics", () => {
  test("toggleList applies and clears bullet numbering", () => {
    const harness = makeControllerHarness();

    harness.controller.toggleList("bullet");
    let para = paraModel(harness.docEdit(), 0);
    expect(para?.numbering?.format).toBe("bullet");
    expect(harness.controller.getFormattingState().listKind).toBe("bullet");

    harness.controller.toggleList("bullet");
    para = paraModel(harness.docEdit(), 0);
    expect(para?.numbering).toBeUndefined();
    expect(harness.controller.getFormattingState().listKind).toBeUndefined();

    harness.destroy();
  });

  test("toggleList applies decimal numbering", () => {
    const harness = makeControllerHarness();

    harness.controller.toggleList("decimal");

    const para = paraModel(harness.docEdit(), 0);
    expect(para?.numbering?.format).toBe("decimal");
    expect(para?.numbering?.text).toBe("1.");
    expect(harness.controller.getFormattingState().listKind).toBe("decimal");

    harness.destroy();
  });

  test("Enter on a non-empty list item continues the list", () => {
    const harness = makeControllerHarness();

    harness.controller.toggleList("bullet");
    harness.input.emitInput({ type: "insertText", data: "Item 1" });
    harness.input.emitInput({ type: "insertParagraph" });

    const doc = harness.docEdit();
    const para0 = paraModel(doc, 0);
    const para1 = paraModel(doc, 1);
    expect(vm(doc).body.length).toBe(2);
    expect(para0?.numbering?.format).toBe("bullet");
    expect(para1?.numbering?.format).toBe("bullet");
    expect(para1?.numbering?.numId).toBe(para0?.numbering?.numId);

    harness.destroy();
  });

  test("Enter on an empty list item exits the list instead of splitting", () => {
    const harness = makeControllerHarness();

    harness.controller.toggleList("bullet");
    harness.input.emitInput({ type: "insertParagraph" });

    const doc = harness.docEdit();
    const para = paraModel(doc, 0);
    expect(vm(doc).body.length).toBe(1);
    expect(para?.numbering).toBeUndefined();
    expect(harness.controller.getFormattingState().listKind).toBeUndefined();

    harness.destroy();
  });

  test("Tab indents the current list item instead of inserting a tab token", () => {
    const harness = makeControllerHarness();

    harness.controller.toggleList("bullet");
    harness.input.emitInput({ type: "insertText", data: "Nested" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 6 });
    harness.input.emitInput({ type: "insertTab" });

    const para = paraModel(harness.docEdit(), 0);
    expect(para?.numbering?.format).toBe("bullet");
    expect(para?.numbering?.level).toBe(1);
    expect(para?.runs.some((run) => run.hasTab)).toBe(false);
    expect(harness.controller.getFormattingState().listLevel).toBe(1);

    harness.destroy();
  });

  test("Shift+Tab outdents the current list item", () => {
    const harness = makeControllerHarness();

    harness.controller.toggleList("decimal");
    harness.input.emitInput({ type: "insertText", data: "Nested" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 6 });
    harness.input.emitInput({ type: "insertTab" });
    harness.input.emitInput({ type: "insertTab", shift: true });

    const para = paraModel(harness.docEdit(), 0);
    expect(para?.numbering?.format).toBe("decimal");
    expect(para?.numbering?.level).toBe(0);
    expect(harness.controller.getFormattingState().listLevel).toBe(0);

    harness.destroy();
  });

  test("Tab on a non-list paragraph still inserts a tab token", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "Body" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 4 });
    harness.input.emitInput({ type: "insertTab" });

    const para = paraModel(harness.docEdit(), 0);
    expect(para?.numbering).toBeUndefined();
    expect(para?.runs.some((run) => run.hasTab)).toBe(true);

    harness.destroy();
  });

  test("mixed inline formatting state does not claim a uniform bold value", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "ABCD" });
    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 0, charOffset: 2 },
    });
    harness.input.emitSelectionChange();
    harness.controller.format("bold");

    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 0, charOffset: 4 },
    });
    harness.input.emitSelectionChange();

    const state = harness.controller.getFormattingState();
    expect(state.bold).toBeUndefined();

    harness.destroy();
  });

  test("mixed paragraph formatting state does not claim a uniform alignment", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "First" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "Second" });

    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 0, charOffset: 5 },
    });
    harness.input.emitSelectionChange();
    harness.controller.setParagraphStyle({ alignment: "center" });

    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 1, charOffset: 6 },
    });
    harness.input.emitSelectionChange();

    const state = harness.controller.getFormattingState();
    expect(state.alignment).toBeUndefined();

    harness.destroy();
  });

  test("decimal numbering renumbers after deleting the first item", () => {
    const harness = makeControllerHarness();

    harness.controller.toggleList("decimal");
    harness.input.emitInput({ type: "insertText", data: "One" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "Two" });
    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 1, charOffset: 0 },
    });
    harness.input.emitInput({ type: "deleteByCut" });

    const list = decimalListParagraphs(harness.docEdit());
    expect(list).toHaveLength(1);
    expectDecimalListParagraph(
      list[0],
      "Two",
      1,
      list[0]?.numbering?.numId ?? -1,
    );

    harness.destroy();
  });

  test("toggling numbering on a middle paragraph continues the surrounding decimal list", () => {
    const harness = makeControllerHarness();

    harness.controller.toggleList("decimal");
    harness.input.emitInput({ type: "insertText", data: "One" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "Gap" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "Two" });
    harness.renderer.setCursor({ bodyIndex: 2, charOffset: 3 });
    harness.input.emitSelectionChange();
    harness.controller.toggleList("decimal");
    harness.renderer.setCursor({ bodyIndex: 1, charOffset: 3 });
    harness.input.emitSelectionChange();
    harness.controller.toggleList("decimal");

    const list = decimalListParagraphs(harness.docEdit());
    expect(list).toHaveLength(3);
    const numId = list[0]?.numbering?.numId ?? -1;
    expectDecimalListParagraph(list[0], "One", 1, numId);
    expectDecimalListParagraph(list[1], "Gap", 2, numId);
    expectDecimalListParagraph(list[2], "Two", 3, numId);

    harness.destroy();
  });
});

describe("controller table/image semantics", () => {
  test("replaceBlank clears prior rich content state", () => {
    const harness = makeControllerHarness();

    expect(harness.controller.insertTable(2, 2)).toBe(true);
    expect(
      harness.controller.insertInlineImage({
        dataUri: `data:image/png;base64,${TINY_PNG_BASE64}`,
        widthPt: 24,
        heightPt: 18,
      }),
    ).toBe(true);

    harness.controller.replaceBlank();

    const model = vm(harness.docEdit()) as unknown as {
      images?: Array<unknown>;
      body?: Array<{ type?: string }>;
    };
    expect(model.images ?? []).toHaveLength(0);
    expect((model.body ?? []).some((item) => item.type === "table")).toBe(
      false,
    );
    expect(firstParaText(harness.docEdit())).toBe("");
    expect(harness.controller.getActiveTableCell()).toBeNull();

    harness.destroy();
  });

  test("replace loads a new document and drops prior rich content state", () => {
    const harness = makeControllerHarness();
    const replacement = DocEdit.blank();

    expect(harness.controller.insertTable(2, 2)).toBe(true);
    replacement.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "Replacement doc",
        anchor: pos(0, 0),
      }),
    );
    const replacementBytes = replacement.save();
    const ms = harness.controller.replace(replacementBytes);

    expect(ms).toBeGreaterThanOrEqual(0);
    expect(firstParaText(harness.docEdit())).toBe("Replacement doc");
    const model = vm(harness.docEdit()) as unknown as {
      images?: Array<unknown>;
      body?: Array<{ type?: string }>;
    };
    expect(model.images ?? []).toHaveLength(0);
    expect((model.body ?? []).some((item) => item.type === "table")).toBe(
      false,
    );
    expect(harness.controller.getActiveTableCell()).toBeNull();

    replacement.free();
    harness.destroy();
  });

  test("insertTable opens the first cell and save/reload preserves cell edits", () => {
    const harness = makeControllerHarness();

    expect(harness.controller.insertTable(2, 2)).toBe(true);
    expect(harness.controller.getActiveTableCell()).toMatchObject({
      row: 0,
      col: 0,
      rowCount: 2,
      colCount: 2,
    });
    expect(harness.controller.setActiveTableCellText("A1")).toBe(true);
    expect(harness.controller.moveActiveTableCell(0, 1)).toBe(true);
    expect(harness.controller.getActiveTableCell()).toMatchObject({
      row: 0,
      col: 1,
    });
    expect(harness.controller.setActiveTableCellText("B1")).toBe(true);
    harness.controller.clearActiveTableCell();

    const table = (
      vm(harness.docEdit()) as unknown as {
        body?: Array<{
          type?: string;
          rows?: Array<{ cells?: Array<{ text?: string }> }>;
        }>;
      }
    ).body?.find((item) => item.type === "table");
    expect(table?.rows?.[0]?.cells?.[0]?.text).toBe("A1");
    expect(table?.rows?.[0]?.cells?.[1]?.text).toBe("B1");

    const reopened = new DocEdit(harness.controller.save());
    const reopenedTable = (
      vm(reopened) as unknown as {
        body?: Array<{
          type?: string;
          rows?: Array<{ cells?: Array<{ text?: string }> }>;
        }>;
      }
    ).body?.find((item) => item.type === "table");
    expect(reopenedTable?.rows?.[0]?.cells?.[0]?.text).toBe("A1");
    expect(reopenedTable?.rows?.[0]?.cells?.[1]?.text).toBe("B1");

    reopened.free();
    harness.destroy();
  });

  test("table row/column insert-remove preserves structure on save/reload", () => {
    const harness = makeControllerHarness();

    expect(harness.controller.insertTable(2, 2)).toBe(true);
    expect(harness.controller.setActiveTableCellText("A1")).toBe(true);
    expect(harness.controller.insertTableRow()).toBe(true);
    expect(harness.controller.getActiveTableCell()).toMatchObject({
      row: 1,
      col: 0,
      rowCount: 3,
      colCount: 2,
    });
    expect(harness.controller.setActiveTableCellText("A2")).toBe(true);
    expect(harness.controller.insertTableColumn()).toBe(true);
    expect(harness.controller.getActiveTableCell()).toMatchObject({
      row: 1,
      col: 1,
      rowCount: 3,
      colCount: 3,
    });
    expect(harness.controller.setActiveTableCellText("B2")).toBe(true);
    expect(harness.controller.removeTableColumn()).toBe(true);
    expect(harness.controller.getActiveTableCell()).toMatchObject({
      row: 1,
      col: 1,
      rowCount: 3,
      colCount: 2,
    });
    expect(harness.controller.removeTableRow()).toBe(true);
    harness.controller.clearActiveTableCell();

    const table = (
      vm(harness.docEdit()) as unknown as {
        body?: Array<{
          type?: string;
          rows?: Array<{ cells?: Array<{ text?: string }> }>;
        }>;
      }
    ).body?.find((item) => item.type === "table");
    expect(table?.rows).toHaveLength(2);
    expect(table?.rows?.[0]?.cells).toHaveLength(2);
    expect(table?.rows?.[0]?.cells?.[0]?.text).toBe("A1");

    const reopened = new DocEdit(harness.controller.save());
    const reopenedTable = (
      vm(reopened) as unknown as {
        body?: Array<{
          type?: string;
          rows?: Array<{ cells?: Array<{ text?: string }> }>;
        }>;
      }
    ).body?.find((item) => item.type === "table");
    expect(reopenedTable?.rows).toHaveLength(2);
    expect(reopenedTable?.rows?.[0]?.cells).toHaveLength(2);
    expect(reopenedTable?.rows?.[0]?.cells?.[0]?.text).toBe("A1");

    reopened.free();
    harness.destroy();
  });

  test("insertInlineImage saves and reloads as an inline image run", () => {
    const harness = makeControllerHarness();

    expect(
      harness.controller.insertInlineImage({
        dataUri: `data:image/png;base64,${TINY_PNG_BASE64}`,
        widthPt: 24,
        heightPt: 18,
        name: "pixel.png",
        description: "pixel",
      }),
    ).toBe(true);

    const model = vm(harness.docEdit()) as unknown as {
      images?: Array<{ contentType?: string }>;
    };
    expect(model.images).toBeDefined();
    expect(model.images ?? []).toHaveLength(1);
    expect(model.images?.[0]?.contentType).toBe("image/png");
    const inlineRun = (paraModel(harness.docEdit(), 0)?.runs ?? []).find(
      (run) =>
        (
          run as {
            inlineImage?: {
              imageIndex?: number;
              widthPt?: number;
              heightPt?: number;
            };
          }
        ).inlineImage != null,
    ) as
      | {
          inlineImage?: {
            imageIndex?: number;
            widthPt?: number;
            heightPt?: number;
          };
        }
      | undefined;
    expect(inlineRun?.inlineImage?.imageIndex).toBe(0);
    expect(inlineRun?.inlineImage?.widthPt).toBe(24);
    expect(inlineRun?.inlineImage?.heightPt).toBe(18);

    const reopened = new DocEdit(harness.controller.save());
    const reopenedModel = vm(reopened) as unknown as {
      images?: Array<{ contentType?: string }>;
    };
    expect(reopenedModel.images).toBeDefined();
    expect(reopenedModel.images ?? []).toHaveLength(1);
    expect(reopenedModel.images?.[0]?.contentType).toBe("image/png");
    const reopenedRun = (paraModel(reopened, 0)?.runs ?? []).find(
      (run) =>
        (
          run as {
            inlineImage?: {
              imageIndex?: number;
              widthPt?: number;
              heightPt?: number;
            };
          }
        ).inlineImage != null,
    ) as
      | {
          inlineImage?: {
            imageIndex?: number;
            widthPt?: number;
            heightPt?: number;
          };
        }
      | undefined;
    expect(reopenedRun?.inlineImage?.imageIndex).toBe(0);
    expect(reopenedRun?.inlineImage?.widthPt).toBe(24);
    expect(reopenedRun?.inlineImage?.heightPt).toBe(18);

    reopened.free();
    harness.destroy();
  });

  test("selected inline image can be resized and save/reload preserves dimensions", () => {
    const harness = makeControllerHarness();

    expect(
      harness.controller.insertInlineImage({
        dataUri: `data:image/png;base64,${TINY_PNG_BASE64}`,
        widthPt: 24,
        heightPt: 18,
        name: "pixel.png",
        description: "pixel",
      }),
    ).toBe(true);

    harness.input.emitPointerDown({
      x: 80,
      y: 12,
      clickCount: 1,
      shift: false,
    });

    expect(harness.controller.getSelectedInlineImage()).toMatchObject({
      bodyIndex: 0,
      charOffset: 0,
      widthPt: 24,
      heightPt: 18,
    });
    expect(harness.controller.resizeSelectedInlineImage(48, 36)).toBe(true);

    const resizedRun = (paraModel(harness.docEdit(), 0)?.runs ?? []).find(
      (run) =>
        (
          run as {
            inlineImage?: {
              widthPt?: number;
              heightPt?: number;
            };
          }
        ).inlineImage != null,
    ) as
      | {
          inlineImage?: {
            widthPt?: number;
            heightPt?: number;
          };
        }
      | undefined;
    expect(resizedRun?.inlineImage?.widthPt).toBe(48);
    expect(resizedRun?.inlineImage?.heightPt).toBe(36);

    const reopened = new DocEdit(harness.controller.save());
    const reopenedRun = (paraModel(reopened, 0)?.runs ?? []).find(
      (run) =>
        (
          run as {
            inlineImage?: {
              widthPt?: number;
              heightPt?: number;
            };
          }
        ).inlineImage != null,
    ) as
      | {
          inlineImage?: {
            widthPt?: number;
            heightPt?: number;
          };
        }
      | undefined;
    expect(reopenedRun?.inlineImage?.widthPt).toBe(48);
    expect(reopenedRun?.inlineImage?.heightPt).toBe(36);

    reopened.free();
    harness.destroy();
  });

  test("selected inline image block alignment persists through save/reload", () => {
    const harness = makeControllerHarness();

    expect(
      harness.controller.insertInlineImage({
        dataUri: `data:image/png;base64,${TINY_PNG_BASE64}`,
        widthPt: 24,
        heightPt: 18,
      }),
    ).toBe(true);

    harness.input.emitPointerDown({
      x: 80,
      y: 12,
      clickCount: 1,
      shift: false,
    });
    expect(harness.controller.setSelectedInlineImageAlignment("center")).toBe(
      true,
    );

    const paragraph = paraModel(harness.docEdit(), 0);
    expect(paragraph?.alignment).toBe("center");

    const reopened = new DocEdit(harness.controller.save());
    expect(paraModel(reopened, 0)?.alignment).toBe("center");

    reopened.free();
    harness.destroy();
  });

  test("undo removes a controller-inserted table", () => {
    const harness = makeControllerHarness();

    expect(harness.controller.insertTable(2, 2)).toBe(true);
    expect(
      vm(harness.docEdit()).body.some((item) => item.type === "table"),
    ).toBe(true);

    harness.input.emitInput({ type: "historyUndo" });

    expect(
      vm(harness.docEdit()).body.some((item) => item.type === "table"),
    ).toBe(false);

    harness.destroy();
  });

  test("undo removes a controller-inserted inline image", () => {
    const harness = makeControllerHarness();

    expect(
      harness.controller.insertInlineImage({
        dataUri: `data:image/png;base64,${TINY_PNG_BASE64}`,
        widthPt: 24,
        heightPt: 18,
        name: "pixel.png",
        description: "pixel",
      }),
    ).toBe(true);

    const modelBeforeUndo = vm(harness.docEdit()) as unknown as {
      images?: Array<{ contentType?: string }>;
    };
    expect(modelBeforeUndo.images ?? []).toHaveLength(1);

    harness.input.emitInput({ type: "historyUndo" });

    const modelAfterUndo = vm(harness.docEdit()) as unknown as {
      images?: Array<{ contentType?: string }>;
    };
    expect(modelAfterUndo.images ?? []).toHaveLength(0);
    const remainingInlineRuns = (
      paraModel(harness.docEdit(), 0)?.runs ?? []
    ).filter(
      (run) =>
        (
          run as {
            inlineImage?: {
              imageIndex?: number;
            };
          }
        ).inlineImage != null,
    );
    expect(remainingInlineRuns).toHaveLength(0);

    harness.destroy();
  });

  test("clicking an inline image selects its token and delete removes it", () => {
    const harness = makeControllerHarness();

    expect(
      harness.controller.insertInlineImage({
        dataUri: `data:image/png;base64,${TINY_PNG_BASE64}`,
        widthPt: 24,
        heightPt: 18,
        name: "pixel.png",
        description: "pixel",
      }),
    ).toBe(true);

    harness.input.emitPointerDown({
      x: 80,
      y: 12,
      clickCount: 1,
      shift: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 0, charOffset: 1 },
    });

    harness.input.emitInput({ type: "deleteByCut" });
    const remainingInlineRuns = (
      paraModel(harness.docEdit(), 0)?.runs ?? []
    ).filter(
      (run) =>
        (
          run as {
            inlineImage?: {
              imageIndex?: number;
            };
          }
        ).inlineImage != null,
    );
    expect(remainingInlineRuns).toHaveLength(0);

    harness.destroy();
  });

  test("pasting an image inserts an inline image run", async () => {
    const harness = makeControllerHarness();
    const file = new File(
      [Uint8Array.from(atob(TINY_PNG_BASE64), (ch) => ch.charCodeAt(0))],
      "pasted.png",
      { type: "image/png" },
    );

    await harness.input.emitPasteImage(file);

    const model = vm(harness.docEdit()) as unknown as {
      images?: Array<{ contentType?: string }>;
    };
    expect(model.images ?? []).toHaveLength(1);
    expect(model.images?.[0]?.contentType).toBe("image/png");
    const inlineRun = (paraModel(harness.docEdit(), 0)?.runs ?? []).find(
      (run) =>
        (
          run as {
            inlineImage?: {
              imageIndex?: number;
            };
          }
        ).inlineImage != null,
    ) as { inlineImage?: { imageIndex?: number } } | undefined;
    expect(inlineRun?.inlineImage?.imageIndex).toBe(0);

    harness.destroy();
  });
});

describe("rich content collaboration", () => {
  test("diff-based sync preserves table cell edits", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    editor1.applyIntent(
      JSON.stringify({
        type: "insertTable",
        anchor: pos(0, 0),
        rows: 2,
        columns: 2,
      }),
    );
    const tableBodyIndex = vm(editor1).body.findIndex(
      (item) => item.type === "table",
    );
    expect(tableBodyIndex).toBeGreaterThanOrEqual(0);
    editor1.applyIntent(
      JSON.stringify({
        type: "setTableCellText",
        bodyIndex: tableBodyIndex,
        row: 0,
        col: 0,
        text: "A1",
      }),
    );
    editor1.applyIntent(
      JSON.stringify({
        type: "setTableCellText",
        bodyIndex: tableBodyIndex,
        row: 1,
        col: 1,
        text: "B2",
      }),
    );

    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    const diff = editor1.encodeDiff(editor2.encodeStateVector());
    editor2.applyUpdate(diff);

    const table = (
      vm(editor2) as unknown as {
        body?: Array<{
          type?: string;
          rows?: Array<{ cells?: Array<{ text?: string }> }>;
        }>;
      }
    ).body?.find((item) => item.type === "table");
    expect(table?.rows?.[0]?.cells?.[0]?.text).toBe("A1");
    expect(table?.rows?.[1]?.cells?.[1]?.text).toBe("B2");

    const reopened = new DocEdit(editor2.save());
    const reopenedTable = (
      vm(reopened) as unknown as {
        body?: Array<{
          type?: string;
          rows?: Array<{ cells?: Array<{ text?: string }> }>;
        }>;
      }
    ).body?.find((item) => item.type === "table");
    expect(reopenedTable?.rows?.[0]?.cells?.[0]?.text).toBe("A1");
    expect(reopenedTable?.rows?.[1]?.cells?.[1]?.text).toBe("B2");

    reopened.free();
    editor1.free();
    editor2.free();
  });

  test("diff-based sync preserves inline image payloads on a peer export path", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    editor1.applyIntent(
      JSON.stringify({
        type: "insertInlineImage",
        anchor: pos(0, 0),
        focus: pos(0, 0),
        data_uri: `data:image/png;base64,${TINY_PNG_BASE64}`,
        width_pt: 24,
        height_pt: 18,
        name: "pixel.png",
        description: "pixel",
      }),
    );

    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    editor2.applyUpdate(editor1.encodeDiff(editor2.encodeStateVector()));

    const syncedModel = vm(editor2) as unknown as {
      images?: Array<{ contentType?: string }>;
    };
    expect(syncedModel.images).toBeDefined();
    expect(syncedModel.images ?? []).toHaveLength(1);
    expect(syncedModel.images?.[0]?.contentType).toBe("image/png");
    const syncedRun = (paraModel(editor2, 0)?.runs ?? []).find(
      (run) =>
        (
          run as {
            inlineImage?: {
              imageIndex?: number;
              widthPt?: number;
              heightPt?: number;
            };
          }
        ).inlineImage != null,
    ) as
      | {
          inlineImage?: {
            imageIndex?: number;
            widthPt?: number;
            heightPt?: number;
          };
        }
      | undefined;
    expect(syncedRun?.inlineImage?.imageIndex).toBe(0);
    expect(syncedRun?.inlineImage?.widthPt).toBe(24);
    expect(syncedRun?.inlineImage?.heightPt).toBe(18);

    const reopened = new DocEdit(editor2.save());
    const reopenedModel = vm(reopened) as unknown as {
      images?: Array<{ contentType?: string }>;
    };
    expect(reopenedModel.images).toBeDefined();
    expect(reopenedModel.images ?? []).toHaveLength(1);
    expect(reopenedModel.images?.[0]?.contentType).toBe("image/png");
    const reopenedRun = (paraModel(reopened, 0)?.runs ?? []).find(
      (run) =>
        (
          run as {
            inlineImage?: {
              imageIndex?: number;
              widthPt?: number;
              heightPt?: number;
            };
          }
        ).inlineImage != null,
    ) as
      | {
          inlineImage?: {
            imageIndex?: number;
            widthPt?: number;
            heightPt?: number;
          };
        }
      | undefined;
    expect(reopenedRun?.inlineImage?.imageIndex).toBe(0);
    expect(reopenedRun?.inlineImage?.widthPt).toBe(24);
    expect(reopenedRun?.inlineImage?.heightPt).toBe(18);

    reopened.free();
    editor1.free();
    editor2.free();
  });

  test("controller-authored table cell edits sync to a peer", () => {
    const harness = makeControllerHarness();
    const peer = DocEdit.blank();
    peer.applyUpdate(harness.docEdit().encodeStateAsUpdate());

    expect(harness.controller.insertTable(2, 2)).toBe(true);
    syncEditors(harness.docEdit(), peer);
    expect(harness.controller.setActiveTableCellText("A1")).toBe(true);
    syncEditors(harness.docEdit(), peer);
    expect(harness.controller.moveActiveTableCell(0, 1)).toBe(true);
    expect(harness.controller.setActiveTableCellText("B1")).toBe(true);
    syncEditors(harness.docEdit(), peer);
    fullResyncEditors(harness.docEdit(), peer);

    const peerTable = (
      vm(peer) as unknown as {
        body?: Array<{
          type?: string;
          rows?: Array<{ cells?: Array<{ text?: string }> }>;
        }>;
      }
    ).body?.find((item) => item.type === "table");
    expect(peerTable?.rows?.[0]?.cells?.[0]?.text).toBe("A1");
    expect(peerTable?.rows?.[0]?.cells?.[1]?.text).toBe("B1");
    expect(JSON.stringify(vm(harness.docEdit()).body)).toBe(
      JSON.stringify(vm(peer).body),
    );

    harness.destroy();
    peer.free();
  });

  test("controller-authored table row/column changes sync to a peer", () => {
    const harness = makeControllerHarness();
    const peer = DocEdit.blank();
    peer.applyUpdate(harness.docEdit().encodeStateAsUpdate());

    expect(harness.controller.insertTable(2, 2)).toBe(true);
    syncEditors(harness.docEdit(), peer);
    expect(harness.controller.insertTableRow()).toBe(true);
    syncEditors(harness.docEdit(), peer);
    expect(harness.controller.insertTableColumn()).toBe(true);
    syncEditors(harness.docEdit(), peer);
    expect(harness.controller.removeTableColumn()).toBe(true);
    syncEditors(harness.docEdit(), peer);
    expect(harness.controller.removeTableRow()).toBe(true);
    syncEditors(harness.docEdit(), peer);
    fullResyncEditors(harness.docEdit(), peer);

    const peerTable = (
      vm(peer) as unknown as {
        body?: Array<{
          type?: string;
          rows?: Array<{ cells?: Array<{ text?: string }> }>;
        }>;
      }
    ).body?.find((item) => item.type === "table");
    expect(peerTable?.rows).toHaveLength(2);
    expect(peerTable?.rows?.[0]?.cells).toHaveLength(2);
    expect(JSON.stringify(vm(harness.docEdit()).body)).toBe(
      JSON.stringify(vm(peer).body),
    );

    harness.destroy();
    peer.free();
  });

  test("controller-authored table row/column edits sync to a peer", () => {
    const harness = makeControllerHarness();
    const peer = DocEdit.blank();
    peer.applyUpdate(harness.docEdit().encodeStateAsUpdate());

    expect(harness.controller.insertTable(2, 2)).toBe(true);
    expect(harness.controller.setActiveTableCellText("A1")).toBe(true);
    syncEditors(harness.docEdit(), peer);

    expect(harness.controller.insertTableRow()).toBe(true);
    expect(harness.controller.insertTableColumn()).toBe(true);
    syncEditors(harness.docEdit(), peer);

    expect(harness.controller.removeTableColumn()).toBe(true);
    expect(harness.controller.removeTableRow()).toBe(true);
    syncEditors(harness.docEdit(), peer);
    fullResyncEditors(harness.docEdit(), peer);

    const peerTable = (
      vm(peer) as unknown as {
        body?: Array<{
          type?: string;
          rows?: Array<{ cells?: Array<{ text?: string }> }>;
        }>;
      }
    ).body?.find((item) => item.type === "table");
    expect(peerTable?.rows).toHaveLength(2);
    expect(peerTable?.rows?.[0]?.cells).toHaveLength(2);
    expect(peerTable?.rows?.[0]?.cells?.[0]?.text).toBe("A1");
    expect(JSON.stringify(vm(harness.docEdit()).body)).toBe(
      JSON.stringify(vm(peer).body),
    );

    harness.destroy();
    peer.free();
  });

  test("controller-authored inline image insert and delete sync to a peer", () => {
    const harness = makeControllerHarness();
    const peer = DocEdit.blank();
    peer.applyUpdate(harness.docEdit().encodeStateAsUpdate());

    expect(
      harness.controller.insertInlineImage({
        dataUri: `data:image/png;base64,${TINY_PNG_BASE64}`,
        widthPt: 24,
        heightPt: 18,
        name: "pixel.png",
        description: "pixel",
      }),
    ).toBe(true);
    syncEditors(harness.docEdit(), peer);

    let peerModel = vm(peer) as unknown as {
      images?: Array<{ contentType?: string }>;
      body?: Array<{
        type?: string;
        runs?: Array<{ inlineImage?: unknown }>;
      }>;
    };
    expect(peerModel.images).toBeDefined();
    expect(peerModel.images ?? []).toHaveLength(1);
    expect(peerModel.images?.[0]?.contentType).toBe("image/png");
    let peerRun = (peerModel.body ?? [])
      .filter((item) => item.type === "paragraph")
      .flatMap((item) => item.runs ?? [])
      .find(
        (run) =>
          (
            run as {
              inlineImage?: {
                imageIndex?: number;
              };
            }
          ).inlineImage != null,
      ) as { inlineImage?: { imageIndex?: number } } | undefined;
    expect(peerRun?.inlineImage?.imageIndex).toBe(0);

    harness.input.emitPointerDown({
      x: 80,
      y: 12,
      clickCount: 1,
      shift: false,
    });
    harness.input.emitInput({ type: "deleteByCut" });
    syncEditors(harness.docEdit(), peer);
    fullResyncEditors(harness.docEdit(), peer);

    peerModel = vm(peer) as unknown as {
      images?: Array<{ contentType?: string }>;
      body?: Array<{
        type?: string;
        runs?: Array<{ inlineImage?: unknown }>;
      }>;
    };
    peerRun = (peerModel.body ?? [])
      .filter((item) => item.type === "paragraph")
      .flatMap((item) => item.runs ?? [])
      .find(
        (run) =>
          (
            run as {
              inlineImage?: {
                imageIndex?: number;
              };
            }
          ).inlineImage != null,
      ) as { inlineImage?: { imageIndex?: number } } | undefined;
    expect(peerRun).toBeUndefined();
    expect(
      (peerModel.body ?? [])
        .flatMap((item) => item.runs ?? [])
        .some((run) => run.inlineImage != null),
    ).toBe(false);

    harness.destroy();
    peer.free();
  });

  test("controller-authored inline image resize and alignment sync to a peer", () => {
    const harness = makeControllerHarness();
    const peer = DocEdit.blank();
    peer.applyUpdate(harness.docEdit().encodeStateAsUpdate());

    expect(
      harness.controller.insertInlineImage({
        dataUri: `data:image/png;base64,${TINY_PNG_BASE64}`,
        widthPt: 24,
        heightPt: 18,
      }),
    ).toBe(true);
    syncEditors(harness.docEdit(), peer);
    harness.input.emitPointerDown({
      x: 80,
      y: 12,
      clickCount: 1,
      shift: false,
    });
    expect(harness.controller.resizeSelectedInlineImage(48, 36)).toBe(true);
    expect(harness.controller.setSelectedInlineImageAlignment("right")).toBe(
      true,
    );
    syncEditors(harness.docEdit(), peer);
    fullResyncEditors(harness.docEdit(), peer);

    const peerParagraphs = vm(peer).body.filter(
      (item): item is VmParagraph => item.type === "paragraph",
    );
    const peerParagraphWithImage = peerParagraphs.find((paragraph) =>
      (paragraph.runs ?? []).some((run) => run.inlineImage != null),
    );
    const peerRun = (peerParagraphWithImage?.runs ?? []).find(
      (run) =>
        (
          run as {
            inlineImage?: {
              widthPt?: number;
              heightPt?: number;
            };
          }
        ).inlineImage != null,
    ) as
      | {
          inlineImage?: {
            widthPt?: number;
            heightPt?: number;
          };
        }
      | undefined;
    expect(peerRun?.inlineImage?.widthPt).toBe(48);
    expect(peerRun?.inlineImage?.heightPt).toBe(36);
    expect(peerParagraphWithImage?.alignment).toBe("right");
    expect(JSON.stringify(vm(harness.docEdit()).body)).toBe(
      JSON.stringify(vm(peer).body),
    );

    harness.destroy();
    peer.free();
  });
});

describe("controller selection semantics", () => {
  test("double click selects the clicked word on canvas", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta" });
    harness.renderer.nextHit = { bodyIndex: 0, charOffset: 7 };
    harness.input.emitPointerDown({
      x: 100,
      y: 24,
      clickCount: 2,
      shift: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 6 },
      focus: { bodyIndex: 0, charOffset: 10 },
    });

    harness.destroy();
  });

  test("triple click selects the entire paragraph on canvas", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta" });
    harness.renderer.nextHit = { bodyIndex: 0, charOffset: 4 };
    harness.input.emitPointerDown({
      x: 100,
      y: 24,
      clickCount: 3,
      shift: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 0, charOffset: 10 },
    });

    harness.destroy();
  });
});

describe("controller clipboard semantics", () => {
  test("copy request returns selected plain text across paragraphs", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "abc" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "def" });
    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 1 },
      focus: { bodyIndex: 1, charOffset: 2 },
    });

    expect(harness.input.requestCopyText()).toBe("bc\nde");

    harness.destroy();
  });

  test("cut request returns selected plain text across paragraphs", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "abc" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "def" });
    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 1 },
      focus: { bodyIndex: 1, charOffset: 2 },
    });

    expect(harness.input.requestCutText()).toBe("bc\nde");

    harness.destroy();
  });

  test("copy request returns rich html for styled selection", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 1 },
      focus: { bodyIndex: 0, charOffset: 4 },
    });
    harness.controller.setTextStyle({
      bold: true,
      color: "#112233",
      highlight: "yellow",
      hyperlink: "https://example.com",
    });

    const html = harness.input.requestCopyHtml();
    expect(html).toContain("<p");
    expect(html).toContain('href="https://example.com"');
    expect(html).toContain("font-weight: 700");
    expect(html).toContain("color: #112233");
    expect(html).toContain("background-color: #ffff00");
    expect(html).toContain(">lph<");

    harness.destroy();
  });

  test("cut request returns rich html across paragraphs", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "abc" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "def" });
    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 1 },
      focus: { bodyIndex: 1, charOffset: 2 },
    });
    harness.controller.setParagraphStyle({ alignment: "center" });

    const html = harness.input.requestCutHtml();
    expect(html).toContain("<p");
    expect(html).toContain("text-align: center");
    expect(html).toContain(">bc<");
    expect(html).toContain(">de<");

    harness.destroy();
  });
});

describe("controller rich html paste", () => {
  test("rich html paste preserves paragraphs, inline styles, heading, and alignment", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({
      type: "insertFromPaste",
      data: "Styled Link\nHeading",
      html: '<p><span style="font-family: \'Noto Serif\'; font-size: 18pt; color: #3366ff; background-color: yellow"><strong>Styled</strong></span> <a href="https://example.com"><u>Link</u></a></p><h2 style="text-align: center">Heading</h2>',
    });

    const model = vm(harness.docEdit());
    expect(model.body.length).toBeGreaterThanOrEqual(2);

    const firstParagraphRuns = paraRuns(harness.docEdit(), 0);
    const styledRun = firstParagraphRuns.find((run) => run.text === "Styled");
    const linkedRun = firstParagraphRuns.find((run) => run.text === "Link");
    expect(paraText(harness.docEdit(), 0)).toBe("Styled Link");
    expect(styledRun?.bold).toBe(true);
    expect(styledRun?.fontFamily).toBe("Noto Serif");
    expect(styledRun?.fontSizePt).toBe(18);
    expect(styledRun?.color).toBe("3366FF");
    expect(styledRun?.highlight).toBe("yellow");
    expect(linkedRun?.underline).toBe(true);
    expect(linkedRun?.hyperlink).toBe("https://example.com");

    const secondParagraph = paraModel(harness.docEdit(), 1)!;
    expect(paraText(harness.docEdit(), 1)).toBe("Heading");
    expect(secondParagraph.headingLevel).toBe(2);
    expect(secondParagraph.alignment).toBe("center");

    harness.destroy();
  });

  test("rich html paste keeps straightforward list cues", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({
      type: "insertFromPaste",
      data: "One\nTwo",
      html: "<ul><li>One</li><li><strong>Two</strong></li></ul>",
    });

    const paragraph0 = paraModel(harness.docEdit(), 0)!;
    const paragraph1 = paraModel(harness.docEdit(), 1)!;
    const paragraph1Runs = paraRuns(harness.docEdit(), 1);
    expect(paraText(harness.docEdit(), 0)).toBe("One");
    expect(paraText(harness.docEdit(), 1)).toBe("Two");
    expect(paragraph0.numbering?.format).toBe("bullet");
    expect(paragraph1.numbering?.format).toBe("bullet");
    expect(paragraph0.numbering?.numId).toBe(paragraph1.numbering?.numId);
    expect(paragraph1Runs.find((run) => run.text === "Two")?.bold).toBe(true);

    harness.destroy();
  });

  test("save roundtrip preserves rich html pasted structure and styling", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({
      type: "insertFromPaste",
      data: "Styled Link\nHeading\nOne\nTwo",
      html: '<p><span style="font-family: \'Noto Serif\'; font-size: 18pt; color: #3366ff; background-color: yellow"><strong>Styled</strong></span> <a href="https://example.com"><u>Link</u></a></p><h2 style="text-align: center">Heading</h2><ul><li>One</li><li><strong>Two</strong></li></ul>',
    });

    const bytes = harness.controller.save();
    const editor2 = new DocEdit(bytes);

    const firstParagraphRuns = paraRuns(editor2, 0);
    const styledRun = firstParagraphRuns.find((run) => run.text === "Styled");
    const linkedRun = firstParagraphRuns.find((run) => run.text === "Link");
    expect(paraText(editor2, 0)).toBe("Styled Link");
    expect(styledRun?.bold).toBe(true);
    expect(styledRun?.fontFamily).toBe("Noto Serif");
    expect(styledRun?.fontSizePt).toBe(18);
    expect(styledRun?.color).toBe("3366FF");
    expect(styledRun?.highlight).toBe("yellow");
    expect(linkedRun?.underline).toBe(true);
    expect(linkedRun?.hyperlink).toBe("https://example.com");

    const headingParagraph = paraModel(editor2, 1)!;
    expect(paraText(editor2, 1)).toBe("Heading");
    expect(headingParagraph.headingLevel).toBe(2);
    expect(headingParagraph.alignment).toBe("center");

    const bullet0 = paraModel(editor2, 2)!;
    const bullet1 = paraModel(editor2, 3)!;
    expect(paraText(editor2, 2)).toBe("One");
    expect(paraText(editor2, 3)).toBe("Two");
    expect(bullet0.numbering?.format).toBe("bullet");
    expect(bullet1.numbering?.format).toBe("bullet");
    expect(bullet0.numbering?.numId).toBe(bullet1.numbering?.numId);
    expect(paraRuns(editor2, 3).find((run) => run.text === "Two")?.bold).toBe(
      true,
    );

    harness.destroy();
    editor2.free();
  });
});

describe("controller navigation semantics", () => {
  test("click places the caret at the hit-tested position", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.renderer.nextHit = { bodyIndex: 0, charOffset: 2 };
    harness.input.emitPointerDown({
      x: 64,
      y: 24,
      clickCount: 1,
      shift: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 2 },
      focus: { bodyIndex: 0, charOffset: 2 },
    });

    harness.destroy();
  });

  test("Cmd/Ctrl+A selects the whole document", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "abc" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "def" });
    harness.input.emitShortcut("a");

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 1, charOffset: 3 },
    });

    harness.destroy();
  });

  test("ArrowRight moves caret forward by one character", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 1 });
    harness.input.emitNavigate({
      key: "ArrowRight",
      shift: false,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 2 },
      focus: { bodyIndex: 0, charOffset: 2 },
    });

    harness.destroy();
  });

  test("ArrowLeft moves caret backward across paragraph boundaries", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "abc" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "def" });
    harness.renderer.setCursor({ bodyIndex: 1, charOffset: 0 });
    harness.input.emitNavigate({
      key: "ArrowLeft",
      shift: false,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 3 },
      focus: { bodyIndex: 0, charOffset: 3 },
    });

    harness.destroy();
  });

  test("Home moves caret to paragraph start", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 7 });
    harness.input.emitNavigate({
      key: "Home",
      shift: false,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 0, charOffset: 0 },
    });

    harness.destroy();
  });

  test("End moves caret to paragraph end", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 2 });
    harness.input.emitNavigate({
      key: "End",
      shift: false,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 10 },
      focus: { bodyIndex: 0, charOffset: 10 },
    });

    harness.destroy();
  });

  test("Ctrl/Cmd+A selects the whole document", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "beta" });
    harness.input.emitShortcut("a");

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 1, charOffset: 4 },
    });

    harness.destroy();
  });

  test("Ctrl/Alt+ArrowLeft moves caret to the prior word boundary", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 10 });
    harness.input.emitNavigate({
      key: "ArrowLeft",
      shift: false,
      meta: false,
      alt: true,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 6 },
      focus: { bodyIndex: 0, charOffset: 6 },
    });

    harness.destroy();
  });

  test("Ctrl/Alt+ArrowRight moves caret to the next word boundary", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 2 });
    harness.input.emitNavigate({
      key: "ArrowRight",
      shift: false,
      meta: false,
      alt: true,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 6 },
      focus: { bodyIndex: 0, charOffset: 6 },
    });

    harness.destroy();
  });

  test("Ctrl/Alt+ArrowLeft crosses paragraph boundaries by word", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 1, charOffset: 0 });
    harness.input.emitNavigate({
      key: "ArrowLeft",
      shift: false,
      meta: false,
      alt: true,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 5 },
      focus: { bodyIndex: 0, charOffset: 5 },
    });

    harness.destroy();
  });

  test("Shift+ArrowRight extends selection forward", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 1 });
    harness.input.emitNavigate({
      key: "ArrowRight",
      shift: true,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 1 },
      focus: { bodyIndex: 0, charOffset: 2 },
    });

    harness.destroy();
  });

  test("Shift+ArrowLeft extends backward selection", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 3 });
    harness.input.emitNavigate({
      key: "ArrowLeft",
      shift: true,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 3 },
      focus: { bodyIndex: 0, charOffset: 2 },
    });

    harness.destroy();
  });

  test("Shift+ArrowLeft extends backward selection across paragraph boundaries", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "abc" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "def" });
    harness.renderer.setCursor({ bodyIndex: 1, charOffset: 0 });
    harness.input.emitNavigate({
      key: "ArrowLeft",
      shift: true,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 1, charOffset: 0 },
      focus: { bodyIndex: 0, charOffset: 3 },
    });

    harness.destroy();
  });

  test("Shift+ArrowRight extends selection forward across paragraph boundaries", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "abc" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "def" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 3 });
    harness.input.emitNavigate({
      key: "ArrowRight",
      shift: true,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 3 },
      focus: { bodyIndex: 1, charOffset: 0 },
    });

    harness.destroy();
  });

  test("ArrowLeft collapses an existing selection to its start", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 1 },
      focus: { bodyIndex: 0, charOffset: 4 },
    });
    harness.input.emitNavigate({
      key: "ArrowLeft",
      shift: false,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 1 },
      focus: { bodyIndex: 0, charOffset: 1 },
    });

    harness.destroy();
  });

  test("ArrowLeft collapses a backward multi-paragraph selection to its start", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "abc" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "def" });
    harness.renderer.setSelection({
      anchor: { bodyIndex: 1, charOffset: 2 },
      focus: { bodyIndex: 0, charOffset: 1 },
    });
    harness.input.emitNavigate({
      key: "ArrowLeft",
      shift: false,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 1 },
      focus: { bodyIndex: 0, charOffset: 1 },
    });

    harness.destroy();
  });

  test("PageDown uses hit-testing to move vertically", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "beta" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 2 });
    harness.renderer.nextHit = { bodyIndex: 1, charOffset: 1 };
    harness.input.emitNavigate({
      key: "PageDown",
      shift: false,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 1, charOffset: 1 },
      focus: { bodyIndex: 1, charOffset: 1 },
    });

    harness.destroy();
  });

  test("PageUp uses hit-testing to move vertically", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "beta" });
    harness.renderer.setCursor({ bodyIndex: 1, charOffset: 2 });
    harness.renderer.nextHit = { bodyIndex: 0, charOffset: 1 };
    harness.input.emitNavigate({
      key: "PageUp",
      shift: false,
      meta: false,
      alt: false,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 1 },
      focus: { bodyIndex: 0, charOffset: 1 },
    });

    harness.destroy();
  });

  test("Ctrl/Alt+Backspace deletes the prior word", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 10 });
    harness.input.emitInput({ type: "deleteWordBackward" });

    expect(firstParaText(harness.docEdit())).toBe("alpha gamma");
    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 6 },
      focus: { bodyIndex: 0, charOffset: 6 },
    });

    harness.destroy();
  });

  test("Ctrl/Alt+Delete deletes the next word", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 6 });
    harness.input.emitInput({ type: "deleteWordForward" });

    expect(firstParaText(harness.docEdit())).toBe("alpha gamma");
    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 6 },
      focus: { bodyIndex: 0, charOffset: 6 },
    });

    harness.destroy();
  });

  test("Alt+ArrowLeft moves caret backward by word", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 16 });
    harness.input.emitNavigate({
      key: "ArrowLeft",
      shift: false,
      meta: false,
      alt: true,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 11 },
      focus: { bodyIndex: 0, charOffset: 11 },
    });

    harness.destroy();
  });

  test("Ctrl+ArrowRight moves caret forward by word", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 0 });
    harness.input.emitNavigate({
      key: "ArrowRight",
      shift: false,
      meta: false,
      alt: false,
      ctrl: true,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 6 },
      focus: { bodyIndex: 0, charOffset: 6 },
    });

    harness.destroy();
  });

  test("Alt+Shift+ArrowRight extends selection by word", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 0 });
    harness.input.emitNavigate({
      key: "ArrowRight",
      shift: true,
      meta: false,
      alt: true,
      ctrl: false,
    });

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 0, charOffset: 6 },
    });

    harness.destroy();
  });

  test("Cmd/Ctrl+A selects the full document", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha" });
    harness.input.emitInput({ type: "insertParagraph" });
    harness.input.emitInput({ type: "insertText", data: "beta" });
    harness.input.emitShortcut("a");

    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 0 },
      focus: { bodyIndex: 1, charOffset: 4 },
    });

    harness.destroy();
  });
});

describe("controller word delete semantics", () => {
  test("Alt+Backspace deletes the previous word", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 16 });
    harness.input.emitInput({ type: "deleteWordBackward" });

    expect(firstParaText(harness.docEdit())).toBe("alpha beta ");
    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 11 },
      focus: { bodyIndex: 0, charOffset: 11 },
    });

    harness.destroy();
  });

  test("Ctrl+Delete deletes the next word", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 6 });
    harness.input.emitInput({ type: "deleteWordForward" });

    expect(firstParaText(harness.docEdit())).toBe("alpha gamma");
    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 6 },
      focus: { bodyIndex: 0, charOffset: 6 },
    });

    harness.destroy();
  });

  test("word delete removes an expanded selection directly", () => {
    const harness = makeControllerHarness();

    harness.input.emitInput({ type: "insertText", data: "alpha beta gamma" });
    harness.renderer.setSelection({
      anchor: { bodyIndex: 0, charOffset: 6 },
      focus: { bodyIndex: 0, charOffset: 10 },
    });
    harness.input.emitInput({ type: "deleteWordBackward" });

    expect(firstParaText(harness.docEdit())).toBe("alpha  gamma");
    expect(harness.renderer.getSelection()).toEqual({
      anchor: { bodyIndex: 0, charOffset: 6 },
      focus: { bodyIndex: 0, charOffset: 6 },
    });

    harness.destroy();
  });
});

describe("controller pending formatting semantics", () => {
  test("collapsed selection formatting applies to the next typed text", () => {
    const harness = makeControllerHarness();

    harness.controller.setTextStyle({
      bold: true,
      fontFamily: "Noto Serif",
      fontSizePt: 18,
      color: "#3366FF",
    });
    harness.input.emitInput({ type: "insertText", data: "A" });

    const run = paraRuns(harness.docEdit(), 0)[0]!;
    expect(run.text).toBe("A");
    expect(run.bold).toBe(true);
    expect(run.fontFamily).toBe("Noto Serif");
    expect(run.fontSizePt).toBe(18);
    expect(run.color).toBe("3366FF");

    harness.destroy();
  });
});

// ---------------------------------------------------------------------------
// 27. Multi-paragraph editing
// ---------------------------------------------------------------------------

describe("multi-paragraph editing", () => {
  test("edit text in second paragraph", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Paragraph one",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 13) });
    applyAndVm(editor, {
      type: "insertText",
      data: "Paragraph two",
      anchor: pos(1, 0),
    });

    expect(paraText(editor, 0)).toBe("Paragraph one");
    expect(paraText(editor, 1)).toBe("Paragraph two");
    editor.free();
  });

  test("format text in second paragraph", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "First",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });
    applyAndVm(editor, {
      type: "insertText",
      data: "Second",
      anchor: pos(1, 0),
    });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(1, 0),
      focus: pos(1, 6),
    });

    // First paragraph should NOT be bold
    const runs0 = paraRuns(editor, 0);
    expect(runs0[0]?.bold).toBeUndefined();

    // Second paragraph should be bold
    const runs1 = paraRuns(editor, 1);
    expect(runs1[0]?.bold).toBe(true);
    expect(runs1[0]?.text).toBe("Second");
    editor.free();
  });

  test("delete in second paragraph", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "First",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });
    applyAndVm(editor, {
      type: "insertText",
      data: "Second",
      anchor: pos(1, 0),
    });

    // Delete "Sec" from second paragraph
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(1, 0),
      focus: pos(1, 3),
    });
    expect(paraText(editor, 1)).toBe("ond");
    expect(paraText(editor, 0)).toBe("First"); // untouched
    editor.free();
  });

  test("three paragraphs via double split", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "AABBCC",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 2) });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(1, 2) });

    expect(vm(editor).body.length).toBe(3);
    expect(paraText(editor, 0)).toBe("AA");
    expect(paraText(editor, 1)).toBe("BB");
    expect(paraText(editor, 2)).toBe("CC");
    editor.free();
  });

  test("save with multiple paragraphs preserves all paragraphs", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Para 1",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 6) });
    applyAndVm(editor, {
      type: "insertText",
      data: "Para 2",
      anchor: pos(1, 0),
    });

    // Save should preserve both the original and new paragraph.
    const bytes = editor.save();
    expect(bytes.length).toBeGreaterThan(0);

    const editor2 = new DocEdit(bytes);
    expect(paraText(editor2, 0)).toBe("Para 1");
    expect(paraText(editor2, 1)).toBe("Para 2");

    editor.free();
    editor2.free();
  });

  test("formattingAt works in second paragraph", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "First",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });
    applyAndVm(editor, { type: "insertText", data: "Bold", anchor: pos(1, 0) });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(1, 0),
      focus: pos(1, 4),
    });

    const fmt = editor.formattingAt(pos(1, 2));
    expect(fmt.bold).toBe(true);

    // First paragraph should not be bold
    const fmt0 = editor.formattingAt(pos(0, 3));
    expect(fmt0.bold).toBeUndefined();
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// 28. Edge cases and stress tests
// ---------------------------------------------------------------------------

describe("edge cases", () => {
  test("insert unicode text (emoji)", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hi 👋",
      anchor: pos(0, 0),
    });
    const text = firstParaText(editor);
    expect(text).toContain("Hi");
    expect(text).toContain("👋");
    editor.free();
  });

  test("insert CJK characters", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "你好世界",
      anchor: pos(0, 0),
    });
    expect(firstParaText(editor)).toBe("你好世界");
    editor.free();
  });

  test("100 rapid inserts", () => {
    const editor = DocEdit.blank();
    for (let i = 0; i < 100; i++) {
      editor.applyIntent(
        JSON.stringify({
          type: "insertText",
          data: "x",
          anchor: DocEdit.encodePosition(0, i),
        }),
      );
    }
    const text = firstParaText(editor);
    expect(text.length).toBe(100);
    expect(text).toBe("x".repeat(100));
    editor.free();
  });

  test("alternating insert and delete", () => {
    const editor = DocEdit.blank();
    for (let i = 0; i < 20; i++) {
      applyAndVm(editor, { type: "insertText", data: "AB", anchor: pos(0, 0) });
      applyAndVm(editor, {
        type: "deleteContentBackward",
        anchor: pos(0, 1),
        focus: pos(0, 1),
      });
    }
    // Each iteration: insert "AB" at 0, then backspace at 1 → "A" remains
    // But the text accumulates since each insert is at offset 0
    // The final result depends on exact offset semantics, just verify no crash
    const text = firstParaText(editor);
    expect(text.length).toBeGreaterThan(0);
    editor.free();
  });

  test("save empty document produces valid docx", () => {
    const editor = DocEdit.blank();
    const bytes = editor.save();
    expect(bytes.length).toBeGreaterThan(0);
    // Reload to verify it's valid
    const editor2 = new DocEdit(bytes);
    expect(vm(editor2).body.length).toBeGreaterThanOrEqual(1);
    editor.free();
    editor2.free();
  });

  test("double save produces valid document both times", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "Test", anchor: pos(0, 0) });

    const bytes1 = editor.save();
    expect(editor.isDirty()).toBe(false);
    expect(bytes1.length).toBeGreaterThan(0);

    // Reload first save to verify content
    const e1 = new DocEdit(bytes1);
    expect(firstParaText(e1)).toBe("Test");

    // Second save (no changes) should still produce valid output
    const bytes2 = editor.save();
    expect(bytes2.length).toBeGreaterThan(0);

    editor.free();
    e1.free();
  });
});

// ---------------------------------------------------------------------------
// CRDT sync
// ---------------------------------------------------------------------------

describe("sync", () => {
  test("encode state vector returns bytes", () => {
    const editor = DocEdit.blank();
    const sv = editor.encodeStateVector();
    expect(sv).toBeInstanceOf(Uint8Array);
    expect(sv.length).toBeGreaterThan(0);
    editor.free();
  });

  test("encode state as update returns bytes", () => {
    const editor = DocEdit.blank();
    const update = editor.encodeStateAsUpdate();
    expect(update).toBeInstanceOf(Uint8Array);
    expect(update.length).toBeGreaterThan(0);
    editor.free();
  });

  test("sync roundtrip: edit -> encode -> apply -> verify", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    // Insert text into editor1
    applyAndVm(editor1, {
      type: "insertText",
      data: "Hello sync",
      anchor: pos(0, 0),
    });
    expect(firstParaText(editor1)).toBe("Hello sync");

    // Encode full state from editor1
    const update = editor1.encodeStateAsUpdate();
    expect(update.length).toBeGreaterThan(0);

    // Apply to editor2
    editor2.applyUpdate(update);

    // editor2 should contain the text somewhere in its body
    // (CRDT merge may reorder paragraphs from different clients)
    const model2 = vm(editor2);
    const allText = model2.body
      .map((item) => (item.runs ?? []).map((r) => r.text).join(""))
      .join("|");
    expect(allText).toContain("Hello sync");

    editor1.free();
    editor2.free();
  });

  test("diff-based sync", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    // Insert text into editor1
    applyAndVm(editor1, {
      type: "insertText",
      data: "Diff sync",
      anchor: pos(0, 0),
    });

    // Get editor2's state vector
    const sv2 = editor2.encodeStateVector();

    // Compute diff from editor1 using editor2's state vector
    const diff = editor1.encodeDiff(sv2);
    expect(diff.length).toBeGreaterThan(0);

    // Apply the diff to editor2
    editor2.applyUpdate(diff);

    // editor2 should now contain the text
    const model2 = vm(editor2);
    const allText = model2.body
      .map((item) => (item.runs ?? []).map((r) => r.text).join(""))
      .join("|");
    expect(allText).toContain("Diff sync");

    editor1.free();
    editor2.free();
  });

  test("full-state sync preserves text formatting attrs", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    applyAndVm(editor1, {
      type: "insertText",
      data: "Styled sync",
      anchor: pos(0, 0),
    });
    applyAndVm(editor1, {
      type: "setTextAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 11),
      attrs: {
        bold: true,
        fontFamily: "Noto Serif",
        fontSizePt: 18,
        color: "3366FF",
      },
    } as const);

    editor2.applyUpdate(editor1.encodeStateAsUpdate());

    const paragraph = vm(editor2).body.find((item) => {
      return (
        item.type === "paragraph" &&
        item.runs?.some((run) => run.text === "Styled sync")
      );
    });
    const run = (paragraph?.runs ?? []).find((candidate) => {
      const styled = candidate as {
        text?: string;
        fontFamily?: string;
        color?: string;
      };
      return (
        styled.text === "Styled sync" ||
        styled.fontFamily === "Noto Serif" ||
        styled.color === "3366FF"
      );
    }) as
      | {
          text: string;
          bold?: boolean;
          fontFamily?: string;
          fontSizePt?: number;
          color?: string;
        }
      | undefined;
    expect(paragraph?.type).toBe("paragraph");
    expect(run?.text).toBe("Styled sync");
    expect(run?.bold).toBe(true);
    expect(run?.fontFamily).toBe("Noto Serif");
    expect(run?.fontSizePt).toBe(18);
    expect(run?.color).toBe("3366FF");

    editor1.free();
    editor2.free();
  });

  test("diff-based sync preserves paragraph heading attrs", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    applyAndVm(editor1, {
      type: "insertText",
      data: "Heading sync",
      anchor: pos(0, 0),
    });
    applyAndVm(editor1, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 12),
      attrs: { headingLevel: 2 },
    } as const);

    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    const diff = editor1.encodeDiff(editor2.encodeStateVector());
    editor2.applyUpdate(diff);

    const paragraph = vm(editor2).body.find((item) => {
      return (
        item.type === "paragraph" &&
        item.runs?.map((run) => run.text).join("") === "Heading sync"
      );
    }) as { headingLevel?: number } | undefined;
    expect(paragraph?.headingLevel).toBe(2);

    editor1.free();
    editor2.free();
  });

  test("diff-based sync preserves paragraph alignment attrs", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();
    applyAndVm(editor1, {
      type: "insertText",
      data: "Aligned sync",
      anchor: pos(0, 0),
    });
    applyAndVm(editor1, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 12),
      attrs: { alignment: "right" },
    } as const);

    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    const diff = editor1.encodeDiff(editor2.encodeStateVector());
    editor2.applyUpdate(diff);

    const paragraph = vm(editor2).body.find((item) => {
      return (
        item.type === "paragraph" &&
        item.runs?.map((run) => run.text).join("") === "Aligned sync"
      );
    }) as { alignment?: string } | undefined;
    expect(paragraph?.alignment).toBe("right");

    editor1.free();
    editor2.free();
  });

  test("diff-based sync preserves hyperlink attrs", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    applyAndVm(editor1, {
      type: "insertText",
      data: "Link sync",
      anchor: pos(0, 0),
    });
    applyAndVm(editor1, {
      type: "setTextAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 9),
      attrs: { hyperlink: "https://example.com" },
    } as const);

    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    const diff = editor1.encodeDiff(editor2.encodeStateVector());
    editor2.applyUpdate(diff);

    const paragraph = vm(editor2).body.find((item) => {
      return (
        item.type === "paragraph" &&
        item.runs?.map((run) => run.text).join("") === "Link sync"
      );
    });
    expect(paragraph?.runs?.length).toBeGreaterThan(0);
    const run = paragraph!.runs![0]!;
    expect(run.text).toBe("Link sync");
    expect(run.hyperlink).toBe("https://example.com");

    editor1.free();
    editor2.free();
  });

  test("diff-based sync preserves highlight attrs", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    applyAndVm(editor1, {
      type: "insertText",
      data: "Highlight sync",
      anchor: pos(0, 0),
    });
    applyAndVm(editor1, {
      type: "setTextAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 14),
      attrs: { highlight: "yellow" },
    });

    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    const diff = editor1.encodeDiff(editor2.encodeStateVector());
    editor2.applyUpdate(diff);

    const paragraph = vm(editor2).body.find((item) => {
      return (
        item.type === "paragraph" &&
        item.runs?.map((run) => run.text).join("") === "Highlight sync"
      );
    });
    const run = paragraph!.runs![0]!;
    expect(run.text).toBe("Highlight sync");
    expect(run.highlight).toBe("yellow");

    editor1.free();
    editor2.free();
  });

  test("diff-based sync preserves paragraph spacing and indent attrs", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    applyAndVm(editor1, {
      type: "insertText",
      data: "Block sync",
      anchor: pos(0, 0),
    });
    applyAndVm(editor1, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 10),
      attrs: {
        spacingBeforePt: 12,
        spacingAfterPt: 18,
        indentLeftPt: 36,
        indentFirstLinePt: 18,
      },
    } as const);

    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    const diff = editor1.encodeDiff(editor2.encodeStateVector());
    editor2.applyUpdate(diff);

    const paragraph = vm(editor2).body.find((item) => {
      return (
        item.type === "paragraph" &&
        item.runs?.map((run) => run.text).join("") === "Block sync"
      );
    }) as
      | {
          spacingBeforePt?: number;
          spacingAfterPt?: number;
          indents?: { leftPt?: number; firstLinePt?: number };
        }
      | undefined;
    expect(paragraph?.spacingBeforePt).toBe(12);
    expect(paragraph?.spacingAfterPt).toBe(18);
    expect(paragraph?.indents?.leftPt).toBe(36);
    expect(paragraph?.indents?.firstLinePt).toBe(18);

    editor1.free();
    editor2.free();
  });

  test("diff-based sync preserves line spacing multiple attrs", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    applyAndVm(editor1, {
      type: "insertText",
      data: "Line space sync",
      anchor: pos(0, 0),
    });
    applyAndVm(editor1, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(0, 15),
      attrs: {
        lineSpacingMultiple: 1.5,
      },
    } as const);

    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    const diff = editor1.encodeDiff(editor2.encodeStateVector());
    editor2.applyUpdate(diff);

    const paragraph = vm(editor2).body.find((item) => {
      return (
        item.type === "paragraph" &&
        item.runs?.map((run) => run.text).join("") === "Line space sync"
      );
    }) as { lineSpacing?: { value?: number; rule?: string } } | undefined;
    expect(paragraph?.lineSpacing?.value).toBe(1.5);
    expect(paragraph?.lineSpacing?.rule).toBe("auto");

    editor1.free();
    editor2.free();
  });

  test("remote diff updates preserve existing decimal list metadata", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    applyAndVm(editor1, {
      type: "insertText",
      data: "One",
      anchor: pos(0, 0),
    });
    applyAndVm(editor1, { type: "insertParagraph", anchor: pos(0, 3) });
    applyAndVm(editor1, {
      type: "insertText",
      data: "Two",
      anchor: pos(1, 0),
    });
    applyAndVm(editor1, {
      type: "setParagraphAttrs",
      anchor: pos(0, 0),
      focus: pos(1, 3),
      attrs: {
        numberingKind: "decimal",
        numberingNumId: 11,
        numberingIlvl: 0,
      },
    } as const);

    editor2.applyUpdate(editor1.encodeStateAsUpdate());

    applyAndVm(editor2, {
      type: "insertText",
      data: " remote",
      anchor: pos(1, 3),
    });

    editor1.applyUpdate(editor2.encodeDiff(editor1.encodeStateVector()));

    const editor1List = decimalListParagraphs(editor1);
    const editor2List = decimalListParagraphs(editor2);
    expect(editor1List).toHaveLength(2);
    expect(editor2List).toHaveLength(2);
    expect(editor1List[0]?.numbering?.text).toBe("1.");
    expect(editor1List[1]?.numbering?.text).toBe("2.");
    expect(editor2List[0]?.numbering?.text).toBe("1.");
    expect(editor2List[1]?.numbering?.text).toBe("2.");
    expect(
      editor1List.every((paragraph) => paragraph.numbering?.numId === 11),
    ).toBe(true);
    expect(
      editor2List.every((paragraph) => paragraph.numbering?.numId === 11),
    ).toBe(true);
    const editor1Texts = editor1List.map((paragraph) =>
      (paragraph.runs ?? []).map((run) => run.text ?? "").join(""),
    );
    const editor2Texts = editor2List.map((paragraph) =>
      (paragraph.runs ?? []).map((run) => run.text ?? "").join(""),
    );
    expect(JSON.stringify(editor1Texts)).toBe(JSON.stringify(editor2Texts));
    expect(editor1Texts.some((text) => text.includes("remote"))).toBe(true);

    editor1.free();
    editor2.free();
  });
});

// ---------------------------------------------------------------------------
// Paragraph merge
// ---------------------------------------------------------------------------

describe("paragraph merge", () => {
  test("backspace at start of paragraph merges with previous", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });
    applyAndVm(editor, {
      type: "insertText",
      data: "World",
      anchor: pos(1, 0),
    });
    expect(vm(editor).body.length).toBe(2);

    // Backspace at start of paragraph 1 -> merge into paragraph 0
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(1, 0),
      focus: pos(1, 0),
    });

    expect(vm(editor).body.length).toBe(1);
    expect(firstParaText(editor)).toBe("HelloWorld");
    editor.free();
  });

  test("delete at end of paragraph merges with next", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });
    applyAndVm(editor, {
      type: "insertText",
      data: "World",
      anchor: pos(1, 0),
    });
    expect(vm(editor).body.length).toBe(2);

    applyAndVm(editor, {
      type: "deleteContentForward",
      anchor: pos(0, 5),
      focus: pos(0, 5),
    });

    expect(vm(editor).body.length).toBe(1);
    expect(firstParaText(editor)).toBe("HelloWorld");
    editor.free();
  });

  test("merge preserves formatting from both paragraphs", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertText", data: "Bold", anchor: pos(0, 0) });
    applyAndVm(editor, {
      type: "formatBold",
      anchor: pos(0, 0),
      focus: pos(0, 4),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 4) });
    applyAndVm(editor, {
      type: "insertText",
      data: "Italic",
      anchor: pos(1, 0),
    });
    applyAndVm(editor, {
      type: "formatItalic",
      anchor: pos(1, 0),
      focus: pos(1, 6),
    });

    // Merge
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(1, 0),
      focus: pos(1, 0),
    });

    expect(vm(editor).body.length).toBe(1);
    const runs = paraRuns(editor, 0);
    // Should have bold "Bold" and italic "Italic"
    const boldRun = runs.find((r) => r.bold === true);
    const italicRun = runs.find((r) => r.italic === true);
    expect(boldRun).toBeDefined();
    expect(boldRun!.text).toBe("Bold");
    expect(italicRun).toBeDefined();
    expect(italicRun!.text).toBe("Italic");
    editor.free();
  });

  test("backspace at start of first paragraph is still a no-op", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(0, 0),
      focus: pos(0, 0),
    });
    expect(firstParaText(editor)).toBe("Hello");
    editor.free();
  });

  test("merge empty paragraph into previous", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertParagraph", anchor: pos(0, 5) });
    // Paragraph 1 is empty, backspace merges (just deletes it)
    applyAndVm(editor, {
      type: "deleteContentBackward",
      anchor: pos(1, 0),
      focus: pos(1, 0),
    });
    expect(vm(editor).body.length).toBe(1);
    expect(firstParaText(editor)).toBe("Hello");
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// Tab insert
// ---------------------------------------------------------------------------

describe("tab insert", () => {
  test("inserts a tab token", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 0),
    });
    applyAndVm(editor, { type: "insertTab", anchor: pos(0, 5) });
    const runs = paraRuns(editor, 0);
    const hasTab = runs.some((r) => r.hasTab === true);
    expect(hasTab).toBe(true);
    editor.free();
  });

  test("tab at start of text", () => {
    const editor = DocEdit.blank();
    applyAndVm(editor, { type: "insertTab", anchor: pos(0, 0) });
    applyAndVm(editor, {
      type: "insertText",
      data: "Hello",
      anchor: pos(0, 1),
    });
    const runs = paraRuns(editor, 0);
    expect(runs.some((r) => r.hasTab)).toBe(true);
    expect(runs.map((r) => r.text).join("")).toContain("Hello");
    editor.free();
  });
});

// ---------------------------------------------------------------------------
// Sync provider integration (CRDT state exchange)
// ---------------------------------------------------------------------------

describe("sync provider integration", () => {
  test("local undo survives a real remote edit", () => {
    const editor1 = DocEdit.blank();
    const editor2 = DocEdit.blank();

    editor1.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "Hello",
        anchor: pos(0, 0),
      }),
    );
    editor2.applyUpdate(editor1.encodeStateAsUpdate());

    editor2.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: " local",
        anchor: pos(0, 5),
      }),
    );

    editor1.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: " remote",
        anchor: pos(0, 5),
      }),
    );
    editor2.applyUpdate(editor1.encodeDiff(editor2.encodeStateVector()));

    editor2.applyIntent(JSON.stringify({ type: "historyUndo" }));

    const text2 = vm(editor2)
      .body.flatMap((item) => (item.runs ?? []).map((r) => r.text))
      .join("");
    expect(text2).toContain("remote");
    expect(text2).not.toContain(" local");

    editor1.free();
    editor2.free();
  });

  test("two editors sync via CRDT updates", () => {
    // Editor1 creates a blank doc and types "Hello".
    const editor1 = DocEdit.blank();
    editor1.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "Hello",
        anchor: pos(0, 0),
      }),
    );

    // Editor2 bootstraps from editor1's full state, so both share
    // the same CRDT history (including the "Hello" insert).
    const editor2 = DocEdit.blank();
    const fullState = editor1.encodeStateAsUpdate();
    editor2.applyUpdate(fullState);

    // Verify editor2 sees "Hello" in its view model.
    const allText2 = vm(editor2)
      .body.flatMap((item) => (item.runs ?? []).map((r) => r.text))
      .join("");
    expect(allText2).toContain("Hello");

    // Editor2 types "XYZ" at position 0 of the first paragraph.
    editor2.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "XYZ",
        anchor: pos(0, 0),
      }),
    );

    // Sync back: editor2 -> editor1 using diff-based approach.
    const sv1 = editor1.encodeStateVector();
    const diff = editor2.encodeDiff(sv1);
    editor1.applyUpdate(diff);

    // Both editors should contain both "Hello" and "XYZ".
    const text1 = vm(editor1)
      .body.flatMap((item) => (item.runs ?? []).map((r) => r.text))
      .join("");
    expect(text1).toContain("Hello");
    expect(text1).toContain("XYZ");

    const text2 = vm(editor2)
      .body.flatMap((item) => (item.runs ?? []).map((r) => r.text))
      .join("");
    expect(text2).toContain("Hello");
    expect(text2).toContain("XYZ");

    // Both editors should converge to the exact same text.
    expect(text1).toBe(text2);

    editor1.free();
    editor2.free();
  });

  test("encodeStateVector returns non-empty bytes", () => {
    const editor = DocEdit.blank();
    const sv = editor.encodeStateVector();
    expect(sv).toBeInstanceOf(Uint8Array);
    expect(sv.length).toBeGreaterThan(0);
    editor.free();
  });

  test("encodeDiff with empty state vector returns full state", () => {
    const editor = DocEdit.blank();
    editor.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "test",
        anchor: pos(0, 0),
      }),
    );

    // Empty state vector = give me everything
    const emptySv = new Uint8Array([0]);
    const diff = editor.encodeDiff(emptySv);
    expect(diff).toBeInstanceOf(Uint8Array);
    expect(diff.length).toBeGreaterThan(0);

    // Apply this diff to a fresh editor and verify content
    const editor2 = DocEdit.blank();
    editor2.applyUpdate(diff);
    const text = vm(editor2)
      .body.map((item) => (item.runs ?? []).map((r) => r.text).join(""))
      .join("|");
    expect(text).toContain("test");

    editor.free();
    editor2.free();
  });

  test("two editors converge after formatting and heading updates with anti-entropy resync", () => {
    const editor1 = DocEdit.blank();
    editor1.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "Shared title",
        anchor: pos(0, 0),
      }),
    );

    const editor2 = DocEdit.blank();
    editor2.applyUpdate(editor1.encodeStateAsUpdate());

    editor1.applyIntent(
      JSON.stringify({
        type: "setTextAttrs",
        anchor: pos(0, 0),
        focus: pos(0, 6),
        attrs: {
          fontFamily: "Noto Serif",
          fontSizePt: 16,
          color: "AA5500",
        },
      }),
    );

    editor2.applyIntent(
      JSON.stringify({
        type: "setParagraphAttrs",
        anchor: pos(0, 0),
        focus: pos(0, 12),
        attrs: { headingLevel: 1 },
      }),
    );

    editor2.applyUpdate(editor1.encodeDiff(editor2.encodeStateVector()));
    editor1.applyUpdate(editor2.encodeDiff(editor1.encodeStateVector()));
    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    editor1.applyUpdate(editor2.encodeStateAsUpdate());

    const model1 = vm(editor1);
    const model2 = vm(editor2);

    const para1 = model1.body.find((item) => {
      const paragraph = item as { type?: string; headingLevel?: number };
      return paragraph.type === "paragraph" && paragraph.headingLevel === 1;
    }) as { headingLevel?: number; runs?: unknown[] } | undefined;
    const para2 = model2.body.find((item) => {
      const paragraph = item as { type?: string; headingLevel?: number };
      return paragraph.type === "paragraph" && paragraph.headingLevel === 1;
    }) as { headingLevel?: number; runs?: unknown[] } | undefined;
    const formattedPara1 = model1.body.find((item) => {
      return (
        item.type === "paragraph" &&
        item.runs?.some((run) => run.fontFamily === "Noto Serif")
      );
    }) as { runs?: unknown[] } | undefined;
    const formattedPara2 = model2.body.find((item) => {
      return (
        item.type === "paragraph" &&
        item.runs?.some((run) => run.fontFamily === "Noto Serif")
      );
    }) as { runs?: unknown[] } | undefined;
    const run1 = (formattedPara1?.runs?.find((run) => {
      const candidate = run as { fontFamily?: string };
      return candidate.fontFamily === "Noto Serif";
    }) ?? null) as {
      fontFamily?: string;
      fontSizePt?: number;
      color?: string;
    } | null;
    const run2 = (formattedPara2?.runs?.find((run) => {
      const candidate = run as { fontFamily?: string };
      return candidate.fontFamily === "Noto Serif";
    }) ?? null) as {
      fontFamily?: string;
      fontSizePt?: number;
      color?: string;
    } | null;

    expect(para1?.headingLevel).toBe(1);
    expect(para2?.headingLevel).toBe(1);
    expect(run1?.fontFamily).toBe("Noto Serif");
    expect(run2?.fontFamily).toBe("Noto Serif");
    expect(run1?.fontSizePt).toBe(16);
    expect(run2?.fontSizePt).toBe(16);
    expect(run1?.color).toBe("AA5500");
    expect(run2?.color).toBe("AA5500");

    editor1.free();
    editor2.free();
  });

  test("existing decimal list paragraphs converge after remote edits and full-state resync", () => {
    const editor1 = DocEdit.blank();
    editor1.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "Alpha",
        anchor: pos(0, 0),
      }),
    );
    editor1.applyIntent(
      JSON.stringify({
        type: "insertParagraph",
        anchor: pos(0, 5),
      }),
    );
    editor1.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "Beta",
        anchor: pos(1, 0),
      }),
    );
    editor1.applyIntent(
      JSON.stringify({
        type: "insertParagraph",
        anchor: pos(1, 4),
      }),
    );
    editor1.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: "Gamma",
        anchor: pos(2, 0),
      }),
    );
    editor1.applyIntent(
      JSON.stringify({
        type: "setParagraphAttrs",
        anchor: pos(0, 0),
        focus: pos(2, 5),
        attrs: {
          numberingKind: "decimal",
          numberingNumId: 21,
          numberingIlvl: 0,
        },
      }),
    );

    const editor2 = DocEdit.blank();
    editor2.applyUpdate(editor1.encodeStateAsUpdate());

    editor1.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: " local",
        anchor: pos(0, 5),
      }),
    );
    editor2.applyIntent(
      JSON.stringify({
        type: "insertText",
        data: " remote",
        anchor: pos(2, 5),
      }),
    );

    editor2.applyUpdate(editor1.encodeDiff(editor2.encodeStateVector()));
    editor1.applyUpdate(editor2.encodeDiff(editor1.encodeStateVector()));
    editor2.applyUpdate(editor1.encodeStateAsUpdate());
    editor1.applyUpdate(editor2.encodeStateAsUpdate());

    const editor1List = decimalListParagraphs(editor1);
    const editor2List = decimalListParagraphs(editor2);
    expect(editor1List).toHaveLength(3);
    expect(editor2List).toHaveLength(3);
    expect(editor1List[0]?.numbering?.text).toBe("1.");
    expect(editor1List[1]?.numbering?.text).toBe("2.");
    expect(editor1List[2]?.numbering?.text).toBe("3.");
    expect(editor2List[0]?.numbering?.text).toBe("1.");
    expect(editor2List[1]?.numbering?.text).toBe("2.");
    expect(editor2List[2]?.numbering?.text).toBe("3.");
    expect(
      editor1List.every((paragraph) => paragraph.numbering?.numId === 21),
    ).toBe(true);
    expect(
      editor2List.every((paragraph) => paragraph.numbering?.numId === 21),
    ).toBe(true);

    const model1 = JSON.stringify(vm(editor1).body);
    const model2 = JSON.stringify(vm(editor2).body);
    expect(model1).toBe(model2);
    const texts = editor1List.map((paragraph) =>
      (paragraph.runs ?? []).map((run) => run.text ?? "").join(""),
    );
    expect(texts.some((text) => text.includes("local"))).toBe(true);
    expect(texts.some((text) => text.includes("remote"))).toBe(true);

    editor1.free();
    editor2.free();
  });

  test("controller-authored decimal list continuation syncs to a peer", () => {
    const harness = makeControllerHarness();
    const peer = DocEdit.blank();
    peer.applyUpdate(harness.docEdit().encodeStateAsUpdate());

    harness.controller.toggleList("decimal");
    syncEditors(harness.docEdit(), peer);

    harness.input.emitInput({ type: "insertText", data: "One" });
    syncEditors(harness.docEdit(), peer);

    harness.input.emitInput({ type: "insertParagraph" });
    syncEditors(harness.docEdit(), peer);

    harness.input.emitInput({ type: "insertText", data: "Two" });
    syncEditors(harness.docEdit(), peer);
    fullResyncEditors(harness.docEdit(), peer);

    const peerList = decimalListParagraphs(peer);
    expect(peerList).toHaveLength(2);
    expectDecimalListParagraph(
      peerList[0],
      "One",
      1,
      peerList[0]?.numbering?.numId ?? -1,
    );
    expectDecimalListParagraph(
      peerList[1],
      "Two",
      2,
      peerList[0]?.numbering?.numId ?? -1,
    );
    expect(JSON.stringify(vm(harness.docEdit()).body)).toBe(
      JSON.stringify(vm(peer).body),
    );

    harness.destroy();
    peer.free();
  });

  test("controller-authored list indent and outdent sync to a peer", () => {
    const harness = makeControllerHarness();
    const peer = DocEdit.blank();
    peer.applyUpdate(harness.docEdit().encodeStateAsUpdate());

    harness.controller.toggleList("bullet");
    harness.input.emitInput({ type: "insertText", data: "Nested" });
    syncEditors(harness.docEdit(), peer);

    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 6 });
    harness.input.emitSelectionChange();
    harness.input.emitInput({ type: "insertTab" });
    syncEditors(harness.docEdit(), peer);

    let peerPara = numberedParagraphs(peer)[0];
    expect(peerPara?.numbering?.format).toBe("bullet");
    expect(peerPara?.numbering?.level).toBe(1);

    harness.renderer.setCursor({ bodyIndex: 0, charOffset: 6 });
    harness.input.emitSelectionChange();
    harness.input.emitInput({ type: "insertTab", shift: true });
    syncEditors(harness.docEdit(), peer);
    fullResyncEditors(harness.docEdit(), peer);

    peerPara = numberedParagraphs(peer)[0];
    expect(peerPara?.numbering?.format).toBe("bullet");
    expect(peerPara?.numbering?.level).toBe(0);
    expect(JSON.stringify(numberedParagraphs(harness.docEdit()))).toBe(
      JSON.stringify(numberedParagraphs(peer)),
    );

    harness.destroy();
    peer.free();
  });

  test("controller-authored empty list exit and toggle-off clear numbering on a peer", () => {
    const exitHarness = makeControllerHarness();
    const exitPeer = DocEdit.blank();
    exitPeer.applyUpdate(exitHarness.docEdit().encodeStateAsUpdate());

    exitHarness.controller.toggleList("bullet");
    syncEditors(exitHarness.docEdit(), exitPeer);
    exitHarness.input.emitInput({ type: "insertParagraph" });
    syncEditors(exitHarness.docEdit(), exitPeer);
    fullResyncEditors(exitHarness.docEdit(), exitPeer);

    expect(paraModel(exitPeer, 0)?.numbering).toBeUndefined();
    expect(JSON.stringify(vm(exitHarness.docEdit()).body)).toBe(
      JSON.stringify(vm(exitPeer).body),
    );

    exitHarness.destroy();
    exitPeer.free();

    const toggleHarness = makeControllerHarness();
    const togglePeer = DocEdit.blank();
    togglePeer.applyUpdate(toggleHarness.docEdit().encodeStateAsUpdate());

    toggleHarness.controller.toggleList("bullet");
    syncEditors(toggleHarness.docEdit(), togglePeer);
    toggleHarness.controller.toggleList("bullet");
    syncEditors(toggleHarness.docEdit(), togglePeer);
    fullResyncEditors(toggleHarness.docEdit(), togglePeer);

    expect(paraModel(togglePeer, 0)?.numbering).toBeUndefined();
    expect(JSON.stringify(vm(toggleHarness.docEdit()).body)).toBe(
      JSON.stringify(vm(togglePeer).body),
    );

    toggleHarness.destroy();
    togglePeer.free();
  });
});
