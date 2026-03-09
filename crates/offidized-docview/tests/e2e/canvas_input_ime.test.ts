import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { CanvasInput } from "../../js/canvas/canvas-input.ts";

type Listener = (event: any) => void;

class FakeEventTarget {
  private listeners = new Map<string, Set<Listener>>();

  addEventListener(type: string, listener: Listener): void {
    const bucket = this.listeners.get(type) ?? new Set<Listener>();
    bucket.add(listener);
    this.listeners.set(type, bucket);
  }

  removeEventListener(type: string, listener: Listener): void {
    this.listeners.get(type)?.delete(listener);
  }

  dispatchEvent(event: Record<string, unknown>): boolean {
    const type = String(event.type ?? "");
    if (!type) {
      throw new Error("FakeEventTarget.dispatchEvent requires an event.type");
    }
    const listeners = [...(this.listeners.get(type) ?? [])];
    const dispatched = event as Record<string, unknown> & {
      defaultPrevented?: boolean;
      preventDefault?: () => void;
      target?: unknown;
      currentTarget?: unknown;
    };
    dispatched.defaultPrevented = false;
    dispatched.preventDefault ??= () => {
      dispatched.defaultPrevented = true;
    };
    dispatched.target = this;
    dispatched.currentTarget = this;
    for (const listener of listeners) {
      listener(dispatched);
    }
    return !dispatched.defaultPrevented;
  }
}

class FakeElement extends FakeEventTarget {
  style: Record<string, string> = {};
  dataset: Record<string, string> = {};
  parentNode: FakeElement | null = null;
  children: FakeElement[] = [];
  value = "";
  tabIndex = -1;
  focused = false;
  attributes = new Map<string, string>();
  scrollLeft = 0;
  scrollTop = 0;

  constructor(readonly tagName: string) {
    super();
  }

  appendChild(child: FakeElement): FakeElement {
    child.parentNode = this;
    this.children.push(child);
    return child;
  }

  removeChild(child: FakeElement): FakeElement {
    const index = this.children.indexOf(child);
    if (index >= 0) {
      this.children.splice(index, 1);
      child.parentNode = null;
    }
    return child;
  }

  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }

  focus(): void {
    this.focused = true;
  }
}

class FakeDocument {
  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName);
  }
}

class RafController {
  private nextId = 1;
  private callbacks = new Map<number, FrameRequestCallback>();

  request = (callback: FrameRequestCallback): number => {
    const id = this.nextId++;
    this.callbacks.set(id, callback);
    return id;
  };

  cancel = (id: number): void => {
    this.callbacks.delete(id);
  };

  flush(): void {
    const entries = [...this.callbacks.entries()];
    this.callbacks.clear();
    for (const [, callback] of entries) {
      callback(0);
    }
  }
}

const originalGlobals = {
  document: globalThis.document,
  window: globalThis.window,
  requestAnimationFrame: globalThis.requestAnimationFrame,
  cancelAnimationFrame: globalThis.cancelAnimationFrame,
};

let raf: RafController;

beforeEach(() => {
  raf = new RafController();
  Object.defineProperty(globalThis, "document", {
    value: new FakeDocument(),
    configurable: true,
    writable: true,
  });
  Object.defineProperty(globalThis, "window", {
    value: new FakeEventTarget(),
    configurable: true,
    writable: true,
  });
  Object.defineProperty(globalThis, "requestAnimationFrame", {
    value: raf.request,
    configurable: true,
    writable: true,
  });
  Object.defineProperty(globalThis, "cancelAnimationFrame", {
    value: raf.cancel,
    configurable: true,
    writable: true,
  });
});

afterEach(() => {
  Object.defineProperty(globalThis, "document", {
    value: originalGlobals.document,
    configurable: true,
    writable: true,
  });
  Object.defineProperty(globalThis, "window", {
    value: originalGlobals.window,
    configurable: true,
    writable: true,
  });
  Object.defineProperty(globalThis, "requestAnimationFrame", {
    value: originalGlobals.requestAnimationFrame,
    configurable: true,
    writable: true,
  });
  Object.defineProperty(globalThis, "cancelAnimationFrame", {
    value: originalGlobals.cancelAnimationFrame,
    configurable: true,
    writable: true,
  });
});

function createSubject() {
  const container = new FakeElement("div");
  const input = new CanvasInput(container as unknown as HTMLElement);
  const textarea = input.getInputElement() as unknown as FakeElement;
  const emitted: Array<{ type: string; data?: string }> = [];
  const navigations: Array<{ key: string; shift: boolean }> = [];
  const shortcuts: Array<{ key: string; shift: boolean }> = [];
  let selectionChanges = 0;

  input.onInput((event) => {
    emitted.push(event);
  });
  input.onNavigate((event) => {
    navigations.push({ key: event.key, shift: event.shift });
  });
  input.onShortcut((key, shift) => {
    shortcuts.push({ key, shift });
  });
  input.onSelectionChange(() => {
    selectionChanges += 1;
  });

  return {
    container,
    input,
    textarea,
    emitted,
    navigations,
    shortcuts,
    getSelectionChanges: () => selectionChanges,
  };
}

function appendCaretOverlay(
  container: FakeElement,
  {
    left,
    top,
    height,
    display = "block",
  }: {
    left: number;
    top: number;
    height: number;
    display?: string;
  },
): FakeElement {
  const caret = new FakeElement("div");
  caret.dataset.docviewCaretOverlay = "1";
  caret.style.left = `${left}px`;
  caret.style.top = `${top}px`;
  caret.style.height = `${height}px`;
  caret.style.display = display;
  container.appendChild(caret);
  return caret;
}

function dispatchBeforeInput(
  target: FakeElement,
  init: {
    inputType: string;
    data?: string | null;
    isComposing?: boolean;
    dataTransfer?: { getData: (kind: string) => string };
  },
) {
  const event = {
    type: "beforeinput",
    inputType: init.inputType,
    data: init.data ?? null,
    isComposing: init.isComposing ?? false,
    dataTransfer: init.dataTransfer,
  };
  target.dispatchEvent(event);
  return event;
}

function dispatchComposition(
  target: FakeElement,
  type: "compositionstart" | "compositionupdate" | "compositionend",
  data?: string,
) {
  const event = {
    type,
    data: data ?? "",
  };
  target.dispatchEvent(event);
  return event;
}

function dispatchKeyboard(
  target: FakeElement,
  type: "keydown" | "keyup",
  init: {
    key: string;
    isComposing?: boolean;
    keyCode?: number;
    shiftKey?: boolean;
    metaKey?: boolean;
    ctrlKey?: boolean;
    altKey?: boolean;
  },
) {
  const event = {
    type,
    key: init.key,
    isComposing: init.isComposing ?? false,
    keyCode: init.keyCode,
    shiftKey: init.shiftKey ?? false,
    metaKey: init.metaKey ?? false,
    ctrlKey: init.ctrlKey ?? false,
    altKey: init.altKey ?? false,
  };
  target.dispatchEvent(event);
  return event;
}

describe("CanvasInput IME handling", () => {
  test("commits composition text exactly once when insertFromComposition fires", () => {
    const { emitted, input, textarea } = createSubject();

    dispatchComposition(textarea, "compositionstart");
    textarea.value = "に";
    dispatchComposition(textarea, "compositionupdate", "に");
    dispatchBeforeInput(textarea, {
      inputType: "insertCompositionText",
      data: "に",
      isComposing: true,
    });
    textarea.value = "日";
    dispatchBeforeInput(textarea, {
      inputType: "insertFromComposition",
      data: "日",
    });
    dispatchComposition(textarea, "compositionend", "日");
    raf.flush();

    expect(emitted).toEqual([{ type: "insertFromComposition", data: "日" }]);
    expect(textarea.value).toBe("");

    input.destroy();
  });

  test("falls back to compositionend commit when insertFromComposition is absent", () => {
    const { emitted, input, textarea } = createSubject();

    dispatchComposition(textarea, "compositionstart");
    textarea.value = "漢";
    dispatchComposition(textarea, "compositionupdate", "漢");
    dispatchBeforeInput(textarea, {
      inputType: "insertCompositionText",
      data: "漢",
      isComposing: true,
    });
    dispatchComposition(textarea, "compositionend", "漢");
    raf.flush();

    expect(emitted).toEqual([{ type: "insertFromComposition", data: "漢" }]);
    expect(textarea.value).toBe("");

    input.destroy();
  });

  test("uses textarea value for committed composition when browsers omit commit data", () => {
    const { emitted, input, textarea } = createSubject();

    dispatchComposition(textarea, "compositionstart");
    textarea.value = "é";
    dispatchComposition(textarea, "compositionupdate", "é");
    dispatchBeforeInput(textarea, {
      inputType: "insertCompositionText",
      data: "é",
      isComposing: true,
    });
    dispatchBeforeInput(textarea, {
      inputType: "insertFromComposition",
      data: "",
    });
    dispatchComposition(textarea, "compositionend", "");
    raf.flush();

    expect(emitted).toEqual([{ type: "insertFromComposition", data: "é" }]);
    expect(textarea.value).toBe("");

    input.destroy();
  });

  test("canceled composition emits nothing and clears the textarea", () => {
    const { emitted, input, textarea } = createSubject();

    dispatchComposition(textarea, "compositionstart");
    textarea.value = "あ";
    dispatchComposition(textarea, "compositionupdate", "あ");
    dispatchBeforeInput(textarea, {
      inputType: "insertCompositionText",
      data: "あ",
      isComposing: true,
    });
    dispatchBeforeInput(textarea, {
      inputType: "deleteCompositionText",
      data: null,
      isComposing: true,
    });
    dispatchComposition(textarea, "compositionend", "");
    raf.flush();

    expect(emitted).toEqual([]);
    expect(textarea.value).toBe("");

    input.destroy();
  });

  test("ignores dead-key and process key events", () => {
    const {
      emitted,
      input,
      textarea,
      navigations,
      shortcuts,
      getSelectionChanges,
    } = createSubject();

    dispatchKeyboard(textarea, "keydown", { key: "Dead" });
    dispatchKeyboard(textarea, "keydown", { key: "Process", keyCode: 229 });
    dispatchKeyboard(textarea, "keydown", {
      key: "ArrowLeft",
      isComposing: true,
    });
    dispatchKeyboard(textarea, "keyup", {
      key: "ArrowLeft",
      isComposing: true,
    });

    expect(emitted).toEqual([]);
    expect(navigations).toEqual([]);
    expect(shortcuts).toEqual([]);
    expect(getSelectionChanges()).toBe(0);

    input.destroy();
  });

  test("preserves textarea content during composition and clears it after commit", () => {
    const { emitted, input, textarea } = createSubject();

    textarea.value = "x";
    dispatchBeforeInput(textarea, { inputType: "insertText", data: "x" });
    expect(emitted).toEqual([{ type: "insertText", data: "x" }]);
    expect(textarea.value).toBe("x");
    raf.flush();
    expect(textarea.value).toBe("");

    dispatchComposition(textarea, "compositionstart");
    textarea.value = "中";
    dispatchBeforeInput(textarea, {
      inputType: "insertCompositionText",
      data: "中",
      isComposing: true,
    });
    raf.flush();
    expect(textarea.value).toBe("中");

    dispatchBeforeInput(textarea, {
      inputType: "insertFromComposition",
      data: "中",
    });
    dispatchComposition(textarea, "compositionend", "中");
    raf.flush();

    expect(emitted).toEqual([
      { type: "insertText", data: "x" },
      { type: "insertFromComposition", data: "中" },
    ]);
    expect(textarea.value).toBe("");

    input.destroy();
  });

  test("anchors the hidden textarea to the visual caret overlay", () => {
    const { container, input, textarea } = createSubject();
    appendCaretOverlay(container, { left: 128, top: 244, height: 27 });

    container.dispatchEvent({
      type: "mousedown",
      button: 0,
      clientX: 128,
      clientY: 244,
    });
    raf.flush();

    expect(textarea.style.left).toBe("128px");
    expect(textarea.style.top).toBe("244px");
    expect(textarea.style.height).toBe("27px");
    expect(textarea.style.lineHeight).toBe("27px");
    expect(textarea.style.fontSize).toBe("27px");

    input.destroy();
  });

  test("keeps the last caret anchor while the overlay is hidden", () => {
    const { container, input, textarea } = createSubject();
    const caret = appendCaretOverlay(container, {
      left: 96,
      top: 188,
      height: 22,
    });

    container.dispatchEvent({
      type: "mousedown",
      button: 0,
      clientX: 96,
      clientY: 188,
    });
    raf.flush();

    caret.style.display = "none";
    dispatchComposition(textarea, "compositionstart");
    raf.flush();

    expect(textarea.style.left).toBe("96px");
    expect(textarea.style.top).toBe("188px");
    expect(textarea.style.height).toBe("22px");

    input.destroy();
  });
});
