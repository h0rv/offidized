import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { HtmlInput } from "../../js/html/html-input.ts";

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
    if (!type) throw new Error("event.type is required");
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
  classList = {
    contains: (_name: string) => false,
  };
  textContent = "";
  isContentEditable = true;
}

class FakeDocument extends FakeEventTarget {
  createRange(): Range {
    throw new Error("not implemented");
  }
}

const originalGlobals = {
  document: globalThis.document,
  window: globalThis.window,
  requestAnimationFrame: globalThis.requestAnimationFrame,
};

beforeEach(() => {
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
    value: (cb: FrameRequestCallback) => {
      cb(0);
      return 1;
    },
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
});

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
  },
) {
  const event = {
    type,
    key: init.key,
    isComposing: init.isComposing ?? false,
    keyCode: init.keyCode,
    metaKey: false,
    ctrlKey: false,
    altKey: false,
    shiftKey: false,
  };
  target.dispatchEvent(event);
  return event;
}

function createSubject() {
  const surface = new FakeElement();
  const input = new HtmlInput(surface as unknown as HTMLElement);
  const emitted: Array<{ type: string; data?: string }> = [];
  const navigations: Array<{ key: string; shift: boolean }> = [];

  input.onInput((event) => {
    emitted.push(event);
  });
  input.onNavigate((event) => {
    navigations.push({ key: event.key, shift: event.shift });
  });

  return { surface, input, emitted, navigations };
}

describe("HtmlInput IME handling", () => {
  test("commits composition text exactly once when insertFromComposition fires", () => {
    const { surface, input, emitted } = createSubject();

    dispatchComposition(surface, "compositionstart");
    surface.textContent = "に";
    dispatchComposition(surface, "compositionupdate", "に");
    dispatchBeforeInput(surface, {
      inputType: "insertCompositionText",
      data: "に",
      isComposing: true,
    });
    surface.textContent = "日";
    dispatchBeforeInput(surface, {
      inputType: "insertFromComposition",
      data: "日",
    });
    dispatchComposition(surface, "compositionend", "日");

    expect(emitted).toEqual([{ type: "insertFromComposition", data: "日" }]);

    input.destroy();
  });

  test("falls back to compositionend commit when insertFromComposition is absent", () => {
    const { surface, input, emitted } = createSubject();

    dispatchComposition(surface, "compositionstart");
    surface.textContent = "漢";
    dispatchComposition(surface, "compositionupdate", "漢");
    dispatchBeforeInput(surface, {
      inputType: "insertCompositionText",
      data: "漢",
      isComposing: true,
    });
    dispatchComposition(surface, "compositionend", "漢");

    expect(emitted).toEqual([{ type: "insertFromComposition", data: "漢" }]);

    input.destroy();
  });

  test("ignores dead-key and process key events", () => {
    const { surface, input, emitted, navigations } = createSubject();

    dispatchKeyboard(surface, "keydown", { key: "Dead" });
    dispatchKeyboard(surface, "keydown", { key: "Process", keyCode: 229 });
    dispatchKeyboard(surface, "keydown", {
      key: "ArrowLeft",
      isComposing: true,
    });
    dispatchKeyboard(surface, "keyup", {
      key: "ArrowLeft",
      isComposing: true,
    });

    expect(emitted).toEqual([]);
    expect(navigations).toEqual([]);

    input.destroy();
  });
});
