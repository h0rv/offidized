import { init } from "../js/doc-edit.ts";
import { CanvasRenderer } from "../js/canvas/canvas-renderer.ts";
import { CanvasInput } from "../js/canvas/canvas-input.ts";
import { createEditorController } from "../js/editor.ts";
import type { DocEditorController } from "../js/adapter.ts";

type Sample = {
  t: number;
  hasCursorPos: boolean;
  cursorVisible: boolean;
  hasCursorRect: boolean;
  cursorTop: number | null;
  overlayVisible: boolean;
  focused: boolean;
  pendingDrawPages: number;
  lastActivityAgeMs: number;
  activeTag: string;
};

type DrawSample = {
  t: number;
  pageIndex: number;
  durationMs: number;
};

type SummaryStats = {
  sampleCount: number;
  drawCallCount: number;
  focusedWithCursorCount: number;
  activeWindowAnomalyCount: number;
  cursorYSpanPx: number;
  finalBodyIndex: number;
  finalCharOffset: number;
  maxPendingDrawPages: number;
  slowDrawsOver16Ms: number;
  slowDrawsOver33Ms: number;
  maxDrawMs: number;
};

const statusEl = document.getElementById("status") as HTMLSpanElement;
const summaryEl = document.getElementById("summary") as HTMLPreElement;
const logEl = document.getElementById("log") as HTMLPreElement;
const container = document.getElementById("editor-container") as HTMLDivElement;

const btnLoad = document.getElementById("btn-load") as HTMLButtonElement;
const btnClick = document.getElementById("btn-click") as HTMLButtonElement;
const btnType = document.getElementById("btn-type") as HTMLButtonElement;
const btnScenario = document.getElementById(
  "btn-scenario",
) as HTMLButtonElement;
const btnExport = document.getElementById("btn-export") as HTMLButtonElement;
const btnClear = document.getElementById("btn-clear") as HTMLButtonElement;

let renderer: CanvasRenderer | null = null;
let input: CanvasInput | null = null;
let controller: DocEditorController | null = null;
let samples: Sample[] = [];
let drawSamples: DrawSample[] = [];
let running = false;
const searchParams = new URLSearchParams(window.location.search);
const autoRun = searchParams.get("autorun") === "1";

function publishAutoStatus(stats: SummaryStats | null): void {
  const body = document.body;
  body.dataset.caretHarnessDone = "1";
  body.dataset.caretHarnessStatus = statusEl.textContent ?? "";
  if (!stats) {
    body.dataset.caretHarnessActiveWindowAnomalies = "-1";
    body.dataset.caretHarnessFocusedCursorSamples = "-1";
    return;
  }
  body.dataset.caretHarnessActiveWindowAnomalies = String(
    stats.activeWindowAnomalyCount,
  );
  body.dataset.caretHarnessFocusedCursorSamples = String(
    stats.focusedWithCursorCount,
  );
  body.dataset.caretHarnessDrawCalls = String(stats.drawCallCount);
  body.dataset.caretHarnessCursorYSpanPx = String(stats.cursorYSpanPx);
  body.dataset.caretHarnessFinalBodyIndex = String(stats.finalBodyIndex);
  body.dataset.caretHarnessFinalCharOffset = String(stats.finalCharOffset);
}

function setStatus(text: string): void {
  statusEl.textContent = text;
}

function appendLog(line: string): void {
  const ts = new Date().toISOString().slice(11, 23);
  logEl.textContent += `[${ts}] ${line}\n`;
  logEl.scrollTop = logEl.scrollHeight;
}

function clearLogs(): void {
  samples = [];
  drawSamples = [];
  logEl.textContent = "";
  summaryEl.textContent = "summary cleared";
}

function patchDrawTelemetry(target: CanvasRenderer): void {
  const anyRenderer = target as unknown as {
    __caretHarnessPatched?: boolean;
    drawPage?: (pageIndex: number) => void;
  };
  if (anyRenderer.__caretHarnessPatched) return;
  if (typeof anyRenderer.drawPage !== "function") return;

  const original = anyRenderer.drawPage.bind(target);
  anyRenderer.drawPage = (pageIndex: number): void => {
    const t0 = performance.now();
    original(pageIndex);
    drawSamples.push({
      t: performance.now(),
      pageIndex,
      durationMs: performance.now() - t0,
    });
  };
  anyRenderer.__caretHarnessPatched = true;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function raf(): Promise<void> {
  return new Promise((resolve) => {
    let done = false;
    const id = requestAnimationFrame(() => {
      done = true;
      resolve();
    });
    setTimeout(() => {
      if (done) return;
      cancelAnimationFrame(id);
      resolve();
    }, 32);
  });
}

function dispatchBeforeInput(
  textarea: HTMLTextAreaElement,
  inputType: string,
  data: string | null,
): boolean {
  let event: InputEvent | Event;
  try {
    event = new InputEvent("beforeinput", {
      bubbles: true,
      cancelable: true,
      inputType,
      data: data ?? null,
    });
  } catch {
    event = new Event("beforeinput", {
      bubbles: true,
      cancelable: true,
    });
    (event as unknown as { inputType: string }).inputType = inputType;
    (event as unknown as { data: string | null }).data = data ?? null;
  }
  return textarea.dispatchEvent(event);
}

function activeElementTag(): string {
  const active = document.activeElement;
  if (!active) return "none";
  return `${active.tagName.toLowerCase()}${active.id ? "#" + active.id : ""}`;
}

function sampleNow(): void {
  if (!renderer) return;

  const anyRenderer = renderer as unknown as {
    cursorPos?: unknown;
    cursorVisible?: boolean;
    pendingDrawPages?: Set<number>;
    lastCursorActivityMs?: number;
    caretOverlay?: HTMLDivElement;
  };

  const cursorRect = (() => {
    try {
      const rect = renderer.getCursorRect();
      return rect ? { y: rect.y } : null;
    } catch {
      return null;
    }
  })();

  const now = performance.now();
  const lastActivity =
    typeof anyRenderer.lastCursorActivityMs === "number"
      ? Math.max(0, now - anyRenderer.lastCursorActivityMs)
      : -1;

  samples.push({
    t: now,
    hasCursorPos: !!anyRenderer.cursorPos,
    cursorVisible: !!anyRenderer.cursorVisible,
    hasCursorRect: cursorRect != null,
    cursorTop: cursorRect?.y ?? null,
    overlayVisible: anyRenderer.caretOverlay?.style.display !== "none",
    focused: renderer.isFocused(),
    pendingDrawPages: anyRenderer.pendingDrawPages?.size ?? 0,
    lastActivityAgeMs: lastActivity,
    activeTag: activeElementTag(),
  });
}

async function runMonitor(durationMs: number): Promise<void> {
  const start = performance.now();
  const useRaf = !autoRun;
  return new Promise((resolve) => {
    const tick = (): void => {
      sampleNow();
      if (performance.now() - start < durationMs) {
        if (useRaf) {
          requestAnimationFrame(tick);
        } else {
          setTimeout(tick, 16);
        }
      } else {
        resolve();
      }
    };
    tick();
  });
}

function summarize(): SummaryStats {
  const focusedWithCursor = samples.filter((s) => s.focused && s.hasCursorPos);
  const activeWindowAnomalies = focusedWithCursor.filter((s) => {
    return (
      s.lastActivityAgeMs >= 0 &&
      s.lastActivityAgeMs < 900 &&
      (!s.cursorVisible || !s.hasCursorRect || !s.overlayVisible)
    );
  });

  const maxPending = samples.reduce(
    (m, s) => Math.max(m, s.pendingDrawPages),
    0,
  );
  const slow16 = drawSamples.filter((d) => d.durationMs > 16).length;
  const slow33 = drawSamples.filter((d) => d.durationMs > 33).length;
  const maxDraw = drawSamples.reduce((m, d) => Math.max(m, d.durationMs), 0);
  const tops = focusedWithCursor
    .map((s) => s.cursorTop)
    .filter((v): v is number => typeof v === "number" && Number.isFinite(v));
  const cursorYSpanPx =
    tops.length > 1
      ? tops.reduce((m, v) => Math.max(m, v), tops[0]!) -
        tops.reduce((m, v) => Math.min(m, v), tops[0]!)
      : 0;
  const finalSelection = renderer?.getSelection();
  const finalBodyIndex = finalSelection?.focus.bodyIndex ?? -1;
  const finalCharOffset = finalSelection?.focus.charOffset ?? -1;

  const stats: SummaryStats = {
    sampleCount: samples.length,
    drawCallCount: drawSamples.length,
    focusedWithCursorCount: focusedWithCursor.length,
    activeWindowAnomalyCount: activeWindowAnomalies.length,
    cursorYSpanPx,
    finalBodyIndex,
    finalCharOffset,
    maxPendingDrawPages: maxPending,
    slowDrawsOver16Ms: slow16,
    slowDrawsOver33Ms: slow33,
    maxDrawMs: maxDraw,
  };

  summaryEl.textContent = [
    `samples: ${stats.sampleCount}`,
    `draw calls: ${stats.drawCallCount}`,
    `focused+cursor samples: ${stats.focusedWithCursorCount}`,
    `active-window anomalies (<900ms since input): ${stats.activeWindowAnomalyCount}`,
    `cursor Y span px: ${stats.cursorYSpanPx.toFixed(2)}`,
    `final cursor: body=${stats.finalBodyIndex} off=${stats.finalCharOffset}`,
    `max pendingDrawPages: ${stats.maxPendingDrawPages}`,
    `slow draws >16ms: ${stats.slowDrawsOver16Ms}`,
    `slow draws >33ms: ${stats.slowDrawsOver33Ms}`,
    `max draw ms: ${stats.maxDrawMs.toFixed(2)}`,
  ].join("\n");

  const first = activeWindowAnomalies.slice(0, 16);
  if (first.length > 0) {
    appendLog("anomaly samples:");
    for (const s of first) {
      appendLog(
        `t=${s.t.toFixed(1)} vis=${s.cursorVisible} rect=${s.hasCursorRect} pending=${s.pendingDrawPages} idle=${s.lastActivityAgeMs.toFixed(1)} active=${s.activeTag}`,
      );
    }
  }
  return stats;
}

function clickFirstPageTarget(): boolean {
  if (!container.firstElementChild) return false;
  const canvas = container.querySelector("canvas");
  if (!(canvas instanceof HTMLCanvasElement)) return false;

  const rect = canvas.getBoundingClientRect();
  const x = rect.left + Math.max(24, rect.width * 0.2);
  const y = rect.top + Math.max(32, rect.height * 0.14);

  const down = new MouseEvent("mousedown", {
    bubbles: true,
    cancelable: true,
    clientX: x,
    clientY: y,
  });
  const up = new MouseEvent("mouseup", {
    bubbles: true,
    cancelable: true,
    clientX: x,
    clientY: y,
  });
  canvas.dispatchEvent(down);
  canvas.dispatchEvent(up);
  return true;
}

async function typeBurst(count: number, delayMs: number): Promise<void> {
  if (!input) return;

  const textarea = input.getInputElement();
  if (!(textarea instanceof HTMLTextAreaElement)) return;
  textarea.focus();

  const alpha = "asdlkjqwepoiuzmxncbv";
  for (let i = 0; i < count; i++) {
    const ch = alpha[i % alpha.length]!;
    dispatchBeforeInput(textarea, "insertText", ch);
    if (i > 0 && i % 37 === 0) {
      dispatchBeforeInput(textarea, "insertParagraph", null);
    }
    if (delayMs > 0) {
      await sleep(delayMs);
    } else if (i % 8 === 0) {
      await raf();
    }
  }
}

async function loadBlankAndPrime(): Promise<void> {
  if (!controller || !renderer) return;
  controller.loadBlank();
  await raf();
  renderer.setCursor({ bodyIndex: 0, charOffset: 0 });
  await raf();
}

async function runScenario(): Promise<SummaryStats | null> {
  if (!renderer || !controller || running) return null;
  running = true;
  setStatus("running scenario...");
  clearLogs();

  await loadBlankAndPrime();

  if (!clickFirstPageTarget()) {
    appendLog("no canvas page found");
    setStatus("failed: no canvas page");
    running = false;
    return null;
  }
  // Exercise blur -> refocus path before typing.
  btnLoad.focus();
  await sleep(40);
  clickFirstPageTarget();

  const monitor = runMonitor(5200);
  await sleep(60);
  await typeBurst(260, 4);
  await sleep(1400);
  await monitor;

  const stats = summarize();
  setStatus("scenario complete");
  running = false;
  return stats;
}

function exportTelemetry(): void {
  const payload = {
    at: new Date().toISOString(),
    samples,
    drawSamples,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "caret-harness-telemetry.json";
  a.click();
  URL.revokeObjectURL(url);
}

async function boot(): Promise<void> {
  setStatus("initializing...");
  await init();

  renderer = await CanvasRenderer.create(container);
  input = new CanvasInput(container);
  controller = createEditorController(renderer, input, container);
  patchDrawTelemetry(renderer);
  await loadBlankAndPrime();

  appendLog("harness ready");
  setStatus("ready");

  if (autoRun) {
    const stats = await runScenario();
    publishAutoStatus(stats);
  }
}

btnLoad.addEventListener("click", async () => {
  await loadBlankAndPrime();
  appendLog("load blank");
});

btnClick.addEventListener("click", () => {
  const ok = clickFirstPageTarget();
  appendLog(ok ? "clicked first-page target" : "click failed");
});

btnType.addEventListener("click", async () => {
  setStatus("typing burst...");
  await typeBurst(120, 3);
  setStatus("ready");
  appendLog("typed burst");
});

btnScenario.addEventListener("click", async () => {
  await runScenario();
});

btnExport.addEventListener("click", () => {
  exportTelemetry();
  appendLog("exported telemetry json");
});

btnClear.addEventListener("click", () => {
  clearLogs();
});

window.addEventListener("beforeunload", () => {
  controller?.destroy();
});

await boot();
