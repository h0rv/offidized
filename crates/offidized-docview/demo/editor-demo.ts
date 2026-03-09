// Demo editor: file upload, drag-drop, save, dirty indicator, collaborative sync.

import {
  init,
  mountEditor,
  type DocEditorController,
  type FormatAction,
  type FormattingState,
  type MountEditorOpts,
  type ParagraphStylePatch,
  type SyncConfig,
  type TextStylePatch,
} from "../js/doc-edit.ts";
import { restoreSelection, selectionToDocPositions } from "../js/position.ts";

// Initialize WASM
await init();

// ---- Sync configuration from URL params ----

const params = new URLSearchParams(window.location.search);
let roomId = params.get("room");
const wsUrl = params.get("ws");

// Read renderer preference from URL: ?renderer=canvas or ?renderer=html.
// Without a param, mountEditor() uses its default renderer (canvas).
const rendererParam = params.get("renderer");
const rendererOpt: "html" | "canvas" | undefined =
  rendererParam === "canvas"
    ? "canvas"
    : rendererParam === "html"
      ? "html"
      : undefined;
const autoRunMode = params.get("autorun");
const autoRunCaret = autoRunMode === "1";
const autoRunOps = autoRunMode === "ops";
const autoRunCopyLink = autoRunMode === "copy-link";
const autoRunCollab = autoRunMode === "collab";
const autoRunCollabPeer = autoRunMode === "collab-peer";
const autoRunCollabPresenceExpiry = autoRunMode === "collab-presence-expiry";
const autoRunCollabPresenceExpiryPeer =
  autoRunMode === "collab-presence-expiry-peer";
const useLocalRelay = params.get("relay") === "local";
const DEMO_FONT_MANIFEST = [
  { family: "PT Sans", url: "/__demo-fonts/pt-sans.ttc" },
  { family: "PT Serif", url: "/__demo-fonts/pt-serif.ttc" },
  { family: "PT Mono", url: "/__demo-fonts/pt-mono.ttc" },
  { family: "Comic Sans MS", url: "/__demo-fonts/comic-sans.ttf" },
] as const;

interface LocalRelayEnvelope {
  sender: string;
  payload: number[];
}

function toUint8Array(data: ArrayBuffer | ArrayBufferView): Uint8Array {
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
}

class LocalRelayWebSocket extends EventTarget {
  static readonly CONNECTING = 0;
  static readonly OPEN = 1;
  static readonly CLOSING = 2;
  static readonly CLOSED = 3;

  readonly CONNECTING = LocalRelayWebSocket.CONNECTING;
  readonly OPEN = LocalRelayWebSocket.OPEN;
  readonly CLOSING = LocalRelayWebSocket.CLOSING;
  readonly CLOSED = LocalRelayWebSocket.CLOSED;

  binaryType: BinaryType = "blob";
  readyState = LocalRelayWebSocket.CONNECTING;
  onopen: ((this: WebSocket, ev: Event) => unknown) | null = null;
  onmessage: ((this: WebSocket, ev: MessageEvent) => unknown) | null = null;
  onerror: ((this: WebSocket, ev: Event) => unknown) | null = null;
  onclose: ((this: WebSocket, ev: CloseEvent) => unknown) | null = null;

  private readonly channel: BroadcastChannel;
  private readonly senderId: string;
  private closed = false;

  constructor(url: string | URL) {
    super();
    const parsed = new URL(String(url), window.location.href);
    const roomMatch = parsed.pathname.match(/^\/doc\/(.+)$/);
    const roomName = roomMatch?.[1] ?? "default";
    this.channel = new BroadcastChannel(`docview-local-relay:${roomName}`);
    this.senderId = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    this.channel.addEventListener("message", (event: MessageEvent) => {
      if (this.closed || this.readyState !== LocalRelayWebSocket.OPEN) return;
      const msg = event.data as LocalRelayEnvelope | undefined;
      if (!msg || msg.sender === this.senderId || !Array.isArray(msg.payload)) {
        return;
      }
      const payload = new Uint8Array(msg.payload);
      const evt = new MessageEvent("message", { data: payload.buffer });
      this.dispatchEvent(evt);
      this.onmessage?.call(this as unknown as WebSocket, evt);
    });

    queueMicrotask(() => {
      if (this.closed) return;
      this.readyState = LocalRelayWebSocket.OPEN;
      const evt = new Event("open");
      this.dispatchEvent(evt);
      this.onopen?.call(this as unknown as WebSocket, evt);
    });
  }

  send(data: ArrayBuffer | ArrayBufferView | Blob | string): void {
    if (this.closed || this.readyState !== LocalRelayWebSocket.OPEN) return;
    if (typeof data === "string" || data instanceof Blob) return;
    const bytes = toUint8Array(data);
    const msg: LocalRelayEnvelope = {
      sender: this.senderId,
      payload: Array.from(bytes),
    };
    this.channel.postMessage(msg);
  }

  close(code?: number, reason?: string): void {
    if (this.closed) return;
    this.closed = true;
    this.readyState = LocalRelayWebSocket.CLOSED;
    this.channel.close();
    const evt = new CloseEvent("close", {
      code: code ?? 1000,
      reason: reason ?? "",
      wasClean: true,
    });
    this.dispatchEvent(evt);
    this.onclose?.call(this as unknown as WebSocket, evt);
  }
}

function installLocalRelayWebSocket(): void {
  if (
    (window as { __docviewLocalRelayInstalled?: boolean })
      .__docviewLocalRelayInstalled
  ) {
    return;
  }
  (
    window as unknown as {
      WebSocket: typeof WebSocket;
      __docviewLocalRelayInstalled?: boolean;
    }
  ).WebSocket = LocalRelayWebSocket as unknown as typeof WebSocket;
  (
    window as { __docviewLocalRelayInstalled?: boolean }
  ).__docviewLocalRelayInstalled = true;
}

if (useLocalRelay) {
  installLocalRelayWebSocket();
}

/** Build sync config when a room ID is present. */
function getSyncConfig(): SyncConfig | undefined {
  if (!roomId) return undefined;
  if (wsUrl) return { roomId, wsUrl };
  return { roomId };
}

// ---- DOM references ----

const fileInput = document.getElementById("file-input") as HTMLInputElement;
const editorContainer = document.getElementById(
  "editor-container",
) as HTMLDivElement;
const emptyState = document.getElementById("empty-state") as HTMLDivElement;
const statusEl = document.getElementById("status") as HTMLSpanElement;
const btnNew = document.getElementById("btn-new") as HTMLButtonElement;
const btnSave = document.getElementById("btn-save") as HTMLButtonElement;
const dirtyIndicator = document.getElementById(
  "dirty-indicator",
) as HTMLSpanElement;
const dropOverlay = document.getElementById("drop-overlay") as HTMLDivElement;

const fmtBold = document.getElementById("fmt-bold") as HTMLButtonElement;
const fmtItalic = document.getElementById("fmt-italic") as HTMLButtonElement;
const fmtLink = document.getElementById("fmt-link") as HTMLButtonElement;
const fmtUnderline = document.getElementById(
  "fmt-underline",
) as HTMLButtonElement;
const fmtStrike = document.getElementById("fmt-strike") as HTMLButtonElement;
const fmtBullet = document.getElementById("fmt-bullet") as HTMLButtonElement;
const fmtNumbered = document.getElementById(
  "fmt-numbered",
) as HTMLButtonElement;
const fmtTable = document.getElementById("fmt-table") as HTMLButtonElement;
const fmtTableRowAdd = document.getElementById(
  "fmt-table-row-add",
) as HTMLButtonElement;
const fmtTableRowRemove = document.getElementById(
  "fmt-table-row-remove",
) as HTMLButtonElement;
const fmtTableColAdd = document.getElementById(
  "fmt-table-col-add",
) as HTMLButtonElement;
const fmtTableColRemove = document.getElementById(
  "fmt-table-col-remove",
) as HTMLButtonElement;
const fmtImage = document.getElementById("fmt-image") as HTMLButtonElement;
const imageInput = document.getElementById("image-input") as HTMLInputElement;
const fmtBlock = document.getElementById("fmt-block") as HTMLSelectElement;
const fmtFontFamily = document.getElementById(
  "fmt-font-family",
) as HTMLSelectElement;
const fmtFontSize = document.getElementById(
  "fmt-font-size",
) as HTMLSelectElement;
const fmtAlign = document.getElementById("fmt-align") as HTMLSelectElement;
const fmtSpaceBefore = document.getElementById(
  "fmt-space-before",
) as HTMLSelectElement;
const fmtSpaceAfter = document.getElementById(
  "fmt-space-after",
) as HTMLSelectElement;
const fmtLineSpacing = document.getElementById(
  "fmt-line-spacing",
) as HTMLSelectElement;
const fmtIndentLeft = document.getElementById(
  "fmt-indent-left",
) as HTMLSelectElement;
const fmtIndentFirstLine = document.getElementById(
  "fmt-indent-first-line",
) as HTMLSelectElement;
const fmtHighlight = document.getElementById(
  "fmt-highlight",
) as HTMLSelectElement;
const fmtColor = document.getElementById("fmt-color") as HTMLInputElement;

const btnCollab = document.getElementById("btn-collab") as HTMLButtonElement;
const collabInfo = document.getElementById("collab-info") as HTMLSpanElement;

const btnRendererHtml = document.getElementById(
  "btn-renderer-html",
) as HTMLButtonElement;
const btnRendererCanvas = document.getElementById(
  "btn-renderer-canvas",
) as HTMLButtonElement;
const rendererStatus = document.getElementById(
  "renderer-status",
) as HTMLSpanElement;

let editor: DocEditorController | null = null;
let currentFileName = "document.docx";
let collabModeLabel = "";
let updatingFormatUI = false;
const TABLE_INSERT_ROWS = 3;
const TABLE_INSERT_COLUMNS = 3;

interface SyncDebugCounters {
  awarenessByeRecv: number;
  awarenessByeSend: number;
  awarenessClear: number;
  awarenessExpire: number;
  awarenessRecv: number;
  awarenessSend: number;
  bcSend: number;
  bcRecv: number;
  divergenceCleared: number;
  divergenceDetected: number;
  repairRequests: number;
  resyncing: boolean;
  divergenceSuspected: boolean;
  sawRepairing: boolean;
  sawDesync: boolean;
  stateRequestSend: number;
  wsSend: number;
  wsRecv: number;
  wsOpen: number;
  wsClose: number;
  wsError: number;
  mode: string;
}

const syncDebug: SyncDebugCounters = {
  awarenessByeRecv: 0,
  awarenessByeSend: 0,
  awarenessClear: 0,
  awarenessExpire: 0,
  awarenessRecv: 0,
  awarenessSend: 0,
  bcSend: 0,
  bcRecv: 0,
  divergenceCleared: 0,
  divergenceDetected: 0,
  repairRequests: 0,
  resyncing: false,
  divergenceSuspected: false,
  sawRepairing: false,
  sawDesync: false,
  stateRequestSend: 0,
  wsSend: 0,
  wsRecv: 0,
  wsOpen: 0,
  wsClose: 0,
  wsError: 0,
  mode: "unknown",
};

function syncDebugSummary(): string {
  return `mode=${syncDebug.mode} repair=${syncDebug.repairRequests} resync=${syncDebug.resyncing ? 1 : 0} diverged=${syncDebug.divergenceSuspected ? 1 : 0} stateReq=${syncDebug.stateRequestSend} div=${syncDebug.divergenceDetected}/${syncDebug.divergenceCleared} bc=${syncDebug.bcSend}/${syncDebug.bcRecv} aw=${syncDebug.awarenessSend}/${syncDebug.awarenessRecv} bye=${syncDebug.awarenessByeSend}/${syncDebug.awarenessByeRecv} clear=${syncDebug.awarenessClear} expire=${syncDebug.awarenessExpire} ws=${syncDebug.wsSend}/${syncDebug.wsRecv} open=${syncDebug.wsOpen} close=${syncDebug.wsClose} err=${syncDebug.wsError}`;
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitFor(
  predicate: () => boolean,
  timeoutMs: number,
  intervalMs = 25,
): Promise<boolean> {
  const started = performance.now();
  while (performance.now() - started <= timeoutMs) {
    if (predicate()) return true;
    await sleep(intervalMs);
  }
  return predicate();
}

function fileToDataUrl(file: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === "string") {
        resolve(reader.result);
        return;
      }
      reject(new Error("image file reader returned a non-string result"));
    };
    reader.onerror = () => {
      reject(reader.error ?? new Error("failed to read image file"));
    };
    reader.readAsDataURL(file);
  });
}

function measureImagePt(
  dataUrl: string,
): Promise<{ widthPt: number; heightPt: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const widthPx = Math.max(1, img.naturalWidth || img.width || 1);
      const heightPx = Math.max(1, img.naturalHeight || img.height || 1);
      resolve({
        widthPt: (widthPx * 72) / 96,
        heightPt: (heightPx * 72) / 96,
      });
    };
    img.onerror = () => reject(new Error("failed to decode image for sizing"));
    img.src = dataUrl;
  });
}

async function insertInlineImageFromFile(file: File): Promise<boolean> {
  const ed = await ensureEditor();
  const dataUrl = await fileToDataUrl(file);
  const size = await measureImagePt(dataUrl);
  const ok = ed.insertInlineImage({
    dataUri: dataUrl,
    widthPt: size.widthPt,
    heightPt: size.heightPt,
    name: file.name || undefined,
    description: file.name || undefined,
  });
  if (!ok) {
    throw new Error("editor rejected inline image insertion");
  }
  updateDirtyState();
  return true;
}

function describeError(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (
    err &&
    typeof err === "object" &&
    "message" in err &&
    typeof (err as { message?: unknown }).message === "string"
  ) {
    return (err as { message: string }).message;
  }
  try {
    return JSON.stringify(err);
  } catch {
    return String(err);
  }
}

function dispatchBeforeInput(
  target: HTMLElement,
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
  return target.dispatchEvent(event);
}

// ---- Collab button ----

function generateRoomId(): string {
  const chars = "abcdefghijklmnopqrstuvwxyz0123456789";
  let result = "";
  for (let i = 0; i < 8; i++) {
    result += chars[Math.floor(Math.random() * chars.length)];
  }
  return result;
}

/** Update the collab UI to reflect current sync state. */
function updateCollabUI(): void {
  if (roomId) {
    const suffix = collabModeLabel ? ` (${collabModeLabel})` : "";
    collabInfo.textContent = `Room: ${roomId}${suffix}`;
    collabInfo.style.display = "inline";
    btnCollab.textContent = "Copy Link";
    btnCollab.title = "Copy collaboration link to clipboard";
  } else {
    collabInfo.style.display = "none";
    btnCollab.textContent = "Collab";
    btnCollab.title = "Start a collaborative editing session";
  }
}

btnCollab.addEventListener("click", () => {
  if (roomId) {
    // Already in a room -- copy the link.
    const url = new URL(window.location.href);
    url.searchParams.set("room", roomId);
    if (wsUrl) {
      url.searchParams.set("ws", wsUrl);
    }
    navigator.clipboard.writeText(url.toString()).then(
      () => {
        statusEl.textContent = "Collaboration link copied!";
      },
      () => {
        statusEl.textContent = "Failed to copy link";
      },
    );
  } else {
    // Start a new room: reload with a room param.
    roomId = generateRoomId();
    const url = new URL(window.location.href);
    url.searchParams.set("room", roomId);
    if (wsUrl) {
      url.searchParams.set("ws", wsUrl);
    }
    window.location.href = url.toString();
  }
});

// Initialize collab UI.
updateCollabUI();

// ---- Mount editor ----

async function ensureEditor(): Promise<DocEditorController> {
  if (!editor) {
    emptyState.style.display = "none";
    const opts: MountEditorOpts = {
      renderer: rendererOpt,
      localFontManifestProvider: () => DEMO_FONT_MANIFEST,
      sync: getSyncConfig(),
    };
    editor = await mountEditor(editorContainer, opts);

    // Highlight format bar buttons based on formatting at cursor.
    editor.onFormattingChange((state: FormattingState) => {
      updatingFormatUI = true;
      fmtBold.classList.toggle("active", !!state.bold);
      fmtItalic.classList.toggle("active", !!state.italic);
      fmtLink.classList.toggle("active", !!state.hyperlink);
      fmtUnderline.classList.toggle("active", !!state.underline);
      fmtStrike.classList.toggle("active", !!state.strike);
      fmtBullet.classList.toggle("active", state.listKind === "bullet");
      fmtNumbered.classList.toggle("active", state.listKind === "decimal");
      fmtBlock.value = headingValueFromState(state);
      fmtFontFamily.value = normalizeFontFamilyValue(state.fontFamily);
      fmtFontSize.value = normalizeFontSizeValue(state.fontSizePt);
      fmtAlign.value = normalizeAlignmentValue(state.alignment);
      fmtSpaceBefore.value = normalizeNumericSelectValue(state.spacingBeforePt);
      fmtSpaceAfter.value = normalizeNumericSelectValue(state.spacingAfterPt);
      fmtLineSpacing.value = normalizeLineSpacingValue(
        state.lineSpacingMultiple,
      );
      fmtIndentLeft.value = normalizeNumericSelectValue(state.indentLeftPt);
      fmtIndentFirstLine.value = normalizeNumericSelectValue(
        state.indentFirstLinePt,
      );
      fmtHighlight.value = normalizeHighlightValue(state.highlight);
      fmtColor.value = normalizeColorValue(state.color);
      updatingFormatUI = false;
    });
    editor.onActiveTableCellChange((state) => {
      setTableActionButtonsEnabled(state != null);
    });
    setTableActionButtonsEnabled(editor.getActiveTableCell() != null);
  }
  return editor;
}

// ---- Dirty tracking ----

function updateDirtyState(): void {
  const dirty = editor?.isDirty() ?? false;
  dirtyIndicator.classList.toggle("visible", dirty);
}

function setTableActionButtonsEnabled(enabled: boolean): void {
  fmtTableRowAdd.disabled = !enabled;
  fmtTableRowRemove.disabled = !enabled;
  fmtTableColAdd.disabled = !enabled;
  fmtTableColRemove.disabled = !enabled;
}

setTableActionButtonsEnabled(false);

// Listen for dirty events from the editor
editorContainer.addEventListener("docedit-dirty", () => {
  updateDirtyState();
});

// Listen for Ctrl+S save events from the editor
editorContainer.addEventListener("docedit-save", () => {
  saveFile();
});

// Surface sync transport state in the collab label.
editorContainer.addEventListener("docedit-sync", (event: Event) => {
  const customEvent = event as CustomEvent<{
    mode?: "offline" | "broadcast" | "websocket" | "hybrid";
    resyncing?: boolean;
    divergenceSuspected?: boolean;
    repairRequests?: number;
  }>;
  const mode = customEvent.detail?.mode;
  const resyncing = customEvent.detail?.resyncing === true;
  const divergenceSuspected = customEvent.detail?.divergenceSuspected === true;
  if (mode === "websocket") collabModeLabel = "WS";
  else if (mode === "broadcast") collabModeLabel = "Local";
  else if (mode === "hybrid") collabModeLabel = "WS+Local";
  else collabModeLabel = "Offline";
  if (divergenceSuspected) collabModeLabel += " Desync";
  else if (resyncing) collabModeLabel += " Repairing";
  syncDebug.mode = mode ?? "offline";
  syncDebug.resyncing = resyncing;
  syncDebug.divergenceSuspected = divergenceSuspected;
  syncDebug.sawRepairing ||= resyncing;
  syncDebug.sawDesync ||= divergenceSuspected;
  syncDebug.repairRequests = customEvent.detail?.repairRequests ?? 0;
  updateCollabUI();
});

window.addEventListener("docedit-sync-debug", (event: Event) => {
  const customEvent = event as CustomEvent<{
    kind?:
      | "awareness-bye-recv"
      | "awareness-bye-send"
      | "awareness-clear"
      | "awareness-expire"
      | "awareness-recv"
      | "awareness-send"
      | "bc-connect"
      | "bc-send"
      | "bc-recv"
      | "divergence-cleared"
      | "divergence-detected"
      | "resync-send"
      | "state-request-send"
      | "ws-open"
      | "ws-close"
      | "ws-error"
      | "ws-send"
      | "ws-recv";
  }>;
  const kind = customEvent.detail?.kind;
  if (kind === "awareness-bye-recv") syncDebug.awarenessByeRecv += 1;
  else if (kind === "awareness-bye-send") syncDebug.awarenessByeSend += 1;
  else if (kind === "awareness-clear") syncDebug.awarenessClear += 1;
  else if (kind === "awareness-expire") syncDebug.awarenessExpire += 1;
  else if (kind === "awareness-recv") syncDebug.awarenessRecv += 1;
  else if (kind === "awareness-send") syncDebug.awarenessSend += 1;
  else if (kind === "bc-send") syncDebug.bcSend += 1;
  else if (kind === "bc-recv") syncDebug.bcRecv += 1;
  else if (kind === "divergence-cleared") syncDebug.divergenceCleared += 1;
  else if (kind === "divergence-detected") syncDebug.divergenceDetected += 1;
  else if (kind === "state-request-send") syncDebug.stateRequestSend += 1;
  else if (kind === "ws-send") syncDebug.wsSend += 1;
  else if (kind === "ws-recv") syncDebug.wsRecv += 1;
  else if (kind === "ws-open") syncDebug.wsOpen += 1;
  else if (kind === "ws-close") syncDebug.wsClose += 1;
  else if (kind === "ws-error") syncDebug.wsError += 1;
});

// ---- File handling ----

async function handleFile(file: File): Promise<void> {
  if (!file.name.endsWith(".docx")) {
    statusEl.textContent = "Only .docx files are supported";
    return;
  }

  try {
    currentFileName = file.name;
    const ed = await ensureEditor();
    const data = new Uint8Array(await file.arrayBuffer());
    const ms = roomId ? ed.replace(data) : ed.load(data);
    statusEl.textContent = `${file.name} \u2014 parsed in ${ms}ms`;
    updateDirtyState();
  } catch (err) {
    statusEl.textContent = `Error: ${(err as Error).message}`;
  }
}

// ---- Save ----

function saveFile(): void {
  if (!editor) return;
  try {
    const bytes = editor.save();
    const blob = new Blob([bytes as BlobPart], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = currentFileName;
    a.click();
    URL.revokeObjectURL(url);
    updateDirtyState();
    statusEl.textContent = `Saved ${currentFileName}`;
  } catch (err) {
    statusEl.textContent = `Save error: ${(err as Error).message}`;
  }
}

btnNew.addEventListener("click", async () => {
  try {
    currentFileName = "untitled.docx";
    const ed = await ensureEditor();
    if (roomId) ed.replaceBlank();
    else ed.loadBlank();
    statusEl.textContent = "New document";
    updateDirtyState();
  } catch (err) {
    statusEl.textContent = `Error: ${(err as Error).message}`;
  }
});

btnSave.addEventListener("click", saveFile);

// ---- File input ----

fileInput.addEventListener("change", () => {
  const file = fileInput.files?.[0];
  if (file) handleFile(file);
});

// ---- Drag and drop ----

let dragCounter = 0;

document.addEventListener("dragenter", (e) => {
  e.preventDefault();
  dragCounter++;
  dropOverlay.classList.add("active");
});

document.addEventListener("dragleave", () => {
  dragCounter--;
  if (dragCounter <= 0) {
    dragCounter = 0;
    dropOverlay.classList.remove("active");
  }
});

document.addEventListener("dragover", (e) => {
  e.preventDefault();
});

document.addEventListener("drop", (e) => {
  e.preventDefault();
  dragCounter = 0;
  dropOverlay.classList.remove("active");

  const file = e.dataTransfer?.files[0];
  if (file) handleFile(file);
});

// ---- Format toolbar ----

function onFormat(action: FormatAction): void {
  editor?.format(action);
}

function normalizeColorValue(color?: string): string {
  if (!color) return "#000000";
  const normalized = color.replace(/^#/, "").slice(0, 6).padStart(6, "0");
  return `#${normalized}`;
}

function normalizeFontFamilyValue(fontFamily?: string): string {
  if (!fontFamily) return "";
  const normalized = fontFamily.trim().toLowerCase();
  if (normalized.includes("comic")) {
    return "Comic Sans MS";
  }
  if (
    normalized.includes("mono") ||
    normalized.includes("code") ||
    normalized.includes("consolas") ||
    normalized.includes("courier")
  ) {
    return "PT Mono";
  }
  if (
    normalized.includes("serif") ||
    normalized.includes("cambria") ||
    normalized.includes("times") ||
    normalized.includes("georgia") ||
    normalized.includes("garamond")
  ) {
    return "PT Serif";
  }
  if (
    normalized.includes("sans") ||
    normalized.includes("calibri") ||
    normalized.includes("arial") ||
    normalized.includes("helvetica")
  ) {
    return "PT Sans";
  }
  const option = Array.from(fmtFontFamily.options).find(
    (item) => item.value === fontFamily,
  );
  return option ? option.value : "";
}

function normalizeFontSizeValue(fontSizePt?: number): string {
  if (typeof fontSizePt !== "number" || !Number.isFinite(fontSizePt)) return "";
  const rounded = Math.round(fontSizePt);
  const option = Array.from(fmtFontSize.options).find(
    (item) => Number(item.value) === rounded,
  );
  return option ? option.value : "";
}

function normalizeAlignmentValue(alignment?: string): string {
  return alignment === "center" ||
    alignment === "right" ||
    alignment === "justify" ||
    alignment === "left"
    ? alignment
    : "";
}

function normalizeNumericSelectValue(value?: number): string {
  if (typeof value !== "number" || !Number.isFinite(value)) return "";
  const rounded = Math.round(value);
  return String(rounded);
}

function normalizeLineSpacingValue(value?: number): string {
  if (typeof value !== "number" || !Number.isFinite(value)) return "";
  const options = Array.from(fmtLineSpacing.options).map((item) => item.value);
  const exact = value.toString();
  if (options.includes(exact)) return exact;
  const rounded = Math.round(value * 100) / 100;
  const normalized = rounded.toString();
  return options.includes(normalized) ? normalized : "";
}

function normalizeHyperlinkValue(value: string): string {
  const trimmed = value.trim();
  if (!trimmed) return "";
  if (
    /^https?:\/\//i.test(trimmed) ||
    /^mailto:/i.test(trimmed) ||
    /^tel:/i.test(trimmed)
  ) {
    return trimmed;
  }
  return `https://${trimmed}`;
}

function normalizeHighlightValue(value?: string): string {
  if (!value) return "";
  const option = Array.from(fmtHighlight.options).find(
    (item) => item.value === value,
  );
  return option ? option.value : "";
}

function headingValueFromState(state: FormattingState): string {
  if (typeof state.headingLevel === "number" && state.headingLevel >= 1) {
    return `heading${Math.min(3, Math.round(state.headingLevel))}`;
  }
  return "paragraph";
}

function setTextStyle(patch: TextStylePatch): void {
  editor?.setTextStyle(patch);
}

function setParagraphStyle(patch: ParagraphStylePatch): void {
  editor?.setParagraphStyle(patch);
}

function refocusEditorInput(): void {
  const tableCellEditor = editorContainer.querySelector<HTMLTextAreaElement>(
    'textarea[data-docview-table-cell-editor="1"]',
  );
  if (
    tableCellEditor &&
    tableCellEditor.style.display !== "none" &&
    editor?.getActiveTableCell()
  ) {
    tableCellEditor.focus({ preventScroll: true });
    return;
  }
  const input = editorContainer.querySelector<HTMLElement>(
    'textarea[data-docview-canvas-input="1"], [contenteditable="true"]',
  );
  input?.focus({ preventScroll: true });
}

fmtBold.addEventListener("mousedown", (e) => {
  e.preventDefault();
  onFormat("bold");
});
fmtItalic.addEventListener("mousedown", (e) => {
  e.preventDefault();
  onFormat("italic");
});
fmtLink.addEventListener("mousedown", (e) => {
  e.preventDefault();
  const current = editor?.getFormattingState().hyperlink ?? "";
  const entered = window.prompt(
    "Set link URL. Leave blank to remove the link.",
    current,
  );
  if (entered == null) return;
  setTextStyle({
    hyperlink: entered.trim() ? normalizeHyperlinkValue(entered) : null,
  });
  queueMicrotask(refocusEditorInput);
});
fmtUnderline.addEventListener("mousedown", (e) => {
  e.preventDefault();
  onFormat("underline");
});
fmtStrike.addEventListener("mousedown", (e) => {
  e.preventDefault();
  onFormat("strikethrough");
});
fmtBullet.addEventListener("mousedown", (e) => {
  e.preventDefault();
  editor?.toggleList("bullet");
  queueMicrotask(refocusEditorInput);
});
fmtNumbered.addEventListener("mousedown", (e) => {
  e.preventDefault();
  editor?.toggleList("decimal");
  queueMicrotask(refocusEditorInput);
});
fmtTable.addEventListener("mousedown", async (e) => {
  e.preventDefault();
  try {
    const ed = await ensureEditor();
    const ok = ed.insertTable(TABLE_INSERT_ROWS, TABLE_INSERT_COLUMNS);
    statusEl.textContent = ok
      ? `Inserted ${TABLE_INSERT_ROWS}x${TABLE_INSERT_COLUMNS} table`
      : "Table insert failed";
    updateDirtyState();
  } catch (err) {
    statusEl.textContent = `Table insert error: ${describeError(err)}`;
  }
});
fmtTableRowAdd.addEventListener("mousedown", (e) => {
  e.preventDefault();
  const ok = editor?.insertTableRow() ?? false;
  statusEl.textContent = ok ? "Inserted row" : "Row insert failed";
  queueMicrotask(refocusEditorInput);
  updateDirtyState();
});
fmtTableRowRemove.addEventListener("mousedown", (e) => {
  e.preventDefault();
  const ok = editor?.removeTableRow() ?? false;
  statusEl.textContent = ok ? "Removed row" : "Row remove failed";
  queueMicrotask(refocusEditorInput);
  updateDirtyState();
});
fmtTableColAdd.addEventListener("mousedown", (e) => {
  e.preventDefault();
  const ok = editor?.insertTableColumn() ?? false;
  statusEl.textContent = ok ? "Inserted column" : "Column insert failed";
  queueMicrotask(refocusEditorInput);
  updateDirtyState();
});
fmtTableColRemove.addEventListener("mousedown", (e) => {
  e.preventDefault();
  const ok = editor?.removeTableColumn() ?? false;
  statusEl.textContent = ok ? "Removed column" : "Column remove failed";
  queueMicrotask(refocusEditorInput);
  updateDirtyState();
});
fmtImage.addEventListener("mousedown", (e) => {
  e.preventDefault();
  imageInput.click();
});
imageInput.addEventListener("change", async () => {
  const file = imageInput.files?.[0];
  imageInput.value = "";
  if (!file) return;
  try {
    await insertInlineImageFromFile(file);
    statusEl.textContent = `Inserted image: ${file.name}`;
    queueMicrotask(refocusEditorInput);
  } catch (err) {
    statusEl.textContent = `Image insert error: ${describeError(err)}`;
  }
});

fmtBlock.addEventListener("change", () => {
  if (updatingFormatUI) return;
  if (fmtBlock.value === "paragraph") {
    setParagraphStyle({ headingLevel: null });
    queueMicrotask(refocusEditorInput);
    return;
  }
  const match = fmtBlock.value.match(/^heading(\d)$/);
  if (!match) return;
  setParagraphStyle({ headingLevel: Number(match[1]) });
  queueMicrotask(refocusEditorInput);
});

fmtFontFamily.addEventListener("change", () => {
  if (updatingFormatUI) return;
  setTextStyle({
    fontFamily: fmtFontFamily.value ? fmtFontFamily.value : null,
  });
  queueMicrotask(refocusEditorInput);
});

fmtFontSize.addEventListener("change", () => {
  if (updatingFormatUI) return;
  setTextStyle({
    fontSizePt: fmtFontSize.value ? Number(fmtFontSize.value) : null,
  });
  queueMicrotask(refocusEditorInput);
});

fmtAlign.addEventListener("change", () => {
  if (updatingFormatUI) return;
  setParagraphStyle({
    alignment: fmtAlign.value
      ? (fmtAlign.value as "left" | "center" | "right" | "justify")
      : null,
  });
  queueMicrotask(refocusEditorInput);
});

fmtSpaceBefore.addEventListener("change", () => {
  if (updatingFormatUI) return;
  setParagraphStyle({
    spacingBeforePt: fmtSpaceBefore.value ? Number(fmtSpaceBefore.value) : null,
  });
  queueMicrotask(refocusEditorInput);
});

fmtSpaceAfter.addEventListener("change", () => {
  if (updatingFormatUI) return;
  setParagraphStyle({
    spacingAfterPt: fmtSpaceAfter.value ? Number(fmtSpaceAfter.value) : null,
  });
  queueMicrotask(refocusEditorInput);
});

fmtLineSpacing.addEventListener("change", () => {
  if (updatingFormatUI) return;
  setParagraphStyle({
    lineSpacingMultiple: fmtLineSpacing.value
      ? Number(fmtLineSpacing.value)
      : null,
  });
  queueMicrotask(refocusEditorInput);
});

fmtIndentLeft.addEventListener("change", () => {
  if (updatingFormatUI) return;
  setParagraphStyle({
    indentLeftPt: fmtIndentLeft.value ? Number(fmtIndentLeft.value) : null,
  });
  queueMicrotask(refocusEditorInput);
});

fmtIndentFirstLine.addEventListener("change", () => {
  if (updatingFormatUI) return;
  setParagraphStyle({
    indentFirstLinePt: fmtIndentFirstLine.value
      ? Number(fmtIndentFirstLine.value)
      : null,
  });
  queueMicrotask(refocusEditorInput);
});

fmtHighlight.addEventListener("change", () => {
  if (updatingFormatUI) return;
  setTextStyle({
    highlight:
      fmtHighlight.value && fmtHighlight.value !== "none"
        ? fmtHighlight.value
        : null,
  });
  queueMicrotask(refocusEditorInput);
});

fmtColor.addEventListener("input", () => {
  if (updatingFormatUI) return;
  setTextStyle({ color: fmtColor.value });
  queueMicrotask(refocusEditorInput);
});

// ---- Renderer toggle ----

/** Switch renderer by reloading the page with the new ?renderer= param. */
function switchRenderer(target: "html" | "canvas"): void {
  const url = new URL(window.location.href);
  url.searchParams.set("renderer", target);
  window.location.href = url.toString();
}

// Highlight the active renderer button
const activeRenderer = rendererOpt ?? "canvas";
btnRendererHtml.classList.toggle("active", activeRenderer === "html");
btnRendererCanvas.classList.toggle("active", activeRenderer === "canvas");
rendererStatus.textContent = activeRenderer === "canvas" ? "Canvas" : "HTML";

btnRendererHtml.addEventListener("click", () => {
  if (activeRenderer !== "html") switchRenderer("html");
});
btnRendererCanvas.addEventListener("click", () => {
  if (activeRenderer !== "canvas") switchRenderer("canvas");
});

// ---- Startup document ----
// Default to a ready-to-edit blank document so the user can type immediately.
const isAutorunMode =
  autoRunCaret ||
  autoRunCopyLink ||
  autoRunOps ||
  autoRunCollab ||
  autoRunCollabPeer ||
  autoRunCollabPresenceExpiry ||
  autoRunCollabPresenceExpiryPeer;
if (!isAutorunMode) {
  ensureEditor()
    .then((ed) => {
      ed.loadBlank();
      updateDirtyState();
      statusEl.textContent = roomId
        ? `Collaborative session: ${roomId}`
        : "New document";
    })
    .catch((err) => {
      statusEl.textContent = `Error: ${(err as Error).message}`;
    });
}

async function runAutoCanvasCursorScenario(): Promise<void> {
  if (!autoRunCaret || rendererOpt === "html") return;

  const publish = (
    status: string,
    anomalies: number,
    ySpan: number,
    finalTop: number,
    stagnantMoves: number,
    textHead: string,
    textTail: string,
  ): void => {
    document.body.dataset.editorCaretDone = "1";
    document.body.dataset.editorCaretStatus = status;
    document.body.dataset.editorCaretAnomalies = String(anomalies);
    document.body.dataset.editorCaretYSpan = String(ySpan);
    document.body.dataset.editorCaretFinalTop = String(finalTop);
    document.body.dataset.editorCaretStagnant = String(stagnantMoves);
    document.body.dataset.editorCaretTextHead = textHead;
    document.body.dataset.editorCaretTextTail = textTail;
  };

  try {
    const ed = await ensureEditor();
    ed.loadBlank();
    await sleep(120);

    const canvas = editorContainer.querySelector("canvas");
    const textarea = editorContainer.querySelector(
      'textarea[data-docview-canvas-input="1"]',
    );
    if (!(canvas instanceof HTMLCanvasElement)) {
      statusEl.textContent = "autorun failed: no canvas";
      publish("failed:no-canvas", -1, -1, -1, -1, "", "");
      return;
    }
    if (!(textarea instanceof HTMLTextAreaElement)) {
      statusEl.textContent = "autorun failed: no textarea";
      publish("failed:no-input", -1, -1, -1, -1, "", "");
      return;
    }

    const rect = canvas.getBoundingClientRect();
    const clickX = rect.left + Math.max(24, rect.width * 0.2);
    const clickY = rect.top + Math.max(32, rect.height * 0.14);
    canvas.dispatchEvent(
      new MouseEvent("mousedown", {
        bubbles: true,
        cancelable: true,
        clientX: clickX,
        clientY: clickY,
      }),
    );
    canvas.dispatchEvent(
      new MouseEvent("mouseup", {
        bubbles: true,
        cancelable: true,
        clientX: clickX,
        clientY: clickY,
      }),
    );
    await sleep(60);

    textarea.focus();
    const sampleTops: number[] = [];
    let anomalyCount = 0;
    let stagnantMoves = 0;
    let lastInputAt = performance.now();
    let monitoring = true;
    const monitorTask = (async () => {
      while (monitoring) {
        const overlay = editorContainer.querySelector(
          'div[data-docview-caret-overlay="1"]',
        ) as HTMLDivElement | null;
        const visible = !!overlay && overlay.style.display !== "none";
        if (visible) {
          const top = Number.parseFloat(overlay.style.top || "NaN");
          if (Number.isFinite(top)) sampleTops.push(top);
        }
        if (performance.now() - lastInputAt < 900 && !visible) {
          anomalyCount += 1;
        }
        await sleep(16);
      }
    })();

    const readOverlayPos = (): { left: number; top: number } | null => {
      const overlay = editorContainer.querySelector(
        'div[data-docview-caret-overlay="1"]',
      ) as HTMLDivElement | null;
      if (!overlay || overlay.style.display === "none") return null;
      const left = Number.parseFloat(overlay.style.left || "NaN");
      const top = Number.parseFloat(overlay.style.top || "NaN");
      if (!Number.isFinite(left) || !Number.isFinite(top)) return null;
      return { left, top };
    };

    const waitForCaretMove = async (
      prev: { left: number; top: number } | null,
      timeoutMs: number,
    ): Promise<boolean> => {
      if (!prev) return true;
      const start = performance.now();
      while (performance.now() - start < timeoutMs) {
        const next = readOverlayPos();
        if (
          next &&
          (Math.abs(next.left - prev.left) > 0.4 ||
            Math.abs(next.top - prev.top) > 0.4)
        ) {
          return true;
        }
        await sleep(8);
      }
      return false;
    };

    const alpha = "asdlkjqwepoiuzmxncbv";
    let prevPos = readOverlayPos();
    for (let i = 0; i < 260; i++) {
      lastInputAt = performance.now();
      dispatchBeforeInput(
        textarea,
        "insertText",
        alpha[i % alpha.length] ?? "a",
      );
      if (i > 0 && i % 16 === 0) {
        const movedAfterBurst = await waitForCaretMove(prevPos, 180);
        if (!movedAfterBurst) stagnantMoves += 1;
        prevPos = readOverlayPos();
      }
      if (i > 0 && i % 37 === 0) {
        dispatchBeforeInput(textarea, "insertParagraph", null);
        const movedAfterParagraph = await waitForCaretMove(prevPos, 240);
        if (!movedAfterParagraph) stagnantMoves += 1;
        prevPos = readOverlayPos();
      }
      await sleep(4);
    }

    await sleep(1200);
    monitoring = false;
    await monitorTask;
    const overlay = editorContainer.querySelector(
      'div[data-docview-caret-overlay="1"]',
    ) as HTMLDivElement | null;
    const finalTop =
      overlay && overlay.style.display !== "none"
        ? Number.parseFloat(overlay.style.top || "NaN")
        : Number.NaN;
    const ySpan =
      sampleTops.length > 1
        ? Math.max(...sampleTops) - Math.min(...sampleTops)
        : 0;
    const docEdit = ed.getDocEdit() as { viewModel?: () => unknown } | null;
    const model = docEdit?.viewModel?.() as
      | { body?: Array<{ type?: string; runs?: Array<{ text?: string }> }> }
      | undefined;
    const text = (model?.body ?? [])
      .filter((item) => item.type === "paragraph")
      .map((item) => (item.runs ?? []).map((r) => r.text ?? "").join(""))
      .join("\n");
    const textHead = text.slice(0, 48);
    const textTail = text.slice(-48);

    statusEl.textContent = "autorun complete";
    if (stagnantMoves > 4) anomalyCount += stagnantMoves;
    publish(
      "complete",
      anomalyCount,
      ySpan,
      finalTop,
      stagnantMoves,
      textHead,
      textTail,
    );
  } catch (err) {
    statusEl.textContent = `autorun error: ${(err as Error).message}`;
    publish("failed:exception", -1, -1, -1, -1, "", "");
  }
}

function dispatchKey(
  target: HTMLElement,
  key: string,
  opts: {
    shift?: boolean;
    meta?: boolean;
    ctrl?: boolean;
    alt?: boolean;
  } = {},
): boolean {
  return target.dispatchEvent(
    new KeyboardEvent("keydown", {
      bubbles: true,
      cancelable: true,
      key,
      shiftKey: opts.shift ?? false,
      metaKey: opts.meta ?? false,
      ctrlKey: opts.ctrl ?? false,
      altKey: opts.alt ?? false,
    }),
  );
}

function makeClipboardEvent(
  type: "copy" | "cut" | "paste",
  text: string,
  html = "",
): { event: ClipboardEvent; data: DataTransfer } {
  const data = new DataTransfer();
  if (text) data.setData("text/plain", text);
  if (html) data.setData("text/html", html);
  let event: ClipboardEvent;
  try {
    event = new ClipboardEvent(type, {
      bubbles: true,
      cancelable: true,
      clipboardData: data,
    });
  } catch {
    const fallback = new Event(type, {
      bubbles: true,
      cancelable: true,
    }) as ClipboardEvent;
    Object.defineProperty(fallback, "clipboardData", { value: data });
    event = fallback;
  }
  return { event, data };
}

function getDocPlainText(ed: DocEditorController): string {
  const model = getDocViewModel(ed);
  return (model?.body ?? [])
    .filter((item) => item.type === "paragraph")
    .map((item) => (item.runs ?? []).map((r) => r.text ?? "").join(""))
    .join("\n");
}

function getDocViewModel(ed: DocEditorController): {
  body?: Array<{
    type?: string;
    runs?: Array<{
      text?: string;
      inlineImage?: {
        imageIndex?: number;
        widthPt?: number;
        heightPt?: number;
      };
    }>;
    rows?: Array<{ cells?: Array<{ text?: string }> }>;
  }>;
  images?: Array<{ contentType?: string }>;
} | null {
  const docEdit = ed.getDocEdit() as { viewModel?: () => unknown } | null;
  return (
    (docEdit?.viewModel?.() as {
      body?: Array<{
        type?: string;
        runs?: Array<{
          text?: string;
          inlineImage?: {
            imageIndex?: number;
            widthPt?: number;
            heightPt?: number;
          };
        }>;
        rows?: Array<{ cells?: Array<{ text?: string }> }>;
      }>;
      images?: Array<{ contentType?: string }>;
    }) ?? null
  );
}

function getCollabRichContentSummary(ed: DocEditorController): {
  tableCellA1: string;
  tableCellB1: string;
  imageCount: number;
  inlineImageRunCount: number;
} {
  const model = getDocViewModel(ed);
  const firstTable = model?.body?.find((item) => item.type === "table");
  const inlineImageRunCount =
    model?.body
      ?.filter((item) => item.type === "paragraph")
      .flatMap((item) => item.runs ?? [])
      .filter((run) => run.inlineImage != null).length ?? 0;
  return {
    tableCellA1: firstTable?.rows?.[0]?.cells?.[0]?.text ?? "",
    tableCellB1: firstTable?.rows?.[0]?.cells?.[1]?.text ?? "",
    imageCount: model?.images?.length ?? 0,
    inlineImageRunCount,
  };
}

function trimBoundaryNewlines(text: string): string {
  return text.replace(/^\n+/, "").replace(/\n+$/, "");
}

interface CanvasHarnessContext {
  canvas?: HTMLCanvasElement;
  input: HTMLElement;
  clickX: number;
  clickY: number;
}

function resolveCanvasHarnessContext(): CanvasHarnessContext | null {
  const canvas = editorContainer.querySelector("canvas");
  const textarea = editorContainer.querySelector(
    'textarea[data-docview-canvas-input="1"]',
  );
  if (
    canvas instanceof HTMLCanvasElement &&
    textarea instanceof HTMLTextAreaElement
  ) {
    const rect = canvas.getBoundingClientRect();
    return {
      canvas,
      input: textarea,
      clickX: rect.left + Math.max(24, rect.width * 0.2),
      clickY: rect.top + Math.max(32, rect.height * 0.14),
    };
  }
  const surface = editorContainer.querySelector('[contenteditable="true"]');
  if (!(surface instanceof HTMLElement)) return null;
  const rect = surface.getBoundingClientRect();
  return {
    input: surface,
    clickX: rect.left + Math.max(24, rect.width * 0.2),
    clickY:
      rect.top + Math.max(32, Math.min(rect.height - 24, rect.height * 0.14)),
  };
}

function dispatchCanvasClick(
  canvas: HTMLCanvasElement,
  clientX: number,
  clientY: number,
): void {
  canvas.dispatchEvent(
    createMouseEvent("mousedown", {
      button: 0,
      buttons: 1,
      clientX,
      clientY,
    }),
  );
  canvas.dispatchEvent(
    createMouseEvent("mouseup", {
      buttons: 0,
      button: 0,
      clientX,
      clientY,
    }),
  );
}

async function focusCanvasInput(
  ctx: CanvasHarnessContext,
  clientX: number = ctx.clickX,
  clientY: number = ctx.clickY,
): Promise<void> {
  if (ctx.canvas) {
    dispatchCanvasClick(ctx.canvas, clientX, clientY);
  } else {
    ctx.input.dispatchEvent(
      createMouseEvent("mousedown", {
        button: 0,
        buttons: 1,
        clientX,
        clientY,
      }),
    );
    ctx.input.dispatchEvent(
      createMouseEvent("mouseup", {
        button: 0,
        buttons: 0,
        clientX,
        clientY,
      }),
    );
  }
  ctx.input.focus();
  await sleep(40);
}

async function dispatchInsertText(
  target: HTMLElement,
  text: string,
  delayMs: number,
): Promise<void> {
  dispatchBeforeInput(target, "insertText", text);
  await sleep(delayMs);
}

async function seedParagraphLines(
  ctx: CanvasHarnessContext,
  lines: string[],
  delayMs: number = 8,
): Promise<void> {
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i] ?? "";
    await dispatchInsertText(ctx.input, line, delayMs);
    if (i < lines.length - 1) {
      dispatchBeforeInput(ctx.input, "insertParagraph", null);
      await sleep(delayMs + 4);
    }
  }
}

async function dragSelectionOnCanvas(
  canvas: HTMLCanvasElement,
  start: { x: number; y: number },
  end: { x: number; y: number },
  stepCount: number = 14,
): Promise<void> {
  canvas.dispatchEvent(
    createMouseEvent("mousedown", {
      button: 0,
      buttons: 1,
      clientX: start.x,
      clientY: start.y,
    }),
  );
  await sleep(12);
  for (let i = 1; i <= stepCount; i++) {
    const t = i / stepCount;
    window.dispatchEvent(
      createMouseEvent("mousemove", {
        button: 0,
        buttons: 1,
        clientX: start.x + (end.x - start.x) * t,
        clientY: start.y + (end.y - start.y) * t,
      }),
    );
    await sleep(8);
  }
  window.dispatchEvent(
    createMouseEvent("mouseup", {
      button: 0,
      buttons: 0,
      clientX: end.x,
      clientY: end.y,
    }),
  );
  await sleep(36);
}

function createMouseEvent(
  type: "mousedown" | "mousemove" | "mouseup",
  opts: {
    button: number;
    buttons: number;
    clientX: number;
    clientY: number;
  },
): MouseEvent {
  const event = new MouseEvent(type, {
    bubbles: true,
    cancelable: true,
    button: opts.button,
    buttons: opts.buttons,
    clientX: opts.clientX,
    clientY: opts.clientY,
  });
  // Headless Chrome may report synthetic mousemove as buttons=0 even when
  // initialized with buttons=1; force expected fields for drag simulation.
  if (event.buttons !== opts.buttons) {
    Object.defineProperty(event, "buttons", { value: opts.buttons });
  }
  if (event.clientX !== opts.clientX) {
    Object.defineProperty(event, "clientX", { value: opts.clientX });
  }
  if (event.clientY !== opts.clientY) {
    Object.defineProperty(event, "clientY", { value: opts.clientY });
  }
  return event;
}

function pressToolbarButton(button: HTMLButtonElement): void {
  button.dispatchEvent(
    new MouseEvent("mousedown", {
      bubbles: true,
      cancelable: true,
      button: 0,
      buttons: 1,
    }),
  );
}

function assignFileInputFile(input: HTMLInputElement, file: File): void {
  const transfer = new DataTransfer();
  transfer.items.add(file);
  Object.defineProperty(input, "files", {
    configurable: true,
    value: transfer.files,
  });
}

async function makeAutorunImageFile(): Promise<File> {
  const canvas = document.createElement("canvas");
  canvas.width = 32;
  canvas.height = 24;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    throw new Error("autorun image canvas context unavailable");
  }
  ctx.fillStyle = "#f3d248";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#0b3d91";
  ctx.fillRect(4, 4, 24, 8);
  ctx.fillStyle = "#d94841";
  ctx.fillRect(10, 14, 14, 6);
  const blob = await new Promise<Blob | null>((resolve) => {
    canvas.toBlob(resolve, "image/png");
  });
  if (!blob) {
    throw new Error("autorun image export failed");
  }
  return new File([blob], "autorun-inline.png", { type: "image/png" });
}

async function runAutoCanvasOpsScenario(): Promise<void> {
  if (!autoRunOps) return;

  interface OpsSummary {
    status: string;
    textAfterEdits: string;
    undoText: string;
    selectDeleteText: string;
    multilineDeleteText: string;
    shiftDeleteText: string;
    crossParagraphHtmlSelectionText: string;
    dragCopiedLen: string;
    dragBeforeLen: string;
    dragAfterLen: string;
    dragCopiedText: string;
    dragAfterText: string;
    reverseDragCopiedText: string;
    reverseDragAfterText: string;
    tableInsertOk: string;
    tableStartCell: string;
    tableAfterTabCell: string;
    tableAfterRowInsert: string;
    tableAfterColInsert: string;
    tableFinalShape: string;
    tableCellA1: string;
    tableCellB1: string;
    tableCellCopyText: string;
    tableCellCutText: string;
    tableCellDeleteText: string;
    tableCellPasteText: string;
    imageInsertOk: string;
    imageCount: string;
    inlineImageRunCount: string;
    imageContentType: string;
    imageDomCount: string;
    copiedText: string;
    cutText: string;
    richCopyHtmlOk: string;
    richCutHtmlOk: string;
    richPasteOk: string;
    richPasteListOk: string;
    backwardSelectionOk: string;
    notes: string;
  }

  const emptySummary = (): OpsSummary => ({
    status: "failed:unknown",
    textAfterEdits: "",
    undoText: "",
    selectDeleteText: "",
    multilineDeleteText: "",
    shiftDeleteText: "",
    crossParagraphHtmlSelectionText: "",
    dragCopiedLen: "",
    dragBeforeLen: "",
    dragAfterLen: "",
    dragCopiedText: "",
    dragAfterText: "",
    reverseDragCopiedText: "",
    reverseDragAfterText: "",
    tableInsertOk: "",
    tableStartCell: "",
    tableAfterTabCell: "",
    tableAfterRowInsert: "",
    tableAfterColInsert: "",
    tableFinalShape: "",
    tableCellA1: "",
    tableCellB1: "",
    tableCellCopyText: "",
    tableCellCutText: "",
    tableCellDeleteText: "",
    tableCellPasteText: "",
    imageInsertOk: "",
    imageCount: "",
    inlineImageRunCount: "",
    imageContentType: "",
    imageDomCount: "",
    copiedText: "",
    cutText: "",
    richCopyHtmlOk: "",
    richCutHtmlOk: "",
    richPasteOk: "",
    richPasteListOk: "",
    backwardSelectionOk: "",
    notes: "",
  });

  const publish = (summary: OpsSummary): void => {
    document.body.dataset.editorOpsDone = "1";
    document.body.dataset.editorOpsStatus = summary.status;
    document.body.dataset.editorOpsTextAfterEdits = summary.textAfterEdits;
    document.body.dataset.editorOpsUndoText = summary.undoText;
    document.body.dataset.editorOpsSelectDeleteText = summary.selectDeleteText;
    document.body.dataset.editorOpsMultilineDeleteText =
      summary.multilineDeleteText;
    document.body.dataset.editorOpsShiftDeleteText = summary.shiftDeleteText;
    document.body.dataset.editorOpsCrossParagraphHtmlSelectionText =
      summary.crossParagraphHtmlSelectionText;
    document.body.dataset.editorOpsDragCopiedLen = summary.dragCopiedLen;
    document.body.dataset.editorOpsDragBeforeLen = summary.dragBeforeLen;
    document.body.dataset.editorOpsDragAfterLen = summary.dragAfterLen;
    document.body.dataset.editorOpsDragCopiedText = summary.dragCopiedText;
    document.body.dataset.editorOpsDragAfterText = summary.dragAfterText;
    document.body.dataset.editorOpsReverseDragCopiedText =
      summary.reverseDragCopiedText;
    document.body.dataset.editorOpsReverseDragAfterText =
      summary.reverseDragAfterText;
    document.body.dataset.editorOpsTableInsertOk = summary.tableInsertOk;
    document.body.dataset.editorOpsTableStartCell = summary.tableStartCell;
    document.body.dataset.editorOpsTableAfterTabCell =
      summary.tableAfterTabCell;
    document.body.dataset.editorOpsTableAfterRowInsert =
      summary.tableAfterRowInsert;
    document.body.dataset.editorOpsTableAfterColInsert =
      summary.tableAfterColInsert;
    document.body.dataset.editorOpsTableFinalShape = summary.tableFinalShape;
    document.body.dataset.editorOpsTableCellA1 = summary.tableCellA1;
    document.body.dataset.editorOpsTableCellB1 = summary.tableCellB1;
    document.body.dataset.editorOpsTableCellCopyText =
      summary.tableCellCopyText;
    document.body.dataset.editorOpsTableCellCutText = summary.tableCellCutText;
    document.body.dataset.editorOpsTableCellDeleteText =
      summary.tableCellDeleteText;
    document.body.dataset.editorOpsTableCellPasteText =
      summary.tableCellPasteText;
    document.body.dataset.editorOpsImageInsertOk = summary.imageInsertOk;
    document.body.dataset.editorOpsImageCount = summary.imageCount;
    document.body.dataset.editorOpsInlineImageRunCount =
      summary.inlineImageRunCount;
    document.body.dataset.editorOpsImageContentType = summary.imageContentType;
    document.body.dataset.editorOpsImageDomCount = summary.imageDomCount;
    document.body.dataset.editorOpsCopiedText = summary.copiedText;
    document.body.dataset.editorOpsCutText = summary.cutText;
    document.body.dataset.editorOpsRichCopyHtmlOk = summary.richCopyHtmlOk;
    document.body.dataset.editorOpsRichCutHtmlOk = summary.richCutHtmlOk;
    document.body.dataset.editorOpsRichPasteOk = summary.richPasteOk;
    document.body.dataset.editorOpsRichPasteListOk = summary.richPasteListOk;
    document.body.dataset.editorOpsBackwardSelectionOk =
      summary.backwardSelectionOk;
    document.body.dataset.editorOpsNotes = summary.notes;
  };

  const assertRichPaste80_20 = (
    ed: DocEditorController,
    summary: OpsSummary,
    label: string,
  ): void => {
    const model = getDocViewModel(ed) as {
      body?: Array<{
        type?: string;
        headingLevel?: number;
        alignment?: string;
        numbering?: { format?: string; numId?: number };
        runs?: Array<{
          text?: string;
          bold?: boolean;
          underline?: boolean;
          fontFamily?: string;
          fontSizePt?: number;
          color?: string;
          highlight?: string;
          hyperlink?: string;
        }>;
      }>;
    } | null;
    const body = model?.body ?? [];
    const first = body[0];
    const second = body[1];
    const third = body[2];
    const fourth = body[3];
    const firstText = (first?.runs ?? []).map((run) => run.text ?? "").join("");
    const secondText = (second?.runs ?? [])
      .map((run) => run.text ?? "")
      .join("");
    const thirdText = (third?.runs ?? []).map((run) => run.text ?? "").join("");
    const fourthText = (fourth?.runs ?? [])
      .map((run) => run.text ?? "")
      .join("");
    const styledRun = (first?.runs ?? []).find((run) => run.text === "Styled");
    const linkedRun = (first?.runs ?? []).find((run) => run.text === "Link");
    if (
      firstText !== "Styled Link" ||
      styledRun?.bold !== true ||
      styledRun.fontFamily !== "Noto Serif" ||
      styledRun.fontSizePt !== 18 ||
      styledRun.color !== "3366FF" ||
      styledRun.highlight !== "yellow" ||
      linkedRun?.underline !== true ||
      linkedRun.hyperlink !== "https://example.com" ||
      secondText !== "Heading" ||
      second?.headingLevel !== 2 ||
      second?.alignment !== "center"
    ) {
      throw new Error(
        `${label} rich paste mismatch: body='${JSON.stringify(body.slice(0, 4))}'`,
      );
    }
    if (
      thirdText !== "One" ||
      fourthText !== "Two" ||
      third?.numbering?.format !== "bullet" ||
      fourth?.numbering?.format !== "bullet" ||
      third?.numbering?.numId !== fourth?.numbering?.numId ||
      (fourth?.runs ?? []).find((run) => run.text === "Two")?.bold !== true
    ) {
      throw new Error(
        `${label} rich paste list mismatch: body='${JSON.stringify(body.slice(0, 4))}'`,
      );
    }
    summary.richPasteOk = "1";
    summary.richPasteListOk = "1";
  };

  const runTableAndImageChecks = async (
    ed: DocEditorController,
    ctx: CanvasHarnessContext,
    summary: OpsSummary,
  ): Promise<void> => {
    ed.loadBlank();
    await sleep(80);
    pressToolbarButton(fmtTable);
    await sleep(80);
    const tableCellEditor = editorContainer.querySelector(
      'textarea[data-docview-table-cell-editor="1"]',
    ) as HTMLTextAreaElement | null;
    const startCell = ed.getActiveTableCell();
    summary.tableInsertOk = startCell ? "1" : "";
    summary.tableStartCell = startCell
      ? `${startCell.row}:${startCell.col}`
      : "";
    if (summary.tableStartCell !== "0:0") {
      throw new Error(
        `table insert did not focus first cell: '${summary.tableStartCell}'`,
      );
    }
    if (
      !(tableCellEditor instanceof HTMLTextAreaElement) ||
      tableCellEditor.style.display === "none"
    ) {
      throw new Error("table cell editor did not become visible");
    }

    const pasteA1 = makeClipboardEvent("paste", "Alpha");
    tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
    tableCellEditor.dispatchEvent(pasteA1.event);
    await sleep(40);
    summary.tableCellPasteText = ed.getActiveTableCell()?.text ?? "";
    if (summary.tableCellPasteText !== "Alpha") {
      throw new Error(
        `table cell paste mismatch: '${summary.tableCellPasteText}'`,
      );
    }

    tableCellEditor.setSelectionRange(1, 4);
    const copyA1 = makeClipboardEvent("copy", "");
    tableCellEditor.dispatchEvent(copyA1.event);
    summary.tableCellCopyText = copyA1.data.getData("text/plain");
    if (summary.tableCellCopyText !== "lph") {
      throw new Error(
        `table cell copy mismatch: '${summary.tableCellCopyText}'`,
      );
    }

    tableCellEditor.setSelectionRange(1, 4);
    const cutA1 = makeClipboardEvent("cut", "");
    tableCellEditor.dispatchEvent(cutA1.event);
    await sleep(40);
    summary.tableCellCutText = cutA1.data.getData("text/plain");
    if (summary.tableCellCutText !== "lph") {
      throw new Error(`table cell cut mismatch: '${summary.tableCellCutText}'`);
    }
    if (ed.getActiveTableCell()?.text !== "Aa") {
      throw new Error(
        `table cell cut did not persist: '${ed.getActiveTableCell()?.text ?? ""}'`,
      );
    }

    tableCellEditor.setSelectionRange(1, 2);
    tableCellEditor.setRangeText("", 1, 2, "start");
    tableCellEditor.dispatchEvent(
      new InputEvent("input", {
        bubbles: true,
        cancelable: false,
        inputType: "deleteContentBackward",
        data: null,
      }),
    );
    await sleep(40);
    summary.tableCellDeleteText = ed.getActiveTableCell()?.text ?? "";
    if (summary.tableCellDeleteText !== "A") {
      throw new Error(
        `table cell delete mismatch: '${summary.tableCellDeleteText}'`,
      );
    }

    tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
    const restoreA1 = makeClipboardEvent("paste", "A1");
    tableCellEditor.dispatchEvent(restoreA1.event);
    await sleep(40);
    if (ed.getActiveTableCell()?.text !== "A1") {
      throw new Error("table A1 restore failed");
    }
    const firstCellRect = tableCellEditor.getBoundingClientRect();
    const firstCellClick = {
      x: firstCellRect.left + Math.max(8, firstCellRect.width * 0.35),
      y: firstCellRect.top + Math.max(8, firstCellRect.height * 0.5),
    };
    if (!ed.moveActiveTableCell(0, 1)) {
      throw new Error("table move to B1 failed");
    }
    await sleep(60);
    const afterTabCell = ed.getActiveTableCell();
    summary.tableAfterTabCell = afterTabCell
      ? `${afterTabCell.row}:${afterTabCell.col}`
      : "";
    if (summary.tableAfterTabCell !== "0:1") {
      throw new Error(
        `table Tab navigation mismatch: '${summary.tableAfterTabCell}'`,
      );
    }
    if (!ed.setActiveTableCellText("B1")) {
      throw new Error("table B1 commit failed");
    }
    pressToolbarButton(fmtTableRowAdd);
    await sleep(60);
    const afterRowInsert = ed.getActiveTableCell();
    summary.tableAfterRowInsert = afterRowInsert
      ? `${afterRowInsert.row}:${afterRowInsert.col}:${afterRowInsert.rowCount}:${afterRowInsert.colCount}`
      : "";
    if (summary.tableAfterRowInsert !== "1:1:4:3") {
      throw new Error(
        `table row insert mismatch: '${summary.tableAfterRowInsert}'`,
      );
    }
    if (document.activeElement !== tableCellEditor) {
      throw new Error("table row insert should keep cell editor focused");
    }
    if (!ed.setActiveTableCellText("B2")) {
      throw new Error("table B2 commit failed");
    }
    pressToolbarButton(fmtTableColAdd);
    await sleep(60);
    const afterColInsert = ed.getActiveTableCell();
    summary.tableAfterColInsert = afterColInsert
      ? `${afterColInsert.row}:${afterColInsert.col}:${afterColInsert.rowCount}:${afterColInsert.colCount}`
      : "";
    if (summary.tableAfterColInsert !== "1:2:4:4") {
      throw new Error(
        `table column insert mismatch: '${summary.tableAfterColInsert}'`,
      );
    }
    if (document.activeElement !== tableCellEditor) {
      throw new Error("table column insert should keep cell editor focused");
    }
    const insertedColumnRect = tableCellEditor.getBoundingClientRect();
    const insertedColumnClick = {
      x: insertedColumnRect.left + Math.max(8, insertedColumnRect.width * 0.35),
      y: insertedColumnRect.top + Math.max(8, insertedColumnRect.height * 0.5),
    };
    if (!ed.setActiveTableCellText("C2")) {
      throw new Error("table C2 commit failed");
    }
    ed.clearActiveTableCell();
    await sleep(80);
    if (ctx.canvas && ctx.canvas instanceof HTMLCanvasElement) {
      const liveCanvas = editorContainer.querySelector("canvas");
      if (!(liveCanvas instanceof HTMLCanvasElement)) {
        throw new Error("canvas table reopen target missing");
      }
      dispatchCanvasClick(liveCanvas, firstCellClick.x, firstCellClick.y);
      await sleep(60);
      const reopenedFirstCell = ed.getActiveTableCell();
      if (reopenedFirstCell?.row !== 0 || reopenedFirstCell?.col !== 0) {
        throw new Error(
          `canvas table cell reopen mismatch: '${reopenedFirstCell ? `${reopenedFirstCell.row}:${reopenedFirstCell.col}` : ""}'`,
        );
      }
      if (document.activeElement !== tableCellEditor) {
        throw new Error("canvas table cell reopen should focus cell editor");
      }
      tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
      const reopenFirstCellPaste = makeClipboardEvent("paste", "A1x");
      tableCellEditor.dispatchEvent(reopenFirstCellPaste.event);
      await sleep(40);
      if (ed.getActiveTableCell()?.text !== "A1x") {
        throw new Error(
          `canvas table cell reopen edit mismatch: '${ed.getActiveTableCell()?.text ?? ""}'`,
        );
      }
      tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
      const restoreFirstCell = makeClipboardEvent("paste", "A1");
      tableCellEditor.dispatchEvent(restoreFirstCell.event);
      await sleep(40);
      ed.clearActiveTableCell();
      await sleep(80);
      const liveCanvasForInsertedColumn =
        editorContainer.querySelector("canvas");
      if (!(liveCanvasForInsertedColumn instanceof HTMLCanvasElement)) {
        throw new Error("canvas inserted column reopen target missing");
      }
      dispatchCanvasClick(
        liveCanvasForInsertedColumn,
        insertedColumnClick.x,
        insertedColumnClick.y,
      );
      await sleep(60);
      const reopenedInsertedColumn = ed.getActiveTableCell();
      if (
        reopenedInsertedColumn?.row !== 1 ||
        reopenedInsertedColumn?.col !== 2
      ) {
        throw new Error(
          `canvas inserted column reopen mismatch: '${reopenedInsertedColumn ? `${reopenedInsertedColumn.row}:${reopenedInsertedColumn.col}` : ""}'`,
        );
      }
      if (document.activeElement !== tableCellEditor) {
        throw new Error(
          "canvas inserted column reopen should focus cell editor",
        );
      }
      tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
      const reopenInsertedColumnPaste = makeClipboardEvent("paste", "C2x");
      tableCellEditor.dispatchEvent(reopenInsertedColumnPaste.event);
      await sleep(40);
      if (ed.getActiveTableCell()?.text !== "C2x") {
        throw new Error(
          `canvas inserted column reopen edit mismatch: '${ed.getActiveTableCell()?.text ?? ""}'`,
        );
      }
      tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
      const restoreInsertedColumn = makeClipboardEvent("paste", "C2");
      tableCellEditor.dispatchEvent(restoreInsertedColumn.event);
      await sleep(40);
    }
    if (!ctx.canvas) {
      const existingCell = editorContainer.querySelector(
        '[data-docview-table-cell="1"][data-docview-table-row="0"][data-docview-table-col="0"]',
      ) as HTMLElement | null;
      if (!existingCell) {
        throw new Error("html table cell reopen target missing");
      }
      existingCell.dispatchEvent(
        new MouseEvent("mousedown", {
          bubbles: true,
          cancelable: true,
          clientX: existingCell.getBoundingClientRect().left + 8,
          clientY: existingCell.getBoundingClientRect().top + 8,
          button: 0,
          buttons: 1,
          detail: 1,
        }),
      );
      existingCell.dispatchEvent(
        new MouseEvent("mouseup", {
          bubbles: true,
          cancelable: true,
          clientX: existingCell.getBoundingClientRect().left + 8,
          clientY: existingCell.getBoundingClientRect().top + 8,
          button: 0,
          buttons: 0,
          detail: 1,
        }),
      );
      await sleep(60);
      const reopenedCell = ed.getActiveTableCell();
      if (reopenedCell?.row !== 0 || reopenedCell?.col !== 0) {
        throw new Error(
          `html table cell reopen mismatch: '${reopenedCell ? `${reopenedCell.row}:${reopenedCell.col}` : ""}'`,
        );
      }
      if (document.activeElement !== tableCellEditor) {
        throw new Error("html table cell reopen should focus cell editor");
      }
      tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
      const reopenPaste = makeClipboardEvent("paste", "A1x");
      tableCellEditor.dispatchEvent(reopenPaste.event);
      await sleep(40);
      if (ed.getActiveTableCell()?.text !== "A1x") {
        throw new Error(
          `html table cell reopen edit mismatch: '${ed.getActiveTableCell()?.text ?? ""}'`,
        );
      }
      tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
      const restoreAfterReopen = makeClipboardEvent("paste", "A1");
      tableCellEditor.dispatchEvent(restoreAfterReopen.event);
      await sleep(40);
      ed.clearActiveTableCell();
      await sleep(80);
      const insertedColumnCell = editorContainer.querySelector(
        '[data-docview-table-cell="1"][data-docview-table-row="1"][data-docview-table-col="2"]',
      ) as HTMLElement | null;
      if (!insertedColumnCell) {
        throw new Error("html inserted column cell target missing");
      }
      insertedColumnCell.dispatchEvent(
        new MouseEvent("mousedown", {
          bubbles: true,
          cancelable: true,
          clientX: insertedColumnCell.getBoundingClientRect().left + 8,
          clientY: insertedColumnCell.getBoundingClientRect().top + 8,
          button: 0,
          buttons: 1,
          detail: 1,
        }),
      );
      insertedColumnCell.dispatchEvent(
        new MouseEvent("mouseup", {
          bubbles: true,
          cancelable: true,
          clientX: insertedColumnCell.getBoundingClientRect().left + 8,
          clientY: insertedColumnCell.getBoundingClientRect().top + 8,
          button: 0,
          buttons: 0,
          detail: 1,
        }),
      );
      await sleep(60);
      const reopenedInsertedColumn = ed.getActiveTableCell();
      if (
        reopenedInsertedColumn?.row !== 1 ||
        reopenedInsertedColumn?.col !== 2
      ) {
        throw new Error(
          `html inserted column reopen mismatch: '${reopenedInsertedColumn ? `${reopenedInsertedColumn.row}:${reopenedInsertedColumn.col}` : ""}'`,
        );
      }
      if (document.activeElement !== tableCellEditor) {
        throw new Error("html inserted column reopen should focus cell editor");
      }
      tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
      const reopenInsertedColumnPaste = makeClipboardEvent("paste", "C2x");
      tableCellEditor.dispatchEvent(reopenInsertedColumnPaste.event);
      await sleep(40);
      if (ed.getActiveTableCell()?.text !== "C2x") {
        throw new Error(
          `html inserted column reopen edit mismatch: '${ed.getActiveTableCell()?.text ?? ""}'`,
        );
      }
      tableCellEditor.setSelectionRange(0, tableCellEditor.value.length);
      const restoreInsertedColumn = makeClipboardEvent("paste", "C2");
      tableCellEditor.dispatchEvent(restoreInsertedColumn.event);
      await sleep(40);
    }
    if (!ed.removeTableColumn()) {
      throw new Error("table column remove failed");
    }
    if (!ed.removeTableRow()) {
      throw new Error("table row remove failed");
    }
    await sleep(60);
    const finalCell = ed.getActiveTableCell();
    summary.tableFinalShape = finalCell
      ? `${finalCell.rowCount}x${finalCell.colCount}`
      : "";
    if (summary.tableFinalShape !== "3x3") {
      throw new Error(
        `table final shape mismatch: '${summary.tableFinalShape}'`,
      );
    }
    ed.clearActiveTableCell();
    await sleep(80);
    const tableModel = getDocViewModel(ed)?.body?.find(
      (item) => item.type === "table",
    );
    summary.tableCellA1 = tableModel?.rows?.[0]?.cells?.[0]?.text ?? "";
    summary.tableCellB1 = tableModel?.rows?.[0]?.cells?.[1]?.text ?? "";
    if (summary.tableCellA1 !== "A1" || summary.tableCellB1 !== "B1") {
      throw new Error(
        `table cell commit mismatch: A1='${summary.tableCellA1}' B1='${summary.tableCellB1}'`,
      );
    }

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    const imageFile = await makeAutorunImageFile();
    pressToolbarButton(fmtImage);
    await sleep(20);
    assignFileInputFile(imageInput, imageFile);
    imageInput.dispatchEvent(new Event("change", { bubbles: true }));
    await waitFor(() => {
      const current = getDocViewModel(ed);
      const imageCount = current?.images?.length ?? 0;
      const runCount =
        current?.body
          ?.filter((item) => item.type === "paragraph")
          .flatMap((item) => item.runs ?? [])
          .filter((run) => run.inlineImage != null).length ?? 0;
      return imageCount >= 1 && runCount >= 1;
    }, 1200);
    const imageModel = getDocViewModel(ed);
    const inlineImageRunCount =
      imageModel?.body
        ?.filter((item) => item.type === "paragraph")
        .flatMap((item) => item.runs ?? [])
        .filter((run) => run.inlineImage != null).length ?? 0;
    summary.imageCount = String(imageModel?.images?.length ?? 0);
    summary.inlineImageRunCount = String(inlineImageRunCount);
    summary.imageContentType = imageModel?.images?.[0]?.contentType ?? "";
    summary.imageDomCount = String(
      editorContainer.querySelectorAll("img[data-docview-inline-image='1']")
        .length,
    );
    summary.imageInsertOk =
      summary.imageCount === "1" && summary.inlineImageRunCount === "1"
        ? "1"
        : "";
    if (summary.imageInsertOk !== "1") {
      throw new Error(
        `image insert mismatch: images=${summary.imageCount} runs=${summary.inlineImageRunCount}`,
      );
    }
    if (summary.imageContentType !== "image/png") {
      throw new Error(
        `image content type mismatch: '${summary.imageContentType}'`,
      );
    }
    if (!ctx.canvas && summary.imageDomCount !== "1") {
      throw new Error(
        `html inline image DOM mismatch: expected 1, got '${summary.imageDomCount}'`,
      );
    }
  };

  try {
    const ed = await ensureEditor();
    const summary = emptySummary();
    const assertSelectionDeleteExact = (
      beforeText: string,
      selectedText: string,
      afterText: string,
      label: string,
    ): void => {
      if (selectedText.length === 0) {
        throw new Error(`${label}: selected text should not be empty`);
      }
      const firstIndex = beforeText.indexOf(selectedText);
      const lastIndex = beforeText.lastIndexOf(selectedText);
      if (firstIndex < 0 || firstIndex !== lastIndex) {
        throw new Error(
          `${label}: selected text must map to a unique range in source`,
        );
      }
      const expectedAfter =
        beforeText.slice(0, firstIndex) +
        beforeText.slice(firstIndex + selectedText.length);
      if (afterText !== expectedAfter) {
        throw new Error(
          `${label}: deleted content mismatch expected '${expectedAfter}' got '${afterText}'`,
        );
      }
    };
    ed.loadBlank();
    await sleep(120);

    const ctx = resolveCanvasHarnessContext();
    if (!ctx) {
      statusEl.textContent = "ops failed: no input";
      publish({
        ...summary,
        status: "failed:no-input",
        notes: "input-missing",
      });
      return;
    }

    if (!ctx.canvas) {
      const surface = ctx.input;
      await focusCanvasInput(ctx);
      for (const ch of "Rich") {
        dispatchBeforeInput(surface, "insertText", ch);
        await sleep(12);
      }
      restoreSelection(surface, 0, 0, 0, 4);
      await sleep(20);
      ed.setTextStyle({
        bold: true,
        color: "#112233",
        highlight: "yellow",
        hyperlink: "https://example.com",
      });
      await sleep(40);
      restoreSelection(surface, 0, 4, 0, 0);
      await sleep(20);
      const restored = window.getSelection();
      const restoredDocSel = restored
        ? selectionToDocPositions(restored)
        : null;
      if (
        !restoredDocSel ||
        restoredDocSel.anchor.bodyIndex !== 0 ||
        restoredDocSel.anchor.charOffset !== 4 ||
        restoredDocSel.focus.bodyIndex !== 0 ||
        restoredDocSel.focus.charOffset !== 0
      ) {
        throw new Error(
          `backward selection restore failed: ${JSON.stringify(restoredDocSel)}`,
        );
      }
      summary.backwardSelectionOk = "1";

      const copy = makeClipboardEvent("copy", "");
      surface.dispatchEvent(copy.event);
      summary.copiedText = copy.data.getData("text/plain");
      const copyHtml = copy.data.getData("text/html");
      if (summary.copiedText !== "Rich") {
        throw new Error(
          `html copy text mismatch: expected 'Rich', got '${summary.copiedText}'`,
        );
      }
      if (
        !copyHtml.includes('href="https://example.com"') ||
        !copyHtml.includes("font-weight: 700") ||
        !copyHtml.includes("background-color: #ffff00")
      ) {
        throw new Error(`html copy rich mismatch: '${copyHtml.slice(0, 240)}'`);
      }
      summary.richCopyHtmlOk = "1";

      const cut = makeClipboardEvent("cut", "");
      surface.dispatchEvent(cut.event);
      await sleep(40);
      summary.cutText = cut.data.getData("text/plain");
      const cutHtml = cut.data.getData("text/html");
      if (summary.cutText !== "Rich") {
        throw new Error(
          `html cut text mismatch: expected 'Rich', got '${summary.cutText}'`,
        );
      }
      if (
        !cutHtml.includes('href="https://example.com"') ||
        !cutHtml.includes("font-weight: 700") ||
        !cutHtml.includes("background-color: #ffff00")
      ) {
        throw new Error(`html cut rich mismatch: '${cutHtml.slice(0, 240)}'`);
      }
      if (getDocPlainText(ed) !== "") {
        throw new Error(
          `html cut should delete selection, got '${getDocPlainText(ed)}'`,
        );
      }

      ed.loadBlank();
      await sleep(80);
      await focusCanvasInput(ctx);
      const richPaste = makeClipboardEvent(
        "paste",
        "Styled Link\nHeading\nOne\nTwo",
        '<p><span style="font-family: \'Noto Serif\'; font-size: 18pt; color: #3366ff; background-color: yellow"><strong>Styled</strong></span> <a href="https://example.com"><u>Link</u></a></p><h2 style="text-align: center">Heading</h2><ul><li>One</li><li><strong>Two</strong></li></ul>',
      );
      surface.dispatchEvent(richPaste.event);
      await sleep(60);
      assertRichPaste80_20(ed, summary, "html");

      ed.loadBlank();
      await sleep(80);
      await focusCanvasInput(ctx);
      for (const line of ["one", "two"]) {
        await dispatchInsertText(surface, line, 10);
        if (line !== "two") {
          dispatchBeforeInput(surface, "insertParagraph", null);
          await sleep(14);
        }
      }
      dispatchKey(surface, "a", { meta: true, ctrl: true });
      await sleep(20);
      const htmlSelectAllCopy = makeClipboardEvent("copy", "");
      surface.dispatchEvent(htmlSelectAllCopy.event);
      if (htmlSelectAllCopy.data.getData("text/plain") !== "one\ntwo") {
        throw new Error(
          `html select-all mismatch: '${htmlSelectAllCopy.data.getData("text/plain")}'`,
        );
      }
      dispatchKey(surface, "Backspace");
      await sleep(30);
      if (getDocPlainText(ed) !== "") {
        throw new Error(
          `html select-all delete failed: got '${getDocPlainText(ed)}'`,
        );
      }

      ed.loadBlank();
      await sleep(80);
      await focusCanvasInput(ctx);
      await dispatchInsertText(surface, "alpha beta gamma", 10);
      dispatchKey(surface, "End");
      await sleep(16);
      dispatchKey(surface, "Backspace", { alt: true, ctrl: true });
      await sleep(20);
      if (getDocPlainText(ed) !== "alpha beta ") {
        throw new Error(
          `html word-backspace mismatch: got '${getDocPlainText(ed)}'`,
        );
      }

      ed.loadBlank();
      await sleep(80);
      await focusCanvasInput(ctx);
      await dispatchInsertText(surface, "alpha beta gamma", 10);
      restoreSelection(surface, 0, 6, 0, 6);
      await sleep(20);
      dispatchKey(surface, "Delete", { alt: true, ctrl: true });
      await sleep(20);
      if (getDocPlainText(ed) !== "alpha gamma") {
        throw new Error(
          `html word-delete mismatch: got '${getDocPlainText(ed)}'`,
        );
      }

      ed.loadBlank();
      await sleep(80);
      await focusCanvasInput(ctx);
      for (const line of ["alpha", "bravo", "charlie"]) {
        await dispatchInsertText(surface, line, 10);
        if (line !== "charlie") {
          dispatchBeforeInput(surface, "insertParagraph", null);
          await sleep(14);
        }
      }
      restoreSelection(surface, 0, 2, 2, 4);
      await sleep(20);
      const multiParagraphCopy = makeClipboardEvent("copy", "");
      surface.dispatchEvent(multiParagraphCopy.event);
      summary.crossParagraphHtmlSelectionText =
        multiParagraphCopy.data.getData("text/plain");
      if (summary.crossParagraphHtmlSelectionText !== "pha\nbravo\nchar") {
        throw new Error(
          `html cross paragraph selection mismatch: '${summary.crossParagraphHtmlSelectionText}'`,
        );
      }
      const beforeHtmlDelete = getDocPlainText(ed);
      const multiParagraphCut = makeClipboardEvent("cut", "");
      surface.dispatchEvent(multiParagraphCut.event);
      await sleep(40);
      assertSelectionDeleteExact(
        beforeHtmlDelete,
        summary.crossParagraphHtmlSelectionText,
        getDocPlainText(ed),
        "html cross paragraph delete",
      );
      await runTableAndImageChecks(ed, ctx, summary);
      summary.richCutHtmlOk = "1";
      summary.status = "complete";
      summary.notes = "ok";
      publish(summary);
      statusEl.textContent = "ops autorun complete";
      return;
    }

    const assertText = (expected: string, step: string): void => {
      const actual = getDocPlainText(ed);
      if (actual !== expected) {
        throw new Error(`${step}: expected '${expected}', got '${actual}'`);
      }
    };

    await focusCanvasInput(ctx);

    for (const ch of "abcd") {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(14);
    }
    assertText("abcd", "type abcd");

    dispatchKey(ctx.input, "ArrowLeft");
    await sleep(12);
    dispatchKey(ctx.input, "ArrowLeft");
    await sleep(12);
    dispatchBeforeInput(ctx.input, "insertText", "X");
    await sleep(20);
    assertText("abXcd", "arrow-left + insert");

    dispatchKey(ctx.input, "Home");
    await sleep(12);
    dispatchBeforeInput(ctx.input, "insertText", "1");
    await sleep(20);
    assertText("1abXcd", "home + insert");

    dispatchKey(ctx.input, "End");
    await sleep(12);
    dispatchBeforeInput(ctx.input, "insertText", "2");
    await sleep(20);
    assertText("1abXcd2", "end + insert");

    dispatchKey(ctx.input, "ArrowLeft", { shift: true });
    await sleep(12);
    const copy = makeClipboardEvent("copy", "");
    ctx.input.dispatchEvent(copy.event);
    summary.copiedText = copy.data.getData("text/plain");
    assertText("1abXcd2", "copy should not mutate");

    const cut = makeClipboardEvent("cut", "");
    ctx.input.dispatchEvent(cut.event);
    await sleep(30);
    summary.cutText = cut.data.getData("text/plain");
    assertText("1abXcd", "cut should delete selection");
    dispatchBeforeInput(ctx.input, "insertText", "2");
    await sleep(20);
    assertText("1abXcd2", "restore tail after cut");

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    for (const ch of "Rich") {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(12);
    }
    dispatchKey(ctx.input, "a", { meta: true, ctrl: true });
    await sleep(20);
    ed.setTextStyle({
      bold: true,
      color: "#112233",
      highlight: "yellow",
      hyperlink: "https://example.com",
    });
    await sleep(40);
    const richCopy = makeClipboardEvent("copy", "");
    ctx.input.dispatchEvent(richCopy.event);
    const richCopyHtml = richCopy.data.getData("text/html");
    if (
      !richCopyHtml.includes('href="https://example.com"') ||
      !richCopyHtml.includes("font-weight: 700") ||
      !richCopyHtml.includes("background-color: #ffff00")
    ) {
      throw new Error(
        `rich copy html mismatch: '${richCopyHtml.slice(0, 240)}'`,
      );
    }
    summary.richCopyHtmlOk = "1";

    const richCut = makeClipboardEvent("cut", "");
    ctx.input.dispatchEvent(richCut.event);
    await sleep(30);
    const richCutHtml = richCut.data.getData("text/html");
    if (
      !richCutHtml.includes('href="https://example.com"') ||
      !richCutHtml.includes("font-weight: 700") ||
      !richCutHtml.includes("background-color: #ffff00")
    ) {
      throw new Error(`rich cut html mismatch: '${richCutHtml.slice(0, 240)}'`);
    }
    if (getDocPlainText(ed) !== "") {
      throw new Error(
        `rich cut should delete selection, got '${getDocPlainText(ed)}'`,
      );
    }
    summary.richCutHtmlOk = "1";

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    for (const ch of "1abXcd2") {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(12);
    }

    dispatchKey(ctx.input, "ArrowLeft", { shift: true });
    await sleep(12);
    const paste = makeClipboardEvent("paste", "Q");
    ctx.input.dispatchEvent(paste.event);
    await sleep(30);
    assertText("1abXcdQ", "paste replace selection");

    dispatchBeforeInput(ctx.input, "deleteContentBackward", null);
    await sleep(20);
    assertText("1abXcd", "backspace");

    dispatchKey(ctx.input, "ArrowLeft");
    await sleep(12);
    dispatchBeforeInput(ctx.input, "deleteContentForward", null);
    await sleep(20);
    assertText("1abXc", "delete forward");
    summary.textAfterEdits = getDocPlainText(ed);

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    for (const ch of "undo") {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(14);
    }
    assertText("undo", "undo seed text");
    dispatchKey(ctx.input, "z", { meta: true, ctrl: true });
    await sleep(40);
    summary.undoText = getDocPlainText(ed);
    if (summary.undoText === "undo") {
      throw new Error("command-z did not change document");
    }

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    for (const ch of "wipe-me") {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(12);
    }
    dispatchKey(ctx.input, "a", { meta: true, ctrl: true });
    await sleep(20);
    dispatchKey(ctx.input, "Backspace");
    await sleep(30);
    summary.selectDeleteText = getDocPlainText(ed);
    if (summary.selectDeleteText !== "") {
      throw new Error(
        `select-all delete failed: expected empty text, got '${summary.selectDeleteText}'`,
      );
    }

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    await seedParagraphLines(ctx, ["abc", "def", "ghi"], 12);
    assertText("abc\ndef\nghi", "multiline seed");
    dispatchKey(ctx.input, "a", { meta: true, ctrl: true });
    await sleep(20);
    const multiCopy = makeClipboardEvent("copy", "");
    ctx.input.dispatchEvent(multiCopy.event);
    const multiCopiedText = multiCopy.data.getData("text/plain");
    if (multiCopiedText !== "abc\ndef\nghi") {
      throw new Error(
        `multiline select-all mismatch: expected 'abc\\ndef\\nghi', got '${multiCopiedText}'`,
      );
    }
    dispatchKey(ctx.input, "Backspace");
    await sleep(40);
    summary.multilineDeleteText = getDocPlainText(ed);
    if (summary.multilineDeleteText !== "") {
      throw new Error(
        `multiline delete failed: expected empty text, got '${summary.multilineDeleteText}'`,
      );
    }

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    for (const ch of "alpha beta gamma") {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(12);
    }
    dispatchKey(ctx.input, "Home");
    await sleep(12);
    dispatchKey(ctx.input, "ArrowRight", { alt: true, ctrl: true });
    await sleep(16);
    dispatchBeforeInput(ctx.input, "insertText", "X");
    await sleep(20);
    assertText("alpha Xbeta gamma", "word-right + insert");

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    for (const ch of "alpha beta gamma") {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(12);
    }
    dispatchKey(ctx.input, "Backspace", { alt: true, ctrl: true });
    await sleep(20);
    assertText("alpha beta ", "word-backspace");

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    for (const ch of "alpha beta gamma") {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(12);
    }
    dispatchKey(ctx.input, "Home");
    await sleep(12);
    dispatchKey(ctx.input, "ArrowRight", { alt: true, ctrl: true });
    await sleep(16);
    dispatchKey(ctx.input, "Delete", { alt: true, ctrl: true });
    await sleep(20);
    assertText("alpha gamma", "word-delete");

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    await seedParagraphLines(ctx, ["abc", "def", "ghi"], 10);
    assertText("abc\ndef\nghi", "shift seed");
    for (let i = 0; i < 5; i++) {
      dispatchKey(ctx.input, "ArrowLeft", { shift: true });
      await sleep(12);
    }
    const shiftCopy = makeClipboardEvent("copy", "");
    ctx.input.dispatchEvent(shiftCopy.event);
    const shiftCopiedText = shiftCopy.data.getData("text/plain");
    if (shiftCopiedText !== "f\nghi") {
      throw new Error(
        `shift selection mismatch: expected 'f\\nghi', got '${shiftCopiedText}'`,
      );
    }
    dispatchKey(ctx.input, "Backspace");
    await sleep(40);
    summary.shiftDeleteText = getDocPlainText(ed);
    if (summary.shiftDeleteText !== "abc\nde") {
      throw new Error(
        `shift delete failed: expected 'abc\\nde', got '${summary.shiftDeleteText}'`,
      );
    }

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    const richPaste = makeClipboardEvent(
      "paste",
      "Styled Link\nHeading\nOne\nTwo",
      '<p><span style="font-family: \'Noto Serif\'; font-size: 18pt; color: #3366ff; background-color: yellow"><strong>Styled</strong></span> <a href="https://example.com"><u>Link</u></a></p><h2 style="text-align: center">Heading</h2><ul><li>One</li><li><strong>Two</strong></li></ul>',
    );
    ctx.input.dispatchEvent(richPaste.event);
    await sleep(60);
    assertRichPaste80_20(ed, summary, "canvas");

    ed.loadBlank();
    await sleep(80);
    await focusCanvasInput(ctx);
    const dragLines = Array.from(
      { length: 60 },
      (_, i) => `L${String(i).padStart(2, "0")}`,
    );
    await seedParagraphLines(ctx, dragLines, 4);
    await sleep(120);
    const beforeDragText = getDocPlainText(ed);
    summary.dragBeforeLen = String(beforeDragText.length);

    if (ctx.canvas) {
      const dragCanvases = Array.from(
        editorContainer.querySelectorAll("canvas"),
      );
      if (dragCanvases.length < 2) {
        throw new Error(
          `drag scenario expected multiple pages, got ${dragCanvases.length}`,
        );
      }
      const firstCanvas = dragCanvases[0] as HTMLCanvasElement;
      const firstRect = firstCanvas.getBoundingClientRect();
      const dragStart = {
        x: firstRect.left + Math.max(30, firstRect.width * 0.18),
        y: firstRect.top + Math.max(34, firstRect.height * 0.15),
      };
      const dragEnd = {
        x: firstRect.left + Math.max(34, firstRect.width * 0.18),
        y: firstRect.top + Math.max(160, firstRect.height * 0.58),
      };
      await dragSelectionOnCanvas(firstCanvas, dragStart, dragEnd);

      const dragCopy = makeClipboardEvent("copy", "");
      ctx.input.dispatchEvent(dragCopy.event);
      const dragCopiedText = dragCopy.data.getData("text/plain");
      summary.dragCopiedText = dragCopiedText;
      summary.dragCopiedLen = String(dragCopiedText.length);
      if (!dragCopiedText.includes("\n") || dragCopiedText.length < 20) {
        throw new Error(
          `drag selection mismatch: expected multiline copied text, got len=${dragCopiedText.length}`,
        );
      }
      dispatchKey(ctx.input, "Backspace");
      await sleep(48);
      const afterDragText = getDocPlainText(ed);
      summary.dragAfterText = afterDragText;
      summary.dragAfterLen = String(afterDragText.length);
      assertSelectionDeleteExact(
        beforeDragText,
        dragCopiedText,
        afterDragText,
        "forward drag delete",
      );

      ed.loadBlank();
      await sleep(80);
      await focusCanvasInput(ctx);
      await seedParagraphLines(ctx, dragLines, 4);
      await sleep(120);
      const beforeReverseDragText = getDocPlainText(ed);
      const reverseDragCanvases = Array.from(
        editorContainer.querySelectorAll("canvas"),
      );
      const secondCanvas = reverseDragCanvases[0] as
        | HTMLCanvasElement
        | undefined;
      if (!secondCanvas) {
        throw new Error("reverse drag scenario missing canvas");
      }
      await dragSelectionOnCanvas(secondCanvas, dragEnd, dragStart);
      const reverseCopy = makeClipboardEvent("copy", "");
      ctx.input.dispatchEvent(reverseCopy.event);
      const reverseCopiedText = reverseCopy.data.getData("text/plain");
      summary.reverseDragCopiedText = reverseCopiedText;
      if (reverseCopiedText !== dragCopiedText) {
        throw new Error(
          `reverse drag selection mismatch: expected '${dragCopiedText}', got '${reverseCopiedText}'`,
        );
      }
      dispatchKey(ctx.input, "Backspace");
      await sleep(48);
      const afterReverseDragText = getDocPlainText(ed);
      summary.reverseDragAfterText = afterReverseDragText;
      assertSelectionDeleteExact(
        beforeReverseDragText,
        reverseCopiedText,
        afterReverseDragText,
        "reverse drag delete",
      );
      if (afterReverseDragText !== afterDragText) {
        throw new Error(
          `reverse drag delete mismatch: expected '${afterDragText}', got '${afterReverseDragText}'`,
        );
      }
    }

    await runTableAndImageChecks(ed, ctx, summary);
    statusEl.textContent = "ops autorun complete";
    summary.status = "complete";
    summary.notes = "ok";
    publish(summary);
  } catch (err) {
    statusEl.textContent = `ops autorun error: ${(err as Error).message}`;
    const summary = emptySummary();
    summary.status = "failed:exception";
    summary.notes = (err as Error).message;
    publish(summary);
  }
}

async function runAutoCopyLinkScenario(): Promise<void> {
  if (!autoRunCopyLink) return;

  const publish = (status: string, copiedLink: string, notes: string): void => {
    document.body.dataset.editorCopyLinkDone = "1";
    document.body.dataset.editorCopyLinkStatus = status;
    document.body.dataset.editorCopyLinkCopiedLink = copiedLink;
    document.body.dataset.editorCopyLinkRoom = roomId ?? "";
    document.body.dataset.editorCopyLinkNotes = notes;
  };

  try {
    if (!roomId) {
      roomId = generateRoomId();
      updateCollabUI();
    }
    let copiedLink = "";
    const clipboard = navigator.clipboard as
      | {
          writeText?: (text: string) => Promise<void>;
        }
      | undefined;
    const originalWriteText = clipboard?.writeText?.bind(clipboard);
    const stubWriteText = async (text: string): Promise<void> => {
      copiedLink = text;
    };
    if (clipboard) {
      clipboard.writeText = stubWriteText;
    } else {
      Object.defineProperty(navigator, "clipboard", {
        value: { writeText: stubWriteText },
        configurable: true,
      });
    }

    btnCollab.click();
    await sleep(120);

    if (originalWriteText && clipboard) {
      clipboard.writeText = originalWriteText;
    }

    const copiedUrl = new URL(copiedLink);
    if (copiedUrl.searchParams.get("room") !== roomId) {
      throw new Error(
        `copied room mismatch: expected '${roomId}' got '${copiedUrl.searchParams.get("room") ?? ""}'`,
      );
    }
    if (wsUrl && copiedUrl.searchParams.get("ws") !== wsUrl) {
      throw new Error(
        `copied ws mismatch: expected '${wsUrl}' got '${copiedUrl.searchParams.get("ws") ?? ""}'`,
      );
    }
    publish("complete", copiedLink, "ok");
  } catch (err) {
    publish("failed:exception", "", describeError(err));
  }
}

async function runAutoCollabSyncScenario(): Promise<void> {
  const runCollabFlow = autoRunCollab || autoRunCollabPresenceExpiry;
  const runPeerFlow = autoRunCollabPeer || autoRunCollabPresenceExpiryPeer;
  if (!runCollabFlow && !runPeerFlow) return;
  if (!roomId) {
    document.body.dataset.editorCollabDone = "1";
    document.body.dataset.editorCollabStatus = "failed:no-room";
    document.body.dataset.editorCollabPresenceCount = "0";
    document.body.dataset.editorCollabPresenceExpired = "0";
    document.body.dataset.editorCollabPresenceExpiredCount = "0";
    document.body.dataset.editorCollabPresenceExpiredCursorCount = "0";
    document.body.dataset.editorCollabPresenceExpiredSelectionCount = "0";
    document.body.dataset.editorCollabNotes = "missing room query parameter";
    return;
  }

  const channel = new BroadcastChannel(`docview-collab-harness:${roomId}`);
  const publish = (
    status: string,
    expected: string,
    peerText: string,
    peerPresenceCount: number,
    peerCursorCount: number,
    peerTableCellA1: string,
    peerTableCellB1: string,
    peerImageCount: number,
    peerInlineImageRunCount: number,
    replaceExpected: string,
    replacePeerText: string,
    replacePeerTableCellA1: string,
    replacePeerTableCellB1: string,
    replacePeerImageCount: number,
    replacePeerInlineImageRunCount: number,
    reconnectExpected: string,
    reconnectPeerText: string,
    reconnectPeerTableCellA1: string,
    reconnectPeerTableCellB1: string,
    reconnectPeerImageCount: number,
    reconnectPeerInlineImageRunCount: number,
    expiredPresence: {
      expired: boolean;
      peerCount: number;
      cursorCount: number;
      selectionCount: number;
    },
    notes: string,
  ): void => {
    document.body.dataset.editorCollabDone = "1";
    document.body.dataset.editorCollabStatus = status;
    document.body.dataset.editorCollabExpected = expected;
    document.body.dataset.editorCollabPeerText = peerText;
    document.body.dataset.editorCollabPresenceCount = String(peerPresenceCount);
    document.body.dataset.editorCollabPresenceCursorCount =
      String(peerCursorCount);
    document.body.dataset.editorCollabPeerTableCellA1 = peerTableCellA1;
    document.body.dataset.editorCollabPeerTableCellB1 = peerTableCellB1;
    document.body.dataset.editorCollabPeerImageCount = String(peerImageCount);
    document.body.dataset.editorCollabPeerInlineImageRunCount = String(
      peerInlineImageRunCount,
    );
    document.body.dataset.editorCollabReplaceExpected = replaceExpected;
    document.body.dataset.editorCollabReplacePeerText = replacePeerText;
    document.body.dataset.editorCollabReplacePeerTableCellA1 =
      replacePeerTableCellA1;
    document.body.dataset.editorCollabReplacePeerTableCellB1 =
      replacePeerTableCellB1;
    document.body.dataset.editorCollabReplacePeerImageCount = String(
      replacePeerImageCount,
    );
    document.body.dataset.editorCollabReplacePeerInlineImageRunCount = String(
      replacePeerInlineImageRunCount,
    );
    document.body.dataset.editorCollabReconnectExpected = reconnectExpected;
    document.body.dataset.editorCollabReconnectPeerText = reconnectPeerText;
    document.body.dataset.editorCollabReconnectPeerTableCellA1 =
      reconnectPeerTableCellA1;
    document.body.dataset.editorCollabReconnectPeerTableCellB1 =
      reconnectPeerTableCellB1;
    document.body.dataset.editorCollabReconnectPeerImageCount = String(
      reconnectPeerImageCount,
    );
    document.body.dataset.editorCollabReconnectPeerInlineImageRunCount = String(
      reconnectPeerInlineImageRunCount,
    );
    document.body.dataset.editorCollabPresenceExpired = expiredPresence.expired
      ? "1"
      : "0";
    document.body.dataset.editorCollabPresenceExpiredCount = String(
      expiredPresence.peerCount,
    );
    document.body.dataset.editorCollabPresenceExpiredCursorCount = String(
      expiredPresence.cursorCount,
    );
    document.body.dataset.editorCollabPresenceExpiredSelectionCount = String(
      expiredPresence.selectionCount,
    );
    document.body.dataset.editorCollabRoom = roomId ?? "";
    document.body.dataset.editorCollabTransport = syncDebug.mode;
    document.body.dataset.editorCollabAwarenessExpire = String(
      syncDebug.awarenessExpire,
    );
    document.body.dataset.editorCollabAwarenessByeRecv = String(
      syncDebug.awarenessByeRecv,
    );
    document.body.dataset.editorCollabStateRequestSend = String(
      syncDebug.stateRequestSend,
    );
    document.body.dataset.editorCollabDivergenceDetected = String(
      syncDebug.divergenceDetected,
    );
    document.body.dataset.editorCollabDivergenceCleared = String(
      syncDebug.divergenceCleared,
    );
    document.body.dataset.editorCollabRepairRequests = String(
      syncDebug.repairRequests,
    );
    document.body.dataset.editorCollabResyncing = syncDebug.resyncing
      ? "1"
      : "0";
    document.body.dataset.editorCollabDivergenceSuspected =
      syncDebug.divergenceSuspected ? "1" : "0";
    document.body.dataset.editorCollabSawRepairing = syncDebug.sawRepairing
      ? "1"
      : "0";
    document.body.dataset.editorCollabSawDesync = syncDebug.sawDesync
      ? "1"
      : "0";
    document.body.dataset.editorCollabNotes = notes;
  };

  const readRemotePresenceSummary = (): {
    visible: boolean;
    peerCount: number;
    cursorCount: number;
    selectionCount: number;
  } => {
    const layer = editorContainer.querySelector(
      'div[data-docview-remote-presence-layer="1"]',
    ) as HTMLDivElement | null;
    if (!layer) {
      return {
        visible: false,
        peerCount: 0,
        cursorCount: 0,
        selectionCount: 0,
      };
    }
    const visible = layer.dataset.docviewRemotePresenceVisible === "1";
    const peerCount = Number.parseInt(
      layer.dataset.docviewRemotePresenceCount ?? "0",
      10,
    );
    if (!visible || !Number.isFinite(peerCount) || peerCount <= 0) {
      return {
        visible: false,
        peerCount: 0,
        cursorCount: 0,
        selectionCount: 0,
      };
    }
    return {
      visible: true,
      peerCount,
      cursorCount: layer.querySelectorAll('[data-docview-remote-cursor="1"]')
        .length,
      selectionCount: layer.querySelectorAll(
        '[data-docview-remote-selection="1"]',
      ).length,
    };
  };

  if (runPeerFlow) {
    try {
      const ed = await ensureEditor();
      ed.loadBlank();
      await sleep(120);
      const peerCtx = resolveCanvasHarnessContext();
      if (peerCtx) {
        await focusCanvasInput(peerCtx);
        await sleep(60);
      }
      channel.addEventListener("message", (event) => {
        const msg = event.data as
          | { type?: string; requestId?: string }
          | undefined;
        if (!msg) return;
        if (msg.type === "read-text" && typeof msg.requestId === "string") {
          const presence = readRemotePresenceSummary();
          const rawText = getDocPlainText(ed);
          const rich = getCollabRichContentSummary(ed);
          const paragraphCount = (
            (
              ed.getDocEdit() as {
                viewModel?: () => { body?: unknown[] };
              } | null
            )?.viewModel?.().body ?? []
          ).length;
          channel.postMessage({
            type: "peer-text",
            requestId: msg.requestId,
            text: rawText,
            debug: `${syncDebugSummary()} raw=${JSON.stringify(rawText)} paras=${paragraphCount}`,
            repairRequests: syncDebug.repairRequests,
            divergenceDetected: syncDebug.divergenceDetected,
            divergenceCleared: syncDebug.divergenceCleared,
            stateRequestSend: syncDebug.stateRequestSend,
            sawRepairing: syncDebug.sawRepairing,
            sawDesync: syncDebug.sawDesync,
            peerPresenceCount: presence.peerCount,
            peerCursorCount: presence.cursorCount,
            peerSelectionCount: presence.selectionCount,
            peerTableCellA1: rich.tableCellA1,
            peerTableCellB1: rich.tableCellB1,
            peerImageCount: rich.imageCount,
            peerInlineImageRunCount: rich.inlineImageRunCount,
          });
          return;
        }
        if (msg.type === "drop-presence" && typeof msg.requestId === "string") {
          channel.postMessage({
            type: "peer-presence-dropped",
            requestId: msg.requestId,
          });
          queueMicrotask(() => {
            ed.destroy();
          });
          return;
        }
        if (
          msg.type === "pause-presence" &&
          typeof msg.requestId === "string"
        ) {
          ed.setAwarenessPausedForTests(true);
          channel.postMessage({
            type: "peer-presence-paused",
            requestId: msg.requestId,
          });
          return;
        }
        if (msg.type === "pause-sync" && typeof msg.requestId === "string") {
          ed.setTransportPausedForTests(true);
          channel.postMessage({
            type: "peer-sync-paused",
            requestId: msg.requestId,
          });
          return;
        }
        if (msg.type === "resume-sync" && typeof msg.requestId === "string") {
          ed.setTransportPausedForTests(false);
          channel.postMessage({
            type: "peer-sync-resumed",
            requestId: msg.requestId,
          });
          return;
        }
        if (msg.type === "shutdown") {
          channel.close();
          window.close();
        }
      });
      statusEl.textContent = "collab peer ready";
      channel.postMessage({ type: "peer-ready" });
      return;
    } catch (err) {
      const message = describeError(err);
      statusEl.textContent = `collab peer error: ${message}`;
      publish(
        "failed:peer-exception",
        "",
        "",
        0,
        0,
        "",
        "",
        0,
        0,
        "",
        "",
        "",
        "",
        0,
        0,
        "",
        "",
        "",
        "",
        0,
        0,
        { expired: false, peerCount: 0, cursorCount: 0, selectionCount: 0 },
        message,
      );
      channel.close();
      return;
    }
  }

  const pendingReads = new Map<
    string,
    (value: {
      text: string;
      debug: string;
      repairRequests: number;
      divergenceDetected: number;
      divergenceCleared: number;
      stateRequestSend: number;
      sawRepairing: boolean;
      sawDesync: boolean;
      peerPresenceCount: number;
      peerCursorCount: number;
      peerSelectionCount: number;
      peerTableCellA1: string;
      peerTableCellB1: string;
      peerImageCount: number;
      peerInlineImageRunCount: number;
    }) => void
  >();
  const pendingPresenceDrop = new Map<string, () => void>();
  const pendingPresencePause = new Map<string, () => void>();
  const pendingSyncPause = new Map<string, () => void>();
  const pendingSyncResume = new Map<string, () => void>();
  let peerReady = false;
  channel.addEventListener("message", (event) => {
    const msg = event.data as
      | {
          type?: string;
          requestId?: string;
          text?: string;
          debug?: string;
          repairRequests?: number;
          divergenceDetected?: number;
          divergenceCleared?: number;
          stateRequestSend?: number;
          sawRepairing?: boolean;
          sawDesync?: boolean;
          peerPresenceCount?: number;
          peerCursorCount?: number;
          peerSelectionCount?: number;
          peerTableCellA1?: string;
          peerTableCellB1?: string;
          peerImageCount?: number;
          peerInlineImageRunCount?: number;
        }
      | undefined;
    if (!msg) return;
    if (msg.type === "peer-ready") {
      peerReady = true;
      return;
    }
    if (
      msg.type === "peer-presence-paused" &&
      typeof msg.requestId === "string"
    ) {
      const resolve = pendingPresencePause.get(msg.requestId);
      if (!resolve) return;
      pendingPresencePause.delete(msg.requestId);
      resolve();
      return;
    }
    if (msg.type === "peer-sync-paused" && typeof msg.requestId === "string") {
      const resolve = pendingSyncPause.get(msg.requestId);
      if (!resolve) return;
      pendingSyncPause.delete(msg.requestId);
      resolve();
      return;
    }
    if (msg.type === "peer-sync-resumed" && typeof msg.requestId === "string") {
      const resolve = pendingSyncResume.get(msg.requestId);
      if (!resolve) return;
      pendingSyncResume.delete(msg.requestId);
      resolve();
      return;
    }
    if (
      msg.type === "peer-presence-dropped" &&
      typeof msg.requestId === "string"
    ) {
      const resolve = pendingPresenceDrop.get(msg.requestId);
      if (!resolve) return;
      pendingPresenceDrop.delete(msg.requestId);
      resolve();
      return;
    }
    if (msg.type === "peer-text" && typeof msg.requestId === "string") {
      const resolve = pendingReads.get(msg.requestId);
      if (!resolve) return;
      pendingReads.delete(msg.requestId);
      resolve({
        text: msg.text ?? "",
        debug: msg.debug ?? "",
        repairRequests:
          typeof msg.repairRequests === "number" ? msg.repairRequests : 0,
        divergenceDetected:
          typeof msg.divergenceDetected === "number"
            ? msg.divergenceDetected
            : 0,
        divergenceCleared:
          typeof msg.divergenceCleared === "number" ? msg.divergenceCleared : 0,
        stateRequestSend:
          typeof msg.stateRequestSend === "number" ? msg.stateRequestSend : 0,
        sawRepairing: msg.sawRepairing === true,
        sawDesync: msg.sawDesync === true,
        peerPresenceCount:
          typeof msg.peerPresenceCount === "number" ? msg.peerPresenceCount : 0,
        peerCursorCount:
          typeof msg.peerCursorCount === "number" ? msg.peerCursorCount : 0,
        peerSelectionCount:
          typeof msg.peerSelectionCount === "number"
            ? msg.peerSelectionCount
            : 0,
        peerTableCellA1: msg.peerTableCellA1 ?? "",
        peerTableCellB1: msg.peerTableCellB1 ?? "",
        peerImageCount:
          typeof msg.peerImageCount === "number" ? msg.peerImageCount : 0,
        peerInlineImageRunCount:
          typeof msg.peerInlineImageRunCount === "number"
            ? msg.peerInlineImageRunCount
            : 0,
      });
    }
  });

  const requestPeerText = (
    timeoutMs: number,
  ): Promise<{
    text: string;
    debug: string;
    repairRequests: number;
    divergenceDetected: number;
    divergenceCleared: number;
    stateRequestSend: number;
    sawRepairing: boolean;
    sawDesync: boolean;
    peerPresenceCount: number;
    peerCursorCount: number;
    peerSelectionCount: number;
    peerTableCellA1: string;
    peerTableCellB1: string;
    peerImageCount: number;
    peerInlineImageRunCount: number;
  }> => {
    return new Promise((resolve, reject) => {
      const requestId = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
      const timer = window.setTimeout(() => {
        pendingReads.delete(requestId);
        reject(new Error("peer read timeout"));
      }, timeoutMs);
      pendingReads.set(requestId, (value) => {
        window.clearTimeout(timer);
        resolve(value);
      });
      channel.postMessage({ type: "read-text", requestId });
    });
  };

  const requestPeerSyncPause = (timeoutMs: number): Promise<void> => {
    return new Promise((resolve, reject) => {
      const requestId = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
      const timer = window.setTimeout(() => {
        pendingSyncPause.delete(requestId);
        reject(new Error("peer sync pause timeout"));
      }, timeoutMs);
      pendingSyncPause.set(requestId, () => {
        window.clearTimeout(timer);
        resolve();
      });
      channel.postMessage({ type: "pause-sync", requestId });
    });
  };

  const requestPeerSyncResume = (timeoutMs: number): Promise<void> => {
    return new Promise((resolve, reject) => {
      const requestId = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
      const timer = window.setTimeout(() => {
        pendingSyncResume.delete(requestId);
        reject(new Error("peer sync resume timeout"));
      }, timeoutMs);
      pendingSyncResume.set(requestId, () => {
        window.clearTimeout(timer);
        resolve();
      });
      channel.postMessage({ type: "resume-sync", requestId });
    });
  };

  const requestPeerPresencePause = (timeoutMs: number): Promise<void> => {
    return new Promise((resolve, reject) => {
      const requestId = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
      const timer = window.setTimeout(() => {
        pendingPresencePause.delete(requestId);
        reject(new Error("peer presence pause timeout"));
      }, timeoutMs);
      pendingPresencePause.set(requestId, () => {
        window.clearTimeout(timer);
        resolve();
      });
      channel.postMessage({ type: "pause-presence", requestId });
    });
  };

  let peerFrame: HTMLIFrameElement | null = null;
  const createPeerFrame = (): HTMLIFrameElement => {
    const peerUrl = new URL(window.location.href);
    peerUrl.searchParams.set("room", roomId ?? "");
    peerUrl.searchParams.set("renderer", rendererOpt ?? "canvas");
    peerUrl.searchParams.set(
      "autorun",
      autoRunCollabPresenceExpiry
        ? "collab-presence-expiry-peer"
        : "collab-peer",
    );
    if (useLocalRelay) {
      peerUrl.searchParams.set("relay", "local");
    } else {
      peerUrl.searchParams.delete("relay");
    }

    const frame = document.createElement("iframe");
    frame.src = peerUrl.toString();
    frame.style.position = "fixed";
    frame.style.left = "-12000px";
    frame.style.top = "0";
    frame.style.width = "1px";
    frame.style.height = "1px";
    frame.style.opacity = "0";
    frame.setAttribute("aria-hidden", "true");
    frame.dataset.docviewCollabPeerFrame = "1";
    document.body.appendChild(frame);
    peerFrame = frame;
    return frame;
  };

  const waitForPeerReady = async (timeoutMs: number): Promise<void> => {
    peerReady = false;
    const deadline = Date.now() + timeoutMs;
    while (!peerReady && Date.now() < deadline) {
      await sleep(50);
    }
    if (!peerReady) {
      throw new Error("peer did not become ready");
    }
  };

  try {
    const ed = await ensureEditor();
    ed.loadBlank();
    await sleep(120);

    const ctx = resolveCanvasHarnessContext();
    if (!ctx) {
      throw new Error("no canvas harness context for page A");
    }
    await focusCanvasInput(ctx);

    const expected = `Collab sync ${Date.now().toString(36)} OK`;
    for (const ch of expected) {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(8);
    }

    let localText = "";
    const localDeadline = Date.now() + 2_000;
    while (Date.now() < localDeadline) {
      localText = trimBoundaryNewlines(getDocPlainText(ed));
      if (localText === expected) break;
      await sleep(30);
    }
    if (localText !== expected) {
      throw new Error(
        `local typing mismatch: expected '${expected}', got '${localText || "<empty>"}'`,
      );
    }

    if (!ed.insertTable(2, 2)) {
      throw new Error("controller rejected collab table insertion");
    }
    await sleep(60);
    if (!ed.setActiveTableCellText("A1")) {
      throw new Error("controller rejected collab table cell A1 edit");
    }
    await sleep(40);
    if (!ed.moveActiveTableCell(0, 1)) {
      throw new Error("controller rejected collab table cell navigation");
    }
    if (!ed.setActiveTableCellText("B1")) {
      throw new Error("controller rejected collab table cell B1 edit");
    }
    ed.clearActiveTableCell();

    const imageFile = await makeAutorunImageFile();
    const imageDataUri = await fileToDataUrl(imageFile);
    const imageSize = await measureImagePt(imageDataUri);
    if (
      !ed.insertInlineImage({
        dataUri: imageDataUri,
        widthPt: imageSize.widthPt,
        heightPt: imageSize.heightPt,
        name: imageFile.name || undefined,
        description: imageFile.name || undefined,
      })
    ) {
      throw new Error("controller rejected collab inline image insertion");
    }

    let peerText = "";
    let peerDebug = "";
    let peerPresenceCount = 0;
    let peerCursorCount = 0;
    let peerSelectionCount = 0;
    let peerTableCellA1 = "";
    let peerTableCellB1 = "";
    let peerImageCount = 0;
    let peerInlineImageRunCount = 0;
    let replaceExpected = "";
    let replacePeerText = "";
    let replacePeerTableCellA1 = "";
    let replacePeerTableCellB1 = "";
    let replacePeerImageCount = 0;
    let replacePeerInlineImageRunCount = 0;
    let reconnectExpected = "";
    let reconnectPeerText = "";
    let reconnectPeerTableCellA1 = "";
    let reconnectPeerTableCellB1 = "";
    let reconnectPeerImageCount = 0;
    let reconnectPeerInlineImageRunCount = 0;
    let reconnectPeerDebug = "";
    let reconnectPeerRepairRequests = 0;
    let reconnectPeerDivergenceDetected = 0;
    let reconnectPeerDivergenceCleared = 0;
    let reconnectPeerStateRequestSend = 0;
    let reconnectPeerSawRepairing = false;
    let reconnectPeerSawDesync = false;

    createPeerFrame();
    await waitForPeerReady(12_000);
    await focusCanvasInput(ctx);
    dispatchKey(ctx.input, "ArrowLeft", { shift: true });
    dispatchKey(ctx.input, "ArrowLeft", { shift: true });
    await sleep(80);

    const syncDeadline = Date.now() + 12_000;
    while (Date.now() < syncDeadline) {
      try {
        const peerResult = await requestPeerText(1200);
        peerText = peerResult.text;
        peerDebug = peerResult.debug;
        peerTableCellA1 = peerResult.peerTableCellA1;
        peerTableCellB1 = peerResult.peerTableCellB1;
        peerImageCount = peerResult.peerImageCount;
        peerInlineImageRunCount = peerResult.peerInlineImageRunCount;
      } catch {
        // Keep the last good peer snapshot; transient harness reads should
        // not erase already-observed convergence.
      }
      const localPresence = readRemotePresenceSummary();
      peerPresenceCount = localPresence.peerCount;
      peerCursorCount = localPresence.cursorCount;
      peerSelectionCount = localPresence.selectionCount;
      if (
        trimBoundaryNewlines(peerText) === expected &&
        peerPresenceCount > 0 &&
        peerCursorCount > 0 &&
        peerTableCellA1 === "A1" &&
        peerTableCellB1 === "B1" &&
        peerImageCount >= 1 &&
        peerInlineImageRunCount >= 1
      ) {
        break;
      }
      await sleep(120);
    }

    const normalizedPeerText = trimBoundaryNewlines(peerText);
    if (
      normalizedPeerText !== expected ||
      peerPresenceCount < 1 ||
      peerCursorCount < 1 ||
      peerTableCellA1 !== "A1" ||
      peerTableCellB1 !== "B1" ||
      peerImageCount < 1 ||
      peerInlineImageRunCount < 1
    ) {
      throw new Error(
        `late-join sync mismatch: text='${normalizedPeerText || "<empty>"}' presence=${peerPresenceCount}/${peerCursorCount} table='${peerTableCellA1}/${peerTableCellB1}' images=${peerImageCount} runs=${peerInlineImageRunCount}; local(${syncDebugSummary()}) peer(${peerDebug || "n/a"})`,
      );
    }

    ed.replaceBlank();
    await sleep(120);
    await focusCanvasInput(ctx);
    replaceExpected = `Replaced ${Date.now().toString(36)} ready`;
    for (const ch of replaceExpected) {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(6);
    }
    const replaceLocalDeadline = Date.now() + 2_500;
    while (Date.now() < replaceLocalDeadline) {
      const localTextAfterReplace = trimBoundaryNewlines(getDocPlainText(ed));
      const localRichAfterReplace = getCollabRichContentSummary(ed);
      if (
        localTextAfterReplace === replaceExpected &&
        localRichAfterReplace.tableCellA1 === "" &&
        localRichAfterReplace.tableCellB1 === "" &&
        localRichAfterReplace.imageCount === 0 &&
        localRichAfterReplace.inlineImageRunCount === 0
      ) {
        break;
      }
      await sleep(40);
    }

    const replaceSyncDeadline = Date.now() + 12_000;
    while (Date.now() < replaceSyncDeadline) {
      try {
        const peerResult = await requestPeerText(1200);
        replacePeerText = trimBoundaryNewlines(peerResult.text);
        replacePeerTableCellA1 = peerResult.peerTableCellA1;
        replacePeerTableCellB1 = peerResult.peerTableCellB1;
        replacePeerImageCount = peerResult.peerImageCount;
        replacePeerInlineImageRunCount = peerResult.peerInlineImageRunCount;
      } catch {
        // Retry until convergence.
      }
      if (
        replacePeerText === replaceExpected &&
        replacePeerTableCellA1 === "" &&
        replacePeerTableCellB1 === "" &&
        replacePeerImageCount === 0 &&
        replacePeerInlineImageRunCount === 0
      ) {
        break;
      }
      await sleep(120);
    }
    if (
      replacePeerText !== replaceExpected ||
      replacePeerTableCellA1 !== "" ||
      replacePeerTableCellB1 !== "" ||
      replacePeerImageCount !== 0 ||
      replacePeerInlineImageRunCount !== 0
    ) {
      throw new Error(
        `replace sync mismatch: text='${replacePeerText || "<empty>"}' table='${replacePeerTableCellA1}/${replacePeerTableCellB1}' images=${replacePeerImageCount} runs=${replacePeerInlineImageRunCount}; local(${syncDebugSummary()}) peer(${peerDebug || "n/a"})`,
      );
    }

    await requestPeerSyncPause(3_000);
    await focusCanvasInput(ctx);
    reconnectExpected = `${replaceExpected} / reconnect`;
    const reconnectSuffix = reconnectExpected.slice(replaceExpected.length);
    for (const ch of reconnectSuffix) {
      dispatchBeforeInput(ctx.input, "insertText", ch);
      await sleep(6);
    }
    if (!ed.insertTable(2, 2)) {
      throw new Error("controller rejected reconnect table insertion");
    }
    await sleep(60);
    if (!ed.setActiveTableCellText("R1")) {
      throw new Error("controller rejected reconnect table cell A1 edit");
    }
    await sleep(40);
    if (!ed.moveActiveTableCell(0, 1)) {
      throw new Error("controller rejected reconnect table cell navigation");
    }
    if (!ed.setActiveTableCellText("R2")) {
      throw new Error("controller rejected reconnect table cell B1 edit");
    }
    ed.clearActiveTableCell();
    if (
      !ed.insertInlineImage({
        dataUri: imageDataUri,
        widthPt: imageSize.widthPt,
        heightPt: imageSize.heightPt,
        name: imageFile.name || undefined,
        description: imageFile.name || undefined,
      })
    ) {
      throw new Error("controller rejected reconnect inline image insertion");
    }

    const reconnectLocalDeadline = Date.now() + 2_500;
    while (Date.now() < reconnectLocalDeadline) {
      const localTextAfterReconnect = trimBoundaryNewlines(getDocPlainText(ed));
      const localRichAfterReconnect = getCollabRichContentSummary(ed);
      if (
        localTextAfterReconnect === reconnectExpected &&
        localRichAfterReconnect.tableCellA1 === "R1" &&
        localRichAfterReconnect.tableCellB1 === "R2" &&
        localRichAfterReconnect.imageCount >= 1 &&
        localRichAfterReconnect.inlineImageRunCount >= 1
      ) {
        break;
      }
      await sleep(40);
    }

    await requestPeerSyncResume(3_000);

    const reconnectSyncDeadline = Date.now() + 12_000;
    while (Date.now() < reconnectSyncDeadline) {
      try {
        const peerResult = await requestPeerText(1200);
        reconnectPeerText = trimBoundaryNewlines(peerResult.text);
        reconnectPeerTableCellA1 = peerResult.peerTableCellA1;
        reconnectPeerTableCellB1 = peerResult.peerTableCellB1;
        reconnectPeerImageCount = peerResult.peerImageCount;
        reconnectPeerInlineImageRunCount = peerResult.peerInlineImageRunCount;
        reconnectPeerDebug = peerResult.debug;
        reconnectPeerRepairRequests = peerResult.repairRequests;
        reconnectPeerDivergenceDetected = peerResult.divergenceDetected;
        reconnectPeerDivergenceCleared = peerResult.divergenceCleared;
        reconnectPeerStateRequestSend = peerResult.stateRequestSend;
        reconnectPeerSawRepairing = peerResult.sawRepairing;
        reconnectPeerSawDesync = peerResult.sawDesync;
      } catch {
        // Retry until convergence.
      }
      if (
        reconnectPeerText === reconnectExpected &&
        reconnectPeerTableCellA1 === "R1" &&
        reconnectPeerTableCellB1 === "R2" &&
        reconnectPeerImageCount >= 1 &&
        reconnectPeerInlineImageRunCount >= 1 &&
        reconnectPeerRepairRequests >= 1 &&
        reconnectPeerDivergenceDetected >= 1 &&
        reconnectPeerDivergenceCleared >= 1 &&
        reconnectPeerStateRequestSend >= 1
      ) {
        break;
      }
      await sleep(120);
    }
    document.body.dataset.editorCollabReconnectPeerDebug = reconnectPeerDebug;
    document.body.dataset.editorCollabReconnectPeerRepairRequests = String(
      reconnectPeerRepairRequests,
    );
    document.body.dataset.editorCollabReconnectPeerDivergenceDetected = String(
      reconnectPeerDivergenceDetected,
    );
    document.body.dataset.editorCollabReconnectPeerDivergenceCleared = String(
      reconnectPeerDivergenceCleared,
    );
    document.body.dataset.editorCollabReconnectPeerStateRequestSend = String(
      reconnectPeerStateRequestSend,
    );
    document.body.dataset.editorCollabReconnectPeerSawRepairing =
      reconnectPeerSawRepairing ? "1" : "0";
    document.body.dataset.editorCollabReconnectPeerSawDesync =
      reconnectPeerSawDesync ? "1" : "0";
    if (
      reconnectPeerText !== reconnectExpected ||
      reconnectPeerTableCellA1 !== "R1" ||
      reconnectPeerTableCellB1 !== "R2" ||
      reconnectPeerImageCount < 1 ||
      reconnectPeerInlineImageRunCount < 1 ||
      reconnectPeerRepairRequests < 1 ||
      reconnectPeerDivergenceDetected < 1 ||
      reconnectPeerDivergenceCleared < 1 ||
      reconnectPeerStateRequestSend < 1
    ) {
      throw new Error(
        `reconnect sync mismatch: text='${reconnectPeerText || "<empty>"}' table='${reconnectPeerTableCellA1}/${reconnectPeerTableCellB1}' images=${reconnectPeerImageCount} runs=${reconnectPeerInlineImageRunCount}; local(${syncDebugSummary()}) peer(${reconnectPeerDebug || "n/a"})`,
      );
    }

    const expiredPresence = {
      expired: false,
      peerCount: peerPresenceCount,
      cursorCount: peerCursorCount,
      selectionCount: peerSelectionCount,
    };

    if (autoRunCollabPresenceExpiry) {
      await requestPeerPresencePause(3_000);
      const expiryDeadline = Date.now() + 12_000;
      while (Date.now() < expiryDeadline) {
        window.dispatchEvent(new Event("resize"));
        const localPresence = readRemotePresenceSummary();
        expiredPresence.expired =
          !localPresence.visible && localPresence.peerCount === 0;
        expiredPresence.peerCount = localPresence.peerCount;
        expiredPresence.cursorCount = localPresence.cursorCount;
        expiredPresence.selectionCount = localPresence.selectionCount;
        if (expiredPresence.expired) {
          break;
        }
        await sleep(150);
      }
      if (!expiredPresence.expired) {
        throw new Error(
          `presence did not expire: peers=${expiredPresence.peerCount} cursors=${expiredPresence.cursorCount} selections=${expiredPresence.selectionCount}; local(${syncDebugSummary()})`,
        );
      }
      channel.postMessage({ type: "shutdown" });
    }

    statusEl.textContent = "collab autorun complete";
    publish(
      "complete",
      expected,
      normalizedPeerText,
      peerPresenceCount,
      peerCursorCount,
      peerTableCellA1,
      peerTableCellB1,
      peerImageCount,
      peerInlineImageRunCount,
      replaceExpected,
      replacePeerText,
      replacePeerTableCellA1,
      replacePeerTableCellB1,
      replacePeerImageCount,
      replacePeerInlineImageRunCount,
      reconnectExpected,
      reconnectPeerText,
      reconnectPeerTableCellA1,
      reconnectPeerTableCellB1,
      reconnectPeerImageCount,
      reconnectPeerInlineImageRunCount,
      expiredPresence,
      "ok",
    );
  } catch (err) {
    const message = describeError(err);
    statusEl.textContent = `collab autorun error: ${message}`;
    publish(
      "failed:exception",
      "",
      "",
      0,
      0,
      "",
      "",
      0,
      0,
      "",
      "",
      "",
      "",
      0,
      0,
      "",
      "",
      "",
      "",
      0,
      0,
      { expired: false, peerCount: 0, cursorCount: 0, selectionCount: 0 },
      message,
    );
  } finally {
    channel.postMessage({ type: "shutdown" });
    channel.close();
    if (peerFrame) {
      peerFrame.remove();
    }
  }
}

void runAutoCanvasCursorScenario();
void runAutoCanvasOpsScenario();
void runAutoCopyLinkScenario();
void runAutoCollabSyncScenario();
