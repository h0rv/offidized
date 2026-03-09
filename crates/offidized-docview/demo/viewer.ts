// Demo viewer: file upload, drag-drop, mode toggle, parse timing.

import { init, mount, type MountedViewer } from "../js/doc-view.ts";

// Initialize WASM
await init();

const fileInput = document.getElementById("file-input") as HTMLInputElement;
const viewerContainer = document.getElementById(
  "viewer-container",
) as HTMLDivElement;
const emptyState = document.getElementById("empty-state") as HTMLDivElement;
const statusEl = document.getElementById("status") as HTMLSpanElement;
const btnContinuous = document.getElementById(
  "btn-continuous",
) as HTMLButtonElement;
const btnPaginated = document.getElementById(
  "btn-paginated",
) as HTMLButtonElement;
const dropOverlay = document.getElementById("drop-overlay") as HTMLDivElement;

// Read renderer preference from URL: ?renderer=canvas or ?renderer=html.
// Without a param, mount() uses its default renderer (canvas).
const params = new URLSearchParams(window.location.search);
const rendererParam = params.get("renderer");
const rendererOpt: "html" | "canvas" | undefined =
  rendererParam === "canvas"
    ? "canvas"
    : rendererParam === "html"
      ? "html"
      : undefined;

let viewer: MountedViewer | null = null;

// ---- Mount viewer ----

async function ensureViewer(): Promise<MountedViewer> {
  if (!viewer) {
    emptyState.style.display = "none";
    viewer = await mount(viewerContainer, { renderer: rendererOpt });
    syncModeButtons(viewer.getMode());
  }
  return viewer;
}

// ---- File handling ----

async function handleFile(file: File): Promise<void> {
  if (!file.name.endsWith(".docx")) {
    statusEl.textContent = "Only .docx files are supported";
    return;
  }

  try {
    const v = await ensureViewer();
    const data = new Uint8Array(await file.arrayBuffer());
    const ms = v.load(data);
    statusEl.textContent = `${file.name} \u2014 parsed in ${ms}ms`;
  } catch (err) {
    statusEl.textContent = `Error: ${(err as Error).message}`;
  }
}

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

// ---- Mode toggle ----

function syncModeButtons(mode: "continuous" | "paginated"): void {
  btnContinuous.classList.toggle("active", mode === "continuous");
  btnPaginated.classList.toggle("active", mode === "paginated");
}

function setMode(mode: "continuous" | "paginated"): void {
  if (!viewer) return;
  viewer.setMode(mode);
  syncModeButtons(viewer.getMode());
}

btnContinuous.addEventListener("click", () => setMode("continuous"));
btnPaginated.addEventListener("click", () => setMode("paginated"));
