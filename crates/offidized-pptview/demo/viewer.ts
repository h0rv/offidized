// Demo viewer: file upload, drag-drop, slide navigation, parse timing.

import { init, mount, type MountedViewer } from "../js/ppt-view.ts";

// Initialize WASM
await init();

const fileInput = document.getElementById("file-input") as HTMLInputElement;
const viewerContainer = document.getElementById(
  "viewer-container",
) as HTMLDivElement;
const emptyState = document.getElementById("empty-state") as HTMLDivElement;
const statusEl = document.getElementById("status") as HTMLSpanElement;
const slideCounter = document.getElementById(
  "slide-counter",
) as HTMLSpanElement;
const btnPrev = document.getElementById("btn-prev") as HTMLButtonElement;
const btnNext = document.getElementById("btn-next") as HTMLButtonElement;
const dropOverlay = document.getElementById("drop-overlay") as HTMLDivElement;

let viewer: MountedViewer | null = null;

// ---- Mount viewer ----

async function ensureViewer(): Promise<MountedViewer> {
  if (!viewer) {
    emptyState.style.display = "none";
    viewer = await mount(viewerContainer);
    viewer.onSlideChange(updateSlideCounter);
  }
  return viewer;
}

function updateSlideCounter(index?: number): void {
  if (!viewer) return;
  const current = (index ?? viewer.currentSlide()) + 1;
  const total = viewer.slideCount();
  slideCounter.textContent = `${current} / ${total}`;
  btnPrev.disabled = current <= 1;
  btnNext.disabled = current >= total;
}

// ---- File handling ----

async function handleFile(file: File): Promise<void> {
  if (!file.name.endsWith(".pptx")) {
    statusEl.textContent = "Only .pptx files are supported";
    return;
  }

  try {
    const v = await ensureViewer();
    const data = new Uint8Array(await file.arrayBuffer());
    const ms = v.load(data);
    statusEl.textContent = `${file.name} \u2014 parsed in ${ms}ms`;
    updateSlideCounter();
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

// ---- Navigation buttons ----

btnPrev.addEventListener("click", () => {
  viewer?.prevSlide();
  updateSlideCounter();
});

btnNext.addEventListener("click", () => {
  viewer?.nextSlide();
  updateSlideCounter();
});
