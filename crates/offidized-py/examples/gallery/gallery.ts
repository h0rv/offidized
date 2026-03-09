// Gallery client — tree rendering, viewer loading, file selection

type ExampleTree = Record<string, { script: string; name: string }[]>;

const sidebar = document.getElementById("sidebar")!;
const viewerContainer = document.getElementById("viewer-container")!;
const generateBtn = document.getElementById(
  "generate-btn",
) as HTMLButtonElement;
const spinner = document.getElementById("spinner")!;
const statusEl = document.getElementById("status")!;
const fileNameEl = document.getElementById("file-name")!;
const renderTimeEl = document.getElementById("render-time")!;

// Viewer state — one instance per format, reused
const viewers: Record<string, { viewer: any; mod: any }> = {};
let activeFormat: string | null = null;
let activeButton: HTMLButtonElement | null = null;

// Script → generated files mapping (set after generate)
// Key is "fmt/scriptname" e.g. "xlsx/beyond_openpyxl"
let scriptFiles: Record<string, string[]> = {};
let generatedFiles = new Set<string>();
let exampleTree: ExampleTree = {};

// ── Module loading ───────────────────────────────────────────────────
function viewerJsPath(ext: string): string {
  return ext === "xlsx"
    ? "/preview/xl-view.js"
    : ext === "docx"
      ? "/preview/doc-view.js"
      : "/preview/ppt-view.js";
}

function wasmPath(ext: string): string {
  return ext === "xlsx"
    ? "/preview/offidized_xlview_bg.wasm"
    : ext === "docx"
      ? "/preview/offidized_docview_bg.wasm"
      : "/preview/offidized_pptview_bg.wasm";
}

async function ensureViewer(ext: string) {
  if (viewers[ext]) return viewers[ext].viewer;

  const jsRes = await fetch(viewerJsPath(ext), { method: "HEAD" });
  if (!jsRes.ok) return null;

  const mod = await import(viewerJsPath(ext));
  await mod.init(wasmPath(ext));
  const viewer = await mod.mount(viewerContainer);
  viewers[ext] = { viewer, mod };
  return viewer;
}

// ── Tree rendering ───────────────────────────────────────────────────
const FORMAT_META: Record<string, { label: string; badge: string }> = {
  pptx: { label: "PowerPoint", badge: "badge-pptx" },
  docx: { label: "Word", badge: "badge-docx" },
  xlsx: { label: "Excel", badge: "badge-xlsx" },
};

function prettyName(name: string): string {
  return name
    .replace(/^\d+_/, "")
    .replace(/_/g, " ")
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

function renderTree(tree: ExampleTree) {
  exampleTree = tree;
  sidebar.innerHTML = "";
  for (const [fmt, scripts] of Object.entries(tree)) {
    const meta = FORMAT_META[fmt] ?? { label: fmt, badge: "" };

    const group = document.createElement("div");
    group.className = "group";

    const heading = document.createElement("div");
    heading.className = "group-header";
    heading.innerHTML = `${meta.label} <span class="group-badge ${meta.badge}">.${fmt}</span>`;
    group.appendChild(heading);

    for (const { name } of scripts) {
      const key = `${fmt}/${name}`;
      const files = scriptFiles[key];

      if (files && files.length > 1) {
        // Script produced multiple files — show each as a sub-item
        const scriptLabel = document.createElement("div");
        scriptLabel.className = "item-parent";
        scriptLabel.textContent = prettyName(name);
        group.appendChild(scriptLabel);

        for (const file of files) {
          const item = document.createElement("button");
          item.className = "item item-sub generated";
          item.textContent = file;
          item.title = file;
          item.dataset.file = file;
          item.dataset.format = fmt;
          item.addEventListener("click", () => loadFile(fmt, file, item));
          group.appendChild(item);
        }
      } else {
        const item = document.createElement("button");
        item.className = "item";
        item.textContent = prettyName(name);
        item.dataset.format = fmt;
        item.dataset.name = name;

        const file = files?.[0];
        if (file) {
          item.classList.add("generated");
          item.dataset.file = file;
          item.title = file;
          item.addEventListener("click", () => loadFile(fmt, file, item));
        } else {
          item.title = `${name}.${fmt}`;
          item.addEventListener("click", () => {
            setStatus(
              `No generated file for ${name} — click Generate All first`,
            );
          });
        }
        group.appendChild(item);
      }
    }

    sidebar.appendChild(group);
  }
}

// ── File loading ─────────────────────────────────────────────────────
async function loadFile(fmt: string, filename: string, btn: HTMLButtonElement) {
  if (activeButton) activeButton.classList.remove("active");
  btn.classList.add("active");
  activeButton = btn;

  setStatus(`Loading ${filename}...`);
  fileNameEl.textContent = filename;
  renderTimeEl.textContent = "";

  try {
    const res = await fetch(`/api/files/${encodeURIComponent(filename)}`);
    if (!res.ok) {
      setStatus(`Failed to load ${filename}: ${res.status}`);
      return;
    }
    const buf = await res.arrayBuffer();

    // Clear empty state on first load
    const emptyState = viewerContainer.querySelector(".empty-state");
    if (emptyState) emptyState.remove();

    // Destroy old viewer if switching formats
    if (activeFormat && activeFormat !== fmt && viewers[activeFormat]) {
      try {
        viewers[activeFormat].viewer.destroy();
      } catch {}
      delete viewers[activeFormat];
      viewerContainer.innerHTML = "";
    }
    activeFormat = fmt;

    const viewer = await ensureViewer(fmt);
    if (!viewer) {
      setStatus(`No viewer available for .${fmt} — build WASM first`);
      return;
    }
    viewerContainer.classList.add("has-viewer");
    const t0 = performance.now();
    viewer.load(new Uint8Array(buf));
    const elapsed = Math.round(performance.now() - t0);
    renderTimeEl.textContent = `${elapsed}ms`;
    setStatus(`Loaded ${filename}`);
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    setStatus(`Error: ${msg}`);
  }
}

// ── Generate all ─────────────────────────────────────────────────────
generateBtn.addEventListener("click", async () => {
  generateBtn.disabled = true;
  generateBtn.classList.add("running");
  generateBtn.textContent = "Generating...";
  spinner.classList.add("visible");
  setStatus("Running all examples...");

  try {
    const res = await fetch("/api/generate", { method: "POST" });
    const data = await res.json();

    // Build script → files mapping
    scriptFiles = {};
    for (const r of data.results) {
      if (!r.ok) continue;
      // r.script is like "xlsx/beyond_openpyxl.py"
      const key = r.script.replace(".py", "");
      scriptFiles[key] = r.files.sort();
    }

    generatedFiles = new Set(data.files);
    // Re-render tree with file mapping
    renderTree(exampleTree);

    const ok = data.results.filter((r: any) => r.ok).length;
    const fail = data.results.filter((r: any) => !r.ok).length;
    setStatus(
      `Generated ${data.files.length} files (${ok} scripts OK${fail > 0 ? `, ${fail} failed` : ""})`,
    );

    for (const r of data.results) {
      if (!r.ok) console.error(`FAIL: ${r.script}`, r.error);
    }
  } catch (e: unknown) {
    const msg = e instanceof Error ? e.message : String(e);
    setStatus(`Generation failed: ${msg}`);
  } finally {
    generateBtn.disabled = false;
    generateBtn.classList.remove("running");
    generateBtn.textContent = "Generate All";
    spinner.classList.remove("visible");
  }
});

// ── Load existing files on startup ───────────────────────────────────
async function loadExistingFiles() {
  try {
    const res = await fetch("/api/files");
    const files: string[] = await res.json();
    if (files.length === 0) return;

    generatedFiles = new Set(files);

    // Without a generate run, we guess the mapping by matching filenames
    // to scripts. For 1:1 scripts (pptx/docx) this works; for multi-file
    // xlsx scripts we just assign all xlsx files to each xlsx script.
    for (const [fmt, scripts] of Object.entries(exampleTree)) {
      const fmtFiles = files.filter((f) => f.endsWith(`.${fmt}`));
      for (const { name } of scripts) {
        const key = `${fmt}/${name}`;
        const exact = fmtFiles.filter((f) => f.startsWith(name));
        if (exact.length > 0) {
          scriptFiles[key] = exact.sort();
        } else if (fmtFiles.length > 0) {
          // Multi-file script (like beyond_openpyxl) — assign all unmatched files
          scriptFiles[key] = fmtFiles.sort();
        }
      }
    }

    renderTree(exampleTree);
    setStatus(`${files.length} files already generated`);
  } catch {}
}

// ── Status helper ────────────────────────────────────────────────────
function setStatus(msg: string) {
  statusEl.textContent = msg;
}

// ── Init ─────────────────────────────────────────────────────────────
async function init() {
  setStatus("Loading examples...");
  const res = await fetch("/api/examples");
  exampleTree = await res.json();
  renderTree(exampleTree);
  await loadExistingFiles();
  if (generatedFiles.size === 0) {
    setStatus("Ready — select an example or click Generate All");
  }
}

init();
