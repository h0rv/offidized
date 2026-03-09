// Test fixture gallery — browse all XLSX test files through the viewer
import { resolve, basename } from "path";
import { readdirSync, existsSync, mkdirSync } from "fs";

const PORT = 3002;
const TEST_DIR = resolve(import.meta.dir);
const XLVIEW_ROOT = resolve(import.meta.dir, "..");
const PREVIEW_DIR = resolve(import.meta.dir, ".preview");

// ── Build viewer assets ─────────────────────────────────────────────
async function buildViewer() {
  const jsPath = resolve(XLVIEW_ROOT, "js/xl-view.ts");
  const wasmName = "offidized_xlview_bg.wasm";
  const wasmSrc = resolve(XLVIEW_ROOT, "pkg", wasmName);

  if (!existsSync(jsPath)) {
    console.log("js/xl-view.ts not found — run from crate root");
    return false;
  }
  if (!existsSync(wasmSrc)) {
    console.log(
      `${wasmName} not built — run: wasm-pack build crates/offidized-xlview --target web`,
    );
    return false;
  }

  mkdirSync(PREVIEW_DIR, { recursive: true });
  const result = await Bun.build({
    entrypoints: [jsPath],
    outdir: PREVIEW_DIR,
    target: "browser",
    naming: "[name].[ext]",
  });

  if (!result.success) {
    console.error("Viewer build failed:", result.logs);
    return false;
  }

  await Bun.write(resolve(PREVIEW_DIR, wasmName), Bun.file(wasmSrc));
  console.log("Viewer assets built");
  return true;
}

const viewerReady = await buildViewer();

// ── Scan test fixtures ──────────────────────────────────────────────
function listFixtures(): string[] {
  return readdirSync(TEST_DIR)
    .filter((f) => f.endsWith(".xlsx"))
    .sort();
}

// ── HTML ─────────────────────────────────────────────────────────────
const html = `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>offidized-xlview Test Fixtures</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    :root {
      --bg: #0c0c14; --surface: #14141f; --surface2: #1a1a2a;
      --border: #2a2a3a; --text: #e4e4ef; --text-dim: #7a7a8f;
      --text-muted: #4a4a5f; --accent: #6366f1; --green: #34d399;
    }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif;
      background: var(--bg); color: var(--text);
      height: 100vh; display: flex; flex-direction: column; overflow: hidden;
    }
    header {
      display: flex; align-items: center; gap: 16px;
      padding: 10px 20px; background: var(--surface);
      border-bottom: 1px solid var(--border); flex-shrink: 0;
    }
    .logo { font-size: 13px; font-weight: 600; letter-spacing: 0.5px; color: var(--text-dim); }
    .logo span { color: var(--accent); }
    .header-right { margin-left: auto; display: flex; align-items: center; gap: 12px; }
    #file-name { font-size: 12px; color: var(--text-dim); }
    #file-size { font-size: 11px; color: var(--text-muted); font-variant-numeric: tabular-nums; }
    .main { display: flex; flex: 1; min-height: 0; }
    #sidebar {
      width: 240px; flex-shrink: 0; background: var(--surface);
      border-right: 1px solid var(--border); overflow-y: auto; padding: 12px 0;
    }
    #sidebar::-webkit-scrollbar { width: 4px; }
    #sidebar::-webkit-scrollbar-thumb { background: var(--border); border-radius: 2px; }
    .group-header {
      padding: 6px 16px; font-size: 10px; font-weight: 700;
      text-transform: uppercase; letter-spacing: 1.2px; color: var(--text-muted);
    }
    .item {
      display: block; width: 100%; text-align: left;
      padding: 5px 16px 5px 24px; font-size: 13px; color: var(--text-dim);
      background: none; border: none; cursor: pointer;
      transition: color 0.1s, background 0.1s;
      white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    }
    .item:hover { color: var(--text); background: var(--surface2); }
    .item.active { color: #fff; background: var(--surface2); }
    .viewer-wrap { flex: 1; display: flex; flex-direction: column; min-width: 0; background: var(--bg); }
    #viewer-container {
      flex: 1; min-height: 0; overflow: auto;
      display: flex; align-items: center; justify-content: center; position: relative;
    }
    #viewer-container.has-viewer { display: block; overflow: hidden; }
    .empty-state { text-align: center; color: var(--text-muted); }
    .empty-state .icon { font-size: 48px; margin-bottom: 12px; opacity: 0.3; }
    .empty-state p { font-size: 13px; line-height: 1.6; }
    footer {
      padding: 5px 20px; background: var(--surface);
      border-top: 1px solid var(--border); flex-shrink: 0;
    }
    #status { font-size: 11px; color: var(--text-muted); }
  </style>
</head>
<body>
  <header>
    <div class="logo"><span>offidized-xlview</span> Test Fixtures</div>
    <div class="header-right">
      <span id="file-name"></span>
      <span id="file-size"></span>
    </div>
  </header>
  <div class="main">
    <aside id="sidebar"></aside>
    <div class="viewer-wrap">
      <div id="viewer-container">
        <div class="empty-state">
          <div class="icon">&#9634;</div>
          <p>Select a test fixture from the sidebar</p>
        </div>
      </div>
    </div>
  </div>
  <footer><span id="status">Loading...</span></footer>

  <script type="module">
    const sidebar = document.getElementById("sidebar");
    const viewerContainer = document.getElementById("viewer-container");
    const fileNameEl = document.getElementById("file-name");
    const fileSizeEl = document.getElementById("file-size");
    const statusEl = document.getElementById("status");
    let viewer = null, mod = null, activeBtn = null;

    async function ensureViewer() {
      if (viewer) return viewer;
      const jsRes = await fetch("/preview/xl-view.js", { method: "HEAD" });
      if (!jsRes.ok) return null;
      mod = await import("/preview/xl-view.js");
      await mod.init("/preview/offidized_xlview_bg.wasm");
      viewer = await mod.mount(viewerContainer);
      return viewer;
    }

    async function loadFile(name, btn) {
      if (activeBtn) activeBtn.classList.remove("active");
      btn.classList.add("active");
      activeBtn = btn;
      statusEl.textContent = "Loading " + name + "...";
      fileNameEl.textContent = name;
      try {
        const res = await fetch("/files/" + encodeURIComponent(name));
        if (!res.ok) { statusEl.textContent = "Failed: " + res.status; return; }
        const buf = await res.arrayBuffer();
        fileSizeEl.textContent = (buf.byteLength / 1024).toFixed(1) + " KB";
        const empty = viewerContainer.querySelector(".empty-state");
        if (empty) empty.remove();
        const v = await ensureViewer();
        if (!v) { statusEl.textContent = "Viewer not available — build WASM first"; return; }
        viewerContainer.classList.add("has-viewer");
        const t0 = performance.now();
        v.load(new Uint8Array(buf));
        statusEl.textContent = name + " — " + Math.round(performance.now() - t0) + "ms";
      } catch (e) { statusEl.textContent = "Error: " + e.message; }
    }

    async function init() {
      const res = await fetch("/api/fixtures");
      const files = await res.json();
      const header = document.createElement("div");
      header.className = "group-header";
      header.textContent = "Test Fixtures (" + files.length + ")";
      sidebar.appendChild(header);
      for (const f of files) {
        const btn = document.createElement("button");
        btn.className = "item";
        btn.textContent = f.replace(".xlsx", "");
        btn.title = f;
        btn.addEventListener("click", () => loadFile(f, btn));
        sidebar.appendChild(btn);
      }
      statusEl.textContent = files.length + " fixtures — select one to preview";
    }
    init();
  </script>
</body>
</html>`;

// ── Server ───────────────────────────────────────────────────────────
const server = Bun.serve({
  port: PORT,
  async fetch(req) {
    const url = new URL(req.url);

    if (url.pathname === "/" || url.pathname === "/index.html") {
      return new Response(html, { headers: { "Content-Type": "text/html" } });
    }

    if (url.pathname === "/api/fixtures") {
      return Response.json(listFixtures());
    }

    if (url.pathname.startsWith("/files/")) {
      const name = decodeURIComponent(url.pathname.slice("/files/".length));
      if (name.includes("..") || name.includes("/"))
        return new Response("Forbidden", { status: 403 });
      const filePath = resolve(TEST_DIR, name);
      if (!existsSync(filePath))
        return new Response("Not found", { status: 404 });
      return new Response(Bun.file(filePath));
    }

    if (url.pathname.startsWith("/preview/")) {
      const name = url.pathname.slice("/preview/".length);
      if (name.includes(".."))
        return new Response("Forbidden", { status: 403 });
      const filePath = resolve(PREVIEW_DIR, name);
      if (!existsSync(filePath))
        return new Response("Not found", { status: 404 });
      const headers: Record<string, string> = {};
      if (name.endsWith(".wasm")) headers["Content-Type"] = "application/wasm";
      else if (name.endsWith(".js"))
        headers["Content-Type"] = "application/javascript";
      return new Response(Bun.file(filePath), { headers });
    }

    return new Response("Not found", { status: 404 });
  },
});

console.log(`Test fixture gallery at http://localhost:${server.port}`);
