import { resolve, basename } from "path";
import {
  readdirSync,
  existsSync,
  mkdirSync,
  renameSync,
  statSync,
  rmSync,
} from "fs";
import homepage from "./index.html";

const PORT = 3001;
const ROOT = resolve(import.meta.dir, "..");
const PY_PROJECT = resolve(import.meta.dir, "../..");
const OUTPUT_DIR = resolve(import.meta.dir, "output");
const PREVIEW_DIR = resolve(import.meta.dir, ".preview");

const OFFICE_EXTS = [".xlsx", ".docx", ".pptx"];

// ── Build viewer assets ──────────────────────────────────────────────
const VIEWER_ROOTS = {
  xlsx: resolve(import.meta.dir, "../../../offidized-xlview"),
  docx: resolve(import.meta.dir, "../../../offidized-docview"),
  pptx: resolve(import.meta.dir, "../../../offidized-pptview"),
};

let previewReady = false;

async function buildViewers() {
  const entrypoints: string[] = [];
  const wasmCopies: Promise<void>[] = [];

  for (const [ext, root] of Object.entries(VIEWER_ROOTS)) {
    const jsFile =
      ext === "xlsx"
        ? "js/xl-view.ts"
        : ext === "docx"
          ? "js/doc-view.ts"
          : "js/ppt-view.ts";
    const jsPath = resolve(root, jsFile);
    if (!existsSync(jsPath)) continue;

    // Only include viewers whose WASM pkg is built
    const wasmName =
      ext === "xlsx"
        ? "offidized_xlview_bg.wasm"
        : ext === "docx"
          ? "offidized_docview_bg.wasm"
          : "offidized_pptview_bg.wasm";
    const wasmSrc = resolve(root, "pkg", wasmName);
    if (!existsSync(wasmSrc)) {
      console.log(`Skipping ${ext} viewer — ${wasmName} not built`);
      continue;
    }

    entrypoints.push(jsPath);
    wasmCopies.push(
      Bun.write(resolve(PREVIEW_DIR, wasmName), Bun.file(wasmSrc)).then(
        () => {},
      ),
    );
  }

  if (entrypoints.length === 0) {
    console.log("No viewer packages found — preview disabled");
    return;
  }

  mkdirSync(PREVIEW_DIR, { recursive: true });
  const result = await Bun.build({
    entrypoints,
    outdir: PREVIEW_DIR,
    target: "browser",
    naming: "[name].[ext]",
  });

  if (result.success) {
    await Promise.all(wasmCopies);
    previewReady = true;
    console.log(`Viewer assets built (${entrypoints.length} viewers)`);
  } else {
    console.error("Viewer build failed:", result.logs);
  }
}

await buildViewers();

// ── Scan examples ────────────────────────────────────────────────────
function scanExamples() {
  const formats = ["pptx", "docx", "xlsx"] as const;
  const tree: Record<string, { script: string; name: string }[]> = {};
  for (const fmt of formats) {
    const dir = resolve(ROOT, fmt);
    if (!existsSync(dir)) continue;
    tree[fmt] = readdirSync(dir)
      .filter((f) => f.endsWith(".py"))
      .sort()
      .map((f) => ({ script: `${fmt}/${f}`, name: f.replace(".py", "") }));
  }
  return tree;
}

// ── Generate examples ────────────────────────────────────────────────

/** Recursively find office files in dir and move them to OUTPUT_DIR. */
function collectOfficeFiles(dir: string): string[] {
  const found: string[] = [];
  for (const entry of readdirSync(dir)) {
    const full = resolve(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) {
      found.push(...collectOfficeFiles(full));
    } else if (OFFICE_EXTS.some((ext) => entry.endsWith(ext))) {
      if (dir !== OUTPUT_DIR) {
        const dest = resolve(OUTPUT_DIR, entry);
        renameSync(full, dest);
      }
      found.push(entry);
    }
  }
  return found;
}

/** List office files currently in OUTPUT_DIR (flat). */
function listOutputFiles(): Set<string> {
  if (!existsSync(OUTPUT_DIR)) return new Set();
  return new Set(
    readdirSync(OUTPUT_DIR).filter((f) =>
      OFFICE_EXTS.some((ext) => f.endsWith(ext)),
    ),
  );
}

async function generateAll() {
  // Clean output so before/after diff is accurate per-script
  if (existsSync(OUTPUT_DIR)) rmSync(OUTPUT_DIR, { recursive: true });
  mkdirSync(OUTPUT_DIR, { recursive: true });
  const examples = scanExamples();
  const results: {
    script: string;
    ok: boolean;
    files: string[];
    error?: string;
  }[] = [];

  for (const [_fmt, scripts] of Object.entries(examples)) {
    for (const { script } of scripts) {
      const scriptPath = resolve(ROOT, script);
      const before = listOutputFiles();
      try {
        // cwd=OUTPUT_DIR so .save("file.ext") lands there.
        // --project so uv finds the pyproject.toml.
        // TMPDIR so xlsx scripts' tempfile.mkdtemp lands here too.
        const proc = Bun.spawn(
          ["uv", "run", "--project", PY_PROJECT, "python", scriptPath],
          {
            cwd: OUTPUT_DIR,
            env: { ...process.env, TMPDIR: OUTPUT_DIR },
            stdout: "pipe",
            stderr: "pipe",
          },
        );
        const exitCode = await proc.exited;
        if (exitCode !== 0) {
          const stderr = await new Response(proc.stderr).text();
          results.push({
            script,
            ok: false,
            files: [],
            error: stderr.slice(-500),
          });
          continue;
        }
        // Hoist files from tempdir subdirs (xlsx scripts)
        collectOfficeFiles(OUTPUT_DIR);
        // Diff to find only the new files this script created
        const after = listOutputFiles();
        const newFiles = [...after].filter((f) => !before.has(f));
        results.push({ script, ok: true, files: newFiles });
      } catch (e: unknown) {
        const msg = e instanceof Error ? e.message : String(e);
        results.push({ script, ok: false, files: [], error: msg });
      }
    }
  }

  // Deduplicated list of all generated files
  const allFiles = [...new Set(results.flatMap((r) => r.files))].sort();
  return { results, files: allFiles };
}

// ── Server ───────────────────────────────────────────────────────────
const server = Bun.serve({
  port: PORT,
  routes: {
    "/": homepage,
  },
  async fetch(req) {
    const url = new URL(req.url);

    // API: list examples
    if (url.pathname === "/api/examples") {
      return Response.json(scanExamples());
    }

    // API: generate all
    if (url.pathname === "/api/generate" && req.method === "POST") {
      const result = await generateAll();
      return Response.json(result);
    }

    // API: list generated files
    if (url.pathname === "/api/files") {
      if (!existsSync(OUTPUT_DIR)) return Response.json([]);
      const files = readdirSync(OUTPUT_DIR).filter((f) =>
        OFFICE_EXTS.some((ext) => f.endsWith(ext)),
      );
      return Response.json(files.sort());
    }

    // API: serve a generated file
    if (url.pathname.startsWith("/api/files/")) {
      const name = decodeURIComponent(url.pathname.slice("/api/files/".length));
      if (name.includes("..") || name.includes("/")) {
        return new Response("Forbidden", { status: 403 });
      }
      const filePath = resolve(OUTPUT_DIR, name);
      if (!existsSync(filePath)) {
        return new Response("Not found", { status: 404 });
      }
      return new Response(Bun.file(filePath));
    }

    // Serve viewer preview assets
    if (url.pathname.startsWith("/preview/")) {
      const name = url.pathname.slice("/preview/".length);
      if (name.includes(".."))
        return new Response("Forbidden", { status: 403 });
      const filePath = resolve(PREVIEW_DIR, name);
      if (!existsSync(filePath)) {
        return new Response("Not found", { status: 404 });
      }
      const headers: Record<string, string> = {};
      if (name.endsWith(".wasm")) headers["Content-Type"] = "application/wasm";
      else if (name.endsWith(".js"))
        headers["Content-Type"] = "application/javascript";
      return new Response(Bun.file(filePath), { headers });
    }

    return new Response("Not found", { status: 404 });
  },
});

console.log(`Gallery server running at http://localhost:${server.port}`);
