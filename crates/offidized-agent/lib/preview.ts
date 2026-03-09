import { resolve } from "path";
import { existsSync } from "fs";

export const PREVIEW_DIR = resolve(import.meta.dir, "..", ".preview");
const VIEWER_ROOTS = {
  xlsx: resolve(import.meta.dir, "../../offidized-xlview"),
  docx: resolve(import.meta.dir, "../../../offidized-docview"),
  pptx: resolve(import.meta.dir, "../../../offidized-pptview"),
};

export let previewReady = false;

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
  if (existsSync(jsPath)) {
    entrypoints.push(jsPath);
    const wasmName =
      ext === "xlsx"
        ? "offidized_xlview_bg.wasm"
        : ext === "docx"
          ? "offidized_docview_bg.wasm"
          : "offidized_pptview_bg.wasm";
    wasmCopies.push(
      Bun.write(
        resolve(PREVIEW_DIR, wasmName),
        Bun.file(resolve(root, "pkg", wasmName)),
      ).then(() => {}),
    );
  }
}

if (entrypoints.length > 0) {
  try {
    const result = await Bun.build({
      entrypoints,
      outdir: PREVIEW_DIR,
      target: "browser",
      naming: "[name].[ext]",
    });

    if (result.success) {
      await Promise.all(wasmCopies);
      previewReady = true;
      console.log(`Preview assets built (${entrypoints.length} viewers)`);
    } else {
      console.error("Preview build failed:", result.logs);
    }
  } catch (e) {
    console.error("Preview build skipped:", e);
  }
} else {
  console.log("Preview disabled — no viewer packages found");
}
