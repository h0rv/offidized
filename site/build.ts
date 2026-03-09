import { mkdir } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";

const ROOT = resolve(import.meta.dir, "..");
const DIST = resolve(import.meta.dir, "dist");

const copies = [
  ["site/index.html", "index.html"],
  ["site/styles.css", "styles.css"],
  [
    "crates/offidized-docview/pkg/offidized_docview_bg.wasm",
    "assets/wasm/offidized_docview_bg.wasm",
  ],
  [
    "crates/offidized-pptview/pkg/offidized_pptview_bg.wasm",
    "assets/wasm/offidized_pptview_bg.wasm",
  ],
  [
    "crates/offidized-xlview/pkg/offidized_xlview_bg.wasm",
    "assets/wasm/offidized_xlview_bg.wasm",
  ],
  [
    "crates/offidized-wasm/pkg/offidized_wasm_bg.wasm",
    "assets/wasm/offidized_wasm_bg.wasm",
  ],
  [
    "crates/offidized-docview/node_modules/canvaskit-wasm/bin/canvaskit.wasm",
    "assets/vendor/canvaskit.wasm",
  ],
  [
    "crates/offidized-py/examples/gallery/output/07_contract.docx",
    "assets/samples/07_contract.docx",
  ],
  [
    "crates/offidized-py/examples/gallery/output/02_tables.docx",
    "assets/samples/02_tables.docx",
  ],
  [
    "crates/offidized-py/examples/gallery/output/07_pitch_deck.pptx",
    "assets/samples/07_pitch_deck.pptx",
  ],
  [
    "crates/offidized-py/examples/gallery/output/04_rich_text_formatting.pptx",
    "assets/samples/04_rich_text_formatting.pptx",
  ],
  [
    "crates/offidized-py/examples/gallery/output/meridian_portfolio.xlsx",
    "assets/samples/meridian_portfolio.xlsx",
  ],
] as const;

async function copyRelative(from: string, to: string): Promise<void> {
  const destination = join(DIST, to);
  await mkdir(dirname(destination), { recursive: true });
  await Bun.write(destination, Bun.file(join(ROOT, from)));
}

const result = await Bun.build({
  entrypoints: [resolve(import.meta.dir, "app.ts")],
  outdir: DIST,
  target: "browser",
  naming: "[name].js",
  minify: false,
  sourcemap: "external",
  splitting: false,
});

if (!result.success) {
  for (const log of result.logs) {
    console.error(log);
  }
  process.exit(1);
}

for (const [from, to] of copies) {
  await copyRelative(from, to);
}

console.log(`Built static demo site in ${DIST}`);
