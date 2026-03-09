#!/usr/bin/env node

/**
 * Post-wasm-pack patcher: updates pkg/package.json with xl-view wrapper entry points.
 * Run after `wasm-pack build --target web`.
 */

import { readFileSync, writeFileSync } from "fs";
import { join } from "path";

const pkgPath = join(import.meta.dirname, "..", "pkg", "package.json");
const pkg = JSON.parse(readFileSync(pkgPath, "utf-8"));

// Set main entry to the wrapper
pkg.main = "xl-view.js";
pkg.types = "xl-view.d.ts";

// Add exports map for both wrapper and raw WASM bindings
pkg.exports = {
  ".": {
    import: "./xl-view.js",
    types: "./xl-view.d.ts",
  },
  "./core": {
    import: "./offidized_xlview.js",
    types: "./offidized_xlview.d.ts",
  },
};

// Ensure xl-view files are in the files array
const extraFiles = ["xl-view.js", "xl-view.d.ts"];
for (const f of extraFiles) {
  if (!pkg.files.includes(f)) {
    pkg.files.push(f);
  }
}

writeFileSync(pkgPath, JSON.stringify(pkg, null, 2) + "\n");
console.log("Patched pkg/package.json with xl-view entry points.");
