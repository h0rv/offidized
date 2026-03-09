// WASM init helper for headless E2E tests (Bun, no browser).
//
// Loads the WASM binary from disk and initialises with `initSync`
// so tests can run without a server or fetch polyfill.

import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { initSync, DocEdit } from "../../pkg/offidized_docview.js";

const wasmPath = resolve(
  import.meta.dir,
  "../../pkg/offidized_docview_bg.wasm",
);
const wasmBytes = readFileSync(wasmPath);
initSync({ module: wasmBytes });

export { DocEdit };

/** Shorthand: view model body item cast to paragraph. */
export interface VmRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
  highlight?: string;
  hyperlink?: string;
  hasBreak?: boolean;
  hasTab?: boolean;
  inlineImage?: {
    imageIndex?: number;
    widthPt?: number;
    heightPt?: number;
    name?: string;
    description?: string;
  };
  footnoteRef?: number;
  endnoteRef?: number;
}

export interface VmParagraph {
  type: "paragraph";
  runs: VmRun[];
  headingLevel?: number;
  alignment?: string;
  spacingBeforePt?: number;
  spacingAfterPt?: number;
  lineSpacing?: {
    value?: number;
    rule?: string;
  };
  indents?: {
    leftPt?: number;
    firstLinePt?: number;
  };
  numbering?: VmNumbering;
}

export interface VmBodyItem {
  type: string;
  runs?: VmRun[];
}

export interface DocViewModel {
  body: VmBodyItem[];
}

export interface VmNumbering {
  numId: number;
  level: number;
  format: string;
  text: string;
}

/** Get the text of paragraph at a specific index. */
export function paraText(editor: DocEdit, index: number): string {
  const model = vm(editor);
  const para = model.body[index];
  if (!para?.runs) return "";
  return para.runs.map((r) => r.text).join("");
}

/** Get the runs of paragraph at a specific index. */
export function paraRuns(editor: DocEdit, index: number): VmRun[] {
  const model = vm(editor);
  const para = model.body[index];
  return (para?.runs ?? []) as VmRun[];
}

/** Get a paragraph model at a specific index. */
export function paraModel(editor: DocEdit, index: number): VmParagraph | null {
  const model = vm(editor);
  const para = model.body[index];
  if (!para || para.type !== "paragraph") return null;
  return para as VmParagraph;
}

/** Get the view model from a DocEdit instance, typed. */
export function vm(editor: DocEdit): DocViewModel {
  return editor.viewModel() as DocViewModel;
}

/** Get the text of the first paragraph in the view model. */
export function firstParaText(editor: DocEdit): string {
  const model = vm(editor);
  const para = model.body[0];
  if (!para?.runs) return "";
  return para.runs.map((r) => r.text).join("");
}

/** Apply an intent and return the updated view model. */
export function applyAndVm(
  editor: DocEdit,
  intent: Record<string, unknown>,
): DocViewModel {
  editor.applyIntent(JSON.stringify(intent));
  return vm(editor);
}
