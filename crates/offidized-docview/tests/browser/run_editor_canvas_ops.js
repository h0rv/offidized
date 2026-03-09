/* global Bun, Response, URL, setTimeout */

const ROOT = new URL("../../", import.meta.url).pathname;

const CHROME_BIN =
  Bun.env.CHROME_BIN ??
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const RENDERER = Bun.env.EDITOR_OPS_RENDERER ?? "canvas";

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForServerReady(server, port, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (server.exitCode != null) {
      throw new Error(
        `Server exited before becoming ready (code=${server.exitCode})`,
      );
    }
    try {
      const res = await fetch(`http://127.0.0.1:${port}/editor`);
      if (res.ok) return;
    } catch {
      // keep polling
    }
    await sleep(150);
  }
  throw new Error(`Timed out waiting for local dev server on :${port}`);
}

function candidatePorts() {
  if (Bun.env.EDITOR_OPS_PORT) {
    const parsed = Number.parseInt(Bun.env.EDITOR_OPS_PORT, 10);
    if (Number.isFinite(parsed)) return [parsed];
  }
  return Array.from(
    { length: 12 },
    (_, i) => 30000 + Math.floor(Math.random() * 30000) + i,
  );
}

async function startServerWithRetry() {
  let lastError = null;
  for (const port of candidatePorts()) {
    const server = Bun.spawn({
      cmd: ["bun", "run", "serve.ts"],
      cwd: ROOT,
      env: {
        ...Bun.env,
        PORT: String(port),
      },
      stdout: "pipe",
      stderr: "pipe",
    });
    try {
      await waitForServerReady(server, port, 15_000);
      return { port, server };
    } catch (err) {
      if (server.exitCode == null) {
        server.kill();
      }
      await server.exited;
      const stdout = await new Response(server.stdout).text();
      const stderr = await new Response(server.stderr).text();
      const combinedMessage = `${String(err)}\n${stderr}`;
      if (combinedMessage.includes("EADDRINUSE")) {
        lastError = new Error(
          `Port ${port} already in use\nserver stdout:\n${stdout.slice(0, 1000)}\nserver stderr:\n${stderr.slice(0, 1000)}`,
        );
        continue;
      }
      throw new Error(
        `${String(err)}\nserver stdout:\n${stdout.slice(0, 2000)}\nserver stderr:\n${stderr.slice(0, 2000)}`,
        { cause: err },
      );
    }
  }
  throw lastError ?? new Error("failed to start local demo server");
}

function readDataAttr(dom, name) {
  const regex = new RegExp(`${name}="([^"]*)"`);
  return dom.match(regex)?.[1] ?? null;
}

function parseSummary(dom) {
  return {
    done: readDataAttr(dom, "data-editor-ops-done"),
    status: readDataAttr(dom, "data-editor-ops-status"),
    textAfterEdits: readDataAttr(dom, "data-editor-ops-text-after-edits") ?? "",
    undoText: readDataAttr(dom, "data-editor-ops-undo-text") ?? "",
    selectDeleteText:
      readDataAttr(dom, "data-editor-ops-select-delete-text") ?? "",
    multilineDeleteText:
      readDataAttr(dom, "data-editor-ops-multiline-delete-text") ?? "",
    shiftDeleteText:
      readDataAttr(dom, "data-editor-ops-shift-delete-text") ?? "",
    crossParagraphHtmlSelectionText:
      readDataAttr(
        dom,
        "data-editor-ops-cross-paragraph-html-selection-text",
      ) ?? "",
    dragCopiedLen: Number.parseInt(
      readDataAttr(dom, "data-editor-ops-drag-copied-len") ?? "-1",
      10,
    ),
    dragBeforeLen: Number.parseInt(
      readDataAttr(dom, "data-editor-ops-drag-before-len") ?? "-1",
      10,
    ),
    dragAfterLen: Number.parseInt(
      readDataAttr(dom, "data-editor-ops-drag-after-len") ?? "-1",
      10,
    ),
    dragCopiedText: readDataAttr(dom, "data-editor-ops-drag-copied-text") ?? "",
    dragAfterText: readDataAttr(dom, "data-editor-ops-drag-after-text") ?? "",
    reverseDragCopiedText:
      readDataAttr(dom, "data-editor-ops-reverse-drag-copied-text") ?? "",
    reverseDragAfterText:
      readDataAttr(dom, "data-editor-ops-reverse-drag-after-text") ?? "",
    tableInsertOk: readDataAttr(dom, "data-editor-ops-table-insert-ok") ?? "",
    tableStartCell: readDataAttr(dom, "data-editor-ops-table-start-cell") ?? "",
    tableAfterTabCell:
      readDataAttr(dom, "data-editor-ops-table-after-tab-cell") ?? "",
    tableAfterRowInsert:
      readDataAttr(dom, "data-editor-ops-table-after-row-insert") ?? "",
    tableAfterColInsert:
      readDataAttr(dom, "data-editor-ops-table-after-col-insert") ?? "",
    tableFinalShape:
      readDataAttr(dom, "data-editor-ops-table-final-shape") ?? "",
    tableCellA1: readDataAttr(dom, "data-editor-ops-table-cell-a1") ?? "",
    tableCellB1: readDataAttr(dom, "data-editor-ops-table-cell-b1") ?? "",
    tableCellCopyText:
      readDataAttr(dom, "data-editor-ops-table-cell-copy-text") ?? "",
    tableCellCutText:
      readDataAttr(dom, "data-editor-ops-table-cell-cut-text") ?? "",
    tableCellDeleteText:
      readDataAttr(dom, "data-editor-ops-table-cell-delete-text") ?? "",
    tableCellPasteText:
      readDataAttr(dom, "data-editor-ops-table-cell-paste-text") ?? "",
    imageInsertOk: readDataAttr(dom, "data-editor-ops-image-insert-ok") ?? "",
    imageCount: Number.parseInt(
      readDataAttr(dom, "data-editor-ops-image-count") ?? "-1",
      10,
    ),
    inlineImageRunCount: Number.parseInt(
      readDataAttr(dom, "data-editor-ops-inline-image-run-count") ?? "-1",
      10,
    ),
    imageContentType:
      readDataAttr(dom, "data-editor-ops-image-content-type") ?? "",
    imageDomCount: Number.parseInt(
      readDataAttr(dom, "data-editor-ops-image-dom-count") ?? "-1",
      10,
    ),
    copiedText: readDataAttr(dom, "data-editor-ops-copied-text") ?? "",
    cutText: readDataAttr(dom, "data-editor-ops-cut-text") ?? "",
    richCopyHtmlOk:
      readDataAttr(dom, "data-editor-ops-rich-copy-html-ok") ?? "",
    richCutHtmlOk: readDataAttr(dom, "data-editor-ops-rich-cut-html-ok") ?? "",
    richPasteOk: readDataAttr(dom, "data-editor-ops-rich-paste-ok") ?? "",
    richPasteListOk:
      readDataAttr(dom, "data-editor-ops-rich-paste-list-ok") ?? "",
    backwardSelectionOk:
      readDataAttr(dom, "data-editor-ops-backward-selection-ok") ?? "",
    notes: readDataAttr(dom, "data-editor-ops-notes") ?? "",
  };
}

async function run() {
  const { port: requestedPort, server } = await startServerWithRetry();
  const PORT = String(requestedPort);
  const TARGET_URL = `http://127.0.0.1:${PORT}/editor?renderer=${RENDERER}&autorun=ops`;

  try {
    const chrome = Bun.spawn({
      cmd: [
        CHROME_BIN,
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        "--window-size=1800,2600",
        "--virtual-time-budget=40000",
        "--dump-dom",
        TARGET_URL,
      ],
      stdout: "pipe",
      stderr: "pipe",
    });

    const dom = await new Response(chrome.stdout).text();
    const chromeErr = await new Response(chrome.stderr).text();
    const chromeExit = await chrome.exited;
    if (chromeExit !== 0) {
      throw new Error(
        `Chrome exited with ${chromeExit}\n${chromeErr.slice(0, 2000)}`,
      );
    }

    const parsed = parseSummary(dom);
    if (parsed.done !== "1") {
      throw new Error(`Editor ops autorun did not finish.`);
    }
    if (parsed.status !== "complete") {
      throw new Error(
        `Editor ops autorun failed. status=${parsed.status ?? "n/a"} notes=${parsed.notes}`,
      );
    }
    if (parsed.richCopyHtmlOk !== "1") {
      throw new Error(
        `Rich copy html assertion failed. got='${parsed.richCopyHtmlOk}'`,
      );
    }
    if (parsed.richCutHtmlOk !== "1") {
      throw new Error(
        `Rich cut html assertion failed. got='${parsed.richCutHtmlOk}'`,
      );
    }
    if (parsed.richPasteOk !== "1") {
      throw new Error(
        `Rich paste assertion failed. got='${parsed.richPasteOk}'`,
      );
    }
    if (parsed.richPasteListOk !== "1") {
      throw new Error(
        `Rich paste list assertion failed. got='${parsed.richPasteListOk}'`,
      );
    }
    if (parsed.tableInsertOk !== "1") {
      throw new Error(
        `Table insert did not report success. got='${parsed.tableInsertOk}'`,
      );
    }
    if (parsed.tableStartCell !== "0:0") {
      throw new Error(
        `Unexpected starting table cell '${parsed.tableStartCell}'`,
      );
    }
    if (parsed.tableAfterTabCell !== "0:1") {
      throw new Error(
        `Unexpected table Tab target '${parsed.tableAfterTabCell}'`,
      );
    }
    if (parsed.tableAfterRowInsert !== "1:1:4:3") {
      throw new Error(
        `Unexpected table row insert state '${parsed.tableAfterRowInsert}'`,
      );
    }
    if (parsed.tableAfterColInsert !== "1:2:4:4") {
      throw new Error(
        `Unexpected table col insert state '${parsed.tableAfterColInsert}'`,
      );
    }
    if (parsed.tableFinalShape !== "3x3") {
      throw new Error(
        `Unexpected final table shape '${parsed.tableFinalShape}'`,
      );
    }
    if (parsed.tableCellA1 !== "A1" || parsed.tableCellB1 !== "B1") {
      throw new Error(
        `Unexpected table cell text A1='${parsed.tableCellA1}' B1='${parsed.tableCellB1}'`,
      );
    }
    if (parsed.tableCellPasteText !== "Alpha") {
      throw new Error(
        `Unexpected table paste text '${parsed.tableCellPasteText}'`,
      );
    }
    if (parsed.tableCellCopyText !== "lph") {
      throw new Error(
        `Unexpected table copy text '${parsed.tableCellCopyText}'`,
      );
    }
    if (parsed.tableCellCutText !== "lph") {
      throw new Error(`Unexpected table cut text '${parsed.tableCellCutText}'`);
    }
    if (parsed.tableCellDeleteText !== "A") {
      throw new Error(
        `Unexpected table delete text '${parsed.tableCellDeleteText}'`,
      );
    }
    if (parsed.imageInsertOk !== "1") {
      throw new Error(
        `Image insert did not report success. got='${parsed.imageInsertOk}'`,
      );
    }
    if (parsed.imageCount !== 1 || parsed.inlineImageRunCount !== 1) {
      throw new Error(
        `Unexpected image model counts images=${parsed.imageCount} runs=${parsed.inlineImageRunCount}`,
      );
    }
    if (parsed.imageContentType !== "image/png") {
      throw new Error(
        `Unexpected image content type '${parsed.imageContentType}'`,
      );
    }
    if (RENDERER === "canvas") {
      if (parsed.imageDomCount !== 0) {
        throw new Error(
          `Canvas renderer should not expose inline image DOM nodes, got ${parsed.imageDomCount}`,
        );
      }
      if (parsed.textAfterEdits !== "1abXc") {
        throw new Error(
          `Unexpected text-after-edits '${parsed.textAfterEdits}' (expected '1abXc')`,
        );
      }
      if (parsed.copiedText !== "2") {
        throw new Error(
          `Unexpected copied text '${parsed.copiedText}' (expected '2')`,
        );
      }
      if (parsed.cutText !== "2") {
        throw new Error(
          `Unexpected cut text '${parsed.cutText}' (expected '2')`,
        );
      }
      if (parsed.undoText === "undo") {
        throw new Error("Undo did not mutate document state");
      }
      if (parsed.selectDeleteText !== "") {
        throw new Error(
          `Select-all delete text should be empty, got '${parsed.selectDeleteText}'`,
        );
      }
      if (parsed.multilineDeleteText !== "") {
        throw new Error(
          `Unexpected multiline delete text '${parsed.multilineDeleteText}' (expected empty text)`,
        );
      }
      if (parsed.shiftDeleteText !== "abc\nde") {
        throw new Error(
          `Unexpected shift delete text '${parsed.shiftDeleteText}' (expected 'abc\\nde')`,
        );
      }
      if (parsed.dragCopiedLen < 20) {
        throw new Error(
          `Unexpected drag copied length '${parsed.dragCopiedLen}' (expected >= 20)`,
        );
      }
      if (!parsed.dragCopiedText.includes("\n")) {
        throw new Error(
          `Forward drag should copy multiline text, got '${parsed.dragCopiedText}'`,
        );
      }
      if (
        parsed.dragBeforeLen < 0 ||
        parsed.dragAfterLen < 0 ||
        parsed.dragAfterLen >= parsed.dragBeforeLen
      ) {
        throw new Error(
          `Drag delete length check failed. before=${parsed.dragBeforeLen} after=${parsed.dragAfterLen}`,
        );
      }
      if (parsed.reverseDragCopiedText !== parsed.dragCopiedText) {
        throw new Error(
          `Reverse drag copied text mismatch. forward='${parsed.dragCopiedText}' reverse='${parsed.reverseDragCopiedText}'`,
        );
      }
      if (parsed.reverseDragAfterText !== parsed.dragAfterText) {
        throw new Error(
          `Reverse drag delete mismatch. forward='${parsed.dragAfterText}' reverse='${parsed.reverseDragAfterText}'`,
        );
      }
    } else {
      if (parsed.copiedText !== "Rich") {
        throw new Error(
          `Unexpected HTML copied text '${parsed.copiedText}' (expected 'Rich')`,
        );
      }
      if (parsed.cutText !== "Rich") {
        throw new Error(
          `Unexpected HTML cut text '${parsed.cutText}' (expected 'Rich')`,
        );
      }
      if (parsed.backwardSelectionOk !== "1") {
        throw new Error(
          `Backward HTML selection assertion failed. got='${parsed.backwardSelectionOk}'`,
        );
      }
      if (parsed.crossParagraphHtmlSelectionText !== "pha\nbravo\nchar") {
        throw new Error(
          `Unexpected HTML cross-paragraph selection text '${parsed.crossParagraphHtmlSelectionText}'`,
        );
      }
      if (parsed.imageDomCount !== 1) {
        throw new Error(
          `HTML renderer should expose one inline image DOM node, got ${parsed.imageDomCount}`,
        );
      }
    }

    console.log(
      `editor-ops OK (${RENDERER}): textAfterEdits='${parsed.textAfterEdits}' undo='${parsed.undoText}' selectDelete='${parsed.selectDeleteText}' multilineDelete='${parsed.multilineDeleteText}' shiftDelete='${parsed.shiftDeleteText}' drag=${parsed.dragBeforeLen}->${parsed.dragAfterLen} copy='${parsed.copiedText}' cut='${parsed.cutText}' richCopy=${parsed.richCopyHtmlOk} richCut=${parsed.richCutHtmlOk}`,
    );
  } finally {
    server.kill();
    await server.exited;
  }
}

await run();
