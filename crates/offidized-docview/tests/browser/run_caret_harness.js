/* global Bun, Response, URL, setTimeout */

const ROOT = new URL("../../", import.meta.url).pathname;

const CHROME_BIN =
  Bun.env.CHROME_BIN ??
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const PORT = Bun.env.CARET_HARNESS_PORT ?? "33133";
const HARNESS_URL = `http://127.0.0.1:${PORT}/harness/caret?autorun=1`;

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForServerReady(timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`http://127.0.0.1:${PORT}/harness/caret`);
      if (res.ok) return;
    } catch {
      // keep polling
    }
    await sleep(150);
  }
  throw new Error("Timed out waiting for local dev server on :3033");
}

function readDataAttr(dom, name) {
  const regex = new RegExp(`${name}="([^"]*)"`);
  return dom.match(regex)?.[1] ?? null;
}

function parseSummary(dom) {
  const done = readDataAttr(dom, "data-caret-harness-done");
  const status = readDataAttr(dom, "data-caret-harness-status");
  const anomalies = Number.parseInt(
    readDataAttr(dom, "data-caret-harness-active-window-anomalies") ?? "-1",
    10,
  );
  const focused = Number.parseInt(
    readDataAttr(dom, "data-caret-harness-focused-cursor-samples") ?? "-1",
    10,
  );
  const drawCalls = Number.parseInt(
    readDataAttr(dom, "data-caret-harness-draw-calls") ?? "-1",
    10,
  );
  const cursorYSpanPx = Number.parseFloat(
    readDataAttr(dom, "data-caret-harness-cursor-y-span-px") ?? "-1",
  );
  const finalBodyIndex = Number.parseInt(
    readDataAttr(dom, "data-caret-harness-final-body-index") ?? "-1",
    10,
  );
  const finalCharOffset = Number.parseInt(
    readDataAttr(dom, "data-caret-harness-final-char-offset") ?? "-1",
    10,
  );
  return {
    done,
    status,
    anomalies,
    focused,
    drawCalls,
    cursorYSpanPx,
    finalBodyIndex,
    finalCharOffset,
  };
}

function extractElementText(dom, id) {
  const regex = new RegExp(`<[^>]*id="${id}"[^>]*>([\\s\\S]*?)<\\/[^>]+>`, "i");
  const raw = dom.match(regex)?.[1] ?? "";
  return raw.replace(/<[^>]+>/g, "").trim();
}

async function run() {
  const server = Bun.spawn({
    cmd: ["bun", "run", "serve.ts"],
    cwd: ROOT,
    env: {
      ...Bun.env,
      PORT,
    },
    stdout: "pipe",
    stderr: "pipe",
  });

  try {
    try {
      await waitForServerReady(15_000);
    } catch (err) {
      const stdout = await new Response(server.stdout).text();
      const stderr = await new Response(server.stderr).text();
      throw new Error(
        `${String(err)}\nserver stdout:\n${stdout.slice(0, 2000)}\nserver stderr:\n${stderr.slice(0, 2000)}`,
        { cause: err },
      );
    }

    const chrome = Bun.spawn({
      cmd: [
        CHROME_BIN,
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        "--virtual-time-budget=35000",
        "--dump-dom",
        HARNESS_URL,
      ],
      stdout: "pipe",
      stderr: "pipe",
    });

    const domPromise = new Response(chrome.stdout).text();
    const stderrPromise = new Response(chrome.stderr).text();
    const chromeExit = await chrome.exited;
    const dom = await domPromise;
    const chromeErr = await stderrPromise;

    if (chromeExit !== 0) {
      throw new Error(
        `Chrome exited with ${chromeExit}\n${chromeErr.slice(0, 2000)}`,
      );
    }

    const parsed = parseSummary(dom);
    const statusText = extractElementText(dom, "status");
    const summaryText = extractElementText(dom, "summary");
    if (parsed.done !== "1") {
      throw new Error(
        `Harness did not finish autorun.\nstatusAttr=${parsed.status ?? "n/a"}\nstatusText=${statusText}\nsummaryText=${summaryText}\nDOM(snippet)=${dom.slice(0, 1200)}`,
      );
    }
    if (parsed.status !== "scenario complete") {
      throw new Error(
        `Unexpected harness status '${parsed.status ?? "n/a"}' (expected 'scenario complete')`,
      );
    }
    if (parsed.focused <= 0) {
      throw new Error(
        `No focused+cursor samples captured (focused=${parsed.focused})`,
      );
    }
    if (parsed.anomalies > 0) {
      throw new Error(
        `Caret anomalies detected: ${parsed.anomalies} (expected 0)`,
      );
    }
    if (parsed.finalBodyIndex < 0 || parsed.finalCharOffset < 0) {
      throw new Error(
        `Invalid final cursor position body=${parsed.finalBodyIndex} off=${parsed.finalCharOffset}`,
      );
    }
    if (parsed.finalBodyIndex === 0 && parsed.finalCharOffset === 0) {
      throw new Error("Cursor stayed at initial position (0,0) after typing");
    }
    if (!Number.isFinite(parsed.cursorYSpanPx) || parsed.cursorYSpanPx < 2) {
      throw new Error(
        `Cursor Y span too small (${parsed.cursorYSpanPx}); cursor may be pinned visually`,
      );
    }

    console.log(
      `caret-harness OK: anomalies=${parsed.anomalies} focused=${parsed.focused} drawCalls=${parsed.drawCalls} final=${parsed.finalBodyIndex}:${parsed.finalCharOffset} ySpan=${parsed.cursorYSpanPx.toFixed(2)}`,
    );
  } finally {
    server.kill();
    await server.exited;
  }
}

await run();
