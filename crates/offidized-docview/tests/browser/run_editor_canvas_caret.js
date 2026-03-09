/* global Bun, Response, URL, setTimeout */

const ROOT = new URL("../../", import.meta.url).pathname;

const CHROME_BIN =
  Bun.env.CHROME_BIN ??
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const PORT = Bun.env.EDITOR_CARET_PORT ?? "33134";

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForServerReady(timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const res = await fetch(`http://127.0.0.1:${PORT}/editor`);
      if (res.ok) return;
    } catch {
      // keep polling
    }
    await sleep(150);
  }
  throw new Error(`Timed out waiting for local dev server on :${PORT}`);
}

function readDataAttr(dom, name) {
  const regex = new RegExp(`${name}="([^"]*)"`);
  return dom.match(regex)?.[1] ?? null;
}

function parseSummary(dom) {
  const done = readDataAttr(dom, "data-editor-caret-done");
  const status = readDataAttr(dom, "data-editor-caret-status");
  const anomalies = Number.parseInt(
    readDataAttr(dom, "data-editor-caret-anomalies") ?? "-1",
    10,
  );
  const stagnant = Number.parseInt(
    readDataAttr(dom, "data-editor-caret-stagnant") ?? "-1",
    10,
  );
  const ySpan = Number.parseFloat(
    readDataAttr(dom, "data-editor-caret-y-span") ?? "-1",
  );
  const finalTop = Number.parseFloat(
    readDataAttr(dom, "data-editor-caret-final-top") ?? "NaN",
  );
  const textHead = readDataAttr(dom, "data-editor-caret-text-head") ?? "";
  const textTail = readDataAttr(dom, "data-editor-caret-text-tail") ?? "";
  return {
    done,
    status,
    anomalies,
    stagnant,
    ySpan,
    finalTop,
    textHead,
    textTail,
  };
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
  const TARGET_URL = `http://127.0.0.1:${PORT}/editor?renderer=canvas&autorun=1`;

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
      throw new Error(
        `Editor autorun did not finish. status=${parsed.status ?? "n/a"}`,
      );
    }
    if (parsed.status !== "complete") {
      throw new Error(
        `Editor autorun status is '${parsed.status ?? "n/a"}' (expected 'complete')`,
      );
    }
    if (parsed.anomalies > 0) {
      throw new Error(
        `Editor caret anomalies detected: ${parsed.anomalies} (expected 0), stagnant=${parsed.stagnant}, ySpan=${parsed.ySpan}, finalTop=${parsed.finalTop}, textHead='${parsed.textHead}', textTail='${parsed.textTail}'`,
      );
    }
    if (parsed.stagnant > 4) {
      throw new Error(
        `Editor caret was stagnant too often (${parsed.stagnant} stalled moves)`,
      );
    }
    if (!Number.isFinite(parsed.ySpan) || parsed.ySpan < 2) {
      throw new Error(
        `Editor cursor Y span too small (${parsed.ySpan}); cursor may be pinned`,
      );
    }
    if (!Number.isFinite(parsed.finalTop)) {
      throw new Error("Editor final caret top is not finite");
    }

    console.log(
      `editor-caret OK: anomalies=${parsed.anomalies} stagnant=${parsed.stagnant} ySpan=${parsed.ySpan.toFixed(2)} finalTop=${parsed.finalTop.toFixed(2)}`,
    );
  } finally {
    server.kill();
    await server.exited;
  }
}

await run();
