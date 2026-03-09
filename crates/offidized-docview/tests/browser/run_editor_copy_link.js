/* global Bun, Response, URL, setTimeout */

const ROOT = new URL("../../", import.meta.url).pathname;

const CHROME_BIN =
  Bun.env.CHROME_BIN ??
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

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
  throw new Error(`Timed out waiting for local demo server on :${port}`);
}

function candidatePorts() {
  if (Bun.env.EDITOR_COPY_LINK_PORT) {
    const parsed = Number.parseInt(Bun.env.EDITOR_COPY_LINK_PORT, 10);
    if (Number.isFinite(parsed)) return [parsed];
  }
  return Array.from(
    { length: 12 },
    (_, i) => 30000 + Math.floor(Math.random() * 30000) + i,
  );
}

async function startServer() {
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
  const raw = dom.match(regex)?.[1] ?? null;
  return raw?.replaceAll("&amp;", "&") ?? null;
}

async function run() {
  const { port, server } = await startServer();
  const room = `copy-link-room-${Date.now().toString(36)}`;
  const ws = "ws://127.0.0.1:9999";
  const targetUrl = `http://127.0.0.1:${port}/editor?renderer=canvas&room=${encodeURIComponent(room)}&ws=${encodeURIComponent(ws)}&autorun=copy-link`;

  try {
    const chrome = Bun.spawn({
      cmd: [
        CHROME_BIN,
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        "--window-size=1600,2200",
        "--virtual-time-budget=15000",
        "--dump-dom",
        targetUrl,
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

    const done = readDataAttr(dom, "data-editor-copy-link-done");
    const status = readDataAttr(dom, "data-editor-copy-link-status");
    const copiedLink = readDataAttr(dom, "data-editor-copy-link-copied-link");
    const copiedRoom = readDataAttr(dom, "data-editor-copy-link-room");
    const notes = readDataAttr(dom, "data-editor-copy-link-notes") ?? "";

    if (done !== "1") {
      throw new Error(`Copy-link autorun did not finish. status=${status}`);
    }
    if (status !== "complete") {
      throw new Error(
        `Copy-link autorun failed. status=${status} notes=${notes}`,
      );
    }
    if (!copiedLink) {
      throw new Error("Copy-link autorun did not capture a link");
    }
    const copiedUrl = new URL(copiedLink);
    if (copiedUrl.searchParams.get("room") !== room) {
      throw new Error(
        `Copied link room mismatch. expected='${room}' got='${copiedUrl.searchParams.get("room") ?? ""}'`,
      );
    }
    if (copiedUrl.searchParams.get("ws") !== ws) {
      throw new Error(
        `Copied link ws mismatch. expected='${ws}' got='${copiedUrl.searchParams.get("ws") ?? ""}'`,
      );
    }
    if (copiedRoom !== room) {
      throw new Error(
        `Copy-link room dataset mismatch. expected='${room}' got='${copiedRoom ?? ""}'`,
      );
    }

    console.log(`editor-copy-link OK: room='${room}' link='${copiedLink}'`);
  } finally {
    server.kill();
    await server.exited;
  }
}

await run();
