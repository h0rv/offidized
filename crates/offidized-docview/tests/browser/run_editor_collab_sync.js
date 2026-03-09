/* global Bun, setTimeout, Response, URL */

const ROOT = new URL("../../", import.meta.url).pathname;

const CHROME_BIN =
  Bun.env.CHROME_BIN ??
  "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const ROOM =
  Bun.env.EDITOR_COLLAB_ROOM ??
  `collab-harness-room-${Date.now().toString(36)}`;

function defaultPort(attempt = 0) {
  if (Bun.env.EDITOR_COLLAB_PORT) return Bun.env.EDITOR_COLLAB_PORT;
  return String(30000 + Math.floor(Math.random() * 30000) + attempt);
}

async function startServerWithRetry(maxAttempts = 5) {
  let lastError = null;
  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const port = defaultPort(attempt);
    const server = Bun.spawn({
      cmd: ["bun", "run", "serve.ts"],
      cwd: ROOT,
      env: {
        ...Bun.env,
        PORT: port,
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
      if (
        combinedMessage.includes("EADDRINUSE") &&
        !Bun.env.EDITOR_COLLAB_PORT &&
        attempt + 1 < maxAttempts
      ) {
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
  throw lastError ?? new Error("failed to start local dev server");
}

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

function readDataAttr(dom, name) {
  const regex = new RegExp(`${name}="([^"]*)"`);
  return dom.match(regex)?.[1] ?? null;
}

function parseSummary(dom) {
  return {
    done: readDataAttr(dom, "data-editor-collab-done"),
    status: readDataAttr(dom, "data-editor-collab-status"),
    expected: readDataAttr(dom, "data-editor-collab-expected") ?? "",
    peerText: readDataAttr(dom, "data-editor-collab-peer-text") ?? "",
    presenceCount: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-presence-count") ?? "0",
      10,
    ),
    presenceCursorCount: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-presence-cursor-count") ?? "0",
      10,
    ),
    presenceSelectionCount: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-presence-selection-count") ?? "0",
      10,
    ),
    peerTableCellA1:
      readDataAttr(dom, "data-editor-collab-peer-table-cell-a1") ?? "",
    peerTableCellB1:
      readDataAttr(dom, "data-editor-collab-peer-table-cell-b1") ?? "",
    peerImageCount: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-peer-image-count") ?? "0",
      10,
    ),
    peerInlineImageRunCount: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-peer-inline-image-run-count") ??
        "0",
      10,
    ),
    replaceExpected:
      readDataAttr(dom, "data-editor-collab-replace-expected") ?? "",
    replacePeerText:
      readDataAttr(dom, "data-editor-collab-replace-peer-text") ?? "",
    replacePeerTableCellA1:
      readDataAttr(dom, "data-editor-collab-replace-peer-table-cell-a1") ?? "",
    replacePeerTableCellB1:
      readDataAttr(dom, "data-editor-collab-replace-peer-table-cell-b1") ?? "",
    replacePeerImageCount: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-replace-peer-image-count") ?? "0",
      10,
    ),
    replacePeerInlineImageRunCount: Number.parseInt(
      readDataAttr(
        dom,
        "data-editor-collab-replace-peer-inline-image-run-count",
      ) ?? "0",
      10,
    ),
    reconnectExpected:
      readDataAttr(dom, "data-editor-collab-reconnect-expected") ?? "",
    reconnectPeerText:
      readDataAttr(dom, "data-editor-collab-reconnect-peer-text") ?? "",
    reconnectPeerTableCellA1:
      readDataAttr(dom, "data-editor-collab-reconnect-peer-table-cell-a1") ??
      "",
    reconnectPeerTableCellB1:
      readDataAttr(dom, "data-editor-collab-reconnect-peer-table-cell-b1") ??
      "",
    reconnectPeerImageCount: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-reconnect-peer-image-count") ?? "0",
      10,
    ),
    reconnectPeerInlineImageRunCount: Number.parseInt(
      readDataAttr(
        dom,
        "data-editor-collab-reconnect-peer-inline-image-run-count",
      ) ?? "0",
      10,
    ),
    reconnectPeerDebug:
      readDataAttr(dom, "data-editor-collab-reconnect-peer-debug") ?? "",
    reconnectPeerRepairRequests: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-reconnect-peer-repair-requests") ??
        "0",
      10,
    ),
    reconnectPeerDivergenceDetected: Number.parseInt(
      readDataAttr(
        dom,
        "data-editor-collab-reconnect-peer-divergence-detected",
      ) ?? "0",
      10,
    ),
    reconnectPeerDivergenceCleared: Number.parseInt(
      readDataAttr(
        dom,
        "data-editor-collab-reconnect-peer-divergence-cleared",
      ) ?? "0",
      10,
    ),
    reconnectPeerStateRequestSend: Number.parseInt(
      readDataAttr(
        dom,
        "data-editor-collab-reconnect-peer-state-request-send",
      ) ?? "0",
      10,
    ),
    reconnectPeerSawRepairing:
      readDataAttr(dom, "data-editor-collab-reconnect-peer-saw-repairing") ??
      "0",
    reconnectPeerSawDesync:
      readDataAttr(dom, "data-editor-collab-reconnect-peer-saw-desync") ?? "0",
    presenceExpired: readDataAttr(dom, "data-editor-collab-presence-expired"),
    expiredPresenceCount: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-presence-expired-count") ?? "0",
      10,
    ),
    expiredPresenceCursorCount: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-presence-expired-cursor-count") ??
        "0",
      10,
    ),
    expiredPresenceSelectionCount: Number.parseInt(
      readDataAttr(
        dom,
        "data-editor-collab-presence-expired-selection-count",
      ) ?? "0",
      10,
    ),
    room: readDataAttr(dom, "data-editor-collab-room") ?? "",
    transport: readDataAttr(dom, "data-editor-collab-transport") ?? "",
    awarenessExpire: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-awareness-expire") ?? "0",
      10,
    ),
    awarenessByeRecv: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-awareness-bye-recv") ?? "0",
      10,
    ),
    stateRequestSend: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-state-request-send") ?? "0",
      10,
    ),
    divergenceDetected: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-divergence-detected") ?? "0",
      10,
    ),
    divergenceCleared: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-divergence-cleared") ?? "0",
      10,
    ),
    repairRequests: Number.parseInt(
      readDataAttr(dom, "data-editor-collab-repair-requests") ?? "0",
      10,
    ),
    resyncing: readDataAttr(dom, "data-editor-collab-resyncing") ?? "0",
    divergenceSuspected:
      readDataAttr(dom, "data-editor-collab-divergence-suspected") ?? "0",
    notes: readDataAttr(dom, "data-editor-collab-notes") ?? "",
  };
}

async function run() {
  const { port: PORT, server } = await startServerWithRetry();
  const TARGET_URL = `http://127.0.0.1:${PORT}/editor?renderer=canvas&room=${encodeURIComponent(ROOM)}&autorun=collab-presence-expiry`;

  try {
    const chrome = Bun.spawn({
      cmd: [
        CHROME_BIN,
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        "--window-size=1800,2600",
        "--virtual-time-budget=90000",
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
        `Collab autorun did not finish. status=${parsed.status ?? "n/a"} notes=${parsed.notes}`,
      );
    }
    if (parsed.status !== "complete") {
      throw new Error(
        `Collab autorun failed. status=${parsed.status ?? "n/a"} notes=${parsed.notes}`,
      );
    }
    if (parsed.room !== ROOM) {
      throw new Error(
        `Collab room mismatch. expected='${ROOM}' got='${parsed.room}'`,
      );
    }
    if (
      parsed.transport !== "broadcast" &&
      parsed.transport !== "websocket" &&
      parsed.transport !== "hybrid"
    ) {
      throw new Error(
        `Collab transport should be connected, got='${parsed.transport}' notes=${parsed.notes}`,
      );
    }
    if (!Number.isFinite(parsed.presenceCount) || parsed.presenceCount < 1) {
      throw new Error(
        `Collab presence missing. expected at least 1 remote overlay, got='${parsed.presenceCount}'`,
      );
    }
    if (
      !Number.isFinite(parsed.presenceCursorCount) ||
      parsed.presenceCursorCount < 1
    ) {
      throw new Error(
        `Collab remote cursor missing. expected at least 1 remote cursor, got='${parsed.presenceCursorCount}'`,
      );
    }
    if (parsed.presenceExpired !== "1") {
      throw new Error(
        `Collab stale presence did not expire. flag='${parsed.presenceExpired}' notes=${parsed.notes}`,
      );
    }
    if (
      !Number.isFinite(parsed.expiredPresenceCount) ||
      parsed.expiredPresenceCount !== 0
    ) {
      throw new Error(
        `Collab stale peer overlay still visible. expected 0 peers, got='${parsed.expiredPresenceCount}'`,
      );
    }
    if (
      !Number.isFinite(parsed.expiredPresenceCursorCount) ||
      parsed.expiredPresenceCursorCount !== 0
    ) {
      throw new Error(
        `Collab stale remote cursor still visible. expected 0 cursors, got='${parsed.expiredPresenceCursorCount}'`,
      );
    }
    if (
      !Number.isFinite(parsed.expiredPresenceSelectionCount) ||
      parsed.expiredPresenceSelectionCount !== 0
    ) {
      throw new Error(
        `Collab stale remote selection still visible. expected 0 selections, got='${parsed.expiredPresenceSelectionCount}'`,
      );
    }
    if (
      !Number.isFinite(parsed.awarenessExpire) ||
      parsed.awarenessExpire < 1
    ) {
      throw new Error(
        `Collab stale expiry should be timer-driven. expected awarenessExpire>=1, got='${parsed.awarenessExpire}'`,
      );
    }
    if (
      !Number.isFinite(parsed.awarenessByeRecv) ||
      parsed.awarenessByeRecv !== 0
    ) {
      throw new Error(
        `Collab stale expiry should not rely on bye. expected awarenessByeRecv=0, got='${parsed.awarenessByeRecv}'`,
      );
    }
    if (parsed.notes !== "ok") {
      throw new Error(
        `Collab notes mismatch. expected='ok' got='${parsed.notes}'`,
      );
    }
    if (parsed.peerTableCellA1 !== "A1" || parsed.peerTableCellB1 !== "B1") {
      throw new Error(
        `Collab table sync mismatch. expected A1/B1, got '${parsed.peerTableCellA1}'/'${parsed.peerTableCellB1}'`,
      );
    }
    if (
      !Number.isFinite(parsed.peerImageCount) ||
      parsed.peerImageCount < 1 ||
      !Number.isFinite(parsed.peerInlineImageRunCount) ||
      parsed.peerInlineImageRunCount < 1
    ) {
      throw new Error(
        `Collab image sync mismatch. images='${parsed.peerImageCount}' runs='${parsed.peerInlineImageRunCount}'`,
      );
    }
    if (
      !parsed.replaceExpected ||
      parsed.replacePeerText !== parsed.replaceExpected
    ) {
      throw new Error(
        `Collab replace text mismatch. expected='${parsed.replaceExpected}' got='${parsed.replacePeerText}'`,
      );
    }
    if (
      parsed.replacePeerTableCellA1 !== "" ||
      parsed.replacePeerTableCellB1 !== ""
    ) {
      throw new Error(
        `Collab replace should clear table state. got '${parsed.replacePeerTableCellA1}'/'${parsed.replacePeerTableCellB1}'`,
      );
    }
    if (
      !Number.isFinite(parsed.replacePeerImageCount) ||
      parsed.replacePeerImageCount !== 0 ||
      !Number.isFinite(parsed.replacePeerInlineImageRunCount) ||
      parsed.replacePeerInlineImageRunCount !== 0
    ) {
      throw new Error(
        `Collab replace should clear image state. images='${parsed.replacePeerImageCount}' runs='${parsed.replacePeerInlineImageRunCount}'`,
      );
    }
    if (
      !parsed.reconnectExpected ||
      parsed.reconnectPeerText !== parsed.reconnectExpected
    ) {
      throw new Error(
        `Collab reconnect text mismatch. expected='${parsed.reconnectExpected}' got='${parsed.reconnectPeerText}'`,
      );
    }
    if (
      parsed.reconnectPeerTableCellA1 !== "R1" ||
      parsed.reconnectPeerTableCellB1 !== "R2"
    ) {
      throw new Error(
        `Collab reconnect table mismatch. expected R1/R2, got '${parsed.reconnectPeerTableCellA1}'/'${parsed.reconnectPeerTableCellB1}'`,
      );
    }
    if (
      !Number.isFinite(parsed.reconnectPeerImageCount) ||
      parsed.reconnectPeerImageCount < 1 ||
      !Number.isFinite(parsed.reconnectPeerInlineImageRunCount) ||
      parsed.reconnectPeerInlineImageRunCount < 1
    ) {
      throw new Error(
        `Collab reconnect image sync mismatch. images='${parsed.reconnectPeerImageCount}' runs='${parsed.reconnectPeerInlineImageRunCount}'`,
      );
    }
    if (
      !Number.isFinite(parsed.reconnectPeerRepairRequests) ||
      parsed.reconnectPeerRepairRequests < 1 ||
      !Number.isFinite(parsed.reconnectPeerDivergenceDetected) ||
      parsed.reconnectPeerDivergenceDetected < 1 ||
      !Number.isFinite(parsed.reconnectPeerDivergenceCleared) ||
      parsed.reconnectPeerDivergenceCleared < 1 ||
      !Number.isFinite(parsed.reconnectPeerStateRequestSend) ||
      parsed.reconnectPeerStateRequestSend < 1
    ) {
      throw new Error(
        `Collab reconnect repair visibility missing. repair='${parsed.reconnectPeerRepairRequests}' divergence='${parsed.reconnectPeerDivergenceDetected}/${parsed.reconnectPeerDivergenceCleared}' stateReq='${parsed.reconnectPeerStateRequestSend}'`,
      );
    }
    if (
      parsed.reconnectPeerSawRepairing !== "1" ||
      parsed.reconnectPeerSawDesync !== "1"
    ) {
      throw new Error(
        `Collab reconnect should visibly enter repair/desync. sawRepairing='${parsed.reconnectPeerSawRepairing}' sawDesync='${parsed.reconnectPeerSawDesync}'`,
      );
    }
    if (parsed.resyncing !== "0" || parsed.divergenceSuspected !== "0") {
      throw new Error(
        `Collab repair flags should be cleared at rest. resyncing='${parsed.resyncing}' divergenceSuspected='${parsed.divergenceSuspected}'`,
      );
    }

    console.log(
      `editor-collab-presence-expiry OK: room='${parsed.room}' transport='${parsed.transport}' presence=${parsed.presenceCount} cursors=${parsed.presenceCursorCount} table=${parsed.peerTableCellA1}/${parsed.peerTableCellB1} images=${parsed.peerImageCount} runs=${parsed.peerInlineImageRunCount} replace='${parsed.replacePeerText}' reconnect='${parsed.reconnectPeerText}' peerRepair=${parsed.reconnectPeerRepairRequests} peerDivergence=${parsed.reconnectPeerDivergenceDetected}/${parsed.reconnectPeerDivergenceCleared} peerStateReq=${parsed.reconnectPeerStateRequestSend} expiredPeers=${parsed.expiredPresenceCount} expiredCursors=${parsed.expiredPresenceCursorCount} expiredSelections=${parsed.expiredPresenceSelectionCount}`,
    );
  } finally {
    server.kill();
    await server.exited;
  }
}

await run();
