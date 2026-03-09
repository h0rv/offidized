/// <reference types="@kitajs/html" />
import { Elysia, t } from "elysia";
import html from "@elysiajs/html";
import { resolve } from "path";
import {
  parseSid,
  makeSetCookie,
  getSession,
  getOrCreateSession,
  hasSession,
  listSessions,
} from "./session";
import { createWorkspace, esc } from "./helpers";
import { PREVIEW_DIR } from "./preview";
import { createStream } from "./agent";
import { Page } from "../components/Page";

const activeStreams = new Map<
  string,
  { result: any; sid: string; createdAt: number; traceId: string }
>();

function logTrace(
  traceId: string,
  event: string,
  data?: Record<string, unknown>,
) {
  console.log(
    JSON.stringify({
      traceId,
      event,
      ts: new Date().toISOString(),
      ...data,
    }),
  );
}

setInterval(() => {
  const MAX_AGE = 5 * 60 * 1000;
  const now = Date.now();
  for (const [id, entry] of activeStreams) {
    if (now - entry.createdAt > MAX_AGE) {
      activeStreams.delete(id);
      logTrace(entry.traceId, "stream_expired", { streamId: id });
    }
  }
}, 60 * 1000);

export function registerRoutes(app: Elysia) {
  return app
    .use(html())
    .get("/", ({ html: renderHtml, request, set }) => {
      const cookieSid = parseSid(request.headers.get("Cookie"));
      const sess = getOrCreateSession(cookieSid);
      // Always set cookie — either fresh or confirming existing
      set.headers["Set-Cookie"] = makeSetCookie(sess.id);
      return renderHtml(<Page currentSid={sess.id} />);
    })
    .get("/preview/:file", ({ params }) => {
      const filePath = resolve(PREVIEW_DIR, params.file);
      const file = Bun.file(filePath);
      const ext = params.file.split(".").pop();
      const mime =
        ext === "wasm" ? "application/wasm" : "text/javascript;charset=utf-8";
      return new Response(file, {
        headers: { "Content-Type": mime, "Cache-Control": "no-cache" },
      });
    })
    .get("/api/sessions", () => {
      return new Response(JSON.stringify(listSessions()), {
        headers: {
          "Content-Type": "application/json",
          "Cache-Control": "no-store",
        },
      });
    })
    .get("/api/switch/:id", ({ params, set }) => {
      const sess = getSession(params.id);
      if (!sess) return new Response("Session not found", { status: 404 });
      set.headers["Set-Cookie"] = makeSetCookie(sess.id);
      // Return conversation HTML + session metadata
      const messages = sess.conversation
        .map((m) => {
          if (m.role === "user") {
            return `<div class="px-4 py-3 border-t border-zinc-800/40">
              <div class="text-[11px] text-zinc-500 mb-0.5">you</div>
              <div class="text-zinc-200 whitespace-pre-wrap">${esc(m.content)}</div>
            </div>`;
          }
          return `<div class="px-4 py-3 border-t border-zinc-800/40">
            <div class="text-[11px] text-zinc-500 mb-0.5">offidized</div>
            <div class="ai-out"><div class="ai-seg">${esc(m.content).replace(/\r?\n/g, "<br>")}</div></div>
          </div>`;
        })
        .join("");
      return new Response(
        JSON.stringify({
          html: messages,
          currentFile: sess.currentFile,
          hasFile: !!sess.currentFilePath,
        }),
        {
          headers: {
            "Content-Type": "application/json",
            "Cache-Control": "no-store",
          },
        },
      );
    })
    .get("/api/file-version", async ({ request }) => {
      const sid = parseSid(request.headers.get("Cookie"));
      if (!sid || !hasSession(sid))
        return new Response("0", { headers: { "Content-Type": "text/plain" } });
      const sess = getSession(sid);
      const filePath = sess?.currentFilePath;
      if (!filePath)
        return new Response("0", { headers: { "Content-Type": "text/plain" } });
      try {
        const mtime = (await Bun.file(filePath).stat()).mtimeMs;
        return new Response(String(mtime), {
          headers: {
            "Content-Type": "text/plain",
            "Cache-Control": "no-store",
          },
        });
      } catch {
        return new Response("0", { headers: { "Content-Type": "text/plain" } });
      }
    })
    .get("/api/file", ({ request }) => {
      const sid = parseSid(request.headers.get("Cookie"));
      if (!sid) return new Response("No session", { status: 404 });
      const sess = getSession(sid);
      if (!sess?.currentFilePath)
        return new Response("No file", { status: 404 });
      return new Response(Bun.file(sess.currentFilePath), {
        headers: { "Cache-Control": "no-store" },
      });
    })
    .post("/api/reset", ({ set }) => {
      // Don't delete the old session — it stays in the sidebar
      // Create a brand new session
      const sess = getOrCreateSession(null);
      set.headers["Set-Cookie"] = makeSetCookie(sess.id);
      return new Response(JSON.stringify({ sid: sess.id }), {
        headers: { "Content-Type": "application/json" },
      });
    })
    .post(
      "/api/chat",
      async ({ body, request, set }) => {
        const cookieSid = parseSid(request.headers.get("Cookie"));
        const sess = getOrCreateSession(cookieSid);
        // Ensure cookie is set if we created a new session
        if (sess.id !== cookieSid) {
          set.headers["Set-Cookie"] = makeSetCookie(sess.id);
        }

        const claudeKey =
          request.headers.get("X-Claude-Key") ||
          process.env.ANTHROPIC_API_KEY ||
          "";
        const googleKey =
          request.headers.get("X-Google-Key") ||
          process.env.GOOGLE_GENERATIVE_AI_API_KEY ||
          "";
        const openrouterKey =
          request.headers.get("X-OpenRouter-Key") ||
          process.env.OPENROUTER_API_KEY ||
          "";

        console.log(
          "API keys - OpenRouter:",
          openrouterKey ? `${openrouterKey.slice(0, 8)}...` : "(none)",
          "Google:",
          googleKey ? `${googleKey.slice(0, 8)}...` : "(none)",
          "Claude:",
          claudeKey ? `${claudeKey.slice(0, 8)}...` : "(none)",
        );

        if (!openrouterKey && !googleKey && !claudeKey) {
          return new Response(
            `<div class="px-4 py-3 border-t border-zinc-800/40">
              <div class="text-[11px] text-red-400 mb-1">error</div>
              <div class="text-zinc-400 text-[13px]">No API key configured. Set OPENROUTER_API_KEY (recommended), GOOGLE_GENERATIVE_AI_API_KEY, or ANTHROPIC_API_KEY env var, or click the key icon.</div>
            </div>`,
            {
              headers: {
                "Content-Type": "text/html",
                "Cache-Control": "no-store",
              },
            },
          );
        }

        const { message, filename = "", filedata = "" } = body as any;

        const traceId = crypto.randomUUID();
        logTrace(traceId, "chat_start", {
          sid: sess.id.slice(0, 8),
          messageLen: message.length,
          filename: filename || null,
        });

        if (!sess.workspaceDir) {
          sess.workspaceDir = await createWorkspace();
          logTrace(traceId, "workspace_created", { dir: sess.workspaceDir });
        }

        if (filename && filedata) {
          sess.currentFile = filename;
          const bytes = Buffer.from(filedata, "base64");
          sess.currentFilePath = resolve(sess.workspaceDir, filename);
          await Bun.write(sess.currentFilePath, bytes);
          logTrace(traceId, "file_uploaded", { filename, size: bytes.length });

          // Snapshot the original upload as v0
          if (!sess.versions.some((v) => v.isOriginal)) {
            const ext = filename.split(".").pop() ?? "xlsx";
            const base = filename.replace(/\.\w+$/, "");
            const origPath = resolve(sess.workspaceDir, `${base}.v0.${ext}`);
            await Bun.write(origPath, bytes);
            sess.versions.push({
              n: 0,
              ts: Date.now(),
              path: origPath,
              isOriginal: true,
            });
            logTrace(traceId, "original_snapshot", {
              file: origPath.split("/").pop(),
            });
          }
        } else if (filename) {
          sess.currentFile = filename;
        }

        // Set session title from first message or filename
        if (!sess.title) {
          if (filename) {
            sess.title = filename;
          } else {
            sess.title =
              message.length > 50 ? message.slice(0, 47) + "..." : message;
          }
        }

        sess.conversation.push({ role: "user", content: message });

        const id = crypto.randomUUID();

        const result = await createStream(sess.id, sess, {
          openrouter: openrouterKey,
          google: googleKey,
          claude: claudeKey,
        });

        activeStreams.set(id, {
          result,
          sid: sess.id,
          createdAt: Date.now(),
          traceId,
        });

        return new Response(
          `<div class="px-4 py-3 border-t border-zinc-800/40">
            <div class="text-[11px] text-zinc-500 mb-0.5">you</div>
            <div class="text-zinc-200 whitespace-pre-wrap">${esc(message)}</div>
          </div>
          <div class="px-4 py-3 border-t border-zinc-800/40">
            <div class="text-[11px] text-zinc-500 mb-0.5">offidized</div>
            <div class="ai-out" data-stream="/api/stream/${id}"></div>
          </div>`,
          { headers: { "Content-Type": "text/html" } },
        );
      },
      {
        body: t.Object({
          message: t.String(),
          filename: t.Optional(t.String()),
          filedata: t.Optional(t.String()),
        }),
      },
    )
    .get("/api/stream/:id", ({ params }) => {
      const entry = activeStreams.get(params.id);
      activeStreams.delete(params.id);

      const enc = new TextEncoder();

      if (!entry) {
        return new Response("event: done\ndata: \n\n", {
          headers: { "Content-Type": "text/event-stream" },
        });
      }

      const { result, sid, traceId } = entry;
      const sess = getSession(sid);
      if (!sess) {
        return new Response("event: done\ndata: \n\n", {
          headers: { "Content-Type": "text/event-stream" },
        });
      }
      logTrace(traceId, "stream_start", { sid: sid.slice(0, 8) });

      let fullText = "";
      let inText = false;
      let mtimeBefore = 0;
      let turnStartMtime = 0;
      const pendingToolBlocks = new Map<
        string,
        { html: string; toolName: string }
      >();

      function extractToolResult(
        toolName: string,
        raw: unknown,
      ): { body: string; isErr: boolean } {
        if (raw == null) return { body: "", isErr: false };
        if (typeof raw === "string") return { body: raw, isErr: false };

        const obj = raw as Record<string, unknown>;

        // bash → show stdout, append stderr if present
        if (toolName === "bash") {
          const stdout = String(obj.stdout ?? "");
          const stderr = String(obj.stderr ?? "");
          const code = obj.exitCode as number | undefined;
          const isErr = code != null && code !== 0;
          let body = stdout;
          if (stderr) {
            body += (body ? "\n" : "") + stderr;
          }
          if (isErr && !body) {
            body = `exit code ${code}`;
          }
          return { body, isErr };
        }

        // readFile → show content directly
        if (toolName === "readFile") {
          return { body: String(obj.content ?? ""), isErr: false };
        }

        // writeFile → short confirmation
        if (toolName === "writeFile") {
          return {
            body: obj.success ? "written" : "write failed",
            isErr: !obj.success,
          };
        }

        // error field on any tool
        if (obj.error) {
          return { body: String(obj.error), isErr: true };
        }

        // Unknown tool — try to extract string values, fall back to JSON
        const stringVals = Object.entries(obj)
          .filter(([, v]) => typeof v === "string" && v.length > 0)
          .map(([, v]) => v as string);
        if (stringVals.length === 1)
          return { body: stringVals[0] ?? "", isErr: false };
        if (stringVals.length > 1)
          return { body: stringVals.join("\n"), isErr: false };

        return { body: JSON.stringify(raw, null, 2), isErr: false };
      }

      function toolInvHtml(name: string, rawArgs: unknown): string {
        const args = (
          rawArgs != null && typeof rawArgs === "object" ? rawArgs : {}
        ) as Record<string, unknown>;
        if (name === "bash") {
          const cmd = esc(String(args.command ?? "").slice(0, 300));
          return `<span class="tool-inv">$ ${cmd}</span>`;
        }
        if (name === "readFile" || name === "writeFile") {
          const p = esc(String(args.path ?? "").replace(/.*[\\/]/, ""));
          return `<span class="tool-inv"><span class="op">${name}</span> <span class="ph">${p}</span></span>`;
        }
        const summary = Object.entries(args)
          .filter(([, v]) => v != null && v !== "")
          .map(([k, v]) => `${k}=${esc(String(v).slice(0, 40))}`)
          .join(" ")
          .slice(0, 120);
        return `<span class="tool-inv">${summary || "(no args)"}</span>`;
      }

      function toolDisplayName(name: string): string {
        return name;
      }

      const stream = new ReadableStream({
        async start(controller) {
          let streamClosed = false;
          const sseEncode = (ev: string, d: string) => {
            const lines = d
              .split("\n")
              .map((l) => `data: ${l}`)
              .join("\n");
            return `event: ${ev}\n${lines}\n\n`;
          };
          const enqueue = (data: string) => {
            if (streamClosed) return;
            try {
              controller.enqueue(enc.encode(data));
            } catch {
              streamClosed = true;
            }
          };
          const send = (d: string) => enqueue(sseEncode("token", d));
          const sendEvent = (ev: string) => enqueue(`event: ${ev}\ndata: \n\n`);
          const sendEventData = (ev: string, data: string) =>
            enqueue(sseEncode(ev, data));
          const openText = () => {
            if (!inText) {
              send(`<div class="ai-seg">`);
              inText = true;
            }
          };
          const closeText = () => {
            if (inText) {
              send(`</div>`);
              inText = false;
            }
          };

          if (sess.currentFilePath) {
            try {
              turnStartMtime = (await Bun.file(sess.currentFilePath).stat())
                .mtimeMs;
              mtimeBefore = turnStartMtime;
            } catch {
              /* file may not exist */
            }
          }

          try {
            for await (const part of result.fullStream) {
              if (
                part.type !== "text-delta" &&
                part.type !== "start-step" &&
                part.type !== "finish-step"
              ) {
                console.log(
                  `[stream ${sid.slice(0, 8)}] part.type=${part.type}`,
                  part.type === "tool-call"
                    ? (part as any).toolName
                    : part.type === "tool-result"
                      ? `id=${(part as any).toolCallId}`
                      : "",
                );
              }
              switch (part.type) {
                case "text-delta":
                  openText();
                  fullText += part.text;
                  send(esc(part.text).replace(/\r?\n/g, "<br>"));
                  break;

                case "tool-call": {
                  closeText();
                  if (part.toolName === "bash" && sess.currentFilePath) {
                    try {
                      mtimeBefore = (
                        await Bun.file(sess.currentFilePath).stat()
                      ).mtimeMs;
                    } catch {
                      /* file may not exist */
                    }
                  }
                  const inv = toolInvHtml(part.toolName, part.args);
                  const head = `<div class="tool-head"><span class="tool-name">${toolDisplayName(part.toolName)}</span>${inv}</div>`;
                  pendingToolBlocks.set(part.toolCallId, {
                    html: head,
                    toolName: part.toolName,
                  });
                  break;
                }

                case "tool-result": {
                  const callId =
                    (part as any).toolCallId ??
                    Math.random().toString(36).slice(2);
                  const pending = pendingToolBlocks.get(callId);
                  const toolName =
                    (part as any).toolName ?? pending?.toolName ?? "";
                  const raw = (part as any).output ?? (part as any).result;
                  const { body, isErr } = extractToolResult(toolName, raw);

                  const maxLen = 8000;
                  const trimmed =
                    body.length > maxLen
                      ? body.slice(0, maxLen) + "\n[truncated]"
                      : body;
                  const lineCount = trimmed.split("\n").length;
                  const CLIP = 6;

                  let outHtml: string;
                  if (!trimmed || trimmed === "(no output)") {
                    outHtml = `<div class="tool-out muted">(no output)</div>`;
                  } else if (lineCount > CLIP) {
                    const short = esc(
                      trimmed.split("\n").slice(0, CLIP).join("\n"),
                    );
                    const full2 = esc(trimmed);
                    outHtml =
                      `<div class="tool-out clipped${isErr ? " err" : ""}" id="to-${callId}">${short}</div>` +
                      `<span class="tool-expand" onclick="` +
                      `var el=document.getElementById('to-${callId}');` +
                      `el.innerHTML=${JSON.stringify(full2)};` +
                      `el.classList.remove('clipped');` +
                      `this.remove()` +
                      `">▾ ${lineCount} lines — expand</span>`;
                  } else {
                    outHtml = `<div class="tool-out${isErr ? " err" : ""}">${esc(trimmed)}</div>`;
                  }

                  const pendingHead = pending?.html ?? "";
                  pendingToolBlocks.delete(callId);
                  send(
                    `<div class="tool-block">${pendingHead}${outHtml}</div>`,
                  );

                  if (sess.currentFilePath) {
                    try {
                      const mtimeAfter = (
                        await Bun.file(sess.currentFilePath).stat()
                      ).mtimeMs;
                      if (mtimeAfter !== mtimeBefore) {
                        sendEvent("file-updated");
                        mtimeBefore = mtimeAfter;
                      }
                    } catch {
                      /* file may not exist */
                    }
                  }
                  break;
                }
              }
            }
          } catch (err: any) {
            console.error("[stream error]", err?.message ?? err);
            const msg = err instanceof Error ? err.message : String(err);
            openText();
            send(`<span style="color:#fca5a5">${esc(msg)}</span>`);
            fullText += `\n[Error: ${msg}]`;
          }

          closeText();

          if (sess.workspaceDir) {
            try {
              const glob = new Bun.Glob("*.{xlsx,docx,pptx}");
              const found: { name: string; mtime: number }[] = [];
              for await (const f of glob.scan({
                cwd: sess.workspaceDir,
                absolute: false,
              })) {
                if (/\.v\d+\./.test(f)) continue;
                const stat = await Bun.file(
                  resolve(sess.workspaceDir, f),
                ).stat();
                found.push({ name: f, mtime: stat.mtimeMs });
              }
              found.sort((a, b) => b.mtime - a.mtime);
              if (found[0]) {
                const newPath = resolve(sess.workspaceDir, found[0].name);
                if (newPath !== sess.currentFilePath) {
                  sess.currentFilePath = newPath;
                  sess.currentFile = found[0].name;
                  logTrace(traceId, "file_autodetected", {
                    filename: found[0].name,
                  });
                  sendEventData("file-created", found[0].name);
                }
              }
            } catch {
              /* file may not exist */
            }
          }

          if (sess.currentFilePath) {
            try {
              const mtimeEnd = (await Bun.file(sess.currentFilePath).stat())
                .mtimeMs;
              if (
                mtimeEnd > 0 &&
                (turnStartMtime === 0 || mtimeEnd !== turnStartMtime)
              ) {
                const vn = sess.versions.length + 1;
                const ext = sess.currentFile.split(".").pop() ?? "xlsx";
                const base = sess.currentFile.replace(/\.\w+$/, "");
                const snapPath = resolve(
                  sess.workspaceDir,
                  `${base}.v${vn}.${ext}`,
                );
                await Bun.write(snapPath, Bun.file(sess.currentFilePath));
                sess.versions.push({ n: vn, ts: Date.now(), path: snapPath });
                logTrace(traceId, "snapshot_created", {
                  version: vn,
                  file: snapPath.split("/").pop(),
                });
                sendEvent("versions-updated");
              }
            } catch {
              /* file may not exist */
            }
          }

          if (fullText) {
            sess.conversation.push({ role: "assistant", content: fullText });
          }
          logTrace(traceId, "stream_done", {
            textLen: fullText.length,
            versions: sess.versions.length,
          });
          const ts = new Date().toLocaleTimeString([], {
            hour: "2-digit",
            minute: "2-digit",
          });
          send(`<div class="turn-end">done ${ts}</div>`);
          sendEvent("session-updated");
          enqueue("event: done\ndata: \n\n");
          if (!streamClosed) {
            try {
              controller.close();
            } catch {
              /* already closed */
            }
          }
          streamClosed = true;
        },
      });

      return new Response(stream, {
        headers: {
          "Content-Type": "text/event-stream",
          "Cache-Control": "no-cache",
        },
      });
    })
    .get("/api/download", ({ request }) => {
      const sid = parseSid(request.headers.get("Cookie"));
      if (!sid) return new Response("No session", { status: 404 });
      const sess = getSession(sid);
      if (!sess?.currentFilePath || !sess.currentFile) {
        return new Response("No file loaded", { status: 404 });
      }
      const file = Bun.file(sess.currentFilePath);
      const ext = sess.currentFile.split(".").pop()?.toLowerCase();
      const mime =
        ext === "xlsx"
          ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          : ext === "docx"
            ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            : ext === "pptx"
              ? "application/vnd.openxmlformats-officedocument.presentationml.presentation"
              : "application/octet-stream";
      return new Response(file, {
        headers: {
          "Content-Type": mime,
          "Content-Disposition": `attachment; filename="${sess.currentFile.replace(/\.(\w+)$/, "_updated.$1")}"`,
        },
      });
    })
    .get("/api/versions", ({ request }) => {
      const sid = parseSid(request.headers.get("Cookie"));
      if (!sid || !hasSession(sid)) {
        return new Response("[]", {
          headers: { "Content-Type": "application/json" },
        });
      }
      const sess = getSession(sid);
      if (!sess) {
        return new Response("[]", {
          headers: { "Content-Type": "application/json" },
        });
      }
      const out = [...sess.versions].reverse().map((v) => ({
        n: v.n,
        label: new Date(v.ts).toLocaleTimeString([], {
          hour: "2-digit",
          minute: "2-digit",
        }),
        isOriginal: !!v.isOriginal,
      }));
      return new Response(JSON.stringify(out), {
        headers: {
          "Content-Type": "application/json",
          "Cache-Control": "no-store",
        },
      });
    })
    .get("/api/download/v/:n", ({ params, request }) => {
      const sid = parseSid(request.headers.get("Cookie"));
      if (!sid) return new Response("Not found", { status: 404 });
      const sess = getSession(sid);
      if (!sess?.versions.length)
        return new Response("Not found", { status: 404 });
      const n = parseInt(params.n, 10);
      const ver = sess.versions.find((v) => v.n === n);
      if (!ver) return new Response("Version not found", { status: 404 });
      const file = Bun.file(ver.path);
      const ext = ver.path.split(".").pop()?.toLowerCase();
      const mime =
        ext === "xlsx"
          ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          : ext === "docx"
            ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            : ext === "pptx"
              ? "application/vnd.openxmlformats-officedocument.presentationml.presentation"
              : "application/octet-stream";
      const snapName =
        ver.path
          .split("/")
          .pop()
          ?.replace(/\.v(\d+)\./, "_v$1.") ?? "file.xlsx";
      return new Response(file, {
        headers: {
          "Content-Type": mime,
          "Content-Disposition": `attachment; filename="${snapName}"`,
        },
      });
    });
}
