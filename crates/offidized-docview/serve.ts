import { join } from "node:path";
import type { Server } from "bun";

import demo from "./demo/index.html";
import editorPage from "./demo/editor.html";
import caretHarnessPage from "./demo/caret-harness.html";

const ROOT_DIR = import.meta.dir;
const DEMO_FONT_PATHS: Record<string, string> = {
  "pt-sans.ttc": "/System/Library/Fonts/Supplemental/PTSans.ttc",
  "pt-serif.ttc": "/System/Library/Fonts/Supplemental/PTSerif.ttc",
  "pt-mono.ttc": "/System/Library/Fonts/Supplemental/PTMono.ttc",
  "comic-sans.ttf": "/System/Library/Fonts/Supplemental/Comic Sans MS.ttf",
};

function resolvePort(requestedPort?: number): number {
  if (typeof requestedPort === "number") return requestedPort;
  const parsed = Number.parseInt(Bun.env.PORT ?? "3033", 10);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : 3033;
}

function logServerUrls(port: number): void {
  console.log(`Doc viewer at http://localhost:${port}`);
  console.log(`Doc editor at http://localhost:${port}/editor`);
  console.log(`Caret harness at http://localhost:${port}/harness/caret`);
}

export function startDemoServer(
  requestedPort?: number,
  development = true,
): Server<undefined> {
  return Bun.serve({
    port: resolvePort(requestedPort),
    routes: {
      "/": demo,
      "/editor": editorPage,
      "/harness/caret": caretHarnessPage,
    },
    async fetch(req) {
      // Serve static files (WASM binary, pkg/ assets)
      const url = new URL(req.url);
      if (url.pathname.startsWith("/__demo-fonts/")) {
        const name = url.pathname.slice("/__demo-fonts/".length);
        const fontPath = DEMO_FONT_PATHS[name];
        if (!fontPath) {
          return new Response("Not found", { status: 404 });
        }
        const file = Bun.file(fontPath);
        if (await file.exists()) {
          return new Response(file);
        }
        return new Response("Not found", { status: 404 });
      }
      const file = Bun.file(join(ROOT_DIR, url.pathname.replace(/^\//, "")));
      if (await file.exists()) {
        return new Response(file);
      }
      return new Response("Not found", { status: 404 });
    },
    development: development
      ? {
          hmr: true,
          console: true,
        }
      : undefined,
  });
}

if (import.meta.main) {
  const server = startDemoServer();
  logServerUrls(server.port ?? resolvePort());
}
