import { join, resolve } from "node:path";

const DIST = resolve(import.meta.dir, "dist");
const basePort = Number(Bun.env.PORT ?? "4173");

function createFetchHandler() {
  return async function fetch(req: Request) {
    const url = new URL(req.url);
    const pathname = url.pathname === "/" ? "/index.html" : url.pathname;
    const file = Bun.file(join(DIST, pathname));

    if (await file.exists()) {
      return new Response(file);
    }

    return new Response("Not found", { status: 404 });
  };
}

let server: ReturnType<typeof Bun.serve> | null = null;
let chosenPort = basePort;

for (let offset = 0; offset < 10; offset += 1) {
  const candidate = basePort + offset;
  try {
    server = Bun.serve({
      port: candidate,
      fetch: createFetchHandler(),
    });
    chosenPort = candidate;
    break;
  } catch (error) {
    if (offset === 9) throw error;
  }
}

if (!server) {
  throw new Error("Could not start static demo server");
}

console.log(`Static demo at http://localhost:${chosenPort}`);
