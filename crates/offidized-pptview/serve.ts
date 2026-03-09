import demo from "./demo/index.html";

const server = Bun.serve({
  port: 8082,
  routes: {
    "/": demo,
  },
  async fetch(req) {
    // Serve static files (WASM binary, pkg/ assets)
    const url = new URL(req.url);
    const file = Bun.file("." + url.pathname);
    if (await file.exists()) {
      return new Response(file);
    }
    return new Response("Not found", { status: 404 });
  },
  development: {
    hmr: true,
    console: true,
  },
});

console.log(`Pptx viewer at http://localhost:${server.port}`);
