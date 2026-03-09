# offidized-agent

Experimental Bun/TypeScript demo app for Office file manipulation via offidized.

This is not part of the core supported library surface. Keep it only if you want a local demo app in the monorepo.

## Setup

```bash
cd crates/offidized-agent
bun install
```

## Run

```bash
bun run server.tsx
```

Starts an HTTP server (default port 3000) that exposes Office file reading, writing, and editing endpoints backed by the offidized WASM bindings. Includes optional document preview when viewer assets are built.

## Configuration

Set environment variables in `.env`:

- `PORT` — Server port (default: 3000)
