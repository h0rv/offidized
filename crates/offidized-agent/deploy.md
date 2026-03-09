# offidized-agent — Demo Deployment Guide

`offidized-agent` is an experimental demo app, not a core supported runtime surface for the OSS library.

## Prerequisites

- Rust toolchain (stable)
- [Bun](https://bun.sh) >= 1.0
- Docker + Docker Compose (for containerised deployment)

## 1. Build the ofx binary

From the workspace root:

```bash
cargo build --release -p offidized-cli
```

The binary lands at `../../target/release/ofx` (relative to this directory).

## 2. Run locally (no Docker)

```bash
cd crates/offidized-agent
cp .env.example .env          # add your ANTHROPIC_API_KEY if desired
bun install
bun run server.tsx
```

Open http://localhost:3000.

## 3. Run with Docker Compose

```bash
# From workspace root — binary must be built first (step 1)
cargo build --release -p offidized-cli

cd crates/offidized-agent
cp .env.example .env          # optional: set ANTHROPIC_API_KEY
docker-compose up --build
```

The container mounts the pre-built `ofx` binary and `.preview/` assets as
read-only volumes, so the image itself stays small and rebuilds stay fast.

## 4. Cloud Deployment

### Option A — fly.io (recommended for public demos)

```bash
cd crates/offidized-agent
fly launch          # detects Bun; follow prompts
fly secrets set ANTHROPIC_API_KEY=sk-ant-...
fly deploy
```

Point a Cloudflare-proxied DNS record at the fly.io hostname.

### Option B — Railway

1. Push this directory (or the whole repo) to GitHub.
2. Create a new Railway project, link the repo, set `ANTHROPIC_API_KEY`.
3. Railway auto-detects Bun and deploys on every push.

### Option C — Cloudflare Tunnel (local server, public URL)

```bash
cloudflared tunnel --url http://localhost:3000
```

No DNS or cloud account required — great for quick demos from your machine.

## Environment Variables

| Variable            | Required | Description                             |
| ------------------- | -------- | --------------------------------------- |
| `ANTHROPIC_API_KEY` | No       | Default key; users can supply their own |
| `PORT`              | No       | Override listen port (default: 3000)    |
