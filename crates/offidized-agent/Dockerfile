# offidized-agent Dockerfile
#
# Pre-build requirements (run from workspace root before `docker-compose up`):
#   cargo build --release -p offidized-cli
#
# The ofx binary and WASM preview assets are expected to be pre-built on the
# host and mounted into the container via docker-compose volumes. Building Rust
# inside Docker is intentionally avoided to keep image builds fast.

FROM oven/bun:1 AS deps
WORKDIR /app
COPY package.json bun.lock* ./
RUN bun install --frozen-lockfile

FROM oven/bun:1
WORKDIR /app

# App source
COPY --from=deps /app/node_modules ./node_modules
COPY server.tsx package.json tsconfig.json ./

# The ofx binary is mounted at /usr/local/bin/ofx via docker-compose volume.
# The WASM preview assets are mounted at /app/.preview via docker-compose volume.

EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:3000/ || exit 1

CMD ["bun", "run", "server.tsx"]
