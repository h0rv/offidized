set shell := ["bash", "-uc"]

mod build   '.just/build.just'
mod dev     '.just/dev.just'
mod release '.just/release.just'
mod skills  '.just/skills.just'
mod stress  '.just/stress.just'

# Show available recipes.
default:
    @just --list --list-submodules

# ── Quality gates ─────────────────────────────────────────────────────────

# Formatting check (workspace).
fmt:
    cargo qfmt

# Strict clippy linting (workspace).
clippy:
    cargo qclippy

# Rust compile checks (workspace).
check:
    cargo qcheck

# Required quality gates (CI-equivalent).
all: fmt clippy check
    @echo "All quality checks passed."

# CI alias.
ci: all

# Full workspace tests.
test:
    cargo test --workspace

# ── Setup ─────────────────────────────────────────────────────────────────

# Install JS deps for the optional demo app.
install:
    cd crates/offidized-agent && bun install

# Full setup: build CLI + install optional demo deps.
setup: install
    just build cli

# ── Demo ──────────────────────────────────────────────────────────────────

# Generate the demo spreadsheet.
demo:
    mkdir -p demo
    cargo run --release -p offidized-xlsx --example pivot_demo

# Fetch the Open XML SDK schema data required for site builds.
site-fetch-openxml-data:
    mkdir -p references
    if [ ! -d references/Open-XML-SDK/.git ]; then git clone --depth 1 --filter=blob:none --sparse https://github.com/dotnet/Open-XML-SDK.git references/Open-XML-SDK; fi
    git -C references/Open-XML-SDK sparse-checkout set data

# Install the JS dependency used by the browser demo build.
site-install-docview-deps:
    cd crates/offidized-docview && bun install --frozen-lockfile

# Generate the sample Office files used by the demo site.
site-build-samples:
    just dev examples

# Build the static browser demo site.
site-build:
    just build wasm-viewers
    just build wasm-xlview-edit
    just build wasm-core
    bun run site/build.ts

# Serve the built static browser demo site.
site-serve: site-build
    bun run site/serve.ts

# Open the demo in LibreOffice.
demo-open: demo
    /opt/homebrew/bin/soffice --calc demo/openpyxl_breaker.xlsx
