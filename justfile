set shell := ["bash", "-uc"]

mod build   '.just/build.just'
mod dev     '.just/dev.just'
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

# Open the demo in LibreOffice.
demo-open: demo
    /opt/homebrew/bin/soffice --calc demo/openpyxl_breaker.xlsx
