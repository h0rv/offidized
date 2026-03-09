#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "${repo_root}"

echo "==> rustfmt"
cargo qfmt

echo "==> clippy (strict)"
cargo qclippy

echo "==> rust check (workspace)"
cargo qcheck
