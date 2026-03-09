root := parent_directory(source_directory())

# Run the AI agent locally (hot reload).
agent: install-cli
    cd {{root}}/crates/offidized-agent && bun install --silent && bun --hot server.tsx

# Install the CLI globally (~/.cargo/bin/ofx).
install-cli:
    cd {{root}} && cargo install --path crates/offidized-cli

# Start the example gallery server (http://localhost:3001).
gallery: _ensure-viewers
    cd {{root}}/crates/offidized-py/examples/gallery && bun server.ts

# Build viewer WASM if missing.
[private]
_ensure-viewers:
    #!/usr/bin/env bash
    set -euo pipefail
    ROOT="{{root}}"
    XLVIEW="$ROOT/../xlview"
    if [ -d "$XLVIEW" ] && [ ! -f "$XLVIEW/pkg/xlview_bg.wasm" ]; then
        echo "Building xlview WASM..."
        wasm-pack build --target web --dev "$XLVIEW" \
            --out-dir "$XLVIEW/pkg/"
    fi
    if [ ! -f "$ROOT/crates/offidized-docview/pkg/offidized_docview_bg.wasm" ]; then
        echo "Building docview WASM..."
        wasm-pack build --target web --dev "$ROOT/crates/offidized-docview" \
            --out-dir "$ROOT/crates/offidized-docview/pkg/"
    fi
    if [ ! -f "$ROOT/crates/offidized-pptview/pkg/offidized_pptview_bg.wasm" ]; then
        echo "Building pptview WASM..."
        wasm-pack build --target web --dev "$ROOT/crates/offidized-pptview" \
            --out-dir "$ROOT/crates/offidized-pptview/pkg/"
    fi

# Run all Python examples (generate files without the server).
examples:
    #!/usr/bin/env bash
    set -euo pipefail
    ROOT="{{root}}"
    PY="$ROOT/crates/offidized-py"
    OUT="$ROOT/crates/offidized-py/examples/gallery/output"
    mkdir -p "$OUT"
    ok=0; fail=0
    for f in "$PY"/examples/pptx/*.py "$PY"/examples/docx/*.py; do
        echo -n "  ${f#$ROOT/} ... "
        if (cd "$OUT" && uv run --project "$PY" python "$f") >/dev/null 2>&1; then
            echo "ok"; ok=$((ok + 1))
        else
            echo "FAIL"; fail=$((fail + 1))
        fi
    done
    export TMPDIR="$OUT"
    for f in "$PY"/examples/xlsx/*.py; do
        echo -n "  ${f#$ROOT/} ... "
        if (cd "$OUT" && uv run --project "$PY" python "$f") >/dev/null 2>&1; then
            # xlsx scripts use tempfile.mkdtemp — hoist files up, remove empty dirs
            find "$OUT" -mindepth 2 -name '*.xlsx' -exec mv {} "$OUT/" \;
            find "$OUT" -mindepth 1 -type d -empty -delete 2>/dev/null || true
            echo "ok"; ok=$((ok + 1))
        else
            echo "FAIL"; fail=$((fail + 1))
        fi
    done
    # Clean up any remaining temp artifacts
    find "$OUT" -mindepth 1 -not -name '*.pptx' -not -name '*.docx' -not -name '*.xlsx' -delete 2>/dev/null || true
    echo ""
    echo "Done: $ok ok, $fail failed"
    echo "Output: $OUT/"
    ls "$OUT"/*.{pptx,docx,xlsx} 2>/dev/null | xargs -n1 basename
