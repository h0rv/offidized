# Build the release CLI binary.
cli:
    cargo build -p offidized-cli --release

# Build WASM viewer panels (docview + pptview).
wasm-viewers:
    wasm-pack build --target web --dev crates/offidized-docview \
        --out-dir crates/offidized-docview/pkg/
    wasm-pack build --target web --dev crates/offidized-pptview \
        --out-dir crates/offidized-pptview/pkg/
