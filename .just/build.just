root := parent_directory(source_directory())

# Build the release CLI binary.
cli:
    cd {{root}} && cargo build -p offidized-cli --release

# Build WASM viewer panels (docview + pptview).
wasm-viewers:
    cd {{root}}/crates/offidized-docview && wasm-pack build --target web --dev --out-dir pkg -- --features editing
    cd {{root}}/crates/offidized-pptview && wasm-pack build --target web --dev --out-dir pkg

# Build the spreadsheet viewer with editing enabled for the demo site.
wasm-xlview-edit:
    cd {{root}}/crates/offidized-xlview && wasm-pack build --target web --dev --out-dir pkg -- --features editing

# Build the generic WASM bindings used by the site for blank-file creation.
wasm-core:
    cd {{root}}/crates/offidized-wasm && wasm-pack build --target web --dev --out-dir pkg
