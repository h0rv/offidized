# Stress tests (ignored suites — real-world corpus files).
stress:
    cargo test --workspace -- --ignored

# Xlsx stress corpus roundtrip.
stress-xlsx:
    cargo test -p offidized-xlsx -- --ignored

# Docx large sweep roundtrip.
stress-docx:
    cargo test -p offidized-docx -- --ignored

# Pptx broader sweep roundtrip.
stress-pptx:
    cargo test -p offidized-pptx -- --ignored

# Differential corpus (Rust-only, compare roundtripped output vs original input).
diff-rust:
    bash scripts/differential_corpus.sh --rust-only --compare-input

# Differential corpus (Rust vs C# Open XML SDK).
diff-csharp:
    bash scripts/differential_corpus.sh
