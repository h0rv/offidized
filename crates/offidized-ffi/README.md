# offidized-ffi

C ABI bindings for offidized OOXML crates.

Part of [offidized](../../README.md).

## Build

```bash
cargo build -p offidized-ffi --release
```

Produces a shared library (`liboffidized_ffi.so` / `.dylib` / `.dll`) that can be linked from C, C++, or any language with C FFI support. Covers xlsx, docx, and pptx read/write operations.
