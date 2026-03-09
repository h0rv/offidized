# offidized-ir

Bidirectional lossless text IR for Office files. Derive a human-readable text representation from `.xlsx`/`.docx`/`.pptx`, edit it, and apply changes back — with full roundtrip fidelity.

Part of [offidized](../../README.md).

## Usage

```rust
use offidized_ir::{derive, apply, DeriveOptions};

let ir = derive("report.xlsx".as_ref(), DeriveOptions::default())?;
// edit the IR string...
apply(&modified_ir, "output.xlsx".as_ref())?;
```

Also provides the unified agent API (`UnifiedDocument`, `UnifiedEdit`) for cross-format edits. See [`docs/design/text-ir.md`](../../docs/design/text-ir.md) and [`docs/design/unified-agent-api.md`](../../docs/design/unified-agent-api.md).
