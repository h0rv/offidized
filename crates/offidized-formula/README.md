# offidized-formula

Excel formula parser and evaluator.

Part of [offidized](../../README.md).

## Usage

```rust
use offidized_formula::{parse, evaluate};

let ast = parse("SUM(A1:A10)")?;
let result = evaluate(&ast, &cell_lookup)?;
```

Supports arithmetic, comparison, string, and reference operators, plus a growing set of built-in Excel functions.
