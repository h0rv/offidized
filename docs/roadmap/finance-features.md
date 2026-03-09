# Finance Feature Depth Roadmap

Status: `active`
Scope: feature depth first (no performance optimization work in this plan)

## Goal

Build institutional-grade finance authoring features (hedge fund / asset manager workflows) on top of `offidized-xlsx`, prioritizing capability breadth and correctness over speed.

## Tracking Conventions

- Status values: `todo`, `in_progress`, `blocked`, `done`
- Every item must have:
  - API surface
  - acceptance tests
  - migration/examples updates

---

## P0 (Foundational Finance Depth)

### P0.1 High-Level Table/Model API

Status: `done`

Declarative model authoring for dimensions, measures, scenarios.

```rust
let model = wb.finance_model("Core")
    .dimension("date", dates)
    .dimension("book", books)
    .measure("gross_exposure", MeasureType::Currency)
    .measure("net_exposure", MeasureType::Currency)
    .measure("pnl", MeasureType::Currency)
    .scenario("base")
    .scenario("stress_5sigma");
```

- `Workbook::finance_model(...)` fluent builder
- `MeasureType` domain enum and model metadata/data sheet materialization

---

### P0.2 Finance-Aware Number Format Semantics

Status: `done`

First-class finance format enums (bps, mm/bn, signed pct, etc.).

```rust
cell.set_finance_format(FinFormat::Bps1);
cell.set_finance_format(FinFormat::UsdMillions2);
cell.set_finance_format(FinFormat::Pct2Signed);
```

- `FinFormat` enum with stable format code mapping: bps, USD millions/billions, signed %, multiples, integer thousands.
- `Style::set_finance_format(FinFormat)` convenience API.

---

### P0.3 Pivot Builder Safety Layer

Status: `done`

Safer high-level pivot API with field validation and index-safe wiring.

```rust
ws.pivot("RiskPivot")
  .source("Positions!A1:Z50000")
  .rows(["Strategy", "Book"])
  .cols(["Month"])
  .filters(["Region", "PM"])
  .values([
      sum("Gross MV").name("Gross"),
      sum("Net MV").name("Net"),
      avg("Leverage").name("Avg Lev"),
  ])
  .validate_fields()?
  .place("A4");
```

- Fixed quoted sheet-name resolution in pivot source handling.
- Fixed `pageField@fld` to use source/cache field index.
- Replaced hardcoded pivot item counts with data-derived unique counts.
- Workbook-level fluent pivot builder: `Workbook::pivot_on(sheet, name)...validate_fields().place(...)`
- Cross-sheet source header validation.

---

### P0.4 Finance Chart Templates

Status: `done`

High-level chart presets for finance reporting: `PnlCurve`, `DrawdownCurve`, `ExposureBars`, `FactorContribStacked`, `WaterfallBridge`.

```rust
ws.chart_template(ChartTemplate::PnlCurve)
  .title("Fund NAV & Drawdown")
  .x("Summary!A2:A253")
  .y("Summary!B2:B253")
  .secondary_y("Summary!C2:C253")
  .place("D2:L20");
```

---

### P0.5 Workbook Lint/Validation Framework

Status: `done`

Preflight checks for broken refs, inconsistent formulas, stale caches, invalid pivot/chart links.

```rust
let report = wb.lint()
  .check_broken_refs()
  .check_formula_consistency()
  .check_pivot_sources()
  .check_named_ranges()
  .run();
```

- Structured report: `LintReport`, `LintFinding`, `LintSeverity`, `LintLocation`
- CLI command: `ofx lint <file.xlsx>`
