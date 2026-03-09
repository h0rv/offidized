# BUG-010: baseColWidth Not Parsed

**Priority**: MEDIUM
**Status**: Open
**Component**: layout / adapter layer

## Problem

The `<sheetFormatPr>` element has a `baseColWidth` attribute that defines the base column width in characters. This is used when `defaultColWidth` is not specified. Currently not parsed.

## Current Behavior

Only `defaultColWidth` is parsed. If absent, it falls back to a hardcoded width
(~8.43 chars / ~64px) set during sheet initialization.

## Expected Behavior

1. Parse `baseColWidth` attribute
2. If `defaultColWidth` is absent but `baseColWidth` is present, use `baseColWidth` for default width calculation
3. The relationship is: `defaultColWidth ~ baseColWidth + padding`

## Where Data Should Come From

```xml
<!-- xl/worksheets/sheet1.xml -->
<sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
```

Per ECMA-376:

- `baseColWidth`: Number of characters of maximum digit width (default 8)
- `defaultColWidth`: Default column width in character units
- If `defaultColWidth` absent, Excel calculates it from `baseColWidth`

## Impact

- Default column width calculation may be wrong when only `baseColWidth` is specified
- Columns may be slightly too narrow or wide

## Proposed Fix

1. Parse `baseColWidth` attribute in `sheetFormatPr` handler
2. If `defaultColWidth` is absent, calculate it:
   ```rust
   // Excel formula (approximately):
   // defaultColWidth = truncate((baseColWidth * maxDigitWidth + 5) / maxDigitWidth * 256) / 256
   ```
3. Use calculated value as default column width

## References

- ECMA-376 Part 1, Section 18.3.1.81 (sheetFormatPr)
- [MS-OI29500] Section 2.1.612
