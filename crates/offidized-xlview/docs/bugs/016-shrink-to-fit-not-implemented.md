# BUG-016: Shrink-To-Fit Not Implemented

**Priority**: MEDIUM
**Status**: Open
**Component**: render layer

## Problem

`shrinkToFit` is treated as overflow clipping rather than reducing font size to fit the cell width.

## Current Behavior

Text is clipped at cell boundaries instead of scaled down to fit.

## Expected Behavior

When `shrinkToFit` is true, Excel reduces the font size (or scales text) until the content fits within the cell width.

## Where Data Should Come From

- Alignment element in `xl/styles.xml`:

```xml
<alignment shrinkToFit="1"/>
```

## Impact

- Text appears truncated rather than scaled
- Layout diverges from Excel in common shrink-to-fit scenarios

## Proposed Fix

1. Measure rendered text width for the cell font.
2. Compute a scale factor (or reduced font size) to fit within the cell width.
3. Apply scaled font size or CSS transform, respecting min font size limits.

## References

- ECMA-376 Part 1, Section 18.8.1 (alignment)
