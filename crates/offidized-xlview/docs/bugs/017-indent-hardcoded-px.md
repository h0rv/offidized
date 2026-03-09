# BUG-017: Indent Uses Hardcoded Pixels

**Priority**: MEDIUM
**Status**: Open
**Component**: render layer

## Problem

Cell indent uses a fixed pixel multiplier rather than Excel's character-based indent units derived from the default font.

## Current Behavior

Indent uses fixed pixel padding rather than font-relative units.

## Expected Behavior

Indent should be based on the width of one character in the default font (max digit width), as Excel defines indent in character units.

## Where Data Should Come From

```xml
<alignment indent="2"/>
```

## Impact

- Indent spacing is too large or too small for non-Default fonts
- Visual alignment differs from Excel

## Proposed Fix

1. Compute per-font max digit width (or use measured text metrics).
2. Convert indent levels to pixel padding using that width.
3. Fall back to a reasonable default only when font metrics are unavailable.

## References

- ECMA-376 Part 1, Section 18.8.1 (alignment)
