# BUG-015: Conditional Formatting Thresholds Ignored

**Priority**: MEDIUM
**Status**: Open
**Component**: conditional formatting (render layer)

## Problem

Conditional formatting rules are parsed, but evaluation is simplified. The renderer uses a min/max of observed numeric values in the range and ignores `cfvo` thresholds, formulas, and many rule settings. Icon sets are rendered as emoji, not Excel glyphs.

## Current Behavior

- Renderer computes `minVal`/`maxVal` from numeric cells and applies:
  - color scales with simple linear interpolation
  - data bars based on min/max
  - icon sets based on percent (emoji icons)
- `cfvo` types (`min`, `max`, `num`, `percent`, `percentile`, `formula`) are ignored.
- `dxfId`/style-based conditional formatting is not applied.

## Expected Behavior

- Evaluate `cfvo` thresholds per ECMA-376 (min/max/percent/percentile/num/formula)
- Apply icon set thresholds and flags like `reverse`, `showValue`
- Apply `dxfId` formatting when rule type requires it
- Use Excel-like icon glyphs (or a closer visual match) instead of emoji

## Where Data Should Come From

```xml
<cfRule type="iconSet" priority="1">
  <iconSet iconSet="3TrafficLights1" showValue="1" reverse="0">
    <cfvo type="percent" val="33"/>
    <cfvo type="percent" val="67"/>
  </iconSet>
</cfRule>
```

## Impact

- Conditional formatting often diverges from Excel
- Visual signals (icons/bars) can be wrong for the same data

## Proposed Fix

1. Parse and carry `cfvo` thresholds into the renderer.
2. Implement threshold evaluation logic for each `cfvo` type.
3. Apply `dxfId` formatting for rule types that require direct styling.
4. Replace emoji icons with a closer Excel-compatible icon set (SVG or sprite).

## References

- ECMA-376 Part 1, Section 18.3.1.10 (conditionalFormatting)
- ECMA-376 Part 1, Section 18.3.1.18 (cfvo)
