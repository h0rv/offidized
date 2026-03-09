# BUG-008: Border Color Default Wrong

**Priority**: MEDIUM
**Status**: Open
**Component**: color resolution (adapter/render layer)

## Problem

When a border has no explicit color, the resolver defaults to black (#000000). Excel's default border color is actually the theme's `dk1` color (which is usually black, but could differ in custom themes).

## Current Behavior

Hardcoded black fallback for missing border colors.

## Expected Behavior

Use theme color `dk1` for default/automatic border colors.

## Where Data Should Come From

```xml
<!-- Border with automatic color (index 64) -->
<border>
  <left style="thin">
    <color indexed="64"/>  <!-- 64 = automatic = use dk1 -->
  </left>
</border>

<!-- Border with no color element (implicit automatic) -->
<border>
  <left style="thin"/>  <!-- No color = automatic = use dk1 -->
</border>
```

## Impact

- Borders may have wrong color in custom-themed workbooks
- Dark-themed workbooks may have invisible borders

## Proposed Fix

1. When border has no color, use theme `dk1` instead of hardcoded black
2. When border has `indexed="64"`, also use theme `dk1`
3. Pass theme reference to border resolution

## References

- ECMA-376 Part 1, Section 18.8.4 (border color)
- Index 64 is defined as "system foreground" per legacy spec
