# BUG-007: Auto Color Always Black

**Priority**: MEDIUM
**Status**: Open
**Component**: color resolution (adapter/render layer)

## Problem

When a color element has `auto="true"`, the resolver always returns black (#000000). In Excel, "auto" color is context-dependent and should use theme colors.

## Current Behavior

Auto color always returns black, regardless of context.

## Expected Behavior

- For **text/foreground**: Use theme color `dk1` (dark 1)
- For **background/fill**: Use theme color `lt1` (light 1)
- These adapt to the workbook's theme (could be dark mode, custom branding, etc.)

## Where Data Should Come From

```xml
<!-- xl/theme/theme1.xml -->
<a:clrScheme name="Office">
  <a:dk1><a:sysClr val="windowText"/></a:dk1>  <!-- Usually black -->
  <a:lt1><a:sysClr val="window"/></a:lt1>      <!-- Usually white -->
</a:clrScheme>
```

## Impact

- In dark-themed workbooks, auto text color might be wrong
- Auto colors may not match Excel's rendering
- Custom branded templates may have contrast issues

## Proposed Fix

1. Add context parameter to color resolution (is this for text or background?)
2. For text contexts, resolve auto to theme `dk1`
3. For background contexts, resolve auto to theme `lt1`

```rust
fn resolve_auto_color(context: ColorContext, theme: &Theme) -> String {
    match context {
        ColorContext::Text | ColorContext::Border => theme.colors[0].clone(), // dk1
        ColorContext::Fill => theme.colors[1].clone(), // lt1
    }
}
```

## References

- ECMA-376 Part 1, Section 18.8.3 (color auto attribute)
