# BUG-011: Color Transforms Not Parsed

**Priority**: MEDIUM
**Status**: Open
**Component**: color resolution (adapter/render layer)

## Problem

Excel/Office XML supports color transformations like tint, shade, saturation modulation, etc. These are specified as child elements of color definitions but are currently ignored.

## Current Behavior

```xml
<a:srgbClr val="4472C4">
  <a:lumMod val="75000"/>  <!-- 75% luminance - IGNORED -->
  <a:lumOff val="25000"/>  <!-- +25% luminance offset - IGNORED -->
</a:srgbClr>
```

The base color #4472C4 is used, but the luminance modification is ignored.

## Expected Behavior

Apply all color transforms in order:

- `<a:tint val="X"/>` - Tint toward white (X/100000 blend)
- `<a:shade val="X"/>` - Shade toward black
- `<a:lumMod val="X"/>` - Multiply luminance by X/100000
- `<a:lumOff val="X"/>` - Add X/100000 to luminance
- `<a:satMod val="X"/>` - Multiply saturation
- `<a:satOff val="X"/>` - Add to saturation
- `<a:alpha val="X"/>` - Set alpha/transparency
- `<a:hueMod val="X"/>` - Modify hue
- `<a:hueOff val="X"/>` - Offset hue

## Where Data Should Come From

Theme colors and any color element can have transform children:

```xml
<!-- Theme color with transform -->
<a:accent1>
  <a:srgbClr val="4472C4">
    <a:lumMod val="60000"/>
    <a:lumOff val="40000"/>
  </a:srgbClr>
</a:accent1>

<!-- Cell color with tint -->
<fgColor theme="4" tint="0.39997558519241921"/>
```

Note: The `tint` attribute on `<fgColor>` IS currently handled, but transforms inside color elements are not.

## Impact

- Theme colors with modifications display incorrectly
- Subtle color variations in professional templates are lost
- Color schemes may look flat/wrong

## Proposed Fix

1. When parsing color elements, check for child transform elements
2. Apply transforms in document order
3. Convert to HSL, apply transforms, convert back to RGB

```rust
fn apply_color_transforms(base_rgb: &str, transforms: &[ColorTransform]) -> String {
    let (h, s, l) = rgb_to_hsl(base_rgb);

    for transform in transforms {
        match transform {
            ColorTransform::LumMod(val) => l *= val / 100000.0,
            ColorTransform::LumOff(val) => l += val / 100000.0,
            ColorTransform::SatMod(val) => s *= val / 100000.0,
            // ... etc
        }
    }

    hsl_to_rgb(h, s.clamp(0.0, 1.0), l.clamp(0.0, 1.0))
}
```

## References

- ECMA-376 Part 1, Section 20.1.2.3 (Color Transforms)
- DrawingML color model documentation
