//! xlsx style mode: derive and apply cell formatting/layout properties.
//!
//! Style mode emits formatting without touching cell content. Properties use
//! additive semantics: properties present in the IR are applied; properties
//! absent are left unchanged.

use crate::{ApplyResult, DeriveOptions, IrError, Result};
use offidized_xlsx::{
    style::{
        BorderSide, ColorReference, HorizontalAlignment, PatternFill, PatternFillType, Style,
        ThemeColor, VerticalAlignment,
    },
    Workbook,
};

// ---------------------------------------------------------------------------
// Color helpers
// ---------------------------------------------------------------------------

/// Serialize a `ColorReference` to IR string.
///
/// Priority: rgb > theme+tint > theme > indexed > auto.
pub(crate) fn serialize_color(color: &ColorReference) -> String {
    if let Some(rgb) = color.rgb() {
        // Strip leading "FF" alpha prefix if present (ARGB → RGB)
        let hex = if rgb.len() == 8 && rgb.get(..2).is_some_and(|a| a.eq_ignore_ascii_case("FF")) {
            &rgb[2..]
        } else {
            rgb
        };
        return format!("#{hex}");
    }
    if let Some(theme) = color.theme() {
        let name = theme_to_ir_name(theme);
        if let Some(tint) = color.tint() {
            if tint >= 0.0 {
                return format!("theme:{name}+{tint:.2}");
            }
            return format!("theme:{name}{tint:.2}");
        }
        return format!("theme:{name}");
    }
    if let Some(indexed) = color.indexed() {
        return format!("indexed:{indexed}");
    }
    if color.auto() == Some(true) {
        return "auto".to_string();
    }
    // Fallback: empty color reference
    "auto".to_string()
}

/// Parse an IR color string to `ColorReference`.
pub(crate) fn parse_color(s: &str) -> Result<ColorReference> {
    let s = s.trim();

    if s == "auto" {
        let mut cr = ColorReference::empty();
        cr.set_auto(true);
        return Ok(cr);
    }

    if let Some(rest) = s.strip_prefix('#') {
        // #RRGGBB → FFRRGGBB (add alpha)
        let argb = if rest.len() == 6 {
            format!("FF{rest}")
        } else {
            rest.to_string()
        };
        return Ok(ColorReference::from_rgb(argb));
    }

    if let Some(rest) = s.strip_prefix("theme:") {
        // theme:name, theme:name+0.40, theme:name-0.25
        if let Some(plus_pos) = rest.find('+') {
            let name = &rest[..plus_pos];
            let tint_str = &rest[plus_pos + 1..];
            let theme = ir_name_to_theme(name)
                .ok_or_else(|| IrError::InvalidBody(format!("unknown theme color: {name}")))?;
            let tint: f64 = tint_str
                .parse()
                .map_err(|_| IrError::InvalidBody(format!("invalid tint: {tint_str}")))?;
            return Ok(ColorReference::from_theme_with_tint(theme, tint));
        }
        // Check for negative tint (theme:name-0.25)
        // Find the last '-' that's not at position 0
        if let Some(minus_pos) = rest.rfind('-') {
            if minus_pos > 0 {
                let name = &rest[..minus_pos];
                if let Some(theme) = ir_name_to_theme(name) {
                    let tint_str = &rest[minus_pos..]; // includes the -
                    if let Ok(tint) = tint_str.parse::<f64>() {
                        return Ok(ColorReference::from_theme_with_tint(theme, tint));
                    }
                }
            }
        }
        // Plain theme:name
        let theme = ir_name_to_theme(rest)
            .ok_or_else(|| IrError::InvalidBody(format!("unknown theme color: {rest}")))?;
        return Ok(ColorReference::from_theme(theme));
    }

    if let Some(rest) = s.strip_prefix("indexed:") {
        let idx: u32 = rest
            .parse()
            .map_err(|_| IrError::InvalidBody(format!("invalid indexed color: {rest}")))?;
        return Ok(ColorReference::from_indexed(idx));
    }

    Err(IrError::InvalidBody(format!("unknown color syntax: {s}")))
}

fn theme_to_ir_name(theme: ThemeColor) -> &'static str {
    match theme {
        ThemeColor::Dark1 => "dark1",
        ThemeColor::Light1 => "light1",
        ThemeColor::Dark2 => "dark2",
        ThemeColor::Light2 => "light2",
        ThemeColor::Accent1 => "accent1",
        ThemeColor::Accent2 => "accent2",
        ThemeColor::Accent3 => "accent3",
        ThemeColor::Accent4 => "accent4",
        ThemeColor::Accent5 => "accent5",
        ThemeColor::Accent6 => "accent6",
        ThemeColor::Hyperlink => "hyperlink",
        ThemeColor::FollowedHyperlink => "followedHyperlink",
    }
}

fn ir_name_to_theme(name: &str) -> Option<ThemeColor> {
    match name {
        "dark1" => Some(ThemeColor::Dark1),
        "light1" => Some(ThemeColor::Light1),
        "dark2" => Some(ThemeColor::Dark2),
        "light2" => Some(ThemeColor::Light2),
        "accent1" => Some(ThemeColor::Accent1),
        "accent2" => Some(ThemeColor::Accent2),
        "accent3" => Some(ThemeColor::Accent3),
        "accent4" => Some(ThemeColor::Accent4),
        "accent5" => Some(ThemeColor::Accent5),
        "accent6" => Some(ThemeColor::Accent6),
        "hyperlink" => Some(ThemeColor::Hyperlink),
        "followedHyperlink" => Some(ThemeColor::FollowedHyperlink),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Property layer
// ---------------------------------------------------------------------------

/// Split a comma-separated property string, respecting quoted values.
fn split_properties(s: &str) -> Vec<&str> {
    let mut result = Vec::new();
    let mut start = 0;
    let mut in_quotes = false;
    let bytes = s.as_bytes();

    for i in 0..bytes.len() {
        match bytes[i] {
            b'"' => in_quotes = !in_quotes,
            b',' if !in_quotes => {
                let part = s[start..i].trim();
                if !part.is_empty() {
                    result.push(part);
                }
                start = i + 1;
            }
            _ => {}
        }
    }

    let last = s[start..].trim();
    if !last.is_empty() {
        result.push(last);
    }

    result
}

/// Convert a `Style` to IR property strings.
fn style_to_properties(style: &Style) -> Vec<String> {
    let mut props = Vec::new();

    // Font properties
    if let Some(font) = style.font() {
        if font.bold() == Some(true) {
            props.push("bold".to_string());
        }
        if font.italic() == Some(true) {
            props.push("italic".to_string());
        }
        if font.underline() == Some(true) {
            props.push("underline".to_string());
        }
        if font.strikethrough() == Some(true) {
            props.push("strikethrough".to_string());
        }
        if let Some(name) = font.name() {
            props.push(format!("font=\"{name}\""));
        }
        if let Some(size) = font.size() {
            props.push(format!("size={size}"));
        }
        if let Some(color_ref) = font.color_ref() {
            props.push(format!("font-color={}", serialize_color(color_ref)));
        } else if let Some(color) = font.color() {
            // Legacy color string (ARGB)
            let hex = if color.len() == 8
                && color.get(..2).is_some_and(|a| a.eq_ignore_ascii_case("FF"))
            {
                &color[2..]
            } else {
                color
            };
            props.push(format!("font-color=#{hex}"));
        }
    }

    // Fill properties
    if let Some(fill) = style.fill() {
        if let Some(pf) = fill.pattern_fill() {
            if let Some(fg_ref) = pf.fg_color() {
                props.push(format!("fill={}", serialize_color(fg_ref)));
            } else if let Some(fg) = fill.foreground_color() {
                let hex =
                    if fg.len() == 8 && fg.get(..2).is_some_and(|a| a.eq_ignore_ascii_case("FF")) {
                        &fg[2..]
                    } else {
                        fg
                    };
                props.push(format!("fill=#{hex}"));
            }
            let pat = pf.pattern_type();
            if pat != PatternFillType::Solid && pat != PatternFillType::None {
                props.push(format!("fill-pattern={}", pattern_fill_to_ir(pat)));
            }
            if let Some(bg_ref) = pf.bg_color() {
                props.push(format!("fill-bg={}", serialize_color(bg_ref)));
            }
        } else {
            if let Some(fg) = fill.foreground_color() {
                let hex =
                    if fg.len() == 8 && fg.get(..2).is_some_and(|a| a.eq_ignore_ascii_case("FF")) {
                        &fg[2..]
                    } else {
                        fg
                    };
                props.push(format!("fill=#{hex}"));
            }
            if let Some(pat) = fill.pattern() {
                if pat != "solid" {
                    props.push(format!("fill-pattern={pat}"));
                }
            }
            if let Some(bg) = fill.background_color() {
                let hex =
                    if bg.len() == 8 && bg.get(..2).is_some_and(|a| a.eq_ignore_ascii_case("FF")) {
                        &bg[2..]
                    } else {
                        bg
                    };
                props.push(format!("fill-bg=#{hex}"));
            }
        }
    }

    // Border properties
    if let Some(border) = style.border() {
        emit_border_side(&mut props, "border-top", border.top());
        emit_border_side(&mut props, "border-bottom", border.bottom());
        emit_border_side(&mut props, "border-left", border.left());
        emit_border_side(&mut props, "border-right", border.right());
    }

    // Alignment
    if let Some(align) = style.alignment() {
        if let Some(h) = align.horizontal() {
            props.push(format!("align={}", horizontal_to_ir(h)));
        }
        if let Some(v) = align.vertical() {
            props.push(format!("valign={}", vertical_to_ir(v)));
        }
        if align.wrap_text() == Some(true) {
            props.push("wrap".to_string());
        }
        if let Some(indent) = align.indent() {
            if indent > 0 {
                props.push(format!("indent={indent}"));
            }
        }
        if let Some(rotation) = align.text_rotation() {
            if rotation > 0 {
                props.push(format!("rotation={rotation}"));
            }
        }
        if align.shrink_to_fit() == Some(true) {
            props.push("shrink-to-fit".to_string());
        }
    }

    // Number format
    if let Some(fmt) = style.custom_format() {
        props.push(format!("format=\"{fmt}\""));
    } else if let Some(nf) = style.number_format() {
        props.push(format!("format=\"{nf}\""));
    }

    // Protection
    if let Some(prot) = style.protection() {
        if prot.locked() == Some(true) {
            props.push("locked".to_string());
        }
        if prot.hidden() == Some(true) {
            props.push("cell-hidden".to_string());
        }
    }

    props
}

fn emit_border_side(props: &mut Vec<String>, prefix: &str, side: Option<&BorderSide>) {
    let Some(side) = side else { return };
    let Some(style_name) = side.style() else {
        return;
    };
    if style_name == "none" {
        return;
    }
    if let Some(color_ref) = side.color_ref() {
        props.push(format!(
            "{prefix}={style_name} {}",
            serialize_color(color_ref)
        ));
    } else if let Some(color) = side.color() {
        let hex =
            if color.len() == 8 && color.get(..2).is_some_and(|a| a.eq_ignore_ascii_case("FF")) {
                &color[2..]
            } else {
                color
            };
        props.push(format!("{prefix}={style_name} #{hex}"));
    } else {
        props.push(format!("{prefix}={style_name}"));
    }
}

fn horizontal_to_ir(h: HorizontalAlignment) -> &'static str {
    match h {
        HorizontalAlignment::General => "general",
        HorizontalAlignment::Left => "left",
        HorizontalAlignment::Center => "center",
        HorizontalAlignment::Right => "right",
        HorizontalAlignment::Fill => "fill",
        HorizontalAlignment::Justify => "justify",
        HorizontalAlignment::CenterContinuous => "centerContinuous",
        HorizontalAlignment::Distributed => "distributed",
    }
}

fn ir_to_horizontal(s: &str) -> Option<HorizontalAlignment> {
    match s {
        "general" => Some(HorizontalAlignment::General),
        "left" => Some(HorizontalAlignment::Left),
        "center" => Some(HorizontalAlignment::Center),
        "right" => Some(HorizontalAlignment::Right),
        "fill" => Some(HorizontalAlignment::Fill),
        "justify" => Some(HorizontalAlignment::Justify),
        "centerContinuous" => Some(HorizontalAlignment::CenterContinuous),
        "distributed" => Some(HorizontalAlignment::Distributed),
        _ => None,
    }
}

fn vertical_to_ir(v: VerticalAlignment) -> &'static str {
    match v {
        VerticalAlignment::Top => "top",
        VerticalAlignment::Center => "center",
        VerticalAlignment::Bottom => "bottom",
        VerticalAlignment::Justify => "justify",
        VerticalAlignment::Distributed => "distributed",
    }
}

fn ir_to_vertical(s: &str) -> Option<VerticalAlignment> {
    match s {
        "top" => Some(VerticalAlignment::Top),
        "center" => Some(VerticalAlignment::Center),
        "bottom" => Some(VerticalAlignment::Bottom),
        "justify" => Some(VerticalAlignment::Justify),
        "distributed" => Some(VerticalAlignment::Distributed),
        _ => None,
    }
}

fn pattern_fill_to_ir(pat: PatternFillType) -> &'static str {
    match pat {
        PatternFillType::None => "none",
        PatternFillType::Solid => "solid",
        PatternFillType::DarkDown => "darkDown",
        PatternFillType::DarkGray => "darkGray",
        PatternFillType::DarkGrid => "darkGrid",
        PatternFillType::DarkHorizontal => "darkHorizontal",
        PatternFillType::DarkTrellis => "darkTrellis",
        PatternFillType::DarkUp => "darkUp",
        PatternFillType::DarkVertical => "darkVertical",
        PatternFillType::Gray0625 => "gray0625",
        PatternFillType::Gray125 => "gray125",
        PatternFillType::LightDown => "lightDown",
        PatternFillType::LightGray => "lightGray",
        PatternFillType::LightGrid => "lightGrid",
        PatternFillType::LightHorizontal => "lightHorizontal",
        PatternFillType::LightTrellis => "lightTrellis",
        PatternFillType::LightUp => "lightUp",
        PatternFillType::LightVertical => "lightVertical",
        PatternFillType::MediumGray => "mediumGray",
    }
}

fn ir_to_pattern_fill(s: &str) -> Option<PatternFillType> {
    match s {
        "none" => Some(PatternFillType::None),
        "solid" => Some(PatternFillType::Solid),
        "darkDown" => Some(PatternFillType::DarkDown),
        "darkGray" => Some(PatternFillType::DarkGray),
        "darkGrid" => Some(PatternFillType::DarkGrid),
        "darkHorizontal" => Some(PatternFillType::DarkHorizontal),
        "darkTrellis" => Some(PatternFillType::DarkTrellis),
        "darkUp" => Some(PatternFillType::DarkUp),
        "darkVertical" => Some(PatternFillType::DarkVertical),
        "gray0625" => Some(PatternFillType::Gray0625),
        "gray125" => Some(PatternFillType::Gray125),
        "lightDown" => Some(PatternFillType::LightDown),
        "lightGray" => Some(PatternFillType::LightGray),
        "lightGrid" => Some(PatternFillType::LightGrid),
        "lightHorizontal" => Some(PatternFillType::LightHorizontal),
        "lightTrellis" => Some(PatternFillType::LightTrellis),
        "lightUp" => Some(PatternFillType::LightUp),
        "lightVertical" => Some(PatternFillType::LightVertical),
        "mediumGray" => Some(PatternFillType::MediumGray),
        _ => None,
    }
}

/// Apply parsed property strings to a mutable `Style`.
fn apply_properties_to_style(style: &mut Style, props: &[&str]) -> Result<()> {
    for prop in props {
        let prop = prop.trim();
        if prop.is_empty() {
            continue;
        }

        // Boolean flags
        match prop {
            "bold" => {
                let mut font = style.font().cloned().unwrap_or_default();
                font.set_bold(true);
                style.set_font(font);
                continue;
            }
            "italic" => {
                let mut font = style.font().cloned().unwrap_or_default();
                font.set_italic(true);
                style.set_font(font);
                continue;
            }
            "underline" => {
                let mut font = style.font().cloned().unwrap_or_default();
                font.set_underline(true);
                style.set_font(font);
                continue;
            }
            "strikethrough" => {
                let mut font = style.font().cloned().unwrap_or_default();
                font.set_strikethrough(true);
                style.set_font(font);
                continue;
            }
            "wrap" => {
                let mut align = style.alignment().cloned().unwrap_or_default();
                align.set_wrap_text(true);
                style.set_alignment(align);
                continue;
            }
            "shrink-to-fit" => {
                let mut align = style.alignment().cloned().unwrap_or_default();
                align.set_shrink_to_fit(true);
                style.set_alignment(align);
                continue;
            }
            "locked" => {
                let mut prot = style.protection().cloned().unwrap_or_default();
                prot.set_locked(true);
                style.set_protection(prot);
                continue;
            }
            "cell-hidden" => {
                let mut prot = style.protection().cloned().unwrap_or_default();
                prot.set_hidden(true);
                style.set_protection(prot);
                continue;
            }
            _ => {}
        }

        // Key=value properties
        if let Some((key, value)) = prop.split_once('=') {
            let key = key.trim();
            let value = value.trim();

            match key {
                "font" => {
                    let name = value.trim_matches('"');
                    let mut font = style.font().cloned().unwrap_or_default();
                    font.set_name(name);
                    style.set_font(font);
                }
                "size" => {
                    let mut font = style.font().cloned().unwrap_or_default();
                    font.set_size(value);
                    style.set_font(font);
                }
                "font-color" => {
                    let color_ref = parse_color(value)?;
                    let mut font = style.font().cloned().unwrap_or_default();
                    font.set_color_ref(color_ref);
                    style.set_font(font);
                }
                "fill" => {
                    let color_ref = parse_color(value)?;
                    let mut fill = style.fill().cloned().unwrap_or_default();
                    let mut pf = fill
                        .pattern_fill()
                        .cloned()
                        .unwrap_or_else(|| PatternFill::new(PatternFillType::Solid));
                    pf.set_fg_color(color_ref);
                    if pf.pattern_type() == PatternFillType::None {
                        pf.set_pattern_type(PatternFillType::Solid);
                    }
                    fill.set_pattern_fill(pf);
                    style.set_fill(fill);
                }
                "fill-pattern" => {
                    if let Some(pat) = ir_to_pattern_fill(value) {
                        let mut fill = style.fill().cloned().unwrap_or_default();
                        let mut pf = fill
                            .pattern_fill()
                            .cloned()
                            .unwrap_or_else(|| PatternFill::new(PatternFillType::None));
                        pf.set_pattern_type(pat);
                        fill.set_pattern_fill(pf);
                        style.set_fill(fill);
                    }
                }
                "fill-bg" => {
                    let color_ref = parse_color(value)?;
                    let mut fill = style.fill().cloned().unwrap_or_default();
                    let mut pf = fill
                        .pattern_fill()
                        .cloned()
                        .unwrap_or_else(|| PatternFill::new(PatternFillType::None));
                    pf.set_bg_color(color_ref);
                    fill.set_pattern_fill(pf);
                    style.set_fill(fill);
                }
                "border-top" => {
                    let (bs_style, bs_color) = parse_border_value(value)?;
                    let mut border = style.border().cloned().unwrap_or_default();
                    let mut side = BorderSide::new();
                    side.set_style(bs_style);
                    if let Some(c) = bs_color {
                        side.set_color_ref(c);
                    }
                    border.set_top(side);
                    style.set_border(border);
                }
                "border-bottom" => {
                    let (bs_style, bs_color) = parse_border_value(value)?;
                    let mut border = style.border().cloned().unwrap_or_default();
                    let mut side = BorderSide::new();
                    side.set_style(bs_style);
                    if let Some(c) = bs_color {
                        side.set_color_ref(c);
                    }
                    border.set_bottom(side);
                    style.set_border(border);
                }
                "border-left" => {
                    let (bs_style, bs_color) = parse_border_value(value)?;
                    let mut border = style.border().cloned().unwrap_or_default();
                    let mut side = BorderSide::new();
                    side.set_style(bs_style);
                    if let Some(c) = bs_color {
                        side.set_color_ref(c);
                    }
                    border.set_left(side);
                    style.set_border(border);
                }
                "border-right" => {
                    let (bs_style, bs_color) = parse_border_value(value)?;
                    let mut border = style.border().cloned().unwrap_or_default();
                    let mut side = BorderSide::new();
                    side.set_style(bs_style);
                    if let Some(c) = bs_color {
                        side.set_color_ref(c);
                    }
                    border.set_right(side);
                    style.set_border(border);
                }
                "align" => {
                    if let Some(h) = ir_to_horizontal(value) {
                        let mut align = style.alignment().cloned().unwrap_or_default();
                        align.set_horizontal(h);
                        style.set_alignment(align);
                    }
                }
                "valign" => {
                    if let Some(v) = ir_to_vertical(value) {
                        let mut align = style.alignment().cloned().unwrap_or_default();
                        align.set_vertical(v);
                        style.set_alignment(align);
                    }
                }
                "indent" => {
                    if let Ok(n) = value.parse::<u32>() {
                        let mut align = style.alignment().cloned().unwrap_or_default();
                        align.set_indent(n);
                        style.set_alignment(align);
                    }
                }
                "rotation" => {
                    if let Ok(n) = value.parse::<u32>() {
                        let mut align = style.alignment().cloned().unwrap_or_default();
                        align.set_text_rotation(n);
                        style.set_alignment(align);
                    }
                }
                "format" => {
                    let fmt = value.trim_matches('"');
                    style.set_custom_format(fmt);
                }
                _ => {
                    // Unknown property — ignore silently
                }
            }
        }
    }

    Ok(())
}

/// Parse border value like "thin #000000" → (style, Option<ColorReference>).
fn parse_border_value(value: &str) -> Result<(String, Option<ColorReference>)> {
    let parts: Vec<&str> = value.splitn(2, ' ').collect();
    let style_name = parts[0].to_string();
    let color = if parts.len() > 1 {
        Some(parse_color(parts[1])?)
    } else {
        None
    };
    Ok((style_name, color))
}

// ---------------------------------------------------------------------------
// Column/cell ref helpers
// ---------------------------------------------------------------------------

/// Convert a 1-based column index to letter(s): 1→A, 26→Z, 27→AA.
fn column_index_to_letter(mut index: u32) -> String {
    let mut result = String::new();
    while index > 0 {
        index -= 1;
        let c = (b'A' + (index % 26) as u8) as char;
        result.insert(0, c);
        index /= 26;
    }
    result
}

/// Parse (row, col) from a cell reference for sorting in row-major order.
fn cell_sort_key(cell_ref: &str) -> (u32, u32) {
    let col_end = cell_ref
        .find(|c: char| c.is_ascii_digit())
        .unwrap_or(cell_ref.len());
    let col_str = &cell_ref[..col_end];
    let row_str = &cell_ref[col_end..];

    let col = col_str.bytes().fold(0u32, |acc, b| {
        acc * 26 + u32::from(b.to_ascii_uppercase() - b'A') + 1
    });
    let row: u32 = row_str.parse().unwrap_or(0);
    (row, col)
}

// ---------------------------------------------------------------------------
// Derive
// ---------------------------------------------------------------------------

/// Append the xlsx style-mode body to `output`.
pub(crate) fn derive_style(wb: &Workbook, options: &DeriveOptions, output: &mut String) {
    for ws in wb.worksheets() {
        // Filter by sheet name if specified
        if let Some(ref sheet_filter) = options.sheet {
            if ws.name() != sheet_filter.as_str() {
                continue;
            }
        }

        output.push('\n');
        output.push_str("=== Sheet: ");
        output.push_str(ws.name());
        output.push_str(" ===\n");

        // Sheet properties
        let mut sheet_prop_lines = Vec::new();

        if let Some(tab_color) = ws.tab_color() {
            sheet_prop_lines.push(format!("tab-color: #{tab_color}"));
        }

        if let Some(freeze) = ws.freeze_pane() {
            let col = column_index_to_letter(freeze.x_split() + 1);
            let row = freeze.y_split() + 1;
            sheet_prop_lines.push(format!("freeze: {col}{row}"));
        }

        if let Some(view_opts) = ws.sheet_view_options() {
            if let Some(show) = view_opts.show_gridlines() {
                if show {
                    sheet_prop_lines.push("gridlines: visible".to_string());
                } else {
                    sheet_prop_lines.push("gridlines: hidden".to_string());
                }
            }
            if let Some(zoom) = view_opts.zoom_scale() {
                sheet_prop_lines.push(format!("zoom: {zoom}"));
            }
        }

        if !sheet_prop_lines.is_empty() {
            output.push_str("\n# Sheet properties\n");
            for line in &sheet_prop_lines {
                output.push_str(line);
                output.push('\n');
            }
        }

        // Columns with non-default properties
        let cols: Vec<_> = ws.columns().collect();
        if !cols.is_empty() {
            let mut col_lines = Vec::new();
            for col in &cols {
                let mut parts = Vec::new();
                if let Some(w) = col.width() {
                    parts.push(format!("width={w}"));
                }
                if col.is_hidden() {
                    parts.push("hidden".to_string());
                }
                if !parts.is_empty() {
                    let letter = column_index_to_letter(col.index());
                    col_lines.push(format!("col {letter}: {}", parts.join(", ")));
                }
            }
            if !col_lines.is_empty() {
                output.push_str("\n# Columns\n");
                for line in &col_lines {
                    output.push_str(line);
                    output.push('\n');
                }
            }
        }

        // Rows with non-default properties
        let mut row_lines = Vec::new();
        for row in ws.rows() {
            let mut parts = Vec::new();
            if let Some(h) = row.height() {
                parts.push(format!("height={h}"));
            }
            if row.is_hidden() {
                parts.push("hidden".to_string());
            }
            if !parts.is_empty() {
                row_lines.push(format!("row {}: {}", row.index(), parts.join(", ")));
            }
        }
        if !row_lines.is_empty() {
            output.push_str("\n# Rows\n");
            for line in &row_lines {
                output.push_str(line);
                output.push('\n');
            }
        }

        // Cell styles
        let mut cell_lines = Vec::new();
        let mut cells: Vec<(&str, &offidized_xlsx::Cell)> = ws.cells().collect();
        cells.sort_by_key(|(cell_ref, _)| cell_sort_key(cell_ref));

        for (cell_ref, cell) in cells {
            let style_id = cell.style_id().unwrap_or(0);
            if style_id == 0 {
                continue;
            }
            if let Some(style) = wb.style(style_id) {
                let props = style_to_properties(style);
                if !props.is_empty() {
                    cell_lines.push(format!("{cell_ref}: {}", props.join(", ")));
                }
            }
        }
        if !cell_lines.is_empty() {
            output.push_str("\n# Cell styles\n");
            for line in &cell_lines {
                output.push_str(line);
                output.push('\n');
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Apply
// ---------------------------------------------------------------------------

/// Apply the style IR body to a workbook, updating formatting.
pub(crate) fn apply_style(body: &str, wb: &mut Workbook) -> Result<ApplyResult> {
    let mut result = ApplyResult::default();
    let mut current_sheet: Option<String> = None;

    for line in body.lines() {
        let line = line.trim_end_matches('\r');

        // Sheet header: === Sheet: NAME ===
        if let Some(name) = parse_sheet_header(line) {
            current_sheet = Some(name.to_string());
            continue;
        }

        // Skip comments, blank lines, and section headers
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        let Some(sheet_name) = current_sheet.as_deref() else {
            continue;
        };

        // Sheet property lines
        if let Some(rest) = line.strip_prefix("tab-color: ") {
            let color_str = rest.trim().trim_start_matches('#');
            if let Some(ws) = wb.sheet_mut(sheet_name) {
                ws.set_tab_color(color_str);
            }
            continue;
        }

        if let Some(rest) = line.strip_prefix("freeze: ") {
            let cell_ref = rest.trim();
            let (row, col) = cell_sort_key(cell_ref);
            if row > 0 || col > 0 {
                let x_split = if col > 0 { col - 1 } else { 0 };
                let y_split = if row > 0 { row - 1 } else { 0 };
                if x_split > 0 || y_split > 0 {
                    if let Some(ws) = wb.sheet_mut(sheet_name) {
                        let _ = ws.set_freeze_panes(x_split, y_split);
                    }
                }
            }
            continue;
        }

        if let Some(rest) = line.strip_prefix("gridlines: ") {
            let visible = rest.trim() == "visible";
            if let Some(ws) = wb.sheet_mut(sheet_name) {
                let mut opts = ws.sheet_view_options().cloned().unwrap_or_default();
                opts.set_show_gridlines(visible);
                ws.set_sheet_view_options(opts);
            }
            continue;
        }

        if let Some(rest) = line.strip_prefix("zoom: ") {
            if let Ok(zoom) = rest.trim().parse::<u32>() {
                if let Some(ws) = wb.sheet_mut(sheet_name) {
                    let mut opts = ws.sheet_view_options().cloned().unwrap_or_default();
                    opts.set_zoom_scale(zoom);
                    ws.set_sheet_view_options(opts);
                }
            }
            continue;
        }

        // Column lines: col A: width=14.0, hidden
        if let Some(rest) = line.strip_prefix("col ") {
            if let Some((col_letter, props_str)) = rest.split_once(": ") {
                let col_idx = parse_column_letter(col_letter.trim());
                if col_idx > 0 {
                    if let Some(ws) = wb.sheet_mut(sheet_name) {
                        if let Ok(col) = ws.column_mut(col_idx) {
                            for prop in split_properties(props_str) {
                                if prop == "hidden" {
                                    col.set_hidden(true);
                                } else if let Some(w) = prop.strip_prefix("width=") {
                                    if let Ok(width) = w.parse::<f64>() {
                                        col.set_width(width);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            continue;
        }

        // Row lines: row 1: height=24.0, hidden
        if let Some(rest) = line.strip_prefix("row ") {
            if let Some((row_str, props_str)) = rest.split_once(": ") {
                if let Ok(row_idx) = row_str.trim().parse::<u32>() {
                    if let Some(ws) = wb.sheet_mut(sheet_name) {
                        if let Ok(row) = ws.row_mut(row_idx) {
                            for prop in split_properties(props_str) {
                                if prop == "hidden" {
                                    row.set_hidden(true);
                                } else if let Some(h) = prop.strip_prefix("height=") {
                                    if let Ok(height) = h.parse::<f64>() {
                                        row.set_height(height);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            continue;
        }

        // Cell style lines: A1: bold, font="Calibri", size=11
        if let Some((cell_ref, props_str)) = parse_cell_line(line) {
            // Auto-create the sheet if it doesn't exist yet.
            if !wb.contains_sheet(sheet_name) {
                wb.add_sheet(sheet_name);
            }

            // Get current style, clone-modify-set
            let current_style_id = {
                let ws = wb.sheet(sheet_name);
                ws.and_then(|ws| ws.cell(cell_ref))
                    .and_then(|c| c.style_id())
                    .unwrap_or(0)
            };

            let mut style = wb.style(current_style_id).cloned().unwrap_or_default();
            let props = split_properties(props_str);
            apply_properties_to_style(&mut style, &props)?;

            let new_id = wb.add_style(style)?;

            // Ensure the cell exists and set its style
            let ws = wb
                .sheet_mut(sheet_name)
                .ok_or_else(|| IrError::InvalidBody(format!("sheet not found: {sheet_name}")))?;

            let cell_existed = ws.cell(cell_ref).is_some();

            let cell = ws.cell_mut(cell_ref).map_err(|e| {
                IrError::InvalidBody(format!("invalid cell reference {cell_ref}: {e}"))
            })?;
            cell.set_style_id(new_id);

            if cell_existed {
                result.cells_updated += 1;
            } else {
                result.cells_created += 1;
            }
        }
    }

    Ok(result)
}

/// Parse a sheet header line like `=== Sheet: Revenue ===`.
fn parse_sheet_header(line: &str) -> Option<&str> {
    let line = line.trim();
    let rest = line.strip_prefix("=== Sheet: ")?;
    let name = rest.strip_suffix(" ===")?;
    if name.is_empty() {
        return None;
    }
    Some(name)
}

/// Parse a cell line like `A1: bold, size=11`.
fn parse_cell_line(line: &str) -> Option<(&str, &str)> {
    let (cell_ref, value) = line.split_once(": ")?;
    let cell_ref = cell_ref.trim();
    if cell_ref.is_empty() {
        return None;
    }
    let first = cell_ref.as_bytes().first()?;
    if !first.is_ascii_alphabetic() {
        return None;
    }
    // Ensure it has digits (to distinguish from "col A" or other prefixes)
    if !cell_ref.bytes().any(|b| b.is_ascii_digit()) {
        return None;
    }
    Some((cell_ref, value))
}

/// Parse a column letter string to a 1-based index: A→1, Z→26, AA→27.
fn parse_column_letter(s: &str) -> u32 {
    s.bytes().fold(0u32, |acc, b| {
        acc * 26 + u32::from(b.to_ascii_uppercase() - b'A') + 1
    })
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    #![allow(clippy::panic_in_result_fn)]

    use super::*;
    use offidized_xlsx::style::Font;
    use offidized_xlsx::Workbook;

    type TestResult = std::result::Result<(), Box<dyn std::error::Error>>;

    // --- Color roundtrip tests ---

    #[test]
    fn color_rgb_roundtrip() -> TestResult {
        let cr = ColorReference::from_rgb("FF4472C4");
        let s = serialize_color(&cr);
        assert_eq!(s, "#4472C4");
        let parsed = parse_color(&s)?;
        assert_eq!(parsed.rgb(), Some("FF4472C4"));
        Ok(())
    }

    #[test]
    fn color_theme_roundtrip() -> TestResult {
        let cr = ColorReference::from_theme(ThemeColor::Accent1);
        let s = serialize_color(&cr);
        assert_eq!(s, "theme:accent1");
        let parsed = parse_color(&s)?;
        assert_eq!(parsed.theme(), Some(ThemeColor::Accent1));
        assert_eq!(parsed.tint(), None);
        Ok(())
    }

    #[test]
    fn color_theme_with_positive_tint_roundtrip() -> TestResult {
        let cr = ColorReference::from_theme_with_tint(ThemeColor::Accent1, 0.40);
        let s = serialize_color(&cr);
        assert_eq!(s, "theme:accent1+0.40");
        let parsed = parse_color(&s)?;
        assert_eq!(parsed.theme(), Some(ThemeColor::Accent1));
        assert!((parsed.tint().unwrap_or(0.0) - 0.40).abs() < 0.01);
        Ok(())
    }

    #[test]
    fn color_theme_with_negative_tint_roundtrip() -> TestResult {
        let cr = ColorReference::from_theme_with_tint(ThemeColor::Dark1, -0.25);
        let s = serialize_color(&cr);
        assert_eq!(s, "theme:dark1-0.25");
        let parsed = parse_color(&s)?;
        assert_eq!(parsed.theme(), Some(ThemeColor::Dark1));
        assert!((parsed.tint().unwrap_or(0.0) - (-0.25)).abs() < 0.01);
        Ok(())
    }

    #[test]
    fn color_indexed_roundtrip() -> TestResult {
        let cr = ColorReference::from_indexed(64);
        let s = serialize_color(&cr);
        assert_eq!(s, "indexed:64");
        let parsed = parse_color(&s)?;
        assert_eq!(parsed.indexed(), Some(64));
        Ok(())
    }

    #[test]
    fn color_auto_roundtrip() -> TestResult {
        let mut cr = ColorReference::empty();
        cr.set_auto(true);
        let s = serialize_color(&cr);
        assert_eq!(s, "auto");
        let parsed = parse_color(&s)?;
        assert_eq!(parsed.auto(), Some(true));
        Ok(())
    }

    // --- Property splitting ---

    #[test]
    fn split_simple_properties() {
        let result = split_properties("bold, italic, underline");
        assert_eq!(result, vec!["bold", "italic", "underline"]);
    }

    #[test]
    fn split_quoted_comma() {
        let result = split_properties("format=\"#,##0.00\", bold");
        assert_eq!(result, vec!["format=\"#,##0.00\"", "bold"]);
    }

    #[test]
    fn split_single_property() {
        let result = split_properties("bold");
        assert_eq!(result, vec!["bold"]);
    }

    // --- Column index conversions ---

    #[test]
    fn column_letter_conversion() {
        assert_eq!(column_index_to_letter(1), "A");
        assert_eq!(column_index_to_letter(26), "Z");
        assert_eq!(column_index_to_letter(27), "AA");
        assert_eq!(column_index_to_letter(52), "AZ");
    }

    #[test]
    fn column_letter_parsing() {
        assert_eq!(parse_column_letter("A"), 1);
        assert_eq!(parse_column_letter("Z"), 26);
        assert_eq!(parse_column_letter("AA"), 27);
    }

    // --- Property parsing ---

    #[test]
    fn apply_bold_property() -> TestResult {
        let mut style = Style::new();
        apply_properties_to_style(&mut style, &["bold"])?;
        assert_eq!(style.font().and_then(|f| f.bold()), Some(true));
        Ok(())
    }

    #[test]
    fn apply_font_and_size() -> TestResult {
        let mut style = Style::new();
        apply_properties_to_style(&mut style, &["font=\"Calibri\"", "size=11"])?;
        assert_eq!(style.font().and_then(|f| f.name()), Some("Calibri"));
        assert_eq!(style.font().and_then(|f| f.size()), Some("11"));
        Ok(())
    }

    #[test]
    fn apply_alignment() -> TestResult {
        let mut style = Style::new();
        apply_properties_to_style(&mut style, &["align=center", "valign=middle"])?;
        assert_eq!(
            style.alignment().and_then(|a| a.horizontal()),
            Some(HorizontalAlignment::Center),
        );
        // "middle" is not valid, should be "center" for vertical — test graceful skip
        Ok(())
    }

    #[test]
    fn apply_format_with_commas() -> TestResult {
        let mut style = Style::new();
        // When parsed by split_properties, this arrives as: format="#,##0.00"
        apply_properties_to_style(&mut style, &["format=\"#,##0.00\""])?;
        assert_eq!(style.custom_format(), Some("#,##0.00"));
        Ok(())
    }

    #[test]
    fn apply_border() -> TestResult {
        let mut style = Style::new();
        apply_properties_to_style(&mut style, &["border-bottom=thin #000000"])?;
        let border = style.border().ok_or("no border")?;
        let bottom = border.bottom().ok_or("no bottom")?;
        assert_eq!(bottom.style(), Some("thin"));
        assert!(bottom.color_ref().is_some());
        Ok(())
    }

    // --- Derive tests ---

    #[test]
    fn derive_empty_workbook() {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");
        let mut output = String::new();
        derive_style(&wb, &DeriveOptions::default(), &mut output);
        assert!(output.contains("=== Sheet: Sheet1 ==="));
        // No cell styles for default workbook
        assert!(!output.contains("# Cell styles"));
    }

    #[test]
    fn derive_styled_cell() -> TestResult {
        let mut wb = Workbook::new();
        let mut style = Style::new();
        let mut font = Font::new();
        font.set_bold(true);
        font.set_name("Calibri");
        font.set_size("11");
        style.set_font(font);
        let id = wb.add_style(style)?;

        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value("test").set_style_id(id);

        let mut output = String::new();
        derive_style(&wb, &DeriveOptions::default(), &mut output);
        assert!(output.contains("A1: bold, font=\"Calibri\", size=11"));
        Ok(())
    }

    // --- Apply tests ---

    #[test]
    fn apply_basic_style() -> TestResult {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value("test");

        let body = "\n=== Sheet: Sheet1 ===\n\n# Cell styles\nA1: bold, size=12\n";
        let result = apply_style(body, &mut wb)?;

        assert_eq!(result.cells_updated, 1);

        let ws = wb.sheet("Sheet1").ok_or("missing")?;
        let cell = ws.cell("A1").ok_or("missing")?;
        let style_id = cell.style_id().ok_or("no style")?;
        let style = wb.style(style_id).ok_or("no style")?;
        assert_eq!(style.font().and_then(|f| f.bold()), Some(true));
        assert_eq!(style.font().and_then(|f| f.size()), Some("12"));

        Ok(())
    }

    #[test]
    fn apply_preserves_existing_content() -> TestResult {
        let mut wb = Workbook::new();
        let ws = wb.add_sheet("Sheet1");
        ws.cell_mut("A1")?.set_value("keep me");

        let body = "\n=== Sheet: Sheet1 ===\n\n# Cell styles\nA1: bold\n";
        apply_style(body, &mut wb)?;

        let ws = wb.sheet("Sheet1").ok_or("missing")?;
        let cell = ws.cell("A1").ok_or("missing")?;
        assert_eq!(
            cell.value(),
            Some(&offidized_xlsx::CellValue::String("keep me".into())),
        );

        Ok(())
    }

    #[test]
    fn apply_column_width() -> TestResult {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let body =
            "\n=== Sheet: Sheet1 ===\n\n# Columns\ncol A: width=14.0\ncol B: width=16.0, hidden\n";
        apply_style(body, &mut wb)?;

        let ws = wb.sheet("Sheet1").ok_or("missing")?;
        let col_a = ws.column(1).ok_or("col A missing")?;
        assert!((col_a.width().unwrap_or(0.0) - 14.0).abs() < 0.01);

        let col_b = ws.column(2).ok_or("col B missing")?;
        assert!((col_b.width().unwrap_or(0.0) - 16.0).abs() < 0.01);
        assert!(col_b.is_hidden());

        Ok(())
    }

    #[test]
    fn apply_row_height() -> TestResult {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let body = "\n=== Sheet: Sheet1 ===\n\n# Rows\nrow 1: height=24.0\nrow 5: hidden\n";
        apply_style(body, &mut wb)?;

        let ws = wb.sheet("Sheet1").ok_or("missing")?;
        let row1 = ws.row(1).ok_or("row 1 missing")?;
        assert!((row1.height().unwrap_or(0.0) - 24.0).abs() < 0.01);

        let row5 = ws.row(5).ok_or("row 5 missing")?;
        assert!(row5.is_hidden());

        Ok(())
    }

    #[test]
    fn apply_sheet_properties() -> TestResult {
        let mut wb = Workbook::new();
        wb.add_sheet("Sheet1");

        let body = "\n=== Sheet: Sheet1 ===\n\n# Sheet properties\ntab-color: #4472C4\nzoom: 150\ngridlines: hidden\n";
        apply_style(body, &mut wb)?;

        let ws = wb.sheet("Sheet1").ok_or("missing")?;
        assert_eq!(ws.tab_color(), Some("4472C4"));

        let view_opts = ws.sheet_view_options().ok_or("no view opts")?;
        assert_eq!(view_opts.zoom_scale(), Some(150));
        assert_eq!(view_opts.show_gridlines(), Some(false));

        Ok(())
    }
}
