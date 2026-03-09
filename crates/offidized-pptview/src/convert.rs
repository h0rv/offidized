//! Converts an [`offidized_pptx::Presentation`] into a
//! [`PresentationViewModel`](crate::model::PresentationViewModel).

use base64::Engine;
use offidized_pptx::shape::{
    BulletStyle, GradientFillType, LineDashStyle, PlaceholderType, Shape, ShapeFill, ShapeGeometry,
    TextAlignment, TextAnchor,
};
use offidized_pptx::slide::SlideBackground;
use offidized_pptx::slide_layout::SlideLayout;
use offidized_pptx::table::CellTextAnchor;
use offidized_pptx::Presentation;

use crate::model::{
    BackgroundModel, BulletModel, ImageModel, InsetsModel, OutlineModel, PresentationViewModel,
    ShapeFillModel, ShapeModel, SlideModel, TableCellModel, TableModel, TableRowModel,
    TextBodyModel, TextParagraphModel, TextRunModel,
};
use crate::units::{angle_to_degrees, emu_to_pt, hundredths_to_pt};

/// Errors that can occur during presentation conversion.
#[derive(Debug, thiserror::Error)]
pub enum ConvertError {
    /// Base64 encoding failed.
    #[error("base64 encoding failed: {0}")]
    Base64(String),
}

/// Result type for conversion operations.
pub type Result<T> = std::result::Result<T, ConvertError>;

/// Default slide width (10 inches = 9144000 EMU = 720pt).
const DEFAULT_SLIDE_WIDTH_PT: f64 = 720.0;
/// Default slide height (7.5 inches = 6858000 EMU = 540pt).
const DEFAULT_SLIDE_HEIGHT_PT: f64 = 540.0;

/// Convert a parsed `Presentation` into a renderer-friendly `PresentationViewModel`.
pub fn convert_presentation(pres: &Presentation) -> Result<PresentationViewModel> {
    let slide_width_pt = pres
        .slide_width_emu()
        .map(emu_to_pt)
        .unwrap_or(DEFAULT_SLIDE_WIDTH_PT);
    let slide_height_pt = pres
        .slide_height_emu()
        .map(emu_to_pt)
        .unwrap_or(DEFAULT_SLIDE_HEIGHT_PT);

    // Collect all images across all slides into a flat vector.
    let mut images = Vec::new();
    let engine = base64::engine::general_purpose::STANDARD;

    let mut slides = Vec::with_capacity(pres.slides().len());

    for slide in pres.slides() {
        // Track per-slide image offset into the global images vec.
        let image_offset = images.len();

        // Resolve the slide layout for placeholder geometry inheritance.
        let layout: Option<&SlideLayout> = slide
            .layout_reference()
            .and_then(|(mi, li)| pres.layout(mi, li));

        // Encode slide images as base64 data URIs.
        for img in slide.images() {
            let encoded = engine.encode(img.bytes());
            let data_uri = format!("data:{};base64,{encoded}", img.content_type());
            images.push(ImageModel {
                data_uri,
                content_type: img.content_type().to_string(),
            });
        }

        // Convert shapes — resolve placeholder geometry from layout when needed.
        let mut shapes = Vec::new();
        for shape in slide.shapes() {
            shapes.push(convert_shape_with_layout(shape, layout));
        }

        // Convert tables into shapes.
        for table in slide.tables() {
            shapes.push(convert_table_shape(table));
        }

        // Convert images into shapes.
        for (img_index, img) in slide.images().iter().enumerate() {
            let name = img.name().map(|n| n.to_string());
            let image_global_index = image_offset + img_index;

            let already_has_image = shapes
                .iter()
                .any(|s| s.image_index == Some(image_global_index));
            if !already_has_image {
                shapes.push(ShapeModel {
                    x_pt: 0.0,
                    y_pt: 0.0,
                    width_pt: 100.0,
                    height_pt: 100.0,
                    rotation: None,
                    name,
                    preset_geometry: None,
                    fill: None,
                    outline: None,
                    text: None,
                    image_index: Some(image_global_index),
                    table: None,
                    hidden: false,
                });
            }
        }

        // Convert grouped shapes (flatten into top-level shapes).
        for group in slide.grouped_shapes() {
            for child_shape in group.shapes() {
                shapes.push(convert_shape_with_layout(child_shape, layout));
            }
        }

        // Background: slide → layout → default white.
        let background = convert_background(slide.background())
            .or_else(|| layout.and_then(|l| convert_background(l.background())));

        slides.push(SlideModel {
            shapes,
            background,
            notes: slide.notes_text().map(|s| s.to_string()),
            hidden: slide.is_hidden(),
        });
    }

    Ok(PresentationViewModel {
        slides,
        slide_width_pt,
        slide_height_pt,
        images,
    })
}

// ---------------------------------------------------------------------------
// Shapes
// ---------------------------------------------------------------------------

/// Convert a shape, resolving geometry from the layout if the shape itself
/// has no `<a:xfrm>` (common for placeholder shapes).
fn convert_shape_with_layout(shape: &Shape, layout: Option<&SlideLayout>) -> ShapeModel {
    let geom = shape
        .geometry()
        .or_else(|| resolve_placeholder_geometry(shape, layout));

    let (x_pt, y_pt, width_pt, height_pt) = match geom {
        Some(g) => (
            emu_to_pt(g.x()),
            emu_to_pt(g.y()),
            emu_to_pt(g.cx()),
            emu_to_pt(g.cy()),
        ),
        None => {
            // Last resort: use default positions based on placeholder type.
            default_placeholder_bounds(shape.placeholder_type())
        }
    };

    let fill = convert_fill(shape.fill());
    let outline = convert_outline(shape.outline());
    let text = convert_text_body(shape);
    let rotation = shape.rotation().map(angle_to_degrees);

    let name_str = shape.name();
    let name = if name_str.is_empty() {
        None
    } else {
        Some(name_str.to_string())
    };

    ShapeModel {
        x_pt,
        y_pt,
        width_pt,
        height_pt,
        rotation,
        name,
        preset_geometry: shape.preset_geometry().map(|s| s.to_string()),
        fill,
        outline,
        text,
        image_index: None,
        table: None,
        hidden: shape.is_hidden(),
    }
}

/// Look up matching placeholder geometry from the slide layout.
fn resolve_placeholder_geometry(
    shape: &Shape,
    layout: Option<&SlideLayout>,
) -> Option<ShapeGeometry> {
    let layout = layout?;
    let ph_type = shape.placeholder_type()?;

    // Find a shape in the layout with the same placeholder type.
    for layout_shape in layout.shapes() {
        if layout_shape.placeholder_type() == Some(ph_type) {
            if let Some(geom) = layout_shape.geometry() {
                return Some(geom);
            }
        }
    }

    // Fallback: match by name pattern (e.g., "Title 1" → title placeholder).
    let shape_name = shape.name().to_lowercase();
    for layout_shape in layout.shapes() {
        let layout_name = layout_shape.name().to_lowercase();
        if !layout_name.is_empty() && layout_name == shape_name {
            if let Some(geom) = layout_shape.geometry() {
                return Some(geom);
            }
        }
    }

    None
}

/// Provide reasonable default bounds (in pt) when no geometry is available.
/// Returns (x_pt, y_pt, width_pt, height_pt).
fn default_placeholder_bounds(ph_type: Option<&PlaceholderType>) -> (f64, f64, f64, f64) {
    match ph_type {
        Some(PlaceholderType::Title | PlaceholderType::CenteredTitle) => {
            // Typical title: across the top
            (36.0, 21.6, 648.0, 90.0)
        }
        Some(PlaceholderType::Subtitle) => {
            // Below title
            (36.0, 126.0, 648.0, 90.0)
        }
        Some(PlaceholderType::Body | PlaceholderType::Object) => {
            // Content area below title
            (36.0, 126.0, 648.0, 378.0)
        }
        Some(PlaceholderType::SlideNumber) => {
            // Bottom right corner
            (612.0, 504.0, 72.0, 27.0)
        }
        Some(PlaceholderType::DateAndTime) => {
            // Bottom left
            (36.0, 504.0, 144.0, 27.0)
        }
        Some(PlaceholderType::Footer) => {
            // Bottom center
            (252.0, 504.0, 216.0, 27.0)
        }
        _ => {
            // Generic shape: centered with reasonable size
            (72.0, 72.0, 576.0, 396.0)
        }
    }
}

fn convert_table_shape(table: &offidized_pptx::table::Table) -> ShapeModel {
    let column_widths_pt: Vec<f64> = table
        .column_widths_emu()
        .iter()
        .map(|&w| emu_to_pt(w))
        .collect();
    let row_heights_pt: Vec<f64> = table
        .row_heights_emu()
        .iter()
        .map(|&h| emu_to_pt(h))
        .collect();

    let total_width: f64 = column_widths_pt.iter().sum();
    let total_height: f64 = row_heights_pt.iter().sum();

    let mut rows = Vec::with_capacity(table.rows());
    for row in 0..table.rows() {
        let mut cells = Vec::with_capacity(table.cols());
        for col in 0..table.cols() {
            if let Some(cell) = table.cell(row, col) {
                cells.push(TableCellModel {
                    text: cell.text().to_string(),
                    fill_color: cell.fill_color_srgb().map(|s| s.to_string()),
                    grid_span: cell.grid_span().unwrap_or(1),
                    row_span: cell.row_span().unwrap_or(1),
                    v_merge: cell.is_v_merge(),
                    vertical_align: cell.vertical_alignment().map(|va| {
                        match va {
                            CellTextAnchor::Top => "top",
                            CellTextAnchor::Middle => "middle",
                            CellTextAnchor::Bottom => "bottom",
                        }
                        .to_string()
                    }),
                });
            }
        }
        rows.push(TableRowModel { cells });
    }

    let table_model = TableModel {
        rows,
        column_widths_pt,
        row_heights_pt,
    };

    ShapeModel {
        x_pt: 0.0,
        y_pt: 0.0,
        width_pt: total_width,
        height_pt: total_height,
        rotation: None,
        name: None,
        preset_geometry: None,
        fill: None,
        outline: None,
        text: None,
        image_index: None,
        table: Some(table_model),
        hidden: false,
    }
}

// ---------------------------------------------------------------------------
// Fill
// ---------------------------------------------------------------------------

fn convert_fill(fill: Option<ShapeFill>) -> Option<ShapeFillModel> {
    match fill? {
        ShapeFill::Solid(color) => Some(ShapeFillModel::Solid { color }),
        ShapeFill::Gradient(gradient) => {
            let css = gradient_to_css(&gradient);
            Some(ShapeFillModel::Gradient { css })
        }
        ShapeFill::NoFill => Some(ShapeFillModel::None),
        ShapeFill::Pattern(_) | ShapeFill::Picture(_) => None,
    }
}

fn gradient_to_css(gradient: &offidized_pptx::shape::GradientFill) -> String {
    let angle = gradient.linear_angle.map(angle_to_degrees).unwrap_or(0.0);

    let stops: Vec<String> = gradient
        .stops
        .iter()
        .map(|stop| {
            let pct = stop.position as f64 / 1000.0;
            format!("#{} {pct:.1}%", stop.color_srgb)
        })
        .collect();

    let direction = match gradient.fill_type {
        Some(GradientFillType::Radial | GradientFillType::Rectangular | GradientFillType::Path) => {
            return format!("radial-gradient(circle, {})", stops.join(", "));
        }
        _ => format!("{angle:.0}deg"),
    };

    format!("linear-gradient({direction}, {})", stops.join(", "))
}

// ---------------------------------------------------------------------------
// Outline
// ---------------------------------------------------------------------------

fn convert_outline(outline: Option<&offidized_pptx::shape::ShapeOutline>) -> Option<OutlineModel> {
    let outline = outline?;
    if !outline.is_set() {
        return None;
    }

    let width_pt = outline.width_emu.map(emu_to_pt);
    let color = outline.color_srgb.clone();
    let dash_style = outline.dash_style.map(|ds| {
        match ds {
            LineDashStyle::Solid => "solid",
            LineDashStyle::Dot | LineDashStyle::SystemDot => "dotted",
            LineDashStyle::Dash
            | LineDashStyle::LargeDash
            | LineDashStyle::SystemDash
            | LineDashStyle::DashDot
            | LineDashStyle::LargeDashDot
            | LineDashStyle::LargeDashDotDot
            | LineDashStyle::SystemDashDot
            | LineDashStyle::SystemDashDotDot => "dashed",
        }
        .to_string()
    });

    Some(OutlineModel {
        width_pt,
        color,
        dash_style,
    })
}

// ---------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------

fn convert_text_body(shape: &Shape) -> Option<TextBodyModel> {
    let paragraphs = shape.paragraphs();
    if paragraphs.is_empty() {
        return None;
    }

    // Check if all paragraphs are empty (no text at all).
    let has_text = paragraphs
        .iter()
        .any(|p| p.runs().iter().any(|r| !r.text().is_empty()));
    if !has_text {
        return None;
    }

    let paragraph_models: Vec<TextParagraphModel> =
        paragraphs.iter().map(convert_paragraph).collect();

    let anchor = shape.text_anchor().map(|a| {
        match a {
            TextAnchor::Top | TextAnchor::TopCentered => "top",
            TextAnchor::Middle | TextAnchor::MiddleCentered => "middle",
            TextAnchor::Bottom | TextAnchor::BottomCentered => "bottom",
        }
        .to_string()
    });

    // Text insets (default in pptx: 91440 EMU left/right = ~7.2pt, 45720 EMU top/bottom = ~3.6pt).
    let has_insets = shape.text_inset_left().is_some()
        || shape.text_inset_top().is_some()
        || shape.text_inset_right().is_some()
        || shape.text_inset_bottom().is_some();

    let insets = if has_insets {
        Some(InsetsModel {
            left_pt: shape.text_inset_left().map(emu_to_pt).unwrap_or(7.2),
            top_pt: shape.text_inset_top().map(emu_to_pt).unwrap_or(3.6),
            right_pt: shape.text_inset_right().map(emu_to_pt).unwrap_or(7.2),
            bottom_pt: shape.text_inset_bottom().map(emu_to_pt).unwrap_or(3.6),
        })
    } else {
        None
    };

    Some(TextBodyModel {
        paragraphs: paragraph_models,
        anchor,
        insets,
    })
}

fn convert_paragraph(para: &offidized_pptx::shape::ShapeParagraph) -> TextParagraphModel {
    let runs: Vec<TextRunModel> = para.runs().iter().map(convert_run).collect();

    let props = para.properties();

    let alignment = props.alignment.map(|a| {
        match a {
            TextAlignment::Left => "left",
            TextAlignment::Center => "center",
            TextAlignment::Right => "right",
            TextAlignment::Justified | TextAlignment::Distributed => "justify",
        }
        .to_string()
    });

    let level = props.level;

    let spacing_before_pt = props
        .space_before
        .as_ref()
        .map(spacing_value_to_pt)
        .or_else(|| props.space_before_pts.map(|v| v as f64 / 100.0));

    let spacing_after_pt = props
        .space_after
        .as_ref()
        .map(spacing_value_to_pt)
        .or_else(|| props.space_after_pts.map(|v| v as f64 / 100.0));

    let line_spacing = props
        .line_spacing
        .as_ref()
        .map(line_spacing_to_multiplier)
        .or_else(|| {
            if let Some(pct) = props.line_spacing_pct {
                Some(pct as f64 / 100_000.0)
            } else {
                props.line_spacing_pts.map(|pts| pts as f64 / 100.0)
            }
        });

    let bullet = convert_bullet(&props.bullet);

    TextParagraphModel {
        runs,
        alignment,
        level,
        spacing_before_pt,
        spacing_after_pt,
        line_spacing,
        bullet,
    }
}

fn spacing_value_to_pt(sv: &offidized_pptx::shape::SpacingValue) -> f64 {
    use offidized_pptx::shape::SpacingUnit;
    match sv.unit {
        SpacingUnit::Points => sv.value as f64 / 100.0,
        SpacingUnit::Percent => {
            // Percentage of font size — approximate as fraction of default 18pt.
            sv.value as f64 / 100_000.0 * 18.0
        }
    }
}

fn line_spacing_to_multiplier(ls: &offidized_pptx::shape::LineSpacing) -> f64 {
    use offidized_pptx::shape::LineSpacingUnit;
    match ls.unit {
        LineSpacingUnit::Percent => ls.value as f64 / 100_000.0,
        LineSpacingUnit::Points => ls.value as f64 / 100.0,
    }
}

fn convert_run(run: &offidized_pptx::text::TextRun) -> TextRunModel {
    TextRunModel {
        text: run.text().to_string(),
        bold: run.is_bold(),
        italic: run.is_italic(),
        underline: run.underline().is_some(),
        strikethrough: run.strikethrough().is_some(),
        font_family: run.font_name().map(|s| s.to_string()),
        font_size_pt: run.font_size().map(hundredths_to_pt),
        color: run.font_color().map(|s| s.to_string()),
        hyperlink: run.hyperlink_url().map(|s| s.to_string()),
    }
}

// ---------------------------------------------------------------------------
// Bullets
// ---------------------------------------------------------------------------

fn convert_bullet(bullet: &offidized_pptx::shape::BulletProperties) -> Option<BulletModel> {
    let style = bullet.style.as_ref()?;
    match style {
        BulletStyle::None => None,
        BulletStyle::Char(ch) => {
            let mapped = map_bullet_char(ch, bullet.font_name.as_deref().unwrap_or(""));
            Some(BulletModel {
                char: Some(mapped),
                auto_num_type: None,
                font_family: bullet.font_name.clone(),
                color: bullet.color_srgb.clone(),
            })
        }
        BulletStyle::AutoNum(num_type) => Some(BulletModel {
            char: None,
            auto_num_type: Some(num_type.clone()),
            font_family: bullet.font_name.clone(),
            color: bullet.color_srgb.clone(),
        }),
    }
}

/// Map PUA bullet characters from Symbol/Wingdings fonts to standard Unicode.
fn map_bullet_char(text: &str, font: &str) -> String {
    let mut chars = text.chars();
    if let Some(ch) = chars.next() {
        if chars.next().is_none() && ('\u{F000}'..='\u{F0FF}').contains(&ch) {
            let code = ch as u32 & 0xFF;
            let font_lower = font.to_lowercase();
            let mapped = if font_lower.contains("wingdings") {
                match code {
                    0x6C => '\u{25CF}', // filled circle
                    0x6E => '\u{25A0}', // filled square
                    0x71 => '\u{25C6}', // filled diamond
                    0x76 => '\u{2714}', // check mark
                    0x77 => '\u{2718}', // cross mark
                    0xA7 => '\u{25A0}', // filled square
                    0xA8 => '\u{25CB}', // white circle
                    0xD8 => '\u{27A2}', // right arrow
                    0xFC => '\u{2714}', // check mark
                    0xFB => '\u{25CF}', // filled circle
                    _ => '\u{2022}',    // fallback: standard bullet
                }
            } else if font_lower.contains("symbol") {
                match code {
                    0xB7 => '\u{2022}', // bullet
                    0x6F => '\u{25CB}', // white circle
                    0xA7 => '\u{2666}', // diamond
                    _ => '\u{2022}',    // fallback
                }
            } else {
                '\u{2022}'
            };
            return mapped.to_string();
        }
    }
    text.to_string()
}

// ---------------------------------------------------------------------------
// Background
// ---------------------------------------------------------------------------

fn convert_background(bg: Option<&SlideBackground>) -> Option<BackgroundModel> {
    match bg? {
        SlideBackground::Solid(color) => Some(BackgroundModel::Solid {
            color: color.clone(),
        }),
        SlideBackground::Gradient(gradient) => {
            let css = gradient_to_css(gradient);
            Some(BackgroundModel::Gradient { css })
        }
        SlideBackground::Pattern { .. } | SlideBackground::Image { .. } => None,
    }
}
