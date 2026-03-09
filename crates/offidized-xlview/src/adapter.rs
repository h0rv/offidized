//! Adapter layer: converts `offidized_xlsx` types into the viewer's internal data model.
//!
//! This module bridges the high-level `offidized_xlsx::Workbook` API with the
//! xlview-specific types used by the renderer and layout engine.

use std::collections::HashMap;
use std::sync::Arc;

use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use base64::Engine;

use crate::cell_ref::parse_cell_ref;
use crate::render::{BorderStyleData, CellStyleData};
use crate::types::chart::{
    BarDirection as ViewerBarDirection, Chart as ViewerChart, ChartAxis as ViewerChartAxis,
    ChartDataRef as ViewerChartDataRef, ChartGrouping as ViewerChartGrouping,
    ChartLegend as ViewerChartLegend, ChartSeries as ViewerChartSeries,
    ChartType as ViewerChartType,
};
use crate::types::content::ValidationType;
use crate::types::content::{DataValidation as ViewerDataValidation, DataValidationRange};
use crate::types::drawing::{Drawing, EmbeddedImage, ImageFormat};
use crate::types::filter::{
    AutoFilter as ViewerAutoFilter, CustomFilter as ViewerCustomFilter,
    CustomFilterOperator as ViewerCustomFilterOperator, FilterColumn as ViewerFilterColumn,
    FilterType as ViewerFilterType,
};
use crate::types::formatting::{
    CFRule, CFRuleType, CFValueObject, ColorScale, ConditionalFormatting as ViewerCF, DataBar,
    DxfStyle, IconSet,
};
use crate::types::rich_text::{RichTextRun, RunStyle};
use crate::types::sparkline::{
    Sparkline as ViewerSparkline, SparklineColors as ViewerSparklineColors,
    SparklineGroup as ViewerSparklineGroup,
};
use crate::types::style::{
    Border, BorderStyle, ColWidth, GradientFill, GradientStop, HAlign, MergeRange, PatternType,
    RowHeight, Style, StyleRef, Theme, VAlign, VertAlign,
};
use crate::types::workbook::{Cell, CellData, CellRawValue, CellType, Comment, Sheet, Workbook};
use crate::types::Hyperlink;

/// Convert an `offidized_xlsx::Workbook` into the viewer's internal `Workbook` model.
pub fn convert_workbook(wb: &offidized_xlsx::Workbook) -> Workbook {
    let theme_colors = wb.theme_colors();
    let indexed_colors = wb.indexed_colors();

    // Build theme
    let theme = Theme {
        colors: theme_colors.to_vec(),
        major_font: wb.major_font().map(String::from),
        minor_font: wb.minor_font().map(String::from),
    };

    // Build resolved styles and numfmt cache
    let styles = wb.styles().styles();
    let mut resolved_styles: Vec<Option<StyleRef>> = Vec::with_capacity(styles.len());
    let mut numfmt_cache: Vec<offidized_xlsx::CompiledFormat> = Vec::with_capacity(styles.len());

    for style in styles {
        resolved_styles.push(Some(convert_style(style, theme_colors, indexed_colors)));
        let compiled = match style.number_format() {
            Some(code) => offidized_xlsx::compile_format(code),
            None => offidized_xlsx::CompiledFormat::General,
        };
        numfmt_cache.push(compiled);
    }

    // Build default style from style index 0
    let default_style = wb
        .style(0)
        .map(|s| convert_style(s, theme_colors, indexed_colors));

    // Build sheets
    let sheets: Vec<Sheet> = wb
        .worksheets()
        .iter()
        .enumerate()
        .map(|(ws_idx, ws)| convert_worksheet(ws, ws_idx, wb, theme_colors, indexed_colors))
        .collect();

    // Build embedded images from all worksheets
    let images = convert_all_images(wb);

    // Build DXF styles
    let dxf_styles = convert_dxf_styles(wb, theme_colors, indexed_colors);

    Workbook {
        sheets,
        shared_strings: Vec::new(),
        numfmt_cache,
        date1904: wb.date1904(),
        images,
        theme,
        dxf_styles,
        resolved_styles,
        default_style,
        sheet_paths: Vec::new(),
    }
}

/// Convert an `offidized_xlsx::Worksheet` into the viewer's `Sheet`.
fn convert_worksheet(
    ws: &offidized_xlsx::Worksheet,
    ws_idx: usize,
    wb: &offidized_xlsx::Workbook,
    theme_colors: &[String],
    indexed_colors: Option<&[String]>,
) -> Sheet {
    // Build hyperlink lookup by cell reference
    let hyperlink_map: HashMap<String, Hyperlink> = ws
        .hyperlinks()
        .iter()
        .filter_map(|h| {
            let target = h.url().or_else(|| h.location())?;
            let is_external = h.url().is_some();
            Some((
                h.cell_ref().to_uppercase(),
                Hyperlink {
                    target: target.to_string(),
                    is_external,
                    display: h.display().map(String::from),
                    tooltip: h.tooltip().map(String::from),
                    location: h.location().map(String::from),
                },
            ))
        })
        .collect();

    // Build comment lookup by cell reference
    let comment_refs: HashMap<String, usize> = ws
        .comments()
        .iter()
        .enumerate()
        .map(|(idx, c)| (c.cell_ref().to_uppercase(), idx))
        .collect();

    // Convert comments
    let comments: Vec<Comment> = ws
        .comments()
        .iter()
        .map(|c| Comment {
            text: c.text().to_string(),
            author: Some(c.author().to_string()),
        })
        .collect();

    // Convert cells
    let mut cells = Vec::new();
    let mut max_row: u32 = 0;
    let mut max_col: u32 = 0;

    for (cell_ref, xlsx_cell) in ws.cells() {
        let Some((col, row)) = parse_cell_ref(cell_ref) else {
            continue;
        };

        // Map CellValue -> CellRawValue
        let raw = xlsx_cell.value().and_then(|val| match val {
            offidized_xlsx::CellValue::Number(n) => Some(CellRawValue::Number(*n)),
            offidized_xlsx::CellValue::String(s) => Some(CellRawValue::String(s.clone())),
            offidized_xlsx::CellValue::Bool(b) => Some(CellRawValue::Boolean(*b)),
            offidized_xlsx::CellValue::Error(e) => Some(CellRawValue::Error(e.clone())),
            offidized_xlsx::CellValue::DateTime(n) => Some(CellRawValue::Date(*n)),
            offidized_xlsx::CellValue::Date(s) => Some(CellRawValue::String(s.clone())),
            offidized_xlsx::CellValue::RichText(runs) => {
                let plain: String = runs.iter().map(|r| r.text()).collect();
                Some(CellRawValue::String(plain))
            }
            offidized_xlsx::CellValue::Blank => None,
        });

        // Convert rich text runs if present
        let rich_text = xlsx_cell.rich_text().map(|runs| {
            runs.iter()
                .map(|run| {
                    let style = if run.has_formatting() {
                        Some(RunStyle {
                            font_family: run.font_name().map(String::from),
                            font_size: run.font_size().and_then(|s| s.parse::<f64>().ok()),
                            font_color: run.color().map(argb_to_css),
                            bold: run.bold(),
                            italic: run.italic(),
                            // TODO: offidized-xlsx RichTextRun needs underline API
                            underline: None,
                            // TODO: offidized-xlsx RichTextRun needs strikethrough API
                            strikethrough: None,
                            // TODO: offidized-xlsx RichTextRun needs vert_align API
                            vert_align: None,
                        })
                    } else {
                        None
                    };
                    RichTextRun {
                        text: run.text().to_string(),
                        style,
                    }
                })
                .collect()
        });

        // Look up hyperlink
        let cell_ref_upper = cell_ref.to_uppercase();
        let hyperlink = hyperlink_map.get(&cell_ref_upper).cloned();

        // Check for comment
        let has_comment = if comment_refs.contains_key(&cell_ref_upper) {
            Some(true)
        } else {
            None
        };

        if row > max_row {
            max_row = row;
        }
        if col > max_col {
            max_col = col;
        }

        // Determine cell type from raw value
        let t = match &raw {
            Some(CellRawValue::Number(_)) => CellType::Number,
            Some(CellRawValue::Date(_)) => CellType::Date,
            Some(CellRawValue::Boolean(_)) => CellType::Boolean,
            Some(CellRawValue::Error(_)) => CellType::Error,
            _ => CellType::String,
        };

        cells.push(CellData {
            r: row,
            c: col,
            cell: Cell {
                v: None,
                cached_display: None,
                raw,
                style_idx: xlsx_cell.style_id(),
                hyperlink,
                has_comment,
                rich_text,
                s: None,
                cached_rich_text: None,
                t,
                formula: None,
            },
        });
    }

    // Column widths
    let col_widths: Vec<ColWidth> = ws
        .columns()
        .filter_map(|c| {
            c.width().map(|w| ColWidth {
                col: c.index().saturating_sub(1),
                width: w,
            })
        })
        .collect();

    // Row heights
    let row_heights: Vec<RowHeight> = ws
        .rows()
        .filter_map(|r| {
            r.height().map(|h| RowHeight {
                row: r.index().saturating_sub(1),
                height: h,
            })
        })
        .collect();

    // Hidden cols/rows
    let hidden_cols: Vec<u32> = ws
        .columns()
        .filter(|c| c.is_hidden())
        .map(|c| c.index().saturating_sub(1))
        .collect();

    let hidden_rows: Vec<u32> = ws
        .rows()
        .filter(|r| r.is_hidden())
        .map(|r| r.index().saturating_sub(1))
        .collect();

    // Merged ranges (1-based to 0-based)
    let merges: Vec<MergeRange> = ws
        .merged_ranges()
        .iter()
        .map(|mr| MergeRange {
            start_row: mr.start_row().saturating_sub(1),
            start_col: mr.start_column().saturating_sub(1),
            end_row: mr.end_row().saturating_sub(1),
            end_col: mr.end_column().saturating_sub(1),
        })
        .collect();

    // Freeze panes
    let (frozen_rows, frozen_cols) = ws
        .freeze_pane()
        .map(|fp| (fp.y_split(), fp.x_split()))
        .unwrap_or((0, 0));

    // Tab color
    let tab_color = ws.tab_color().map(|c| format!("#{c}"));

    // Comments by cell index
    let comments_by_cell: HashMap<String, usize> = comment_refs;

    // Charts
    let charts = convert_charts(ws, wb);

    // Drawings/Images
    let drawings = convert_drawings(ws, ws_idx);

    // Sparklines
    let sparkline_groups = convert_sparkline_groups(ws);

    // Conditional formatting
    let conditional_formatting = convert_conditional_formattings(ws, theme_colors, indexed_colors);

    // Auto-filter
    let auto_filter = convert_auto_filter(ws);

    // Data validations
    let data_validations = convert_data_validations(ws);

    let mut sheet = Sheet {
        name: ws.name().to_string(),
        tab_color,
        cells,
        col_widths,
        row_heights,
        hidden_cols,
        hidden_rows,
        merges,
        frozen_rows,
        frozen_cols,
        max_row,
        max_col,
        drawings,
        charts,
        data_validations,
        conditional_formatting,
        conditional_formatting_cache: Vec::new(),
        sparkline_groups,
        auto_filter,
        comments,
        comments_by_cell,
        cells_by_row: Vec::new(),
        default_col_width: 8.43,
        default_row_height: 15.0,
        hyperlinks: Vec::new(),
    };
    sheet.rebuild_cell_index();
    sheet
}

// ============================================================================
// Feature conversion functions
// ============================================================================

/// Convert charts from an `offidized_xlsx::Worksheet` into viewer chart types.
///
/// Accepts the full workbook so that chart formula references (e.g.
/// `'Charts'!$B$2:$B$5`) can be resolved to actual cell values when
/// the chart XML has no `<numCache>`.
fn convert_charts(
    ws: &offidized_xlsx::Worksheet,
    wb: &offidized_xlsx::Workbook,
) -> Vec<ViewerChart> {
    ws.charts()
        .iter()
        .map(|chart| {
            let chart_type = convert_chart_type(chart.chart_type());
            let bar_direction = chart.bar_direction().map(convert_bar_direction);
            let grouping = chart.grouping().map(convert_chart_grouping);

            let series: Vec<ViewerChartSeries> = chart
                .series()
                .iter()
                .map(|s| ViewerChartSeries {
                    idx: s.idx(),
                    order: s.order(),
                    name: s.name().map(String::from),
                    name_ref: s.name_ref().map(String::from),
                    categories: s.categories().map(|dr| convert_chart_data_ref(dr, wb)),
                    values: s.values().map(|dr| convert_chart_data_ref(dr, wb)),
                    x_values: s.x_values().map(|dr| convert_chart_data_ref(dr, wb)),
                    bubble_sizes: s.bubble_sizes().map(|dr| convert_chart_data_ref(dr, wb)),
                    fill_color: s.fill_color().map(argb_to_css),
                    line_color: s.line_color().map(argb_to_css),
                    series_type: s.series_type().map(convert_chart_type),
                })
                .collect();

            let axes: Vec<ViewerChartAxis> = chart
                .axes()
                .iter()
                .map(|a| ViewerChartAxis {
                    id: a.id(),
                    axis_type: a.axis_type().to_string(),
                    position: Some(a.position().to_string()),
                    title: a.title().map(String::from),
                    min: a.min(),
                    max: a.max(),
                    major_unit: a.major_unit(),
                    minor_unit: a.minor_unit(),
                    major_gridlines: a.major_gridlines(),
                    minor_gridlines: a.minor_gridlines(),
                    crosses_ax: a.crosses_ax(),
                    num_fmt: a.num_fmt().map(String::from),
                    deleted: a.deleted(),
                })
                .collect();

            let legend = chart.legend().map(|l| ViewerChartLegend {
                position: l.position().to_string(),
                overlay: l.overlay(),
            });

            // Detect one-cell anchors: the reader leaves `to_*` at zero when
            // there is no `<to>` marker, and the meaningful size comes from the
            // outer anchor `<ext>`. This also handles charts anchored at A1.
            let has_positive_extent = matches!(
                (chart.extent_cx(), chart.extent_cy()),
                (Some(cx), Some(cy)) if cx > 0 && cy > 0
            );
            let is_one_cell = has_positive_extent
                && chart.to_col() == 0
                && chart.to_row() == 0
                && chart.to_col_off() == 0
                && chart.to_row_off() == 0;

            ViewerChart {
                chart_type,
                bar_direction,
                grouping,
                title: chart.title().map(String::from),
                series,
                axes,
                legend,
                vary_colors: Some(chart.vary_colors()),
                from_col: Some(chart.from_col()),
                from_row: Some(chart.from_row()),
                to_col: if is_one_cell {
                    None
                } else {
                    Some(chart.to_col())
                },
                to_row: if is_one_cell {
                    None
                } else {
                    Some(chart.to_row())
                },
                name: chart.name().map(String::from),
                extent_cx: chart.extent_cx(),
                extent_cy: chart.extent_cy(),
            }
        })
        .collect()
}

/// Convert an `offidized_xlsx::ChartDataRef` into the viewer's `ChartDataRef`.
///
/// When the data ref has a formula but no cached values (common when the
/// chart XML omits `<numCache>`), resolves the formula against the workbook
/// to populate `num_values` / `str_values`.
fn convert_chart_data_ref(
    data_ref: &offidized_xlsx::ChartDataRef,
    wb: &offidized_xlsx::Workbook,
) -> ViewerChartDataRef {
    let mut num_values = data_ref.num_values().to_vec();
    let mut str_values = data_ref.str_values().to_vec();

    // Resolve formula reference if no cached values are present
    if num_values.is_empty() && str_values.is_empty() {
        if let Some(formula) = data_ref.formula() {
            let (resolved_nums, resolved_strs) = resolve_chart_formula(formula, wb);
            num_values = resolved_nums;
            str_values = resolved_strs;
        }
    }

    ViewerChartDataRef {
        formula: data_ref.formula().map(String::from),
        num_values,
        str_values,
    }
}

/// Resolve a chart formula reference like `'Charts'!$B$2:$B$5` against workbook cells.
///
/// Returns `(num_values, str_values)`. Numbers go into `num_values`, strings into
/// `str_values`. If all resolved values are numeric, `str_values` will be empty and
/// vice versa.
fn resolve_chart_formula(
    formula: &str,
    wb: &offidized_xlsx::Workbook,
) -> (Vec<Option<f64>>, Vec<String>) {
    let mut num_values = Vec::new();
    let mut str_values = Vec::new();

    // Parse "SheetName!CellRange" — sheet name may be quoted with single quotes
    let Some((sheet_name, range_part)) = parse_sheet_formula(formula) else {
        return (num_values, str_values);
    };

    // Find the worksheet by name
    let Some(ws) = wb.worksheets().iter().find(|w| w.name() == sheet_name) else {
        return (num_values, str_values);
    };

    // Parse the cell range (handles $A$2:$B$5 and single cells like B1)
    let range_str = range_part.replace('$', "");
    let parts: Vec<&str> = range_str.split(':').collect();

    let (start_ref, end_ref) = if parts.len() == 2 {
        let Some(s) = parts.first() else {
            return (num_values, str_values);
        };
        let Some(e) = parts.get(1) else {
            return (num_values, str_values);
        };
        (*s, *e)
    } else if parts.len() == 1 {
        let Some(s) = parts.first() else {
            return (num_values, str_values);
        };
        (*s, *s)
    } else {
        return (num_values, str_values);
    };

    let Some((start_col, start_row)) = parse_cell_ref(start_ref) else {
        return (num_values, str_values);
    };
    let Some((end_col, end_row)) = parse_cell_ref(end_ref) else {
        return (num_values, str_values);
    };

    // Iterate cells in the range and collect values
    let mut has_strings = false;
    let mut has_numbers = false;

    for row in start_row..=end_row {
        for col in start_col..=end_col {
            let cell_ref = format!("{}{}", col_index_to_letters(col), row + 1);
            if let Some(cell) = ws.cell(&cell_ref) {
                match cell.value() {
                    Some(offidized_xlsx::CellValue::Number(n)) => {
                        num_values.push(Some(*n));
                        has_numbers = true;
                    }
                    Some(offidized_xlsx::CellValue::DateTime(n)) => {
                        num_values.push(Some(*n));
                        has_numbers = true;
                    }
                    Some(offidized_xlsx::CellValue::String(s)) => {
                        str_values.push(s.clone());
                        has_strings = true;
                    }
                    Some(offidized_xlsx::CellValue::Bool(b)) => {
                        num_values.push(Some(if *b { 1.0 } else { 0.0 }));
                        has_numbers = true;
                    }
                    _ => {
                        num_values.push(None);
                    }
                }
            } else {
                num_values.push(None);
            }
        }
    }

    // If all values were strings (categories), clear num_values
    if has_strings && !has_numbers {
        num_values.clear();
    }
    // If all values were numeric, clear str_values
    if has_numbers && !has_strings {
        str_values.clear();
    }

    (num_values, str_values)
}

/// Parse a chart formula like `'Sheet Name'!$A$1:$B$5` or `Sheet1!A1` into
/// `(sheet_name, range_part)`.
fn parse_sheet_formula(formula: &str) -> Option<(&str, &str)> {
    let excl = formula.find('!')?;
    let sheet_part = formula.get(..excl)?;
    let range_part = formula.get(excl + 1..)?;

    // Strip surrounding single quotes from sheet name if present
    let sheet_name = sheet_part
        .strip_prefix('\'')
        .and_then(|s| s.strip_suffix('\''))
        .unwrap_or(sheet_part);

    Some((sheet_name, range_part))
}

/// Convert a 0-indexed column number to Excel column letters (A, B, ..., Z, AA, AB, ...).
fn col_index_to_letters(mut col: u32) -> String {
    let mut result = String::new();
    loop {
        let rem = col % 26;
        // Safe: rem is 0..25, so 'A' + rem is always valid ASCII
        result.insert(0, (b'A' + rem as u8) as char);
        if col < 26 {
            break;
        }
        col = col / 26 - 1;
    }
    result
}

/// Convert drawings/images and charts from an `offidized_xlsx::Worksheet` into viewer drawings.
fn convert_drawings(ws: &offidized_xlsx::Worksheet, ws_idx: usize) -> Vec<Drawing> {
    let mut drawings: Vec<Drawing> = ws
        .images()
        .iter()
        .enumerate()
        .map(|(img_idx, img)| {
            let anchor_type = match img.anchor_type() {
                offidized_xlsx::ImageAnchorType::TwoCell => "twoCellAnchor",
                offidized_xlsx::ImageAnchorType::OneCell => "oneCellAnchor",
                offidized_xlsx::ImageAnchorType::Absolute => "absoluteAnchor",
            };

            let (from_col, from_row, from_col_off, from_row_off) =
                if let Some(anchor) = img.from_anchor() {
                    (
                        Some(anchor.col()),
                        Some(anchor.row()),
                        Some(anchor.col_offset()),
                        Some(anchor.row_offset()),
                    )
                } else {
                    (None, None, None, None)
                };

            let (to_col, to_row, to_col_off, to_row_off) = if let Some(anchor) = img.to_anchor() {
                (
                    Some(anchor.col()),
                    Some(anchor.row()),
                    Some(anchor.col_offset()),
                    Some(anchor.row_offset()),
                )
            } else {
                (None, None, None, None)
            };

            Drawing {
                anchor_type: anchor_type.to_string(),
                drawing_type: "picture".to_string(),
                name: img.name().map(String::from),
                description: img.description().map(String::from),
                title: None,
                from_col,
                from_row,
                from_col_off,
                from_row_off,
                to_col,
                to_row,
                to_col_off,
                to_row_off,
                pos_x: img.position_x(),
                pos_y: img.position_y(),
                extent_cx: img.extent_cx(),
                extent_cy: img.extent_cy(),
                edit_as: None,
                image_id: Some(format!("ws{ws_idx}_img{img_idx}")),
                chart_id: None,
                shape_type: None,
                fill_color: None,
                line_color: None,
                text_content: None,
                rotation: None,
                flip_h: None,
                flip_v: None,
                hyperlink: None,
                xfrm_x: None,
                xfrm_y: None,
                xfrm_cx: img.extent_cx(),
                xfrm_cy: img.extent_cy(),
            }
        })
        .collect();

    // Also create Drawing entries for charts so they appear in the drawings list
    for (idx, chart) in ws.charts().iter().enumerate() {
        drawings.push(Drawing {
            anchor_type: "twoCellAnchor".to_string(),
            drawing_type: "chart".to_string(),
            name: chart.name().map(String::from),
            description: None,
            title: chart.title().map(String::from),
            from_col: Some(chart.from_col()),
            from_row: Some(chart.from_row()),
            from_col_off: Some(chart.from_col_off()),
            from_row_off: Some(chart.from_row_off()),
            to_col: Some(chart.to_col()),
            to_row: Some(chart.to_row()),
            to_col_off: Some(chart.to_col_off()),
            to_row_off: Some(chart.to_row_off()),
            pos_x: None,
            pos_y: None,
            extent_cx: None,
            extent_cy: None,
            edit_as: None,
            image_id: None,
            chart_id: Some(format!("chart-{idx}")),
            shape_type: None,
            fill_color: None,
            line_color: None,
            text_content: None,
            rotation: None,
            flip_h: None,
            flip_v: None,
            hyperlink: None,
            xfrm_x: None,
            xfrm_y: None,
            xfrm_cx: None,
            xfrm_cy: None,
        });
    }

    drawings
}

/// Convert all worksheet images into `EmbeddedImage` entries for the workbook.
fn convert_all_images(wb: &offidized_xlsx::Workbook) -> Vec<EmbeddedImage> {
    let mut images = Vec::new();
    for (ws_idx, ws) in wb.worksheets().iter().enumerate() {
        for (img_idx, img) in ws.images().iter().enumerate() {
            let bytes = img.bytes();
            let format = ImageFormat::from_magic_bytes(bytes);
            let mime_type = if img.content_type().is_empty() {
                format.mime_type().to_string()
            } else {
                img.content_type().to_string()
            };
            let data = BASE64_STANDARD.encode(bytes);
            let id = format!("ws{ws_idx}_img{img_idx}");

            images.push(EmbeddedImage {
                id,
                mime_type,
                data,
                filename: None,
                width: None,
                height: None,
            });
        }
    }
    images
}

/// Convert sparkline groups from an `offidized_xlsx::Worksheet`.
fn convert_sparkline_groups(ws: &offidized_xlsx::Worksheet) -> Vec<ViewerSparklineGroup> {
    ws.sparkline_groups()
        .iter()
        .map(|sg| {
            let sparkline_type = sg.sparkline_type().as_str().to_string();

            let sparklines: Vec<ViewerSparkline> = sg
                .sparklines()
                .iter()
                .map(|s| ViewerSparkline {
                    location: s.location().to_string(),
                    data_range: s.data_range().to_string(),
                })
                .collect();

            let colors = ViewerSparklineColors {
                series: sg.colors().series.as_ref().map(|c| argb_to_css(c)),
                negative: sg.colors().negative.as_ref().map(|c| argb_to_css(c)),
                axis: sg.colors().axis.as_ref().map(|c| argb_to_css(c)),
                markers: sg.colors().markers.as_ref().map(|c| argb_to_css(c)),
                first: sg.colors().first.as_ref().map(|c| argb_to_css(c)),
                last: sg.colors().last.as_ref().map(|c| argb_to_css(c)),
                high: sg.colors().high.as_ref().map(|c| argb_to_css(c)),
                low: sg.colors().low.as_ref().map(|c| argb_to_css(c)),
            };

            let display_empty_cells_as = Some(sg.display_empty_cells_as().as_str().to_string());

            ViewerSparklineGroup {
                sparkline_type,
                sparklines,
                colors,
                display_empty_cells_as,
                markers: sg.markers(),
                high_point: sg.high_point(),
                low_point: sg.low_point(),
                first_point: sg.first_point(),
                last_point: sg.last_point(),
                negative_points: sg.negative_points(),
                display_x_axis: sg.display_x_axis(),
                display_hidden: Some(sg.display_hidden()),
                right_to_left: sg.right_to_left(),
                line_weight: sg.line_weight(),
                min_axis_type: Some(sg.min_axis_type().as_str().to_string()),
                max_axis_type: Some(sg.max_axis_type().as_str().to_string()),
                manual_min: sg.manual_min(),
                manual_max: sg.manual_max(),
                date_axis: Some(sg.date_axis()),
            }
        })
        .collect()
}

/// Convert conditional formatting rules from an `offidized_xlsx::Worksheet`.
fn convert_conditional_formattings(
    ws: &offidized_xlsx::Worksheet,
    _theme_colors: &[String],
    _indexed_colors: Option<&[String]>,
) -> Vec<ViewerCF> {
    // offidized-xlsx stores one ConditionalFormatting per cfRule, but the viewer
    // expects rules grouped by sqref (matching the original OOXML structure).
    // Group rules that share the same sqref into a single ViewerCF.
    let mut grouped: Vec<ViewerCF> = Vec::new();

    for cf in ws.conditional_formattings() {
        // Build sqref string from ranges, collapsing single-cell ranges.
        // Fall back to raw_sqref for ranges CellRange can't represent (e.g. "A:A").
        let sqref = if cf.sqref().is_empty() {
            cf.raw_sqref().unwrap_or_default().to_string()
        } else {
            cf.sqref()
                .iter()
                .map(|r| {
                    if r.start() == r.end() {
                        r.start().to_string()
                    } else {
                        format!("{}:{}", r.start(), r.end())
                    }
                })
                .collect::<Vec<_>>()
                .join(" ")
        };

        let rule_type = convert_cf_rule_type(cf.rule_type());

        // Build the single rule from this ConditionalFormatting
        let mut rule = CFRule {
            rule_type,
            priority: cf.priority().unwrap_or(1),
            color_scale: None,
            data_bar: None,
            icon_set: None,
            formula: cf.formulas().first().map(String::from),
            operator: cf.operator().map(convert_cf_operator),
            dxf_id: cf.dxf_id(),
            rank: cf.rank(),
            percent: cf.cf_percent(),
            bottom: cf.cf_bottom(),
            above_average: cf.above_average(),
            equal_average: cf.equal_average(),
            std_dev: cf.std_dev(),
            time_period: cf.time_period().map(String::from),
        };

        // Color scale
        if !cf.color_scale_stops().is_empty() {
            let cfvo: Vec<CFValueObject> = cf
                .color_scale_stops()
                .iter()
                .map(|stop| CFValueObject {
                    cfvo_type: convert_cfvo_type(stop.cfvo.value_type),
                    val: stop.cfvo.value.clone(),
                })
                .collect();
            let colors: Vec<String> = cf
                .color_scale_stops()
                .iter()
                .map(|stop| argb_to_css(&stop.color))
                .collect();
            rule.color_scale = Some(ColorScale { cfvo, colors });
        }

        // Data bar
        if let (Some(db_min), Some(db_max)) = (cf.data_bar_min(), cf.data_bar_max()) {
            let cfvo = vec![
                CFValueObject {
                    cfvo_type: convert_cfvo_type(db_min.value_type),
                    val: db_min.value.clone(),
                },
                CFValueObject {
                    cfvo_type: convert_cfvo_type(db_max.value_type),
                    val: db_max.value.clone(),
                },
            ];
            let color = cf
                .data_bar_color()
                .map(argb_to_css)
                .unwrap_or_else(|| "#638EC6".to_string());
            rule.data_bar = Some(DataBar {
                cfvo,
                color,
                show_value: cf.data_bar_show_value(),
                min_length: cf.data_bar_min_length(),
                max_length: cf.data_bar_max_length(),
            });
        }

        // Icon set — default name is "3TrafficLights1" per OOXML spec
        if cf.icon_set_name().is_some() || !cf.icon_set_values().is_empty() {
            let name = cf.icon_set_name().unwrap_or("3TrafficLights1").to_string();
            let cfvo: Vec<CFValueObject> = cf
                .icon_set_values()
                .iter()
                .map(|v| CFValueObject {
                    cfvo_type: convert_cfvo_type(v.value_type),
                    val: v.value.clone(),
                })
                .collect();
            rule.icon_set = Some(IconSet {
                icon_set: name,
                cfvo,
                show_value: cf.icon_set_show_value(),
                reverse: cf.icon_set_reverse(),
            });
        }

        // Group rules by sqref: if the last ViewerCF has the same sqref, append the rule
        if let Some(last) = grouped.last_mut() {
            if last.sqref == sqref {
                last.rules.push(rule);
                continue;
            }
        }
        grouped.push(ViewerCF {
            sqref,
            rules: vec![rule],
        });
    }

    grouped
}

/// Convert DXF (differential formatting) styles from the workbook.
fn convert_dxf_styles(
    wb: &offidized_xlsx::Workbook,
    _theme_colors: &[String],
    _indexed_colors: Option<&[String]>,
) -> Vec<DxfStyle> {
    // TODO: offidized-xlsx needs a StyleTable.dxfs() API to expose DXF entries.
    // For now, build empty DXF entries matching the count referenced by CF rules
    // and auto-filter color filters. This avoids index-out-of-bounds when rules
    // reference dxf_id.
    let max_cf_dxf = wb
        .worksheets()
        .iter()
        .flat_map(|ws| ws.conditional_formattings())
        .filter_map(|cf| cf.dxf_id())
        .max();

    let max_af_dxf = wb
        .worksheets()
        .iter()
        .filter_map(|ws| ws.auto_filter())
        .flat_map(|af| af.filter_columns())
        .filter_map(|fc| fc.dxf_id())
        .max();

    let count = match (max_cf_dxf, max_af_dxf) {
        (Some(a), Some(b)) => std::cmp::max(a, b) as usize + 1,
        (Some(a), None) => a as usize + 1,
        (None, Some(b)) => b as usize + 1,
        (None, None) => return Vec::new(),
    };

    vec![DxfStyle::default(); count]
}

/// Convert auto-filter from an `offidized_xlsx::Worksheet`.
fn convert_auto_filter(ws: &offidized_xlsx::Worksheet) -> Option<ViewerAutoFilter> {
    let af = ws.auto_filter()?;
    let range_obj = af.range()?;

    let range = format!("{}:{}", range_obj.start(), range_obj.end());
    // CellRange stores 1-based indices; convert to 0-based for the viewer
    let start_row = range_obj.start_row().saturating_sub(1);
    let start_col = range_obj.start_column().saturating_sub(1);
    let end_row = range_obj.end_row().saturating_sub(1);
    let end_col = range_obj.end_column().saturating_sub(1);

    let filter_columns: Vec<ViewerFilterColumn> = af
        .filter_columns()
        .iter()
        .map(|fc| {
            let filter_type = convert_filter_type(fc.filter_type());

            let custom_filters: Vec<ViewerCustomFilter> = fc
                .custom_filters()
                .iter()
                .map(|cf| ViewerCustomFilter {
                    operator: convert_custom_filter_operator(cf.operator()),
                    val: cf.val().to_string(),
                })
                .collect();

            ViewerFilterColumn {
                col_id: fc.col_id(),
                has_filter: fc.has_filter(),
                filter_type,
                show_button: fc.show_button(),
                blank: fc.blank(),
                values: fc.values().iter().map(|v| v.to_string()).collect(),
                custom_filters,
                custom_filters_and: fc.custom_filters_and(),
                dxf_id: fc.dxf_id(),
                cell_color: fc.cell_color(),
                // Viewer icon_set is Option<u32> but source is Option<&str> (icon set name).
                // We don't have a numeric mapping, so leave as None.
                icon_set: None,
                icon_id: fc.icon_id(),
                dynamic_type: fc.dynamic_type().map(String::from),
                top: fc.top(),
                percent: fc.percent(),
                top10_val: fc.top10_val(),
            }
        })
        .collect();

    Some(ViewerAutoFilter {
        range,
        start_row,
        start_col,
        end_row,
        end_col,
        filter_columns,
    })
}

/// Convert data validations from an `offidized_xlsx::Worksheet`.
fn convert_data_validations(ws: &offidized_xlsx::Worksheet) -> Vec<DataValidationRange> {
    ws.data_validations()
        .iter()
        .map(|dv| {
            let sqref = dv
                .sqref()
                .iter()
                .map(|r| format!("{}:{}", r.start(), r.end()))
                .collect::<Vec<_>>()
                .join(" ");

            let validation_type = convert_validation_type(dv.validation_type());

            // Parse inline list values from formula1 for list-type validations
            let list_values = if matches!(
                dv.validation_type(),
                offidized_xlsx::DataValidationType::List
            ) {
                let f1 = dv.formula1();
                // Inline comma-separated values when formula doesn't reference a range
                if f1.starts_with('"') || (!f1.contains('!') && !f1.starts_with('$')) {
                    Some(
                        f1.split(',')
                            .map(|v| v.trim().trim_matches('"').to_string())
                            .collect(),
                    )
                } else {
                    None
                }
            } else {
                None
            };

            let validation = ViewerDataValidation {
                validation_type,
                operator: None, // TODO: offidized-xlsx DataValidation needs operator API
                formula1: Some(dv.formula1().to_string()),
                formula2: dv.formula2().map(String::from),
                allow_blank: true,
                show_dropdown: matches!(
                    dv.validation_type(),
                    offidized_xlsx::DataValidationType::List
                ),
                show_input_message: dv.show_input_message().unwrap_or(true),
                show_error_message: dv.show_error_message().unwrap_or(true),
                error_title: dv.error_title().map(String::from),
                error_message: dv.error_message().map(String::from),
                prompt_title: dv.prompt_title().map(String::from),
                prompt_message: dv.prompt_message().map(String::from),
                list_values,
            };

            DataValidationRange { sqref, validation }
        })
        .collect()
}

// ============================================================================
// Style conversion
// ============================================================================

/// Convert an `offidized_xlsx::Style` into the viewer's `StyleRef`.
fn convert_style(
    style: &offidized_xlsx::Style,
    theme_colors: &[String],
    indexed_colors: Option<&[String]>,
) -> StyleRef {
    let mut s = Style::default();

    // Font
    if let Some(font) = style.font() {
        s.font_family = font.name().map(String::from);
        s.font_size = font.size().and_then(|sz| sz.parse::<f64>().ok());
        s.bold = font.bold();
        s.italic = font.italic();

        // Underline: offidized-xlsx Font only exposes a bool, not the style variant.
        // TODO: offidized-xlsx Font needs underline_style() API for Double/SingleAccounting/etc.
        if font.underline() == Some(true) {
            s.underline = Some(crate::types::style::UnderlineStyle::Single);
        }
        s.strikethrough = font.strikethrough();

        // Font vertical alignment (superscript/subscript)
        s.vert_align = font.vertical_align().map(|va| match va {
            offidized_xlsx::FontVerticalAlign::Superscript => VertAlign::Superscript,
            offidized_xlsx::FontVerticalAlign::Subscript => VertAlign::Subscript,
            offidized_xlsx::FontVerticalAlign::Baseline => VertAlign::Baseline,
        });

        // Resolve font color
        s.font_color = font
            .color_ref()
            .and_then(|cr| offidized_xlsx::resolve_color(cr, theme_colors, indexed_colors));
    }

    // Fill
    if let Some(fill) = style.fill() {
        if let Some(pf) = fill.pattern_fill() {
            let pt = convert_pattern_type(pf.pattern_type());

            let fg = pf
                .fg_color()
                .and_then(|cr| offidized_xlsx::resolve_color(cr, theme_colors, indexed_colors));
            let bg = pf
                .bg_color()
                .and_then(|cr| offidized_xlsx::resolve_color(cr, theme_colors, indexed_colors));

            // For solid fills, just set bg_color to the foreground color without
            // emitting pattern_type — the renderer treats bg_color alone as a
            // solid background.
            if pt == PatternType::Solid {
                s.bg_color = fg;
            } else {
                s.pattern_type = Some(pt);
                s.fg_color = fg;
                s.bg_color = bg;
            }
        }
    }

    // Gradient fill
    if let Some(gf) = style.gradient_fill() {
        let gradient_type = match gf.gradient_type() {
            offidized_xlsx::GradientFillType::Linear => "linear",
            offidized_xlsx::GradientFillType::Path => "path",
        };
        let stops: Vec<GradientStop> = gf
            .stops()
            .iter()
            .map(|stop| {
                // Resolve color: prefer theme/indexed resolution, fall back to raw ARGB
                let color = if let Some(cr) = stop.color_ref() {
                    offidized_xlsx::resolve_color(cr, theme_colors, indexed_colors)
                        .unwrap_or_else(|| argb_to_css(stop.color()))
                } else {
                    argb_to_css(stop.color())
                };
                GradientStop {
                    position: stop.position(),
                    color,
                }
            })
            .collect();
        s.gradient = Some(GradientFill {
            gradient_type: gradient_type.to_string(),
            degree: gf.degree(),
            left: gf.left(),
            right: gf.right(),
            top: gf.top(),
            bottom: gf.bottom(),
            stops,
        });
    }

    // Border
    if let Some(border) = style.border() {
        s.border_left = convert_border_side(border.left(), theme_colors, indexed_colors);
        s.border_right = convert_border_side(border.right(), theme_colors, indexed_colors);
        s.border_top = convert_border_side(border.top(), theme_colors, indexed_colors);
        s.border_bottom = convert_border_side(border.bottom(), theme_colors, indexed_colors);
        s.border_diagonal = convert_border_side(border.diagonal(), theme_colors, indexed_colors);
        s.diagonal_up = border.diagonal_up();
        s.diagonal_down = border.diagonal_down();
    }

    // Alignment
    if let Some(alignment) = style.alignment() {
        s.align_h = alignment.horizontal().map(convert_halign);
        s.align_v = alignment.vertical().map(convert_valign);
        s.wrap = alignment.wrap_text();
        s.indent = alignment.indent();
        s.rotation = alignment.text_rotation().map(|r| r as i32);
        s.shrink_to_fit = alignment.shrink_to_fit();
        s.reading_order = alignment
            .reading_order()
            .and_then(|ro| u8::try_from(ro).ok());
    }

    // Protection
    if let Some(prot) = style.protection() {
        s.locked = prot.locked();
        s.hidden = prot.hidden();
    }

    StyleRef(Arc::new(s))
}

/// Convert an `offidized_xlsx::Style` directly to a `CellStyleData` for the render cache.
pub fn style_to_render(
    style: &offidized_xlsx::Style,
    theme_colors: &[String],
    indexed_colors: Option<&[String]>,
) -> CellStyleData {
    let sr = convert_style(style, theme_colors, indexed_colors);
    style_ref_to_cell_style_data(&sr)
}

/// Build the render style cache from an `offidized_xlsx::Workbook`.
pub fn build_render_style_cache(wb: &offidized_xlsx::Workbook) -> Vec<Option<CellStyleData>> {
    let theme_colors = wb.theme_colors();
    let indexed_colors = wb.indexed_colors();

    wb.styles()
        .styles()
        .iter()
        .map(|style| {
            let sr = convert_style(style, theme_colors, indexed_colors);
            Some(style_ref_to_cell_style_data(&sr))
        })
        .collect()
}

/// Convert a `StyleRef` to `CellStyleData` for the render pipeline.
#[allow(clippy::cast_possible_truncation)]
fn style_ref_to_cell_style_data(style: &StyleRef) -> CellStyleData {
    CellStyleData {
        bg_color: style.bg_color.clone(),
        font_color: style.font_color.clone(),
        font_size: style.font_size.map(|f| f as f32),
        font_family: style.font_family.clone(),
        bold: style.bold,
        italic: style.italic,
        underline: style.underline.is_some().then_some(true),
        strikethrough: style.strikethrough,
        rotation: style.rotation,
        indent: style.indent,
        align_h: style.align_h.as_ref().map(halign_to_string),
        align_v: style.align_v.as_ref().map(valign_to_string),
        wrap_text: style.wrap,
        border_top: style.border_top.as_ref().map(|b| BorderStyleData {
            style: Some(border_style_to_string(&b.style)),
            color: Some(b.color.clone()),
        }),
        border_right: style.border_right.as_ref().map(|b| BorderStyleData {
            style: Some(border_style_to_string(&b.style)),
            color: Some(b.color.clone()),
        }),
        border_bottom: style.border_bottom.as_ref().map(|b| BorderStyleData {
            style: Some(border_style_to_string(&b.style)),
            color: Some(b.color.clone()),
        }),
        border_left: style.border_left.as_ref().map(|b| BorderStyleData {
            style: Some(border_style_to_string(&b.style)),
            color: Some(b.color.clone()),
        }),
        border_diagonal_down: if style.diagonal_down == Some(true) {
            style.border_diagonal.as_ref().map(|b| BorderStyleData {
                style: Some(border_style_to_string(&b.style)),
                color: Some(b.color.clone()),
            })
        } else {
            None
        },
        border_diagonal_up: if style.diagonal_up == Some(true) {
            style.border_diagonal.as_ref().map(|b| BorderStyleData {
                style: Some(border_style_to_string(&b.style)),
                color: Some(b.color.clone()),
            })
        } else {
            None
        },
        pattern_type: style.pattern_type.as_ref().map(pattern_type_to_string),
        pattern_fg_color: style.fg_color.clone(),
        pattern_bg_color: style.bg_color.clone(),
    }
}

// ============================================================================
// Helper conversion functions
// ============================================================================

/// Convert an ARGB hex string (e.g. "FF0000FF") to a CSS color ("#0000FF").
fn argb_to_css(argb: &str) -> String {
    if argb.len() == 8 {
        // Strip the alpha channel (first 2 chars)
        format!("#{}", &argb[2..])
    } else {
        format!("#{argb}")
    }
}

/// Convert `offidized_xlsx::ChartType` to the viewer's `ChartType`.
fn convert_chart_type(ct: offidized_xlsx::ChartType) -> ViewerChartType {
    match ct {
        offidized_xlsx::ChartType::Bar => ViewerChartType::Bar,
        offidized_xlsx::ChartType::Line => ViewerChartType::Line,
        offidized_xlsx::ChartType::Pie => ViewerChartType::Pie,
        offidized_xlsx::ChartType::Area => ViewerChartType::Area,
        offidized_xlsx::ChartType::Scatter => ViewerChartType::Scatter,
        offidized_xlsx::ChartType::Doughnut => ViewerChartType::Doughnut,
        offidized_xlsx::ChartType::Radar => ViewerChartType::Radar,
        offidized_xlsx::ChartType::Bubble => ViewerChartType::Bubble,
        offidized_xlsx::ChartType::Stock => ViewerChartType::Stock,
        offidized_xlsx::ChartType::Surface => ViewerChartType::Surface,
        offidized_xlsx::ChartType::Combo => ViewerChartType::Combo,
    }
}

/// Convert `offidized_xlsx::BarDirection` to the viewer's `BarDirection`.
fn convert_bar_direction(bd: offidized_xlsx::BarDirection) -> ViewerBarDirection {
    match bd {
        offidized_xlsx::BarDirection::Column => ViewerBarDirection::Col,
        offidized_xlsx::BarDirection::Bar => ViewerBarDirection::Bar,
    }
}

/// Convert `offidized_xlsx::ChartGrouping` to the viewer's `ChartGrouping`.
fn convert_chart_grouping(g: offidized_xlsx::ChartGrouping) -> ViewerChartGrouping {
    match g {
        offidized_xlsx::ChartGrouping::Standard => ViewerChartGrouping::Standard,
        offidized_xlsx::ChartGrouping::Stacked => ViewerChartGrouping::Stacked,
        offidized_xlsx::ChartGrouping::PercentStacked => ViewerChartGrouping::PercentStacked,
        offidized_xlsx::ChartGrouping::Clustered => ViewerChartGrouping::Clustered,
    }
}

/// Convert `offidized_xlsx::ConditionalFormattingRuleType` to the viewer's `CFRuleType`.
fn convert_cf_rule_type(rt: offidized_xlsx::ConditionalFormattingRuleType) -> CFRuleType {
    match rt {
        offidized_xlsx::ConditionalFormattingRuleType::CellIs => CFRuleType::CellIs,
        offidized_xlsx::ConditionalFormattingRuleType::Expression => CFRuleType::Expression,
        offidized_xlsx::ConditionalFormattingRuleType::ColorScale => CFRuleType::ColorScale,
        offidized_xlsx::ConditionalFormattingRuleType::DataBar => CFRuleType::DataBar,
        offidized_xlsx::ConditionalFormattingRuleType::IconSet => CFRuleType::IconSet,
        offidized_xlsx::ConditionalFormattingRuleType::Top10 => CFRuleType::Top10,
        offidized_xlsx::ConditionalFormattingRuleType::AboveAverage => CFRuleType::AboveAverage,
        offidized_xlsx::ConditionalFormattingRuleType::TimePeriod => CFRuleType::TimePeriod,
        offidized_xlsx::ConditionalFormattingRuleType::DuplicateValues => {
            CFRuleType::DuplicateValues
        }
        offidized_xlsx::ConditionalFormattingRuleType::UniqueValues => CFRuleType::UniqueValues,
        offidized_xlsx::ConditionalFormattingRuleType::ContainsText => CFRuleType::ContainsText,
        offidized_xlsx::ConditionalFormattingRuleType::NotContainsText => {
            CFRuleType::NotContainsText
        }
        offidized_xlsx::ConditionalFormattingRuleType::BeginsWith => CFRuleType::BeginsWith,
        offidized_xlsx::ConditionalFormattingRuleType::EndsWith => CFRuleType::EndsWith,
        offidized_xlsx::ConditionalFormattingRuleType::ContainsBlanks => CFRuleType::ContainsBlanks,
        offidized_xlsx::ConditionalFormattingRuleType::NotContainsBlanks => {
            CFRuleType::NotContainsBlanks
        }
        offidized_xlsx::ConditionalFormattingRuleType::ContainsErrors => CFRuleType::ContainsErrors,
        offidized_xlsx::ConditionalFormattingRuleType::NotContainsErrors => {
            CFRuleType::NotContainsErrors
        }
    }
}

/// Convert a `CfValueObjectType` to its XML string representation.
///
/// The upstream `as_xml_value()` method is `pub(crate)`, so we replicate the
/// mapping here for the viewer adapter.
fn convert_cfvo_type(vt: offidized_xlsx::CfValueObjectType) -> String {
    match vt {
        offidized_xlsx::CfValueObjectType::Num => "num".to_string(),
        offidized_xlsx::CfValueObjectType::Percent => "percent".to_string(),
        offidized_xlsx::CfValueObjectType::Max => "max".to_string(),
        offidized_xlsx::CfValueObjectType::Min => "min".to_string(),
        offidized_xlsx::CfValueObjectType::Formula => "formula".to_string(),
        offidized_xlsx::CfValueObjectType::Percentile => "percentile".to_string(),
        offidized_xlsx::CfValueObjectType::AutoMin => "autoMin".to_string(),
        offidized_xlsx::CfValueObjectType::AutoMax => "autoMax".to_string(),
    }
}

/// Convert a conditional formatting operator to its XML string representation.
fn convert_cf_operator(op: offidized_xlsx::ConditionalFormattingOperator) -> String {
    match op {
        offidized_xlsx::ConditionalFormattingOperator::LessThan => "lessThan".to_string(),
        offidized_xlsx::ConditionalFormattingOperator::LessThanOrEqual => {
            "lessThanOrEqual".to_string()
        }
        offidized_xlsx::ConditionalFormattingOperator::Equal => "equal".to_string(),
        offidized_xlsx::ConditionalFormattingOperator::NotEqual => "notEqual".to_string(),
        offidized_xlsx::ConditionalFormattingOperator::GreaterThanOrEqual => {
            "greaterThanOrEqual".to_string()
        }
        offidized_xlsx::ConditionalFormattingOperator::GreaterThan => "greaterThan".to_string(),
        offidized_xlsx::ConditionalFormattingOperator::Between => "between".to_string(),
        offidized_xlsx::ConditionalFormattingOperator::NotBetween => "notBetween".to_string(),
        offidized_xlsx::ConditionalFormattingOperator::ContainsText => "containsText".to_string(),
        offidized_xlsx::ConditionalFormattingOperator::NotContains => "notContains".to_string(),
        offidized_xlsx::ConditionalFormattingOperator::BeginsWith => "beginsWith".to_string(),
        offidized_xlsx::ConditionalFormattingOperator::EndsWith => "endsWith".to_string(),
    }
}

/// Convert `offidized_xlsx::FilterType` to the viewer's `FilterType`.
fn convert_filter_type(ft: offidized_xlsx::FilterType) -> ViewerFilterType {
    match ft {
        offidized_xlsx::FilterType::None => ViewerFilterType::None,
        offidized_xlsx::FilterType::Values => ViewerFilterType::Values,
        offidized_xlsx::FilterType::Custom => ViewerFilterType::Custom,
        offidized_xlsx::FilterType::Top10 => ViewerFilterType::Top10,
        offidized_xlsx::FilterType::Dynamic => ViewerFilterType::Dynamic,
        offidized_xlsx::FilterType::Color => ViewerFilterType::Color,
        offidized_xlsx::FilterType::Icon => ViewerFilterType::Icon,
    }
}

/// Convert `offidized_xlsx::CustomFilterOperator` to the viewer's `CustomFilterOperator`.
fn convert_custom_filter_operator(
    op: offidized_xlsx::CustomFilterOperator,
) -> ViewerCustomFilterOperator {
    match op {
        offidized_xlsx::CustomFilterOperator::Equal => ViewerCustomFilterOperator::Equal,
        offidized_xlsx::CustomFilterOperator::NotEqual => ViewerCustomFilterOperator::NotEqual,
        offidized_xlsx::CustomFilterOperator::GreaterThan => {
            ViewerCustomFilterOperator::GreaterThan
        }
        offidized_xlsx::CustomFilterOperator::GreaterThanOrEqual => {
            ViewerCustomFilterOperator::GreaterThanOrEqual
        }
        offidized_xlsx::CustomFilterOperator::LessThan => ViewerCustomFilterOperator::LessThan,
        offidized_xlsx::CustomFilterOperator::LessThanOrEqual => {
            ViewerCustomFilterOperator::LessThanOrEqual
        }
    }
}

/// Convert `offidized_xlsx::DataValidationType` to the viewer's `ValidationType`.
fn convert_validation_type(vt: offidized_xlsx::DataValidationType) -> ValidationType {
    match vt {
        offidized_xlsx::DataValidationType::List => ValidationType::List,
        offidized_xlsx::DataValidationType::Whole => ValidationType::Whole,
        offidized_xlsx::DataValidationType::Decimal => ValidationType::Decimal,
        offidized_xlsx::DataValidationType::Date => ValidationType::Date,
        offidized_xlsx::DataValidationType::TextLength => ValidationType::TextLength,
        offidized_xlsx::DataValidationType::Custom => ValidationType::Custom,
        offidized_xlsx::DataValidationType::Time => ValidationType::Time,
    }
}

/// Convert `PatternFillType` to the viewer's `PatternType`.
fn convert_pattern_type(pt: offidized_xlsx::PatternFillType) -> PatternType {
    match pt {
        offidized_xlsx::PatternFillType::None => PatternType::None,
        offidized_xlsx::PatternFillType::Solid => PatternType::Solid,
        offidized_xlsx::PatternFillType::Gray125 => PatternType::Gray125,
        offidized_xlsx::PatternFillType::Gray0625 => PatternType::Gray0625,
        offidized_xlsx::PatternFillType::DarkGray => PatternType::DarkGray,
        offidized_xlsx::PatternFillType::MediumGray => PatternType::MediumGray,
        offidized_xlsx::PatternFillType::LightGray => PatternType::LightGray,
        offidized_xlsx::PatternFillType::DarkHorizontal => PatternType::DarkHorizontal,
        offidized_xlsx::PatternFillType::DarkVertical => PatternType::DarkVertical,
        offidized_xlsx::PatternFillType::DarkDown => PatternType::DarkDown,
        offidized_xlsx::PatternFillType::DarkUp => PatternType::DarkUp,
        offidized_xlsx::PatternFillType::DarkGrid => PatternType::DarkGrid,
        offidized_xlsx::PatternFillType::DarkTrellis => PatternType::DarkTrellis,
        offidized_xlsx::PatternFillType::LightHorizontal => PatternType::LightHorizontal,
        offidized_xlsx::PatternFillType::LightVertical => PatternType::LightVertical,
        offidized_xlsx::PatternFillType::LightDown => PatternType::LightDown,
        offidized_xlsx::PatternFillType::LightUp => PatternType::LightUp,
        offidized_xlsx::PatternFillType::LightGrid => PatternType::LightGrid,
        offidized_xlsx::PatternFillType::LightTrellis => PatternType::LightTrellis,
    }
}

/// Convert a `BorderSide` from offidized-xlsx to the viewer's `Border`.
fn convert_border_side(
    side: Option<&offidized_xlsx::BorderSide>,
    theme_colors: &[String],
    indexed_colors: Option<&[String]>,
) -> Option<Border> {
    let side = side?;
    let style_str = side.style()?;
    let border_style = parse_border_style(style_str);
    if border_style == BorderStyle::None {
        return None;
    }
    let color = side
        .color_ref()
        .and_then(|cr| offidized_xlsx::resolve_color(cr, theme_colors, indexed_colors))
        .unwrap_or_else(|| "#000000".to_string());
    Some(Border {
        style: border_style,
        color,
    })
}

/// Parse a border style string into a `BorderStyle` enum.
fn parse_border_style(s: &str) -> BorderStyle {
    match s {
        "thin" => BorderStyle::Thin,
        "medium" => BorderStyle::Medium,
        "thick" => BorderStyle::Thick,
        "dashed" => BorderStyle::Dashed,
        "dotted" => BorderStyle::Dotted,
        "double" => BorderStyle::Double,
        "hair" => BorderStyle::Hair,
        "mediumDashed" => BorderStyle::MediumDashed,
        "dashDot" => BorderStyle::DashDot,
        "mediumDashDot" => BorderStyle::MediumDashDot,
        "dashDotDot" => BorderStyle::DashDotDot,
        "mediumDashDotDot" => BorderStyle::MediumDashDotDot,
        "slantDashDot" => BorderStyle::SlantDashDot,
        "none" | "" => BorderStyle::None,
        _ => BorderStyle::None,
    }
}

/// Convert `HorizontalAlignment` to `HAlign`.
fn convert_halign(ha: offidized_xlsx::HorizontalAlignment) -> HAlign {
    match ha {
        offidized_xlsx::HorizontalAlignment::General => HAlign::General,
        offidized_xlsx::HorizontalAlignment::Left => HAlign::Left,
        offidized_xlsx::HorizontalAlignment::Center => HAlign::Center,
        offidized_xlsx::HorizontalAlignment::Right => HAlign::Right,
        offidized_xlsx::HorizontalAlignment::Fill => HAlign::Fill,
        offidized_xlsx::HorizontalAlignment::Justify => HAlign::Justify,
        offidized_xlsx::HorizontalAlignment::CenterContinuous => HAlign::CenterContinuous,
        offidized_xlsx::HorizontalAlignment::Distributed => HAlign::Distributed,
    }
}

/// Convert `VerticalAlignment` to `VAlign`.
fn convert_valign(va: offidized_xlsx::VerticalAlignment) -> VAlign {
    match va {
        offidized_xlsx::VerticalAlignment::Top => VAlign::Top,
        offidized_xlsx::VerticalAlignment::Center => VAlign::Center,
        offidized_xlsx::VerticalAlignment::Bottom => VAlign::Bottom,
        offidized_xlsx::VerticalAlignment::Justify => VAlign::Justify,
        offidized_xlsx::VerticalAlignment::Distributed => VAlign::Distributed,
    }
}

/// Convert `HAlign` to a lowercase string for the render pipeline.
fn halign_to_string(h: &HAlign) -> String {
    match h {
        HAlign::General => "general".to_string(),
        HAlign::Left => "left".to_string(),
        HAlign::Center => "center".to_string(),
        HAlign::Right => "right".to_string(),
        HAlign::Fill => "fill".to_string(),
        HAlign::Justify => "justify".to_string(),
        HAlign::CenterContinuous => "centercontinuous".to_string(),
        HAlign::Distributed => "distributed".to_string(),
    }
}

/// Convert `VAlign` to a lowercase string for the render pipeline.
fn valign_to_string(v: &VAlign) -> String {
    match v {
        VAlign::Top => "top".to_string(),
        VAlign::Center => "center".to_string(),
        VAlign::Bottom => "bottom".to_string(),
        VAlign::Justify => "justify".to_string(),
        VAlign::Distributed => "distributed".to_string(),
    }
}

/// Convert `BorderStyle` to a lowercase string for the render pipeline.
fn border_style_to_string(bs: &BorderStyle) -> String {
    match bs {
        BorderStyle::None => "none".to_string(),
        BorderStyle::Thin => "thin".to_string(),
        BorderStyle::Medium => "medium".to_string(),
        BorderStyle::Thick => "thick".to_string(),
        BorderStyle::Dashed => "dashed".to_string(),
        BorderStyle::Dotted => "dotted".to_string(),
        BorderStyle::Double => "double".to_string(),
        BorderStyle::Hair => "hair".to_string(),
        BorderStyle::MediumDashed => "mediumdashed".to_string(),
        BorderStyle::DashDot => "dashdot".to_string(),
        BorderStyle::MediumDashDot => "mediumdashdot".to_string(),
        BorderStyle::DashDotDot => "dashdotdot".to_string(),
        BorderStyle::MediumDashDotDot => "mediumdashdotdot".to_string(),
        BorderStyle::SlantDashDot => "slantdashdot".to_string(),
    }
}

/// Convert `PatternType` to a lowercase string for the render pipeline.
fn pattern_type_to_string(p: &PatternType) -> String {
    match p {
        PatternType::None => "none".to_string(),
        PatternType::Solid => "solid".to_string(),
        PatternType::Gray125 => "gray125".to_string(),
        PatternType::Gray0625 => "gray0625".to_string(),
        PatternType::DarkGray => "darkgray".to_string(),
        PatternType::MediumGray => "mediumgray".to_string(),
        PatternType::LightGray => "lightgray".to_string(),
        PatternType::DarkHorizontal => "darkhorizontal".to_string(),
        PatternType::DarkVertical => "darkvertical".to_string(),
        PatternType::DarkDown => "darkdown".to_string(),
        PatternType::DarkUp => "darkup".to_string(),
        PatternType::DarkGrid => "darkgrid".to_string(),
        PatternType::DarkTrellis => "darktrellis".to_string(),
        PatternType::LightHorizontal => "lighthorizontal".to_string(),
        PatternType::LightVertical => "lightvertical".to_string(),
        PatternType::LightDown => "lightdown".to_string(),
        PatternType::LightUp => "lightup".to_string(),
        PatternType::LightGrid => "lightgrid".to_string(),
        PatternType::LightTrellis => "lighttrellis".to_string(),
    }
}
